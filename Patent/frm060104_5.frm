VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_5 
   BorderStyle     =   1  '單線固定
   Caption         =   " 外專發文-變更事項"
   ClientHeight    =   7035
   ClientLeft      =   240
   ClientTop       =   585
   ClientWidth     =   8625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   8625
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   5
      Top             =   5850
      Width           =   375
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4875
      Left            =   180
      TabIndex        =   27
      Top             =   960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8599
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "申請人"
      TabPicture(0)   =   "frm060104_5.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17(14)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17(13)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label17(12)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label17(11)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label17(10)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label17(9)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label17(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label17(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label17(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label17(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label17(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label17(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label17(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label17(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label17(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(8)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(9)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(10)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(11)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(12)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(13)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(14)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(15)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(16)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(17)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(18)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label1(19)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "地址"
      TabPicture(1)   =   "frm060104_5.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(34)"
      Tab(1).Control(1)=   "Label1(33)"
      Tab(1).Control(2)=   "Label1(32)"
      Tab(1).Control(3)=   "Label1(31)"
      Tab(1).Control(4)=   "Label1(30)"
      Tab(1).Control(5)=   "Label1(29)"
      Tab(1).Control(6)=   "Label1(28)"
      Tab(1).Control(7)=   "Label1(27)"
      Tab(1).Control(8)=   "Label1(26)"
      Tab(1).Control(9)=   "Label1(25)"
      Tab(1).Control(10)=   "Label1(24)"
      Tab(1).Control(11)=   "Label1(23)"
      Tab(1).Control(12)=   "Label1(22)"
      Tab(1).Control(13)=   "Label1(21)"
      Tab(1).Control(14)=   "Label1(20)"
      Tab(1).Control(15)=   "Label61"
      Tab(1).Control(16)=   "Label63"
      Tab(1).Control(17)=   "Label65"
      Tab(1).Control(18)=   "Label67"
      Tab(1).Control(19)=   "Label69"
      Tab(1).Control(20)=   "Label71"
      Tab(1).Control(21)=   "Label73"
      Tab(1).Control(22)=   "Label75"
      Tab(1).Control(23)=   "Label77"
      Tab(1).Control(24)=   "Label79"
      Tab(1).Control(25)=   "Label81"
      Tab(1).Control(26)=   "Label83"
      Tab(1).Control(27)=   "Label85"
      Tab(1).Control(28)=   "Label87"
      Tab(1).Control(29)=   "Label89"
      Tab(1).ControlCount=   30
      TabCaption(2)   =   "中譯文/印鑑"
      TabPicture(2)   =   "frm060104_5.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label17(31)"
      Tab(2).Control(1)=   "Label17(30)"
      Tab(2).Control(2)=   "Label17(29)"
      Tab(2).Control(3)=   "Label17(28)"
      Tab(2).Control(4)=   "Label17(27)"
      Tab(2).Control(5)=   "Label17(26)"
      Tab(2).Control(6)=   "Label17(25)"
      Tab(2).Control(7)=   "Label17(24)"
      Tab(2).Control(8)=   "Label17(23)"
      Tab(2).Control(9)=   "Label17(22)"
      Tab(2).Control(10)=   "Label17(21)"
      Tab(2).Control(11)=   "Label17(20)"
      Tab(2).Control(12)=   "Label17(19)"
      Tab(2).Control(13)=   "Label17(18)"
      Tab(2).Control(14)=   "Label17(17)"
      Tab(2).Control(15)=   "Label17(16)"
      Tab(2).Control(16)=   "Label17(15)"
      Tab(2).Control(17)=   "Label1(58)"
      Tab(2).Control(18)=   "Label1(57)"
      Tab(2).Control(19)=   "Label1(56)"
      Tab(2).Control(20)=   "Label1(55)"
      Tab(2).Control(21)=   "Label1(54)"
      Tab(2).Control(22)=   "Label1(53)"
      Tab(2).Control(23)=   "Label1(52)"
      Tab(2).Control(24)=   "Label1(51)"
      Tab(2).Control(25)=   "Label1(48)"
      Tab(2).Control(26)=   "Label1(49)"
      Tab(2).Control(27)=   "Label1(47)"
      Tab(2).Control(28)=   "Label1(46)"
      Tab(2).Control(29)=   "Label1(45)"
      Tab(2).Control(30)=   "Label1(44)"
      Tab(2).Control(31)=   "Label1(43)"
      Tab(2).Control(32)=   "Label1(42)"
      Tab(2).Control(33)=   "Label1(41)"
      Tab(2).ControlCount=   34
      TabCaption(3)   =   "代表人1"
      TabPicture(3)   =   "frm060104_5.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(67)"
      Tab(3).Control(1)=   "Label1(66)"
      Tab(3).Control(2)=   "Label1(65)"
      Tab(3).Control(3)=   "Label1(64)"
      Tab(3).Control(4)=   "Label1(63)"
      Tab(3).Control(5)=   "Label1(62)"
      Tab(3).Control(6)=   "Label1(61)"
      Tab(3).Control(7)=   "Label1(60)"
      Tab(3).Control(8)=   "Label1(59)"
      Tab(3).Control(9)=   "Label92(14)"
      Tab(3).Control(10)=   "Label92(13)"
      Tab(3).Control(11)=   "Label92(12)"
      Tab(3).Control(12)=   "Label92(11)"
      Tab(3).Control(13)=   "Label92(10)"
      Tab(3).Control(14)=   "Label92(9)"
      Tab(3).Control(15)=   "Label92(8)"
      Tab(3).Control(16)=   "Label92(7)"
      Tab(3).Control(17)=   "Label92(6)"
      Tab(3).Control(18)=   "Label92(5)"
      Tab(3).Control(19)=   "Label92(4)"
      Tab(3).Control(20)=   "Label92(3)"
      Tab(3).Control(21)=   "Label92(2)"
      Tab(3).Control(22)=   "Label92(1)"
      Tab(3).Control(23)=   "Label92(0)"
      Tab(3).Control(24)=   "Label1(35)"
      Tab(3).Control(25)=   "Label1(36)"
      Tab(3).Control(26)=   "Label1(37)"
      Tab(3).Control(27)=   "Label1(38)"
      Tab(3).Control(28)=   "Label1(39)"
      Tab(3).Control(29)=   "Label1(40)"
      Tab(3).ControlCount=   30
      TabCaption(4)   =   "代表人2"
      TabPicture(4)   =   "frm060104_5.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label92(17)"
      Tab(4).Control(1)=   "Label1(70)"
      Tab(4).Control(2)=   "Label1(82)"
      Tab(4).Control(3)=   "Label1(81)"
      Tab(4).Control(4)=   "Label1(80)"
      Tab(4).Control(5)=   "Label1(79)"
      Tab(4).Control(6)=   "Label1(78)"
      Tab(4).Control(7)=   "Label1(77)"
      Tab(4).Control(8)=   "Label1(76)"
      Tab(4).Control(9)=   "Label1(75)"
      Tab(4).Control(10)=   "Label1(74)"
      Tab(4).Control(11)=   "Label1(73)"
      Tab(4).Control(12)=   "Label1(72)"
      Tab(4).Control(13)=   "Label1(71)"
      Tab(4).Control(14)=   "Label1(69)"
      Tab(4).Control(15)=   "Label1(68)"
      Tab(4).Control(16)=   "Label92(29)"
      Tab(4).Control(17)=   "Label92(28)"
      Tab(4).Control(18)=   "Label92(27)"
      Tab(4).Control(19)=   "Label92(26)"
      Tab(4).Control(20)=   "Label92(25)"
      Tab(4).Control(21)=   "Label92(24)"
      Tab(4).Control(22)=   "Label92(23)"
      Tab(4).Control(23)=   "Label92(22)"
      Tab(4).Control(24)=   "Label92(21)"
      Tab(4).Control(25)=   "Label92(20)"
      Tab(4).Control(26)=   "Label92(19)"
      Tab(4).Control(27)=   "Label92(18)"
      Tab(4).Control(28)=   "Label92(16)"
      Tab(4).Control(29)=   "Label92(15)"
      Tab(4).ControlCount=   30
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(日)6:"
         Height          =   180
         Index           =   17
         Left            =   -74760
         TabIndex        =   188
         Top             =   915
         Width           =   972
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   70
         Left            =   -73560
         TabIndex        =   187
         Top             =   915
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   82
         Left            =   -73560
         TabIndex        =   186
         Top             =   4425
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   81
         Left            =   -73560
         TabIndex        =   185
         Top             =   4185
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   80
         Left            =   -73560
         TabIndex        =   184
         Top             =   3915
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   79
         Left            =   -73560
         TabIndex        =   183
         Top             =   3555
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   78
         Left            =   -73560
         TabIndex        =   182
         Top             =   3285
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   77
         Left            =   -73560
         TabIndex        =   181
         Top             =   3015
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   76
         Left            =   -73560
         TabIndex        =   180
         Top             =   2655
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   75
         Left            =   -73560
         TabIndex        =   179
         Top             =   2385
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   74
         Left            =   -73560
         TabIndex        =   178
         Top             =   2115
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   73
         Left            =   -73560
         TabIndex        =   177
         Top             =   1785
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   72
         Left            =   -73560
         TabIndex        =   176
         Top             =   1515
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   71
         Left            =   -73560
         TabIndex        =   175
         Top             =   1245
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   69
         Left            =   -73560
         TabIndex        =   174
         Top             =   645
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   68
         Left            =   -73560
         TabIndex        =   173
         Top             =   375
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(日)10:"
         Height          =   180
         Index           =   29
         Left            =   -74760
         TabIndex        =   172
         Top             =   4425
         Width           =   1068
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(英)10:"
         Height          =   180
         Index           =   28
         Left            =   -74760
         TabIndex        =   171
         Top             =   4185
         Width           =   1068
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(中)10:"
         Height          =   180
         Index           =   27
         Left            =   -74760
         TabIndex        =   170
         Top             =   3915
         Width           =   1068
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(日)9:"
         Height          =   180
         Index           =   26
         Left            =   -74760
         TabIndex        =   169
         Top             =   3555
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(英)9:"
         Height          =   180
         Index           =   25
         Left            =   -74760
         TabIndex        =   168
         Top             =   3285
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(中)9:"
         Height          =   180
         Index           =   24
         Left            =   -74760
         TabIndex        =   167
         Top             =   3015
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(日)8:"
         Height          =   180
         Index           =   23
         Left            =   -74760
         TabIndex        =   166
         Top             =   2655
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(英)8:"
         Height          =   180
         Index           =   22
         Left            =   -74760
         TabIndex        =   165
         Top             =   2385
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(中)8:"
         Height          =   180
         Index           =   21
         Left            =   -74760
         TabIndex        =   164
         Top             =   2115
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(日)7:"
         Height          =   180
         Index           =   20
         Left            =   -74760
         TabIndex        =   163
         Top             =   1785
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(英)7:"
         Height          =   180
         Index           =   19
         Left            =   -74760
         TabIndex        =   162
         Top             =   1515
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(中)7:"
         Height          =   180
         Index           =   18
         Left            =   -74760
         TabIndex        =   161
         Top             =   1245
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(英)6:"
         Height          =   180
         Index           =   16
         Left            =   -74760
         TabIndex        =   160
         Top             =   645
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(中)6:"
         Height          =   180
         Index           =   15
         Left            =   -74760
         TabIndex        =   159
         Top             =   375
         Width           =   972
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   67
         Left            =   -73560
         TabIndex        =   158
         Top             =   4425
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   66
         Left            =   -73560
         TabIndex        =   157
         Top             =   4185
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   65
         Left            =   -73560
         TabIndex        =   156
         Top             =   3915
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   64
         Left            =   -73560
         TabIndex        =   155
         Top             =   3555
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   63
         Left            =   -73560
         TabIndex        =   154
         Top             =   3285
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   62
         Left            =   -73560
         TabIndex        =   153
         Top             =   3015
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   61
         Left            =   -73560
         TabIndex        =   152
         Top             =   2655
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   60
         Left            =   -73560
         TabIndex        =   151
         Top             =   2385
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   59
         Left            =   -73560
         TabIndex        =   150
         Top             =   2115
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(日)5:"
         Height          =   180
         Index           =   14
         Left            =   -74760
         TabIndex        =   149
         Top             =   4425
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(英)5:"
         Height          =   180
         Index           =   13
         Left            =   -74760
         TabIndex        =   148
         Top             =   4185
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(中)5:"
         Height          =   180
         Index           =   12
         Left            =   -74760
         TabIndex        =   147
         Top             =   3915
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(日)4:"
         Height          =   180
         Index           =   11
         Left            =   -74760
         TabIndex        =   146
         Top             =   3555
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(英)4:"
         Height          =   180
         Index           =   10
         Left            =   -74760
         TabIndex        =   145
         Top             =   3285
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(中)4:"
         Height          =   180
         Index           =   9
         Left            =   -74760
         TabIndex        =   144
         Top             =   3015
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(日)3:"
         Height          =   180
         Index           =   8
         Left            =   -74760
         TabIndex        =   143
         Top             =   2655
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(英)3:"
         Height          =   180
         Index           =   7
         Left            =   -74760
         TabIndex        =   142
         Top             =   2385
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(中)3:"
         Height          =   180
         Index           =   6
         Left            =   -74760
         TabIndex        =   141
         Top             =   2115
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(日)2:"
         Height          =   180
         Index           =   5
         Left            =   -74760
         TabIndex        =   140
         Top             =   1785
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(英)2:"
         Height          =   180
         Index           =   4
         Left            =   -74760
         TabIndex        =   139
         Top             =   1515
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(中)2:"
         Height          =   180
         Index           =   3
         Left            =   -74760
         TabIndex        =   138
         Top             =   1245
         Width           =   975
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(日)1:"
         Height          =   180
         Index           =   2
         Left            =   -74760
         TabIndex        =   137
         Top             =   915
         Width           =   972
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(英)1:"
         Height          =   180
         Index           =   1
         Left            =   -74760
         TabIndex        =   136
         Top             =   645
         Width           =   972
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文10"
         Height          =   180
         Index           =   31
         Left            =   -74760
         TabIndex        =   135
         Top             =   4185
         Width           =   1260
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文9"
         Height          =   180
         Index           =   30
         Left            =   -74760
         TabIndex        =   134
         Top             =   3915
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文8"
         Height          =   180
         Index           =   29
         Left            =   -74760
         TabIndex        =   133
         Top             =   3645
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文7"
         Height          =   180
         Index           =   28
         Left            =   -74760
         TabIndex        =   132
         Top             =   3375
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文6"
         Height          =   180
         Index           =   27
         Left            =   -74760
         TabIndex        =   131
         Top             =   3105
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文5"
         Height          =   180
         Index           =   26
         Left            =   -74760
         TabIndex        =   130
         Top             =   2835
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文4"
         Height          =   180
         Index           =   25
         Left            =   -74760
         TabIndex        =   129
         Top             =   2565
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文3"
         Height          =   180
         Index           =   24
         Left            =   -74760
         TabIndex        =   128
         Top             =   2295
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文2"
         Height          =   180
         Index           =   23
         Left            =   -74760
         TabIndex        =   127
         Top             =   2025
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人中譯文5"
         Height          =   180
         Index           =   22
         Left            =   -74760
         TabIndex        =   126
         Top             =   1425
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人中譯文4"
         Height          =   180
         Index           =   21
         Left            =   -74760
         TabIndex        =   125
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人中譯文3"
         Height          =   180
         Index           =   20
         Left            =   -74760
         TabIndex        =   124
         Top             =   915
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人中譯文2"
         Height          =   180
         Index           =   19
         Left            =   -74760
         TabIndex        =   123
         Top             =   645
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人印鑑"
         Height          =   180
         Index           =   18
         Left            =   -70920
         TabIndex        =   122
         Top             =   4530
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "代表人印鑑"
         Height          =   180
         Index           =   17
         Left            =   -74760
         TabIndex        =   121
         Top             =   4530
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人中譯文1"
         Height          =   180
         Index           =   16
         Left            =   -74760
         TabIndex        =   120
         Top             =   372
         Width           =   1176
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文1"
         Height          =   180
         Index           =   15
         Left            =   -74760
         TabIndex        =   119
         Top             =   1755
         Width           =   1170
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   58
         Left            =   -73440
         TabIndex        =   118
         Top             =   4185
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   57
         Left            =   -73440
         TabIndex        =   117
         Top             =   3915
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   56
         Left            =   -73440
         TabIndex        =   116
         Top             =   3645
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   55
         Left            =   -73440
         TabIndex        =   115
         Top             =   3375
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   54
         Left            =   -73440
         TabIndex        =   114
         Top             =   3105
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   53
         Left            =   -73440
         TabIndex        =   113
         Top             =   2835
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   52
         Left            =   -73440
         TabIndex        =   112
         Top             =   2565
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   51
         Left            =   -73440
         TabIndex        =   111
         Top             =   2295
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "代表人(中)1:"
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   110
         Top             =   375
         Width           =   972
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   35
         Left            =   -73560
         TabIndex        =   109
         Top             =   375
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   36
         Left            =   -73560
         TabIndex        =   108
         Top             =   645
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   37
         Left            =   -73560
         TabIndex        =   107
         Top             =   915
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   38
         Left            =   -73560
         TabIndex        =   106
         Top             =   1245
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   39
         Left            =   -73560
         TabIndex        =   105
         Top             =   1515
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   40
         Left            =   -73560
         TabIndex        =   104
         Top             =   1785
         Width           =   6540
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11536;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   48
         Left            =   -73440
         TabIndex        =   103
         Top             =   4530
         Width           =   2340
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "4128;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   49
         Left            =   -69960
         TabIndex        =   101
         Top             =   4530
         Width           =   2940
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5186;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   47
         Left            =   -73440
         TabIndex        =   100
         Top             =   2025
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   46
         Left            =   -73440
         TabIndex        =   99
         Top             =   1755
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   45
         Left            =   -73440
         TabIndex        =   98
         Top             =   1425
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   44
         Left            =   -73440
         TabIndex        =   97
         Top             =   1155
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   43
         Left            =   -73440
         TabIndex        =   96
         Top             =   915
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   42
         Left            =   -73440
         TabIndex        =   95
         Top             =   645
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   41
         Left            =   -73440
         TabIndex        =   94
         Top             =   375
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   34
         Left            =   -73440
         TabIndex        =   93
         Top             =   4500
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   33
         Left            =   -73440
         TabIndex        =   92
         Top             =   4215
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   32
         Left            =   -73440
         TabIndex        =   91
         Top             =   3945
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   31
         Left            =   -73440
         TabIndex        =   90
         Top             =   3585
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   30
         Left            =   -73440
         TabIndex        =   89
         Top             =   3315
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   29
         Left            =   -73440
         TabIndex        =   88
         Top             =   3045
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   28
         Left            =   -73440
         TabIndex        =   87
         Top             =   2685
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   27
         Left            =   -73440
         TabIndex        =   86
         Top             =   2415
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   26
         Left            =   -73440
         TabIndex        =   85
         Top             =   2145
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   25
         Left            =   -73440
         TabIndex        =   84
         Top             =   1815
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   24
         Left            =   -73440
         TabIndex        =   83
         Top             =   1545
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   23
         Left            =   -73440
         TabIndex        =   82
         Top             =   1275
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   22
         Left            =   -73440
         TabIndex        =   81
         Top             =   930
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   21
         Left            =   -73440
         TabIndex        =   80
         Top             =   645
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   20
         Left            =   -73440
         TabIndex        =   79
         Top             =   375
         Width           =   6450
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11377;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   19
         Left            =   1320
         TabIndex        =   78
         Top             =   4425
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   18
         Left            =   1320
         TabIndex        =   77
         Top             =   4185
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   17
         Left            =   1320
         TabIndex        =   76
         Top             =   3915
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   16
         Left            =   1320
         TabIndex        =   75
         Top             =   3555
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   15
         Left            =   1320
         TabIndex        =   74
         Top             =   3285
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   14
         Left            =   1320
         TabIndex        =   73
         Top             =   3015
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   13
         Left            =   1320
         TabIndex        =   72
         Top             =   2655
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   12
         Left            =   1320
         TabIndex        =   71
         Top             =   2385
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   11
         Left            =   1320
         TabIndex        =   70
         Top             =   2115
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   10
         Left            =   1320
         TabIndex        =   69
         Top             =   1785
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   9
         Left            =   1320
         TabIndex        =   68
         Top             =   1515
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   8
         Left            =   1320
         TabIndex        =   67
         Top             =   1245
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   66
         Top             =   915
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   65
         Top             =   645
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   64
         Top             =   375
         Width           =   6690
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11800;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(中)1:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   57
         Top             =   372
         Width           =   1152
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(英)1:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   56
         Top             =   645
         Width           =   1155
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(日)1:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   55
         Top             =   930
         Width           =   1155
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(中)2:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   54
         Top             =   1275
         Width           =   1155
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(英)2:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   53
         Top             =   1545
         Width           =   1155
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(日)2:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   52
         Top             =   1815
         Width           =   1155
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(中)3:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   51
         Top             =   2145
         Width           =   1155
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(英)3:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   50
         Top             =   2415
         Width           =   1155
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(日)3:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   49
         Top             =   2685
         Width           =   1155
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(中)4:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   48
         Top             =   3045
         Width           =   1155
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(英)4:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   47
         Top             =   3315
         Width           =   1155
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(日)4:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   46
         Top             =   3585
         Width           =   1155
      End
      Begin VB.Label Label85 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(中)5:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   45
         Top             =   3945
         Width           =   1155
      End
      Begin VB.Label Label87 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(英)5:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   44
         Top             =   4215
         Width           =   1155
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(日)5:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   43
         Top             =   4485
         Width           =   1155
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(中)1:"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   372
         Width           =   972
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(英)1:"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   41
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(日)1:"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   40
         Top             =   915
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(中)2:"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   39
         Top             =   1245
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(英)2:"
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   38
         Top             =   1515
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(日)2:"
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   37
         Top             =   1785
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(中)3:"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   36
         Top             =   2115
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(英)3:"
         Height          =   180
         Index           =   7
         Left            =   240
         TabIndex        =   35
         Top             =   2385
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(日)3:"
         Height          =   180
         Index           =   8
         Left            =   240
         TabIndex        =   34
         Top             =   2655
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(中)4:"
         Height          =   180
         Index           =   9
         Left            =   240
         TabIndex        =   33
         Top             =   3015
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(英)4:"
         Height          =   180
         Index           =   10
         Left            =   240
         TabIndex        =   32
         Top             =   3285
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(日)4:"
         Height          =   180
         Index           =   11
         Left            =   240
         TabIndex        =   31
         Top             =   3555
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(中)5:"
         Height          =   180
         Index           =   12
         Left            =   240
         TabIndex        =   30
         Top             =   3915
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(英)5:"
         Height          =   180
         Index           =   13
         Left            =   240
         TabIndex        =   29
         Top             =   4185
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人(日)5:"
         Height          =   180
         Index           =   14
         Left            =   240
         TabIndex        =   28
         Top             =   4455
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   6480
      TabIndex        =   12
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   360
      Index           =   1
      Left            =   7320
      TabIndex        =   13
      Top             =   60
      Width           =   1200
   End
   Begin VB.CheckBox Check12 
      Caption         =   "其他:"
      Height          =   180
      Left            =   210
      TabIndex        =   10
      Top             =   7530
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   1650
      MaxLength       =   9
      TabIndex        =   11
      Top             =   7530
      Width           =   3375
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Check1"
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   6180
      Width           =   255
   End
   Begin VB.CheckBox Check10 
      Caption         =   "圖樣"
      Height          =   180
      Left            =   7080
      TabIndex        =   4
      Top             =   5880
      Width           =   735
   End
   Begin VB.CheckBox Check9 
      Caption         =   "正商標號數"
      Height          =   180
      Left            =   3180
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   4440
      MaxLength       =   9
      TabIndex        =   3
      Top             =   5850
      Width           =   1455
   End
   Begin VB.CheckBox Check8 
      Caption         =   "專利種類"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1620
      MaxLength       =   9
      TabIndex        =   1
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1116
      MaxLength       =   3
      TabIndex        =   18
      Top             =   48
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1596
      MaxLength       =   6
      TabIndex        =   17
      Top             =   48
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2436
      MaxLength       =   1
      TabIndex        =   16
      Top             =   48
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2676
      MaxLength       =   2
      TabIndex        =   15
      Top             =   48
      Width           =   375
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   2
      Left            =   1620
      TabIndex        =   9
      Top             =   6660
      Width           =   6735
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "11880;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   1
      Left            =   1620
      TabIndex        =   8
      Top             =   6420
      Width           =   6735
      VariousPropertyBits=   671105051
      MaxLength       =   180
      Size            =   "11880;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   0
      Left            =   1620
      TabIndex        =   7
      Top             =   6180
      Width           =   6735
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "11880;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   285
      Index           =   50
      Left            =   2040
      TabIndex        =   102
      Top             =   5880
      Width           =   480
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   285
      Index           =   4
      Left            =   4590
      TabIndex        =   63
      Top             =   660
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3228;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   285
      Index           =   3
      Left            =   1110
      TabIndex        =   62
      Top             =   660
      Width           =   2010
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3545;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   285
      Index           =   2
      Left            =   1110
      TabIndex        =   61
      Top             =   348
      Width           =   2010
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3545;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   285
      Index           =   1
      Left            =   4590
      TabIndex        =   60
      Top             =   45
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3228;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   285
      Index           =   0
      Left            =   4596
      TabIndex        =   59
      Top             =   348
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3228;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label129 
      Height          =   255
      Left            =   7380
      TabIndex        =   58
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label133 
      AutoSize        =   -1  'True
      Caption         =   "案件外文名稱:"
      Height          =   180
      Left            =   420
      TabIndex        =   26
      Top             =   6690
      Width           =   1125
   End
   Begin VB.Label Label132 
      AutoSize        =   -1  'True
      Caption         =   "案件英文名稱:"
      Height          =   180
      Left            =   420
      TabIndex        =   25
      Top             =   6450
      Width           =   1125
   End
   Begin VB.Label Label131 
      AutoSize        =   -1  'True
      Caption         =   "案件中文名稱:"
      Height          =   180
      Left            =   420
      TabIndex        =   14
      Top             =   6210
      Width           =   1245
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3630
      TabIndex        =   24
      Top             =   660
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   150
      TabIndex        =   23
      Top             =   660
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   3636
      TabIndex        =   22
      Top             =   348
      Width           =   768
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   156
      TabIndex        =   21
      Top             =   348
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號:"
      Height          =   180
      Left            =   3636
      TabIndex        =   20
      Top             =   48
      Width           =   768
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   156
      TabIndex        =   19
      Top             =   48
      Width           =   768
   End
End
Attribute VB_Name = "frm060104_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/15 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String, intWhere As Integer
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim intGo As Integer


Public Function FormSave() As Boolean
'Add By Cheng 2003/04/10
Dim i As Integer, intStep As Integer, strTxt(1 To 20) As String, j As Integer
Dim strCe(99) As String, bolChk As Boolean
Dim strOldData As String
Dim strTmp(1 To 5) As String

strOldData = Empty
 strExc(0) = ""
   If Check8.Value = 1 Then
      If Text6 = "" Then
         MsgBox "專利種類不得空白 !", vbCritical
         Text6.SetFocus
         Exit Function
      End If
   End If
   If Check9.Value = 1 Then
      If Text5 = "" Then
         MsgBox "正商標號數不得空白 !", vbCritical
         Text5.SetFocus
         Exit Function
      End If
   End If
   If Check10.Value = 1 Then
      If Text8 = "" Then
         MsgBox "圖樣不得空白 !", vbCritical
         Text8.SetFocus
         Exit Function
      End If
   End If
   If Check11.Value = 1 Then
      If Text7(0) = "" And Text7(1) = "" And Text7(2) = "" Then
         MsgBox "案件名稱不得空白 !", vbCritical
         Text7(0).SetFocus
         Exit Function
      End If
   End If
   If Check12.Value = 1 Then
      If Text10 = "" Then
         MsgBox "其他不得空白 !", vbCritical
         Exit Function
      End If
   End If
    '專利種類
   If Check8.Value = 1 Then
      strExc(0) = strExc(0) & "CE40='1',CE39=" & CNULL(ChgSQL(Text6)) & ","
   Else
      strExc(0) = strExc(0) & "CE40='',CE39='',"
   End If
    '正商標號數
   If Check9.Value = 1 Then
      strExc(0) = strExc(0) & "CE58='1',CE57=" & CNULL(ChgSQL(Text5)) & ","
   Else
      strExc(0) = strExc(0) & "CE58='',CE57='',"
   End If
    '圖樣
   If Check10.Value = 1 Then
      strExc(0) = strExc(0) & "CE60='1',CE59=" & CNULL(Text8) & ","
   Else
      strExc(0) = strExc(0) & "CE60='',CE59='',"
   End If
    '案件名稱
   If Check11.Value = 1 Then
      strExc(0) = strExc(0) & "CE44='1',CE41=" & CNULL(Text7(0)) & "," & _
      "CE42=" & CNULL(Text7(1)) & ",CE43=" & CNULL(Text7(2)) & ","
   Else
      strExc(0) = strExc(0) & "CE44='',CE41='',CE42='',CE43='',"
   End If
   If Check12.Value = 1 Then
      strExc(0) = strExc(0) & "CE62='1',CE61=" & CNULL(Text10)
   Else
      strExc(0) = strExc(0) & "CE62='',CE61=''"
   End If
   If Right(strExc(0), 1) = "," Then strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
   If strExc(0) <> "" Then
      strExc(1) = "UPDATE CHANGEEVENT SET " & strExc(0) & " WHERE CE01='" & strReceiveNo & "'"
      FormSave = ClsLawExecSQL(1, strExc)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ExecSQL(1, strExc)
        'Add By Cheng 2003/04/10
        strExc(0) = "SELECT * FROM CHANGEEVENT WHERE CE01='" & strReceiveNo & "'"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            With RsTemp
               For i = 1 To 99
                  If IsNull(.Fields(i - 1)) Then
                     strCe(i) = ""
                  Else
                     strCe(i) = .Fields(i - 1)
                  End If
               Next
            End With
            strExc(1) = ""
            strExc(2) = ""
            strExc(3) = ""
            
            '申請日 10
            If strCe(2) <> "" Then
               strExc(1) = strExc(1) & "申請日 : " & strCe(2) & ","
               strExc(2) = strExc(2) & "PA10=" & strCe(2) & ","
               strExc(3) = strExc(3) & "CE03='1',"
               ' 90.07.17 modify by louis (變更事項舊資料)
               strOldData = strOldData & "申請日 : " & pa(10) & " "
            End If
            
            '申請人 26-30
            bolChk = False
            For i = 4 To 8
               If strCe(i) <> "" Then
                  bolChk = True
                  Exit For
               End If
            Next
            If bolChk = True Then
               ' 90.07.17 modify by louis (變更事項舊資料)
               strOldData = strOldData & "申請人 : "
               strExc(1) = strExc(1) & "申請人 : "
               For i = 4 To 8
                  If strCe(i) <> "" Then
                     strExc(1) = strExc(1) & strCe(i) & ","
                     'edit by nickc 2007/02/02 不用 dll 了
                     'If objPublicData.GetCustomerNameAndAddress(strCe(i), strTmp(5), strTmp(1), strTmp(2), strTmp(3)) Then
                     If ClsPDGetCustomerNameAndAddress(strCe(i), strTmp(5), strTmp(1), strTmp(2), strTmp(3)) Then
                        strExc(2) = strExc(2) & "PA" & i + 27 & "=" & CNULL(ChgSQL(strTmp(1))) & ",PA" & i + 32 & "=" & CNULL(ChgSQL(strTmp(2))) & ",PA" & i + 37 & "=" & CNULL(ChgSQL(strTmp(3))) & ","
                     End If
                  End If
                    'Modify By Cheng 2003/05/13
'                  strExc(2) = strExc(2) & "PA" & i + 22 & "=" & CNULL(strCe(i)) & ","
                  strExc(2) = strExc(2) & "PA" & i + 22 & "=" & CNULL(ChangeCustomerL(strCe(i))) & ","
               Next
               ' 90.07.17 modify by louis (變更事項舊資料)
               If IsEmptyText(strCe(4)) = False Then
                  strOldData = strOldData & pa(26) & " "
               End If
               If IsEmptyText(strCe(5)) = False Then
                  strOldData = strOldData & pa(27) & " "
               End If
               If IsEmptyText(strCe(6)) = False Then
                  strOldData = strOldData & pa(28) & " "
               End If
               If IsEmptyText(strCe(7)) = False Then
                  strOldData = strOldData & pa(29) & " "
               End If
               If IsEmptyText(strCe(8)) = False Then
                  strOldData = strOldData & pa(30) & " "
               End If
               strExc(3) = strExc(3) & "CE09='1',"
            Else
               '申請地址 31-45
               bolChk = False
               For i = 23 To 37
                  If strCe(i) <> "" Then
                     bolChk = True
                     Exit For
                  End If
               Next
               If bolChk = True Then
                  ' 90.07.17 modify by louis (變更事項舊資料)
                  strOldData = strOldData & "申請地址 : "
                  strExc(1) = strExc(1) & "申請地址 : "
                  For i = 23 To 37
                     If strCe(i) <> "" Then
                        strExc(1) = strExc(1) & strCe(i) & ","
                     End If
                     strExc(2) = strExc(2) & "PA" & i + 8 & "=" & CNULL(strCe(i)) & ","
                  Next
                  strExc(3) = strExc(3) & "CE38='1',"
                  ' 90.07.17 modify by louis (變更事項舊資料)
                  If IsEmptyText(strCe(23)) = False Then
                     strOldData = strOldData & pa(31) & " "
                  End If
                  If IsEmptyText(strCe(24)) = False Then
                     strOldData = strOldData & pa(36) & " "
                  End If
                  If IsEmptyText(strCe(25)) = False Then
                     strOldData = strOldData & pa(41) & " "
                  End If
                  If IsEmptyText(strCe(26)) = False Then
                     strOldData = strOldData & pa(32) & " "
                  End If
                  If IsEmptyText(strCe(27)) = False Then
                     strOldData = strOldData & pa(37) & " "
                  End If
                  If IsEmptyText(strCe(28)) = False Then
                     strOldData = strOldData & pa(42) & " "
                  End If
                  If IsEmptyText(strCe(29)) = False Then
                     strOldData = strOldData & pa(33) & " "
                  End If
                  If IsEmptyText(strCe(30)) = False Then
                     strOldData = strOldData & pa(38) & " "
                  End If
                  If IsEmptyText(strCe(31)) = False Then
                     strOldData = strOldData & pa(43) & " "
                  End If
                  If IsEmptyText(strCe(32)) = False Then
                     strOldData = strOldData & pa(34) & " "
                  End If
                  If IsEmptyText(strCe(33)) = False Then
                     strOldData = strOldData & pa(39) & " "
                  End If
                  If IsEmptyText(strCe(34)) = False Then
                     strOldData = strOldData & pa(44) & " "
                  End If
                  If IsEmptyText(strCe(35)) = False Then
                     strOldData = strOldData & pa(35) & " "
                  End If
                  If IsEmptyText(strCe(36)) = False Then
                     strOldData = strOldData & pa(40) & " "
                  End If
                  If IsEmptyText(strCe(37)) = False Then
                     strOldData = strOldData & pa(45) & " "
                  End If
               End If
            End If
            
            '專利商標種類代號 08
            If strCe(39) <> "" Then
               ' 90.07.17 modify by louis (舊的資料)
               strOldData = strOldData & "專利商標種類代號 : " & pa(8) & " "
               strExc(1) = strExc(1) & "專利商標種類代號 : " & strCe(39) & ","
               strExc(2) = strExc(2) & "PA08='" & strCe(39) & "',"
               strExc(3) = strExc(3) & "CE40='1',"
            End If
            
            '案件名稱 05-07
            bolChk = False
            For i = 41 To 43
               If strCe(i) <> "" Then
                  bolChk = True
                  Exit For
               End If
            Next
            If bolChk = True Then
               ' 90.07.17 modify by louis (舊的資料)
               strOldData = strOldData & "案件名稱 : "
               strExc(1) = strExc(1) & "案件名稱 : "
               For i = 41 To 43
                  If strCe(i) <> "" Then
                     strExc(1) = strExc(1) & strCe(i) & ","
                  End If
                  strExc(2) = strExc(2) & "PA" & i - 36 & "=" & CNULL(strCe(i)) & ","
               Next
               strExc(3) = strExc(3) & "CE44='1',"
               ' 90.07.17 modify by louis (舊的資料)
               If IsEmptyText(strCe(41)) = False Then
                  strOldData = strOldData & pa(5) & " "
               End If
               If IsEmptyText(strCe(42)) = False Then
                  strOldData = strOldData & pa(6) & " "
               End If
               If IsEmptyText(strCe(43)) = False Then
                  strOldData = strOldData & pa(7) & " "
               End If
            End If
            
            '代表人 79-84
            bolChk = False
            For i = 10 To 15
               If strCe(i) <> "" Then
                  bolChk = True
                  Exit For
               End If
            Next
            If Not bolChk Then
               For i = 68 To 91
                  If strCe(i) <> "" Then
                     bolChk = True
                     Exit For
                  End If
               Next
            End If
            
            If bolChk Then
               ' 90.07.17 modify by louis (舊的資料)
               strOldData = strOldData & "代表人 : "
               strExc(1) = strExc(1) & "代表人 : "
               For i = 10 To 15
                  If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
                  strExc(2) = strExc(2) & "PA" & i + 69 & "=" & CNULL(strCe(i)) & ","
                  ' 90.07.17 modify by louis (舊的資料)
                  If IsEmptyText(strCe(i)) Then
                     strOldData = strOldData & pa(i + 69) & " "
                  End If
               Next
               For i = 68 To 91
                  If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
                  strExc(2) = strExc(2) & "PA" & i + 41 & "=" & CNULL(strCe(i)) & ","
                  ' 90.07.17 modify by louis (舊的資料)
                  If IsEmptyText(strCe(i)) Then
                     strOldData = strOldData & pa(i + 41) & " "
                  End If
               Next
               strExc(3) = strExc(3) & "CE16='1',"
            End If
            
            '代表人中譯文
            If Not bolChk Then
               bolChk = False
               For i = 63 To 64
                  If strCe(i) <> "" Then
                     bolChk = True
                     Exit For
                  End If
               Next
               If Not bolChk Then
                  For i = 92 To 99
                     If strCe(i) <> "" Then
                        bolChk = True
                        Exit For
                     End If
                  Next
               End If
               If bolChk Then
                  strExc(1) = strExc(1) & "代表人中譯文 : "
                  strExc(2) = strExc(2) & "PA79=" & CNULL(strCe(63)) & ",PA82=" & CNULL(strCe(64)) & "," & _
                     "PA109=" & CNULL(strCe(92)) & ",PA112=" & CNULL(strCe(93)) & ",PA115=" & CNULL(strCe(94)) & "," & _
                     "PA118=" & CNULL(strCe(95)) & ",PA121=" & CNULL(strCe(96)) & ",PA124=" & CNULL(strCe(97)) & "," & _
                     "PA127=" & CNULL(strCe(98)) & ",PA130=" & CNULL(strCe(99)) & ","
                  For i = 63 To 64
                     If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
                  Next
                  For i = 92 To 99
                     If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
                  Next
                  strExc(3) = strExc(3) & "CE65='1',"
               End If
            End If
            
            ' 90.07.17 modify by louis
            ' 申請人中議文
            bolChk = False
            For i = 17 To 21
               If strCe(i) <> "" Then
                  bolChk = True
                  Exit For
               End If
            Next
            If bolChk = True Then
               strExc(3) = strExc(3) & "CE22='1',"
            End If
            
            If strExc(1) <> "" Then
               For i = 2 To 3
                  If Right(strExc(i), 1) = "," Then strExc(i) = Left(strExc(i), Len(strExc(i)) - 1)
               Next
               intStep = intStep + 1
               ' 90.07.17 modify by louis (儲存在案件進度檔的是舊資料)
               'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP64=CP64||'" & strExc(1) & "' WHERE CP09='" & strReceiveNo & "'"
               strTxt(intStep) = "UPDATE CASEPROGRESS SET CP64=CP64||'" & strOldData & "' WHERE CP09='" & strReceiveNo & "'"
               
              '911105 nick transation
              cnnConnection.Execute strTxt(intStep)
            
               intStep = intStep + 1
               strTxt(intStep) = "UPDATE PATENT SET " & strExc(2) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
               
              '911105 nick transation
              cnnConnection.Execute strTxt(intStep)
               
               intStep = intStep + 1
'               strTxt(intStep) = "UPDATE CHANGEEVENT SET " & strExc(3) & " WHERE CE01='" & strReceiveNo & "'"
               
              '911105 nick transation
'              cnnConnection.Execute strTxt(intStep)
               
               intStep = intStep + 1
            End If
        End If
   Else
      FormSave = True
   End If
End Function

Private Sub cmdOK_Click(Index As Integer)
   Dim bolSaveOK As Boolean 'Add by Morgan 2006/6/8
   
   
   If Index = 0 Then
      'Add by Sindy 2021/11/15 檢查畫面上的物件是否含有Unicode文字
      If PUB_ChkUniText(Me, True, True) = False Then
         Exit Sub
      End If
      
      If FormSave = False Then
         MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
      Else
         bolSaveOK = True 'Add by Morgan 2006/6/8
      End If
   End If
      
   Select Case intGo
      Case 3
         frm060104_3.Show
         frm060104_3.m_bolSaveChgEvent = bolSaveOK 'Add by Morgan 2006/6/8
      Case 7
         frm060104_7.Show
      Case 10
         frm060104_a.Show
      Case 11
         frm060104_c.Show
   End Select
   Unload Me
End Sub

Public Sub LoadMe(ByVal RecNo As String, ByVal txt1 As String, ByVal txt2 As String, _
   ByVal txt3 As String, ByVal txt4 As String, ByVal iGo As Integer)
   Text1 = txt1
   Text2 = txt2
   Text3 = txt3
   Text4 = txt4
   strReceiveNo = RecNo
   intGo = iGo
   ReadPatent
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_5 = Nothing
End Sub

Private Sub ReadPatent()
 Dim Lbl As Object, strTmp As String, i As Integer
   
   For Each Lbl In Label1
      Lbl = ""
   Next
   Label5 = ""
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   
   strExc(0) = "SELECT PA08 FROM PATENT WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields(0)) Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, rsTemp.Fields(0), strTmp, , 台灣國家代號) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, RsTemp.Fields(0), strTmp, , 台灣國家代號) = 1 Then
            Label1(2) = strTmp
         End If
      End If
   End If

   strExc(0) = "select cp45,cpm03,staff.st02 as st1,staff1.st02 as st2 from caseprogress,casepropertymap," & _
      "staff,staff staff1 where cp09='" & strReceiveNo & "' and cp01=cpm01(+) and cp10=cpm02(+) and " & _
      "cp14=staff.st01(+) and cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
      If intI = 1 Then
         For i = 0 To 3
            If Not IsNull(.Fields(i)) Then Label1(i + 1) = .Fields(i)
         Next
      End If
   End With
   
   strExc(0) = "SELECT CE04,CE05,CE06,CE07,CE08,CE23,CE24,CE25,CE26,CE27,CE28,CE29,CE30,CE31,CE32," & _
                      "CE33,CE34,CE35,CE36,CE37,CE10,CE11,CE12,CE13,CE14,CE15,CE17,CE18,CE19,CE20," & _
                      "CE21,CE63,CE64,CE53,CE51,CE09,CE38,CE16,CE22,CE65,CE54,CE52,CE39,CE40,CE57," & _
                      "CE58,CE59,CE60,CE41,CE42,CE43,CE44,CE61,CE62,CE92,CE93,CE94,CE95,CE96,CE97," & _
                      "CE98,CE99,CE68,CE69,CE70,CE71,CE72,CE73,CE74,CE75,CE76,CE77,CE78,CE79,CE80," & _
                      "CE81,CE82,CE83,CE84,CE85,CE86,CE87,CE88,CE89,CE90,CE91 FROM CHANGEEVENT WHERE CE01='" & strReceiveNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      
      For i = 0 To 4
         If Not IsNull(.Fields(i)) Then
            'edit by nickc 2007/02/05 不用 dll 了
            'If objLawDll.GetCusCAJnam(.Fields(i), strExc(1), strExc(2), strExc(3)) Then
            If ClsLawGetCusCAJnam(.Fields(i), strExc(1), strExc(2), strExc(3)) Then
               Label1(5 + i * 3) = strExc(1)
               Label1(6 + i * 3) = strExc(2)
               Label1(7 + i * 3) = strExc(3)
            End If
         End If
      Next
      
      For i = 5 To 34
         If Not IsNull(.Fields(i)) Then Label1(i + 15) = .Fields(i)
      Next
      
'      For i = 35 To 41
'         If .Fields(i) = 1 Then Check1(i - 35).Value = 1
'      Next
      
      If Not IsNull(.Fields(42)) Then Text6 = .Fields(42)
      If .Fields(43) = 1 Then Check8.Value = 1
      
      If Not IsNull(.Fields(44)) Then Text5 = .Fields(44)
      If .Fields(45) = 1 Then Check9.Value = 1
      
      If Not IsNull(.Fields(46)) Then Text8 = .Fields(46)
      If .Fields(47) = 1 Then Check10.Value = 1
      
      For i = 0 To 2
         If Not IsNull(.Fields(48 + i)) Then Text7(i) = .Fields(48 + i)
      Next
      If .Fields(51) = 1 Then Check11.Value = 1
      
      If Not IsNull(.Fields(52)) Then Text10 = .Fields(52)
      If .Fields(53) = 1 Then Check12.Value = 1
      
      '代表人中譯文3-10
      For i = 54 To 61
         If Not IsNull(.Fields(i)) Then Label1(i - 3) = .Fields(i)
      Next
      
      '代表人3-10
      For i = 62 To 85
         If Not IsNull(.Fields(i)) Then Label1(i - 3) = .Fields(i)
      Next
      
   Else
      strExc(1) = "INSERT INTO CHANGEEVENT (CE01) VALUES ('" & strReceiveNo & "')"
      'edit by nickc 2007/02/05 不用 dll 了
      'If Not objLawDll.ExecSQL(1, strExc) Then
      If Not ClsLawExecSQL(1, strExc) Then
         
      End If
   End If
   End With
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
 Dim strTmp As String
   If Check8.Value = 1 Then
      If Text6 = "" Then
         MsgBox "專利種類代號不可空白 !", vbCritical
         Cancel = True
      Else
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, Text6, strTmp, , 台灣國家代號) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, Text6, strTmp, , 台灣國家代號) = 1 Then
            Label1(50) = strTmp
         Else
            Label1(50) = ""
         End If
      End If
   Else
      Label1(50) = ""
   End If
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Check9.Value = 1 Then
      If Text5 = "" Then
         MsgBox "正商標號數不可空白 !", vbCritical
         Cancel = True
      End If
   End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Check12.Value = 1 Then
      If Text10 = "" Then
         MsgBox "其他不可空白 !", vbCritical
         Cancel = True
      End If
   End If
End Sub

Private Sub Text7_Validate(Index As Integer, Cancel As Boolean)
   If Check11.Value = 1 Then
      If Index = 2 Then
         If Text7(0) = "" And Text7(1) = "" And Text7(2) = "" Then
            MsgBox "案件名稱不可同時空白 !", vbCritical
            Cancel = True
         End If
      End If
   End If
End Sub
