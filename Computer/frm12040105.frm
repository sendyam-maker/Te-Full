VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040105 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工檔維護"
   ClientHeight    =   6190
   ClientLeft      =   2570
   ClientTop       =   2560
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm12040105.frx":0000
   ScaleHeight     =   6190
   ScaleWidth      =   8950
   Begin TabDlg.SSTab SSTab1 
      Height          =   5420
      Left            =   30
      TabIndex        =   32
      Top             =   660
      Width           =   8870
      _ExtentX        =   15646
      _ExtentY        =   9543
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm12040105.frx":0342
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(13)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblName(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(12)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblName(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(11)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(5)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(14)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label5(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblName(15)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblName(14)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label5(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblName(13)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label5(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label4"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label3"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(10)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label12"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label7"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblName(58)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label5(3)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(15)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label8"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label10"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblName(62)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label9"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lblName(61)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label13"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label1(7)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label22"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label5(4)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "lblName(69)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(2)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text1(1)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text1(0)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(16)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text1(6)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text1(5)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text1(4)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text1(3)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text1(15)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Text1(14)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Text1(13)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Text1(12)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Text1(11)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Text1(10)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Text1(17)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Text1(57)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Text1(58)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Text1(60)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Text1(62)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Text1(61)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Text1(63)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Text1(68)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Text1(69)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Label1(8)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Text1(92)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "lblName(92)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "Label5(5)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Label24"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).ControlCount=   67
      TabCaption(1)   =   "其他"
      TabPicture(1)   =   "frm12040105.frx":035E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "lblName(56)"
      Tab(1).Control(2)=   "Label19"
      Tab(1).Control(3)=   "Label20"
      Tab(1).Control(4)=   "Label23"
      Tab(1).Control(5)=   "Text1(56)"
      Tab(1).Control(6)=   "Text1(64)"
      Tab(1).Control(7)=   "Text1(65)"
      Tab(1).Control(8)=   "Frame1"
      Tab(1).ControlCount=   9
      Begin VB.Frame Frame1 
         Caption         =   "期限管制人"
         Height          =   1695
         Left            =   -74895
         TabIndex        =   58
         Top             =   450
         Width           =   8610
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   54
            Left            =   990
            TabIndex        =   26
            Top             =   1290
            Width           =   735
            VariousPropertyBits=   671105051
            MaxLength       =   6
            Size            =   "1296;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   53
            Left            =   990
            TabIndex        =   25
            Top             =   960
            Width           =   735
            VariousPropertyBits=   671105051
            MaxLength       =   6
            Size            =   "1296;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   52
            Left            =   990
            TabIndex        =   24
            Top             =   630
            Width           =   735
            VariousPropertyBits=   671105051
            MaxLength       =   6
            Size            =   "1296;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   51
            Left            =   990
            TabIndex        =   23
            Top             =   300
            Width           =   735
            VariousPropertyBits=   671105051
            MaxLength       =   6
            Size            =   "1296;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "國外部同仁時，請注意專業代號是否須同時修改；外商同時改組別"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   2910
            TabIndex        =   81
            Top             =   340
            Width           =   5220
         End
         Begin MSForms.Label lblName 
            Height          =   255
            Index           =   54
            Left            =   1800
            TabIndex        =   66
            Top             =   1313
            Width           =   1065
            BackColor       =   -2147483634
            VariousPropertyBits=   27
            Size            =   "1879;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblName 
            Height          =   255
            Index           =   53
            Left            =   1800
            TabIndex        =   65
            Top             =   983
            Width           =   1065
            BackColor       =   -2147483634
            VariousPropertyBits=   27
            Size            =   "1879;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblName 
            Height          =   255
            Index           =   52
            Left            =   1800
            TabIndex        =   64
            Top             =   653
            Width           =   1065
            BackColor       =   -2147483634
            VariousPropertyBits=   27
            Size            =   "1879;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblName 
            Height          =   255
            Index           =   51
            Left            =   1800
            TabIndex        =   63
            Top             =   323
            Width           =   1065
            BackColor       =   -2147483634
            VariousPropertyBits=   27
            Size            =   "1879;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label18 
            Caption         =   "第五級"
            Height          =   180
            Left            =   270
            TabIndex        =   62
            Top             =   1350
            Width           =   540
         End
         Begin VB.Label Label17 
            Caption         =   "第四級"
            Height          =   180
            Left            =   270
            TabIndex        =   61
            Top             =   1020
            Width           =   540
         End
         Begin VB.Label Label16 
            Caption         =   "第三級"
            Height          =   180
            Left            =   270
            TabIndex        =   60
            Top             =   690
            Width           =   540
         End
         Begin VB.Label Label15 
            Caption         =   "第二級"
            Height          =   180
            Left            =   270
            TabIndex        =   59
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Null:有特定英核主管或不需英核)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   6090
         TabIndex        =   91
         Top             =   3090
         Width           =   2570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "新部門別"
         Height          =   180
         Index           =   5
         Left            =   3510
         TabIndex        =   90
         Top             =   3220
         Width           =   720
      End
      Begin MSForms.Label lblName 
         Height          =   250
         Index           =   92
         Left            =   5030
         TabIndex        =   89
         Top             =   3190
         Width           =   1780
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "3133;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   92
         Left            =   4270
         TabIndex        =   14
         Top             =   3160
         Width           =   730
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1291;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專業代號 －對國外：最好不要重覆"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   8
         Left            =   4920
         TabIndex        =   88
         Top             =   2040
         Width           =   2750
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   69
         Left            =   8430
         TabIndex        =   21
         Top             =   4676
         Width           =   315
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "556;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   68
         Left            =   7575
         TabIndex        =   86
         Top             =   2244
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   65
         Left            =   -73065
         TabIndex        =   29
         Top             =   2910
         Width           =   3030
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "5345;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   64
         Left            =   -73065
         TabIndex        =   28
         Top             =   2580
         Width           =   3030
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5345;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   63
         Left            =   5730
         TabIndex        =   12
         Top             =   2820
         Width           =   290
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   61
         Left            =   2115
         TabIndex        =   9
         Top             =   2548
         Width           =   735
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   62
         Left            =   2115
         TabIndex        =   11
         Top             =   2852
         Width           =   735
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   60
         Left            =   6270
         TabIndex        =   10
         Top             =   2550
         Width           =   290
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   58
         Left            =   5475
         TabIndex        =   17
         Top             =   3764
         Width           =   735
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   57
         Left            =   6570
         TabIndex        =   3
         Top             =   1028
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   56
         Left            =   -72720
         TabIndex        =   27
         Top             =   2220
         Width           =   735
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   17
         Left            =   2115
         TabIndex        =   22
         Top             =   4980
         Width           =   4830
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "8520;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   10
         Left            =   2115
         TabIndex        =   13
         Top             =   3156
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   11
         Left            =   2115
         TabIndex        =   15
         Top             =   3460
         Width           =   4815
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "8493;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   12
         Left            =   2115
         TabIndex        =   16
         Top             =   3764
         Width           =   1215
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   13
         Left            =   2115
         TabIndex        =   18
         Top             =   4065
         Width           =   2415
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4260;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   14
         Left            =   2115
         TabIndex        =   19
         Top             =   4372
         Width           =   735
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   15
         Left            =   2115
         TabIndex        =   20
         Top             =   4676
         Width           =   735
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   2115
         TabIndex        =   4
         Top             =   1332
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   4
         Left            =   2115
         TabIndex        =   5
         Top             =   1636
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   2115
         TabIndex        =   6
         Top             =   1940
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   6
         Left            =   2115
         TabIndex        =   7
         Top             =   2244
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   16
         Left            =   4845
         TabIndex        =   8
         Top             =   2244
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   2115
         TabIndex        =   0
         Top             =   420
         Width           =   735
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   2115
         TabIndex        =   1
         Top             =   724
         Width           =   1575
         VariousPropertyBits=   671105051
         MaxLength       =   12
         Size            =   "2778;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   2
         Left            =   2115
         TabIndex        =   2
         Top             =   1028
         Width           =   735
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblName 
         Height          =   255
         Index           =   69
         Left            =   7560
         TabIndex        =   87
         Top             =   5003
         Width           =   1170
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "2064;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "國外部小組/專利處組別"
         Height          =   180
         Index           =   4
         Left            =   6570
         TabIndex        =   85
         Top             =   4736
         Width           =   1845
      End
      Begin VB.Label Label23 
         Caption         =   "外翻人員用來記錄聯絡人(含稱謂)"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -69960
         TabIndex        =   84
         Top             =   2618
         Width           =   4155
      End
      Begin VB.Label Label22 
         Caption         =   "專業代號 －對開拓"
         Height          =   195
         Left            =   6060
         TabIndex        =   83
         Top             =   2297
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "多人時以;區隔"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   7
         Left            =   7620
         TabIndex        =   82
         Top             =   4125
         Width           =   1125
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "定稿智權人員區別："
         Height          =   180
         Left            =   -74775
         TabIndex        =   80
         Top             =   2970
         Width           =   1620
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "定稿智權人員名稱："
         Height          =   180
         Left            =   -74775
         TabIndex        =   79
         Top             =   2640
         Width           =   1620
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "是否設定英核表       (N:未設定但固定由英文顧問做英核"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   4440
         TabIndex        =   78
         Top             =   2880
         Width           =   4360
      End
      Begin MSForms.Label lblName 
         Height          =   260
         Index           =   61
         Left            =   2880
         TabIndex        =   77
         Top             =   2570
         Width           =   1280
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "2258;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label9 
         Caption         =   "草圖核稿人"
         Height          =   180
         Left            =   540
         TabIndex        =   76
         Top             =   2608
         Width           =   1095
      End
      Begin MSForms.Label lblName 
         Height          =   260
         Index           =   62
         Left            =   2880
         TabIndex        =   75
         Top             =   2880
         Width           =   1280
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "2258;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label10 
         Caption         =   "繪圖判發人"
         Height          =   180
         Left            =   540
         TabIndex        =   74
         Top             =   2912
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "專利處是否承辦日文案件        (Y:是)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   4260
         TabIndex        =   73
         Top             =   2610
         Width           =   2810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "外商FCT承辦要加入非業務但有收文檔(nsbhc)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   15
         Left            =   3840
         TabIndex        =   72
         Top             =   720
         Width           =   3540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "打卡異常郵件收件人："
         Height          =   180
         Index           =   3
         Left            =   3660
         TabIndex        =   71
         Top             =   3824
         Width           =   1800
      End
      Begin MSForms.Label lblName 
         Height          =   255
         Index           =   58
         Left            =   6240
         TabIndex        =   70
         Top             =   3787
         Width           =   1215
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "2143;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "是否可自動收文        (N:否)"
         Height          =   180
         Left            =   5265
         TabIndex        =   69
         Top             =   1088
         Width           =   2085
      End
      Begin MSForms.Label lblName 
         Height          =   255
         Index           =   56
         Left            =   -71910
         TabIndex        =   68
         Top             =   2243
         Width           =   1845
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "3254;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "外商FC案件程序管制人："
         Height          =   180
         Left            =   -74790
         TabIndex        =   67
         Top             =   2280
         Width           =   2025
      End
      Begin VB.Label Label12 
         Caption         =   "外部信箱"
         Height          =   195
         Left            =   540
         TabIndex        =   57
         Top             =   5033
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Group 別"
         Height          =   180
         Index           =   10
         Left            =   540
         TabIndex        =   56
         Top             =   3216
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "英文名"
         Height          =   180
         Left            =   540
         TabIndex        =   55
         Top             =   3520
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "員工到職日"
         Height          =   180
         Left            =   540
         TabIndex        =   54
         Top             =   3824
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "內部郵件收件員工編號"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   53
         Top             =   4128
         Width           =   1800
      End
      Begin MSForms.Label lblName 
         Height          =   255
         Index           =   13
         Left            =   4590
         TabIndex        =   52
         Top             =   4095
         Width           =   2865
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "5054;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "收文所屬部門"
         Height          =   180
         Index           =   1
         Left            =   540
         TabIndex        =   51
         Top             =   4432
         Width           =   1080
      End
      Begin MSForms.Label lblName 
         Height          =   255
         Index           =   14
         Left            =   2880
         TabIndex        =   50
         Top             =   4395
         Width           =   1155
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "2037;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblName 
         Height          =   255
         Index           =   15
         Left            =   2880
         TabIndex        =   49
         Top             =   4699
         Width           =   1170
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "2064;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "專利考核部門/國外部組別"
         Height          =   180
         Index           =   2
         Left            =   45
         TabIndex        =   48
         Top             =   4736
         Width           =   2025
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "舊制與工程師每週完稿明細用"
         Height          =   180
         Left            =   4080
         TabIndex        =   47
         Top             =   4736
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利處新工程師考核部門設為CFP"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   14
         Left            =   4080
         TabIndex        =   46
         Top             =   4432
         Width           =   2640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "在職、離職"
         Height          =   180
         Index           =   3
         Left            =   540
         TabIndex        =   45
         Top             =   1392
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "等級"
         Height          =   180
         Index           =   4
         Left            =   540
         TabIndex        =   44
         Top             =   1696
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工所屬所別                      ( 1.北 2.中 3.南 4.高 5.其他 )"
         Height          =   180
         Index           =   5
         Left            =   540
         TabIndex        =   43
         Top             =   2000
         Width           =   4215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專業代號 －對國內"
         Height          =   180
         Index           =   6
         Left            =   540
         TabIndex        =   42
         Top             =   2304
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(1:在職 2:離職)"
         Height          =   180
         Index           =   11
         Left            =   2595
         TabIndex        =   41
         Top             =   1392
         Width           =   1155
      End
      Begin MSForms.Label lblName 
         Height          =   255
         Index           =   4
         Left            =   2610
         TabIndex        =   40
         Top             =   1659
         Width           =   4815
         BackColor       =   16761024
         VariousPropertyBits=   27
         Size            =   "8493;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         Caption         =   "專業代號 －對國外"
         Height          =   195
         Left            =   3330
         TabIndex        =   39
         Top             =   2297
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "若為業務員離職,記得要刪隔月起之目標"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   12
         Left            =   4190
         TabIndex        =   38
         Top             =   1390
         Width           =   3100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號"
         Height          =   180
         Index           =   0
         Left            =   540
         TabIndex        =   37
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Index           =   1
         Left            =   540
         TabIndex        =   36
         Top             =   784
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工部門別"
         Height          =   180
         Index           =   2
         Left            =   540
         TabIndex        =   35
         Top             =   1088
         Width           =   900
      End
      Begin MSForms.Label lblName 
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   34
         Top             =   1051
         Width           =   1620
         BackColor       =   -2147483634
         VariousPropertyBits=   27
         Size            =   "2857;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利處試用期間工程師,國內專業代號設為 專99"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   13
         Left            =   3840
         TabIndex        =   33
         Top             =   480
         Width           =   3690
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1056
      Top             =   -96
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040105.frx":037A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040105.frx":0696
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040105.frx":09B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040105.frx":0B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040105.frx":0EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040105.frx":11C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040105.frx":14E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040105.frx":17FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040105.frx":1B1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040105.frx":1E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040105.frx":2152
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   8950
      _ExtentX        =   15787
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
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
   Begin VB.Label Label14 
      Caption         =   "內部信箱"
      Height          =   195
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   1035
   End
End
Attribute VB_Name = "frm12040105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/14 改成Form2.0 ; Text1(index)、lblName(index)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

'Remove by Morgan 2008/6/24
''edit by nick 2004/07/13 加欄位
''Const iFieldTotal = 15
'Const iFieldTotal = 18
'Dim TmpField(0 To iFieldTotal) As String
'end 2008/6/24
Dim RcMain As New ADODB.Recordset, cp As New ADODB.Recordset
Dim ActionEdit As Integer
Dim Bmk As Variant, i As Integer

' 90.07.16 modify by Ken (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim oText As Object, idx As Integer
Dim m_MeTrackMode  As String 'Added by Lydia 2021/10/20 Form2.0 記錄鍵盤傳入順序
'Dim bolMsgEnter As Boolean 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值; 因為MsgBox的Enter鍵都會觸發Toolbar的”確定KeyF9”動作反之用滑鼠就不會

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'Memo by Lydia 2021/10/20 原程式搬到Form_KeyUp
    
   'Remove by Lydia 2021/10/25
   'Call PUB_SaveMeTrackMode(m_MeTrackMode, 0, KeyCode)  'Added by Lydia 2021/10/20 Form2.0 記錄鍵盤傳入順序

End Sub

'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
     
'Remove by Lydia 2021/10/20 改到Form_KeyUp
'    Select Case KeyAscii
'      Case vbKeyReturn:
'         If ActionEdit <> 3 Then
'            'KeyAscii = 0
'            'Form_KeyDown vbKeyF9, 0
'         End If
'    End Select
End Sub

'Added by Lydia 2021/10/20
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Remove by Lydia 2021/10/25
    'Call PUB_SaveMeTrackMode(m_MeTrackMode, 1, KeyCode)  'Added by Lydia 2021/10/20 Form2.0 記錄鍵盤傳入順序
    
'Memo by Lydia 2021/10/20 從Form_KeyDown搬來
Select Case KeyCode
        Case vbKeyF2
        Text1(0).SetFocus
        RcEdit 0
        Case vbKeyF3
        Text1(1).SetFocus
        Text1(0).TabStop = False
        RcEdit 1
        Case vbKeyF5
        RcEdit 2
        Case vbKeyF4
        RcEdit 5
        Case vbKeyHome
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                ActionRc 0
             End If
        Case vbKeyPageUp
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                 ActionRc 1
             End If
        Case vbKeyPageDown
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                 ActionRc 2
             End If
        Case vbKeyEnd
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                  ActionRc 3
             End If
        Case vbKeyF9
         'Modified by Morgan 2019/3/11 變的有點慢,改資料抓法
         'If Text1(0) = "" Then MsgBox "員工代號不可為空值", vbInformation: Text1(0).SetFocus: Exit Sub
         'If ActionEdit = 0 Or ActionEdit = 1 Then
         'If Text1(1) = "" Then MsgBox "姓名", vbInformation: Text1(1).SetFocus: Exit Sub
         'If Text1(2) = "" Then MsgBox "員工部門別不可為空值", vbInformation: Text1(2).SetFocus: Exit Sub
         'If Text1(3) = "" Then MsgBox "在職，離職不可為空值", vbInformation: Text1(3).SetFocus: Exit Sub
         ''If Text1(4) = "" Then MsgBox "等級不可為空值", vbInformation: Text1(4).SetFocus: Exit Sub
         'If Text1(5) = "" Then MsgBox "員工所屬所別不可為空值", vbInformation: Text1(5).SetFocus: Exit Sub
         'If Text1(14) = "" Then Text1(14) = Text1(2)
         ''If Text1(10) = "" Then MsgBox "Group別不可為空值", vbInformation: Text1(10).SetFocus: Exit Sub
         'End If
         'RcEdit 3
         'Text1(0).TabStop = True
         'RcMain.ReQuery
         'RcMain.Find "st01='" & Text1(0) & "'", 0, adSearchForward, 1
         OnAction vbKeyF9
         'end 2019/3/11
        
        Case vbKeyF10
         RcEdit 4
        
        'Added by Lydia 2021/10/20 從Form_KeyPress改到這裡
        'Remove by Lydia 2021/11/22 取消以ENTER控制為換行的功能 (Form2.0修改之維護資料功能Toolbar之修改統一)
        'Case vbKeyReturn:
        '   If ActionEdit <> 3 Then
        '      'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值; 因為MsgBox的Enter鍵都會觸發Toolbar的”確定KeyF9”動作
        '      If bolMsgEnter = True Then
        '          bolMsgEnter = False
        '      Else
        '      'end 2021/10/25
        '          OnAction vbKeyF9
        '      End If 'Added by Lydia 2021/10/25
        '   End If
        ''end 2021/10/20
        'end 2021/11/22
        
        Case vbKeyEscape
        Unload Me
        Set frm12040105 = Nothing
End Select
'   ' Ken 90.07.19 -- Start
'   If KeyCode <> vbKeyF2 And KeyCode <> vbKeyF3 And KeyCode <> vbKeyF4 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyEscape Then
'         If m_bInsert Then
'             TBar1.Buttons(1).Enabled = True
'         Else
'             TBar1.Buttons(1).Enabled = False
'         End If
'         If m_bUpdate Then
'             TBar1.Buttons(2).Enabled = True
'         Else
'             TBar1.Buttons(2).Enabled = False
'         End If
'         If m_bDelete Then
'             TBar1.Buttons(3).Enabled = True
'         Else
'             TBar1.Buttons(3).Enabled = False
'         End If
'   End If
'   ' Ken 90.07.19 -- End
End Sub

Private Sub Form_Load()
   ' 90.07.16 modify by Ken (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040105", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040105", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040105", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040105", strFind, False)
   ' Ken 90.07.16 -- End
   
   MoveFormToCenter Me
   cp.CursorLocation = adUseClient
   
'Modified by Morgan 2019/3/11 變的有點慢,改資料抓法
'   If RcMain.State = adStateOpen Then RcMain.Close
'   'edit by nick 2004/07/13 加欄位
'   'strExc(0) = "SELECT ST01,ST02,ST03,ST04,ST05,ST06,ST07,ST08,ST09,ST10,ST11," & _
'      "ST12,ST13,ST14,ST15,ST16 FROM STAFF ORDER BY ST01,ST02,ST03"
'   'Modify by Morgan 2008/6/24
'   'strExc(0) = "SELECT ST01,ST02,ST03,ST04,ST05,ST06,ST07,ST08,ST09,ST10,ST11," & _
'      "ST12,ST13,ST14,ST15,ST16,ST17,ST18,ST19 FROM STAFF ORDER BY ST01,ST02,ST03"
'   strExc(0) = "SELECT * FROM STAFF ORDER BY ST01,ST02,ST03"
'   RcMain.CursorType = adOpenDynamic
'   RcMain.CursorLocation = adUseClient
'   RcMain.LockType = adLockBatchOptimistic
'   RcMain.Open strExc(0), cnnConnection
'   If Not RcMain.BOF Then ActionRc 0
   ActionRc 0
'END 2019/3/11
   
   TxtSitu True
   ActionEdit = 3
   
   ' Ken 90.07.16 -- start
   If m_bInsert Then
       TBar1.Buttons(1).Enabled = True
   Else
       TBar1.Buttons(1).Enabled = False
   End If
   If m_bUpdate Then
       TBar1.Buttons(2).Enabled = True
   Else
       TBar1.Buttons(2).Enabled = False
   End If
   If m_bDelete Then
       TBar1.Buttons(3).Enabled = True
   Else
       TBar1.Buttons(3).Enabled = False
   End If
   
   Dim objLbl As Object
   For Each objLbl In lblName
      objLbl.BackColor = &H8000000F
   Next
   
   SSTab1.Tab = 0 'Added by Lydia 2017/03/07
End Sub

Private Sub ActionRc(ByVal Sty As Integer)
   
'Modified by Morgan 2019/3/11 變的有點慢，改一次只要抓一筆
'TxtLock 2
'   If RcMain.EOF And RcMain.BOF Then MsgBox "資料庫內無資料 !", vbInformation: Exit Sub
'   With RcMain
'      Select Case Sty
'         Case 0
'           .MoveFirst
'         Case 1
'               .MovePrevious
'            If .BOF Then
'               Beep
'               DataErrorMessage (6)
'               .MoveFirst
'            End If
'         Case 2
'               .MoveNext
'            If .EOF Then
'               Beep
'               DataErrorMessage (7)
'               .MoveLast
'            End If
'         Case 3
'            .MoveLast
'      End Select
'   End With
'   SetTxtValue
   Select Case Sty
   Case 0 '第一筆
      strExc(0) = "SELECT * FROM STAFF WHERE ST01=(SELECT Min(B.ST01) FROM STAFF B)"
   Case 1 '上一筆
      strExc(0) = "SELECT * FROM STAFF WHERE ST01=(SELECT Max(B.ST01) FROM STAFF B WHERE B.ST01<'" & Text1(0) & "')"
   Case 2 '下一筆
      strExc(0) = "SELECT * FROM STAFF WHERE ST01=(SELECT Min(B.ST01) FROM STAFF B WHERE B.ST01>'" & Text1(0) & "')"
   Case 3 '最後筆
      strExc(0) = "SELECT * FROM STAFF WHERE ST01=(SELECT Max(B.ST01) FROM STAFF B)"
   Case 4 '查詢
      strExc(0) = "SELECT * FROM STAFF WHERE ST01='" & Text1(0) & "'"
   End Select
   intI = 1

   If RcMain.State = adStateOpen Then RcMain.Close
   RcMain.CursorType = adOpenDynamic
   RcMain.CursorLocation = adUseClient
   RcMain.LockType = adLockBatchOptimistic
   RcMain.Open strExc(0), cnnConnection
   If Not RcMain.BOF Then
      TxtLock 2
      SetTxtValue
      'Modified by Lydia 2025/08/14
      'Text1(0).Tag = Text1(0)
      'Text1(14).Tag = Text1(14)   'add by sonia 2022/5/9
      ''add by sonia 2022/5/24
      'Text1(51).Tag = Text1(51)
      'Text1(52).Tag = Text1(52)
      'Text1(53).Tag = Text1(53)
      'Text1(54).Tag = Text1(54)
      'end 2022/5/24
      For Each oText In Text1
         oText.Tag = oText.Text
      Next
      'end 2025/08/14
   Else
      If Sty = 4 Then
         Beep
         'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
         MsgBox "查無資料！", vbInformation
         
         '帶前次查詢資料,沒有則帶最後筆
         If Text1(0).Tag <> Text1(0) Then
            If Text1(0).Tag <> "" Then
               Text1(0) = Text1(0).Tag
               ActionRc 4
            Else
               ActionRc 3
            End If
         End If
         
      ElseIf Sty = 1 Then
         Beep
         DataErrorMessage (6)
         ActionRc 0
         
      ElseIf Sty = 2 Then
         Beep
         DataErrorMessage (7)
         ActionRc 3
         
      End If
   End If
'END 2019/3/11
End Sub

'Modify by Morgan 2008/6/24 改寫法
Private Sub SetTxtValue()
   For Each oText In Text1
      idx = oText.Index
      'Modified by Lydia 2024/01/04
      'If IsNull(RcMain(idx)) Then
      If IsNull(RcMain.Fields("ST" & Format(idx + 1, "00"))) Then
         oText = ""
      Else
         If idx = 12 Then
            'Modified by Lydia 2024/01/04
            'oText = ChangeWStringToTString(RcMain(idx))
            oText = ChangeWStringToTString(RcMain.Fields("ST" & Format(idx + 1, "00")))
         Else
            'Modified by Lydia 2024/01/04
            'oText = RcMain(idx)
            oText = "" & RcMain.Fields("ST" & Format(idx + 1, "00"))
            'Add By Sindy 2024/1/9
            If idx + 1 = 3 Then '=ST03
               If Left(oText, 2) = "F2" Then
                  Label8.Caption = "外專歷程判發為三級主管        (Y:是)"
               Else
                  Label8.Caption = "專利處是否承辦日文案件        (Y:是)"
               End If
            End If
            '2024/1/9 END
            Select Case idx
               Case 2, 14
                  lblName(idx).Caption = ChgType(0, oText.Text)
               'Added by Lydia 2017/03/07 +ST14
               Case 13
                  lblName(idx).Caption = ChgType(2, oText.Text)
               'end 2017/03/07
               Case 15
                  '2010/1/8 MODIFY BY SONIA
                  'Select Case oText.Text
                  '   'Modify by Morgan 2005/5/25 外專工程師組別
                  '   Case "1", "2", "3", "4"
                  '      'Modify by Morgan 2008/1/4 加外翻F51,內翻F52  2008/4/8 加F81
                  '      If Text1(2) = "F21" Or Text1(2) = "F51" Or Text1(2) = "F52" Or Text1(2) = "F81" Then
                  '         lblName(idx) = PUB_GetFCPGrpName(oText.Text)
                  '      ElseIf Mid(Text1(2), 1, 2) = "F1" Then
                  '         lblName(idx) = PUB_GetFCTGrpName(oText.Text)
                  '      End If
                  '   Case Else
                  '      lblName(idx).Caption = ChgType(1, oText.Text)
                  'End Select
                  If Text1(2) = "F21" Or Text1(2) = "F51" Or Text1(2) = "F52" Or Text1(2) = "F81" Then
                     lblName(idx) = PUB_GetFCPGrpName(oText.Text, True)  '2010年起何季陵改第5組其他
                  ElseIf Mid(Text1(2), 1, 2) = "F1" Then
                     'Mark by Amy 2021/04/12 名稱改至lblName(69)顯示
                     'lblName(idx) = PUB_GetFCTGrpName(oText.Text)
                  '2015/12/21 add by sonia 加入外法組別控制,可不輸入
                  'Modify By Sindy 2019/8/5 + Or Text1(2) = "F23"
                  'modify by sonia 2020/4/3 +L01
                  ElseIf (Text1(2) = "F31" Or Text1(2) = "L02" Or Text1(2) = "L01" Or Text1(2) = "F23") And oText.Text <> "" Then
                     lblName(15).Caption = PUB_GetFCLGrpName(oText.Text)
                  'add by sonia 2023/12/18加F62英文顧問、F71日文顧問、F72德文顧問
                  ElseIf Text1(2) = "F62" Or Text1(2) = "F71" Or Text1(2) = "F72" Then
                     lblName(15) = PUB_GetFEMPGrpName(oText.Text)
                  'end 2023/12/18
                  Else
                  '2015/12/21 end
                     lblName(idx).Caption = ChgType(1, oText.Text)
                  End If
                  '2010/1/8 END
                  
               'Added by Morgan 2019/3/12
               Case 69
                   'Add by Amy 2021/04/12 +外商組別 原於lblName(15)顯示
                  If Mid(Text1(2), 1, 2) = "F1" Then
                     lblName(69) = PUB_GetFCTGrpName(Text1(15), Text1(69))
                  ElseIf Text1(2) = "P10" Or Text1(2) = "P11" Then
                     lblName(69) = PUB_CST70(Text1(69), Text1(2))
                  End If
               'end 2019/3/12
               'end 2019/3/12
               'Added by Lydia 2024/01/04
               Case 92  '新部門別ST93
                  lblName(idx).Caption = ChgType(3, oText.Text)
               'end 2024/01/04
            End Select
         End If
      End If
   Next
End Sub

Private Sub RcEdit(Situ As Integer)
Dim i As Integer
Dim m_FnoMSG As String, m_MailText As String 'Add By Sindy 2019/8/16
   
   Select Case Situ
      Case 0 'add
         TxtClear
         TxtSitu False
         ActionEdit = 0
         TextInverse Text1(0)
         
      Case 1 'modi
         TxtSitu False
         ActionEdit = 1
         'Remove by Morgan 2008/6/24
         'For i = 0 To iFieldTotal
         '   If i = 12 Then
         '   TmpField(i) = ChangeTStringToWString(Text1(i))
         '   Else
         '   TmpField(i) = Text1(i).Text
         '   End If
         'Next
         Text1(0).Locked = True
         
      Case 2 'delete
         'Added by Lydia 2021/10/20 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         'Remove by Lydia 2021/10/25
         'If PUB_ChkMeTrackMode(m_MeTrackMode) = False Then
         '    Exit Sub
         'End If
         ''end 2021/10/20
         'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
         If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            'Add By Sindy 2019/8/16 人員刪除時要檢查的資料
            m_FnoMSG = StaffLeaveChkData(Text1(0).Text, "03離職")
            m_MailText = "員工編號：" + Text1(0) + " " + Text1(1) & vbCrLf & vbCrLf & m_FnoMSG
            
           'Added by Lydia 2025/08/14 利益衝突案件：檢查利益衝突案件之權限，若人事有異動留下相關記錄
            Call PUB_SaveCUFA_Staff_Log(True, Text1(0), "02", Me.Name, pub_HostName)
            Call PUB_SendMailCache
            'end 2025/08/14
            RcMain.Delete
            RcMain.UpdateBatch
            
            'Add By Sindy 2019/8/16 發E-Mail通知
            If m_FnoMSG <> "" Then
               'Modified by Lydia 2025/08/14 "83002"=>Pub_GetSpecMan("程式管理人員")
               PUB_SendMail strUserNum, Pub_GetSpecMan("程式管理人員"), "", "員工刪除,異動資料通知！", m_MailText & vbCrLf
            End If
            
            'Modified by Morgan 2019/3/11 變的有點慢,改資料抓法
            'If RcMain.EOF = True Then
            '   ActionRc 1
            'Else
            '   ActionRc 2
            'End If
            Text1(0).Tag = ""
            ActionRc 2
            'end 2019/3/11
         End If
      Case 3 'update
         If ActionEdit = 0 Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
         
            RcMain.AddNew
            If GetVal = False Then Exit Sub
            
            'Modified by Morgan 2019/3/11 新增完應該停在該筆才對
            'ActionRc 3
            ActionRc 4
            'end 2019/3/11
         ElseIf ActionEdit = 1 Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            If GetVal = False Then Exit Sub
         ElseIf ActionEdit = 2 Then
         
'Modified by Morgan 2019/3/11 變的有點慢,改資料抓法
'            RcMain.Find "ST01='" & Text1(0).Text & "'", 0, adSearchForward, 1
'            If RcMain.EOF Then
'               MsgBox "無此記錄之資料 !", vbCritical
'               RcMain.Bookmark = Bmk
'           ' Else
'            '   RcMain.Find "ST02='" & Text1(1).Text & "'", 0, adSearchForward, RcMain.Bookmark
'             '  If RcMain.EOF Then
'              '    MsgBox "無此記錄之資料 !", vbCritical
'               '   RcMain.Bookmark = Bmk
'              ' Else
'              '    RcMain.Find "ST03='" & Text1(2).Text & "'", 0, adSearchForward, RcMain.Bookmark
'              '    If RcMain.EOF Then
'              '       MsgBox "無此記錄之資料 !", vbCritical
'              '       RcMain.Bookmark = Bmk
'              '    End If
'           '    End If
'            End If
'            SetTxtValue
            ActionRc 4
'end 2019/3/11

         End If
         TxtSitu True
         ActionEdit = 3
      Case 4 'cancel
         'Added by Lydia 2021/10/20 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         'Remove by Lydia 2021/10/25
         'If PUB_ChkMeTrackMode(m_MeTrackMode) = False Then
         '    Exit Sub
         'End If
         ''end 2021/10/20
         'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
         If MsgBox("妳並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
            TxtSitu True
            If ActionEdit = 0 Then
               'Modified by Morgan 2019/3/11 新增取消應該停在前次那一筆才對
               'ActionRc 3
               Text1(0) = Text1(0).Tag
               ActionRc 4
               'end 2019/3/11
            ElseIf ActionEdit = 1 Then
               'Modify by Morgan 2008/6/24
               'For i = 0 To iFieldTotal
               '   If i = 12 Then
               '   Text1(i).Text = ChangeWStringToTString(TmpField(i))
               '   Else
               '   Text1(i).Text = TmpField(i)
               '   End If
               'Next
               ReadText
               'end 2008/6/24
            ElseIf ActionEdit = 2 Then
               'Modified by Morgan 2019/3/11 變的有點慢,改資料抓法
               'RcMain.Bookmark = Bmk
               Text1(0) = Text1(0).Tag
               ActionRc 4
               'end 2019/3/11
               SetTxtValue
            End If
            ActionEdit = 3
         Else
            Exit Sub
         End If
      Case 5 'query
         'Bmk = RcMain.Bookmark 'Removed by Morgan 2019/3/11 改資料抓法用不到了
         TxtSitu False
         TxtLock 2
         ActionEdit = 2
         Text1(0).Locked = False
         Text1(0).SetFocus
   End Select
End Sub

Private Function GetVal() As Boolean
   Dim i As Integer
   
On Error GoTo ErrHand
   
   'Modify by Morgan 2008/6/24
   'For i = 0 To iFieldTotal
   '   If i = 12 Then
   '   If Text1(i).Text <> "" Then
   '      RcMain.Fields(i).Value = ChangeTStringToWString(Text1(i).Text)
   '   Else
   '      RcMain.Fields(i).Value = Null
   '   End If
   '   Else
   '   If Text1(i).Text <> "" Then
   '      RcMain.Fields(i).Value = Text1(i).Text
   '   Else
   '      RcMain.Fields(i).Value = Null
   '   End If
   '   End If
   'Next
   For Each oText In Text1
      idx = oText.Index
      If oText = "" Then
         'Modified by Lydia 2024/01/04
         'RcMain.Fields(idx).Value = Null
         RcMain.Fields("ST" & Format(idx + 1, "00")) = Null
      Else
         If idx = 12 Then
            'Modified by Lydia 2024/01/04
            'RcMain.Fields(idx).Value = ChangeTStringToWString(oText.Text)
            RcMain.Fields("ST" & Format(idx + 1, "00")).Value = ChangeTStringToWString(oText.Text)
         Else
            'Modified by Lydia 2024/01/04
            'RcMain.Fields(idx).Value = oText
            RcMain.Fields("ST" & Format(idx + 1, "00")).Value = oText
         End If
      End If
   Next
   'end 2008/6/24
   
   'Add by Morgan 2008/12/29
   '若離職但無離職日時自動上系統日
   'Modified by Lydia 2024/01/04
   'If RcMain.Fields(3) = "2" And IsNull(RcMain.Fields(50)) Then
   '   RcMain.Fields(50) = strSrvDate(1)
   If RcMain.Fields("ST04") = "2" And IsNull(RcMain.Fields("ST51")) Then
      RcMain.Fields("ST51") = strSrvDate(1)
   End If
   
   RcMain.UpdateBatch
   GetVal = True
   'Added by Lydia 2025/08/14 利益衝突案件：檢查利益衝突案件之權限，若人事有異動留下相關記錄
   strSql = ""
   If ActionEdit = 0 Then '新增
      strSql = "01" '新增
   ElseIf ActionEdit = 1 Then
      If Text1(3).Tag <> Text1(3) Then
         If Text1(3) = "1" Then
            strSql = "01"
         Else
            strSql = "02" '離職
         End If
      Else
         If Text1(3) = "2" Then
            strSql = ""
         Else
            strSql = "03"  '修改其他
         End If
      End If
   End If
   If strSql <> "" Then
      Call PUB_SaveCUFA_Staff_Log(True, Text1(0), strSql, Me.Name, pub_HostName, Text1(2), Text1(2).Tag, Text1(4), Text1(4).Tag, Text1(15), Text1(15).Tag, Text1(69), Text1(69).Tag)
      Call PUB_SendMailCache
   End If
   
   For Each oText In Text1
      oText.Tag = oText.Text
   Next
   'end 2025/08/14
   
   Exit Function
ErrHand:
   GetVal = False
   RcMain.CancelUpdate
   RcMain.ReQuery
   If Err.Number = -2147217887 Then
      MsgBox "資料錯誤，新增失敗 !", vbInformation
   Else
      MsgBox "錯誤 : " & Err.Description, vbInformation
   End If
End Function

Private Sub TxtLock(ByVal Lt As Integer)
 Dim txt As Object, i As Integer
   Select Case Lt
      Case 0
         For Each txt In frm12040105.Text1
            txt.Locked = True
         Next
         
      Case 1
         For Each txt In frm12040105.Text1
            txt.Locked = False
         Next
         
      Case 2
         For Each txt In frm12040105.Text1
            txt.Locked = True
         Next
         TxtClear
   End Select
End Sub

Private Sub TxtClear()
   Dim txt As Object, Lbl As Object
   For Each txt In frm12040105.Text1
      txt.Text = ""
   Next
   For Each Lbl In lblName
      Lbl = ""
   Next
End Sub

Private Sub TxtSitu(ByVal TF As Boolean)
 Dim i As Integer, txt As Object
   If TF = True Then
      TxtLock 0
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = True
         TBar1.Buttons(i + 5).Enabled = True
      Next
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
   Else
      TxtLock 1
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   'Add By Cheng 2002/07/18
   Set frm12040105 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 'Remove by Lydia 2021/10/25
 'Call Pub_SaveMeToolBar(m_MeTrackMode, Me.TBar1, Button.Index) 'Added by Lydia 2021/10/20 若有交錯使用Function鍵和Toolbar鍵會失去記錄造成無法判斷，所以ToolBar鍵另外記錄
   
 Select Case Button.Index
      Case 1
         Text1(0).SetFocus
         RcEdit 0
      Case 2
         Text1(1).SetFocus
         Text1(0).TabStop = False
         RcEdit 1
      Case 3
         RcEdit 2
      Case 4
         RcEdit 5
      Case 6
         ActionRc 0
      Case 7
         ActionRc 1
      Case 8
         ActionRc 2
      Case 9
         ActionRc 3
      Case 11
        'Modified by Morgan 2019/3/11 程式重複,改寫函數,
        'If Text1(0) = "" Then MsgBox "員工代號不可為空值", vbInformation: Text1(0).SetFocus: Exit Sub
        'If ActionEdit = 0 Or ActionEdit = 1 Then
        ''If Text1(1) = "" Then MsgBox "姓名不可為空值", vbInformation: Text1(1).SetFocus: Exit Sub
        ''If Text1(2) = "" Then MsgBox "員工部門別不可為空值", vbInformation: Text1(2).SetFocus: Exit Sub
        'If Text1(3) = "" Then MsgBox "在職，離職不可為空值", vbInformation: Text1(3).SetFocus: Exit Sub
        ''If Text1(4) = "" Then MsgBox "等級不可為空值", vbInformation: Text1(4).SetFocus: Exit Sub
        'If Text1(5) = "" Then MsgBox "員工所屬所別不可為空值", vbInformation: Text1(5).SetFocus: Exit Sub
        ''If Text1(10) = "" Then MsgBox "Group別不可為空值", vbInformation: Text1(10).SetFocus: Exit Sub
        'End If
        'RcEdit 3
        'Text1(0).TabStop = True
        'RcMain.ReQuery
        'RcMain.Find "st01='" & Text1(0) & "'", 0, adSearchForward, 1
        OnAction vbKeyF9
        'end 2019/3/11
      Case 12
        RcEdit 4
      Case 14
         Unload Me
         Set frm12040105 = Nothing
   End Select
   
If ActionEdit = 3 Then 'Addded by Morgan 2019/3/12 恢復瀏覽模式才要重設,因若檢查失敗時仍要維持當下的狀態
   ' Ken 90.07.16 -- Start
   If Button.Index <> 14 And Button.Index <> 1 And Button.Index <> 2 And Button.Index <> 3 And Button.Index <> 4 Then
      If m_bInsert Then
          TBar1.Buttons(1).Enabled = True
      Else
          TBar1.Buttons(1).Enabled = False
      End If
      If m_bUpdate Then
          TBar1.Buttons(2).Enabled = True
      Else
          TBar1.Buttons(2).Enabled = False
      End If
      If m_bDelete Then
          TBar1.Buttons(3).Enabled = True
      Else
          TBar1.Buttons(3).Enabled = False
      End If
   End If
   ' Ken 90.07.16 -- End
End If 'Addded by Morgan 2019/3/12

End Sub

Private Sub Text1_Change(Index As Integer)
   Select Case Index
       Case 4
         If cp.State = adStateOpen Then cp.Close
         strExc(1) = "select sl02 from staff_level where sl01='" & Text1(4) & "'"
         cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
         If cp.BOF And cp.EOF Then
            lblName(4).Caption = ""
         Else
            If IsNull(cp.Fields(0).Value) Then
                lblName(4).Caption = ""
            Else
                lblName(4).Caption = cp.Fields(0).Value
            End If
         End If
        cp.Close
      
      'Modify By Sindy 2013/7/25 +58
      'Modify By Sindy 2015/4/20 +, 61, 62
      'Modified by Lydia 2017/03/07 -13
      Case 51, 52, 53, 54, 56, 58, 61, 62 'Modify By Sindy 2009/09/10
         'Modified by Morgan 2016/3/16 改3碼以上(原5) Ex.M12 接待室
         If Len(Text1(Index)) >= 3 Then
            lblName(Index).Caption = GetStaffName(Text1(Index), True)
         Else
            lblName(Index).Caption = ""
         End If
   End Select
   
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   Select Case Index
      Case 1, 6, 16, 68
        'edit by nickc 2007/07/11 切換輸入法改用API
        'Text1(Index).IMEMode = 1
        OpenIme
      Case Else
        'edit by nickc 2007/07/11 切換輸入法改用API
        'Text1(Index).IMEMode = 2
        CloseIme
   End Select
End Sub

'Modified by Lydia 2021/10/14 改成Form 2.0
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      '93.5.24 MODIFY BY SONIA 因有內外翻需輸入 Fxxxx
      'Case 0, 1, 2, 14, 15
      'Modified by Lydia 2024/01/04 +92 新部門別ST93
      Case 0, 10, 4, 13, 51, 52, 53, 54, 56, 92
         KeyAscii = UpperCase(KeyAscii)
      Case 1, 2, 14, 15
      '93.5.24 END
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii = 13 And ActionEdit = 2 Then
            RcEdit 3
         End If
      'Case 5
      '   If (KeyAscii > 53 Or KeyAscii < 49) And KeyAscii <> 8 Then
      '      KeyAscii = 0
      '      Beep
      '   End If
      'Modify By Sindy 2015/4/23 +, 63
      Case 57, 63 'Add by Morgan 2011/1/10
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
      Case 60 'Add By Sindy 2015/3/13
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
            Beep
         End If
      Case 69 'Add By Sindy 2018/11/7
         KeyAscii = UpperCase(KeyAscii)
         'Added by Morgan 2019/3/8
         If Text1(2) = "P10" Or Text1(2) = "P11" Then
            If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
               KeyAscii = 0
               Beep
            End If
         'end 2019/3/8
         'modify by sonia +5
         ElseIf KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> Asc("4") And KeyAscii <> Asc("5") Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTmp As String, i As Integer
   If ActionEdit = 3 Then Exit Sub
   If Index = 15 Then lblName(15).Caption = ""
   If Index = 92 Then lblName(92).Caption = ""  'Added by Lydia 2024/01/04
   If Text1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0
         If ActionEdit = 0 Then
            If cp.State = adStateOpen Then cp.Close
            strExc(1) = "select count(st01) from staff where st01='" & Text1(0) & "'"
            cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
            If cp.Fields(0).Value <> "0" Then
               'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
               MsgBox "此員工代號己存在"
               Cancel = True
            Else
               Cancel = False
            End If
            cp.Close
         End If
      Case 1
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(1).IMEMode = 2
         CloseIme
      Case 2
         If Text1(Index) = "" Then lblName(2) = "": Exit Sub
         lblName(2).Caption = ChgType(0, Text1(Index).Text)
         If lblName(2).Caption = "" Then Cancel = True
      Case 14
         If Text1(Index) = "" Then lblName(14) = "": Exit Sub
         lblName(14).Caption = ChgType(0, Text1(Index).Text)
         If lblName(14).Caption = "" Then Cancel = True
      Case 3
         If Not (Text1(3).Text = "1" Or Text1(3).Text = "2") Then
            'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
            MsgBox "輸入錯誤"
            Cancel = True
         Else
            Cancel = False
         End If
     Case 4
         If cp.State = adStateOpen Then cp.Close
         strExc(1) = "select sl02 from staff_level where sl01='" & Text1(4) & "'"
         cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
         If cp.BOF And cp.EOF And Text1(4) <> "" Then
            'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
            MsgBox "輸入錯誤"
            Cancel = True
         Else
            Cancel = False
            lblName(4).Caption = ""
            If Text1(4) <> "" Then
                lblName(4).Caption = cp.Fields(0).Value
            End If
         End If
         cp.Close
     Case 5
         If Not (Text1(5).Text >= "1" And Text1(5) <= "5") Then
            'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
            MsgBox "員工所屬所別只可輸入1~5", vbInformation
            Cancel = True
         Else
            Cancel = False
         End If
      Case 6, 16, 68
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(6).IMEMode = 2
         CloseIme
         If CheckLengthIsOK(Text1(Index), Text1(Index).MaxLength) = False Then
            Cancel = True
         Else
            Cancel = False
         End If
         
      'Add By Sindy 2018/11/7
      Case 69
         'Added by Morgan 2019/3/8
         '專利處組別
         lblName(69) = ""
         'Add by Amy 2021/04/12 +外商組別 原於lblName(15)顯示
         If Mid(Text1(2), 1, 2) = "F1" Then
             lblName(69) = PUB_GetFCTGrpName(Text1(15), Text1(Index))
         ElseIf Text1(2) = "P10" Or Text1(2) = "P11" Then
            lblName(69) = PUB_CST70(Text1(69), Text1(2))
         'end 2019/3/8
         ElseIf Not (Text1(69).Text = "1" Or Text1(69).Text = "2" Or Text1(69).Text = "3" Or Text1(69).Text = "4" Or Text1(69).Text = "5") Then
            'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
            MsgBox "輸入錯誤"
            Cancel = True
         Else
            Cancel = False
         End If
      Case 10
        If cp.State = adStateOpen Then cp.Close
        strExc(1) = "select count(sg01) from staff_group where sg01='" & Text1(10) & "'"
        cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
        If cp.Fields(0) = "0" And Text1(10) <> "" Then
            'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
            MsgBox "此Group別不存在"
            Cancel = True
        Else
            Cancel = False
        End If
'Modify by Morgan 2006/7/6 改放英文名
'      Case 11
'         If cp.State = adStateOpen Then cp.Close
'         strExc(1) = "select st02 from staff where st01='" & Text1(11) & "'"
'         cp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
'         If cp.EOF And cp.BOF And Text1(11) <> "" Then
'            MsgBox "此帶人主管不存在"
'            Cancel = True
'         Else
'            Cancel = False
'            If IsNull(cp.Fields(0)) Or cp.BOF Or cp.EOF Then
'                Label6.Caption = ""
'            Else
'                Label6.Caption = cp.Fields(0).Value
'            End If
'         End If
'         cp.Close
      Case 12
         If CheckIsTaiwanDate(Text1(12)) = False And Text1(12) <> "" Then
            Cancel = True
         Else
            Cancel = False
         End If
      'Added by Lydia 2017/03/07 ST14從v2(6)改為v2(20)
      Case 13
         If Text1(Index) = "" Then lblName(13) = "": Exit Sub
         Text1(Index).Text = Replace(Text1(Index).Text, ",", ";")
         lblName(13).Caption = ChgType(2, Text1(Index).Text)
         If lblName(13).Caption = "" Then Cancel = True
         'Add By Sindy 2018/5/28
         If InStr(Text1(Index), "99997") > 0 And Text1(58) = "" Then
            'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
            MsgBox "內部郵件不收信時，打卡異常郵件收件人欄，不可空白!!!"
            Text1(58).SetFocus 'Remove by Lydia 2021/10/25 造成重複彈訊息 'ReMark by Lydia 2021/11/22 恢復控制
            Cancel = True
         End If
         '2018/5/28 END
         
      'end 2017/03/07
      'end 2017/03/07
      'Modify By Sindy 2013/7/25 +58
      'Modified by Lydia 2017/03/07 - 13
      Case 51, 52, 53, 54, 56, 58 'Modify By Sindy 2009/09/10
         If Text1(Index) <> "" And lblName(Index) = "" Then
            Cancel = True
         End If
      Case 15
         'Modify by Morgan 2005/5/26 外專工程師為組別
         'Modify by Morgan 2008/1/4 內外翻也是  2008/4/8 加F81
         'If Text1(2) = "F21" Then
         If Text1(2) = "F21" Or Text1(2) = "F51" Or Text1(2) = "F52" Or Text1(2) = "F81" Then
            'Modify by Morgan 2007/5/30 改Call公用程式
            'Select Case Text1(Index)
            '   Case "1"
            '      lblName(15).Caption = "機電組"
            '   Case "2"
            '      lblName(15).Caption = "化學組"
            '   Case "3"
            '      lblName(15).Caption = "日文組"
            '   Case Else
            '      lblName(15).Caption = ""
            '      MsgBox "組別只可輸1,2,3！", vbCritical
            'End Select
            '2010/1/8 MODIFY BY SONIA 何季陵改第5組其他
            'lblName(15).Caption = PUB_GetFCPGrpName(Text1(Index))
            lblName(15).Caption = PUB_GetFCPGrpName(Text1(Index), True)
            If lblName(15).Caption = "" Then
               Cancel = True
            End If
            '2010/1/8 END
            'end 2007/5/30
         Else
            '2007/12/24 modify by sonia 加入外商承辦組別控制,只分組別無特定名稱
            'lblName(15).Caption = ChgType(1, Text1(Index).Text)
            If Mid(Text1(2), 1, 2) = "F1" Then
               'Mark by Amy 2021/04/12 名稱改至lblName(69)顯示
               'lblName(15).Caption = PUB_GetFCTGrpName(Text1(Index))
            '2015/12/21 add by sonia 加入外法組別控制,可不輸入
            'Modify By Sindy 2019/8/5 + Or Text1(2) = "F23"
            'modify by sonia 2020/3/30 +L01
            ElseIf (Text1(2) = "F31" Or Text1(2) = "L02" Or Text1(2) = "L01" Or Text1(2) = "F23") And Text1(Index) <> "" Then
               lblName(15).Caption = PUB_GetFCLGrpName(Text1(Index))
            'add by sonia 2023/12/18 加F62英文顧問、F71日文顧問、F72德文顧問
            ElseIf Text1(2) = "F62" Or Text1(2) = "F71" Or Text1(2) = "F72" Then
               lblName(15).Caption = PUB_GetFEMPGrpName(Text1(Index))
            'end 2023/12/18
            Else
            '2015/12/21 end
               lblName(15).Caption = ChgType(1, Text1(Index).Text)
            End If
            '2007/12/24 end
         End If
         '2007/12/24 modify by sonia
         'If lblName(15).Caption = "" Then Cancel = True
         '2009/8/17 CANCEL BY SONIA 改在TxtValidate檢查
         'If lblName(15).Caption = "" Then
         '   If MsgBox("專利考核部門/國外部組別錯誤或空白, 請確認是否儲存此欄位??", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         '      Text1(Index).Text = ""
         '      Exit Sub
         '   Else
         '      Cancel = True
         '   End If
         'End If
         '2009/8/17 END
         '2007/12/24 end
      'Added by Lydia 2024/01/04
      Case 92  '新部門別
         If Trim(Text1(Index)) = "" Then
            lblName(Index).Caption = ""
         Else
            lblName(Index).Caption = ChgType(3, Text1(Index))
            If lblName(Index).Caption = "" Then
               MsgBox "新部門別代號錯誤", vbCritical
               Cancel = True
            End If
         End If
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub

Private Function ChgType(ByVal Sty As Integer, ByVal txt As String) As String
 Dim strTmp As String
 'Added by Lydia 2017/03/07
 Dim tmpArr As Variant
 Dim inS As Integer

   Select Case Sty
      Case 0
         'edit by nickc 2007/02/09 不用 dll 了
         'If objLawDll.GetStaffDeptName(txt, strTmp) Then
         If ClsPDGetStaffDeptName(txt, strTmp) Then
            ChgType = strTmp
         Else
            ChgType = ""
         End If
      Case 1
         'edit by nickc 2007/02/09 不用 dll 了
         'If objLawDll.GetStaffDeptName(txt, strTmp) Then
         If ClsPDGetStaffDeptName(txt, strTmp) Then
            ChgType = strTmp
         Else
            ChgType = ""
         End If
      'Added by Lydia 2017/03/07
      Case 2  'ST14
        If Len(txt) > 0 Then
           tmpArr = Split(Replace(txt, ",", ";"), ";")
           strExc(0) = ""
           For inS = 0 To UBound(tmpArr)
              If Trim(tmpArr(inS)) <> "" Then
                 strExc(1) = GetStaffName(Trim(tmpArr(inS)), True)
                 If strExc(1) = "" Then
                    'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
                    MsgBox Trim(tmpArr(inS)) & "查無資料!"
                 Else
                    strExc(0) = strExc(0) & IIf(strExc(0) <> "", ";", "") & strExc(1)
                 End If
              End If
           Next inS
           ChgType = strExc(0)
        Else
           ChgType = ""
        End If
    'end 2017/03/07
      'Added by Lydia 2024/01/04
      Case 3
        ChgType = GetDeptA09(txt, "22", True)
      'end 2024/01/04
   End Select
End Function

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Text1
   If objTxt.Index <> 0 Then
      If objTxt.Enabled = True Then
         Cancel = False
         Text1_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   End If
Next

'add by sonia 2025/9/2
If (Text1(51) = "" And Text1(52) <> "") Or (Text1(52) = "" And Text1(53) <> "") Then
   MsgBox "第二~四級期限管制人請依序填寫，向前遞補！", vbCritical
   Exit Function
End If
'end 2025/9/2

'2008/11/5 add by sonia
If Text1(3).Text = "1" And (Mid(Text1(2), 1, 2) = "F2" Or Mid(Text1(2), 1, 2) = "F8") And Text1(0) <> "81040" Then
   If Text1(54) = "" Then
      Cancel = True
      'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
      MsgBox "外專員工請輸入第五級期限管制人！", vbCritical
      Exit Function
   End If
End If
If Text1(3).Text = "1" And (Text1(2) = "F21" Or Text1(2) = "F51" Or Text1(2) = "F52" Or Text1(2) = "F81" Or Mid(Text1(2), 1, 2) = "F1") Then
   'modify by sonia 2021/5/26 應改用lblName(69)判斷
   'If lblName(15).Caption = "" Then
   'modify by sonia 2021/5/26 改用lblName(15),lblName(69)都要判斷
   If lblName(69).Caption = "" And lblName(15).Caption = "" Then
      'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
      If MsgBox("專利考核部門/國外部組別錯誤或空白, 請確認是否儲存此欄位??", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Text1(15).Text = ""
      Else
         Cancel = True
         Exit Function
      End If
   End If
End If
'2008/11/5 end

'Added by Morgan 2015/1/14
If Text1(2) = "F51" And Text1(14) <> Text1(2) Then
   'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
   MsgBox "員工部門別為" & Text1(2) & "(" & lblName(2) & ")時，收文所屬部門必須設相同！" & vbCrLf & vbCrLf & "(補充保費會依照收文所屬部門判斷是否為兼職所得)", vbExclamation
   Text1(14).SetFocus 'Remove by Lydia 2021/10/25 造成重複彈訊息  'ReMark by Lydia 2021/11/22 恢復控制
   Exit Function
End If
'end 2015/1/14

'Add By Sindy 2018/5/31
If Text1(2) = "P13" And Text1(61) = "" Then
   'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
   MsgBox "員工部門別為" & Text1(2) & "(" & lblName(2) & ")時，草圖核稿人不可空白！", vbExclamation
   Text1(61).SetFocus 'Remove by Lydia 2021/10/25 造成重複彈訊息  'ReMark by Lydia 2021/11/22 恢復控制
   Exit Function
End If
If Text1(2) = "P13" And Text1(62) = "" Then
   'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
   MsgBox "員工部門別為" & Text1(2) & "(" & lblName(2) & ")時，繪圖判發人不可空白！", vbExclamation
   Text1(62).SetFocus 'Remove by Lydia 2021/10/25 造成重複彈訊息  'ReMark by Lydia 2021/11/22 恢復控制
   Exit Function
End If
If Text1(68) <> "" And InStr(Text1(68), "/") = 0 Then
   'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
   MsgBox "專業代號 －對開拓，必須輸入完整一組代號！", vbExclamation
   Text1(68).SetFocus 'Remove by Lydia 2021/10/25 造成重複彈訊息  'ReMark by Lydia 2021/11/22 恢復控制
   Exit Function
End If
'2018/5/31 END

'Add By Sindy 2018/11/7
'modify by sonia 2019/10/7
'If ((Left(Text1(2), 2) = "F1" And Text1(15) = "2") Or
If ((Left(Text1(2), 2) = "F1" And (Text1(15) = "2" Or Text1(15) = "4")) Or _
    (Left(Text1(2), 2) = "F2" And Text1(15) = "3") Or _
    (Text1(2) = "F52" And Text1(15) = "3")) _
   And Text1(69) = "" Then
   'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
   MsgBox "小組不可空白！", vbExclamation
   Text1(69).SetFocus  'Remove by Lydia 2021/10/25 造成重複彈訊息  'ReMark by Lydia 2021/11/22 恢復控制
   Exit Function
   
'Added by Morgan 2019/3/12
ElseIf (Text1(2) = "P10" Or Text1(2) = "P11") And Text1(3) = "1" And Text1(69) = "" Then
   'bolMsgEnter = True 'Added by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值 'Remove by Lydia 2021/11/22
   MsgBox "專利處組別不可空白！", vbExclamation
   Text1(69).SetFocus 'Remove by Lydia 2021/10/25 造成重複彈訊息  'ReMark by Lydia 2021/11/22 恢復控制
   Exit Function
   
End If
'2018/11/7 END

'Add By Sindy 2023/9/21
If Text1(2) = "F11" Then
   MsgBox "提醒：此人為外商承辦，請確認「系統特殊設定的(外商信件需經主核名單)」是否需要調整！", vbInformation
End If
'2023/9/21 END

'Added by Lydia 2021/10/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

TxtValidate = True
End Function

'Add by Morgan 2008/6/24
Private Sub ReadText()
   For Each oText In Text1
      idx = oText.Index
      'Modified by Lydia 2024/01/04
      'If IsNull(RcMain(idx)) Then
      If IsNull(RcMain.Fields("ST" & Format(idx + 1, "00"))) Then
         oText = ""
      Else
         If idx = 12 Then
            'Modified by Lydia 2024/01/04
            'oText = ChangeWStringToTString(RcMain(idx))
            oText = ChangeWStringToTString(RcMain.Fields("ST" & Format(idx + 1, "00")))
         Else
            'Modified by Lydia 2024/01/04
            'oText = RcMain(idx)
            oText = "" & RcMain.Fields("ST" & Format(idx + 1, "00"))
         End If
      End If
      oText.Tag = oText 'Added by Lydia 2025/08/14
   Next
End Sub

'Added by Morgan 2019/3/11
' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   
   If KeyCode = vbKeyF9 Then
      'Modified by Lydia 2021/10/25 區分確定存檔和MsgBox的回傳值; bolMsgEnter = True; 造成重複彈訊息,所以去掉SetFocus
       'ReMark by Lydia 2021/11/22 恢復控制
      If Text1(0) = "" Then MsgBox "員工代號不可為空值", vbInformation: Text1(0).SetFocus: Exit Sub
      If ActionEdit = 0 Or ActionEdit = 1 Then
         If Text1(1) = "" Then MsgBox "姓名不可為空值", vbInformation: Text1(1).SetFocus: Exit Sub
         If Text1(2) = "" Then MsgBox "員工部門別不可為空值", vbInformation: Text1(2).SetFocus: Exit Sub
         If Text1(3) = "" Then MsgBox "在職，離職不可為空值", vbInformation: Text1(3).SetFocus: Exit Sub
         If Text1(5) = "" Then MsgBox "員工所屬所別不可為空值", vbInformation: Text1(5).SetFocus: Exit Sub
         'add by sonia 2022/5/9
         If Text1(14).Tag <> "" And Text1(14).Tag <> Text1(14) Then
            MsgBox "修改收文所屬部門時，請執行智權人員調區作業，同時調整客戶之業務區，以免後續有跨區收文問題！", vbInformation
         End If
         'end 2022/5/9
         'add by sonia 2022/5/24
         If Left(Text1(2), 2) <> "F5" And (Text1(51).Tag <> Text1(51) Or Text1(52).Tag <> Text1(52) Or Text1(53).Tag <> Text1(53) Or Text1(54).Tag <> Text1(54)) Then
            MsgBox "修改帶人主管時，請同時確認是否有外翻編號，該筆資料是否要修改！", vbInformation
            'Add By Sindy 2025/6/23    '2025/9/15 modify by sonia +接洽單簽核主管
            If Trim(Text1(2)) = "F23" Then
               MsgBox "修改帶人主管時，請同時確認是否需要修改<專業代號>！還有接洽單簽核主管(法律所案源通知副本) ", vbInformation
            End If
            '2025/6/23 END
         End If
         'end 2022/5/24
      End If
      'Mark by Lydia 2021/11/22
      'If Text1(0) = "" Then bolMsgEnter = True: MsgBox "員工代號不可為空值", vbInformation:  Exit Sub
      'If ActionEdit = 0 Or ActionEdit = 1 Then
      '    If Text1(1) = "" Then bolMsgEnter = True: MsgBox "姓名不可為空值", vbInformation: Exit Sub
      '    If Text1(2) = "" Then bolMsgEnter = True: MsgBox "員工部門別不可為空值", vbInformation:  Exit Sub
      '    If Text1(3) = "" Then bolMsgEnter = True: MsgBox "在職，離職不可為空值", vbInformation: Exit Sub
      '    If Text1(5) = "" Then bolMsgEnter = True: MsgBox "員工所屬所別不可為空值", vbInformation: Exit Sub
      'End If
      ''end 2021/10/25
      
      RcEdit 3
      Text1(0).TabStop = True
      
   End If
End Sub
