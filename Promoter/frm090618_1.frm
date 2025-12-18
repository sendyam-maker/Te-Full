VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090618_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "季考核"
   ClientHeight    =   6432
   ClientLeft      =   -2856
   ClientTop       =   2088
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6432
   ScaleWidth      =   9324
   Begin TabDlg.SSTab SSTab1 
      Height          =   5955
      Left            =   45
      TabIndex        =   8
      Top             =   435
      Width           =   9225
      _ExtentX        =   16277
      _ExtentY        =   10499
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   14.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "瀏覽資料"
      TabPicture(0)   =   "frm090618_1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "詳細資料"
      TabPicture(1)   =   "frm090618_1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label22"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label24"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label6"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label8"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label9"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label10"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label11"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label12"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label13"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label14"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label15"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label16"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label17"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label18"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label19"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label20"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label21"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Line1(0)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Line1(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Line1(2)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Line1(3)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Line1(4)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Line1(5)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Line1(6)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Line1(7)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Line1(8)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Line1(9)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Line1(10)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Line1(11)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Line1(12)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Line1(13)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Line1(14)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Line1(15)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Line1(16)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Line1(17)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Line1(18)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "lblPromoter"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "lbl2(1)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "lbl2(2)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "lbl2(3)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "lbl2(4)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "lbl2(5)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "lbl2(9)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "lbl2(10)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "lbl2(11)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "lbl2(12)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "lbl2(13)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "lbl2(14)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "lbl2(15)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "lbl2(16)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "lbl2(17)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "lbl2(18)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "lbl2(19)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "lbl2(20)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "lbl2(21)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "lbl2(22)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "lbl2(23)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "lbl3(0)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "lbl3(1)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "lbl3(2)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "lbl3(3)"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "lblst01"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "Line1(19)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "Label23"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "lbl2(6)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "lbl2(7)"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "lbl2(8)"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "Label25"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "lbl2(24)"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "Label26"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "lbl2(25)"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "Line1(20)"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "Line1(21)"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "Line1(22)"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "Label27"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "Label28"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "txt1(0)"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "txt1(2)"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "txt1(3)"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "txt1(1)"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "txt1(4)"
      Tab(1).Control(86).Enabled=   0   'False
      Tab(1).Control(87)=   "txt1(5)"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).Control(88)=   "grd2"
      Tab(1).Control(88).Enabled=   0   'False
      Tab(1).Control(89)=   "cmdok(1)"
      Tab(1).Control(89).Enabled=   0   'False
      Tab(1).Control(90)=   "txtInput"
      Tab(1).Control(90).Enabled=   0   'False
      Tab(1).ControlCount=   91
      Begin VB.TextBox txtInput 
         Height          =   375
         Left            =   3930
         TabIndex        =   71
         Text            =   "Text3"
         Top             =   4005
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   5565
         Left            =   -74910
         TabIndex        =   69
         Top             =   330
         Width           =   9060
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
            Height          =   5475
            Left            =   15
            TabIndex        =   70
            Top             =   45
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   9652
            _Version        =   393216
            Rows            =   3
            FixedRows       =   2
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            HighLight       =   2
            AllowUserResizing=   1
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
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "存檔(&S)"
         Height          =   400
         Index           =   1
         Left            =   7860
         TabIndex        =   6
         Top             =   75
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
         Height          =   2475
         Left            =   135
         TabIndex        =   72
         Top             =   3330
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10181
         _ExtentY        =   4360
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
         _Band(0).Cols   =   1
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.TextBox txt1 
         Height          =   705
         Index           =   5
         Left            =   5925
         TabIndex        =   4
         Top             =   4410
         Width           =   3150
         VariousPropertyBits=   -1466941413
         ForeColor       =   16711935
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "5556;1244"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   4
         Left            =   7500
         TabIndex        =   5
         Top             =   5145
         Width           =   780
         VariousPropertyBits=   671107099
         ForeColor       =   16711935
         Size            =   "1376;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   1
         Left            =   7500
         TabIndex        =   1
         Top             =   3015
         Width           =   780
         VariousPropertyBits=   671107099
         ForeColor       =   16711935
         Size            =   "1376;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   3
         Left            =   7500
         TabIndex        =   3
         Top             =   3720
         Width           =   780
         VariousPropertyBits=   671107099
         ForeColor       =   16711935
         Size            =   "1376;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   2
         Left            =   7500
         TabIndex        =   2
         Top             =   3375
         Width           =   780
         VariousPropertyBits=   671107099
         ForeColor       =   16711935
         Size            =   "1376;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   0
         Left            =   4875
         TabIndex        =   0
         Top             =   1260
         Width           =   930
         VariousPropertyBits=   671107099
         ForeColor       =   16711935
         Size            =   "1640;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label28 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "此頁之發文點數會包含銷案點數"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   540
         TabIndex        =   68
         Top             =   3075
         Width           =   2730
      End
      Begin VB.Label Label27 
         Alignment       =   1  '靠右對齊
         Caption         =   "帶人主管備註："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6030
         TabIndex        =   67
         Top             =   4110
         Width           =   1425
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   22
         X1              =   120
         X2              =   9075
         Y1              =   5820
         Y2              =   5820
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   21
         X1              =   120
         X2              =   9075
         Y1              =   5475
         Y2              =   5475
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   20
         X1              =   120
         X2              =   9075
         Y1              =   5115
         Y2              =   5115
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   25
         Left            =   4875
         TabIndex        =   66
         Top             =   2715
         Width           =   930
      End
      Begin VB.Label Label26 
         Alignment       =   1  '靠右對齊
         Caption         =   "銷案點數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3375
         TabIndex        =   65
         Top             =   2715
         Width           =   1425
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   24
         Left            =   1650
         TabIndex        =   64
         Top             =   2715
         Width           =   930
      End
      Begin VB.Label Label25 
         Alignment       =   1  '靠右對齊
         Caption         =   "銷案基數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   63
         Top             =   2715
         Width           =   1425
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   1635
         TabIndex        =   62
         Top             =   1650
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   7515
         TabIndex        =   61
         Top             =   1305
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   60
         Top             =   1305
         Width           =   930
      End
      Begin VB.Label Label23 
         Alignment       =   1  '靠右對齊
         Caption         =   "可增加的基數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6030
         TabIndex        =   58
         Top             =   1305
         Width           =   1425
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   19
         X1              =   120
         X2              =   9075
         Y1              =   4395
         Y2              =   4395
      End
      Begin VB.Label lblst01 
         Caption         =   "lblst01"
         Height          =   255
         Left            =   2415
         TabIndex        =   56
         Top             =   4740
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lbl3 
         Caption         =   "±"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   3
         Left            =   8325
         TabIndex        =   55
         Top             =   5175
         Width           =   705
      End
      Begin VB.Label lbl3 
         Caption         =   "±"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   2
         Left            =   8340
         TabIndex        =   54
         Top             =   3750
         Width           =   705
      End
      Begin VB.Label lbl3 
         Caption         =   "±"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   1
         Left            =   8340
         TabIndex        =   53
         Top             =   3405
         Width           =   705
      End
      Begin VB.Label lbl3 
         Caption         =   "±"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   0
         Left            =   8340
         TabIndex        =   52
         Top             =   3045
         Width           =   705
      End
      Begin VB.Label lbl2 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   23
         Left            =   7500
         TabIndex        =   51
         Top             =   5535
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   22
         Left            =   7500
         TabIndex        =   50
         Top             =   5175
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   21
         Left            =   7515
         TabIndex        =   49
         Top             =   3750
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   20
         Left            =   7515
         TabIndex        =   48
         Top             =   3405
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   19
         Left            =   7515
         TabIndex        =   47
         Top             =   3045
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   18
         Left            =   7515
         TabIndex        =   46
         Top             =   2700
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   17
         Left            =   1635
         TabIndex        =   45
         Top             =   4395
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   16
         Left            =   7515
         TabIndex        =   44
         Top             =   2355
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   15
         Left            =   4920
         TabIndex        =   43
         Top             =   2355
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   14
         Left            =   1635
         TabIndex        =   42
         Top             =   2355
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   13
         Left            =   7515
         TabIndex        =   41
         Top             =   1995
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   12
         Left            =   4920
         TabIndex        =   40
         Top             =   1995
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   11
         Left            =   1635
         TabIndex        =   39
         Top             =   1995
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   10
         Left            =   7515
         TabIndex        =   38
         Top             =   1650
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   4920
         TabIndex        =   37
         Top             =   1650
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   1635
         TabIndex        =   36
         Top             =   1305
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   7515
         TabIndex        =   35
         Top             =   945
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   34
         Top             =   945
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   1635
         TabIndex        =   33
         Top             =   945
         Width           =   930
      End
      Begin VB.Label lbl2 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   32
         Top             =   615
         Width           =   930
      End
      Begin VB.Label lblPromoter 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1635
         TabIndex        =   31
         Top             =   615
         Width           =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   18
         X1              =   9075
         X2              =   9075
         Y1              =   555
         Y2              =   5820
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   17
         X1              =   7470
         X2              =   7470
         Y1              =   555
         Y2              =   5835
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   16
         X1              =   5910
         X2              =   5910
         Y1              =   555
         Y2              =   5835
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   15
         X1              =   4830
         X2              =   4830
         Y1              =   555
         Y2              =   5835
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   14
         X1              =   2670
         X2              =   2670
         Y1              =   555
         Y2              =   5835
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   13
         X1              =   1605
         X2              =   1605
         Y1              =   570
         Y2              =   5835
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   12
         X1              =   120
         X2              =   120
         Y1              =   555
         Y2              =   5835
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   11
         X1              =   120
         X2              =   9075
         Y1              =   4755
         Y2              =   4755
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   10
         X1              =   120
         X2              =   9075
         Y1              =   4050
         Y2              =   4050
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   9
         X1              =   120
         X2              =   9075
         Y1              =   3705
         Y2              =   3705
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   8
         X1              =   120
         X2              =   9075
         Y1              =   3345
         Y2              =   3345
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   7
         X1              =   120
         X2              =   9075
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   6
         X1              =   120
         X2              =   9075
         Y1              =   2655
         Y2              =   2655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   5
         X1              =   120
         X2              =   9075
         Y1              =   2295
         Y2              =   2295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   4
         X1              =   120
         X2              =   9075
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   3
         X1              =   120
         X2              =   9075
         Y1              =   1605
         Y2              =   1605
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   2
         X1              =   120
         X2              =   9075
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   1
         X1              =   120
         X2              =   9075
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   120
         X2              =   9075
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label Label21 
         Alignment       =   1  '靠右對齊
         Caption         =   "速度得分："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   29
         Top             =   4395
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label20 
         Alignment       =   1  '靠右對齊
         Caption         =   "季考核得分："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6015
         TabIndex        =   28
         Top             =   5535
         Width           =   1425
      End
      Begin VB.Label Label19 
         Alignment       =   1  '靠右對齊
         Caption         =   "部門主管評分："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6015
         TabIndex        =   27
         Top             =   5175
         Width           =   1425
      End
      Begin VB.Label Label18 
         Alignment       =   1  '靠右對齊
         Caption         =   "帶人主管評分："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6030
         TabIndex        =   26
         Top             =   3750
         Width           =   1425
      End
      Begin VB.Label Label17 
         Alignment       =   1  '靠右對齊
         Caption         =   "核稿人評分："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6030
         TabIndex        =   25
         Top             =   3405
         Width           =   1425
      End
      Begin VB.Label Label16 
         Alignment       =   1  '靠右對齊
         Caption         =   "自我比較評分："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6030
         TabIndex        =   24
         Top             =   3045
         Width           =   1425
      End
      Begin VB.Label Label15 
         Alignment       =   1  '靠右對齊
         Caption         =   "速度考核得分："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6030
         TabIndex        =   23
         Top             =   2700
         Width           =   1425
      End
      Begin VB.Label Label14 
         Alignment       =   1  '靠右對齊
         Caption         =   "發文點數得分："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6030
         TabIndex        =   22
         Top             =   2355
         Width           =   1425
      End
      Begin VB.Label Label13 
         Alignment       =   1  '靠右對齊
         Caption         =   "發文點數達成率："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2790
         TabIndex        =   21
         Top             =   2355
         Width           =   2010
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         Caption         =   "發文點數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   20
         Top             =   2355
         Width           =   1425
      End
      Begin VB.Label Label11 
         Alignment       =   1  '靠右對齊
         Caption         =   "發文張數得分："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6030
         TabIndex        =   19
         Top             =   1995
         Width           =   1425
      End
      Begin VB.Label Label10 
         Alignment       =   1  '靠右對齊
         Caption         =   "發文張數達成率："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2790
         TabIndex        =   18
         Top             =   1995
         Width           =   2010
      End
      Begin VB.Label Label9 
         Alignment       =   1  '靠右對齊
         Caption         =   "發文張數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   1995
         Width           =   1425
      End
      Begin VB.Label Label8 
         Alignment       =   1  '靠右對齊
         Caption         =   "考核基數得分："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6030
         TabIndex        =   16
         Top             =   1650
         Width           =   1425
      End
      Begin VB.Label Label7 
         Alignment       =   1  '靠右對齊
         Caption         =   "考核基數達成率："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2790
         TabIndex        =   15
         Top             =   1650
         Width           =   2010
      End
      Begin VB.Label Label6 
         Alignment       =   1  '靠右對齊
         Caption         =   "發文基數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   1305
         Width           =   1425
      End
      Begin VB.Label Label5 
         Alignment       =   1  '靠右對齊
         Caption         =   "目標點數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6030
         TabIndex        =   13
         Top             =   945
         Width           =   1425
      End
      Begin VB.Label Label4 
         Alignment       =   1  '靠右對齊
         Caption         =   "目標張數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2790
         TabIndex        =   12
         Top             =   945
         Width           =   2010
      End
      Begin VB.Label Label3 
         Alignment       =   1  '靠右對齊
         Caption         =   "目標基數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   945
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   1  '靠右對齊
         Caption         =   "考核季別："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2790
         TabIndex        =   10
         Top             =   615
         Width           =   2010
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "承辦人："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   615
         Width           =   1425
      End
      Begin VB.Label Label24 
         Alignment       =   1  '靠右對齊
         Caption         =   "考核基數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   59
         Top             =   1650
         Width           =   1425
      End
      Begin VB.Label Label22 
         Alignment       =   1  '靠右對齊
         Caption         =   "修改及複雜案件時數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2790
         TabIndex        =   57
         Top             =   1305
         Width           =   2010
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   405
      Index           =   0
      Left            =   8064
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lbl4 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4560
      TabIndex        =   73
      Top             =   120
      Width           =   45
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   300
      TabIndex        =   30
      Top             =   105
      Width           =   165
   End
End
Attribute VB_Name = "frm090618_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/13 改成Form2.0 (grd1,grd2,txt1,lblPromoter,Printer列印未改)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Modify by Morgan 2010/12/30 件數->基數
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
'Modify by Morgan 2011/4/18 數字格式 "#.#" 若 0 會變成 "." 非數字,故改為 "0.0" 格式
Option Explicit

Dim i As Integer
Dim SWPColor As String, SWPColor2 As String, SWPRow As String, SWPRow2 As String
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim PLeft(0 To 17) As Integer, iPrint As Integer, Page As Integer
Dim MaxD As Integer
Dim MinD As Integer
Dim MaxE As Integer
Dim MinE As Integer
Dim MaxF As Integer
Dim MinF As Integer
Dim MaxG As Integer
Dim MinG As Integer
Dim m_IsRun As Boolean
'暫存季考核比重
Dim m_AR02 As Double    '工程師
Dim m_AR03 As Double
Dim m_AR04 As Double
Dim m_AR05 As Double
Dim m_AR06 As Double
Dim m_AR07 As Double
Dim m_AR08 As Double
Dim m_AR13 As Double   '繪圖
Dim m_AR14 As Double
Dim m_AR15 As Double
Dim m_AR16 As Double
Dim m_AR17 As Double
Dim m_AR18 As Double
Dim m_AR19 As Double
Dim ii As Integer
Dim iRow As Integer '本次點選列數
Dim iCol As Integer '智權人員名稱欄位
'控制輸入方塊用
Dim txtInputMax As Integer
Dim txtInputMin As Integer
Dim txtInputState As String
Dim AdoRecordSet33 As New ADODB.Recordset
Dim BeginDayCP As String
Dim EndDayCP As String
'add by nickc 2007/11/26 判斷該季是否協理有評過任一人
Dim IsHaveClose As Boolean
Dim IsOverSeason As Boolean   '2010/1/19 add by sonia 判斷該考核季與系統日比較是否已過期,否則未過期先存檔,資料會少抓
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限


Private Sub cmdok_Click(Index As Integer)
Dim StrSqlaa As String
Dim StrSqlbb As String
Select Case Index
Case 0
      If cmdok(0).Caption = "結束(&X)" Then
         Unload frm090618
         Unload Me
      Else
         Me.Hide
         frm090618.Show
         Unload Me
      End If
Case 1
         'add by nickc  2005/04/14
         Screen.MousePointer = vbHourglass
         grd1.MousePointer = flexArrowHourGlass
         
         With AdoRecordSet33
               If ProSysState = "1" Then '承辦人
                  CheckOC33
                  strSql = "select * from engineerassess where ea01='" & lblst01.Caption & "' and ea02=" & Val(frm090618.txt1(0)) + 1911 & " and ea03=" & Val(frm090618.txt1(1)) & " "
                  .CursorLocation = adUseClient
                  .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If .RecordCount <> 0 Then
                        'edit by nickc 2005/04/27 加銷案點數
                        'strSQL = "update engineerassess set ea04=" & ChgNull(lbl2(2)) & ",ea05=" & ChgNull(lbl2(4)) & ",ea06=" & ChgNull(lbl2(5)) & ",ea18=" & ChgNull(lbl2(6)) & ",ea07=" & ChgNull(IIf(Val(lbl2(2)) = 0, Null, Val(lbl2(5)) / Val(lbl2(2)) * 100)) & ",ea08=" & ChgNull(lbl2(10)) & ",ea09=" & ChgNull(lbl2(14)) & ",ea10=" & ChgNull(lbl2(15)) & ",ea11=" & ChgNull(lbl2(16)) & ",ea12=" & ChgNull(lbl2(18)) & ",ea13=" & ChgNull(lbl2(19)) & ",ea14=" & ChgNull(lbl2(20)) & ",ea15=" & ChgNull(lbl2(21)) & ",ea16=" & ChgNull(lbl2(22)) & ",ea17=" & ChgNull(lbl2(23)) & ",ea19=" & ChgNull(lbl2(24)) & " where ea01='" & lblst01.Caption & "' and ea02=" & Val(frm090618.txt1(0)) + 1911 & " and ea03=" & Val(frm090618.txt1(1)) & " "
                        'Modified by Morgan 2014/7/14
                        'strSql = "update engineerassess set ea04=" & ChgNull(lbl2(2)) & ",ea05=" & ChgNull(lbl2(4)) & ",ea06=" & ChgNull(lbl2(5)) & ",ea18=" & ChgNull(lbl2(6)) & ",ea07=" & ChgNull(IIf(Val(lbl2(2)) = 0, Null, Val(lbl2(5)) / Val(lbl2(2)) * 100)) & ",ea08=" & ChgNull(lbl2(10)) & ",ea09=" & ChgNull(Val(lbl2(14)) - Val(lbl2(25))) & ",ea10=" & ChgNull(lbl2(15)) & ",ea11=" & ChgNull(lbl2(16)) & ",ea12=" & ChgNull(lbl2(18)) & ",ea13=" & ChgNull(lbl2(19)) & ",ea14=" & ChgNull(lbl2(20)) & ",ea15=" & ChgNull(lbl2(21)) & ",ea16=" & ChgNull(lbl2(22)) & ",ea17=" & ChgNull(lbl2(23)) & ",ea19=" & ChgNull(lbl2(24)) & ",ea20=" & ChgNull(lbl2(25)) & " where ea01='" & lblst01.Caption & "' and ea02=" & Val(frm090618.txt1(0)) + 1911 & " and ea03=" & Val(frm090618.txt1(1)) & " "
                        strSql = "update engineerassess set ea04=" & ChgNull(lbl2(2)) & ",ea05=" & ChgNull(lbl2(4)) & ",ea06=" & ChgNull(lbl2(5)) & ",ea18=" & ChgNull(lbl2(6)) & ",ea07=decode(" & Val(lbl2(2)) & ",0, Null," & Val(lbl2(5)) & "/" & Val(lbl2(2)) & "*100),ea08=" & ChgNull(lbl2(10)) & ",ea09=" & ChgNull(Val(lbl2(14)) - Val(lbl2(25))) & ",ea10=" & ChgNull(lbl2(15)) & ",ea11=" & ChgNull(lbl2(16)) & ",ea12=" & ChgNull(lbl2(18)) & ",ea13=" & ChgNull(lbl2(19)) & ",ea14=" & ChgNull(lbl2(20)) & ",ea15=" & ChgNull(lbl2(21)) & ",ea16=" & ChgNull(lbl2(22)) & ",ea17=" & ChgNull(lbl2(23)) & ",ea19=" & ChgNull(lbl2(24)) & ",ea20=" & ChgNull(lbl2(25)) & " where ea01='" & lblst01.Caption & "' and ea02=" & Val(frm090618.txt1(0)) + 1911 & " and ea03=" & Val(frm090618.txt1(1)) & " "
                        'end 2014/7/14
                        cnnConnection.Execute strSql
                  Else
                        'edit by nickc 2005/04/27 加銷案點數
                        'strSQL = "insert into engineerassess (ea01,ea02,ea03,ea04,ea05,ea06,ea07,ea08,ea09,ea10,ea11,ea12,ea13,ea14,ea15,ea16,ea17,ea18,ea19) values ('" & lblst01.Caption & "'," & Val(frm090618.txt1(0)) + 1911 & "," & Val(frm090618.txt1(1)) & " ," & ChgNull(lbl2(2)) & "," & ChgNull(lbl2(4)) & "," & ChgNull(lbl2(5)) & "," & ChgNull(IIf(Val(lbl2(2)) = 0, Null, Val(lbl2(5)) / Val(lbl2(2)) * 100)) & "," & ChgNull(lbl2(10)) & "," & ChgNull(lbl2(14)) & "," & ChgNull(lbl2(15)) & "," & ChgNull(lbl2(16)) & "," & ChgNull(lbl2(18)) & "," & ChgNull(lbl2(19)) & "," & ChgNull(lbl2(20)) & "," & ChgNull(lbl2(21)) & "," & ChgNull(lbl2(22)) & "," & ChgNull(lbl2(23)) & "," & ChgNull(lbl2(6)) & "," & ChgNull(lbl2(24)) & ") "
                        'Modified by Morgan 2014/7/14
                        'strSql = "insert into engineerassess (ea01,ea02,ea03,ea04,ea05,ea06,ea07,ea08,ea09,ea10,ea11,ea12,ea13,ea14,ea15,ea16,ea17,ea18,ea19,ea20) values ('" & lblst01.Caption & "'," & Val(frm090618.txt1(0)) + 1911 & "," & Val(frm090618.txt1(1)) & " ," & ChgNull(lbl2(2)) & "," & ChgNull(lbl2(4)) & "," & ChgNull(lbl2(5)) & "," & ChgNull(IIf(Val(lbl2(2)) = 0, Null, Val(lbl2(5)) / Val(lbl2(2)) * 100)) & "," & ChgNull(lbl2(10)) & "," & ChgNull(Val(lbl2(14)) - Val(lbl2(25))) & "," & ChgNull(lbl2(15)) & "," & ChgNull(lbl2(16)) & "," & ChgNull(lbl2(18)) & "," & ChgNull(lbl2(19)) & "," & ChgNull(lbl2(20)) & "," & ChgNull(lbl2(21)) & "," & ChgNull(lbl2(22)) & "," & ChgNull(lbl2(23)) & "," & ChgNull(lbl2(6)) & "," & ChgNull(lbl2(24)) & "," & ChgNull(lbl2(25)) & ") "
                        strSql = "insert into engineerassess (ea01,ea02,ea03,ea04,ea05,ea06,ea07,ea08,ea09,ea10,ea11,ea12,ea13,ea14,ea15,ea16,ea17,ea18,ea19,ea20) values ('" & lblst01.Caption & "'," & Val(frm090618.txt1(0)) + 1911 & "," & Val(frm090618.txt1(1)) & " ," & ChgNull(lbl2(2)) & "," & ChgNull(lbl2(4)) & "," & ChgNull(lbl2(5)) & ",decode(" & Val(lbl2(2)) & ",0,Null," & Val(lbl2(5)) & "/" & Val(lbl2(2)) & "*100)," & ChgNull(lbl2(10)) & "," & ChgNull(Val(lbl2(14)) - Val(lbl2(25))) & "," & ChgNull(lbl2(15)) & "," & ChgNull(lbl2(16)) & "," & ChgNull(lbl2(18)) & "," & ChgNull(lbl2(19)) & "," & ChgNull(lbl2(20)) & "," & ChgNull(lbl2(21)) & "," & ChgNull(lbl2(22)) & "," & ChgNull(lbl2(23)) & "," & ChgNull(lbl2(6)) & "," & ChgNull(lbl2(24)) & "," & ChgNull(lbl2(25)) & ") "
                        'end 2014/7/14
                        cnnConnection.Execute strSql
                  End If
                  If ProState = "2" Then  '管理
                        '加儲存核稿跟帶人主管資料
                        For ii = 0 To grd2.Rows - 1
                           grd2.row = ii
                           grd2.col = 0
                           StrSqlaa = ""
                           StrSqlbb = ""
                           If grd2.Text = "核稿人" Then
                                 StrSqlaa = StrSqlaa & " and ab05='1' "
                                 StrSqlbb = StrSqlbb & "'1' "
                           ElseIf grd2.Text = "帶人主管" Then
                                 StrSqlaa = StrSqlaa & " and ab05='2' "
                                 StrSqlbb = StrSqlbb & "'2' "
                           End If
                           If StrSqlaa <> "" Then
                                 grd2.col = 3
                                 StrSqlaa = StrSqlaa & " and ab06='" & grd2.Text & "' "
                                 StrSqlbb = StrSqlbb & ",'" & grd2.Text & "' "
                                 CheckOC33
                                 grd2.col = 2
                                 strSql = "select * from assessboss where ab01='" & lblst01.Caption & "' and ab02=" & Val(frm090618.txt1(0)) + 1911 & " and ab03=" & Val(frm090618.txt1(1)) & " and ab04='1' " & StrSqlaa
                                 .CursorLocation = adUseClient
                                 .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                                 If .RecordCount <> 0 Then
                                       'edit by nickc 2005/04/26 加備註
                                       'strSQL = "update assessboss set ab07=" & grd2.Text & " where ab01='" & lblst01.Caption & "' and ab02=" & Val(frm090618.txt1(0)) + 1911 & " and ab03=" & Val(frm090618.txt1(1)) & " and ab04='1' " & StrSqlaa
                                       'Modify by Morgan 2008/10/28 可刪除本人所打的分數(輸錯人...)
                                       If Trim(grd2.Text) = "" Then
                                          strSql = "delete assessboss where ab01='" & lblst01.Caption & "' and ab02=" & Val(frm090618.txt1(0)) + 1911 & " and ab03=" & Val(frm090618.txt1(1)) & " and ab04='1' " & StrSqlaa
                                       Else
                                          strSql = "update assessboss set ab07=" & grd2.Text & ",ab08='" & ChgSQL(grd2.TextMatrix(ii, 4)) & "' where ab01='" & lblst01.Caption & "' and ab02=" & Val(frm090618.txt1(0)) + 1911 & " and ab03=" & Val(frm090618.txt1(1)) & " and ab04='1' " & StrSqlaa
                                       End If
                                       cnnConnection.Execute strSql
                                 Else
                                       'edit by nickc 2005/04/26 加備註
                                       'strSQL = "insert into assessboss (ab01,ab02,ab03,ab04,ab05,ab06,ab07) values ('" & lblst01.Caption & "'," & Val(frm090618.txt1(0)) + 1911 & "," & Val(frm090618.txt1(1)) & ",'1'," & StrSqlbb & "," & grd2.Text & ") "
                                       If Trim(grd2.Text) <> "" Then
                                          strSql = "insert into assessboss (ab01,ab02,ab03,ab04,ab05,ab06,ab07,ab08) values ('" & lblst01.Caption & "'," & Val(frm090618.txt1(0)) + 1911 & "," & Val(frm090618.txt1(1)) & ",'1'," & StrSqlbb & "," & grd2.Text & ",'" & ChgSQL(grd2.TextMatrix(ii, 4)) & "') "
                                          cnnConnection.Execute strSql
                                       End If
                                 End If
                                 CheckOC33
                           End If
                        Next ii
                  End If
               Else '繪圖
                  CheckOC33
                  strSql = "select * from drawassess where da01='" & lblst01.Caption & "' and da02=" & Val(frm090618.txt1(0)) + 1911 & " and da03=" & Val(frm090618.txt1(1)) & " "
                  .CursorLocation = adUseClient
                  .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If .RecordCount <> 0 Then
                        'edit by nickc 2005/04/27 加銷案點數
                        'strSQL = "update drawassess set da04=" & ChgNull(lbl2(2)) & ",da05=" & ChgNull(lbl2(3)) & ",da06=" & ChgNull(lbl2(4)) & ",da07=" & ChgNull(lbl2(5)) & ",da22=" & ChgNull(lbl2(6)) & ",da08=" & ChgNull(IIf(Val(lbl2(2)) = 0, Null, Val(lbl2(5)) / Val(lbl2(2)) * 100)) & ",da09=" & ChgNull(lbl2(10)) & ",da10=" & ChgNull(lbl2(11)) & ",da11=" & ChgNull(lbl2(12)) & ",da12=" & ChgNull(lbl2(13)) & ",da13=" & ChgNull(lbl2(14)) & ",da14=" & ChgNull(lbl2(15)) & ",da15=" & ChgNull(lbl2(16)) & ",da16=" & ChgNull(lbl2(18)) & ",da17=" & ChgNull(lbl2(19)) & ",da19=" & ChgNull(lbl2(21)) & ",da20=" & ChgNull(lbl2(22)) & ",da21=" & ChgNull(lbl2(23)) & ",da18=" & ChgNull(lbl2(24)) & " where  da01='" & lblst01.Caption & "' and da02=" & Val(frm090618.txt1(0)) + 1911 & " and da03=" & Val(frm090618.txt1(1)) & " "
                        'Modified by Morgan 2014/7/14
                        'strSql = "update drawassess set da04=" & ChgNull(lbl2(2)) & ",da05=" & ChgNull(lbl2(3)) & ",da06=" & ChgNull(lbl2(4)) & ",da07=" & ChgNull(lbl2(5)) & ",da22=" & ChgNull(lbl2(6)) & ",da08=" & ChgNull(IIf(Val(lbl2(2)) = 0, Null, Val(lbl2(5)) / Val(lbl2(2)) * 100)) & ",da09=" & ChgNull(lbl2(10)) & ",da10=" & ChgNull(lbl2(11)) & ",da11=" & ChgNull(lbl2(12)) & ",da12=" & ChgNull(lbl2(13)) & ",da13=" & ChgNull(Val(lbl2(14)) - Val(lbl2(25))) & ",da14=" & ChgNull(lbl2(15)) & ",da15=" & ChgNull(lbl2(16)) & ",da16=" & ChgNull(lbl2(18)) & ",da17=" & ChgNull(lbl2(19)) & ",da19=" & ChgNull(lbl2(21)) & ",da20=" & ChgNull(lbl2(22)) & ",da21=" & ChgNull(lbl2(23)) & ",da18=" & ChgNull(lbl2(24)) & ",da23=" & ChgNull(lbl2(25)) & " where  da01='" & lblst01.Caption & "' and da02=" & Val(frm090618.txt1(0)) + 1911 & " and da03=" & Val(frm090618.txt1(1)) & " "
                        strSql = "update drawassess set da04=" & ChgNull(lbl2(2)) & ",da05=" & ChgNull(lbl2(3)) & ",da06=" & ChgNull(lbl2(4)) & ",da07=" & ChgNull(lbl2(5)) & ",da22=" & ChgNull(lbl2(6)) & ",da08=decode(" & Val(lbl2(2)) & ",0, Null," & Val(lbl2(5)) & "/" & Val(lbl2(2)) & "*100),da09=" & ChgNull(lbl2(10)) & ",da10=" & ChgNull(lbl2(11)) & ",da11=" & ChgNull(lbl2(12)) & ",da12=" & ChgNull(lbl2(13)) & ",da13=" & ChgNull(Val(lbl2(14)) - Val(lbl2(25))) & ",da14=" & ChgNull(lbl2(15)) & ",da15=" & ChgNull(lbl2(16)) & ",da16=" & ChgNull(lbl2(18)) & ",da17=" & ChgNull(lbl2(19)) & ",da19=" & ChgNull(lbl2(21)) & ",da20=" & ChgNull(lbl2(22)) & ",da21=" & ChgNull(lbl2(23)) & ",da18=" & ChgNull(lbl2(24)) & ",da23=" & ChgNull(lbl2(25)) & " where  da01='" & lblst01.Caption & "' and da02=" & Val(frm090618.txt1(0)) + 1911 & " and da03=" & Val(frm090618.txt1(1)) & " "
                        'end 2014/7/14
                        cnnConnection.Execute strSql
                  Else
                        'edit by nickc 2005/04/27 加銷案點數
                        'strSQL = "insert into drawassess (da01,da02,da03,da04,da05,da06,da07,da08,da09,da10,da11,da12,da13,da14,da15,da16,da17,da19,da20,da21,da22,da18) values ('" & lblst01.Caption & "'," & Val(frm090618.txt1(0)) + 1911 & "," & Val(frm090618.txt1(1)) & " ," & ChgNull(lbl2(2)) & "," & ChgNull(lbl2(3)) & "," & ChgNull(lbl2(4)) & "," & ChgNull(lbl2(5)) & "," & ChgNull(IIf(Val(lbl2(2)) = 0, Null, Val(lbl2(5)) / Val(lbl2(2)) * 100)) & "," & ChgNull(lbl2(10)) & "," & ChgNull(lbl2(11)) & "," & ChgNull(lbl2(12)) & "," & ChgNull(lbl2(13)) & "," & ChgNull(lbl2(14)) & "," & ChgNull(lbl2(15)) & "," & ChgNull(lbl2(16)) & "," & ChgNull(lbl2(18)) & "," & ChgNull(lbl2(19)) & "," & ChgNull(lbl2(21)) & "," & ChgNull(lbl2(22)) & "," & ChgNull(lbl2(23)) & "," & ChgNull(lbl2(6)) & "," & ChgNull(lbl2(24)) & ") "
                        'Modified by Morgan 2014/7/14
                        'strSql = "insert into drawassess (da01,da02,da03,da04,da05,da06,da07,da08,da09,da10,da11,da12,da13,da14,da15,da16,da17,da19,da20,da21,da22,da18,da23) values ('" & lblst01.Caption & "'," & Val(frm090618.txt1(0)) + 1911 & "," & Val(frm090618.txt1(1)) & " ," & ChgNull(lbl2(2)) & "," & ChgNull(lbl2(3)) & "," & ChgNull(lbl2(4)) & "," & ChgNull(lbl2(5)) & "," & ChgNull(IIf(Val(lbl2(2)) = 0, Null, Val(lbl2(5)) / Val(lbl2(2)) * 100)) & "," & ChgNull(lbl2(10)) & "," & ChgNull(lbl2(11)) & "," & ChgNull(lbl2(12)) & "," & ChgNull(lbl2(13)) & "," & ChgNull(Val(lbl2(14)) - Val(lbl2(25))) & "," & ChgNull(lbl2(15)) & "," & ChgNull(lbl2(16)) & "," & ChgNull(lbl2(18)) & "," & ChgNull(lbl2(19)) & "," & ChgNull(lbl2(21)) & "," & ChgNull(lbl2(22)) & "," & ChgNull(lbl2(23)) & "," & ChgNull(lbl2(6)) & "," & ChgNull(lbl2(24)) & "," & ChgNull(lbl2(25)) & ") "
                        strSql = "insert into drawassess (da01,da02,da03,da04,da05,da06,da07,da08,da09,da10,da11,da12,da13,da14,da15,da16,da17,da19,da20,da21,da22,da18,da23) values ('" & lblst01.Caption & "'," & Val(frm090618.txt1(0)) + 1911 & "," & Val(frm090618.txt1(1)) & " ," & ChgNull(lbl2(2)) & "," & ChgNull(lbl2(3)) & "," & ChgNull(lbl2(4)) & "," & ChgNull(lbl2(5)) & ",decode(" & Val(lbl2(2)) & ",0, Null," & Val(lbl2(5)) & "/" & Val(lbl2(2)) & "*100)," & ChgNull(lbl2(10)) & "," & ChgNull(lbl2(11)) & "," & ChgNull(lbl2(12)) & "," & ChgNull(lbl2(13)) & "," & ChgNull(Val(lbl2(14)) - Val(lbl2(25))) & "," & ChgNull(lbl2(15)) & "," & ChgNull(lbl2(16)) & "," & ChgNull(lbl2(18)) & "," & ChgNull(lbl2(19)) & "," & ChgNull(lbl2(21)) & "," & ChgNull(lbl2(22)) & "," & ChgNull(lbl2(23)) & "," & ChgNull(lbl2(6)) & "," & ChgNull(lbl2(24)) & "," & ChgNull(lbl2(25)) & ") "
                        'end 2014/7/14
                        cnnConnection.Execute strSql
                  End If
                  If ProState = "2" Then  '管理
                        '加儲存核稿跟帶人主管資料
                        For ii = 0 To grd2.Rows - 1
                           grd2.row = ii
                           grd2.col = 0
                           StrSqlaa = ""
                           StrSqlbb = ""
                           If grd2.Text = "核稿人" Then
                                 StrSqlaa = StrSqlaa & " and ab05='1' "
                                 StrSqlbb = StrSqlbb & "'1' "
                           ElseIf grd2.Text = "帶人主管" Then
                                 StrSqlaa = StrSqlaa & " and ab05='2' "
                                 StrSqlbb = StrSqlbb & "'2' "
                           End If
                           If StrSqlaa <> "" Then
                                 grd2.col = 3
                                 StrSqlaa = StrSqlaa & " and ab06='" & grd2.Text & "' "
                                 StrSqlbb = StrSqlbb & ",'" & grd2.Text & "' "
                                 CheckOC33
                                 grd2.col = 2
                                 strSql = "select * from assessboss where ab01='" & lblst01.Caption & "' and ab02=" & Val(frm090618.txt1(0)) + 1911 & " and ab03=" & Val(frm090618.txt1(1)) & " and ab04='2' " & StrSqlaa
                                 .CursorLocation = adUseClient
                                 .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                                 If .RecordCount <> 0 Then
                                       'edit by nickc 2005/04/26 加備註
                                       'strSQL = "update assessboss set ab07=" & grd2.Text & " where ab01='" & lblst01.Caption & "' and ab02=" & Val(frm090618.txt1(0)) + 1911 & " and ab03=" & Val(frm090618.txt1(1)) & " and ab04='2' " & StrSqlaa
                                       'Modify by Morgan 2008/10/28 可刪除本人所打的分數(輸錯人...)
                                       If Trim(grd2.Text) = "" Then
                                          strSql = "delete assessboss where ab01='" & lblst01.Caption & "' and ab02=" & Val(frm090618.txt1(0)) + 1911 & " and ab03=" & Val(frm090618.txt1(1)) & " and ab04='2' " & StrSqlaa
                                       Else
                                       'end 2008/10/24
                                          strSql = "update assessboss set ab07=" & grd2.Text & ",ab08='" & ChgSQL(grd2.TextMatrix(ii, 4)) & "' where ab01='" & lblst01.Caption & "' and ab02=" & Val(frm090618.txt1(0)) + 1911 & " and ab03=" & Val(frm090618.txt1(1)) & " and ab04='2' " & StrSqlaa
                                       End If
                                       cnnConnection.Execute strSql
                                 Else
                                       'edit by nickc 2005/04/26 加備註
                                       'strSQL = "insert into assessboss (ab01,ab02,ab03,ab04,ab05,ab06,ab07) values ('" & lblst01.Caption & "'," & Val(frm090618.txt1(0)) + 1911 & "," & Val(frm090618.txt1(1)) & ",'2'," & StrSqlbb & "," & grd2.Text & ") "
                                       If Trim(grd2.Text) <> "" Then
                                          strSql = "insert into assessboss (ab01,ab02,ab03,ab04,ab05,ab06,ab07,ab08) values ('" & lblst01.Caption & "'," & Val(frm090618.txt1(0)) + 1911 & "," & Val(frm090618.txt1(1)) & ",'2'," & StrSqlbb & "," & grd2.Text & ",'" & ChgSQL(grd2.TextMatrix(ii, 4)) & "') "
                                          cnnConnection.Execute strSql
                                       End If
                                 End If
                                 CheckOC33
                           End If
                        Next ii
                  End If
               End If
         End With
         MsgBox "存檔成功！", vbOKOnly, "季考核存檔"
         SSTab1.Visible = False 'Added by Morgan 2018/5/24 避免畫面會閃
         StrMenu
         SSTab1.Visible = True 'Added by Morgan 2018/5/24
         'add by nickc 2005/04/14 若瀏覽畫面可以用的話，應該要回瀏覽畫面
         If SSTab1.TabEnabled(0) = True Then
            SSTab1.Tab = 0
         End If
         Screen.MousePointer = vbDefault
         grd1.MousePointer = flexDefault
Case Else
End Select
End Sub

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
If m_IsRun = False Then
   m_IsRun = True
      If frm090618.txt1(4) = "2" Then
         Me.Hide
      End If
      lbl1.Caption = "考核年月、季別：" & frm090618.txt1(0) & " 年 " & frm090618.txt1(1) & "  季 "
      If ProState = "2" Then
         Me.SSTab1.Tab = 0
      End If
      Screen.MousePointer = vbHourglass
      Me.grd1.MousePointer = flexHourglass
      DoEvents
      If StrMenu = False Then
            If frm090618.txt1(4) = "2" Then
               grd1.MousePointer = flexDefault
               Screen.MousePointer = vbDefault
               Unload Me
            Else
               grd1.MousePointer = flexDefault
               Screen.MousePointer = vbDefault
               cmdok_Click (0)
            End If
           Exit Sub
      End If
      grd1.MousePointer = flexDefault
      Screen.MousePointer = vbDefault
      If ProState <> "2" Then
         SSTab1.Tab = 1
         SSTab1.TabEnabled(0) = False
         txt1(2).Visible = False
         txt1(3).Visible = False
         txt1(4).Visible = False
         Label17.Visible = False
         Label18.Visible = False
         Label19.Visible = False
         Label27.Visible = False
         lbl3(1).Visible = False
         lbl3(2).Visible = False
         lbl3(3).Visible = False
         lbl2(20).Visible = False
         lbl2(21).Visible = False
         lbl2(22).Visible = False
         If Val(Trim(lbl2(20).Caption)) + Val(Trim(lbl2(21).Caption)) + Val(Trim(lbl2(22).Caption)) <> 0 Then
            txt1(0).Enabled = False
            txt1(1).Enabled = False
         End If
      Else
         'Modified by Morgan 2023/3/25
         'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "00" Then
         'Modified by Morgan 2025/2/4 +P10部門
         'Modified by Morgan 2025/6/26 +79075
         If Pub_strUserST05 = "71" Or Pub_strUserST05 = "72" Or Pub_strUserST05 = "00" Or Pub_StrUserSt03 = "P10" Or strUserNum = "79075" Then
         'end 2023/3/25
         Else
               lbl2(22).Visible = False
               lbl3(3).Visible = False
               Label19.Visible = False
               txt1(4).Visible = False
            If Val(Trim(lbl2(22).Caption)) <> 0 Then
                txt1(0).Visible = False
                txt1(1).Visible = False
                txt1(2).Visible = False
                txt1(3).Visible = False
                'add by nickc 2007/11/26
                txt1(5).Visible = False
            End If
         End If
      End If
      Me.Show
End If
End Sub

Private Sub Form_Load()
m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
m_IsRun = False
MoveFormToCenter Me
End Sub

Private Sub SetGrd1()
Dim j As Integer
With grd1
    .Visible = False
    If ProSysState = "1" Then
         'edit by nickc 2005/04/27
         '.Cols = 19
         .Cols = 20
         .row = 0
         .col = 0:   .Text = "承辦人"
         .ColWidth(0) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 1:   .Text = "目標"
         .ColWidth(1) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 2:   .Text = "目標"
         .ColWidth(2) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 3:   .Text = "發文"
         .ColWidth(3) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 4:   .Text = "發文"
         .ColWidth(4) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 5:   .Text = "發文"
         .ColWidth(5) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 6:   .Text = "發文"
         .ColWidth(6) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 7:   .Text = "發文"
         .ColWidth(7) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 8:   .Text = "發文"
         .ColWidth(8) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 9:   .Text = "速度考核"
         .ColWidth(9) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 10:   .Text = "自我比較"
         .ColWidth(10) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 11:   .Text = "核稿人"
         If ProState = "2" Then
            .ColWidth(11) = 800
         Else
            .ColWidth(11) = 0
         End If
         .CellAlignment = flexAlignCenterCenter
         .col = 12:  .Text = "帶人主管"
         If ProState = "2" Then
            .ColWidth(12) = 800
         Else
            .ColWidth(12) = 0
         End If
         .CellAlignment = flexAlignCenterCenter
         .col = 13:  .Text = "部門主管"
         If ProState = "2" Then
            'Modified by Morgan 2023/3/25
            'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "00" Then
            'Modified by Morgan 2025/2/4 +P10部門
            If Pub_strUserST05 = "71" Or Pub_strUserST05 = "72" Or Pub_strUserST05 = "00" Or Pub_StrUserSt03 = "P10" Then
            'end 2023/3/25
               .ColWidth(13) = 800
            Else
               .ColWidth(13) = 0
            End If
         Else
            .ColWidth(13) = 0
         End If
         .CellAlignment = flexAlignCenterCenter
         .col = 14:  .Text = "季考核"
         .ColWidth(14) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 15:  .Text = ""
         .ColWidth(15) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 16:  .Text = ""
         .ColWidth(16) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 17:  .Text = ""
         .ColWidth(17) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 18:  .Text = ""
         .ColWidth(18) = 0
         .CellAlignment = flexAlignCenterCenter
         'add by nickc 2005/04/27
         .col = 19:  .Text = ""
         .ColWidth(19) = 0
         .CellAlignment = flexAlignCenterCenter
         
         .row = 1
         .col = 0:   .Text = "承辦人"
         .ColWidth(0) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 1:   .Text = "基數"
         .ColWidth(1) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 2:   .Text = "點數"
         .ColWidth(2) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 3:   .Text = "基數"
         .ColWidth(3) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 4:   .Text = "基數達成率%"
         .ColWidth(4) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 5:   .Text = "基數得分"
         .ColWidth(5) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 6:   .Text = "點數"
         .ColWidth(6) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 7:   .Text = "點數達成率%"
         .ColWidth(7) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 8:   .Text = "點數得分"
         .ColWidth(8) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 9:   .Text = "得分"
         .ColWidth(9) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 10:   .Text = "評分"
         .ColWidth(10) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 11:   .Text = "評分"
         If ProState = "2" Then
            .ColWidth(11) = 800
         Else
            .ColWidth(11) = 0
         End If
         .CellAlignment = flexAlignCenterCenter
         .col = 12:  .Text = "評分"
         If ProState = "2" Then
            .ColWidth(12) = 800
         Else
            .ColWidth(12) = 0
         End If
         .CellAlignment = flexAlignCenterCenter
         .col = 13:  .Text = "評分"
         If ProState = "2" Then
            'Modified by Morgan 2023/3/25
            'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "00" Then
            'Modified by Morgan 2025/2/4 +P10部門
            If Pub_strUserST05 = "71" Or Pub_strUserST05 = "72" Or Pub_strUserST05 = "00" Or Pub_StrUserSt03 = "P10" Then
            'end 2023/3/25
               .ColWidth(13) = 800
            Else
               .ColWidth(13) = 0
            End If
         Else
            .ColWidth(13) = 0
         End If
         .CellAlignment = flexAlignCenterCenter
         .col = 14:  .Text = "得分"
         .ColWidth(14) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 15:  .Text = ""
         .ColWidth(15) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 16:  .Text = ""
         .ColWidth(16) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 17:  .Text = ""
         .ColWidth(17) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 18:  .Text = ""
         .ColWidth(18) = 0
         .CellAlignment = flexAlignCenterCenter
         'add by nickc 2005/04/27
         .col = 19:  .Text = ""
         .ColWidth(19) = 0
         .CellAlignment = flexAlignCenterCenter
   Else
         'edit by nickc 2005/04/27
         '.Cols = 23
         .Cols = 24
         .row = 0
         .col = 0:   .Text = "繪圖人員"
         .ColWidth(0) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 1:   .Text = "目標"
         .ColWidth(1) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 2:   .Text = "目標"
         .ColWidth(2) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 3:   .Text = "目標"
         .ColWidth(3) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 4: .Text = "發文"
         .ColWidth(4) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 5:   .Text = "發文"
         .ColWidth(5) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 6:   .Text = "發文"
         .ColWidth(6) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 7:   .Text = "發文"
         .ColWidth(7) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 8:   .Text = "發文"
         .ColWidth(8) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 9:   .Text = "發文"
         .ColWidth(9) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 10:   .Text = "發文"
         .ColWidth(10) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 11:   .Text = "發文"
         .ColWidth(11) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 12:   .Text = "發文"
         .ColWidth(12) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 13:   .Text = "速度考核"
         .ColWidth(13) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 14:   .Text = "自我比較"
         .ColWidth(14) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 15:   .Text = "帶人主管"
         If ProState = "2" Then
            .ColWidth(15) = 800
         Else
            .ColWidth(15) = 0
         End If
         .CellAlignment = flexAlignCenterCenter
         .col = 16:   .Text = "部門主管"
         If ProState = "2" Then
            'Modified by Morgan 2023/3/25
            'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "00" Then
            'Modified by Morgan 2025/2/4 +P10部門
            If Pub_strUserST05 = "71" Or Pub_strUserST05 = "72" Or Pub_strUserST05 = "00" Or Pub_StrUserSt03 = "P10" Then
            'end 2023/3/25
               .ColWidth(16) = 800
            Else
               .ColWidth(16) = 0
            End If
         Else
            .ColWidth(16) = 0
         End If
         .CellAlignment = flexAlignCenterCenter
         .col = 17:   .Text = "季考核"
         .ColWidth(17) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 18:  .Text = ""
         .ColWidth(18) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 19:  .Text = ""
         .ColWidth(19) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 20:  .Text = ""
         .ColWidth(20) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 21:  .Text = ""
         .ColWidth(21) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 22:  .Text = ""
         .ColWidth(22) = 0
         .CellAlignment = flexAlignCenterCenter
         'add by nickc 2005/04/27
         .col = 23:  .Text = ""
         .ColWidth(23) = 0
         .CellAlignment = flexAlignCenterCenter
         
         .row = 1
         .col = 0:   .Text = "繪圖人員"
         .ColWidth(0) = 1000
         .CellAlignment = flexAlignCenterCenter
         .col = 1:   .Text = "基數"
         .ColWidth(1) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 2:   .Text = "張數"
         .ColWidth(2) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 3:   .Text = "點數"
         .ColWidth(3) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 4:   .Text = "基數"
         .ColWidth(4) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 5:   .Text = "達成率%"
         .ColWidth(5) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 6:   .Text = "得分"
         .ColWidth(6) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 7:   .Text = "張數"
         .ColWidth(7) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 8:   .Text = "達成率%"
         .ColWidth(8) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 9:   .Text = "得分"
         .ColWidth(9) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 10:   .Text = "點數"
         .ColWidth(10) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 11:   .Text = "達成率%"
         .ColWidth(11) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 12:   .Text = "得分"
         .ColWidth(12) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 13:   .Text = "得分"
         .ColWidth(13) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 14:   .Text = "評分"
         .ColWidth(14) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 15:   .Text = "評分"
         If ProState = "2" Then
            .ColWidth(15) = 800
         Else
            .ColWidth(15) = 0
         End If
         .CellAlignment = flexAlignCenterCenter
         .col = 16:   .Text = "評分"
         If ProState = "2" Then
            'Modified by Morgan 2023/3/25
            'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "00" Then
            'Modified by Morgan 2025/2/4 +P10部門
            If Pub_strUserST05 = "71" Or Pub_strUserST05 = "72" Or Pub_strUserST05 = "00" Or Pub_StrUserSt03 = "P10" Then
            'end 2023/3/25
               .ColWidth(16) = 800
            Else
               .ColWidth(16) = 0
            End If
         Else
            .ColWidth(16) = 0
         End If
         .CellAlignment = flexAlignCenterCenter
         .col = 17:   .Text = "得分"
         .ColWidth(17) = 800
         .CellAlignment = flexAlignCenterCenter
         .col = 18:  .Text = ""
         .ColWidth(18) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 19:  .Text = ""
         .ColWidth(19) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 20:  .Text = ""
         .ColWidth(20) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 21:  .Text = ""
         .ColWidth(21) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 22:  .Text = ""
         .ColWidth(22) = 0
         .CellAlignment = flexAlignCenterCenter
         'add by nickc 2005/04/27
         .col = 23:  .Text = ""
         .ColWidth(23) = 0
         .CellAlignment = flexAlignCenterCenter
   End If
   .MergeCells = flexMergeRestrictRows
   .MergeRow(0) = True
   .MergeCol(0) = True

   .MergeCol(1) = True
    .Visible = True
End With
   With Me.grd1
      .row = 2
         For j = 1 To .Cols - 1
             .col = j
             .CellBackColor = &HFFC0C0
         Next j
         Process .row
      SWPColor2 = SWPColor
      SWPRow2 = .row
   End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090618_1 = Nothing
End Sub

Function StrMenu() As Boolean

StrMenu = True
Dim strSql As String
Dim strSQL1 As String
Dim strSQL2 As String
Dim CalMonth As Integer
Dim j As Integer
Dim BeginDay As String
Dim EndDay As String
Select Case Trim(frm090618.txt1(1))
Case "1"
         BeginDay = Mid(Trim(Val(Trim(frm090618.txt1(0)) & "0101") + 19110000), 1, 6)
         EndDay = Mid(Trim(Val(Trim(frm090618.txt1(0)) & "0331") + 19110000), 1, 6)
Case "2"
         BeginDay = Mid(Trim(Val(Trim(frm090618.txt1(0)) & "0401") + 19110000), 1, 6)
         EndDay = Mid(Trim(Val(Trim(frm090618.txt1(0)) & "0631") + 19110000), 1, 6)
Case "3"
         BeginDay = Mid(Trim(Val(Trim(frm090618.txt1(0)) & "0701") + 19110000), 1, 6)
         EndDay = Mid(Trim(Val(Trim(frm090618.txt1(0)) & "0931") + 19110000), 1, 6)
Case "4"
         BeginDay = Mid(Trim(Val(Trim(frm090618.txt1(0)) & "1001") + 19110000), 1, 6)
         EndDay = Mid(Trim(Val(Trim(frm090618.txt1(0)) & "1231") + 19110000), 1, 6)
Case Else
End Select
BeginDayCP = BeginDay & "01"
EndDayCP = EndDay & "31"

strSql = ""
strSQL1 = ""
strSQL2 = ""
CalMonth = 3
If Len(Trim(frm090618.txt1(6))) <> 0 Then
   strSQL1 = strSQL1 & " and ma01='" & frm090618.txt1(6) & "' "
   strSQL2 = strSQL2 & " and pe01='" & frm090618.txt1(6) & "' "
End If
If Len(Trim(frm090618.txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " and st03>='" & frm090618.txt1(2) & "' "
End If
If Len(Trim(frm090618.txt1(3))) <> 0 Then
   strSQL1 = strSQL1 & " and st03<='" & frm090618.txt1(3) & "' "
End If
If Len(Trim(frm090618.txt1(7))) <> 0 Then
   strSQL1 = strSQL1 & " and st06>='" & frm090618.txt1(7) & "' "
End If
If Len(Trim(frm090618.txt1(8))) <> 0 Then
   strSQL1 = strSQL1 & " and st06<='" & frm090618.txt1(8) & "' "
End If
'add by nickc 2005/04/18
strSQL1 = strSQL1 & " and st04='1' "
strSQL1 = strSQL1 & " and ma03='" & ProSysState & "' "
'add by nickc 2007/11/26 假如協理評過該季任一人，皆不可以修改
IsHaveClose = False
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open "select * from engineerassess where ea02=" & Val(frm090618.txt1(0)) + 1911 & " and ea03=" & frm090618.txt1(1) & " and ea16>0 ", cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
    IsHaveClose = True
End If
'2010/1/19 add by sonia 該考核季與系統日比較,尚未過時者不可按存檔
IsOverSeason = True
If EndDayCP < strSrvDate(1) Then
   IsOverSeason = False
   lbl4.Caption = ""
Else
   lbl4.Caption = "此季尚未達考核時間, 已取消存檔功能 !"
End If
'2010/1/19 end

'MODIFY BY SONIA 2014/4/11 加入 pe02 in ('P','CFP') 杜燕文有T的目標
'Modified by Morgan 2019/3/19 108考核(逾期件數修改每件倒扣0.5分，不再除以「當月達成率」。)
If ProSysState = "1" Then '承辦人
      'edit by nickc 2005/04/27 加銷案點數
      'strSQL = " select  nvl(ea04,A1),nvl(ea05,A2),nvl(ea06,ma37),nvl(ea07,decode(A1,0,0,round(ma37/A1 * 100,2))),nvl(ea08,0),nvl(ea09,ma40),nvl(ea10,decode(A2,0,0,round(ma40/A2 * 100,2))),nvl(ea11,0),nvl(ea12,round(ma35/" & CalMonth & ",2)),ea13,ea14,ea15,ea16,nvl(ea17,0),st02,ea18,st01,nvl(ea19,CancelCount)  from (select pe01,sum(nvl(decode(pe02,'CFP',pe05 * 2 ,pe05),0) + nvl(decode(pe02,'CFP',pe07*2,pe07),0)) as A1, sum(nvl(pe06,0) + nvl(pe08,0)) as A2,sum(nvl(pe09,0)) as A3,sum(nvl(pe10,0)) as A4,sum(nvl(pe11,0)) as A5 from performance where pe03>=" & BeginDay & " and pe03<=" & EndDay & " " & strSQL2 & " group by pe01) APE ,("
      strSql = " select  nvl(ea04,A1),nvl(ea05,A2),nvl(ea06,ma37),nvl(ea07,decode(A1,0,0,round(ma37/A1 * 100,2))),nvl(ea08,0),nvl(ea09,ma40),nvl(ea10,decode(A2,0,0,round(ma40/A2 * 100,2))),nvl(ea11,0),nvl(ea12,round(ma35/" & CalMonth & ",2)),ea13,ea14,ea15,ea16,nvl(ea17,0),st02,ea18,st01,nvl(ea19,CancelCount),nvl(ea20,CancelPoint)  from (select pe01,sum(nvl(decode(pe02,'CFP',pe05 * 2 ,pe05),0) + nvl(decode(pe02,'CFP',pe07*2,pe07),0)) as A1, sum(nvl(pe06,0) + nvl(pe08,0)) as A2,sum(nvl(pe09,0)) as A3,sum(nvl(pe10,0)) as A4,sum(nvl(pe11,0)) as A5 from performance where pe02 in ('P','CFP') And pe03>=" & BeginDay & " and pe03<=" & EndDay & " " & strSQL2 & " group by pe01) APE ,("
      strSql = strSql & " select st01,st02,ma03,sum(nvl(ma04,0)) as ma04,sum(nvl(ma37,0)) as ma37,sum(nvl(ma40,0)) as ma40,sum(nvl(ma43,0)) as ma43,sum(nvl(ma35 - " & IIf(Val(BeginDayCP) >= Val(PUB_108RuleDate), "0.5*ma51", "decode(ma44,0,0,(0.5/(ma44))*ma51)") & ",0)) as ma35,sum(nvl(ma47,0)) as ma47,sum(nvl(ma51,0)) as ma51 from monthassess,staff where ma01=st01(+) and ma02>=" & BeginDay & " and ma02<=" & EndDay & " " & strSQL1
      strSql = strSql & " group by st01,st02,ma03) AAA ,("
      strSql = strSql & " select * from engineerassess where ea02=" & Val(frm090618.txt1(0)) + 1911 & " and ea03=" & frm090618.txt1(1) & "  ) BBB,("
      'edit by nickc 2005/04/20 加判斷完稿日
      'strSQL = strSQL & " select cp14,sum(nvl(cp97,0) * nvl(cp98,0)) as CancelCount from caseprogress where cp57>=" & BeginDayCP & " and cp57<=" & EndDayCP & " group by cp14) CCC where AAA.st01=APE.pe01(+) and AAA.st01=BBB.ea01(+) and AAA.st01=CCC.cp14(+)  order by st01"
      'edit by nickc 2005/04/27 加銷案點數
      'strSQL = strSQL & " select cp14,sum(decode(ep09,null,0,nvl(cp97,0) * nvl(cp98,0))) as CancelCount from caseprogress,engineerprogress where cp09=ep02(+) and cp57>=" & BeginDayCP & " and cp57<=" & EndDayCP & " group by cp14) CCC where AAA.st01=APE.pe01(+) and AAA.st01=BBB.ea01(+) and AAA.st01=CCC.cp14(+)  order by st01"
      'edit by nickc 2005/05/13
      'strSQL = strSQL & " select cp14,sum(decode(ep09,null,0,nvl(cp97,0) * nvl(cp98,0))) as CancelCount,sum(decode(ep09,null,0,cp18-(A1u.a1u07/1000))) as CancelPoint from caseprogress,engineerprogress,(select a1u03,sum(a1u07) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp57>=" & BeginDayCP & " and cp57<=" & EndDayCP & " ) group by a1u03) A1u where cp09=ep02(+) and cp57>=" & BeginDayCP & " and cp57<=" & EndDayCP & " and cp09=A1u.a1u03(+)  group by cp14) CCC where AAA.st01=APE.pe01(+) and AAA.st01=BBB.ea01(+) and AAA.st01=CCC.cp14(+)  order by st01"
      strSql = strSql & " select cp14,sum(decode(ep09,null,0,nvl(cp97,0) * nvl(cp98,0))) as CancelCount,sum(decode(ep09,null,0,cp18-(A1u.a1u07/1000))) as CancelPoint from caseprogress,engineerprogress,(select a1u03,sum(a1u07) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp57>=" & BeginDayCP & " and cp57<=" & EndDayCP & " and cp27 is null) group by a1u03) A1u where cp09=ep02(+) and cp57>=" & BeginDayCP & " and cp57<=" & EndDayCP & " and cp09=A1u.a1u03(+)  group by cp14) CCC where AAA.st01=APE.pe01(+) and AAA.st01=BBB.ea01(+) and AAA.st01=CCC.cp14(+)  order by st01"
Else
      'edit by nickc 2005/04/27 加銷案點數
      'strSQL = " select  nvl(da04,A3),nvl(da05,A4),nvl(da06,A5),nvl(da07,ma37),nvl(da08,decode(A3,0,0,round(ma37/A3 * 100,2))),nvl(da09,0),nvl(da10,ma47),nvl(da11,decode(A4,0,0,round(ma47/A4 * 100,2))),nvl(da12,0),nvl(da13,ma40),nvl(da14,decode(A5,0,0,round(ma40/A5 * 100,2))),nvl(da15,0),nvl(da16,round(ma35/2/" & CalMonth & ",2)),da17,da19,da20,da21,st02,nvl(da22,ma53),st01,nvl(da18,CancelCount) from (select pe01,sum(nvl(decode(pe02,'CFP',pe05 * 2 ,pe05),0) + nvl(decode(pe02,'CFP',pe07*2,pe07),0)) as A1, sum(nvl(pe06,0) + nvl(pe08,0)) as A2,sum(nvl(pe09,0)) as A3,sum(nvl(pe10,0)) as A4,sum(nvl(pe11,0)) as A5 from performance where pe03>=" & BeginDay & " and pe03<=" & EndDay & " " & strSQL2 & " group by pe01) APE ,("
      strSql = " select  nvl(da04,A3),nvl(da05,A4),nvl(da06,A5),nvl(da07,ma37),nvl(da08,decode(A3,0,0,round(ma37/A3 * 100,2))),nvl(da09,0),nvl(da10,ma47),nvl(da11,decode(A4,0,0,round(ma47/A4 * 100,2))),nvl(da12,0),nvl(da13,ma40),nvl(da14,decode(A5,0,0,round(ma40/A5 * 100,2))),nvl(da15,0),nvl(da16,round(ma35/2/" & CalMonth & ",2)),da17,da19,da20,da21,st02,nvl(da22,ma53),st01,nvl(da18,CancelCount),nvl(da23,CancelPoint) from (select pe01,sum(nvl(decode(pe02,'CFP',pe05 * 2 ,pe05),0) + nvl(decode(pe02,'CFP',pe07*2,pe07),0)) as A1, sum(nvl(pe06,0) + nvl(pe08,0)) as A2,sum(nvl(pe09,0)) as A3,sum(nvl(pe10,0)) as A4,sum(nvl(pe11,0)) as A5 from performance where pe02 in ('P','CFP') And pe03>=" & BeginDay & " and pe03<=" & EndDay & " " & strSQL2 & " group by pe01) APE ,("
      strSql = strSql & " select st01,st02,ma03,sum(nvl(ma04,0)) as ma04,sum(nvl(ma37,0)) as ma37,sum(nvl(ma40,0)) as ma40,sum(nvl(ma43,0)) as ma43,sum(nvl(ma35 - decode(ma44,0,0," & IIf(Val(BeginDayCP) >= Val(PUB_108RuleDate), "0.5*ma51", "(0.5/(ma44))*ma51") & "),0)) as ma35,sum(nvl(ma47,0)) as ma47,sum(nvl(ma51,0)) as ma51,sum(nvl(ma52,0)) as ma52,sum(nvl(ma53,0)) as ma53 from monthassess,staff where ma01=st01(+) and ma02>=" & BeginDay & " and ma02<=" & EndDay & " " & strSQL1
      strSql = strSql & " group by st01,st02,ma03) AAA,("
      strSql = strSql & " select * from drawassess where da02=" & Val(frm090618.txt1(0)) + 1911 & " and da03=" & frm090618.txt1(1) & " ) BBB,("
      'edit by nickc 2005/05/13
      'strSQL = strSQL & " select ep13,sum(decode(ep18,null,decode(ep15,null,0,decode(ep20,null,nvl(cp100,0) * nvl(cp101,0),0)),decode(ep29,null,nvl(cp103,0) * nvl(cp104,0),decode(ep15,null,0,decode(ep20,null,nvl(cp100,0) * nvl(cp101,0),0))))) as CancelCount,sum(decode(ep18,null,decode(ep15,null,0,decode(ep20,null,cp18-(A1u.a1u07/1000),0)),decode(ep29,null,cp18-(A1u.a1u07/1000),decode(ep15,null,0,decode(ep20,null,cp18-(A1u.a1u07/1000),0))))) as CancelPoint  from caseprogress ,engineerprogress,(select a1u03,sum(a1u07) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp57>=" & BeginDayCP & " and cp57<=" & EndDayCP & ") group by a1u03) A1u where cp09=ep02(+) and ep13 is not null and cp57>=" & BeginDayCP & " and cp57<=" & EndDayCP & " and cp09=A1u.a1u03(+) group by ep13) CCC where AAA.st01=APE.pe01(+) and AAA.st01=BBB.da01(+) and AAA.st01=CCC.ep13(+) order by st01 "
      strSql = strSql & " select ep13,sum(decode(ep18,null,decode(ep15,null,0,decode(ep20,null,nvl(cp100,0) * nvl(cp101,0),0)),decode(ep29,null,nvl(cp103,0) * nvl(cp104,0),decode(ep15,null,0,decode(ep20,null,nvl(cp100,0) * nvl(cp101,0),0))))) as CancelCount,sum(decode(ep18,null,decode(ep15,null,0,decode(ep20,null,cp18-(A1u.a1u07/1000),0)),decode(ep29,null,cp18-(A1u.a1u07/1000),decode(ep15,null,0,decode(ep20,null,cp18-(A1u.a1u07/1000),0))))) as CancelPoint  from caseprogress ,engineerprogress,(select a1u03,sum(a1u07) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where cp57>=" & BeginDayCP & " and cp57<=" & EndDayCP & " and cp27 is null ) group by a1u03) A1u where cp09=ep02(+) and ep13 is not null and cp57>=" & BeginDayCP & " and cp57<=" & EndDayCP & " and cp09=A1u.a1u03(+) group by ep13) CCC where AAA.st01=APE.pe01(+) and AAA.st01=BBB.da01(+) and AAA.st01=CCC.ep13(+) order by st01 "
End If
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then

      Set grd1.Recordset = adoRecordset
      
      '算得分
      strSql = "select * from assessrate where ar01 in (select max(ar01) from assessrate where ar01<=" & BeginDayCP & " ) "
      CheckOC33
      AdoRecordSet33.CursorLocation = adUseClient
      AdoRecordSet33.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet33.RecordCount <> 0 Then
            '寫入暫存季考核資料
            m_AR02 = Val("" & AdoRecordSet33.Fields("ar02").Value)
            m_AR03 = Val("" & AdoRecordSet33.Fields("ar03").Value)
            m_AR04 = Val("" & AdoRecordSet33.Fields("ar04").Value)
            m_AR05 = Val("" & AdoRecordSet33.Fields("ar05").Value)
            m_AR06 = Val("" & AdoRecordSet33.Fields("ar06").Value)
            m_AR07 = Val("" & AdoRecordSet33.Fields("ar07").Value)
            m_AR08 = Val("" & AdoRecordSet33.Fields("ar08").Value)
            m_AR13 = Val("" & AdoRecordSet33.Fields("ar13").Value)
            m_AR14 = Val("" & AdoRecordSet33.Fields("ar14").Value)
            m_AR15 = Val("" & AdoRecordSet33.Fields("ar15").Value)
            m_AR16 = Val("" & AdoRecordSet33.Fields("ar16").Value)
            m_AR17 = Val("" & AdoRecordSet33.Fields("ar17").Value)
            m_AR18 = Val("" & AdoRecordSet33.Fields("ar18").Value)
            m_AR19 = Val("" & AdoRecordSet33.Fields("ar19").Value)
            With grd1
                  For j = 2 To grd1.Rows - 1
                     If ProSysState = "1" Then '承辦人
                        .TextMatrix(j, 0) = .TextMatrix(j, 15)
                        
                        MaxD = Val("" & AdoRecordSet33.Fields("ar05").Value)
                        MinD = MaxD * -1
                        lbl3(0) = "± " & Trim(MaxD)
                        MaxE = Val("" & AdoRecordSet33.Fields("ar06").Value)
                        MinE = MaxE * -1
                        lbl3(1) = "± " & Trim(MaxE) & " %"
                        MaxF = Val("" & AdoRecordSet33.Fields("ar07").Value)
                        MinF = 0
                        lbl3(2) = "0~" & Trim(MaxF)
                        MaxG = Val("" & AdoRecordSet33.Fields("ar08").Value)
                        MinG = MaxG * -1
                        lbl3(3) = "± " & Trim(MaxG) & " %"
                        
                        '發文件數得分    .TextMatrix(j, 16)==>已填過的修改及複雜件數
                        'Modified by Morgan2019/2/19  判斷及公式有錯,且此處應顯示發文基數得分即可(若要顯示考核得分則基數與達成率也要改)--有跟柄佑確認
                        'If (Val(.TextMatrix(j, 4)) + Val(.TextMatrix(j, 16))) < 100 Then
                        '   .TextMatrix(j, 5) = Format(((Val(.TextMatrix(j, 4)) + (Val(.TextMatrix(j, 16)) / 4)) / 100) ^ 2 * 0.8 * (AdoRecordSet33.Fields("ar02").Value), "0.00")
                        'Else
                        '   .TextMatrix(j, 5) = Format(((Val(.TextMatrix(j, 4)) + (Val(.TextMatrix(j, 16)) / 4)) / 100) * 0.8 * (AdoRecordSet33.Fields("ar02").Value), "0.00")
                        'End If
                        'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式)
                        'If Val(.TextMatrix(j, 4)) < 100 Then
                        If Val(.TextMatrix(j, 4)) < 100 And Val(BeginDayCP) < Val(PUB_108RuleDate) Then
                        'end 2019/3/20
                           .TextMatrix(j, 5) = Format((Val(.TextMatrix(j, 4)) / 100) ^ 2 * 0.8 * (AdoRecordSet33.Fields("ar02").Value), "0.00")
                        Else
                           .TextMatrix(j, 5) = Format((Val(.TextMatrix(j, 4)) / 100) * 0.8 * (AdoRecordSet33.Fields("ar02").Value), "0.00")
                        End If
                        'end 2019/2/18
                        
                        '發文點數得分
                        'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式)
                        'If Val(.TextMatrix(j, 7)) < 100 Then
                        If Val(.TextMatrix(j, 7)) < 100 And Val(BeginDayCP) < Val(PUB_108RuleDate) Then
                        'end 2019/3/20
                           .TextMatrix(j, 8) = Format((Val(.TextMatrix(j, 7)) / 100) ^ 2 * 0.8 * (AdoRecordSet33.Fields("ar03").Value), "0.00")
                        Else
                           .TextMatrix(j, 8) = Format((Val(.TextMatrix(j, 7)) / 100) * 0.8 * (AdoRecordSet33.Fields("ar03").Value), "0.00")
                           '點數有上限
                           If Val(.TextMatrix(j, 8)) > ((AdoRecordSet33.Fields("ar03").Value) * 1.5) Then
                              .TextMatrix(j, 8) = Format((AdoRecordSet33.Fields("ar03").Value) * 1.5, "0.00")
                           End If
                        End If
                        '速度考核最低分是 0 分
                        If Val(.TextMatrix(j, 9)) < 0 Then
                              .TextMatrix(j, 9) = "0.00"
                        End If
'                        '考核得分
'                         .TextMatrix(j, 14) = Format((((Val(.TextMatrix(j, 5)) + Val(.TextMatrix(j, 8)) + Val(.TextMatrix(j, 9)) + Val(.TextMatrix(j, 10))) * (1 + Val(.TextMatrix(j, 11))) + Val(.TextMatrix(j, 12))) * (1 + Val(.TextMatrix(j, 13)))), "0.00")
                         '主管是抓正確，不用比對人
                         lblst01.Caption = .TextMatrix(j, 17)
                         Process j
                         'Modified by Morgan 2023/3/25
                         'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "00" Then
                         'Modified by Morgan 2025/2/4 +P10部門
                         If Pub_strUserST05 = "71" Or Pub_strUserST05 = "72" Or Pub_strUserST05 = "00" Or Pub_StrUserSt03 = "P10" Then
                         'end 2023/3/25
                         Else
                              .TextMatrix(j, 11) = txt1(2).Text
                              .TextMatrix(j, 12) = txt1(3).Text
                         End If
                         .TextMatrix(j, 14) = lbl2(23).Caption
                     Else
                        .TextMatrix(j, 0) = .TextMatrix(j, 18)
                        
                        MaxD = Val("" & AdoRecordSet33.Fields("ar17").Value)
                        MinD = MaxD * -1
                        lbl3(0) = "± " & Trim(MaxD)
                        MaxF = Val("" & AdoRecordSet33.Fields("ar18").Value)
                        MinF = 0
                        lbl3(2) = "0~" & Trim(MaxF)
                        MaxG = Val("" & AdoRecordSet33.Fields("ar19").Value)
                        MinG = MaxG * -1
                        lbl3(3) = "± " & Trim(MaxG) & " %"
                        
                        '發文件數得分
                        
                        If Trim(.TextMatrix(j, 19)) = "" Then   '未填過修改，或未存檔
                              'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式),公式有錯一併修改
                              'If  (Val(.TextMatrix(j, 5)) + Val(.TextMatrix(j, 21))) < 100 Then
                              '   .TextMatrix(j, 6) = Format(((Val(.TextMatrix(j, 5)) + (Val(.TextMatrix(j, 21)) / 4)) / 100) ^ 2 * 0.8 * (AdoRecordSet33.Fields("ar13").Value), "0.00")
                              'Else
                              '   .TextMatrix(j, 6) = Format(((Val(.TextMatrix(j, 5)) + (Val(.TextMatrix(j, 21)) / 4)) / 100) * 0.8 * (AdoRecordSet33.Fields("ar13").Value), "0.00")
                              'End If
                              If Val(.TextMatrix(j, 5)) < 100 And Val(BeginDayCP) < Val(PUB_108RuleDate) Then
                                 .TextMatrix(j, 6) = Format((Val(.TextMatrix(j, 5)) / 100) ^ 2 * 0.8 * (AdoRecordSet33.Fields("ar13").Value), "0.00")
                              Else
                                 .TextMatrix(j, 6) = Format((Val(.TextMatrix(j, 5)) / 100) * 0.8 * (AdoRecordSet33.Fields("ar13").Value), "0.00")
                              End If
                              'end 2019/3/20
                         Else
                              'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式),公式有錯一併修改
                              'If  (Val(.TextMatrix(j, 5)) + Val(.TextMatrix(j, 19))) < 100 Then
                              '   .TextMatrix(j, 6) = Format(((Val(.TextMatrix(j, 5)) + (Val(.TextMatrix(j, 19)) / 4)) / 100) ^ 2 * 0.8 * (AdoRecordSet33.Fields("ar13").Value), "0.00")
                              'Else
                              '   .TextMatrix(j, 6) = Format(((Val(.TextMatrix(j, 5)) + (Val(.TextMatrix(j, 19)) / 4)) / 100) * 0.8 * (AdoRecordSet33.Fields("ar13").Value), "0.00")
                              'End If
                              If Val(.TextMatrix(j, 5)) < 100 And Val(BeginDayCP) < Val(PUB_108RuleDate) Then
                                 .TextMatrix(j, 6) = Format((Val(.TextMatrix(j, 5)) / 100) ^ 2 * 0.8 * (AdoRecordSet33.Fields("ar13").Value), "0.00")
                              Else
                                 .TextMatrix(j, 6) = Format((Val(.TextMatrix(j, 5)) / 100) * 0.8 * (AdoRecordSet33.Fields("ar13").Value), "0.00")
                              End If
                              'end 2019/3/20
                         End If
                        '發文張數得分
                        'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式)
                        'If Val(.TextMatrix(j, 8)) < 100 Then
                        If Val(.TextMatrix(j, 8)) < 100 And Val(BeginDayCP) < Val(PUB_108RuleDate) Then
                        'end 2019/3/20
                           .TextMatrix(j, 9) = Format((Val(.TextMatrix(j, 8)) / 100) ^ 2 * 0.8 * (AdoRecordSet33.Fields("ar14").Value), "0.00")
                        Else
                           .TextMatrix(j, 9) = Format((Val(.TextMatrix(j, 8)) / 100) * 0.8 * (AdoRecordSet33.Fields("ar14").Value), "0.00")
                        End If
                        '發文點數得分
                        'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式)
                        'If Val(.TextMatrix(j, 11)) < 100 Then
                        If Val(.TextMatrix(j, 11)) < 100 And Val(BeginDayCP) < Val(PUB_108RuleDate) Then
                        'end 2019/3/20
                           .TextMatrix(j, 12) = Format((Val(.TextMatrix(j, 11)) / 100) ^ 2 * 0.8 * (AdoRecordSet33.Fields("ar15").Value), "0.00")
                        Else
                           .TextMatrix(j, 12) = Format((Val(.TextMatrix(j, 11)) / 100) * 0.8 * (AdoRecordSet33.Fields("ar15").Value), "0.00")
                           '點數有上限
                           If Val(.TextMatrix(j, 12)) > ((AdoRecordSet33.Fields("ar15").Value) * 1.5) Then
                              .TextMatrix(j, 12) = Format((AdoRecordSet33.Fields("ar15").Value) * 1.5, "0.00")
                           End If
                        End If
                        '速度考核最低分是 0 分
                       If Val(.TextMatrix(j, 13)) < 0 Then
                              .TextMatrix(j, 13) = "0.00"
                        End If
                        '考核得分
'                        .TextMatrix(j, 17) = Format((Val(.TextMatrix(j, 6)) + Val(.TextMatrix(j, 9)) + Val(.TextMatrix(j, 12)) + Val(.TextMatrix(j, 13)) + Val(.TextMatrix(j, 14)) + Val(.TextMatrix(j, 15))) * (1 + Val(.TextMatrix(j, 16))), "0.00")
                        '主管是抓正確，不用比對人
                        lblst01.Caption = .TextMatrix(j, 20)
                        Process j
                        'Modified by Morgan 2023/3/25
                        'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "00" Then
                        'Modified by Morgan 2025/2/4 +P10部門
                        If Pub_strUserST05 = "71" Or Pub_strUserST05 = "72" Or Pub_strUserST05 = "00" Or Pub_StrUserSt03 = "P10" Then
                        'end 2023/3/25
                        Else
                              .TextMatrix(j, 15) = txt1(3).Text
                        End If
                        .TextMatrix(j, 17) = lbl2(23).Caption
                     End If
                  Next j
            End With
       End If
      If ProSysState = "1" Then '承辦人
         grd1.col = 14
      Else
         grd1.col = 17
      End If
      grd1.Sort = 4
      SetGrd1
      If frm090618.txt1(4).Text = "2" Then '列印
         PrintData
         StrMenu = False
      End If
Else
   ShowNoData
   StrMenu = False
End If
End Function

Sub PrintData()
Dim iCol As Integer
Dim iRow As Integer
iPrint = 0
Page = 1
GetPleft
If ProSysState = "1" Then
   Printer.FontSize = 12
Else
   Printer.FontSize = 10
End If
PrintTitle
With grd1
   For iRow = 2 To .Rows - 1
      .row = iRow
      Process iRow
      For iCol = 0 To .Cols - IIf(ProSysState = "1", 6, 7)
         If iCol = 0 Then
            Printer.CurrentX = PLeft(iCol)
            Printer.CurrentY = iPrint
            Printer.Print .TextMatrix(iRow, iCol)
         Else
            'edit by nickc 2005/04/21 考核件數，抓詳細
            If (iCol = 4 And ProSysState = "2") Or (iCol = 3 And ProSysState = "1") Then
               Printer.CurrentX = PLeft(iCol) + 500 - Printer.TextWidth(Format(Val(lbl2(8).Caption), "0.00"))
               Printer.CurrentY = iPrint
               Printer.Print Format(Val(lbl2(8).Caption), "0.00")
            ElseIf (iCol = 5 And ProSysState = "2") Or (iCol = 4 And ProSysState = "1") Then
               Printer.CurrentX = PLeft(iCol) + 500 - Printer.TextWidth(Format(Val(lbl2(9).Caption), "0.00"))
               Printer.CurrentY = iPrint
               Printer.Print Format(Val(lbl2(9).Caption), "0.00")
            ElseIf (iCol = 6 And ProSysState = "2") Or (iCol = 5 And ProSysState = "1") Then
               Printer.CurrentX = PLeft(iCol) + 500 - Printer.TextWidth(Format(Val(lbl2(10).Caption), "0.00"))
               Printer.CurrentY = iPrint
               Printer.Print Format(Val(lbl2(10).Caption), "0.00")
            Else
               Printer.CurrentX = PLeft(iCol) + 500 - Printer.TextWidth(Format(Val(.TextMatrix(iRow, iCol)), "0.00"))
               Printer.CurrentY = iPrint
               Printer.Print Format(Val(.TextMatrix(iRow, iCol)), "0.00")
            End If
         End If
      Next iCol
      iPrint = iPrint + 300
      If iPrint >= 10000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle
      End If
   Next iRow
End With
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft()
Erase PLeft
'定陣列
If ProSysState = "1" Then '承辦人
      PLeft(0) = 500 - 500
      PLeft(1) = 1500 - 500
      PLeft(2) = 2500 - 500
      PLeft(3) = 3500 - 500
      PLeft(4) = 4500 - 500
      PLeft(5) = 5500 - 500
      PLeft(6) = 6500 - 500
      PLeft(7) = 7500 - 500
      PLeft(8) = 8500 - 500
      PLeft(9) = 9500 - 500
      PLeft(10) = 10600 - 500
      PLeft(11) = 11700 - 500
      PLeft(12) = 12800 - 500
      PLeft(13) = 13900 - 500
      PLeft(14) = 15000 - 500
Else
      PLeft(0) = 500 - 500
      PLeft(1) = 1200 + 300 - 500
      PLeft(2) = 2100 + 300 - 500
      PLeft(3) = 3000 + 300 - 500
      PLeft(4) = 3900 + 300 - 500
      PLeft(5) = 4800 + 300 - 500
      PLeft(6) = 5700 + 300 - 500
      PLeft(7) = 6600 + 300 - 500
      PLeft(8) = 7500 + 300 - 500
      PLeft(9) = 8400 + 300 - 500
      PLeft(10) = 9300 + 300 - 500
      PLeft(11) = 10200 + 300 - 500
      PLeft(12) = 11100 + 300 - 500
      PLeft(13) = 12000 + 300 - 500
      PLeft(14) = 12900 + 300 - 500
      PLeft(15) = 13800 + 300 - 500
      PLeft(16) = 14700 + 300 - 500
      PLeft(17) = 15600 + 300 - 500
End If
End Sub

Sub PrintTitle() '列印抬頭
iPrint = 0
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print IIf(ProSysState = "1", "承辦人", "繪圖人員") & "季考核表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
'Printer.Print "年月：" & Format(Format(str(Val(frm090618.txt1(0)) + 191100) & "01", "####/##/##"), "ee/MM") & "-" & Format(Format(str(Val(frm090618.txt1(1)) + 191100) & "01", "####/##/##"), "ee/MM")
Printer.Print "季別：" & str(Val(frm090618.txt1(0))) & " 年，第" & str(Val(frm090618.txt1(1))) & " 季 "
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
If ProSysState = "1" Then
   Printer.CurrentX = 12500
Else
   Printer.CurrentX = 13800
End If
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print IIf(ProSysState = "1", "承辦人", "繪圖人員") & "：" & IIf(Trim(frm090618.lbl1.Caption) = "", "所有", frm090618.lbl1.Caption)
If ProSysState = "1" Then
   Printer.CurrentX = 12500
Else
   Printer.CurrentX = 13800
End If
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
If ProSysState = "1" Then
   Printer.FontSize = 12
Else
   Printer.FontSize = 10
End If
If ProSysState = "1" Then
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print "承辦人"
      Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1)) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "目標"
      Printer.Line (PLeft(1), iPrint + 290)-(PLeft(3) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(3) + ((PLeft(5) - PLeft(3)) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "考核"
      Printer.Line (PLeft(3), iPrint + 290)-(PLeft(6) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(6) + ((PLeft(8) - PLeft(6)) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "發文"
      Printer.Line (PLeft(6), iPrint + 290)-(PLeft(9) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(9)
      Printer.CurrentY = iPrint
      Printer.Print "速度考核"
      Printer.CurrentX = PLeft(10)
      Printer.CurrentY = iPrint
      Printer.Print "自我比較"
      Printer.CurrentX = PLeft(11)
      Printer.CurrentY = iPrint
      Printer.Print "核稿人"
      Printer.CurrentX = PLeft(12)
      Printer.CurrentY = iPrint
      Printer.Print "帶人主管"
      Printer.CurrentX = PLeft(13)
      Printer.CurrentY = iPrint
      Printer.Print "部門主管"
      Printer.CurrentX = PLeft(14)
      Printer.CurrentY = iPrint
      Printer.Print "季考核"
      iPrint = iPrint + 300
      If iPrint >= 9000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle
          Exit Sub
      End If
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print ""
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iPrint
      Printer.Print "基數"
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = iPrint
      Printer.Print "點數"
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = iPrint
      Printer.Print "基數"
      Printer.CurrentX = PLeft(4)
      Printer.CurrentY = iPrint
      Printer.Print "達成率%"
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = iPrint
      Printer.Print "得分"
      Printer.CurrentX = PLeft(6)
      Printer.CurrentY = iPrint
      Printer.Print "點數"
      Printer.CurrentX = PLeft(7)
      Printer.CurrentY = iPrint
      Printer.Print "達成率%"
      Printer.CurrentX = PLeft(8)
      Printer.CurrentY = iPrint
      Printer.Print "得分"
      Printer.CurrentX = PLeft(9)
      Printer.CurrentY = iPrint
      Printer.Print "得分"
      Printer.CurrentX = PLeft(10)
      Printer.CurrentY = iPrint
      Printer.Print "評分"
      Printer.CurrentX = PLeft(11)
      Printer.CurrentY = iPrint
      Printer.Print "評分"
      Printer.CurrentX = PLeft(12)
      Printer.CurrentY = iPrint
      Printer.Print "評分"
      Printer.CurrentX = PLeft(13)
      Printer.CurrentY = iPrint
      Printer.Print "評分"
      Printer.CurrentX = PLeft(14)
      Printer.CurrentY = iPrint
      Printer.Print "總分"
      iPrint = iPrint + 300
      If iPrint >= 9000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle
          Exit Sub
      End If
      ShowLine
   If iPrint >= 9000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       Exit Sub
   End If
Else
      
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print "繪圖"
      Printer.CurrentX = PLeft(1) + ((PLeft(3) - PLeft(1)) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "目標"
      Printer.Line (PLeft(1), iPrint + 290)-(PLeft(4) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(4) + ((PLeft(7) - PLeft(4)) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "考核"
      Printer.Line (PLeft(4), iPrint + 290)-(PLeft(7) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(7) + ((PLeft(13) - PLeft(8)) / 2)
      Printer.CurrentY = iPrint
      Printer.Print "發文"
      Printer.Line (PLeft(7), iPrint + 290)-(PLeft(13) - 100, iPrint + 290)
      Printer.CurrentX = PLeft(13)
      Printer.CurrentY = iPrint
      Printer.Print "速度考核"
      Printer.CurrentX = PLeft(14)
      Printer.CurrentY = iPrint
      Printer.Print "自我比較"
      Printer.CurrentX = PLeft(15)
      Printer.CurrentY = iPrint
      Printer.Print "帶人主管"
      Printer.CurrentX = PLeft(16)
      Printer.CurrentY = iPrint
      Printer.Print "部門主管"
      Printer.CurrentX = PLeft(17)
      Printer.CurrentY = iPrint
      Printer.Print "季考核"
      iPrint = iPrint + 300
      If iPrint >= 9000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle
          Exit Sub
      End If
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print "人員"
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iPrint
      Printer.Print "基數"
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = iPrint
      Printer.Print "張數"
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = iPrint
      Printer.Print "點數"
      Printer.CurrentX = PLeft(4)
      Printer.CurrentY = iPrint
      Printer.Print "基數"
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = iPrint
      Printer.Print "達成率%"
      Printer.CurrentX = PLeft(6)
      Printer.CurrentY = iPrint
      Printer.Print "得分"
      Printer.CurrentX = PLeft(7)
      Printer.CurrentY = iPrint
      Printer.Print "張數"
      Printer.CurrentX = PLeft(8)
      Printer.CurrentY = iPrint
      Printer.Print "達成率%"
      Printer.CurrentX = PLeft(9)
      Printer.CurrentY = iPrint
      Printer.Print "得分"
      Printer.CurrentX = PLeft(10)
      Printer.CurrentY = iPrint
      Printer.Print "點數"
      Printer.CurrentX = PLeft(11)
      Printer.CurrentY = iPrint
      Printer.Print "達成率%"
      Printer.CurrentX = PLeft(12)
      Printer.CurrentY = iPrint
      Printer.Print "得分"
      Printer.CurrentX = PLeft(13)
      Printer.CurrentY = iPrint
      Printer.Print "得分"
      Printer.CurrentX = PLeft(14)
      Printer.CurrentY = iPrint
      Printer.Print "評分"
      Printer.CurrentX = PLeft(15)
      Printer.CurrentY = iPrint
      Printer.Print "評分"
      Printer.CurrentX = PLeft(16)
      Printer.CurrentY = iPrint
      Printer.Print "評分"
      Printer.CurrentX = PLeft(17)
      Printer.CurrentY = iPrint
      Printer.Print "得分"
      iPrint = iPrint + 300
      If iPrint >= 9000 Then
          Page = Page + 1
          Printer.NewPage
          PrintTitle
          Exit Sub
      End If
      ShowLine
   If iPrint >= 9000 Then
       Page = Page + 1
       Printer.NewPage
       PrintTitle
       Exit Sub
   End If
End If
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
If ProSysState = "1" Then
   Printer.Line (0, iPrint + 150)-(16500, iPrint + 150)
Else
   Printer.Line (0, iPrint + 150)-(16500, iPrint + 150)
End If
iPrint = iPrint + 300
End Sub

Private Sub GRD1_DblClick()

Me.Enabled = False
Screen.MousePointer = vbHourglass
    If Me.grd1.MouseRow > 1 Then
        If Me.grd1.Rows > 2 Then
            SWPRow = str(grd1.MouseRow)
            Process Val(SWPRow)
            SSTab1.Tab = 1
        End If
    End If
Screen.MousePointer = vbDefault
Me.Enabled = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Strindex As Integer
Dim j As Integer
Dim oMouseCol As Integer
If Me.grd1.MouseRow < 0 Then Exit Sub
If Button = 1 Then
    Screen.MousePointer = vbHourglass
    SWPRow = str(grd1.MouseRow)
    oMouseCol = grd1.MouseCol
    If Val(SWPRow) < 2 Then
        Select Case oMouseCol
        Case 0
            If m_blnColOrderAsc = True Then
                Me.grd1.Sort = 5 '昇冪
                m_blnColOrderAsc = False
            Else
                Me.grd1.Sort = 6 '降冪
                m_blnColOrderAsc = True
            End If
        Case Else
            If m_blnColOrderAsc = True Then
                Me.grd1.Sort = 3 '昇冪
                m_blnColOrderAsc = False
            Else
                Me.grd1.Sort = 4 '降冪
                m_blnColOrderAsc = True
            End If
        End Select
    End If
    Strindex = SWPRow
    With grd1
        DoEvents
        .Visible = False
         If Val(SWPRow) = 0 Or Val(SWPRow) = 1 Then
            For j = 2 To .Rows - 1
               .row = j
               .col = 1
               If .CellBackColor = &HFFC0C0 Then
                  SWPRow2 = j
                  .Visible = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            Next j
         End If
        If SWPRow2 <> "" Then
           .row = SWPRow2
           For j = 1 To .Cols - 1
               .col = j
               .CellBackColor = QBColor(15)
           Next j
        End If
        .col = 0
        If Strindex <> 0 Then
            .row = Strindex
        Else
            .row = .MouseRow
        End If
        If .row = 0 Or .row = 1 Then
            .row = 2
        End If
         For j = 1 To .Cols - 1
             .col = j
             .CellBackColor = &HFFC0C0
         Next j
        SWPColor2 = SWPColor
        SWPRow2 = .row
        Process Val(SWPRow2)
        .Visible = True
    End With
    Screen.MousePointer = vbDefault
End If
End Sub
'帶單筆資料
Sub Process(oRow As Integer)
With grd1
      If ProSysState = "1" Then '承辦人
         Label4.Visible = False
         lbl2(3).Visible = False
         lbl2(11).Visible = False
         lbl2(12).Visible = False
         lbl2(13).Visible = False
         Label9.Visible = False
         Label10.Visible = False
         Label11.Visible = False
         Label21.Visible = False
         lbl2(17).Visible = False
         lblst01.Visible = False
         lblst01.Caption = .TextMatrix(oRow, 17)
         lblPromoter.Caption = .TextMatrix(oRow, 0)
         lbl2(1).Caption = frm090618.txt1(0) & " 年 " & frm090618.txt1(1) & "  季 "
         lbl2(2).Caption = .TextMatrix(oRow, 1)
         lbl2(4).Caption = .TextMatrix(oRow, 2)
         lbl2(5).Caption = .TextMatrix(oRow, 3)
         lbl2(6).Caption = .TextMatrix(oRow, 16)
         '增加銷案件數
         lbl2(24).Caption = .TextMatrix(oRow, 18)
         'add by nickc 2005/04/27  加銷案點數
         lbl2(25).Caption = .TextMatrix(oRow, 19)
         If Val(lbl2(6).Caption) <> 0 Then
            lbl2(7).Caption = Format(Val(lbl2(6).Caption) / Val(4), "0.00")
         Else
            lbl2(7).Caption = "0.00"
         End If
         lbl2(8).Caption = Trim(Val(lbl2(5)) + Val(lbl2(7)) + Val(lbl2(24)))
         '達成率
         If Val(lbl2(2).Caption) <> 0 Then
            lbl2(9).Caption = Format(Trim((Val(lbl2(8).Caption) / Val(lbl2(2).Caption)) * 100), "0.00")
         Else
            lbl2(9).Caption = "0"
         End If
         '考核件數得分
         'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式)
         'If Val(lbl2(9).Caption) > 100 Then
         If Val(lbl2(9).Caption) > 100 Or Val(BeginDayCP) >= Val(PUB_108RuleDate) Then
         'end 2019/3/20
            lbl2(10).Caption = Format(Val(lbl2(9).Caption) / 100 * 0.8 * m_AR02, "0.00")
         Else
            lbl2(10).Caption = Format(((Val(lbl2(9).Caption) / 100) ^ 2) * 0.8 * m_AR02, "0.00")
         End If
         '發文點數
         'edit by nickc 2005/04/27 點數要加入銷案點數，且達成率也要重算
         'lbl2(14).Caption = .TextMatrix(oRow, 6)
         'lbl2(15).Caption = .TextMatrix(oRow, 7)
         lbl2(14).Caption = Trim(Val(.TextMatrix(oRow, 6)) + Val(lbl2(25).Caption))
         '達成率
         If Val(lbl2(4).Caption) <> 0 Then
            lbl2(15).Caption = Format(Trim((Val(lbl2(14).Caption) / Val(lbl2(4).Caption)) * 100), "0.00")
         Else
            lbl2(15).Caption = "0"
         End If
         'edit end
         
         'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式)
         'If Val(lbl2(15).Caption) > 100 Then
         If Val(lbl2(15).Caption) > 100 Or Val(BeginDayCP) >= Val(PUB_108RuleDate) Then
         'end 2019/3/20
            lbl2(16).Caption = Format(Val(lbl2(15).Caption) / 100 * 0.8 * m_AR03, "0.00")
         Else
            lbl2(16).Caption = Format(((Val(lbl2(15).Caption) / 100) ^ 2) * 0.8 * m_AR03, "0.00")
         End If
         lbl2(18).Caption = .TextMatrix(oRow, 9)
         lbl2(19).Caption = .TextMatrix(oRow, 10)
         lbl2(20).Caption = .TextMatrix(oRow, 11)
         lbl2(21).Caption = .TextMatrix(oRow, 12)
         lbl2(22).Caption = .TextMatrix(oRow, 13)
         lbl2(23).Caption = .TextMatrix(oRow, 14)
         txt1(0).Text = lbl2(6).Caption
         txt1(1).Text = lbl2(19).Caption
         txt1(2).Text = lbl2(20).Caption
         txt1(3).Text = lbl2(21).Caption
         txt1(4).Text = lbl2(22).Caption
         txt1(5).Text = ""
         '若是核稿人或帶人主管時，帶資料到 txt1
         Call GetAB(lblst01.Caption, Val(frm090618.txt1(0).Text) + 1911, Val(frm090618.txt1(1).Text), "1")
         ReCal
         If ProState <> "2" Then    '個人
               grd2.Visible = False
               lbl2(22).Visible = False
               lbl2(21).Visible = False
               lbl2(20).Visible = False
               Label19.Visible = False
               Label18.Visible = False
               Label27.Visible = False
               Label17.Visible = False
               lbl3(3).Visible = False
               lbl3(2).Visible = False
               lbl3(1).Visible = False
               '2010/1/19 modify by sonia 考核季未過不可存檔
               'If lbl2(22).Caption <> "" Or lbl2(21).Caption <> "" Or lbl2(20).Caption <> "" Then   '部門主管評過，不允許修改
               If IsOverSeason Or lbl2(22).Caption <> "" Or lbl2(21).Caption <> "" Or lbl2(20).Caption <> "" Then   '部門主管評過，不允許修改
                  txt1(0).Visible = False
                  txt1(1).Visible = False
                  txt1(2).Visible = False
                  txt1(3).Visible = False
                  txt1(4).Visible = False
                  txt1(5).Visible = False
                  cmdok(1).Visible = False
               Else
                  txt1(0).Visible = True
                  txt1(1).Visible = True
                  cmdok(1).Visible = True
               End If
         Else  '管理
               '2010/1/19 add by sonia 考核季未過不可存檔
               If IsOverSeason Then
                  cmdok(1).Visible = False
               End If
               '2010/1/19 end
               'Modified by Morgan 2023/3/25
               'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "00" Then
               'Modified by Morgan 2025/2/4 +P10部門
               If Pub_strUserST05 = "71" Or Pub_strUserST05 = "72" Or Pub_strUserST05 = "00" Or Pub_StrUserSt03 = "P10" Then
               'end 2023/3/25
                   grd2.Visible = True
                    '協理
               Else
                  '核稿人或帶人主管
                  grd2.Visible = False
                 ThisUserData
                  'edit by nickc 2007/11/26 協理評過該季，不允許再動，且不可以自評
                  'If lbl2(22).Caption <> "" Then    '部門主管評過，不允許修改
                  '2010/1/19 modify by sonia 考核季未過不可存檔
                  'If IsHaveClose Or lblst01.Caption = strUserNum Then
                  If IsOverSeason Or IsHaveClose Or lblst01.Caption = strUserNum Then
                     lbl2(22).Visible = False
                     txt1(0).Visible = False
                     txt1(1).Visible = False
                     txt1(2).Visible = False
                     txt1(3).Visible = False
                     txt1(4).Visible = False
                     Label19.Visible = False
                     lbl3(3).Visible = False
                     txt1(5).Visible = False
                     cmdok(1).Visible = False
                  Else
                     txt1(0).Enabled = True
                     txt1(0).Visible = True
                     txt1(1).Enabled = True
                     txt1(1).Visible = True
                     txt1(2).Enabled = True
                     txt1(2).Visible = True
                     txt1(3).Enabled = True
                     txt1(3).Visible = True
                     txt1(5).Enabled = True
                     txt1(5).Visible = True
                     cmdok(1).Visible = True
                  End If
               End If
          End If
      Else
         lblst01.Caption = .TextMatrix(oRow, 20)
         Label1.Caption = "繪圖人員："
         Label21.Visible = False
         Label17.Visible = False
         lbl3(1).Visible = False
         lbl2(17).Visible = False
         lbl2(20).Visible = False
         lblst01.Visible = False
         txt1(2).Visible = False
         lblPromoter.Caption = .TextMatrix(oRow, 0)
         lbl2(1).Caption = frm090618.txt1(0) & " 年 " & frm090618.txt1(1) & "  季 "
         lbl2(2).Caption = .TextMatrix(oRow, 1)
         lbl2(3).Caption = .TextMatrix(oRow, 2)
         lbl2(4).Caption = .TextMatrix(oRow, 3)
         lbl2(5).Caption = .TextMatrix(oRow, 4)
         lbl2(6).Caption = .TextMatrix(oRow, 19)
         '增加銷案件數
         lbl2(24).Caption = .TextMatrix(oRow, 21)
         'add by nickc 2005/04/27 加入銷案點數
         lbl2(25).Caption = .TextMatrix(oRow, 22)
         
         If Val(lbl2(6).Caption) <> 0 Then
            lbl2(7).Caption = Format(Val(lbl2(6).Caption) / Val(4), "0.00")
         Else
            lbl2(7).Caption = "0.00"
         End If
         lbl2(8).Caption = Trim(Val(lbl2(5)) + Val(lbl2(7)) + Val(lbl2(24)))
         '達成率
         If Val(lbl2(2).Caption) <> 0 Then
            lbl2(9).Caption = Format(Trim((Val(lbl2(8).Caption) / Val(lbl2(2).Caption)) * 100), "0.00")
         Else
            lbl2(9).Caption = "100.00"
         End If
         '考核件數得分
         'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式)
         'If Val(lbl2(9).Caption) > 100 Then
         If Val(lbl2(9).Caption) > 100 Or Val(BeginDayCP) >= Val(PUB_108RuleDate) Then
         'end 2019/3/20
            lbl2(10).Caption = Format(Val(lbl2(9).Caption) / 100 * 0.8 * m_AR13, "0.00")
         Else
            lbl2(10).Caption = Format(((Val(lbl2(9).Caption) / 100) ^ 2) * 0.8 * m_AR13, "0.00")
         End If
         lbl2(11).Caption = .TextMatrix(oRow, 7)
         lbl2(12).Caption = .TextMatrix(oRow, 8)
         'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式)
         'If Val(lbl2(12).Caption) > 100 Then
         If Val(lbl2(12).Caption) > 100 Or Val(BeginDayCP) >= Val(PUB_108RuleDate) Then
         'end 2019/3/20
            lbl2(13).Caption = Format(Val(lbl2(12).Caption) / 100 * 0.8 * m_AR14, "0.00")
         Else
            lbl2(13).Caption = Format(((Val(lbl2(12).Caption) / 100) ^ 2) * 0.8 * m_AR14, "0.00")
         End If
         'edit by nickc 2005/04/27 加入銷案點數，達成率也要重算
'         lbl2(14).Caption = .TextMatrix(oRow, 10)
'         lbl2(15).Caption = .TextMatrix(oRow, 11)
         lbl2(14).Caption = Trim(Val(.TextMatrix(oRow, 10)) + Val(lbl2(25).Caption))
         If Val(lbl2(4).Caption) <> 0 Then
            lbl2(15).Caption = Format(Trim((Val(lbl2(14).Caption) / Val(lbl2(4).Caption)) * 100), "0.00")
         Else
            lbl2(15).Caption = "0"
         End If
         'edit end
         'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式)
         'If Val(lbl2(15).Caption) > 100 Then
         If Val(lbl2(15).Caption) > 100 Or Val(BeginDayCP) >= Val(PUB_108RuleDate) Then
         'end 2019/3/20
            lbl2(16).Caption = Format(Val(lbl2(15).Caption) / 100 * 0.8 * m_AR15, "0.00")
         Else
            lbl2(16).Caption = Format(((Val(lbl2(15).Caption) / 100) ^ 2) * 0.8 * m_AR15, "0.00")
         End If
         lbl2(18).Caption = .TextMatrix(oRow, 13)
         lbl2(19).Caption = .TextMatrix(oRow, 14)
         lbl2(21).Caption = .TextMatrix(oRow, 15)
         lbl2(22).Caption = .TextMatrix(oRow, 16)
         lbl2(23).Caption = .TextMatrix(oRow, 17)
         txt1(0).Text = lbl2(6).Caption
         txt1(1).Text = lbl2(19).Caption
         txt1(3).Text = lbl2(21).Caption
         txt1(4).Text = lbl2(22).Caption
         txt1(5).Text = ""
         Call GetAB(lblst01.Caption, Val(frm090618.txt1(0).Text) + 1911, Val(frm090618.txt1(1).Text), "2")
         ReCal
         If ProState <> "2" Then    '個人
               grd2.Visible = False
               lbl2(22).Visible = False
               lbl2(21).Visible = False
               lbl2(20).Visible = False
               Label19.Visible = False
               Label18.Visible = False
               Label27.Visible = False
               Label17.Visible = False
               lbl3(3).Visible = False
               lbl3(2).Visible = False
               lbl3(1).Visible = False
               '2010/1/19 modify by sonia 考核季未過不可存檔
               'If lbl2(22).Caption <> "" Or lbl2(21).Caption <> "" Or lbl2(20).Caption <> "" Then   '部門主管評過，不允許修改
               If IsOverSeason Or lbl2(22).Caption <> "" Or lbl2(21).Caption <> "" Or lbl2(20).Caption <> "" Then   '部門主管評過，不允許修改
                  txt1(0).Visible = False
                  txt1(1).Visible = False
                  txt1(3).Visible = False
                  txt1(4).Visible = False
                  txt1(5).Visible = False
                  cmdok(1).Visible = False
               Else
                  txt1(0).Visible = True
                  txt1(1).Visible = True
                  cmdok(1).Visible = True
               End If
         Else  '管理
               '2010/1/19 add by sonia 考核季未過不可存檔
               If IsOverSeason Then
                  cmdok(1).Visible = False
               End If
               '2010/1/19 end
               'Modified by Morgan 2023/3/25
               'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "00" Then
               'Modified by Morgan 2025/2/4 +P10部門
               If Pub_strUserST05 = "71" Or Pub_strUserST05 = "72" Or Pub_strUserST05 = "00" Or Pub_StrUserSt03 = "P10" Then
               'end 2023/3/25
                   grd2.Visible = True
                    '協理
               Else
                  '核稿人或帶人主管
                  '核稿人或帶人主管
                  'edit by nickc 2005/04/26
                  'modify by sonia 2016/3/3 改72006為73022
                  'Modified by Morgan 2025/2/19 改73022為87025，能看其他繪圖人員的所有主管評分--游協理
                  'If strUserNum = "73022" Then
                  If strUserNum = "87025" Then
                  'end 2025/2/19
                     grd2.Visible = True
                  Else
                     grd2.Visible = False
                  End If
                  ThisUserData
                  'edit by nickc 2007/11/26 協理評過該季，不允許再動，且不可以自評
                  'If lbl2(22).Caption <> "" Then    '部門主管評過，不允許修改
                  '2010/1/19 modify by sonia 考核季未過不可存檔
                  '2010/1/19 modify by sonia 考核季未過不可存檔
                  'If IsHaveClose Or lblst01.Caption = strUserNum Then
                  If IsOverSeason Or IsHaveClose Or lblst01.Caption = strUserNum Then
                     lbl2(22).Visible = False
                     txt1(0).Visible = False
                     txt1(1).Visible = False
                     txt1(2).Visible = False
                     txt1(3).Visible = False
                     txt1(4).Visible = False
                     txt1(5).Visible = False
                     Label19.Visible = False
                     lbl3(3).Visible = False
                     cmdok(1).Visible = False
                  Else
                     txt1(0).Enabled = True
                     txt1(0).Visible = True
                     txt1(1).Enabled = True
                     txt1(1).Visible = True
                     txt1(3).Enabled = True
                     txt1(3).Visible = True
                     txt1(5).Enabled = True
                     txt1(5).Visible = True
                     cmdok(1).Visible = True
                  End If
               End If
          End If
      End If
End With
End Sub

Private Sub grd2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grd2, x, y, nCol, nRow
grd2.col = nCol
grd2.row = nRow
End Sub


Private Sub txtInput_GotFocus()
txtInput.SelStart = 0
txtInput.SelLength = Len(txtInput)
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   Dim Cancel  As Boolean
   If KeyAscii = vbKeyReturn Then
      Cancel = False
      txtInputValidate Cancel
      If Cancel = False Then
         grd2.TextMatrix(iRow, iCol) = Format(txtInput.Text, "0.00")
         Call ReCal
         grd2.SetFocus
         grd2.Refresh
         txtInput.Visible = False
      End If
   ElseIf KeyAscii = vbKeyEscape Then
      grd2.SetFocus
      txtInput.Visible = False
   End If
   
End Sub

Private Sub txtInput_LostFocus()
   txtInput.Visible = False
   txtInput.Tag = ""
End Sub
Private Sub SetBox()
   
   Dim lngLeft As Long, lngTop As Long
   
   With grd2
      If .row > 0 And .col = 2 Then
         If .TextMatrix(.row, 2) <> "" Then
            txtInput.FontName = .CellFontName
            txtInput.FontSize = .CellFontSize
            txtInput.Alignment = .CellAlignment \ 5
            txtInput.Text = .TextMatrix(.row, .col)
            txtInput.Tag = txtInput.Text
            txtInput.Width = .ColWidth(.col)
            txtInput.Height = .RowHeight(.row)
            iRow = .row: iCol = .col
            txtInput.Visible = True
            txtInput.Enabled = True
            txtInput.SetFocus
            txtInput.SelStart = 0
            txtInput.SelLength = Len(txtInput)
            lngLeft = .Left + 25
            lngTop = .Top + 25 '+ .RowHeight(iRow)
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            txtInput.Left = lngLeft: txtInput.Top = lngTop
         End If
      End If
   End With
End Sub

Private Sub SetDataListWidth()
   
   With grd2
         .Visible = False
         .Cols = 5
         .row = 0
         .col = 0:   .Text = "主管類別"
         .ColWidth(0) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 1:   .Text = "名稱"
         .ColWidth(1) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 2:   .Text = "評分"
         .ColWidth(2) = 1200
         .CellAlignment = flexAlignCenterCenter
         .col = 3:   .Text = ""
         .ColWidth(3) = 0
         .CellAlignment = flexAlignCenterCenter
         .col = 4:   .Text = "備註"
         .ColWidth(4) = 2000
         .CellAlignment = flexAlignCenterCenter
         'Modified by Morgan 2023/3/25
         'If PUB_GetST05(strUserNum) = "71" Or PUB_GetST05(strUserNum) = "00" Then
         'Modified by Morgan 2025/2/4 +P10部門
         If Pub_strUserST05 = "71" Or Pub_strUserST05 = "72" Or Pub_strUserST05 = "00" Or Pub_StrUserSt03 = "P10" Then
         'end 2023/3/25
            .Visible = True
         'modify by sonia 2016/3/3 改72006為73022
         'Modified by Morgan 2025/2/20 改73022為87025
         ElseIf strUserNum = "87025" Then
            For i = 1 To .Rows - 1
               'Removed by Morgan 2022/8/2 游經理應該要能所有繪圖主管的評分
               'Modified by Morgan 2025/2/20 除了自己的不能看，開放能看其他繪圖人員的所有主管評分--游協理
               'If .TextMatrix(i, 3) <> "78007" And .TextMatrix(i, 3) <> "82018" And .TextMatrix(i, 3) <> "72006" Then
               If lblst01.Caption = strUserNum Then
                     .RowHeight(i) = 0
               End If
               'end 2025/2/20
            Next i
            'end 2022/8/2
            .Visible = True
         End If
   End With
End Sub

Private Sub grd2_DblClick()
'modify by sonia 2016/3/3 改72006為73022
If strUserNum = "73022" Then Exit Sub

Screen.MousePointer = vbHourglass
        grd2.Visible = False
'   grd2.Row = grd2.MouseRow
'   grd2.col = grd2.MouseCol
        
        txtInput.Visible = False
    If Me.grd2.row > 0 Then
        txtInputState = grd2.TextMatrix(grd2.row, 0)
        If txtInputState = "核稿人" Then
            txtInputMin = MinE
            txtInputMax = MaxE
        Else
            txtInputMin = MinF
            txtInputMax = MaxF
        End If
        SetBox
    End If
    grd2.Visible = True
Screen.MousePointer = vbDefault
End Sub

Private Sub grd2_KeyDown(KeyCode As Integer, Shift As Integer)
'modify by sonia 2016/3/3 改72006為73022
If strUserNum = "73022" Then Exit Sub

   Dim iNextRow As Integer, iNextCol As Integer
   If KeyCode = 13 Or (Shift = 0 And KeyCode >= 37 And KeyCode <= 40) Then
      With grd2
         iNextRow = .row
         iNextCol = .col
         Select Case KeyCode
            Case 13
               SetBox
            Case 38 '上
               iNextRow = .row - 1
            Case 40 '下
               iNextRow = .row + 1
            Case 37 '左
               iNextCol = .col - 1
            Case 39 '右
               iNextCol = .col + 1
         End Select
         If iNextRow > 1 And iNextRow < .Rows And iNextCol > 0 And iNextCol < .Cols - 1 Then
'            .Row = iNextRow:
            .col = iNextCol
         End If
      End With
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
txt1(Index).Tag = txt1(Index) 'Add by Morgan 2008/10/24
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
'add by nickc 2005/04/26
If Index = 5 Then Exit Sub
Select Case KeyAscii
Case 46, 48 To 57, 8, 43, 45
Case Else
         KeyAscii = 0
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 0
            If Val(txt1(Index)) < 0 Or Val(txt1(Index)) > 999 Then
               MsgBox "修改及複雜案件的時數 請介於 0 到 999 ！", , "錯誤！"
               Cancel = True
            Else
               lbl2(6).Caption = txt1(Index).Text
               ReCal
            End If
   Case 1
            If Val(txt1(Index)) < MinD Or Val(txt1(Index)) > MaxD Then
               MsgBox "自我評分請介於 " & Trim(MinD) & " 到 " & Trim(MaxD) & " ！", , "錯誤！"
               Cancel = True
            Else
               lbl2(19).Caption = txt1(Index).Text
               ReCal
            End If
   Case 2
            'Modify by Morgan 2008/10/28 可刪除本人所打的分數(輸錯人...)
            If Trim(txt1(Index)) = "" And txt1(Index).Tag <> "" Then
               If MsgBox("是否確定要刪除核稿人評分資料！", vbYesNo + vbDefaultButton2) = vbNo Then
                  Cancel = True
               End If
            ElseIf Val(txt1(Index)) < MinE Or Val(txt1(Index)) > MaxE Then
               MsgBox "核稿人評分請介於 " & Trim(MinE) & " 到 " & Trim(MaxE) & " ！", , "錯誤！"
               Cancel = True
            End If
            
            If Cancel = False Then
               IntoGrid "1", txt1(Index), True
               ReCal
               txt1(Index).Tag = txt1(Index)
            End If
   Case 3
            'Modify by Morgan 2008/10/28 可刪除本人所打的分數(輸錯人...)
            If Trim(txt1(Index)) = "" And txt1(Index).Tag <> "" Then
               If MsgBox("是否確定要刪除帶人主管評分資料！", vbYesNo + vbDefaultButton2) = vbNo Then
                  Cancel = True
               Else
                  txt1(5) = ""
               End If
            ElseIf Val(txt1(Index)) < MinF Or Val(txt1(Index)) > MaxF Then
               MsgBox "帶人主管評分請介於 " & Trim(MinF) & " 到 " & Trim(MaxF) & " ！", , "錯誤！"
               Cancel = True
            End If
            If Cancel = False Then
               IntoGrid "2", txt1(Index), True
               ReCal
               txt1(Index).Tag = txt1(Index)
            End If
   Case 4
            If Val(txt1(Index)) < MinG Or Val(txt1(Index)) > MaxG Then
               MsgBox "部門主管評分請介於 " & Trim(MinG) & " 到 " & Trim(MaxG) & " ！", , "錯誤！"
               Cancel = True
            Else
               lbl2(22).Caption = txt1(Index).Text
               ReCal
            End If
   'add by nickc 2005/04/26
   Case 5
            IntoGrid "2", txt1(3)
   Case Else
   End Select
If Cancel = True Then
   txt1(Index).SetFocus
   txt1_GotFocus Index
Else
   ReCal
End If
End Sub

'計算季考核得分 全部重算，因為有可能會修改  修改及複雜案件時數
Sub ReCal()
Dim tmplbl20 As String    '核稿人
Dim tmplbl21 As String    '帶人主管


      If ProSysState = "1" Then '承辦人
      
         If Val(txt1(0).Text) <> 0 Then
            lbl2(7).Caption = Format(Val(txt1(0).Text) / Val(4), "0.00")
         Else
            lbl2(7).Caption = "0.00"
         End If
         'edit by nickc 2005/03/25 加入銷案件數
         'lbl2(8).Caption = Trim(Val(lbl2(5)) + Val(lbl2(7)))
         lbl2(8).Caption = Trim(Val(lbl2(5)) + Val(lbl2(7)) + Val(lbl2(24)))
         '達成率
         If Val(lbl2(2).Caption) <> 0 Then
            lbl2(9).Caption = Format(Trim((Val(lbl2(8).Caption) / Val(lbl2(2).Caption)) * 100), "0.00")
         Else
            lbl2(9).Caption = "0"
         End If
         '考核件數得分
         'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式)
         'If Val(lbl2(9).Caption) > 100 Then
         If Val(lbl2(9).Caption) > 100 Or Val(BeginDayCP) >= Val(PUB_108RuleDate) Then
         'end 2019/3/20
            lbl2(10).Caption = Format(Val(lbl2(9).Caption) / 100 * 0.8 * m_AR02, "0.00")
         Else
            lbl2(10).Caption = Format(((Val(lbl2(9).Caption) / 100) ^ 2) * 0.8 * m_AR02, "0.00")
         End If
         tmplbl20 = ReCalSum("1")
         tmplbl21 = ReCalSum("2")
         lbl2(20).Caption = tmplbl20
         lbl2(21).Caption = tmplbl21
         'add by nickc 2007/11/26
         If grd2.Visible = True Then
            txt1(2) = tmplbl20
            txt1(3) = tmplbl21
         End If
         lbl2(23).Caption = Format((((Val(lbl2(10)) + Val(lbl2(16)) + Val(lbl2(18)) + Val(lbl2(19))) * (1 + (Val(tmplbl20) / 100)) + Val(tmplbl21)) * (1 + (Val(lbl2(22)) / 100))), "0.00")
      Else   '繪圖
         If Val(txt1(0).Text) <> 0 Then
            lbl2(7).Caption = Format(Val(txt1(0).Text) / Val(4), "0.00")
         Else
            lbl2(7).Caption = "0.00"
         End If
         'edit by nickc 2005/03/25 加入銷案件數
         'lbl2(8).Caption = Trim(Val(lbl2(5)) + Val(lbl2(7)))
         lbl2(8).Caption = Trim(Val(lbl2(5)) + Val(lbl2(7)) + Val(lbl2(24)))
         '達成率
         If Val(lbl2(2).Caption) <> 0 Then
            lbl2(9).Caption = Format(Trim((Val(lbl2(8).Caption) / Val(lbl2(2).Caption)) * 100), "0.00")
         Else
            lbl2(9).Caption = "100.00"
         End If
         '考核件數得分
         'Modified by Morgan 2019/3/20 108考核(得分取消(達成率)^2的計算方式)
         'If Val(lbl2(9).Caption) > 100 Then
         If Val(lbl2(9).Caption) > 100 Or Val(BeginDayCP) >= Val(PUB_108RuleDate) Then
         'end 2019/3/20
            lbl2(10).Caption = Format(Val(lbl2(9).Caption) / 100 * 0.8 * m_AR13, "0.00")
         Else
            lbl2(10).Caption = Format(((Val(lbl2(9).Caption) / 100) ^ 2) * 0.8 * m_AR13, "0.00")
         End If
         tmplbl21 = ReCalSum("2")
         lbl2(21).Caption = tmplbl21
         'add by nickc 2007/11/26
         If grd2.Visible = True Then
            txt1(3) = tmplbl21
         End If
         lbl2(23) = Format((Val(lbl2(10)) + Val(lbl2(13)) + Val(lbl2(16)) + Val(lbl2(18)) + Val(lbl2(19)) + Val(tmplbl21)) * (1 + (Val(lbl2(22)) / 100)), "0.00")
      End If
      
End Sub

Sub GetAB(oAB01 As String, oAB02 As Integer, oAB03 As Integer, oAB04 As String)
Dim tmpRS3 As New ADODB.Recordset
Set tmpRS3 = New ADODB.Recordset
grd2.Clear
grd2.Rows = 2
grd2.Refresh
'strSQL = "select decode(ab05,'1','核稿人','2','帶人主管',''),st02,ab07,ab06 from Assessboss,staff where ab01='" & oAB01 & "' and ab02=" & oAB02 & " and ab03=" & oAB03 & " and ab04='" & oAB04 & "' and ab06=st01(+) "
strSql = "select decode(ab05,'1','核稿人','2','帶人主管',''),st02,ab07,ab06,ab08 from Assessboss,staff where ab01='" & oAB01 & "' and ab02=" & oAB02 & " and ab03=" & oAB03 & " and ab04='" & oAB04 & "' and ab06=st01(+) "
With tmpRS3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      SetDataListWidth
      Set grd2.Recordset = tmpRS3
      SetDataListWidth
End With

End Sub

'重新計算或是帶人主管分數
'oKind = 1   核稿人
'oKind = 2   帶人主管
Function ReCalSum(oKind As String) As String
Dim S_KeyWord As String
Dim S_Count As Integer
Dim S_Sum As Integer
If oKind = "1" Then
   S_KeyWord = "核稿人"
ElseIf oKind = "2" Then
   S_KeyWord = "帶人主管"
End If
S_Sum = 0
S_Count = 0
For ii = 0 To grd2.Rows - 1
   grd2.row = ii
   grd2.col = 0
   If grd2.Text = S_KeyWord And Trim(grd2.TextMatrix(ii, 2)) <> "" Then
      grd2.col = 2
      S_Count = S_Count + 1
      S_Sum = S_Sum + Val(grd2.Text)
   End If
Next ii
If S_Count <> 0 Then
   ReCalSum = Format(S_Sum / S_Count, "0.00")
Else
   ReCalSum = ""
End If
End Function

Sub ThisUserData()
If ProSysState = "1" Then '承辦人
   txt1(2).Text = ""
   lbl2(20).Caption = ""
   txt1(3).Text = ""
   lbl2(21).Caption = ""
   For ii = 0 To grd2.Rows - 1
      grd2.row = ii
      grd2.col = 0
      If grd2.Text = "核稿人" Then
         grd2.col = 3
         If grd2.Text = strUserNum Then
            grd2.col = 2
            txt1(2).Text = grd2.Text
            lbl2(20).Caption = grd2.Text
            Exit For
         End If
      End If
   Next ii
   For ii = 0 To grd2.Rows - 1
      grd2.row = ii
      grd2.col = 0
      If grd2.Text = "帶人主管" Then
         grd2.col = 3
         If grd2.Text = strUserNum Then
            grd2.col = 2
            txt1(3).Text = grd2.Text
            lbl2(21).Caption = grd2.Text
            'add by nickc 2005/04/26
            grd2.col = 4
            txt1(5).Text = grd2.Text
            Exit For
         End If
      End If
   Next ii
Else
   txt1(3).Text = ""
   lbl2(21).Caption = ""
   For ii = 0 To grd2.Rows - 1
      grd2.row = ii
      grd2.col = 0
      If grd2.Text = "帶人主管" Then
         grd2.col = 3
         If grd2.Text = strUserNum Then
            grd2.col = 2
            txt1(3).Text = grd2.Text
            lbl2(21).Caption = grd2.Text
            'add by nickc 2005/04/26
            grd2.col = 4
            txt1(5).Text = grd2.Text
            Exit For
         End If
      End If
   Next ii
End If
End Sub

Private Sub txtInputValidate(Cancel As Boolean)
If Val(txtInput.Text) > txtInputMax Or Val(txtInput.Text) < txtInputMin Then
      MsgBox txtInputState & "評分請介於 " & Trim(txtInputMin) & " 到 " & Trim(txtInputMax) & " ！", , "錯誤！"
      txtInput.SetFocus
      txtInput.SelStart = 0
      txtInput.SelLength = Len(txtInput)
      Cancel = True
End If
End Sub

Sub IntoGrid(oKind As String, oValue As String, Optional ByVal bDelete As Boolean)

'Modify by Morgan 2008/10/24 改刪除時可空白
If oValue = "" Then
   If bDelete = False Then
      Exit Sub
   End If
End If

Dim S_KeyWord As String
Dim S_Ok As Boolean
Dim S_Sum As Integer
If oKind = "1" Then
   S_KeyWord = "核稿人"
ElseIf oKind = "2" Then
   S_KeyWord = "帶人主管"
End If
S_Ok = False
For ii = 0 To grd2.Rows - 1
   grd2.row = ii
   grd2.col = 0
   If grd2.Text = S_KeyWord Then
      grd2.col = 3
      If grd2.Text = strUserNum Then
         S_Ok = True
         grd2.col = 2
         grd2.Text = oValue
         'add by nickc 2005/04/26
         If S_KeyWord = "帶人主管" Then
            grd2.col = 4
            grd2.Text = Me.txt1(5).Text
         End If
      End If
   End If
Next ii
If S_Ok = False Then
   grd2.row = grd2.Rows - 1
   grd2.col = 0
   If grd2.Text <> "" Then
      grd2.Rows = grd2.Rows + 1
   End If
   grd2.row = grd2.Rows - 1
   grd2.col = 0
   grd2.Text = S_KeyWord
   grd2.col = 1
   grd2.Text = strUserName
   grd2.col = 2
   grd2.Text = oValue
   grd2.col = 3
   grd2.Text = strUserNum
   'add by nickc 2005/04/26
   If S_KeyWord = "帶人主管" Then
      grd2.col = 4
      grd2.Text = Me.txt1(5).Text
   End If
End If
grd2.Refresh
End Sub

Function ChgNull(oStr As String) As String
If Len(Trim(oStr)) = 0 Then
   ChgNull = "null"
Else
   ChgNull = oStr
End If
End Function

Public Sub CheckOC33()           'NICK
If AdoRecordSet33.State <> 0 Then
   AdoRecordSet33.Close
End If
End Sub

