VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_d 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護"
   ClientHeight    =   5810
   ClientLeft      =   3370
   ClientTop       =   2940
   ClientWidth     =   10020
   ControlBox      =   0   'False
   FillColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5810
   ScaleWidth      =   10020
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   2070
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   60
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   2
      Left            =   8220
      TabIndex        =   16
      Top             =   50
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Height          =   375
      Index           =   1
      Left            =   7500
      TabIndex        =   15
      Top             =   50
      Width           =   675
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5265
      Left            =   60
      TabIndex        =   19
      Top             =   480
      Width           =   9930
      _ExtentX        =   17498
      _ExtentY        =   9296
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "瀏覽資料"
      TabPicture(0)   =   "frm090201_d.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Combo3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grd1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdok2(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdok2(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Combo1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "詳細資料"
      TabPicture(1)   =   "frm090201_d.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(7)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(31)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(29)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(6)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(27)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(24)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(23)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(32)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(30)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(11)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(13)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label1(14)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(15)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label1(17)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label1(18)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label1(19)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(20)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label1(21)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label1(8)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label1(1)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label1(12)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "lblClose"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label1(46)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label1(39)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label1(26)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label1(47)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label1(4)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label1(5)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "lblEApp"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Label18"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "lbl1(1)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "lbl1(0)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "lbl1(3)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "lbl1(5)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "lbl1(7)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "lbl1(9)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "lbl1(11)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "lbl1(15)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "lbl1(17)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "lbl1(19)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "lbl1(21)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "lbl1(23)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "lbl1(8)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "lbl1(10)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "lbl1(6)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "lbl1(28)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "txtCP64"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "txtEP12"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "Combo2"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "Combo6"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "txt1(8)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "txt1(4)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "txt1(1)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "txt1(18)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "txt1(3)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "cmdOK(4)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "cmd(5)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "txt1(7)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "cmd(7)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "txt1(2)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "txt1(12)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).ControlCount=   62
      TabCaption(2)   =   "待辦歷程"
      TabPicture(2)   =   "frm090201_d.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(48)"
      Tab(2).Control(1)=   "Label16"
      Tab(2).Control(2)=   "grd2"
      Tab(2).Control(3)=   "Combo5"
      Tab(2).Control(4)=   "cmdDetail"
      Tab(2).Control(5)=   "cmdQuery"
      Tab(2).ControlCount=   6
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   12
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   0
         Top             =   465
         Width           =   930
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   1
         Top             =   735
         Width           =   930
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "未完稿暫存區"
         Height          =   285
         Index           =   7
         Left            =   7350
         Style           =   1  '圖片外觀
         TabIndex        =   13
         Top             =   1050
         Width           =   1260
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   7
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   8
         Top             =   3090
         Width           =   915
      End
      Begin VB.CommandButton cmd 
         Caption         =   "承辦歷程(&E)"
         Height          =   285
         Index           =   5
         Left            =   7350
         TabIndex        =   12
         Top             =   540
         Width           =   1290
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "畫面更新(&Q)"
         Height          =   360
         Left            =   -66960
         TabIndex        =   54
         Top             =   480
         Width           =   1125
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "明細資料(&D)"
         Height          =   360
         Left            =   -68130
         TabIndex        =   53
         Top             =   480
         Width           =   1125
      End
      Begin VB.ComboBox Combo5 
         Height          =   260
         ItemData        =   "frm090201_d.frx":0054
         Left            =   -69150
         List            =   "frm090201_d.frx":0064
         Style           =   2  '單純下拉式
         TabIndex        =   52
         Top             =   510
         Width           =   960
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "完整卷宗"
         Height          =   285
         Index           =   4
         Left            =   7350
         Style           =   1  '圖片外觀
         TabIndex        =   14
         Top             =   1920
         Width           =   870
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1350
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   18
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   5
         Top             =   1920
         Width           =   915
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         ItemData        =   "frm090201_d.frx":0083
         Left            =   -70260
         List            =   "frm090201_d.frx":0096
         TabIndex        =   48
         Top             =   390
         Width           =   2430
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   4410
         Left            =   -74925
         TabIndex        =   20
         Top             =   750
         Width           =   9795
         _ExtentX        =   17268
         _ExtentY        =   7796
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   2
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
         _Band(0).Cols   =   1
      End
      Begin VB.CommandButton cmdok2 
         Caption         =   "未發文"
         Height          =   400
         Index           =   1
         Left            =   -66648
         TabIndex        =   18
         Top             =   348
         Width           =   852
      End
      Begin VB.CommandButton cmdok2 
         Caption         =   "當月資料"
         Height          =   400
         Index           =   0
         Left            =   -67656
         TabIndex        =   17
         Top             =   348
         Width           =   972
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   5010
         MaxLength       =   1
         TabIndex        =   4
         Top             =   1635
         Width           =   480
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   4
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   6
         Top             =   2205
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   8
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   10
         Top             =   3390
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
         Height          =   4335
         Left            =   -74940
         TabIndex        =   55
         Top             =   840
         Width           =   9825
         _ExtentX        =   17339
         _ExtentY        =   7638
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|目次|流程日期|本所案號|案件名稱|國家|案件性質|本所期限|承辦人|承辦期限|智權人員|目前流程狀態|不顯示"
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
         _Band(0).Cols   =   14
      End
      Begin MSForms.ComboBox Combo6 
         Height          =   300
         Left            =   7350
         TabIndex        =   9
         Top             =   3090
         Width           =   2220
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3916;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Left            =   5010
         TabIndex        =   2
         Top             =   1020
         Width           =   2220
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3916;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   300
         Left            =   -74160
         TabIndex        =   79
         Top             =   390
         Width           =   2730
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "4815;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEP12 
         Height          =   600
         Left            =   5010
         TabIndex        =   7
         Top             =   2490
         Width           =   4845
         VariousPropertyBits=   -1466939365
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "8546;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCP64 
         Height          =   960
         Left            =   5010
         TabIndex        =   11
         Top             =   4050
         Width           =   4845
         VariousPropertyBits=   -1466939367
         MaxLength       =   2000
         ScrollBars      =   3
         Size            =   "8546;1693"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   28
         Left            =   5010
         TabIndex        =   78
         Top             =   3720
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   6
         Left            =   5970
         TabIndex        =   77
         Top             =   1710
         Visible         =   0   'False
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   10
         Left            =   5970
         TabIndex        =   76
         Top             =   1380
         Visible         =   0   'False
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   8
         Left            =   5970
         TabIndex        =   75
         Top             =   780
         Visible         =   0   'False
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   23
         Left            =   1020
         TabIndex        =   74
         Top             =   4020
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   21
         Left            =   1020
         TabIndex        =   73
         Top             =   3720
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   19
         Left            =   1020
         TabIndex        =   72
         Top             =   3360
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   17
         Left            =   1020
         TabIndex        =   71
         Top             =   3060
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   15
         Left            =   1020
         TabIndex        =   70
         Top             =   2760
         Width           =   2775
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "4895;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   11
         Left            =   1470
         TabIndex        =   69
         Top             =   2430
         Width           =   615
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1085;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   555
         Index           =   9
         Left            =   1020
         TabIndex        =   68
         Top             =   1740
         Width           =   2715
         Caption         =   "LblFM2"
         Size            =   "4789;979"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   7
         Left            =   1020
         TabIndex        =   67
         Top             =   1440
         Width           =   1695
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2990;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   5
         Left            =   1020
         TabIndex        =   66
         Top             =   1140
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   3
         Left            =   1020
         TabIndex        =   65
         Top             =   840
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   64
         Top             =   540
         Width           =   615
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1085;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   63
         Top             =   540
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label18 
         Caption         =   "可不跑承辦歷程"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   7350
         TabIndex        =   62
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblEApp 
         Caption         =   "電子送件"
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
         Height          =   180
         Left            =   3030
         TabIndex        =   61
         Top             =   1050
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "核稿人："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   5
         Left            =   4260
         TabIndex        =   60
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "會稿完成日："
         Height          =   180
         Index           =   4
         Left            =   3900
         TabIndex        =   59
         Top             =   3135
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "判發人："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   47
         Left            =   6615
         TabIndex        =   58
         Top             =   3165
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "註：雙擊選取時，開啟承辦歷程。"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   -74940
         TabIndex        =   57
         Top             =   330
         Width           =   3315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "最近聯絡："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   48
         Left            =   -70050
         TabIndex        =   56
         Top             =   570
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "完稿日："
         Height          =   180
         Index           =   26
         Left            =   4260
         TabIndex        =   51
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "指定會稿日："
         Height          =   180
         Index           =   39
         Left            =   3915
         TabIndex        =   50
         Top             =   1950
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "進度備註："
         Height          =   180
         Index           =   46
         Left            =   4095
         TabIndex        =   49
         Top             =   4080
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "顏色說明："
         Height          =   225
         Left            =   -71160
         TabIndex        =   47
         Top             =   432
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "承辦人： "
         Height          =   180
         Index           =   0
         Left            =   -74904
         TabIndex        =   46
         Top             =   432
         Width           =   792
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
         Height          =   180
         Left            =   3030
         TabIndex        =   45
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   195
         Index           =   12
         Left            =   75
         TabIndex        =   44
         Top             =   3720
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "目次："
         Height          =   180
         Index           =   1
         Left            =   75
         TabIndex        =   43
         Top             =   570
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "承辦人："
         Height          =   180
         Index           =   8
         Left            =   1275
         TabIndex        =   41
         Top             =   570
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "總收文號："
         Height          =   180
         Index           =   21
         Left            =   75
         TabIndex        =   40
         Top             =   855
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "收文日："
         Height          =   180
         Index           =   20
         Left            =   75
         TabIndex        =   39
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   180
         Index           =   19
         Left            =   75
         TabIndex        =   38
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱："
         Height          =   180
         Index           =   18
         Left            =   75
         TabIndex        =   37
         Top             =   1755
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "是否算案件數："
         Height          =   195
         Index           =   17
         Left            =   75
         TabIndex        =   36
         Top             =   2445
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   195
         Index           =   15
         Left            =   75
         TabIndex        =   35
         Top             =   2775
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限："
         Height          =   195
         Index           =   14
         Left            =   75
         TabIndex        =   34
         Top             =   3075
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限："
         Height          =   195
         Index           =   13
         Left            =   75
         TabIndex        =   33
         Top             =   3390
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "點數："
         Height          =   195
         Index           =   11
         Left            =   75
         TabIndex        =   32
         Top             =   4035
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y/N)"
         Height          =   180
         Index           =   30
         Left            =   5520
         TabIndex        =   29
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "(N：不算)"
         Height          =   195
         Index           =   32
         Left            =   2205
         TabIndex        =   28
         Top             =   2445
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "齊備日："
         Height          =   180
         Index           =   23
         Left            =   4260
         TabIndex        =   27
         Top             =   795
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "會稿日："
         Height          =   180
         Index           =   24
         Left            =   4260
         TabIndex        =   26
         Top             =   2205
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "是否會稿："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   27
         Left            =   4035
         TabIndex        =   25
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "發文日："
         Height          =   180
         Index           =   6
         Left            =   4260
         TabIndex        =   24
         Top             =   3420
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "取消收文日："
         Height          =   180
         Index           =   29
         Left            =   3615
         TabIndex        =   23
         Top             =   3750
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "承辦備註："
         Height          =   180
         Index           =   31
         Left            =   4095
         TabIndex        =   22
         Top             =   2655
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "請點選""確定""按鈕存檔!!"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5970
         TabIndex        =   21
         Top             =   30
         Width           =   3225
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "承辦期限："
         Height          =   180
         Index           =   7
         Left            =   4035
         TabIndex        =   42
         Top             =   525
         Width           =   960
      End
   End
   Begin VB.Label Label4 
      Caption         =   "申請國家："
      Enabled         =   0   'False
      Height          =   180
      Left            =   2715
      TabIndex        =   30
      Top             =   495
      Width           =   900
   End
End
Attribute VB_Name = "frm090201_d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/9/27 Form2.0已修改
'Modify By Sindy 2021/9/15 參考 frm090201_b
Option Explicit

Public TextOk As Boolean
Public Combo1_String As String
Dim s As Integer, i As Integer, k As Integer
Dim SWPRow As String, SWPRow2 As String, SWPColor As String, SWPColor2 As String
Dim strTemp(0 To 26) As String
Dim Tmp001 As String, Tmp002 As String, Tmp003 As String, Tmp004 As String
Dim SeekTmpBk As String
Dim ChkNoData As Boolean, ChkData As Boolean
Dim Fobj As FileSystemObject
Dim StrGrp090201 As String
Dim Adorecordset99 As New ADODB.Recordset
Dim m_SqlGrpStr1 As String, m_SqlGrpStr2 As String, m_SqlGrpStr3 As String, m_SqlGrpStr4 As String, m_SqlGrpStr5 As String
Dim m_strCP09 As String '總收文號
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String '本所案號
Dim m_CP10 As String
Dim m_CP13 As String
Dim m_CP14 As String
Dim m_CP31 As String
Dim m_CP43 As String
Dim m_CP44 As String
Dim m_CP112 As String
Dim m_CP159 As String 'Add By Sindy 2020/1/31
Dim strSQL1 As String
Dim strSQL2 As String
Dim StrSQL6 As String
Dim StrSQL61 As String
Dim StrSQL62 As String
Dim StrSQL63 As String
Dim StrSQL64 As String
Dim StrSTM As String
Dim StrSLC As String
Dim StrSHC As String
Dim StrSSP As String
Dim m_Country As String
Dim m_CaseName As String
Dim m_SaleArea As String
Dim m_CuNo As String
Dim m_FieldList() As FIELDITEM
'紀錄 mail 資料，在 trans 後發
Dim skMail() As SeekMails
Dim m_NA03 As String
Dim bolInsert As Boolean, bolUpdate As Boolean, bolDelete As Boolean, bolSelect As Boolean, bolPrint As Boolean
Dim m_CPM05 As String
Dim m_blnClkSure As Boolean '判斷是否按下確定
Dim m_CP149 As String 'Add By Sindy 2012/10/24
Public cmdState As Integer '紀錄作用按鍵
Dim m_CP140 As String
'Added by Lydia 2015/11/12 新增查名單對應
Public Tmpfrm090130 As Form
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限
'Add By Sindy 2018/4/17
Public intBackTab As Integer
Dim dblPrevRow As Double
Dim m_intRow As Integer, m_intCol As Integer
'2018/4/17 END
'Add By Sindy 2018/4/20
Dim m_PP04 As String '預設核稿人
Dim m_PP05 As String '預設判發人
Dim m_PP01 As String '系統別
Dim m_PP03 As String '案件性質或核判分類
Dim m_EP39 As String '核稿完成日
Dim m_CPM28 As String
Dim m_CPM29 As String
Dim m_CP27 As String '發文日
Public m_chkcmdok1 As Boolean '記錄確定鍵是否存檔成功
Public m_Flow As String '欲新增的下一流程
'2018/4/20 END
Dim m_AttachPath As String 'Add By Sindy 2020/1/31
Dim m_CP16 As String 'Add By Sindy 2020/11/17
Dim m_CP163 As String 'Add By Sindy 2020/12/2


Private Sub cmd_Click(Index As Integer)
Dim strTempName As String
Dim nFrm As Form 'Add By Sindy 2018/4/17
Dim ET01 As String, ET03 As String 'Add By Sindy 2020/10/20

On Error GoTo ErrHand

Select Case Index
Case 5 '承辦歷程
      
      '重新檢查欄位有效性
      If TxtValidate = True Then
         
         If SetColTag(False) = False Then
            cmdOK(1).Enabled = False
            Call cmdOK_Click(1)
            cmdOK(1).Enabled = True
            If m_chkcmdok1 = False Then Exit Sub
         Else
            Call Process(lbl1(3)) '要重新查詢資料 Add By Sindy 2018/10/4
         End If
         
'         '檢查表單是否已開啟，若是，則關閉
'         For Each nFrm In Forms
'            If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'               Unload frm090202_2
'               Exit For
'            End If
'         Next
         If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
         intBackTab = 1
         frm090202_2.Hide
         frm090202_2.m_EEP01 = lbl1(3) '總收文號
         frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) '案件流程所屬人員
         frm090202_2.intReceiveKind = 0 '0.承辦人工作進度
         frm090202_2.SetParent Me
         If frm090202_2.QueryData = True Then
            frm090202_2.Show
            Me.Hide
         End If
      End If
'2018/4/17 END

'Add By Sindy 2020/3/17
Case 7 '原始檔暫存區
   'Call PUB_ChkFormIsClose("frm100101_M")
   frm100101_M.m_strKey = lbl1(3).Caption '總收文號
   frm100101_M.SetParent Me
   If frm100101_M.QueryData = True Then
      frm100101_M.Show
      Me.Hide
   End If
'2020/3/17 END
Case Else
End Select
Exit Sub

ErrHand:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub SetComboData(mCombo As Object)
Dim blnMatch As Boolean
Dim ii As Integer
Dim strST03Con As String
   
   'Modify By Sindy 2023/11/7
'   If Left(Pub_StrUserSt03, 1) = "W" Then
      strST03Con = " and substr(ST03,1,1)='W' and ST03<>'W00'"
'   Else
'      strST03Con = " (ST03>='P20' and ST03<='P21') AND ST20<='52'"
'      'Modify By Sindy 2023/7/10 Mark
'      'strCon = " or ST01='76012'"
'   End If
   
   '代主任以上的人員
   'strExc(0) = "SELECT st01||' ==> '||st02 FROM STAFF WHERE (" & strST03Con & " AND ST04='1' AND ST20<='52')" & strCon & " ORDER BY ST01"
   strExc(0) = "SELECT st01||' ==> '||st02 FROM STAFF WHERE st04='1' and st01>'63' and st01<'F'" & strST03Con & " ORDER BY ST01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      .MoveFirst
      Do While .EOF = False
         blnMatch = False
         For ii = 0 To mCombo.ListCount - 1
             blnMatch = False
             If Trim(Left(mCombo.List(ii), 6)) = Left(.Fields(0), 5) Then
                 mCombo.ListIndex = ii
                 blnMatch = True
                 Exit For
             End If
         Next ii
         If blnMatch = False Then
            intI = mCombo.ListCount
            mCombo.AddItem "" & .Fields(0), intI
         End If
         .MoveNext
      Loop
      End With
   End If
   blnMatch = False
   For ii = 0 To mCombo.ListCount - 1
      If Trim(Left(mCombo.List(ii), 6)) = mCombo.Tag Then
          mCombo.ListIndex = ii
          blnMatch = True
          Exit For
      End If
   Next ii
   If blnMatch = False Then mCombo.ListIndex = 0
   mCombo.Tag = mCombo.Text
End Sub

Sub Process(strText As String)
Dim stVTB As String
Dim oLbl As Object
Dim oTxt1 As Object
'Add By Sindy 2018/4/20
Dim strRefEEP02 As String
'2018/4/20 END
Dim tmpBol As Boolean 'Added by Lydia 2019/05/02
Dim rsTmp As New ADODB.Recordset
   
   Me.Enabled = False
   
   '法務
   stVTB = " SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112" & _
            ",NVL(LC05,NVL(LC06,LC07)) C10,'' C14,'' C26,'' C28,LC08,cp49 as C33,LC15 as m_country,LC11 as cuno,CP43,CP140,CP118,CP143,cp159,cp16,cp163" & _
            " FROM CASEPROGRESS,LawCase WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr3 & ") AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04"
   '顧問
   stVTB = stVTB + " UNION all SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112" & _
            ",HC06 C10,'' C14,'' C26,'' C28,HC09,'*' C33,'000' as m_country,HC05 as cuno,CP43,CP140,CP118,CP143,cp159,cp16,cp163" & _
            " FROM CASEPROGRESS,HireCase WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr4 & ") AND HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04"
   
   'Add By Sindy 2022/11/18
   '法務
   stVTB = stVTB + " UNION all SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112" & _
            ",SP05 C10,'' C14,'' C26,'' C28,SP15,'*' C33,'000' as m_country,SP08 as cuno,CP43,CP140,CP118,CP143,cp159,cp16,cp163" & _
            " FROM CASEPROGRESS,ServicePractice WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr5 & ") AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04"
   '2022/11/18 END
   
   strSql = "SELECT EP01,S1.ST02 C2,sqldateT(CP48) C3,CP09,EP13,sqldateT(cp05) C6,EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04 C8" & _
      ",EP06,C10,EP09,CP26,EP07,C14,EP04,decode(na01,'000',cpm03,cpm04) C16,EP03,sqldateT(CP06) C18,EP08,sqldateT(CP07) C20,CP27" & _
      ",S5.ST02 C22,EP11,CP18,EP12,C26,Nvl(EP35,0) C27,C28,sqldateT(CP57) C29,CP10,CP15,'' PA57,C33,EP27,EP31,cp13,ep05,m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99" & _
      ",cp106,cuno,cp111,cp112,ep28,ep32,ep33,na03,cp64,cpm05,cp44,ibf01,S3.ST02 EP04N,pp04,s6.st02 pp04N,s2.st02 EP13N,s4.st02 EP03N" & _
      ",NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CuName,CP43,CP140,CP118,CP143,EP38,EP39,pp05,EP40,cpm28,cpm29,cpm23,cp159,cp16,cp163" & _
      " from (" & stVTB & ") X,ENGINEERPROGRESS,CASEPROPERTYMAP,nation" & _
      ",STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,customer,imgbytefile,promoterproofreader,staff S6" & _
      " where EP02(+)=CP09 and cpm01(+)=CP01 and cpm02(+)=CP10 AND na01(+)=m_country" & _
      " AND S1.ST01(+)=EP05 AND S2.ST01(+)=EP13 AND S3.ST01(+)=EP04 AND S4.ST01(+)=EP03 AND S5.ST01(+)=CP13" & _
      " and cu01(+)=substr(cuno,1,8) and cu02(+)=substr(cuno,9) and pp01(+)='T' and pp02(+)=cp14 and pp03(+)=cp10 and s6.st01(+)=pp04" & _
      " and ibf01(+)=cp01 and ibf02(+)=cp02 and ibf03(+)=cp03 and ibf04(+)=cp04 and ibf05(+)='1'"
   CheckOC
   With adoRecordset
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   
   '***** 清除欄位值 *****
   m_CP163 = ""
   m_CPM28 = ""
   m_CPM29 = ""
   m_EP39 = "" '核稿完成日
   m_CP01 = "": m_CP02 = "": m_CP03 = "": m_CP04 = ""
   m_CP159 = ""
   For Each oLbl In lbl1
      oLbl.Caption = ""
   Next
   Me.lblClose.Caption = ""
   For Each oTxt1 In txt1
      oTxt1.Text = ""
   Next
   txtCP64.Text = ""
   Combo2.Clear: Combo2.Tag = ""
   Combo6.Clear: Combo6.Tag = ""
   '***** END *****
   
   If .RecordCount <> 0 And .RecordCount > 0 Then
      .MoveFirst
      m_CP163 = "" & .Fields("CP163") 'Add By Sindy 2020/12/2
      m_CP43 = "" & .Fields("CP43")
      m_CP01 = SystemNumber(Trim(.Fields("C8")), 1)
      m_CP02 = SystemNumber(Trim(.Fields("C8")), 2)
      m_CP03 = SystemNumber(Trim(.Fields("C8")), 3)
      m_CP04 = SystemNumber(Trim(.Fields("C8")), 4)
      
      'Add By Sindy 2018/4/20
      m_EP39 = "" & .Fields("EP39")
      m_CPM28 = "" & .Fields("CPM28")
      m_CPM29 = "" & .Fields("CPM29")
      m_CP159 = "" & .Fields("CP159") 'Add By Sindy 2020/1/31
      '電子送件
      If Not IsNull(.Fields("CP118")) Then
         lblEApp.Visible = True
      Else
         lblEApp.Visible = False
      End If
      '2018/4/20 END
      
      For i = 0 To 29
         'Modify by Morgan 2008/10/13 原來值由lablel改為放tag或text
         '會稿日
         If i = 12 Then
            txt1(4).Tag = ChangeWStringToTString(CheckStr(.Fields(i)))
            txt1(4).Text = txt1(4).Tag
         '會稿完成日
         ElseIf i = 18 Then
            txt1(7).Tag = ChangeWStringToTString(CheckStr(.Fields(i)))
            txt1(7).Text = txt1(7).Tag
         '發文日
         ElseIf i = 20 Then
            txt1(8).Text = ChangeWStringToTString(CheckStr(.Fields(i)))
'         '是否通知客戶
'         ElseIf i = 22 Then
'            txt1(9).Text = CheckStr(.Fields(i))
         '承辦備註
         ElseIf i = 24 Then
            txtEP12.Text = CheckStr(.Fields(i))
         '承辦期限
         ElseIf i = 2 Then
            txt1(12).Text = ChangeTDateStringToTString(CheckStr(.Fields(i)))
         '案件性質
         ElseIf i = 15 Then
            If Not IsNull(.Fields("CP43")) Then '有相關總收文號
               lbl1(i) = CheckStr(.Fields(i)) & PUB_GetRelateCasePropertyName(strText, "1")
            Else
               lbl1(i) = CheckStr(.Fields(i))
            End If
         Else
            If i <> 4 And i <> 14 And i <> 16 And i <> 18 And i <> 26 And i <> 25 And i <> 27 And i <> 13 And i <> 22 And i <> 29 Then
               lbl1(i) = CheckStr(.Fields(i))
            End If
         End If
      Next i
      
      m_CuNo = "" & .Fields("CuName")
      
      m_CP13 = "" & .Fields("cp13").Value '智權人員
      
      m_CP14 = "" & .Fields("ep05").Value
      m_CP10 = "" & .Fields("cp10").Value
      m_CP16 = "" & .Fields("cp16").Value 'Add By Sindy 2020/11/17
      
      m_NA03 = "" & .Fields("NA03").Value
      
      m_CPM05 = "" & .Fields("cpm05")
      m_CP112 = "" & .Fields("cp112")
     
      m_CP44 = "" & .Fields("cp44")
        
      '進度備註
      txtCP64 = CheckStr(.Fields("cp64"))
      
      '指定會稿日
      txt1(18).Tag = ""
      txt1(18).Tag = ChangeWStringToTString(CheckStr(.Fields("ep28")))
      txt1(18) = ChangeWStringToTString(CheckStr(.Fields("ep28")))
      
      m_Country = "" & .Fields("m_country").Value
      m_CP31 = "" & .Fields("cp31").Value
      '案件名稱
      m_CaseName = "" & .Fields(9).Value
      m_SaleArea = "" & .Fields("Area").Value
      
      '收文號
      m_strCP09 = Me.lbl1(3).Caption
      If Len(Trim(CheckStr(.Fields(20)))) <> 0 Then
         m_CP27 = .Fields(20) 'Add By Sindy 2018/4/25 發文日
      Else
         m_CP27 = "" 'Add By Sindy 2018/4/25 發文日
      End If
      If IsNull(.Fields(31).Value) <> 0 Then
          Me.lblClose.Caption = ""
      Else
          Me.lblClose.Caption = "已閉卷"
      End If
      
      'Add By Sindy 2018/4/20
      m_PP04 = "" '核判表設定的核稿人
      m_PP05 = "" '核判表設定的判發人
      Call PUB_ChkIsSetPromoterReader(m_CP14, m_CP01, m_CP10, m_PP04, m_PP05, m_strCP09, m_Country, m_PP01, m_PP03)
      If m_PP04 = Trim(Left("" & Combo1.Text, 6)) Then m_PP04 = "" '為自行核稿,不需再將自己ID放入核稿人欄位
      If m_PP05 = Trim(Left("" & Combo1.Text, 6)) Then m_PP05 = "" '為自行判發,不需再將自己ID放入判發人欄位
      '核稿人:
      Combo2.AddItem "", 0
      '有完稿日時,則不用再預設核稿人
      If Val("" & .Fields("EP09")) <= 0 And Len("" & .Fields("EP04")) = 0 Then
         Combo2.Tag = m_PP04
      Else
         'Add By Sindy 2018/4/25
         If m_CP27 = "" And "" & .Fields("EP04") = "" And m_PP04 <> "" Then
            Combo2.Tag = m_PP04
         Else
         '2018/4/25 END
            Combo2.Tag = "" & .Fields("EP04")
         End If
      End If
      If Combo2.Tag <> "" Then
         Combo2.AddItem Combo2.Tag & " ==> " & GetPrjSalesNM(Combo2.Tag), 1
      End If
      Call SetComboData(Combo2) 'Add By Sindy 2018/10/1
'      blnMatch = False
'      For ii = 0 To Me.Combo2.ListCount - 1
'         If Trim(Left(Me.Combo2.List(ii), 6)) = Combo2.Tag Then
'            Me.Combo2.ListIndex = ii
'            blnMatch = True
'            Exit For
'         End If
'      Next ii
'      If blnMatch = False Then Me.Combo2.ListIndex = 0
'      Combo2.Tag = Combo2.Text
      '判發人:
      Combo6.AddItem "", 0
      '不用完稿日為預設判發的基準點,改用檢查有無送判或判發歷程
      If PUB_ChkEmpFlowExists(lbl1(3), EMP_送判) = False And _
         PUB_ChkEmpFlowExists(lbl1(3), EMP_判發) = False And _
         Len("" & .Fields("EP40")) = 0 Then
         Combo6.Tag = m_PP05
      Else
         Combo6.Tag = "" & .Fields("EP40")
      End If
      If Combo6.Tag <> "" Then
         Combo6.AddItem Combo6.Tag & " ==> " & GetPrjSalesNM(Combo6.Tag), 1
      End If
      Call SetComboData(Combo6) 'Add By Sindy 2018/10/1
'      blnMatch = False
'      For ii = 0 To Me.Combo6.ListCount - 1
'         If Trim(Left(Me.Combo6.List(ii), 6)) = Combo6.Tag Then
'            Me.Combo6.ListIndex = ii
'            blnMatch = True
'            Exit For
'         End If
'      Next ii
'      If blnMatch = False Then Me.Combo6.ListIndex = 0
'      Combo6.Tag = Combo6.Text
      '2018/4/20 END
   End If
   End With
   
'   'Add By Sindy 2020/12/1
'   txtNote.Visible = False
'   If m_CP163 <> "" Then
'      If m_CP163 <> lbl1(3) Then
'         strSql = "Select CP01,CP02,CP03,CP04,Decode('" & m_Country & "','000',CPM03,CPM04) as 案件性質" & _
'                  " from caseprogress,CasePropertyMap" & _
'                  " where cp09='" & m_CP163 & "' And CP01=CPM01(+) And CP10=CPM02(+)"
'         If rsTmp.State = 1 Then rsTmp.Close
'         rsTmp.CursorLocation = adUseClient
'         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsTmp.RecordCount > 0 Then
'            txtNote.Text = "※此案屬多案歷程，請參" & rsTmp.Fields("cp01") & "-" & rsTmp.Fields("cp02") & IIf(rsTmp.Fields("cp03") & rsTmp.Fields("cp04") = "000", "", "-" & rsTmp.Fields("cp03") & "-" & rsTmp.Fields("cp04")) & _
'                           "(" & rsTmp.Fields("案件性質") & ")"
'            txtNote.Width = 5000
'            txtNote.Visible = True
'         End If
'         If rsTmp.State = 1 Then rsTmp.Close
'      End If
'   End If
'   '2020/12/1 END
   
   CheckOC
   InitialField
   'Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & lbl1(3).Caption & "' "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_CP149 = "" & rsTmp.Fields("CP149")
      UpdateFieldOldData rsTmp
   End If
   If rsTmp.State = 1 Then rsTmp.Close
   '是否會稿
   txt1(1).Text = lbl1(6).Caption
   '齊備日
   txt1(2).Text = ChangeWStringToTString(lbl1(8).Caption)
   '完稿日
   txt1(3).Text = ChangeWStringToTString(lbl1(10).Caption)
   If PUB_GetST05(strUserNum) = "91" Or PUB_GetST05(strUserNum) = "92" Then
      '指定會稿日
      txt1(18).Enabled = True
   Else
      txt1(18).Enabled = False
   End If
   
   If ProState = "2" Then
      frm090614.TextOk = True
   End If
   
   If m_blnClkSure = False Then
      'Modified by  Lydia 2019/05/02
      'If Me.Txt1(12).Text = "" And Me.Txt1(2).Text <> "" Then Call txt1_LostFocus(2)
      If Me.txt1(12).Text = "" And Me.txt1(2).Text <> "" Then
         tmpBol = False
         Call txt1_Validate(2, tmpBol)
      End If
   End If
   
   cmd(5).Tag = "" 'Added by Lydia 2018/12/10 加註-基本權限
   'Add By Sindy 2018/4/20
   '個人案件不可用主管權限操作
   If ProState = "2" And m_CP14 = strUserNum Then  '2.主管
      cmd(5).Enabled = False
      cmdDetail.Enabled = False
      cmd(5).Tag = "N" 'Added by Lydia 2018/12/10 加註-基本權限
'   '無齊備日,不可使用歷程 (Sindy 2019/8/1:無齊備日時,進歷程也只能做附加歷程和聯絡)
'   ElseIf Val(txt1(2)) = 0 Then
'      cmd(5).Enabled = False
'      cmdDetail.Enabled = False
   Else
      cmd(5).Enabled = True
      cmdDetail.Enabled = True
      cmd(5).Tag = "Y" 'Added by Lydia 2018/12/10 加註-基本權限
   End If
   
   If m_CPM29 = "N" Then
      Label18.Visible = False '先不顯示
   Else
      Label18.Visible = False
   End If
   '若為個人工作管理及承辦人下拉選單為操作者
   '完稿日
   Me.txt1(3).Enabled = True
   '會稿日
   Me.txt1(4).Enabled = True
   '會稿完成日
   Me.txt1(7).Enabled = True
   '發文日
   If Left(m_strCP09, 1) = "C" Then Me.txt1(8).Enabled = True
   If ProState = "1" Or Trim(Left("" & Combo1.Text, 6)) = strUserNum Then
      If m_CPM29 = "" Then '要電子簽核的案件性質
         '完稿日
         Me.txt1(3).Enabled = False
         '會稿日
         Me.txt1(4).Enabled = False
         '會稿完成日
         Me.txt1(7).Enabled = False
         '發文日
         Me.txt1(8).Enabled = False
         '不自動更新會完日時,則開放可以自行輸入會稿完成日
         If Me.txt1(7).Text = "" Then
            If PUB_ChkEmpFlowExists(lbl1(3), EMP_送會, , strRefEEP02) = True Then
               If PUB_ChkEmpFlowExists(lbl1(3), EMP_不自動更新會完日, strRefEEP02) = True Then
                  Me.txt1(7).Enabled = True
               End If
            End If
         End If
      End If
   End If
   
   '是否會稿
   txt1(1).Enabled = True 'Add Sindy 2018/9/18
   If txt1(1) = "Y" And PUB_ChkEmpFlowExists(lbl1(3), EMP_送會) = True Then
      txt1(1).Enabled = False
   End If
   'Add By Sindy 2019/7/25 是否會稿,空白才要預設
   If Trim(lbl1(6).Caption) = "" Then
   '2019/7/25 END
      'Add By Sindy 2018/4/20 不會稿案件性質,則預設為N
      '商爭案是否會稿由智權人員在填寫接洽單時決定(收文,分案)
      If txt1(1).Text = "" Then
         txt1(1).Text = m_CPM28
         If m_CP10 = "101" Then '申請案(101)一定要會稿
            txt1(1).Text = "Y"
         End If
      End If
      '2018/4/20 END
      'Add By Sindy 2019/7/12 不電子簽核的案件性質,是否會稿預設為N
      If m_CPM29 = "N" Then
         txt1(1).Text = "N"
      End If
   End If
   
   Call SetColTag(True)
   Me.Enabled = True
End Sub

'Add By Sindy 2018/4/25
'bolSetTag=true : 將輸入欄位值記錄至.tag裡面
'bolSetTag=false : 比較輸入欄位值.Tag與畫面上資料是否一致
Private Function SetColTag(bolSetTag As Boolean) As Boolean
   If bolSetTag = True Then
      txt1(12).Tag = txt1(12)
      txt1(2).Tag = txt1(2)
      txt1(3).Tag = txt1(3)
      txt1(1).Tag = txt1(1)
      txt1(18).Tag = txt1(18)
      txt1(4).Tag = txt1(4)
      txtEP12.Tag = txtEP12
      txt1(7).Tag = txt1(7)
      txt1(8).Tag = txt1(8)
      txtCP64.Tag = txtCP64
      Combo6.Tag = Combo6.Text '判發人
      Combo2.Tag = Combo2.Text '核稿人
   Else
      SetColTag = True
      'Add By Sindy 2024/9/23 + Or Trim(lbl1(6).Caption) = ""
      If txt1(1) = "" Or Trim(lbl1(6).Caption) = "" Then SetColTag = False: Exit Function '是否會稿欄位空白時,確定鍵會update成Y
      If txt1(12).Tag <> txt1(12) Then SetColTag = False: Exit Function
      If txt1(2).Tag <> txt1(2) Then SetColTag = False: Exit Function
      If txt1(3).Tag <> txt1(3) Then SetColTag = False: Exit Function
      If txt1(1).Tag <> txt1(1) Then SetColTag = False: Exit Function
      If txt1(18).Tag <> txt1(18) Then SetColTag = False: Exit Function
      If txt1(4).Tag <> txt1(4) Then SetColTag = False: Exit Function
      If txtEP12.Tag <> txtEP12 Then SetColTag = False: Exit Function
      If txt1(7).Tag <> txt1(7) Then SetColTag = False: Exit Function
      If txt1(8).Tag <> txt1(8) Then SetColTag = False: Exit Function
      If txtCP64.Tag <> txtCP64 Then SetColTag = False: Exit Function
      If Left(Combo6.Tag, 5) <> Left(Combo6.Text, 5) Then SetColTag = False: Exit Function '判發人
      If Left(Combo2.Tag, 5) <> Left(Combo2.Text, 5) Then SetColTag = False: Exit Function '核稿人
   End If
End Function

'Add By Sindy 2018/4/25
Private Sub ChkEP34ToEP07EP08()
Dim bolChkEmp As Boolean
   
   'add by nickc 2006/09/26 若是輸入不會稿，直接按存檔，他不會自動代
   'If txt1(1) = "N" Then txt1(4) = txt1(3): txt1(7) = txt1(3)
   If txt1(1) = "N" Then
      bolChkEmp = False
      '要電子簽核的案件或有電子歷程的案件
      If m_CPM29 = "" Or _
         m_Flow = EMP_送核 Or _
         m_Flow = EMP_送英核 Or _
         ((PUB_ChkEmpFlowExists(lbl1(3), EMP_送核) = True Or PUB_ChkEmpFlowExists(lbl1(3), EMP_送英核) = True) And PUB_ChkEmpFlowExists(lbl1(3), EMP_核完) = False) Then
         '有核稿人
         'If txt1(5).Text <> "" And txt1(5).Text <> m_CP14 Then
         If Combo2.Text <> "" And Left(Trim(Combo2.Text), 5) <> m_CP14 Then
            bolChkEmp = True
'                     strExc(0) = "select ep39 From engineerprogress where ep02='" & lbl1(3) & "'"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        If Val("" & RsTemp.Fields("ep39")) <= 0 Then
'                           bolUpdDate = False
'                        End If
'                     End If
         End If
      End If
      If bolChkEmp = False Then '無電子簽核或無核稿主管
         'Modify By Sindy 2016/5/19 發現核完N不會稿時,系統會上(會稿日)和(會稿完成日),所以要檢查日期是否已有值,以免重新覆蓋掉 ex:P-113197
         'txt1(4) = txt1(3): txt1(7) = txt1(3)
         If Trim(txt1(4).Text) = "" Then
            txt1(4).Text = txt1(3).Text
         End If
         If Trim(txt1(7).Text) = "" Then
            txt1(7).Text = txt1(3).Text
         End If
         '2016/5/19 END
'               Else
'                  If bolUpdDate = True Then
'                     If Trim(txt1(4)) = "" Then
'                        txt1(4) = strSrvDate(2)
'                     End If
'                     If Trim(txt1(7)) = "" Then
'                        txt1(7) = strSrvDate(2)
'                     End If
'                  End If
      End If
   End If
   '2013/10/4 END
End Sub

Public Sub cmdOK_Click(Index As Integer)
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Public Sub PubShowNextData()
'Dim rsA As New ADODB.Recordset
'Dim stFileName As String
'Dim hLocalFile As Long
   
   '***2008/11/21 加註BY SONIA 按確定後很快按結束會因為DoEvents造成錯誤,因使用者未反應故暫不取消DoEvents
   Dim iMouse As Integer
   iMouse = Screen.MousePointer
   
   Select Case cmdState
   Case 1 '確定
      Select Case ProState
      Case "1", "2"
         m_chkcmdok1 = False 'Add By Sindy 2013/6/7 進入承辦歷程時會先執行一次確定鍵,因有可能已在此畫面先修改資料,且有些日期檢查條件須先執行
         If SSTab1.Tab = 0 Then Exit Sub
         
         'Add By Sindy 2019/7/12 輸入發文日,未輸入完稿日,則完稿日=發文日
         If Len(txt1(3)) = 0 And txt1(8).Enabled = True Then txt1(3) = txt1(8)
         'Add By Sindy 2018/4/25
         If txt1(1) = "" Then txt1(1) = "N"
         Call ChkEP34ToEP07EP08
         '2018/4/25 END
         
         Screen.MousePointer = vbHourglass
         'Modify By Sindy 2018/4/25
         'If SSTab1.Tab = 1 Then
         If SSTab1.Tab = 1 Or Me.m_Flow <> "" Then
         '2018/4/25 END
            If ChkNoData = False Then
               '重新檢查欄位有效性
               If TxtValidate = True Then
                  DoEvents
                  Me.Enabled = False
                  If FormSave = True Then
                     '集中發信
                     'Modify By Sindy 2018/6/20 + If m_Flow = "" Then
                     If m_Flow = "" Then BatctMail
                     
                     '更新mdb暫存資料及第一畫面的Grid內容
                     UpdEngMdb
                     TextOk = False
                     Call SetColTag(True) 'Add By Sindy 2018/4/25
                     m_chkcmdok1 = True 'Add By Sindy 2018/4/25
                  End If
                  Me.Enabled = True
                  'SSTab1.Tab = 0
'                  'Add By Sindy 2018/4/26
'                  If cmdOK(1).Enabled = True Then
'                  '2018/4/26 END
'                     SSTab1.Tab = 0
'                  End If
               'Add By Sindy 2019/12/3
               Else
                  Screen.MousePointer = vbDefault
                  Exit Sub
               '2019/12/3 END
               End If
            End If
         Else
            SSTab1.Tab = 1
         End If
         'Add By Sindy 2018/5/15
         If SSTab1.Tab = 1 Then
            Call Process(lbl1(3).Caption)
         End If
         '2018/5/15 END
         
         'Add By Sindy 2020/10/8
         If Me.m_Flow <> "" Then
            '重查資料,多案單筆歷程時,要更新瀏覽資料日期欄位值
            'Modify By Sindy 2021/2/2 嘉雯在做延展時,會點本所期限做排序再進行歷程,為不影響到畫面資料順序,改寫
            'Call Combo1_Click
            If frm090202_2.m_RetrunRecvSub <> "" Then
               Call StrMenuOneRec_RecvSub(frm090202_2.m_RetrunRecvSub)
            End If
            '2021/2/2 END
         '2020/10/8 END
            Me.m_Flow = "" 'Add By Sindy 2018/4/25
         End If
         Screen.MousePointer = iMouse
      Case Else
      End Select
         
   Case 2 '結束
      Select Case ProState
      Case "1"
          Unload Me
          Exit Sub
      Case "2"
           frm090614.Show
           Unload Me
           Exit Sub
      Case "3"
      Case Else
      End Select
      
   Case 3 '接洽單
      Screen.MousePointer = vbHourglass
      If m_CP140 <> "" Then
         '查詢接洽記錄單
         'Modify By Sindy 2022/12/23 改用共用函數
         Call PUB_Queryfrm090801(m_CP140, DBDATE(lbl1(5).Caption), Me)
'         'Modify By Sindy 2022/9/5
'         If DBDATE(Lbl1(5)) >= 接洽單電子收文啟用日 Then
'            frm090801_Q.SetParent Me
'            frm090801_Q.m_blnCallPrint = True
'            frm090801_Q.Text5 = m_CP140
'            Call frm090801_Q.cmdOK_Click(4)
'            'frm090801_Q.ZOrder
'            frm090801_Q.Show vbModal
'         Else
'         '2022/9/5 END
'            frm090801.SetParent Me
'            frm090801.m_blnCallPrint = True 'Add By Sindy 2022/10/19
'            frm090801.Text5 = m_CP140
'            frm090801.m_blnCallPrint_CRL119 = True '是否列印特殊收據頁
'            Call frm090801.cmdOK_Click(4)
'            frm090801.cmdOK(2).Visible = False
'            frm090801.cmdOK(0).Visible = False
'            frm090801.txtPCnt.Visible = False
'            Me.Hide
'         End If
         '2022/12/23 END
         cmdState = 99 '結束
'      Else
'         '檢查是否有接洽單.pdf
'         strExc(0) = "select *" & _
'                     " From casepaperpdf" & _
'                     " where cpp01='" & m_EEP01 & "' and instr(upper(cpp02),upper('" & EMP_接洽單 & ".pdf'))>0 and cpp10<>'D'"
'         rsA.CursorLocation = adUseClient
'         rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount > 0 Then
'            '讀取檔案名稱
'            stFileName = rsA.Fields("cpp02")
'   '         If GetAttachFile_CPP(m_EEP01, stFileName, m_AttachPath & "\" & stFileName) = False Then
'   '            MsgBox "無法儲存欲開啟的檔案[ " & stFileName & " ]！"
'   '         End If
'            If PUB_GetAttachFile_CPP(m_EEP01, stFileName, m_AttachPath) = True Then
'               '開啟檔案
'               ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
'            End If
'         Else
'            MsgBox "無接洽單！"
'         End If
'         rsA.Close
'         Set rsA = Nothing
      End If
      Screen.MousePointer = vbDefault
   
   Case 4 '完整卷宗
      Screen.MousePointer = vbHourglass
      frm100101_L.m_strKey = lbl1(7).Caption
      frm100101_L.SetParent Me
      If frm100101_L.QueryData = True Then
         frm100101_L.Show
         Me.Hide
      Else
         Unload frm100101_L
      End If
      Screen.MousePointer = vbDefault
   Case Else
   End Select
End Sub

Private Sub cmdok2_Click(Index As Integer)
Dim iMouse As Integer
iMouse = Screen.MousePointer

Screen.MousePointer = vbHourglass
grd1.Visible = False
Select Case Index
Case 0 '當月資料
      'Modify By Sindy 2015/9/10 增加CP140電子表單單號做排序,自動收文在前面
'      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028) or r110028=0,1,r110028)),'0.00') ,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032 FROM R090614 " & _
'                    " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' ORDER BY R110002 desc,R110003,R110004 "
      'Modify By Sindy 2023/2/8 原案件清單排列,以電子收文為優先
      '但現皆為電子收文案件,會導致當日智慧局來文不易被發現,請調整為以"目次"由大至小排列
      '取消 R110033 desc,
      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028) or r110028=0,1,r110028)),'0.00') ,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032 FROM R090614 " & _
                    " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
                    " ORDER BY R110002 desc,R110003,R110004 "
      CheckOC
      With adoRecordset
          .CursorLocation = adUseClient
          .Open strSql, adoEng, adOpenStatic, adLockReadOnly
          If .RecordCount <> 0 And .RecordCount > 0 Then
              Set grd1.Recordset = adoRecordset
              ChgGrdColor
          Else
             grd1.Clear
             grd1.Rows = 2
          End If
      End With
      CheckOC
      SWPRow2 = 1
Case 1 '未發文
      'Modify By Sindy 2015/9/10 增加CP140電子表單單號做排序,自動收文在前面
      'Modify By Sindy 2023/2/8 原案件清單排列,以電子收文為優先
      '但現皆為電子收文案件,會導致當日智慧局來文不易被發現,請調整為以"目次"由大至小排列
      '取消 R110033 desc,
      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028) or r110028=0,1,r110028)),'0.00') ,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032 FROM R090614 " & _
                  " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' AND R110018='' and (R110024='' or R110024='0')" & _
                  " order by R110002 desc "
      CheckOC
      With adoRecordset
          .CursorLocation = adUseClient
          .Open strSql, adoEng, adOpenStatic, adLockReadOnly
          If .RecordCount <> 0 And .RecordCount > 0 Then
              Set grd1.Recordset = adoRecordset
              ChgGrdColor
          Else
               grd1.Clear
               grd1.Rows = 2
               SetGrd1
          End If
      End With
      CheckOC
      SWPRow2 = 1
Case Else
End Select
'Modify By Sindy 2018/8/23
'MouseClick (1)
MouseClick_1 (1)
'2018/8/23 END
grd1.Visible = True
Screen.MousePointer = iMouse
End Sub

'Add By Sindy 2018/4/17
Private Sub cmdDetail_Click()
   Call grd2_DblClick
End Sub

'Add By Sindy 2018/4/17
Private Sub cmdQuery_Click()
   If QueryData(True) = False Then ShowNoData
End Sub

Public Function QueryData(bolFirst As Boolean) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer
Dim strQuyDate As String
Dim strVal As String
   
   m_blnColOrderAsc = True
   QueryData = True
   
   If Combo5.ListIndex = 0 Then
      strQuyDate = CompWorkDay(3, strSrvDate(1), 1) '不含當天,3個工作天
   ElseIf Combo5.ListIndex = 1 Then
      strQuyDate = CompWorkDay(5, strSrvDate(1), 1) '不含當天,5個工作天
   ElseIf Combo5.ListIndex = 2 Then
      strQuyDate = CompWorkDay(7, strSrvDate(1), 1) '不含當天,7個工作天
   Else
      '全部
   End If
   
   grd2.Clear
   SetGrd2
   
   Screen.MousePointer = vbHourglass
   
   strVal = "select eep01,eep02 from EmpElectronProcess where EEP13='Y' and EEP09 is null and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ",'" & EMP_判發 & "')" & _
            " union select eep01,eep02 from EmpElectronProcess where EEP13='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and EEP04='" & EMP_聯絡 & "'" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "")
   strSql = "Select ' ' as V,EP01 as 目次,SqlDateT(EEP06)||' '||sqltime(EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,LC05||LC06||LC07 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(LC15,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "EEP01 as 總收文號,EEP02 as 序號,EP08,EP38,' ' as 不顯示,EEP06 a,EEP07 b,EEP15,EEP11" & _
            " From EmpElectronProcess,CaseProgress,EngineerProgress,LawCase," & _
            "staff s1,staff s2,nation,CasePropertyMap,allcode" & _
            " Where (EEP01,EEP02) in(" & strVal & ")" & _
            " and EEP01=CP09(+)" & _
            " and cp158=0 and cp159=0" & _
            " And EEP01=EP02(+)" & _
            " And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And LC15=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And EEP04=ac02(+)"
   'If ProState = "1" Then '個人
      strSql = strSql & " And CP14='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   'End If
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(EEP06)||' '||sqltime(EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,HC06 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode('000','000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "EEP01 as 總收文號,EEP02 as 序號,EP08,EP38,' ' as 不顯示,EEP06 a,EEP07 b,EEP15,EEP11" & _
            " From EmpElectronProcess,CaseProgress,EngineerProgress,HireCase," & _
            "staff s1,staff s2,nation,CasePropertyMap,allcode" & _
            " Where (EEP01,EEP02) in(" & strVal & ")" & _
            " and EEP01=CP09(+)" & _
            " and cp158=0 and cp159=0" & _
            " And EEP01=EP02(+)" & _
            " And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And '000'=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And EEP04=ac02(+)" & _
            IIf(Pub_StrUserSt15 = "P22", " And EEP04 not in('" & EMP_判發 & "','" & EMP_退件重送 & "')", "")
   'If ProState = "1" Then '個人
      strSql = strSql & " And CP14='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   'End If
   'Add By Sindy 2022/11/18
   '法務
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(EEP06)||' '||sqltime(EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,SP05 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode('000','000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "EEP01 as 總收文號,EEP02 as 序號,EP08,EP38,' ' as 不顯示,EEP06 a,EEP07 b,EEP15,EEP11" & _
            " From EmpElectronProcess,CaseProgress,EngineerProgress,ServicePractice," & _
            "staff s1,staff s2,nation,CasePropertyMap,allcode" & _
            " Where (EEP01,EEP02) in(" & strVal & ")" & _
            " and EEP01=CP09(+)" & _
            " and cp158=0 and cp159=0" & _
            " And EEP01=EP02(+)" & _
            " And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And '000'=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And EEP04=ac02(+)" & _
            IIf(Pub_StrUserSt15 = "P22", " And EEP04 not in('" & EMP_判發 & "','" & EMP_退件重送 & "')", "")
   'If ProState = "1" Then '個人
      strSql = strSql & " And CP14='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   'End If
   '2022/11/18 END
   strSql = strSql & " order by a desc,b desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd2.Recordset = rsTmp
      For i = 1 To grd2.Rows - 1
         Call SetColColor(i)
      Next i
      cmdDetail.Enabled = True
   Else
      cmdDetail.Enabled = False
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
      
ExitQuery:
   '若有資料時游標停在第一筆
   If bolFirst = True Then
      grd2.Visible = False
      grd2.col = 0
      grd2.row = 1
      If rsTmp.RecordCount > 0 Then
         dblPrevRow = grd2.row
         grd2.Text = "V"
         m_intRow = 1: m_intCol = 0
         For i = 0 To grd2.Cols - 1
            'Modify By Sindy 2020/11/30
            If i <> 4 Then
            '2020/11/30 END
               grd2.col = i
               If grd2.CellBackColor <> &H8080FF Then
                  grd2.CellBackColor = &HFFC0C0
               End If
            End If
         Next i
      End If
      grd2.Visible = True
   End If
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'Add By Sindy 2018/4/17
Private Sub SetColColor(intRow As Integer)
Dim i As Integer
   
   If intRow < 1 Then Exit Sub
   grd2.row = intRow
   '退件要淺紅色表示
   If grd2.TextMatrix(intRow, 12) = "退件" Then
      grd2.col = 12
      grd2.CellBackColor = &HC0C0FF
   End If
   'Add By Sindy 2020/11/30 多案時,案件名稱變桃粉色
   If grd2.TextMatrix(intRow, 20) <> "" And _
      InStr(grd2.TextMatrix(intRow, 21), "多案單筆歷程") > 0 Then
      grd2.col = 4
      grd2.CellBackColor = &HFF00FF 'QBColor(Rnd * 5)
   End If
End Sub

'Add By Sindy 2018/4/17
Private Sub SetGrd2()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2020/11/30 + EEP15,EEP11
   arrGridHeadText = Array("V", "目次", "流程日期", "本所案號", "案件名稱", _
                           "國家", "種類", "案件性質", "本所期限", "承辦人", _
                           "承辦期限", "智權人員", "目前流程狀態", _
                           "總收文號", "序號", "EP08", "EP38", "不顯示", "EEP06 a", "EEP07 b", "EEP15", "EEP11")
   arrGridHeadWidth = Array(200, 400, 800, 1400, 1000, _
                            700, 0, 900, 800, 600, _
                            800, 600, 600, _
                            0, 0, 0, 0, 600, 0, 0, 0, 0)
   grd2.Visible = False
   grd2.Cols = UBound(arrGridHeadText) + 1
   grd2.Rows = 2
   For iRow = 0 To grd2.Cols - 1
      grd2.row = 0
      grd2.col = iRow
      grd2.Text = arrGridHeadText(iRow)
      grd2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      If iRow = 11 Or iRow = 12 Then
         grd2.CellAlignment = flexAlignLeftCenter
      Else
         grd2.CellAlignment = flexAlignCenterCenter
      End If
   Next
   grd2.Visible = True
End Sub

'Add By Sindy 2018/4/20 核稿人
Private Sub Combo2_LostFocus()
   Dim ii As Integer
   Dim blnMatch As Boolean
   
   For ii = 0 To Me.Combo2.ListCount - 1
       blnMatch = False
       If Trim(Left(Me.Combo2.List(ii), 6)) = Trim(Left(Me.Combo2.Text, 6)) Then
           Me.Combo2.ListIndex = ii
           blnMatch = True
           Exit For
       End If
   Next ii
   If blnMatch = False Then
      If Trim(Left(Me.Combo2.Text, 6)) = "" Then
'         Me.Combo2.ListIndex = 0
      Else
         If Len(GetPrjSalesNM(Trim(Left(Me.Combo2.Text, 6)))) = 0 Then
            'Modify By Sindy 2022/12/6
            'Call ShowStaffErr(Trim(Left(Me.Combo2.Text, 6)))
            Call PUB_GetStaffNameDept(Trim(Left(Me.Combo2.Text, 6)), strExc(10), strExc(0), True, False)
            '2022/12/6 END
            Me.Combo2.SetFocus
            Exit Sub
         Else
            Combo2.Text = Trim(Left(Me.Combo2.Text, 6)) & " ==> " & GetPrjSalesNM(Trim(Left(Me.Combo2.Text, 6)))
         End If
      End If
   End If
End Sub
'2018/4/20 END

Private Sub Combo5_Click()
   If Me.Visible = True Then
      If QueryData(True) = False Then ShowNoData 'Add By Sindy 2023/4/12
   End If
End Sub

'Add By Sindy 2018/4/20 判發人
Private Sub Combo6_LostFocus()
   Dim ii As Integer
   Dim blnMatch As Boolean
   
   For ii = 0 To Me.Combo6.ListCount - 1
       blnMatch = False
       If Trim(Left(Me.Combo6.List(ii), 6)) = Trim(Left(Me.Combo6.Text, 6)) Then
           Me.Combo6.ListIndex = ii
           blnMatch = True
           Exit For
       End If
   Next ii
   If blnMatch = False Then
      If Trim(Left(Me.Combo6.Text, 6)) = "" Then
         Me.Combo6.ListIndex = 0
      Else
         If Len(GetPrjSalesNM(Trim(Left(Me.Combo6.Text, 6)))) = 0 Then
            'Modify By Sindy 2022/12/6
            'Call ShowStaffErr(Trim(Left(Me.Combo6.Text, 6)))
            Call PUB_GetStaffNameDept(Trim(Left(Me.Combo6.Text, 6)), strExc(10), strExc(0), True, False)
            '2022/12/6 END
            Me.Combo6.SetFocus
            Exit Sub
         Else
            Combo6.Text = Trim(Left(Me.Combo6.Text, 6)) & " ==> " & GetPrjSalesNM(Trim(Left(Me.Combo6.Text, 6)))
         End If
      End If
   End If
End Sub
'2015/5/21 END

'Add By Sindy 2018/4/17
Private Sub grd2_DblClick()
Dim nFrm As Form
   
   If m_intRow <> 0 Then
      If m_intCol <> 17 Then
         If cmdDetail.Enabled = False Then Exit Sub

         If dblPrevRow = 0 Then
            MsgBox "請點選一筆資料列!", vbExclamation
            Exit Sub
         End If
         
         If grd2.TextMatrix(dblPrevRow, 0) = "V" Then
            Call Process(grd2.TextMatrix(dblPrevRow, 13)) '要重新查詢資料,因核稿人及判發人有預設問題
            If Me.cmd(5).Enabled = True Then
               '重新檢查欄位有效性
               If TxtValidate = True Then
'                  '檢查表單是否已開啟，若是，則關閉
'                  For Each nFrm In Forms
'                     If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'                        Unload frm090202_2
'                        Exit For
'                     End If
'                  Next
                  If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
                  intBackTab = 2
                  frm090202_2.Hide
                  frm090202_2.m_EEP01 = grd2.TextMatrix(dblPrevRow, 13) '總收文號
                  frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) '案件流程所屬人員
                  frm090202_2.intReceiveKind = 0 '0.承辦人工作進度
                  frm090202_2.SetParent Me
                  If frm090202_2.QueryData = True Then
                     frm090202_2.Show
                     Me.Hide
                  End If
               End If
            Else
               Me.SSTab1.Tab = 1
            End If
         End If
      End If
   End If
End Sub

'Add By Sindy 2018/4/17
Private Sub GRD2_SelChange()
Dim j As Integer

grd2.Visible = False
If grd2.MouseRow = 0 Then
   '已選取的資料列清除反白
   For j = 1 To grd2.Rows - 1
      If grd2.TextMatrix(j, 0) = "V" Then
         grd2.col = 0
         grd2.row = j
         grd2.Text = ""
         For i = 0 To grd2.Cols - 1
            'Modify By Sindy 2020/11/30
            If i <> 4 Then
            '2020/11/30 END
               grd2.col = i
               grd2.CellBackColor = QBColor(15)
            End If
         Next i
         Call SetColColor(j)
         Exit For
      End If
   Next j
Else
   '上一筆資料列清除反白
   If dblPrevRow > 0 And dblPrevRow <= (grd2.Rows - 1) Then
      grd2.col = 0
      grd2.row = dblPrevRow
      grd2.Text = ""
      For i = 0 To grd2.Cols - 1
         'Modify By Sindy 2020/11/30
         If i <> 4 Then
         '2020/11/30 END
            grd2.col = i
            If grd2.CellBackColor <> &H8080FF Then
               grd2.CellBackColor = QBColor(15)
            End If
         End If
      Next i
      Call SetColColor(CStr(dblPrevRow))
   End If
   '目前資料列反白
   grd2.col = 0
   grd2.row = grd2.MouseRow
   dblPrevRow = grd2.row
   If grd2.TextMatrix(grd2.row, 1) <> "" Then
      grd2.Text = "V"
      For i = 0 To grd2.Cols - 1
         'Modify By Sindy 2020/11/30
         If i <> 4 Then
         '2020/11/30 END
            grd2.col = i
            If grd2.CellBackColor <> &H8080FF Then
               grd2.CellBackColor = &HFFC0C0
            End If
         End If
      Next i
   End If
End If
grd2.Visible = True
End Sub

'Add By Sindy 2018/4/17
Private Sub grd2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow grd2, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   'grd2.col = nCol
   grd2.row = nRow
   If Me.grd2.row < 1 And Me.grd2.Text <> "V" Then
      If Me.grd2.Text = "目次" Then
         If m_blnColOrderAsc = True Then
            Me.grd2.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.grd2.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.grd2.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.grd2.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

'Add By Sindy 2018/4/17 不顯示功能
Private Sub grd2_Click()
   m_intRow = grd2.MouseRow
   m_intCol = grd2.MouseCol
   If m_intRow <> 0 Then
      If m_intCol = 17 Then '不顯示
         If grd2.TextMatrix(m_intRow, 13) <> "" And _
            grd2.TextMatrix(m_intRow, 12) <> "核修" And _
            grd2.TextMatrix(m_intRow, 12) <> "核完" And _
            grd2.TextMatrix(m_intRow, 12) <> "會修" And _
            grd2.TextMatrix(m_intRow, 12) <> "會完" And _
            grd2.TextMatrix(m_intRow, 12) <> "繪圖判發" And _
            grd2.TextMatrix(m_intRow, 12) <> "判發" And _
            grd2.TextMatrix(m_intRow, 12) <> "退回" And _
            grd2.TextMatrix(m_intRow, 12) <> "退件" And _
            grd2.TextMatrix(m_intRow, 12) <> "圖修" And _
            grd2.TextMatrix(m_intRow, 12) <> "圖完" Then
            grd2.TextMatrix(m_intRow, 17) = "V"
            If MsgBox("請再次確定不顯示 " & vbCrLf & grd2.TextMatrix(m_intRow, 3) & " " & grd2.TextMatrix(m_intRow, 12) & " 嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               grd2.TextMatrix(m_intRow, 17) = ""
            Else
               strExc(0) = "update EmpElectronProcess set eep13=null" & _
                           " where eep01='" & grd2.TextMatrix(m_intRow, 13) & "'" & _
                             " and eep02=" & grd2.TextMatrix(m_intRow, 14)
               Pub_SeekTbLog strExc(0) 'Add By Sindy 2020/11/30
               cnnConnection.Execute strExc(0)
               grd2.RowHeight(m_intRow) = 0
            End If
         End If
      End If
   End If
End Sub

Private Sub Combo1_Click()
   Dim iMouse As Integer
   iMouse = Screen.MousePointer
   
   Me.grd1.Visible = False
   Screen.MousePointer = vbHourglass
   Me.MousePointer = vbHourglass
   grd1.MousePointer = flexArrowHourGlass
   Me.Enabled = False
   Combo1.Enabled = False
'   Label3.Caption = GetPrjSalesNM(Trim(Left("" & Combo1.Text, 6)))
   StrMenu1
   StrMenu
'   If ChkNoData = True Then
'      For s = 0 To 10
'         txt1(s).Enabled = False
'      Next s
'   Else
'      For s = 0 To 10
'         txt1(s).Enabled = True
'      Next s
'   End If
   SetGrd1
   DoEvents
   'cmdok2(0).SetFocus
   Combo1.Enabled = True
   Me.Enabled = True
   grd1.MousePointer = flexDefault
   Me.MousePointer = vbDefault
   Screen.MousePointer = iMouse
   
   Me.grd1.Visible = True
End Sub

Private Sub Form_Activate()
   ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_Load()
Dim iMouse As Integer
Dim nFrm As Form
   
   m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
   iMouse = Screen.MousePointer
   ReDim m_FieldList(TF_CP)
   
   InitialField
   MoveFormToCenter Me
   '讀取各基本檔可用系統別
   m_SqlGrpStr1 = SQLGrpStr("", 1)
   m_SqlGrpStr2 = SQLGrpStr("", 2)
   m_SqlGrpStr3 = SQLGrpStr("", 3)
   m_SqlGrpStr4 = SQLGrpStr("", 4)
   m_SqlGrpStr5 = SQLGrpStr("", 5)
   
   ReDim skMail(0) As SeekMails
   
   If PUB_GetST05(strUserNum) = "91" Or PUB_GetST05(strUserNum) = "92" Then
      '指定會稿日
      txt1(18).Enabled = True
   Else
      txt1(18).Enabled = False
   End If
   
   Select Case ProState
   Case "1" '個人
      '讀取使用權限
      Me.Caption = "工作進度資料維護 (個人)" 'Add By Sindy 2018/8/13
      bolInsert = IsUserHasRightOfFunction("frm090201_4", strAdd, False)
      bolUpdate = IsUserHasRightOfFunction("frm090201_4", strEdit, False)
      bolDelete = IsUserHasRightOfFunction("frm090201_4", strDel, False)
      bolSelect = IsUserHasRightOfFunction("frm090201_4", strFind, False)
      bolPrint = IsUserHasRightOfFunction("frm090201_4", strPrint, False)
      
      TextOk = True
      '統計年月(個人抓系統日的年月)
      Text1.Text = Mid(strSrvDate(1), 1, 6)
   Case "2" '主管 承辦人管理工作進度資料查詢
      Me.Caption = "工作進度資料維護 (主管)" 'Add By Sindy 2018/8/13
      bolInsert = IsUserHasRightOfFunction("frm090614", strAdd, False)
      bolUpdate = IsUserHasRightOfFunction("frm090614", strEdit, False)
      bolDelete = IsUserHasRightOfFunction("frm090614", strDel, False)
      bolSelect = IsUserHasRightOfFunction("frm090614", strFind, False)
      bolPrint = IsUserHasRightOfFunction("frm090614", strPrint, False)
      
      frm090614.TextOk = True
      cmdOK(2).Caption = "回前畫面"
      '統計年月(管理抓查詢畫面輸入的發文年月)
      Text1.Text = Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2))
   Case "3" '分所
      bolInsert = IsUserHasRightOfFunction("frm090201_4", strAdd, False)
      bolUpdate = IsUserHasRightOfFunction("frm090201_4", strEdit, False)
      bolDelete = IsUserHasRightOfFunction("frm090201_4", strDel, False)
      bolSelect = IsUserHasRightOfFunction("frm090201_4", strFind, False)
      bolPrint = IsUserHasRightOfFunction("frm090201_4", strPrint, False)
   Case Else
   End Select
   
   Screen.MousePointer = vbHourglass
   Select Case ProState
   Case "1"
         'Add By Sindy 2018/4/18
         Combo1.AddItem strUserNum & " " & "(" & strUserName & ")", 0
         Combo1.Text = Combo1.List(0)
         '2013/9/17 END
         StrMenu1 'Modify By Sindy 2016/9/6 因前句Combo1就會run 到 StrMenu1
         SetEngineer '設定承辦人選單
         '檢查當時是否需要為他人職代
         Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
         '2018/4/18 END
         'StrMenu1
   Case "2" '承辦人管理工作進度資料查詢
         frm090614.Process3
         StrMenu1
   Case "3"
   Case Else
   End Select
   
'   Label3.Caption = GetPrjSalesNM(Trim(Left("" & Combo1.Text, 6)))
   
   StrMenu
   
   Select Case ProState
   Case "1"
      If TextOk = False Then Screen.MousePointer = iMouse: GoTo EXITSUB
      'Add By Sindy 2018/4/18
      'Combo1.Enabled = False
      Combo1.Enabled = True
      '2018/4/18 END
   Case "2"
      If frm090614.TextOk = False Then Screen.MousePointer = iMouse: TextOk = True: GoTo EXITSUB
      Combo1.Enabled = True
   Case "3"
   Case Else
   End Select
   
   SetGrd1
   'Modify By Sindy 2018/8/23
   'MouseClick (1)
   MouseClick_1 (1)
   '2018/8/23 END
   Screen.MousePointer = iMouse
   SSTab1.Tab = 0
   Me.Combo3.ListIndex = 0
   
   'Add By Sindy 2018/4/26
   SSTab1.Tab = 2
   If QueryData(True) = False Then
      SSTab1.Tab = 0
   End If
   '2018/4/26 END
   
   If bolUpdate = False Then
      cmdOK(1).Visible = False
   End If
   
   Me.txt1(12).Enabled = False 'Added by Lydia 2019/05/03 承辦期限不可修改(應該是沒有人強調,所以都沒限制)
   
   Exit Sub

EXITSUB:
   Me.Hide
   Select Case ProState
   Case "1"
        Me.Hide
   Case "2"
        frm090614.Show
        Me.Hide
   Case "3"
   Case Else
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   
   'Add by Sindy 2022/12/17 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q") = True Then
      Unload frm090801_Q
   End If
   '2022/12/17 END
   
   Set Fobj = New FileSystemObject
   Fobj.DeleteFile DocTempPath & "\*.doc", False
   ClearFieldList
   Set Fobj = Nothing
   
   Set frm090201_d = Nothing
End Sub

Sub StrMenu1()
Me.Enabled = False
DoEvents
On Error GoTo ErrHnd 'Add By Sindy 2024/3/14
adoEng.Execute "drop table R090614 "
'Modify By Sindy 2015/9/10 +,R110033 text
RunCreateTable: 'Add By Sindy 2024/3/14
adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text" & _
               ",R110006 text,R110007 text,R110008 text,R110009 text,R110010 text" & _
               ",R110011 text,R110012 text,R110013 text,R110014 text,R110015 text" & _
               ",R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo" & _
               ",R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text" & _
               ",R110026 double,R110027 double,R110028 double,R110029 text,R110030 text" & _
               ",R110031 text,R110032 double,R110033 text)"
On Error GoTo 0 'Add by Sindy 2024/3/14 還原錯誤控制

Select Case ProState
Case "1" '承辦人個人工作進度資料維護
      StrGrp090201 = ""
      StrSQL6 = ""
      strSQL1 = ""
      strSQL2 = ""
      StrSQL61 = ""
      StrSQL62 = ""
      StrSQL63 = ""
      StrSQL64 = ""
      StrSTM = ""
      StrSLC = ""
      StrSHC = ""
      StrSSP = ""
        
      StrSQL6 = StrSQL6 & " and CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' and cp05>=19980101 "
      StrSQL61 = StrSQL61 & " and CP27 IS NULL  and CP57 IS NULL  "
      StrSQL62 = StrSQL62 & " and CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 "
      StrSQL63 = StrSQL63 & " and CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp27 is null   "
      StrSQL64 = StrSQL64 & " and CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP05<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp57 is null And CP05>CP27  "
      StrSTM = StrSTM & " and ((tm30>=" & Mid(strSrvDate(1), 1, 6) & "01 AND tm30<=" & Mid(strSrvDate(1), 1, 6) & "31) or tm30 is null) "
      StrSLC = StrSLC & " and ((lc09>=" & Mid(strSrvDate(1), 1, 6) & "01 AND lc09<=" & Mid(strSrvDate(1), 1, 6) & "31) or lc09 is null) "
      StrSHC = StrSHC & " and ((hc10>=" & Mid(strSrvDate(1), 1, 6) & "01 AND hc10<=" & Mid(strSrvDate(1), 1, 6) & "31) or hc10 is null) "
      StrSSP = StrSSP & " and ((sp16>=" & Mid(strSrvDate(1), 1, 6) & "01 AND sp16<=" & Mid(strSrvDate(1), 1, 6) & "31) or sp16 is null) "

Case "2" '承辦人管理工作進度資料查詢
      StrGrp090201 = frm090614.ManaGrp
      '改成收文日要小於等於發文年月當月的最後一天
      StrSQL6 = " and cp05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 "
      strSQL1 = ""
      strSQL2 = ""
      StrSQL61 = ""
      StrSQL62 = ""
      StrSQL63 = ""
      StrSQL64 = ""
      StrSTM = " and ((tm30>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and tm30<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or tm30 is null) "
      StrSLC = " and ((lc09>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and lc09<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or lc09 is null) "
      StrSHC = " and ((hc10>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and hc10<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or hc10 is null) "
      StrSSP = " and ((sp16>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and sp16<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or sp16 is null) "
      If frm090614.txt1(8) = "N" Then
         StrSQL6 = StrSQL6 & " and CP14 IN (" & Combo1_String & ")  and cp05>=19980101 "
         StrSQL61 = StrSQL61 & " and CP27 IS NULL  and CP57 IS NULL "
         StrSQL62 = StrSQL62 & " and CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01  "
         StrSQL63 = StrSQL63 & " and CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27 is null "
         StrSQL64 = StrSQL64 & " and CP05>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null And CP05>CP27 "
      Else
         StrSQL6 = StrSQL6 & " and CP14='" & Trim(Left("" & Combo1.Text, 6)) & "'  and cp05>=19980101 "
         StrSQL61 = StrSQL61 & " and CP27 IS NULL  and CP57 IS NULL "
         StrSQL62 = StrSQL62 & " and CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 "
         StrSQL63 = StrSQL63 & " and CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27 is null "
         StrSQL64 = StrSQL64 & " and CP05>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null And CP05>CP27 "
      End If
Case Else
End Select

CheckOC

'and cp05>=19980101 and CP27 IS NULL  and CP57 IS NULL
'and cp05>=19980101 and CP27>=20181001 AND CP27<=20181031
'and cp05>=19980101 and CP57>=20181001 AND CP57<=20181031 and cp27 is null
'and cp05>=19980101 and CP05>=20181001 AND CP05<=20181031 and cp57 is null And CP05>CP27

'Modify By Sindy 2015/9/10 增加讀取cp140
'第一次
'Modify By Sindy 2018/8/1 當月資料剔除核准1001(已發文才剔除)及註冊證1701資料(+  and cp10<>'1701' and not(cp10='1001' and cp158>0))
strSql = "SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL61 & StrSTM & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
strSql = strSql + " UNION   SELECT CP14,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL61 & StrSLC & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL61 & StrSHC & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL61 & StrSSP & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
'第二次
'Modify By Sindy 2018/8/1 當月資料剔除核准1001(已發文才剔除)及註冊證1701資料(+  and cp10<>'1701' and not(cp10='1001' and cp158>0))
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL62 & StrSTM & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
strSql = strSql + " UNION   SELECT CP14,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL62 & StrSLC & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL62 & StrSHC & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL62 & StrSSP & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
'第三次
'Modify By Sindy 2018/8/1 當月資料剔除核准1001(已發文才剔除)及註冊證1701資料(+  and cp10<>'1701' and not(cp10='1001' and cp158>0))
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL63 & StrSTM & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
strSql = strSql + " UNION   SELECT CP14,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL63 & StrSLC & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL63 & StrSHC & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL63 & StrSSP & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
'第四次
'Modify By Sindy 2018/8/1 當月資料剔除核准1001(已發文才剔除)及註冊證1701資料(+  and cp10<>'1701' and not(cp10='1001' and cp158>0))
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL64 & StrSTM & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
strSql = strSql + " UNION   SELECT CP14,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL64 & StrSLC & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL64 & StrSHC & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
                " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL64 & StrSSP & _
                " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
AddToMdb (strSql)
Me.Enabled = True

'Add By Sindy 2024/3/14
Exit Sub

ErrHnd:
   GoTo RunCreateTable
'2024/3/14 END
End Sub

Sub AddToMdb(oStrSQL As String)
Dim strCP09s As String

CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 And .RecordCount > 0 Then
      .MoveFirst
      k = 0
      strCP09s = "''"
      Do While .EOF = False
         strCP09s = strCP09s & ",'" & .Fields("cp09") & "'"
         For i = 0 To 26
            strTemp(i) = CheckStr(.Fields(i))
            If Len(strTemp(i)) = 8 Then
               If Mid(strTemp(i), 3, 1) = "/" And Mid(strTemp(i), 6, 1) = "/" Then
                  strTemp(i) = " " & strTemp(i)
               End If
            End If
         Next i
         'Modify By Sindy 2015/9/10 +,'" & .Fields("cp140").Value & "'
         strSql = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "','" & .Fields(28).Value & "'," & Val("" & .Fields("cp97")) & "," & Val("" & .Fields("cp98")) & "," & Val("" & .Fields("cp111")) & ",'" & "" & .Fields("ep34").Value & "','" & "" & .Fields("cp112").Value & "','" & .Fields("ep28").Value & "',0,'" & .Fields("cp140").Value & "') "
         adoEng.Execute strSql
         .MoveNext
      Loop
   End If
End With
CheckOC
End Sub

Sub ChgGrdColor(Optional iRow As Integer = -1)
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim ColorFlag As String
Dim iStart As Integer, iEnd As Integer

With grd1
   If iRow >= 0 Then
      iStart = iRow
      iEnd = iRow
   Else
      iStart = 1
      iEnd = .Rows - 1
   End If
   For i = iStart To iEnd
      DoEvents
      .row = i
      .col = 21 '承辦人備註
'      ColorFlag = Mid(.Text, 1, 1)
      '.Text = Mid(.Text, 2)
      .Text = Mid(.Text, 1)
'      If ColorFlag = "1" And Mid(Pub_StrUserSt15, 1, 2) <> "P2" Then
'         .col = 4
'         .CellBackColor = QBColor(10) '淡綠色
'      End If
      .col = 24
      Tmp003 = Trim(.Text)
      '若有取消收文日期
      If Tmp003 <> "" Then
         '灰色
         .col = 3
         .CellBackColor = QBColor(8)
         .col = 10
         .CellBackColor = QBColor(8)
         .col = 11
         .CellBackColor = QBColor(8)
         .col = 13
         .CellBackColor = QBColor(8)
      Else
         If .TextMatrix(i, 1) & SystemNumber(.TextMatrix(i, 3), 1) <> "CP" And .TextMatrix(i, 25) <> "N" Then
            .col = 10
            Tmp001 = Trim(.Text)
            .col = 16
            Tmp002 = Trim(.Text)
            .col = 24
            Tmp003 = Trim(.Text)
            '若有承辦期限, 無會稿日及取消收文日期
            If Tmp001 <> "" And Tmp002 = "" And Tmp003 = "" Then
               If Val(ChangeTStringToWString(ChangeTDateStringToTString(Tmp001))) < Val(strSrvDate(1)) Then
                  '黃色
                  .col = 3
                  .CellBackColor = &H80FFFF
                  .col = 10
                  .CellBackColor = &H80FFFF
                  .col = 11
                  .CellBackColor = &H80FFFF
                  .col = 13
                  .CellBackColor = &H80FFFF
               End If
            Else
               '若是有會稿日，且過承辦期限，給淡黃色
               If Tmp001 <> "" And Tmp002 <> "" And Tmp003 = "" Then
                  If Val(ChangeTStringToWString(ChangeTDateStringToTString(Tmp001))) < Val(ChangeTStringToWString(ChangeTDateStringToTString(Tmp002))) Then
                     '淡黃色
                     .col = 3
                     .CellBackColor = &HC0FFFF
                     .col = 10
                     .CellBackColor = &HC0FFFF
                     .col = 11
                     .CellBackColor = &HC0FFFF
                     .col = 13
                     .CellBackColor = &HC0FFFF
                   End If
               End If
            End If
         End If
         .col = 19
         '若無發文日
         If .Text = "" Then
            .col = 9
            '若系統日大於等於本所期限且本所期限有值(逾本所期限未發文)
            If Val(ChangeTStringToWString(ChangeTDateStringToTString(Trim(.Text)))) <= Val(strSrvDate(1)) And Trim(.Text) <> "" Then
               '淺紅色
               .col = 3
               .CellBackColor = &HC0C0FF
               .col = 10
               .CellBackColor = &HC0C0FF
               .col = 11
               .CellBackColor = &HC0C0FF
               .col = 13
               .CellBackColor = &HC0C0FF
            End If
         '若有發文日
         Else
            .col = 13
            Tmp001 = Trim(.Text)
            .col = 14
            Tmp002 = Trim(.Text)
            .col = 16
            Tmp003 = Trim(.Text)
            .col = 18
            Tmp004 = Trim(.Text)
            .col = 24
            If (Tmp001 = "" Or Tmp002 = "" Or Tmp003 = "" Or Tmp004 = "") Then
               .col = 18
               .Text = " ******"
            End If
         End If
      End If
   Next i
   '預設目前在第一筆的位置
   With Me.grd1
      .row = 1
      .col = 0
      .CellBackColor = &HFFC0C0
      .col = 12
      .CellBackColor = &HFFC0C0
      SWPColor2 = SWPColor
      SWPRow2 = .row
   End With
   SetGrd1
End With
End Sub

Sub StrMenu()
Dim iMouse As Integer
iMouse = Screen.MousePointer
Select Case ProState
Case "1"
      'Modify By Sindy 2015/9/10 增加CP140電子表單單號做排序,自動收文在前面
      'Modify By Sindy 2018/8/1 + AND R110018='' and (R110024='' or R110024='0') : 進入後只出現未發文案件
      'Modify By Sindy 2023/2/8 原案件清單排列,以電子收文為優先
      '但現皆為電子收文案件,會導致當日智慧局來文不易被發現,請調整為以"目次"由大至小排列
      '取消 R110033 desc,
      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028)  or r110028=0,1,r110028))+R110032,'0.00'),R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032 FROM R090614 " & _
                    " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' AND R110018='' and (R110024='' or R110024='0')" & _
                    " ORDER BY R110002 desc,R110003,R110004 "

Case "2"
      If frm090614.txt1(8) = "N" Then 'N：不區分個人
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028)  or r110028=0,1,r110028))+R110032,'0.00'),R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032 FROM R090614 " & _
                     " WHERE ID='" & strUserNum & "' AND R110001 IN (" & Combo1_String & ") ORDER BY R110005,R110002 desc,  R110004 "
      Else
         'Modify By Sindy 2015/9/10 增加CP140電子表單單號做排序,自動收文在前面
         'Modify By Sindy 2023/2/8 原案件清單排列,以電子收文為優先
         '但現皆為電子收文案件,會導致當日智慧局來文不易被發現,請調整為以"目次"由大至小排列
         '取消 R110033 desc,
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028) or r110028=0,1,r110028))+R110032,'0.00'),R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032 FROM R090614 " & _
                     " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
                     " ORDER BY R110002 desc, R110003, R110004 "
      End If
Case "3"
Case Else
End Select
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        If ProState = "2" Then
            InsertQueryLog (.RecordCount)
        End If
        Set grd1.Recordset = adoRecordset
        ChkNoData = False
    Else
        If ProState = "2" Then
            InsertQueryLog (0)
        End If
        ChkNoData = True
        grd1.Clear
        grd1.Rows = 2
        Screen.MousePointer = iMouse
        Exit Sub
    End If
End With
CheckOC
ChgGrdColor
    SWPRow2 = "1"
    grd1.row = Val(SWPRow2)
    grd1.col = 1
End Sub

Private Sub SetGrd1()
With grd1
    .Visible = False
    .Cols = 29
    .row = 0
    .col = 0:   .Text = "目次"
    .ColWidth(0) = 350
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "收文類別"
    .ColWidth(1) = 200
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "收文日"
    .ColWidth(2) = 795
    .ColAlignment(2) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "本所案號"
    .ColWidth(3) = 1005
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "案件名稱"
    .ColWidth(4) = 1155
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "國家"
    .ColWidth(5) = 450
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "種類"
    .ColWidth(6) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "案件性質"
    .ColWidth(7) = 795
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "Y/N"
    .ColWidth(8) = 285
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "本所期限"
    .ColWidth(9) = 795
    .ColAlignment(9) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 10:   .Text = "承辦期限"
    .ColWidth(10) = 795
    .ColAlignment(10) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 11:  .Text = "考核值"
    .ColWidth(11) = 0 '435
    .ColAlignment(11) = flexAlignRightCenter
    .CellAlignment = flexAlignLeftCenter
    .col = 12:  .Text = "法定期限"
    .ColWidth(12) = 0
    .ColAlignment(12) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 13:  .Text = "齊備日"
    .ColWidth(13) = 795
    .ColAlignment(13) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 14:  .Text = "完稿日"
    .ColWidth(14) = 795
    .ColAlignment(14) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 15:  .Text = "指會日"
    .ColWidth(15) = 795
    .ColAlignment(15) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 16:  .Text = "會稿日"
    .ColWidth(16) = 795
    .ColAlignment(16) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 17:  .Text = "核稿人"
    .ColWidth(17) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 18:  .Text = "會稿完成日"
    .ColWidth(18) = 795
    .ColAlignment(18) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 19:  .Text = "發文日"
    .ColWidth(19) = 795
    .ColAlignment(19) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 20:  .Text = "承辦天數"
    .ColWidth(20) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 21:  .Text = "備註"
    .ColWidth(21) = 2000
    .CellAlignment = flexAlignCenterCenter
    .col = 22:  .Text = "智權人員"
    .ColWidth(22) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 23:  .Text = ""
    .ColWidth(23) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 24:  .Text = ""
    .ColWidth(24) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 25:  .Text = ""
    .ColWidth(25) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 26:  .Text = ""
    .ColWidth(26) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 27:  .Text = ""
    .ColWidth(27) = 0
    .CellAlignment = flexAlignCenterCenter
    For intI = 28 To .Cols - 1
      .ColWidth(intI) = 0
    Next
    .Visible = True
End With
   '預設目前在第一筆的位置
   With Me.grd1
      .row = 1
      .col = 0
      .CellBackColor = &HFFC0C0
      .col = 12
      .CellBackColor = &HFFC0C0
      SWPColor2 = SWPColor
      SWPRow2 = .row
   End With
End Sub

Private Sub GRD1_DblClick()
    If Me.grd1.MouseRow > 0 Then
        '若有資料
        If Me.grd1.Rows > 1 Then
            SWPRow = str(grd1.MouseRow)
            '若點選的那筆無資料, 則退出函式
            If Me.grd1.TextMatrix(SWPRow, 1) = "" Then Exit Sub
            SSTab1.Tab = 1
        End If
    End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Strindex As Integer
Dim iMouse As Integer

iMouse = Screen.MousePointer

If Me.grd1.MouseRow <= 0 Then Exit Sub
If Button = 1 Then
    Screen.MousePointer = vbHourglass
    SWPRow = str(grd1.MouseRow)
    Strindex = SWPRow
    With grd1
        DoEvents
        .Visible = False
        If SWPRow2 <> "" Then
           .row = SWPRow2
           .col = 0
           .CellBackColor = QBColor(15)
           .col = 12
           .CellBackColor = QBColor(15)
        End If
        .col = 0
        If Strindex <> 0 Then
            .row = Strindex
        Else
            .row = .MouseRow
        End If
        If .row = 0 Then
            .row = 1
        End If
        .col = 0
        .CellBackColor = &HFFC0C0
        .col = 12
        .CellBackColor = &HFFC0C0
        SWPColor2 = SWPColor
        SWPRow2 = .row
        .Visible = True
    End With
    Screen.MousePointer = iMouse
End If
End Sub

Sub MouseClick(Optional Strindex As Integer = 0)
    Dim iMouse As Integer
    
    iMouse = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    With grd1
        DoEvents
        .Visible = False
        If SWPRow2 <> "" Then
           .row = SWPRow2
           .col = 0
           .CellBackColor = QBColor(15)
           .col = 12
           .CellBackColor = QBColor(15)
        End If
        .col = 0
        If Strindex <> 0 Then
            .row = Strindex
        Else
            .row = .MouseRow
        End If
        If .row = 0 Then
            .row = 1
        End If
        .col = 23
        Process (.Text)
        .col = 0
        .CellBackColor = &HFFC0C0
        .col = 12
        .CellBackColor = &HFFC0C0
        SWPColor2 = SWPColor
        SWPRow2 = .row
        .Visible = True
    End With
    Screen.MousePointer = iMouse
End Sub

'存檔時使用
Sub MouseClick_1(Optional Strindex As Integer = 0)
Dim iMouse As Integer
   
   iMouse = Screen.MousePointer
   
   Screen.MousePointer = vbHourglass
   With grd1
       DoEvents
       .Visible = False
       If SWPRow2 <> "" Then
          .row = SWPRow2
          .col = 0
          .CellBackColor = QBColor(15)
          .col = 12
          .CellBackColor = QBColor(15)
       End If
       .col = 0
       If Strindex <> 0 Then
           .row = Strindex
       Else
           .row = .MouseRow
       End If
       If .row = 0 Then
           .row = 1
       End If
       .col = 0
       .CellBackColor = &HFFC0C0
       .col = 12
       .CellBackColor = &HFFC0C0
       SWPColor2 = SWPColor
       SWPRow2 = .row
       .Visible = True
   End With
   
   Screen.MousePointer = iMouse
End Sub

Private Sub grd1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.grd1.MouseRow < 1 Then
        Select Case Me.grd1.MouseCol
        Case 0
            If m_blnColOrderAsc = True Then
                Me.grd1.Sort = 3 '昇冪
                m_blnColOrderAsc = False
            Else
                Me.grd1.Sort = 4 '降冪
                m_blnColOrderAsc = True
            End If
        Case Else
            If m_blnColOrderAsc = True Then
                Me.grd1.Sort = 5 '昇冪
                m_blnColOrderAsc = False
            Else
                Me.grd1.Sort = 6 '降冪
                m_blnColOrderAsc = True
            End If
        End Select
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 Dim aD1 As Integer 'Added by Lydia 2016/05/06
Dim ii As Integer 'Add By Sindy 2018/4/26
   
   'Add By Sindy 2018/4/26
   If SSTab1.Tab = 2 Then
      Call QueryData(True)
   Else
      Call QueryData(False)
   End If
   If PreviousTab = 2 Then
      '若有資料
      If (Me.grd2.Rows - 1) < dblPrevRow Then dblPrevRow = 0 'Add By Sindy 2018/10/2
      If Me.grd2.Rows > 1 And dblPrevRow > 0 Then
         If Me.grd2.TextMatrix(dblPrevRow, 1) <> "" Then
            For i = 1 To Me.grd1.Rows - 1
               If Me.grd2.TextMatrix(dblPrevRow, 1) = Me.grd1.TextMatrix(i, 0) Then
                  SWPRow = i
                  Exit For
               End If
            Next i
            MouseClick Val(SWPRow)
            If SSTab1.Tab = 1 Then
               SSTab1.Tab = 1
            End If
         End If
      End If
   End If
   If PreviousTab = 0 Or PreviousTab = 1 Then
      '若有資料
      If (Me.grd1.Rows - 1) < Val(SWPRow) Then SWPRow = 0 'Add By Sindy 2018/10/2
      If Me.grd1.Rows > 1 Then
         '若點選的那筆無資料, 則退出函式
         If Me.grd1.TextMatrix(Val("0" & SWPRow), 1) = "" Then SSTab1.Tab = 0: Exit Sub
         If Val(SWPRow) > 0 Then
            '上一筆資料列清除反白
            If dblPrevRow > 0 And dblPrevRow <= (grd2.Rows - 1) Then
               grd2.col = 0
               grd2.row = dblPrevRow
               grd2.Text = ""
               For ii = 0 To grd2.Cols - 1
                  'Modify By Sindy 2020/11/30
                  If ii <> 4 Then
                  '2020/11/30 END
                     grd2.col = ii
                     If grd2.CellBackColor <> &H8080FF Then
                        grd2.CellBackColor = QBColor(15)
                     End If
                  End If
               Next ii
               dblPrevRow = 0
               Call SetColColor(CStr(dblPrevRow))
            End If
            For i = 1 To Me.grd2.Rows - 1
               If Me.grd2.TextMatrix(i, 1) = Me.grd1.TextMatrix(Val("0" & SWPRow), 0) Then
                  '目前資料列反白
                  dblPrevRow = i
                  grd2.col = 0
                  grd2.row = dblPrevRow
                  If grd2.TextMatrix(grd2.row, 1) <> "" Then
                     grd2.Text = "V"
                     For ii = 0 To grd2.Cols - 1
                        'Modify By Sindy 2020/11/30
                        If ii <> 4 Then
                        '2020/11/30 END
                           grd2.col = ii
                           If grd2.CellBackColor <> &H8080FF Then
                              grd2.CellBackColor = &HFFC0C0
                           End If
                        End If
                     Next ii
                  End If
                  Exit For
               End If
            Next i
         End If
      End If
   End If
   '2018/4/26 END
   
   'If PreviousTab = 0 Then
   If SSTab1.Tab = 1 Then
      '若有資料
      If Me.grd1.Rows > 1 Then
         '若點選的那筆無資料, 則退出函式
         If Me.grd1.TextMatrix(Val("0" & SWPRow), 1) = "" Then SSTab1.Tab = 0: Exit Sub
         MouseClick Val(SWPRow)
         SSTab1.Tab = 1
         If cmd(5).Tag = "N" Then
            cmd(5).Enabled = False 'Add By Sindy 2018/5/23
         End If
      End If
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Index = 1 Then
      If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
         KeyAscii = 0
         Beep
      End If
   'Add By Sindy 2018/8/8
   ElseIf Index = 2 Or Index = 3 Or Index = 4 Or Index = 7 Or Index = 8 Or Index = 12 Then
      KeyAscii = Pub_NumAscii(KeyAscii)
   '2018/8/8 END
   End If
End Sub

Private Sub txt1_LostFocus(Index As Integer)
'Add by Morgan 2010/9/6 若回第一頁籤時不檢查,否則若有錯誤時會無窮回圈
If Me.SSTab1.Tab = 0 Then Exit Sub

Select Case Index
Case 1 '是否會稿
     Select Case Trim(txt1(1))
     Case "Y", ""
     Case "N"
         'Add By Sindy 2018/4/25
'         Call ChkEP34ToEP07EP08
         txt1_LostFocus (4)
         '2018/4/25 END
     Case Else
         s = MsgBox("是否會稿只能輸入 Y 或 N !!", , "USER 輸入錯誤")
         txt1(1).SetFocus
         txt1(1).SelStart = 0
         txt1(1).SelLength = Len(txt1(1))
         Exit Sub
     End Select
Case 2 '齊備日
'Mark by Lydia 2019/05/02 改到Validate
'     If Len(Txt1(Index)) <> 0 Then
'         If Not ChkWorkDay(ChangeTStringToWString(Txt1(Index))) Then
'            ShowDateErr
'            txt1(Index).SetFocus
'            txt1(Index).SelLength = Len(txt1(Index))
'            Exit Sub
'         End If
''     Else
''        '若未輸入齊備日則清空承辦期限
''        Me.txt1(12).Text = ""
'     End If
Case 3 '完稿日
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
     End If
Case 4 '會稿日
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
     End If
Case 7 '會稿完成日
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
     End If
Case 8 '發文日
     If Len(txt1(Index)) <> 0 Then
        '若發文日為111111則不檢查是否為工作日
        If Me.txt1(Index).Text <> "111111" Then
            If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
               ShowDateErr
               txt1(Index).SetFocus
               txt1(Index).SelLength = Len(txt1(Index))
               Exit Sub
            End If
        End If
        If txt1(1) = "Y" And Len(txt1(4)) = 0 Then txt1(1) = "N" 'Add By Sindy 2019/7/12
     End If

Case 12 '承辦期限
   'Add By Sindy 2018/8/8
   'Modify By Sindy 2018/9/25 + 排除(102)延展案
   If Len(txt1(Index)) <> 0 And m_CP10 <> "102" Then
        If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
           ShowDateErr
           txt1(Index).SetFocus
           txt1(Index).SelLength = Len(txt1(Index))
           Exit Sub
        End If
    End If
    '2018/8/8 END
    '若有承辦期限
    If Me.txt1(12).Text <> "" And Me.lbl1(17).Caption <> "" Then
        '若承辦期限大於本所期限
        If Val(txt1(12).Text) > Val(Replace(lbl1(17).Caption, "/", "")) Then
            Me.txt1(12).Text = Replace(lbl1(17).Caption, "/", "")
        End If
    End If
Case Else
End Select

'Add By Sindy 2018/4/25
If Index = 1 Or Index = 3 Then
   Select Case Trim(txt1(1))
   Case "N"
'         Call ChkEP34ToEP07EP08
   Case Else
   End Select
End If
'2018/4/25 END
End Sub

Sub ChkTxt(Strindex As String)
    ChkData = False
    '齊備日
    If Strindex = "2" Or Strindex = "" Then
         If Len(txt1(2)) = 0 Then
             If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Or Len(txt1(8)) <> 0 Then
                 ShowDateRanErr
                 txt1(2).SetFocus
                 txt1_GotFocus (2)
                 Exit Sub
             End If
         End If
    End If
    
    '完稿日
    If Strindex = "3" Or Strindex = "" Then
        If Len(txt1(3)) = 0 Then
            If Len(txt1(4)) <> 0 Or Len(txt1(8)) <> 0 Then
                ShowDateRanErr
                txt1(3).SetFocus
                Exit Sub
            End If
        End If
    End If
    
    '會稿日
    If Strindex = "4" Or Strindex = "" Then
        '無會稿日
        If Len(txt1(4)) = 0 Then
            '有發文日
            If Len(txt1(8)) <> 0 Then
                ShowDateRanErr
                txt1(4).SetFocus
                Exit Sub
            End If
        End If
    End If
    
    '承辦備註
    If Strindex = "10" Or Strindex = "" Then
        If Not CheckLengthIsOK(txtEP12, 2000) Then
            txtEP12.SetFocus
            txt1_GotFocus (10)
            Exit Sub
        End If
    End If
    
    '承辦期限
    If Strindex = "12" Or Strindex = "" Then
        If CheckIsTaiwanDate(Me.txt1(12).Text) = False Then
            MsgBox "承辦期限輸入錯誤！", vbExclamation
            Me.txt1(12).SetFocus
            txt1_GotFocus 12
            Exit Sub
        End If
    End If
    
    ChkData = True
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
'Added by Lydia 2019/05/02 從Lostfocus移過來
Case 2
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            Cancel = True
            Exit Sub
         End If
     End If
Case 3, 4, 8, 9
    '若欄位無資料則不檢查
    If Me.txt1(Index).Text = "" Then Exit Sub
    ChkTxt "" & Index
    If ChkData = False Then
        Cancel = True
        Exit Sub
    End If
    'Add By Sindy 2019/12/3 C類發文日只能小於等於系統日+2個工作天
    'Modify By Sindy 2021/5/24 改判斷不是 系統日 和 19221111 就彈詢問訊息
    If Me.txt1(Index).Enabled = True And Val(Me.txt1(Index).Text) > 0 Then
      If Index = 8 Then '發文日
         If Left(lbl1(3).Caption, 1) = "C" Then
'            If Val(DBDATE(Me.txt1(Index).Text)) > Val(CompWorkDay(3, strSrvDate(1))) Then
'               MsgBox "發文日不可大於系統日+2個工作天！"
'               txt1(8).SetFocus
'               txt1_GotFocus 8
'               Cancel = True
'               Exit Sub
'            End If
            If Not (Val(DBDATE(Me.txt1(Index).Text)) = strSrvDate(1) Or _
                    Val(DBDATE(Me.txt1(Index).Text)) = 19221111) Then
               If MsgBox("確定發文日為 " & ChangeTStringToTDateString(Me.txt1(Index).Text) & " 嗎？" & vbCrLf & vbCrLf & "（注意：不發文應該輸入 11/11/11）", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  txt1(8).SetFocus
                  txt1_GotFocus 8
                  Cancel = True
                  Exit Sub
               End If
            End If
         End If
      End If
    End If
    '2019/12/3 END
    
'指定會稿日
Case 18
     If txt1(Index).Enabled = True And Trim(txt1(Index).Text) <> "" Then
         If ChkWork(ChangeTStringToWString(txt1(Index))) = False Then
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            Cancel = True
            Exit Sub
         End If
         
         If CheckIsTaiwanDate(txt1(Index).Text) = False Then
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            Cancel = True
            Exit Sub
         End If
         
         If txt1(Index) <> txt1(Index).Tag Then
            If Val(DBDATE(txt1(Index))) < Val(strSrvDate(1)) Then
               MsgBox "指定會稿日不可早於系統日！"
               Cancel = True
               Exit Sub
            End If
         End If
     End If
Case Else
End Select
End Sub

Private Function TxtValidate() As Boolean
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim rsB As New ADODB.Recordset
Dim StrSqlB As String
Dim arrCaseNo '本所案號
Dim ii As Integer
Dim blnMatch As Boolean
Dim tmpBol As Boolean  'Added by Lydia 2019/05/02
TxtValidate = False

''檢查承辦期限
'If Me.txt1(12).Text <> "" And txt1(12).Enabled = True Then
'    If Me.txt1(2).Text = "" Then
'        MsgBox "無齊備日不可輸入承辦期限!!!", vbExclamation + vbOKOnly
'        Exit Function
'    End If
'End If

'add by nickc 2006/10/23 有齊備日的，承辦期限若是空白再抓一次
'Modified by Lydia 2019/05/02 T-217900齊備日輸入5/1(勞動節放假),因為載入資料時直接SetFocus會程式出錯,所以到存檔前才檢查
'If txt1(2).Text <> "" And txt1(12) = "" Then txt1_LostFocus 2
If txt1(2).Text <> "" Then
    tmpBol = False
    Call txt1_Validate(2, tmpBol)
    If tmpBol = True Then
        txt1(2).SetFocus
        txt1_GotFocus 2
        Exit Function
    End If
End If
'end 2019/05/02

'Add By Sindy 2019/12/3
If txt1(8).Text <> "" Then
    tmpBol = False
    Call txt1_Validate(8, tmpBol)
    If tmpBol = True Then
        txt1(8).SetFocus
        txt1_GotFocus 8
        Exit Function
    End If
End If
'2019/12/3 END

'Add By Sindy 2018/4/20
'核稿人不可與承辦人相同
If Combo2.Enabled = True And Val(m_CP27) = 0 Then
   '若核判表有設定核稿人時只可以修改但不可以空白
   If Trim(m_PP04) <> "" And Trim(Left(Combo2.Text, 6)) = "" Then
      MsgBox "核稿人不可空白!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1 'Add By Sindy 2024/3/15
      Combo2.SetFocus
      Exit Function
   End If
   If Trim(Left(Combo2.Text, 6)) <> "" Then
      '增加檢查核稿人是否離職
      If PUB_ChkEmpFlowExists(lbl1(3), EMP_送核) = True And m_EP39 = "" Then
         If ChkStaffST04(Trim(Left(Combo2.Text, 6))) = True Then
            SSTab1.Tab = 1
            Combo2.SetFocus
            Exit Function
         End If
      End If
      If UCase(Trim(Left("" & Combo1.Text, 6))) = UCase(Trim(Left(Combo2.Text, 6))) And _
         Not (UCase(m_PP04) = UCase(Trim(Left(Combo2.Text, 6)))) Then
         MsgBox "核稿人不可與承辦人相同!!!", vbExclamation + vbOKOnly
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         Combo2.SetFocus
         Exit Function
      End If
      'Modify By Sindy 2018/10/1 不鎖權限
'      '只要非系統設定的人員均要檢查權限
'      '承辦人非程序人員時,才需檢查核判權限
'      'Modify By Sindy 2018/9/19 And m_PP03 <> "" ==>有設定核判表的案件性質才要檢查權限
'      If GetStaffDepartment(Trim(Left("" & Combo1.Text, 6))) <> "P22" And m_PP03 <> "" Then
'         If Combo2.Tag <> Combo2.Text And ProState = "1" Then
'            arrCaseNo = Split(Me.Lbl1(7).Caption, "-")
'            If Trim(m_PP04) <> Trim(Left(Combo2.Text, 6)) Then
'               If PUB_ChkPromoterReader(IIf(m_PP01 <> "", m_PP01, arrCaseNo(0)), IIf(m_PP03 <> "", m_PP03, m_CP10), "1", Trim(Left(Combo2.Text, 6))) = False And _
'                  PUB_ChkPromoterReader(IIf(m_PP01 <> "", m_PP01, arrCaseNo(0)), m_CP10, "1", Trim(Left(Combo2.Text, 6))) = False Then
'                  MsgBox "此人無核稿權限，請重新輸入！"
'                  Combo2.SetFocus
'                  Exit Function
'               End If
'            End If
'         End If
'      End If
   End If
End If
If Combo6.Enabled = True And Val(m_CP27) = 0 Then
   '若核判表有設定判發人時只可以修改但不可以空白
   If Trim(m_PP05) <> "" And Trim(Left(Combo6.Text, 6)) = "" Then
      MsgBox "判發人不可空白!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1 'Add By Sindy 2024/3/15
      Combo6.SetFocus
      Exit Function
   End If
   If Trim(Left(Combo6.Text, 6)) <> "" Then
      '增加檢查判發人是否離職
      If PUB_ChkEmpFlowExists(lbl1(3), EMP_送判) = True And _
         PUB_ChkEmpFlowExists(lbl1(3), EMP_判發) = False Then
         If ChkStaffST04(Trim(Left(Combo6.Text, 6))) = True Then
            SSTab1.Tab = 1
            Combo6.SetFocus
            Exit Function
         End If
      End If
      If UCase(Trim(Left(Combo1.Text, 6))) = UCase(Trim(Left(Combo6.Text, 6))) And _
         Not (UCase(m_PP05) = UCase(Trim(Left(Combo6.Text, 6)))) Then
         MsgBox "判發人不可與承辦人相同!!!", vbExclamation + vbOKOnly
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         Combo6.SetFocus
         Exit Function
      End If
      'Modify By Sindy 2018/10/1 不鎖權限
'      '當代理狀況時，檢查輸入的判發人是否有判發權限
'      '承辦人非程序人員時,才需檢查核判權限
'      'Modify By Sindy 2018/9/19 And m_PP03 <> "" ==>有設定核判表的案件性質才要檢查權限
'      If GetStaffDepartment(Trim(Left(Combo1.Text, 6))) <> "P22" And m_PP03 <> "" Then
'         If Combo6.Tag <> Combo6.Text And ProState = "1" Then
'            arrCaseNo = Split(Me.Lbl1(7).Caption, "-")
'            '只要非系統設定的人員均要檢查權限
'            If Trim(m_PP05) <> Trim(Left(Combo6.Text, 6)) Then
'               If PUB_ChkPromoterReader(IIf(m_PP01 <> "", m_PP01, arrCaseNo(0)), IIf(m_PP03 <> "", m_PP03, m_CP10), "2", Trim(Left(Combo6.Text, 6))) = False And _
'                  PUB_ChkPromoterReader(IIf(m_PP01 <> "", m_PP01, arrCaseNo(0)), m_CP10, "2", Trim(Left(Combo6.Text, 6))) = False Then
'                  For ii = 0 To Me.Combo6.ListCount - 1
'                      blnMatch = False
'                      If Trim(Left(Me.Combo6.List(ii), 6)) = Trim(Left(Me.Combo6.Text, 6)) Then
'                          Me.Combo6.ListIndex = ii
'                          blnMatch = True
'                          Exit For
'                      End If
'                  Next ii
'                  If blnMatch = False Then
'                     MsgBox "此人無判發權限，請重新輸入！"
'                     Combo6.SetFocus
'                     Exit Function
'                  End If
'               End If
'            End If
'         End If
'      End If
   End If
End If
'加入日期檢查
If Trim(txt1(2)) = "" And Trim(txt1(3)) & Trim(txt1(4)) & Trim(txt1(7)) & Trim(txt1(8)) <> "" And _
   txt1(2).Enabled = True Then
    MsgBox "有下列日期，齊備日不能空白！" & vbCrLf & "完稿日、會稿日、會稿完成日、發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(2).SetFocus
    Exit Function
End If
If Trim(txt1(3)) = "" And Trim(txt1(4)) & Trim(txt1(7)) & Trim(txt1(8)) <> "" And _
   txt1(3).Enabled = True Then
    MsgBox "有下列日期，完稿日不能空白！" & vbCrLf & "會稿日、會稿完成日、發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(3).SetFocus
    Exit Function
End If
If Trim(txt1(4)) = "" And Trim(txt1(7)) & Trim(txt1(8)) <> "" And _
   txt1(4).Enabled = True Then
    MsgBox "有下列日期，會稿日不能空白！" & vbCrLf & "會稿完成日、發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(4).SetFocus
    Exit Function
End If
If Trim(txt1(7)) = "" And Trim(txt1(8)) <> "" And txt1(7).Enabled = True Then
    MsgBox "有下列日期，會稿完成日不能空白！" & vbCrLf & "發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(7).SetFocus
    Exit Function
End If
'2018/4/20 END

Set rsA = Nothing
Set rsB = Nothing
TxtValidate = True
End Function

'Add By Sindy 2021/2/2 確定鍵,多案單筆歷程時,要更新瀏覽資料日期欄位值
Sub StrMenuOneRec_RecvSub(strRecvSub As String)
Dim ii As Integer, intCnt As Integer
Dim PicRs As New ADODB.Recordset
Dim arrID As Variant
   
   arrID = Split(strRecvSub, ",")
   For intCnt = 0 To UBound(arrID)
      For ii = 1 To Me.grd1.Rows - 1
         '依收文號更新畫面欄位值
         If Me.grd1.TextMatrix(ii, 23) = arrID(intCnt) Then
            strSql = "SELECT * from engineerprogress,caseprogress where ep02='" & arrID(intCnt) & "' and ep02=cp09"
            PicRs.CursorLocation = adUseClient
            PicRs.Open strSql, cnnConnection, adOpenStatic, adLockOptimistic
            If PicRs.RecordCount <> 0 Then
               '承辦期限
               If Val("" & PicRs.Fields("cp48")) > 0 Then
                  Me.grd1.TextMatrix(ii, 10) = ChangeWStringToTDateString(PicRs.Fields("cp48"))
               Else
                  Me.grd1.TextMatrix(ii, 10) = ""
               End If
               '齊備日
               If Val("" & PicRs.Fields("ep06")) > 0 Then
                  Me.grd1.TextMatrix(ii, 13) = ChangeWStringToTDateString(PicRs.Fields("ep06"))
               Else
                  Me.grd1.TextMatrix(ii, 13) = ""
               End If
               '完稿日
               If Val("" & PicRs.Fields("ep06")) > 0 Then
                  Me.grd1.TextMatrix(ii, 14) = ChangeWStringToTDateString(PicRs.Fields("ep09"))
               Else
                  Me.grd1.TextMatrix(ii, 14) = ""
               End If
               '指會日
               If Val("" & PicRs.Fields("EP28")) > 0 Then
                  Me.grd1.TextMatrix(ii, 15) = ChangeWStringToTDateString(PicRs.Fields("EP28"))
               Else
                  Me.grd1.TextMatrix(ii, 15) = ""
               End If
               '會稿日
               If Val("" & PicRs.Fields("EP07")) > 0 Then
                  Me.grd1.TextMatrix(ii, 16) = ChangeWStringToTDateString(PicRs.Fields("EP07"))
               Else
                  Me.grd1.TextMatrix(ii, 16) = ""
               End If
               '核稿人
               If "" & PicRs.Fields("EP04") <> "" Then
                  Me.grd1.TextMatrix(ii, 17) = GetPrjSalesNM(PicRs.Fields("EP04"))
               Else
                  Me.grd1.TextMatrix(ii, 17) = ""
               End If
               '會稿完成日
               If Val("" & PicRs.Fields("EP08")) > 0 Then
                  Me.grd1.TextMatrix(ii, 18) = ChangeWStringToTDateString(PicRs.Fields("EP08"))
               Else
                  Me.grd1.TextMatrix(ii, 18) = ""
               End If
               '發文日
               If Val("" & PicRs.Fields("cp27")) > 0 Then
                  Me.grd1.TextMatrix(ii, 19) = ChangeWStringToTDateString(PicRs.Fields("cp27"))
               Else
                  Me.grd1.TextMatrix(ii, 19) = ""
               End If
               '承辦備註
               Me.grd1.TextMatrix(ii, 21) = "" & PicRs.Fields("ep12")
               
               '修正日期欄位排序問題(小於100年的前面補空白)
               For intI = 10 To 21
                  If Len(grd1.TextMatrix(ii, intI)) = 8 Then
                    If Mid(grd1.TextMatrix(ii, intI), 3, 1) = "/" And Mid(grd1.TextMatrix(ii, intI), 6, 1) = "/" Then
                       grd1.TextMatrix(ii, intI) = " " & grd1.TextMatrix(ii, intI)
                    End If
                  End If
               Next
               
               ChgGrdColor ii
               PicRs.Close
               Exit For
            End If
            PicRs.Close
         End If
      Next ii
   Next intCnt
   Set PicRs = Nothing
End Sub

Sub StrMenuOneRec(Optional ByVal Strindex As Integer = 1)
Dim ii As Integer
   For ii = 1 To Me.grd1.Rows - 1
      '若目次相同, 收文號也相同
      If Me.grd1.TextMatrix(ii, 0) = Me.lbl1(0).Caption And Me.grd1.TextMatrix(ii, 23) = m_strCP09 Then
         '承辦期限
         Me.grd1.TextMatrix(ii, 10) = ChangeTStringToTDateString(Me.txt1(12).Text)
         '齊備日
         Me.grd1.TextMatrix(ii, 13) = ChangeTStringToTDateString(Me.txt1(2).Text)
         '完稿日
         Me.grd1.TextMatrix(ii, 14) = ChangeTStringToTDateString(Me.txt1(3).Text)
         '指會日
         Me.grd1.TextMatrix(ii, 15) = ChangeTStringToTDateString(Me.txt1(18).Text)
         '會稿日
         Me.grd1.TextMatrix(ii, 16) = ChangeTStringToTDateString(Me.txt1(4).Text)
         '核稿人 Add By Sindy 2018/4/25
         Me.grd1.TextMatrix(ii, 17) = IIf(Combo2.Text = "", "", IIf(InStr(Combo2, "==>") > 0, Trim(Mid(Combo2, 10)), Trim(Mid(Combo2, 6))))
         '會稿完成日 Add By Sindy 2018/4/25
         Me.grd1.TextMatrix(ii, 18) = ChangeTStringToTDateString(Me.txt1(7).Text)
         '發文日
         Me.grd1.TextMatrix(ii, 19) = ChangeTStringToTDateString(Me.txt1(8).Text)
         '承辦備註
         Me.grd1.TextMatrix(ii, 21) = Me.txtEP12.Text
         
         '修正日期欄位排序問題(小於100年的前面補空白)
         For intI = 10 To 21
            If Len(grd1.TextMatrix(ii, intI)) = 8 Then
              If Mid(grd1.TextMatrix(ii, intI), 3, 1) = "/" And Mid(grd1.TextMatrix(ii, intI), 6, 1) = "/" Then
                 grd1.TextMatrix(ii, intI) = " " & grd1.TextMatrix(ii, intI)
              End If
            End If
         Next
         
         ChgGrdColor ii
         Exit For
      End If
   Next ii
   
   SWPRow2 = Strindex
   grd1.row = Val(SWPRow2)
   grd1.col = 1
End Sub

' 控制只跟 DB 溝通一次
' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To TF_CP
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CP" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      '定義數字
      Select Case nIndex
         Case 5, 6, 7, 15, 16, 17, 18, 19, 25, 27, 33, 34, 46, 47, 48, 53, 54, 57, 66, 67, 69, 70, 73, 74, 75, 76, 77, 78, 79, 82, 84, 85, 97, 98, 100, 101, 103, 104, 108, 109, 111:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
End Sub

Private Sub ClearFieldList()
   Erase m_FieldList
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To TF_CP - 1
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   For nIndex = 0 To TF_CP - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

Private Function FormSave() As Boolean
Dim iMouse As Integer
   
   iMouse = Screen.MousePointer
   
On Error GoTo ErrHand

If m_Flow = "" Then cnnConnection.BeginTrans

cnnConnection.Execute "begin user_data.user_formname:='" & Me.Name & "';end;"
   
   '目次
   SeekTmpBk = Trim(lbl1(0).Caption)
   'Modify By Sindy 2024/5/23 調整檢查欄位有異動再儲存
'****************************************************************
'更新EP
'****************************************************************
   '有預設欄位值
   strSql = "EP04='" & Trim(Left("" & Combo2.Text, 6)) & "',EP40='" & Trim(Left("" & Combo6.Text, 6)) & "'" & _
            ",EP34='" & txt1(1) & "'"
   If txt1(2).Tag <> txt1(2).Text Then
      Pub_SaveLog strUserNum, "齊備日異動：" & DBDATE(lbl1(8)) & "==>" & DBDATE(txt1(2)) & " ", SystemNumber(lbl1(7).Caption, 1), SystemNumber(lbl1(7).Caption, 2), SystemNumber(lbl1(7).Caption, 3), SystemNumber(lbl1(7).Caption, 4), lbl1(3).Caption
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP06=" & IIf(ChangeTStringToWString(txt1(2)) = "", "NULL", ChangeTStringToWString(txt1(2)))
   End If
   If txtEP12.Tag <> txtEP12 Then
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP12='" & txtEP12 & "'"
   End If
   If Val(txt1(3).Tag) <> Val(txt1(3).Text) Then
      Pub_SaveLog strUserNum, "完稿日異動：" & DBDATE(Trim(txt1(3).Tag)) & "==>" & DBDATE(Trim(txt1(3).Text)) & " ", SystemNumber(lbl1(7).Caption, 1), SystemNumber(lbl1(7).Caption, 2), SystemNumber(lbl1(7).Caption, 3), SystemNumber(lbl1(7).Caption, 4), lbl1(3).Caption
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP09=" & IIf(ChangeTStringToWString(txt1(3)) = "", "NULL", ChangeTStringToWString(txt1(3)))
   End If
   If Val(txt1(4).Tag) <> Val(txt1(4).Text) Then
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP07=" & IIf(ChangeTStringToWString(txt1(4)) = "", "NULL", ChangeTStringToWString(txt1(4)))
   End If
   If Val(txt1(7).Tag) <> Val(txt1(7).Text) Then
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP08=" & IIf(ChangeTStringToWString(txt1(7)) = "", "NULL", ChangeTStringToWString(txt1(7)))
   End If
   '指定會稿日異動
   If Trim(txt1(18).Tag) <> Trim(txt1(18).Text) Then
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP28=" & CNULL(ChangeTStringToWString(txt1(18)))
   End If
   If strSql <> "" Then
      strSql = "Update EngineerProgress Set " & strSql & " Where EP02='" & lbl1(3).Caption & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   '2024/5/23 END
'****************************************************************
   '加入 核稿人異動時，紀錄
   If Combo2.Tag <> Combo2.Text Then
      Pub_SaveLog strUserNum, "核稿人異動：" & Combo2.Tag & "==>" & Combo2.Text & " ", SystemNumber(lbl1(7).Caption, 1), SystemNumber(lbl1(7).Caption, 2), SystemNumber(lbl1(7).Caption, 3), SystemNumber(lbl1(7).Caption, 4), lbl1(3).Caption
   End If
   '加入 判發人異動時，紀錄
   If Combo6.Tag <> Combo6.Text Then
      Pub_SaveLog strUserNum, "判發人異動：" & Combo6.Tag & "==>" & Combo6.Text & " ", SystemNumber(lbl1(7).Caption, 1), SystemNumber(lbl1(7).Caption, 2), SystemNumber(lbl1(7).Caption, 3), SystemNumber(lbl1(7).Caption, 4), lbl1(3).Caption
   End If
   
   If Mid(lbl1(3).Caption, 1, 1) = "C" Then
      '發文日
      If Trim(txt1(8).Tag) <> Trim(txt1(8).Text) Then
         SetFieldNewData "CP27", IIf(ChangeTStringToWString(txt1(8)) = "", "", ChangeTStringToWString(txt1(8)))
      End If
   End If
   '承辦期限
   If Trim(txt1(12).Tag) <> Trim(txt1(12).Text) Then
      SetFieldNewData "CP48", IIf(txt1(12) <> "", ChangeTStringToWString(txt1(12)), "")
   End If
   
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   strSql = " UPDATE CASEPROGRESS SET "
   bFirst = True
   bDifference = False
      
   For nIndex = 0 To TF_CP - 1
      strTmp = Empty
      If nIndex < 64 Or nIndex > 69 Then
         If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
            If m_FieldList(nIndex).fiType = 0 Then
               If m_FieldList(nIndex).fiNewData = Empty Then
                  strTmp = m_FieldList(nIndex).fiName & " = NULL "
               Else
                  strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
               End If
            Else
               If m_FieldList(nIndex).fiNewData = Empty Then
                  strTmp = m_FieldList(nIndex).fiName & " = NULL "
               Else
                  strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
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
      End If
   Next nIndex
                  
   strSql = strSql & " " & _
      "WHERE CP09 = '" & Me.lbl1(3).Caption & "' "
     
   If bDifference = True Then
      cnnConnection.Execute strSql
   End If
   
   cnnConnection.Execute "begin user_data.user_formname:=Null;end;"
   
   If m_Flow = "" Then cnnConnection.CommitTrans
      
   FormSave = True
   Exit Function
   
ErrHand:
   cnnConnection.Execute "begin user_data.user_formname:=Null;end;"
   If m_Flow = "" Then cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

'批次發Mail
'Modify By Sindy 2018/6/20
'Private Sub BatctMail()
Public Sub BatctMail()
'2018/6/20 END
   Dim i As Integer
   For i = 1 To UBound(skMail)
        PUB_SendMail skMail(i).fiSender, skMail(i).fiReceiver, skMail(i).fiRecriverNo, skMail(i).fiSubject, skMail(i).fiContent
   Next i
   ReDim skMail(0) As SeekMails
   'Trigger 也會產生待發郵件
   PUB_SendMailCache
End Sub

'更新mdb暫存資料及第一畫面的Grid內容
Private Sub UpdEngMdb()

On Error GoTo ErrHand
   
   'R110013.齊備日
   'R110014.完稿日
   'R110015.會稿日
   'R110010.承辦期限
   'R110018.發文日
   'R110020.承辦備註
   strSql = "UPDATE R090614 SET " & _
      "R110013='" & IIf(txt1(2) = "", "", Right(" " & ChangeTStringToTDateString(txt1(2)), 9)) & "'," & _
      "R110014='" & IIf(txt1(3) = "", "", Right(" " & ChangeTStringToTDateString(txt1(3)), 9)) & "'," & _
      "R110015='" & IIf(txt1(4) = "", "", Right(" " & ChangeTStringToTDateString(txt1(4)), 9)) & "'," & _
      "R110017='" & IIf(txt1(7) = "", "", Right(" " & ChangeTStringToTDateString(txt1(7)), 9)) & "'," & _
      "R110016='" & IIf(Combo2.Text = "", "", IIf(InStr(Combo2, "==>") > 0, Trim(Mid(Combo2, 10)), Trim(Mid(Combo2, 6)))) & "'," & _
      "R110010='" & IIf(txt1(12) = "", "", Right(" " & ChangeTStringToTDateString(txt1(12)), 9)) & "'," & _
      "R110018='" & IIf(txt1(8) = "", "", Right(" " & ChangeTStringToTDateString(txt1(8)), 9)) & "'," & _
      "R110020='" & txtEP12 & "' " & _
      " WHERE ID='" & strUserNum & "' AND R110022='" & lbl1(3).Caption & "' "
   adoEng.Execute strSql, intI
   
   m_blnClkSure = True
   For i = 1 To grd1.Rows - 1
      grd1.row = i
      grd1.col = 0
      '若目次相同, 收文號也相同
      If grd1.Text = SeekTmpBk And Me.grd1.TextMatrix(i, 23) = m_strCP09 Then
         MouseClick_1 (i)
         StrMenuOneRec SWPRow2
         Exit For
      End If
   Next i
   m_blnClkSure = False
      
ErrHand:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'設定承辦人選單
Private Sub SetEngineer()
   strSql = "SELECT Distinct (R110001&' '&'(' & R110025&')') FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Trim(strUserNum) & "' ORDER BY (R110001&' '&'(' & R110025&')') "
   CheckOC
   i = 0
   Combo1.Clear
   Combo1_String = ""
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, adoEng, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
         Do While .EOF = False
           Combo1.AddItem "" & .Fields(0), i
           i = i + 1
           If Combo1_String = "" Then
              Combo1_String = "'" & Trim(Left("" & .Fields(0), 6)) & "'"
           Else
              Combo1_String = Combo1_String + ",'" & Trim(Left("" & .Fields(0), 6)) & "'"
           End If
           .MoveNext
         Loop
         Combo1.Text = Combo1.List(0)
       End If
   End With
End Sub

'Add By Sindy 2021/9/27
Private Sub txtEP12_GotFocus()
   txtEP12.SelStart = 0
   txtEP12.SelLength = Len(txtEP12)
End Sub
'承辦備註
Private Sub txtEP12_Validate(Cancel As Boolean)
   '若欄位無資料則不檢查
   If Me.txtEP12.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtEP12, txtEP12.MaxLength) Then
       txtEP12.SetFocus
       txtEP12_GotFocus
       Cancel = True
       Exit Sub
   End If
End Sub
Private Sub txtCP64_GotFocus()
   txtCP64.SelStart = 0
   txtCP64.SelLength = Len(txtCP64)
End Sub
'進度備註
Private Sub txtCP64_Validate(Cancel As Boolean)
   '若欄位無資料則不檢查
   If Me.txtCP64.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCP64, txtCP64.MaxLength) Then
       txtCP64.SetFocus
       txtCP64_GotFocus
       Cancel = True
       Exit Sub
   End If
End Sub
