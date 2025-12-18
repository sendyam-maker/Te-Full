VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010605_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "證書號數輸入"
   ClientHeight    =   6600
   ClientLeft      =   456
   ClientTop       =   948
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6590.496
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   8593
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   401
      Index           =   2
      Left            =   7515
      TabIndex        =   28
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   401
      Index           =   0
      Left            =   5460
      TabIndex        =   26
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   401
      Index           =   1
      Left            =   6300
      TabIndex        =   27
      Top             =   15
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4560
      TabIndex        =   33
      Top             =   451
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   29
      Top             =   451
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   30
      Top             =   451
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   31
      Top             =   451
      Width           =   255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   32
      Top             =   451
      Width           =   375
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4530
      Left            =   30
      TabIndex        =   43
      Top             =   2010
      Width           =   8505
      _ExtentX        =   15007
      _ExtentY        =   7980
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "證書資料"
      TabPicture(0)   =   "frm06010605_2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPA14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label25(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label25(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label24"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label23"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label31"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label30"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label29"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label28"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label27"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label26"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label21"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label20"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label19"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label18"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label17"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text33(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text33(5)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text33(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text33(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text33(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text33(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text33(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text10(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text10(1)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text10(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtPA14"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text9"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Combo2(1)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Combo2(0)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text8(1)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text8(0)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text7"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text6"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).ControlCount=   36
      TabCaption(1)   =   "發明人資料"
      TabPicture(1)   =   "frm06010605_2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "cmdUpdRow"
      Tab(1).Control(2)=   "cmdAddRow"
      Tab(1).Control(3)=   "cmdDelRow"
      Tab(1).Control(4)=   "Combo4"
      Tab(1).Control(5)=   "txtIN11"
      Tab(1).Control(6)=   "GRD1"
      Tab(1).Control(7)=   "GRDtmp"
      Tab(1).Control(8)=   "txtInvField(0)"
      Tab(1).Control(9)=   "txtInvField(1)"
      Tab(1).Control(10)=   "txtInvField(2)"
      Tab(1).Control(11)=   "Lb_IN11N"
      Tab(1).Control(12)=   "Lb_Inv(0)"
      Tab(1).Control(13)=   "Lb_Inv(1)"
      Tab(1).Control(14)=   "Lb_Inv(2)"
      Tab(1).Control(15)=   "Lb_Inv(3)"
      Tab(1).Control(16)=   "Lb_IN11"
      Tab(1).ControlCount=   17
      Begin VB.Frame Frame1 
         Caption         =   "移動順序:"
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   -69990
         TabIndex        =   69
         Top             =   1410
         Width           =   2025
         Begin VB.CommandButton cmdUp 
            Caption         =   "▲"
            Height          =   255
            Left            =   960
            TabIndex        =   71
            Top             =   90
            Width           =   375
         End
         Begin VB.CommandButton cmdDown 
            Caption         =   "▼"
            Height          =   255
            Left            =   1410
            TabIndex        =   70
            Top             =   90
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdUpdRow 
         Caption         =   "修改"
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72180
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdAddRow 
         Caption         =   "加入"
         Height          =   285
         Left            =   -73845
         TabIndex        =   23
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdDelRow 
         Caption         =   "刪除"
         Height          =   285
         Left            =   -73020
         TabIndex        =   24
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         ItemData        =   "frm06010605_2.frx":0038
         Left            =   -74110
         List            =   "frm06010605_2.frx":003A
         Style           =   2  '單純下拉式
         TabIndex        =   18
         Top             =   350
         Width           =   5535
      End
      Begin VB.TextBox txtIN11 
         Height          =   270
         Left            =   -68040
         MaxLength       =   3
         TabIndex        =   22
         Top             =   708
         Width           =   400
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   0
         Top             =   365
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   3885
         MaxLength       =   7
         TabIndex        =   1
         Top             =   365
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Index           =   0
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   3
         Top             =   635
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Index           =   1
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   4
         Top             =   635
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   0
         Left            =   4560
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   905
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   1
         Left            =   4560
         Style           =   2  '單純下拉式
         TabIndex        =   10
         Top             =   1730
         Width           =   2295
      End
      Begin VB.TextBox Text9 
         Height          =   264
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   17
         Top             =   3720
         Width           =   5652
      End
      Begin VB.TextBox txtPA14 
         Height          =   270
         Left            =   6210
         MaxLength       =   7
         TabIndex        =   2
         Top             =   365
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   2715
         Left            =   -74940
         TabIndex        =   68
         Top             =   1740
         Width           =   8355
         _ExtentX        =   14732
         _ExtentY        =   4784
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "V|發明人編號|中文名稱|英文名稱|日文名稱|國籍|申請人1"
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
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRDtmp 
         Height          =   825
         Left            =   -74910
         TabIndex        =   72
         Top             =   1500
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1820
         _ExtentY        =   1461
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "V|發明人編號|中文名稱|英文名稱|日文名稱|國籍|申請人1"
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
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.TextBox Text10 
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   14
         Top             =   2835
         Width           =   5655
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "9975;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text10 
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   3135
         Width           =   5655
         VariousPropertyBits=   671105051
         MaxLength       =   250
         Size            =   "9975;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text10 
         Height          =   300
         Index           =   2
         Left            =   1440
         TabIndex        =   16
         Top             =   3420
         Width           =   5655
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "9975;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtInvField 
         Height          =   285
         Index           =   0
         Left            =   -74115
         TabIndex        =   19
         Top             =   600
         Width           =   5535
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "9763;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtInvField 
         Height          =   285
         Index           =   1
         Left            =   -74115
         TabIndex        =   20
         Top             =   870
         Width           =   5535
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "9763;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtInvField 
         Height          =   285
         Index           =   2
         Left            =   -74115
         TabIndex        =   21
         Top             =   1140
         Width           =   5535
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "9763;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   5
         Top             =   915
         Width           =   2775
         VariousPropertyBits=   671105051
         Size            =   "4895;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   1185
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   8
         Top             =   1455
         Width           =   5655
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9975;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   9
         Top             =   1740
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   11
         Top             =   2010
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   12
         Top             =   2295
         Width           =   5655
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9975;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   13
         Top             =   2565
         Width           =   5655
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "9975;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Lb_IN11N 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "Lb_IN11N"
         Height          =   180
         Left            =   -68500
         TabIndex        =   67
         Top             =   1058
         Width           =   975
      End
      Begin VB.Label Lb_Inv 
         AutoSize        =   -1  'True
         Caption         =   "發明人"
         Height          =   180
         Index           =   0
         Left            =   -74730
         TabIndex        =   66
         Top             =   390
         Width           =   540
      End
      Begin VB.Label Lb_Inv 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   1
         Left            =   -74470
         TabIndex        =   65
         Top             =   638
         Width           =   345
      End
      Begin VB.Label Lb_Inv 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   2
         Left            =   -74470
         TabIndex        =   64
         Top             =   908
         Width           =   345
      End
      Begin VB.Label Lb_Inv 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   3
         Left            =   -74470
         TabIndex        =   63
         Top             =   1168
         Width           =   345
      End
      Begin VB.Label Lb_IN11 
         AutoSize        =   -1  'True
         Caption         =   "國籍:"
         Height          =   180
         Left            =   -68510
         TabIndex        =   62
         Top             =   708
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "專利號數:"
         Height          =   180
         Left            =   150
         TabIndex        =   61
         Top             =   410
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "發證日期:"
         Height          =   180
         Left            =   3030
         TabIndex        =   60
         Top             =   410
         Width           =   765
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "專用期限: "
         Height          =   180
         Left            =   150
         TabIndex        =   59
         Top             =   680
         Width           =   810
      End
      Begin VB.Label Label20 
         Caption         =   "~"
         Height          =   255
         Left            =   2760
         TabIndex        =   58
         Top             =   643
         Width           =   135
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "下次繳費日:"
         Height          =   180
         Left            =   4560
         TabIndex        =   57
         Top             =   680
         Width           =   945
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(中):"
         Height          =   180
         Left            =   150
         TabIndex        =   56
         Top             =   965
         Width           =   975
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(英):"
         Height          =   180
         Left            =   150
         TabIndex        =   55
         Top             =   1235
         Width           =   975
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(日):"
         Height          =   180
         Left            =   150
         TabIndex        =   54
         Top             =   1505
         Width           =   975
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(中):"
         Height          =   180
         Left            =   150
         TabIndex        =   53
         Top             =   1790
         Width           =   975
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(英):"
         Height          =   180
         Left            =   150
         TabIndex        =   52
         Top             =   2060
         Width           =   975
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(日):"
         Height          =   180
         Left            =   150
         TabIndex        =   51
         Top             =   2345
         Width           =   975
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(中):"
         Height          =   180
         Left            =   150
         TabIndex        =   50
         Top             =   2885
         Width           =   1065
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(英):"
         Height          =   180
         Left            =   150
         TabIndex        =   49
         Top             =   3180
         Width           =   1065
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(外):"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   48
         Top             =   3465
         Width           =   1065
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   7
         Left            =   5610
         TabIndex        =   47
         Top             =   675
         Width           =   1170
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2064;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號:"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   46
         Top             =   3765
         Width           =   1125
      End
      Begin VB.Label lblPA14 
         AutoSize        =   -1  'True
         Caption         =   "公告日:"
         Height          =   180
         Left            =   5535
         TabIndex        =   45
         Top             =   410
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日):"
         Height          =   180
         Left            =   150
         TabIndex        =   44
         Top             =   2615
         Width           =   1245
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Left            =   240
      TabIndex        =   79
      Top             =   1380
      Width           =   675
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Left            =   3480
      TabIndex        =   78
      Top             =   1380
      Width           =   675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Left            =   240
      TabIndex        =   77
      Top             =   1695
      Width           =   675
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   4560
      TabIndex        =   76
      Top             =   1065
      Width           =   2220
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3916;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   75
      Top             =   1380
      Width           =   2220
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3916;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   5
      Left            =   4560
      TabIndex        =   74
      Top             =   1380
      Width           =   2220
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3916;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   73
      Top             =   1695
      Width           =   2220
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3916;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   42
      Top             =   1065
      Width           =   2220
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3916;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   41
      Top             =   750
      Width           =   2220
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3916;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   40
      Top             =   451
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Left            =   3480
      TabIndex        =   39
      Top             =   1065
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Left            =   240
      TabIndex        =   38
      Top             =   1065
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   3480
      TabIndex        =   37
      Top             =   750
      Width           =   945
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   36
      Top             =   750
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3480
      TabIndex        =   35
      Top             =   451
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   34
      Top             =   750
      Width           =   2220
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3916;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm06010605_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/22 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim strReceiveNo As String, strSales As String
'Modify by Morgan 2006/10/20 改動態
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim intWhere As Integer, bolFormLoad As Boolean
'Add By Cheng 2001/12/31
Dim m_bln_keyinValidate As Boolean

'Add by Morgan 2004/8/4
Dim m_bolNew As Boolean '是否用新法
Dim m_str421CP09 As String '技術報告總收文號
Dim m_str421EP06 As String '技術報告文件齊備日
Dim m_str421CP48 As String '技術報告承辦期限
Dim DATE1 As String, DATE2 As String '專用期起訖
Dim m_intLastYear As Integer  '最後繳費年度
Dim m_strNextFeeDate As String  '下次繳費日本所期限
Dim m_strNextDueDate As String  '下次繳費日法定期限
Dim m_strAgreeOnDate As String  '下次繳費日約定期限 Add By Sindy 2021/8/17

Dim m_928Upd As Boolean '是否更新重新委任准駁
Dim m_928CP09 As String '重新委任收文號
Dim m_bolIs117 As Boolean '是否為積體電路佈局案件
Dim strName As String '20140306ADD By eric
Dim m_otxt As Object '20140306ADD By eric    共用物件

'Add by Lydia 2014/10/22 發明人輸入比對兼自動代入(模糊比對)
' 宣告發明人
Private Type INVENTOR
   iN01 As String
   iN02 As String
   iN04 As String
   IN05 As String
   IN06 As String
End Type

Dim m_InventorList() As INVENTOR
Dim m_InventorListCount As Integer
Dim m_LetterLanguage As String 'add by sonia 2014/4/22
Dim pPrevRow As Integer 'Add By Sindy 2014/11/11
Dim nResponse 'Modified by Lydia 2015/01/06
'Added by Morgan 2015/5/20
Dim m_bolDualApply As Boolean '是否一案兩請
Dim m_stUPA(4) As String '一案兩請新型案號
Dim m_bolDualApp1 As Boolean, m_bolDualApp2 As Boolean '一案兩請 1:可自動閉卷(有擇一發文且放棄新型)但有未完成事項 2:不符合自動閉卷
'Added by Morgan 2023/1/16 電子公文
Public m_DocWord As String
Public m_DocNo As String
'end 2023/1/16

'Add by Lydia 2014/10/22 發明人輸入比對兼自動代入(模糊比對)
' 增加發明人
Private Sub AddInventor(ByVal strInventor As String, Optional ByVal mIN02 As String, Optional ByVal mIN04 As String, Optional ByVal mIN05 As String, Optional ByVal mIN06 As String)

   Dim strIN01 As String
   
    ' 字串補滿八碼或只取八碼
    If Len(strInventor) > 8 Then
       strIN01 = Mid(strInventor, 1, 8)
    Else
       strIN01 = strInventor & String(8 - Len(strInventor), "0")
    End If
    
     m_InventorList(m_InventorListCount).iN01 = strIN01 '客戶編號(8碼)
     m_InventorList(m_InventorListCount).iN02 = mIN02  '發明人代號
     m_InventorList(m_InventorListCount).iN04 = mIN04  '(發明人)中文名稱
     m_InventorList(m_InventorListCount).IN05 = mIN05  '(發明人)英文名稱
     m_InventorList(m_InventorListCount).IN06 = mIN06  '(發明人)日文名稱
    
     m_InventorListCount = m_InventorListCount + 1
End Sub

'Add By Sindy 2014/11/11
Private Sub cmdAddRow_Click()
Dim bolChk As Boolean
Dim ii As Integer
Dim Cancel As Boolean
Dim rsTmp  As New ADODB.Recordset
Dim strNo As String
   
   '檢查發明人
   strExc(1) = Replace(Right(Combo4.Text, 11), ")", "")
   If strExc(1) = "" Then
      If txtInvField(0) = "" And txtInvField(1) = "" And txtInvField(2) = "" Then
         Exit Sub
      Else
         '判斷國籍是否有輸入
         If txtIN11.Visible = True Then
            If txtIN11 = "" Then
                MsgBox "請輸入國藉！", vbExclamation
                SSTab1.Tab = 1
                txtIN11.SetFocus
                Exit Sub
            Else
                Cancel = False
                txtIN11_Validate Cancel
                If Cancel = True Then
                  SSTab1.Tab = 1
                  txtIN11.SetFocus
                  TextInverse txtIN11
                  Exit Sub
                End If
            End If
         End If
         '判斷客戶發明人檔是否有重覆資料:發明人會有造字無法存檔時會加空白,所以改在語法內trim
         strNo = pa(26) & String(8 - Len(Left(pa(26), 8)), "0")
         strSql = "Select * From Inventor Where IN01=" & CNULL(strNo) & " and (rtrim(IN04)=rtrim('" + ChgSQL(txtInvField(0)) & "')" & _
                  " OR upper(rtrim(IN05))=rtrim('" & ChgSQL(UCase(txtInvField(1))) & "') OR rtrim(IN06)=rtrim('" & ChgSQL(txtInvField(2)) & "'))"
         Set rsTmp = ClsPDReadRst(strSql)
         If Not rsTmp.EOF Then
            If Trim(txtInvField(0)) = Trim("" & rsTmp.Fields("IN04")) Then
               If MsgBox("發明人名稱中文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  rsTmp.Close
                  txtInvField(0).SetFocus
                  TextInverse txtInvField(0)
                  Exit Sub
               End If
            End If
            If Trim(UCase(txtInvField(1))) = UCase(Trim("" & rsTmp.Fields("IN05"))) Then
               If MsgBox("發明人名稱英文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  rsTmp.Close
                  txtInvField(1).SetFocus
                  TextInverse txtInvField(1)
                  Exit Sub
               End If
            End If
            If Trim(txtInvField(2)) = Trim("" & rsTmp.Fields("IN06")) Then
               If MsgBox("發明人名稱日文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  rsTmp.Close
                  txtInvField(2).SetFocus
                  TextInverse txtInvField(2)
                  Exit Sub
               End If
            End If
         End If
         rsTmp.Close
         
         If m_LetterLanguage = "3" And txtInvField(2) = "" Then
            If MsgBox("定稿語文為日文, 發明人的日文空白, 是否要輸入日文名稱 ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
               txtInvField(2).SetFocus
               TextInverse txtInvField(2)
               Exit Sub
            End If
         End If
         
         If Trim(txtInvField(0)) <> "" Then
            bolChk = True
            For ii = 1 To GRD1.Rows - 1
               If GRD1.TextMatrix(ii, 2) = Trim(txtInvField(0)) Then
                  bolChk = False
                  Exit For
               End If
            Next ii
            If Not bolChk Then
               MsgBox "發明人中文名稱不可重覆 !", vbCritical
               Combo4.SetFocus
               Exit Sub
            End If
         End If
         
         If Trim(txtInvField(1)) <> "" Then
            bolChk = True
            For ii = 1 To GRD1.Rows - 1
               If UCase(GRD1.TextMatrix(ii, 3)) = Trim(UCase(txtInvField(1))) Then
                  bolChk = False
                  Exit For
               End If
            Next ii
            If Not bolChk Then
               MsgBox "發明人英文名稱不可重覆 !", vbCritical
               Combo4.SetFocus
               Exit Sub
            End If
         End If
         
         If Trim(txtInvField(2)) <> "" Then
            bolChk = True
            For ii = 1 To GRD1.Rows - 1
               If GRD1.TextMatrix(ii, 4) = Trim(txtInvField(2)) Then
                  bolChk = False
                  Exit For
               End If
            Next ii
            If Not bolChk Then
               MsgBox "發明人日文名稱不可重覆 !", vbCritical
               Combo4.SetFocus
               Exit Sub
            End If
         End If
      End If
   Else
      bolChk = True
      For ii = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(ii, 1) = strExc(1) Then
            bolChk = False
            Exit For
         End If
      Next ii
      If Not bolChk Then
         MsgBox "發明人不可重覆 !", vbCritical
         Combo4.SetFocus
         Exit Sub
      End If
      If m_LetterLanguage = "3" And txtInvField(2) = "" Then
         If MsgBox("定稿語文為日文, 發明人的日文空白, 若需要日文請選 是 並自行至客戶發明人資料維護補輸日文, 不需要日文請選 否 !", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'            txtInvField(2).SetFocus
'            TextInverse txtInvField(2)
            Exit Sub
         End If
      End If
   End If
   If Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) <> "" Or Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 6) <> "" Then
      GRD1.AddItem ""
   End If
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) = strExc(1)
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 2) = txtInvField(0)
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 3) = txtInvField(1)
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 4) = txtInvField(2)
   If strExc(1) = "" Then
      Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 5) = txtIN11
      Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 6) = strNo '申請人1_ID
   End If
   cmdAddRow.Tag = "I" '記錄有異動資料
   'Call cmdUpdRow_Click 'Add By Sindy 2015/3/5
   '清空欄位
   Combo4.ListIndex = 0
   For Each m_otxt In txtInvField
      m_otxt.Text = ""
      m_otxt.Tag = "" 'Add By Sindy 2015/3/5
   Next
   txtIN11.Text = "" 'Add By Sindy 2015/12/4
End Sub

'Add By Sindy 2014/11/11
Private Sub cmdDelRow_Click()
   If pPrevRow = 1 And GRD1.Rows = 2 Then
      GRD1.TextMatrix(pPrevRow, 0) = ""
      GRD1.TextMatrix(pPrevRow, 1) = ""
      GRD1.TextMatrix(pPrevRow, 2) = ""
      GRD1.TextMatrix(pPrevRow, 3) = ""
      GRD1.TextMatrix(pPrevRow, 4) = ""
      GRD1.TextMatrix(pPrevRow, 5) = ""
      GRD1.TextMatrix(pPrevRow, 6) = ""
   Else
      If pPrevRow > 0 Then
         Call GRD1.RemoveItem(pPrevRow)
      Else
         Exit Sub
      End If
   End If
   pPrevRow = pPrevRow - 1
   cmdDelRow.Tag = "D" '記錄有異動資料
   '清空欄位
   Combo4.ListIndex = 0
   For Each m_otxt In txtInvField
      m_otxt.Text = ""
      m_otxt.Tag = "" 'Add By Sindy 2015/3/5
   Next
   txtIN11.Text = "" 'Add By Sindy 2015/12/4
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim nFrm As Form 'Added by Lydia 2019/12/03

   Select Case Index
      Case 0
        'Add By Cheng 2002/11/04
        If Me.Text8(0).Text = "" Then
            MsgBox "請輸入專用期起日!!!", vbExclamation + vbOKOnly
            Me.Text8(0).SetFocus
            Text8_GotFocus 0
            Exit Sub
        End If
        If Me.Text8(1).Text = "" Then
            MsgBox "請輸入專用期止日!!!", vbExclamation + vbOKOnly
            Me.Text8(1).SetFocus
            Text8_GotFocus 1
            Exit Sub
        End If
        
         'Add By Cheng 2001/12/31
         '檢查專案截止日
         Text8_Validate 1, False
         If Not m_bln_keyinValidate Then Exit Sub
         
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
                  
         'Added by Morgan 2015/5/20
         '一案兩請發明案發證,新型案自動閉卷
         m_bolDualApply = False
         m_bolDualApp1 = False: m_bolDualApp2 = False 'Added by Morgan 2019/7/18
         If pa(8) = "1" Then
            If PUB_IsDualApply(pa, m_stUPA) Then
               '新型未閉卷
               strExc(0) = "SELECT pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo FROM PATENT WHERE PA01='" & m_stUPA(1) & "' AND PA02='" & m_stUPA(2) & "' AND PA03='" & m_stUPA(3) & "' AND PA04='" & m_stUPA(4) & "' AND PA57 IS NULL"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  'Modified by Morgan 2016/6/14 有收文通知擇一且有選擇放棄新型才要閉卷--敏莉
                  If pa(60) = "Y" Then
                     'Modified by Morgan 2017/5/12
                     '有擇一申復發文也可閉卷--David Ex.FCP-49159,FCP-049365
                     'If PUB_ChkCPExist(pa(), "1232") = True Then
                     'Modified by Morgan 2019/7/18 改判斷有擇一發文(因可能來審查意見函) FCP-50781,FCP-50782
                     'If PUB_ChkCPExist(pa(), "1232") = True Or PUB_ChkCPExist(pa(), "239", 2) = True Then
                     If PUB_ChkCPExist(pa(), "239", 2) = True Then
                     'end 2017/5/12
                        
                        'Added by Morgan 2017/1/23 新型案閉卷檢查--何淑華
                        If CloseCheck(m_stUPA(1), m_stUPA(2), m_stUPA(3), m_stUPA(4)) = False Then
                           MsgBox "一案兩請新型案(" & RsTemp("CaseNo") & ")有未完成程序,請處理完再上閉卷", vbInformation
                           m_bolDualApp1 = True
                        Else
                        'end 2017/1/23
                        
                           MsgBox "本案與新型案(" & RsTemp("CaseNo") & ")為一案兩請，新型案將自動上閉卷！(若該新型案尚有年費期限也將會自動不續辦)", vbInformation
                           m_bolDualApply = True
                           
                        End If 'Added by Morgan 2017/1/23
                     End If
                  End If
                  'end 2016/6/14
                  
                  'Added by Morgan 2019/7/18
                  If Not (m_bolDualApp1 Or m_bolDualApply) Then
                     MsgBox "請確認是否閉卷新型(" & RsTemp("CaseNo") & ")", vbInformation
                     m_bolDualApp2 = True
                  End If
                  'end 2019/7/18
               End If
            End If
         End If
         'end 2015/5/20
         
         'Add by Sindy 2021/11/22 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If

         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Added by Lydia 2021/01/20 檢查"核對已准專利"是否有請款單號，若有，請彈訊息
         'Modified by Morgan 2021/12/14 請款單號應為CP60非CP64
         strExc(0) = "select cp09,cp60 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp159=0 and cp10='926' and cp60 is not null order by cp05"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
             MsgBox "本案二核已請款，請注意!", vbInformation
         End If
         'end 2021/01/20
         
         'Added by Lydia 2019/12/03 一併產生公告通知函定稿(E化及非E化)電子檔,放至卷宗區公告公報1228,副檔名:.CUS(通知函)
         '檢查表單是否已開啟，若是，則關閉
         For Each nFrm In Forms
            If StrComp(nFrm.Name, "frm060302", vbTextCompare) = 0 Then
               Unload frm060302
               Exit For
            End If
         Next
         frm06010605_2.Hide
         frm060302.m_KeyCP01 = pa(1)
         frm060302.m_KeyCP02 = pa(2)
         frm060302.m_KeyCP03 = pa(3)
         frm060302.m_KeyCP04 = pa(4)
         frm060302.m_KeyDate = Label2(1)
         frm060302.Show
         Call frm060302.cmdok_Click(0)
         Unload frm060302
         'end 2019/12/03
         
         'Mark by Lydia 2019/03/19 併入承辦單備註FcpEMPbill
         'ShowPrompt 'Added by Morgan 2018/5/28
         
         'Modified by Morgan 2023/1/16
         'frm06010605_1.Show
         If m_DocNo <> "" Then
            Unload frm06010605_1
            frm060119.GoNext
         Else
            frm06010605_1.Show
         End If
         'end 2023/1/16
         Unload Me
      Case 1
         frm06010605_1.Show
         Unload Me
      Case 2
         Unload frm06010605_1
         Unload Me
   End Select
End Sub

'Added by Morgan 2017/1/23 新型案閉卷檢查
Private Function CloseCheck(pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String) As Boolean
   Dim stSQL As String, iR As Integer
   Dim rsQuery As ADODB.Recordset
   
   '若"有"(A)"未請款程序" (B)"已收文未發文"則新型案不可自動上閉卷
   stSQL = "select * from caseprogress where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and ((cp27>0 and cp16>0 and cp20||cp57||cp60 is null) or cp27 is null) and cp57 is null"
   iR = 1
   Set rsQuery = ClsLawReadRstMsg(iR, stSQL)
   If iR = 0 Then
      CloseCheck = True
   End If
   Set rsQuery = Nothing
End Function

Private Function FormSave() As Boolean
Dim strTmp(1 To 3) As String
Dim stCP12 As String, stCP13 As String, stPA14 As String, stCP09 As String
Dim strCP20 As String, strCP16 As String
Dim strMemo605 As String   '2011/10/12 add by sonia
'20140306START ADD By eric
Dim i As Integer
Dim m_PA30 As String
Dim stUpdates As String
'20140306END
Dim ii As Integer
'Added by Lydia 2019/03/04
Dim m_strMemo As String '備註
Dim bolSpecMsg As Boolean  '是否彈訊息
Dim strCP14 As String, strCP48 As String  'Added by Lydia 2019/05/31 預設承辦人和承辦期限
Dim bolCaseIs959 As Boolean 'Added by Lydia 2021/08/16 外專-藥品專利連結：是否為藥品專利連結告代之案件

   m_928Upd = PUB_928Check(pa, m_928CP09) 'Add by Morgan 2007/7/18
   
   'Added by Lydia 2019/03/04 通知函承辦單備註設定(FCPEmpBill)
   'Modified by Lydia 2019/04/17 補足客戶編號
   'If PUB_GetFcpEMPBillSpec(pa(1) & pa(2) & pa(3) & pa(4), "03", pa(75), pa(26), bolSpecMsg, m_strMemo) = True Then
   strExc(1) = ChangeCustomerL(pa(26)) & IIf(pa(27) <> "", "," & ChangeCustomerL(pa(27)), "") & _
                    IIf(pa(28) <> "", "," & ChangeCustomerL(pa(28)), "") & _
                    IIf(pa(29) <> "", "," & ChangeCustomerL(pa(29)), "") & _
                    IIf(pa(30) <> "", "," & ChangeCustomerL(pa(30)), "")
   If PUB_GetFcpEMPBillSpec(pa(1) & pa(2) & pa(3) & pa(4), "03", ChangeCustomerL(pa(75)), strExc(1), bolSpecMsg, m_strMemo) = True Then
        If bolSpecMsg = True Then '彈訊息+列印在承辦單
            MsgBox m_strMemo, vbInformation, "承辦單備註"
        End If
   End If
   'end 2019/03/04
   
   'Added by Lydia 2021/08/16 外專-藥品專利連結：是否為藥品專利連結告代之案件
   If InStr("Y20412,Y45493,", Left(pa(75), 6)) > 0 Then '特定代理人
       bolCaseIs959 = True
   Else
       If PUB_ChkCPExist(pa, "959") = True Then  '有收文藥品專利連結告代959
           bolCaseIs959 = True
       End If
   End If
   'end 2021/08/16
   
On Error GoTo CheckingErr

   cnnConnection.BeginTrans
   
   'Add by Morgan 2007/7/18
   If m_928Upd = True And m_928CP09 <> "" Then
      PUB_928Update pa, m_928CP09
   End If
   'end 2007/7/18
 
   'Modify by Morgan 2004/8/4
   '台灣無公告日或公告日>9307012的同時更新公告日
   If m_bolNew = True And txtPA14 <> "" Then
      stPA14 = Format(Val(txtPA14) + 19110000)
   Else
   '   stPA14 = "PA14"
       stPA14 = "0"
   End If
   
   'Modify by Morgan 2006/10/20 加PA139聯絡人部門
   'Modify by Morgan 2010/6/1 聯絡人會有單引號要轉
   '20140306START MODIFY By eric 發明人新增要update patent
   'Modified by Morgan 2014/5/2 發明人欄位下面已有寫程式更新,此處改不更新(選單已改有帶名稱會發生錯誤)
   'Memo by Lydia 2021/08/17 刪除舊程式碼：專利發明人在專利基本檔60~69
   strExc(1) = "UPDATE PATENT SET PA17='Y',PA22='" & Text6 & "',PA21=" & TransDate(Text7, 2) & _
      ",PA24=" & TransDate(Text8(0), 2) & ",PA25=" & TransDate(Text8(1), 2) & _
      ",PA05=" & CNULL(ChgSQL(Text10(0))) & ",PA06=" & CNULL(ChgSQL(Text10(1))) & _
      ",PA07=" & CNULL(ChgSQL(Text10(2))) & ",PA51=" & CNULL(ChgSQL(Text33(0))) & ",PA52=" & CNULL(ChgSQL(Text33(1))) & _
      ",PA53=" & CNULL(ChgSQL(Text33(2))) & ",PA54=" & CNULL(ChgSQL(Text33(3))) & ",PA55=" & CNULL(ChgSQL(Text33(4))) & _
      ",PA56=" & CNULL(ChgSQL(Text33(5))) & _
      ",PA48=" & CNULL(ChgSQL(Text9)) & _
      ",PA14=" & stPA14 & ",PA139=" & CNULL(ChgSQL(Text33(6))) & _
      " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   'end 2014/5/2
   '20140306END
      
   cnnConnection.Execute strExc(1)
   
'2012/10/2 CANCEL BY SONIA 業務區改抓最新智權人員之業務區
'   strExc(0) = "SELECT CP12,CP13 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP31='Y'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      '業務區
'      If Not IsNull(RsTemp.Fields(0)) Then strTmp(2) = RsTemp.Fields(0)
'      '智權人員
'      If Not IsNull(RsTemp.Fields(1)) Then strTmp(1) = RsTemp.Fields(1)
'   Else
'      strExc(0) = ""
'      'edit by nickc 2007/02/02 不用 dll 了
'      'If objPublicData.GetStaffArea(strUserNum, strExc(0)) Then
'      If ClsPDGetStaffArea(strUserNum, strExc(0)) Then
'        '智權人員
'         strTmp(1) = strUserNum
'         '業務區
'         strTmp(2) = strExc(0)
'      End If
'   End If

   stCP09 = AutoNo("C", 6)
   '2012/10/2 CANCEL BY SONIA 業務區改抓最新智權人員之業務區
   'strExc(2) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13," & _
      "CP14,CP20,CP26,CP32,CP27) VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & _
      CNULL(TransDate(Label2(1), 2)) & ",'" & stCP09 & "'," & 專利證書 & "," & CNULL(ChgSQL(strTmp(2))) & _
      "," & CNULL(ChgSQL(PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)))) & ",'" & strUserNum & "','N','N','N'," & strSrvDate(1) & ")"
   'Modified by Lydia 2019/03/04 +進度備註cp64
   'strExc(2) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13," & _
      "CP14,CP20,CP26,CP32,CP27) VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & _
      CNULL(TransDate(Label2(1), 2)) & ",'" & stCP09 & "'," & 專利證書 & "," & CNULL(ChgSQL(GetSalesArea(PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))))) & _
      "," & CNULL(ChgSQL(PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)))) & ",'" & strUserNum & "','N','N','N'," & strSrvDate(1) & " )"
   'Modified by Lydia 2019/05/31 外專程序工作大項先不上發文日(整批發文)
   'strExc(2) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13," & _
      "CP14,CP20,CP26,CP32,CP27,CP64) VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & _
      CNULL(TransDate(Label2(1), 2)) & ",'" & stCP09 & "'," & 專利證書 & "," & CNULL(ChgSQL(GetSalesArea(PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))))) & _
      "," & CNULL(ChgSQL(PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)))) & ",'" & strUserNum & "','N','N','N'," & strSrvDate(1) & ", " & CNULL(IIf(m_strMemo <> "", "承辦單備註:" & m_strMemo, "")) & ")"
   stCP13 = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
   stCP12 = GetSalesArea(stCP13)
   'Modified by Lydia 2021/03/29 取消設定，回歸到各區程序
   'strCP14 = Pub_GetSpecMan("外專程序-專利證書")
    strCP14 = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
   
   'Modified by Lydia 2019/06/17 已上閉卷的案件，各項大批進度檔發文日請先上111111
   'strCP48 = CompDate(2, 10, strSrvDate(1))
   'strExc(2) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13," & _
      "CP14,CP20,CP26,CP32,CP27,CP48,CP64) VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & _
      CNULL(TransDate(Label2(1), 2)) & ",'" & stCP09 & "'," & 專利證書 & "," & CNULL(stCP12) & _
      "," & CNULL(stCP13) & "," & CNULL(strCP14) & ",'N','N','N',null," & CNULL(strCP48, True) & ", " & CNULL(IIf(m_strMemo <> "", "承辦單備註:" & m_strMemo, "")) & ")"
   'end 2019/05/31
   strExc(3) = ""
   If Trim(pa(57) & pa(108)) <> "" Then
       strExc(3) = "19221111"
   Else
       strCP48 = CompDate(2, 10, strSrvDate(1))
       strCP48 = CompWorkDay(1, strCP48, 1)     'add by sonia 2025/3/14 若遇假日則提前至前一工作日
       'Added by Lydia 2021/08/16 外專-藥品專利連結：藥品專利連結告代之案件=>通知證書承辦期限=系統日+3工作天
       If bolCaseIs959 = True Then
           strCP48 = CompWorkDay(4, strSrvDate(1))
       End If
       'end 2021/08/16
   End If
   'Modified by Lydia 2021/08/16 +ChgSQL
   'Modified by Lydia 2023/02/22 請取消key證書來函時將備註回寫到進度檔之設定，在對外通知函\證書函則直接抓當時的承辦單設定
   'strExc(2) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13," & _
      "CP14,CP20,CP26,CP32,CP27,CP48,CP64) VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & _
      CNULL(TransDate(Label2(1), 2)) & ",'" & stCP09 & "'," & 專利證書 & "," & CNULL(stCP12) & _
      "," & CNULL(stCP13) & "," & CNULL(strCP14) & ",'N','N','N'," & CNULL(strExc(3), True) & "," & CNULL(strCP48, True) & ", " & CNULL(ChgSQL(IIf(m_strMemo <> "", "承辦單備註:" & m_strMemo, ""))) & ")"
   strExc(2) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13," & _
      "CP14,CP20,CP26,CP32,CP27,CP48) VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & _
      CNULL(TransDate(Label2(1), 2)) & ",'" & stCP09 & "'," & 專利證書 & "," & CNULL(stCP12) & _
      "," & CNULL(stCP13) & "," & CNULL(strCP14) & ",'N','N','N'," & CNULL(strExc(3), True) & "," & CNULL(strCP48, True) & ")"
   cnnConnection.Execute strExc(2)
   'end 2019/06/17
   'Modify by Morgan 2007/7/23 CP20改抓CPM的設定
   'Modify by Morgan 2008/3/27 +pa75
   'Modify by Morgan 2008/4/10 +本所案號
   strCP20 = PUB_GetCP20(pa(1), 專利證書, strCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
   If strCP20 = "" Then
      strSql = "update caseprogress set cp20=NULL,cp16=" & strCP16 & ",cp17=0,cp18=" & strCP16 / 1000 & _
         " where cp09='" & stCP09 & "'"
      cnnConnection.Execute strSql, intI
   End If
   'end 2007/7/23
      
   '更新相同案號的案件進度檔案件性質為"領證及繳年費"(601)且"發文日"有值者, 其實際結果欄為"1"(准)
   strSql = "Update CASEPROGRESS SET CP24='1' WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='601' AND CP27 IS NOT NULL "
   cnnConnection.Execute strSql
   
   'Added by Lydia 2021/08/16 外專-藥品專利連結：藥品專利連結告代之案件=>二次核對承辦期限=系統日+5工作天
   If bolCaseIs959 = True Then
       strExc(4) = CompWorkDay(6, strSrvDate(1))
       strSql = "Update CaseProgress set cp48=" & strExc(4) & " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND cp10='926' and cp159=0 and cp158=0 "
       cnnConnection.Execute strSql
   End If
   'end 2021/08/16
   
   'Add by Morgan 2004/7/12
   '若用新法則更新下一程序年費期限
   If m_bolNew = True Then

      '2008/10/13 modify by sonia
      'strSQL = "update nextprogress set NP08=" & m_strNextFeeDate & ", np09=" & m_strNextDueDate & _
         " where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
         " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
      '2008/11/26 MODIFY BY SONIA 改為關係企業
      'If ChangeCustomerL(pa(75)) = "Y33944010" Then
      '2009/2/5 modify by sonia 若年費代理人非Y33944的關係企業則不掛np15,FCP-011795
      '2009/8/4 MODIFY BY SONIA 加Y48840,Y48196,Y20624
      '2009/8/20 MODIFY BY SONIA 加Y21099
      'Modify by Morgan 2011/3/22 改先存變數才不用重複抓相同資料
      'If Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y33944" Or _
         Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y48840" Or _
         Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y48196" Or _
         Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y20624" Or _
         Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y21099" Then
      'Modified by Morgan 2013/5/2 函數要傳9碼
      'strExc(9) = Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6)
      strExc(9) = PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1")
      'end 2013/5/2
      
'2011/10/12 MODIFY BY SONIA 改用MODULE
'      If strExc(9) = "Y33944" Or strExc(9) = "Y48840" Or strExc(9) = "Y48196" Or _
'         strExc(9) = "Y20624" Or strExc(9) = "Y21099" Then
'
'      '2008/11/26 END
'         strSql = "update nextprogress set NP08=" & m_strNextFeeDate & ", np09=" & m_strNextDueDate & ", NP15='信函要傳真;'||NP15 " & _
'            " where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
'            " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
'
'      '2009/8/4 ADD BY SONIA Y49083年費備註
'      'Modify by Morgan 2011/3/22 改先存變數才不用重複抓相同資料
'      'ElseIf Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y49083" Then
'      ElseIf strExc(9) = "Y49083" Then
'         strSql = "update nextprogress set NP08=" & m_strNextFeeDate & ", np09=" & m_strNextDueDate & ", NP15='只需銀龍加蓋年費回傳章;'||NP15 " & _
'            " where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
'            " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
'      '2009/8/4 END
'
'      'Add by Morgan 2011/3/22 代理人 2011.03.19 指示信--Susan
'      ElseIf strExc(9) = "Y30011" Then
'         strSql = "update nextprogress set NP08=" & m_strNextFeeDate & ", np09=" & m_strNextDueDate & ", NP15='年費函需以EMail傳送,不寄紙本;'||NP15 " & _
'            " where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
'            " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
'
'      Else
'         strSql = "update nextprogress set NP08=" & m_strNextFeeDate & ", np09=" & m_strNextDueDate & _
'            " where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
'            " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
'      End If
      'Modified by Morgan 2012/6/4 +pa26
      'Modified by Morgan 2013/9/11 改抓設定檔
      'strMemo605 = PUB_Get605Memo(strExc(9), ChangeCustomerL(pa(26)), pa(1) & pa(2) & pa(3) & pa(4))
      'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
      'strMemo605 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), ChangeCustomerL(pa(26)))
      strMemo605 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30))
      
      '同時更新智權人員
      'Modified by Lydia 2021/08/16 +ChgSQL
      'Modify By Sindy 2021/8/17 + ,np23=" & m_strAgreeOnDate
      'Modified by Lydia 2022/08/02 備註另外更新；
      'strSql = "update nextprogress set NP08=" & m_strNextFeeDate & ",np09=" & m_strNextDueDate & ",np23=" & m_strAgreeOnDate & _
               ",NP10='" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "',np15=decode(NP15,null,'" & ChgSQL(strMemo605) & "','" & strMemo605 & "'||NP15) " & _
               " where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
               " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
      strSql = "update nextprogress set NP08=" & m_strNextFeeDate & ",np09=" & m_strNextDueDate & ",np23=" & m_strAgreeOnDate & _
               ",NP10='" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "' " & _
               " where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
               " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
'2011/10/12 END
      '2008/10/13 end
      cnnConnection.Execute strSql
      'Added by Lydia 2022/08/02 備註另外更新；
      strSql = "update nextprogress set NP15='" & ChgSQL(strMemo605) & "'||';'||np15 where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
               " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and instr(np15,'" & ChgSQL(strMemo605) & "') = 0 "
      cnnConnection.Execute strSql
      'end 2022/08/02
      'Add by Morgan 2004/7/15
      '若有未發文技術報告時更新文件齊備日及承辦期限
      If PUB_ChkCPExist(pa, "421", 1, m_str421CP09) = True Then
         m_str421EP06 = strSrvDate(1)
         'Modify by Morgan 2007/10/12 承辦期限改呼叫共用函數計算
         'm_str421CP48 = PUB_GetEngDueDate(m_str421EP06, pa(1), "000", "421")
         m_str421CP48 = Pub_GetHandleDay(pa(1), "000", "421", m_str421EP06, , m_str421CP09)
         'end 2007/10/12
         '更新文件齊備日
         strSql = "Update EngineerProgress Set EP06=" & strSrvDate(1) & " Where EP02='" & m_str421CP09 & "' AND EP06 IS NULL"
         cnnConnection.Execute strSql
         If Val(m_str421CP48) > 0 Then
            '更新承辦期限
            strSql = "Update CaseProgress Set CP48=" & m_str421CP48 & " Where CP09='" & m_str421CP09 & "' AND CP48 IS NULL"
            cnnConnection.Execute strSql
         End If
      End If
      'END 2004/7/14
   End If

   '2008/5/16 ADD BY SONIA 更新相關收文號為'A'類最小的總收文號,配合P同時修改
   strSql = "UPDATE CASEPROGRESS A" & _
      " SET A.CP43=(SELECT MIN(B.CP09) FROM CASEPROGRESS B WHERE B.CP01=A.CP01 AND B.CP02=A.CP02 AND B.CP03=A.CP03 AND B.CP04=A.CP04 AND B.CP09<'B')" & _
      " WHERE A.CP09='" & stCP09 & "' AND A.CP43 IS NULL"
   cnnConnection.Execute strSql
   '2008/5/16 END
   
'   '20140306START ADD By eric 更新 PATENT中發明人資料
'   For i = 60 To 69
'      strTmp(1) = txtInvField((i - 60) * 3)      '發明人中文名稱
'      strTmp(2) = txtInvField((i - 60) * 3 + 1)  '發明人英文名稱
'      strTmp(3) = txtInvField((i - 60) * 3 + 2)  '發明人日文名稱
'
'      If Combo4(i - 60) = "" And (strTmp(1) <> "" Or strTmp(2) <> "" Or strTmp(3) <> "") Then
'          '自行輸入則客戶發明人檔IN01=PA26
'          'Modified by Morgan 2014/4/22
'          'InsInventor m_PA30, pa(26), strTmp(1), strTmp(2), strTmp(3), txtIN11(i - 60)
'          InsInventor m_PA30, Left(pa(26), 8), strTmp(1), strTmp(2), strTmp(3), txtIN11(i - 60)
'          pa(i) = m_PA30: stUpdates = stUpdates & ",pa" & i & "=" & CNULL(ChgSQL(pa(i)))
'      Else
'          pa(i) = Replace(Right(Combo4(i - 60).Text, 11), ")", ""): stUpdates = stUpdates & ",pa" & i & "=" & CNULL(ChgSQL(pa(i)))
'      End If
'
'      If stUpdates <> "" Then
'          stUpdates = Mid(stUpdates, 2)
'          strSql = "UPDATE PATENT SET " & stUpdates & " WHERE PA01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
'          Pub_SeekTbLog strSql
'          cnnConnection.Execute strSql
'          stUpdates = ""
'      End If
'   Next
'   '20140306END
   'Add By Sindy 2014/11/11
   If cmdAddRow.Tag = "I" Or cmdDelRow.Tag = "D" Then '有異動發明人資料
      '全部刪除,重新新增
      strSql = "delete from patentInventor where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4))
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      For ii = 1 To GRD1.Rows - 1
         '自行輸入則客戶發明人檔IN01=PA26
         If Trim(GRD1.TextMatrix(ii, 1)) = "" And _
            (Trim(GRD1.TextMatrix(ii, 2)) <> "" Or Trim(GRD1.TextMatrix(ii, 3)) <> "" Or Trim(GRD1.TextMatrix(ii, 4)) <> "") Then
            'Modified by Morgan 2015/12/14 造字後面可能會加空白不可用Trim
            'InsInventor m_PA30, pa(26) & String(8 - Len(Left(pa(26), 8)), "0"), Trim(GRD1.TextMatrix(ii, 2)), Trim(GRD1.TextMatrix(ii, 3)), Trim(GRD1.TextMatrix(ii, 4)), Trim(GRD1.TextMatrix(ii, 5))
            InsInventor m_PA30, pa(26) & String(8 - Len(Left(pa(26), 8)), "0"), LTrim(GRD1.TextMatrix(ii, 2)), Trim(GRD1.TextMatrix(ii, 3)), LTrim(GRD1.TextMatrix(ii, 4)), Trim(GRD1.TextMatrix(ii, 5))
            'end 2015/12/14
            GRD1.TextMatrix(ii, 1) = m_PA30
         End If
         strSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
                  CNULL(pa(1)) & "," & CNULL(pa(2)) & "," & CNULL(pa(3)) & "," & CNULL(pa(4)) & "," & ii & ",'" & GRD1.TextMatrix(ii, 1) & "')"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      Next ii
   End If
   '2014/11/11 END
   
   'Added by Morgan 2015/5/20
   If m_bolDualApply Then
      strSql = "UPDATE PATENT SET PA57='Y',PA58=" & strSrvDate(1) & ",PA59='99', PA91='發明案(" & pa(1) & pa(2) & pa(3) & pa(4) & ")已發證系統自動上閉卷;'||PA91 WHERE PA01='" & m_stUPA(1) & "' AND PA02='" & m_stUPA(2) & "' AND PA03='" & m_stUPA(3) & "' AND PA04='" & m_stUPA(4) & "' AND PA57 IS NULL"
      cnnConnection.Execute strSql, intI
   'Added by Morgan 2019/7/18
   End If
   If m_bolDualApply Or m_bolDualApp1 Then
   'end 2019/7/18
      '下一程序年費期限自動不續辦
      strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & strSrvDate(1) & ",NP12='99', NP15='發明案(" & pa(1) & pa(2) & pa(3) & pa(4) & ")已發證系統自動不續辦;'||NP15 WHERE NP02='" & m_stUPA(1) & "' AND NP03='" & m_stUPA(2) & "' AND NP04='" & m_stUPA(3) & "' AND NP05='" & m_stUPA(4) & "' AND NP06 IS NULL AND NP07='605' AND NP09>=" & strSrvDate(1)
      cnnConnection.Execute strSql, intI
   End If
   'end 2015/5/20
   'Added by Morgan 2019/7/18
   If m_bolDualApp1 Or m_bolDualApp2 Then
      '提醒人員:程序(管制人),承辦
      strExc(3) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) & "," & stCP13
      If m_bolDualApp1 Then
         '期限: 5工作天
         strExc(1) = CompWorkDay(5, strSrvDate(1))
         '事由
         strExc(4) = "優先請款或處理未發文程序完再上閉卷"
         PUB_AddFCPStaffCalendar strExc(1), "1", strExc(3), strExc(4), strExc(3), "1", m_stUPA(1), m_stUPA(2), m_stUPA(3), m_stUPA(4)
      Else
         '期限: 當天
         strExc(1) = strSrvDate(1)
         '事由
         strExc(4) = "因一案二請新型(" & m_stUPA(1) & "-" & m_stUPA(2) & IIf(m_stUPA(3) & m_stUPA(4) = "000", "", "-" & m_stUPA(3) & "-" & m_stUPA(4)) & ")未自動閉卷,請確認相關資料是否完整,再人工閉卷" & vbCrLf & _
            "(無擇一申復發文及基本檔是否放棄新型為""Y"",且無非屬相同創作)"
         PUB_AddFCPStaffCalendar strExc(1), "1", strExc(3), strExc(4), strExc(3), "1", pa(1), pa(2), pa(3), pa(4)
      End If
      
   End If
   'end 2019/7/18
   'Added by Lydia 2022/08/11 遇特定代理人當Key完證書函 (紙本公文), 按確定時, 請系統自動發一封Email通知核對已准專利承辦的工程師
   If InStr("Y52212000,Y45799050,Y45799070,Y52341010,Y28343010", ChangeCustomerL(pa(75))) > 0 Then
       strExc(0) = "select x1.*,oman from ( select cp09,cp14,st04,decode(st16,'1','T','2','R','3','S','4','T1','O') st16 from caseprogress, staff " & _
                        "where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' " & _
                        "and cp159=0 and cp10='926' and cp14=st01(+)) X1, setspecman where st16=ocode(+) order by cp09 desc "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
           If "" & RsTemp.Fields("st04") <> "1" Then
               '若無或離職，則改發給該組的主管
               strExc(1) = "" & RsTemp.Fields("oman")
           Else
               strExc(1) = "" & RsTemp.Fields("cp14")
           End If
           strExc(2) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "請提供1. 已准claim數量2. 已准版本英文Claim之WORD檔"
           'Add by Amy 2025/08/05 後續准駁簡單報告=Y,輸C類來函[主旨]最前面加【請簡單報告】-Winfrey
           If pa(89) = "Y" Then strExc(2) = "【請簡單報告】" & strExc(2)
           
           strExc(3) = "Dear " & PUB_ReadUserData(strExc(1)) & "，" & vbCrLf & vbCrLf & _
                           pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "依客戶指示，請您儘速提供" & vbCrLf & _
                           "1. 已准Claim數量 " & vbCrLf & _
                           "2. 已准版本英文Claim之WORD檔" & vbCrLf & vbCrLf & _
                           "請將上述資訊Email給該區程序優先寄證書 , 謝謝!"
           strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                       " values( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                       ",'" & strExc(2) & "','" & strExc(3) & "',null)"
           cnnConnection.Execute strSql, intI
       End If
   End If
   'end 2022/08/11
   
   'Added by Lydia 2024/05/30 勘誤公報控管：有「更正402」並且未有發文日時，帶入「公告公報1228」總收文號
   strSql = Pub_GetProcCRC("1", pa(1), pa(2), pa(3), pa(4))
   If strSql <> "" Then
      cnnConnection.Execute strSql, intI
   End If
   'end 2024/05/30
   
   'Added by Morgan 2023/1/16 電子公文
   If m_DocNo <> "" Then
      If m_DocWord <> "" Then
         strSql = "UPDATE CASEPROGRESS set cp08='" & m_DocWord & "字第" & m_DocNo & "號'" & _
            " WHERE CP09='" & stCP09 & "'"
         cnnConnection.Execute strSql, intI
      End If
      PUB_UpdateEdocRec m_DocNo, stCP09, pa(1), pa(2), pa(3), pa(4), 專利證書
   End If
   'end 2023/1/16
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   FormSave = False
   MsgBox Err.Description
End Function

'20140306ADD By eric
Private Sub InsInventor(ByRef m_PA30, ByVal InvNo As String, ByVal InvCh As String, ByVal InvEng As String, ByVal InvJP As String, ByVal IN11 As String)
    Dim strIns As String, m_IN01 As String, m_IN02 As String
    
    'Modified by Morgan 2015/1/6 有更名會錯
    'm_IN01 = InvNo & String(8 - Len(InvNo), "0")
    m_IN01 = Left(ChangeCustomerL(InvNo), 8)
    'end 2014/1/6
    m_IN02 = PUB_GetNewIN02(m_IN01)
    m_PA30 = m_IN01 & m_IN02
    strIns = "Insert Into Inventor (IN01,IN02,IN04,IN05,IN06,IN11) Values(" & CNULL(ChgSQL(m_IN01)) & "," & CNULL(ChgSQL(m_IN02)) & "," & _
                CNULL(ChgSQL(InvCh)) & "," & CNULL(ChgSQL(InvEng)) & "," & CNULL(ChgSQL(InvJP)) & "," & CNULL(ChgSQL(IN11)) & ")"
    cnnConnection.Execute strIns
End Sub

''Add By Sindy 2015/3/5 修改發明人的中英日名稱時可存檔
'Private Sub UpdateInventor()
'    Dim strUpd As String, m_IN01 As String, m_IN02 As String
'
'    m_IN01 = Left(Combo4.Text, 8)
'    m_IN02 = Mid(Combo4.Text, 9, 2)
'    strUpd = "update Inventor set" & _
'             " IN04=" & CNULL(txtInvField(0)) & ",IN05=" & CNULL(txtInvField(1)) & ",IN06=" & CNULL(txtInvField(2)) & _
'             " where IN01='" & m_IN01 & "' and IN02='" & m_IN02 & "'"
'    cnnConnection.Execute strUpd
'End Sub

'20140306ADD By eric 清除資料表
Private Sub FormClear()
 Dim Lbl As Object
 Dim Cmb4 As Object
 
   SSTab1.TabEnabled(1) = True
 
   For Each Lbl In Label2
      Lbl.Caption = ""
   Next
   
   Combo4.Clear
   
   For Each m_otxt In txtInvField
      m_otxt.Text = ""
      m_otxt.Tag = "" 'Add By Sindy 2015/3/5
   Next
   
   'IN11 國籍
   txtIN11 = ""
   Lb_IN11N = "" 'Add By Sindy 2015/12/4
'   txtIN11.Visible = False
'   Lb_IN11.Visible = False
'   Lb_IN11N.Visible = False
End Sub

'Add By Sindy 2017/3/15 向上移
Private Sub cmdUp_Click()
Dim ii As Integer, jj As Integer
   
   If pPrevRow > 1 And GRD1.Rows - 1 > 0 Then
      If GRD1.TextMatrix(pPrevRow, 0) <> "" Then '點選的資料列有資料
         '記錄暫存Grid
         GRDtmp.Clear
         Call SetGrd(GRDtmp): GRDtmp.Visible = False
         'Set GRDtmp.DataSource = GRD1.Recordset
         For ii = 1 To GRD1.Rows - 1
            If ii > 1 Then GRDtmp.AddItem ""
            For jj = 0 To GRD1.Cols - 1
               GRDtmp.TextMatrix(ii, jj) = GRD1.TextMatrix(ii, jj)
            Next jj
         Next ii
         GRD1.Enabled = False
         '處理上移後的上方資料
         For ii = 1 To pPrevRow - 2
            For jj = 0 To GRD1.Cols - 1
               GRD1.TextMatrix(ii, jj) = GRDtmp.TextMatrix(ii, jj)
            Next jj
         Next ii
         '對換資料列
         For jj = 0 To GRD1.Cols - 1
            GRD1.TextMatrix(pPrevRow - 1, jj) = GRDtmp.TextMatrix(pPrevRow, jj)
         Next jj
         For jj = 0 To GRD1.Cols - 1
            GRD1.TextMatrix(pPrevRow, jj) = GRDtmp.TextMatrix(pPrevRow - 1, jj)
         Next jj
         '處理上移後的下方資料
         For ii = pPrevRow + 1 To GRD1.Rows - 1
            For jj = 0 To GRD1.Cols - 1
               GRD1.TextMatrix(ii, jj) = GRDtmp.TextMatrix(ii, jj)
            Next jj
         Next ii
         cmdAddRow.Tag = "I" '記錄有異動資料
         Call SetGrd1SelRow(pPrevRow - 1)
         GRD1.Enabled = True
      End If
   ElseIf pPrevRow = 1 Then
      MsgBox "已到第一筆！", vbCritical + vbOKOnly, MsgText(9001)
   Else
      MsgBox "欲移動資料項目，請選擇一筆資料！", vbCritical + vbOKOnly, MsgText(9001)
   End If
End Sub

'Add By Sindy 2017/3/15 向下移
Private Sub cmdDown_Click()
Dim ii As Integer, jj As Integer
   
   If (pPrevRow > 0 And pPrevRow < GRD1.Rows - 1) And GRD1.Rows - 1 > 0 Then
      If GRD1.TextMatrix(pPrevRow, 0) <> "" Then '點選的資料列有資料
         '記錄暫存Grid
         GRDtmp.Clear
         Call SetGrd(GRDtmp): GRDtmp.Visible = False
         'Set GRDtmp.DataSource = GRD1.Recordset
         For ii = 1 To GRD1.Rows - 1
            If ii > 1 Then GRDtmp.AddItem ""
            For jj = 0 To GRD1.Cols - 1
               GRDtmp.TextMatrix(ii, jj) = GRD1.TextMatrix(ii, jj)
            Next jj
         Next ii
         GRD1.Enabled = False
         '處理下移後的上方資料
         For ii = 1 To pPrevRow - 1
            For jj = 0 To GRD1.Cols - 1
               GRD1.TextMatrix(ii, jj) = GRDtmp.TextMatrix(ii, jj)
            Next jj
         Next ii
         '對換資料列
         For jj = 0 To GRD1.Cols - 1
            GRD1.TextMatrix(pPrevRow + 1, jj) = GRDtmp.TextMatrix(pPrevRow, jj)
         Next jj
         For jj = 0 To GRD1.Cols - 1
            GRD1.TextMatrix(pPrevRow, jj) = GRDtmp.TextMatrix(pPrevRow + 1, jj)
         Next jj
         '處理下移後的下方資料
         For ii = pPrevRow + 2 To GRD1.Rows - 1
            For jj = 0 To GRD1.Cols - 1
               GRD1.TextMatrix(ii, jj) = GRDtmp.TextMatrix(ii, jj)
            Next jj
         Next ii
         cmdAddRow.Tag = "I" '記錄有異動資料
         Call SetGrd1SelRow(pPrevRow + 1)
         GRD1.Enabled = True
      End If
   ElseIf pPrevRow = GRD1.Rows - 1 Then
      MsgBox "已到最末筆！", vbCritical + vbOKOnly, MsgText(9001)
   Else
      MsgBox "欲移動資料項目，請選擇一筆資料！", vbCritical + vbOKOnly, MsgText(9001)
   End If
End Sub

'Add By Sindy 2017/3/15
Private Sub SetGrd1SelRow(intSelRow As Integer)
Dim nRow As Integer, nCol As Integer
Dim iCol As Integer
   
   With GRD1
      .Visible = False
      nRow = intSelRow
      If nRow > 0 Then
         nCol = .col
         If pPrevRow > 0 Then
            If pPrevRow <> nRow Then
               .row = pPrevRow
               .TextMatrix(pPrevRow, 0) = ""
               If .FixedCols > 0 Then
                  .col = .FixedCols - 1
                  .CellBackColor = .BackColorFixed
                  .CellForeColor = .ForeColor
               End If
               For iCol = .FixedCols To .Cols - 1
                  .col = iCol
                  .CellBackColor = .BackColor
               Next
            End If
         End If
         If nRow > 0 Then
            .row = nRow
            .TextMatrix(nRow, 0) = "V"
            If .FixedCols > 0 Then
               .col = .FixedCols - 1
               .CellBackColor = .BackColorSel
               .CellForeColor = .ForeColorSel
            End If
            For iCol = .FixedCols To .Cols - 1
              .col = iCol
              .CellBackColor = &HFFC0C0
            Next
         End If
         '.col = nCol
         pPrevRow = intSelRow
         Call SetCombo4Data(.TextMatrix(nRow, 1))
         If .TextMatrix(nRow, 1) = "" Then
            txtInvField(0) = .TextMatrix(nRow, 2)
            txtInvField(1) = .TextMatrix(nRow, 3)
            txtInvField(2) = .TextMatrix(nRow, 4)
            txtIN11 = .TextMatrix(nRow, 5)
            cmdUpdRow.Enabled = True
            cmdAddRow.Enabled = False
         End If
      End If
      .Visible = True
   End With
End Sub

'Add By Sindy 2015/3/5
Private Sub cmdUpdRow_Click()
   'If Trim(Combo4.Text) = "" Then
      Me.GRD1.TextMatrix(pPrevRow, 2) = txtInvField(0)
      Me.GRD1.TextMatrix(pPrevRow, 3) = txtInvField(1)
      Me.GRD1.TextMatrix(pPrevRow, 4) = txtInvField(2)
      Me.GRD1.TextMatrix(pPrevRow, 5) = txtIN11
      cmdUpdRow.Enabled = False
   'End If
End Sub

'20140306START MODIFY By eric
'Private Sub Combo1_Click(Index As Integer)
'   If bolFormLoad Then ChgType Index + 10
'End Sub
'20140306END

Private Sub Combo2_Click(Index As Integer)
 Dim i As Integer, strTmp As String
   If Combo2(Index) = "" Then
      For i = 0 To 2
         Text33(i + Index * 3) = ""
      Next
      Exit Sub
   End If
   
   strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
   Select Case Text2
      Case "FCP"
         If pa(75) <> "" Then
            Select Case strTmp
               Case "1"
                  strExc(1) = "FA07,FA08,FA09"
               Case "2"
                  strExc(1) = "FA52,FA53,FA54"
            End Select
         Else
            Select Case strTmp
               Case "1"
                  strExc(1) = "CU58,CU59,CU60"
               Case "2"
                  strExc(1) = "CU61,CU62,CU63"
            End Select
         End If
      Case "FG"
         If pa(26) <> "" Then
            Select Case strTmp
               Case "1"
                  strExc(1) = "FA07"
               Case "2"
                  strExc(1) = "FA52"
            End Select
         Else
            Select Case strTmp
               Case "1"
                  strExc(1) = "CU58"
               Case "2"
                  strExc(1) = "CU61"
            End Select
         End If
   End Select
   
   strExc(2) = ChgFagent(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
   strExc(3) = ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
   Select Case Text2
      Case "FCP"
         If pa(75) <> "" Then
            strExc(0) = "SELECT " & strExc(1) & " FROM FAGENT WHERE " & strExc(2)
         Else
            strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & strExc(3)
         End If
      Case "FG"
         If pa(26) <> "" Then
            strExc(0) = "SELECT " & strExc(1) & " FROM FAGENT WHERE " & strExc(2)
         Else
            strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & strExc(3)
         End If
   End Select
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Select Case Text2
         Case "FCP"
            For i = 0 To 2
               If Not IsNull(RsTemp.Fields(i)) Then
                  Text33(i + Index * 3) = RsTemp.Fields(i)
               Else
                  Text33(i + Index * 3) = ""
               End If
            Next
         Case "FG"
            If Not IsNull(RsTemp.Fields(0)) Then Text33(0) = RsTemp.Fields(0)
      End Select
   End If
End Sub

'20140306ADD By eric
Private Sub Combo4_Click()
   Dim strMain As String, i As Integer
   strMain = Replace(Right(Combo4.Text, 11), ")", "")
   For i = 0 To 2
      txtInvField(i).Text = ""
      txtInvField(i).Tag = "" 'Add By Sindy 2015/3/5
   Next
'   Lb_IN11.Visible = False
'   txtIN11.Visible = False
'   Lb_IN11N.Visible = False
   cmdUpdRow.Enabled = False 'Add By Sindy 2015/3/5
   cmdAddRow.Enabled = True 'Add By Sindy 2015/3/5
   If Len(strMain) > 0 Then
      'cmdUpdRow.Enabled = True 'Add By Sindy 2015/3/5
      If ClsLawGetInventor(strMain, strExc) Then
         For i = 0 To 2
            txtInvField(i).Text = strExc(i + 1)
            txtInvField(i).Tag = txtInvField(i).Text 'Add By Sindy 2015/3/5
         Next
         'Add By Sindy 2015/12/4
         txtIN11 = strExc(6)
         Call txtIN11_Validate(False)
         '2015/12/4 END
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   '20140306START ADD By eric
   FormClear
   SSTab1.Tab = 0
   '20140306END
   Text2 = strExc(1)
   Text3 = strExc(2)
   Text4 = strExc(3)
   Text5 = strExc(4)
   'Add by Lydia 2014/10/27 日文定稿頁籤切換提示發明人語系
   m_LetterLanguage = PUB_GetLanguage(Text2, Text3, Text4, Text5)
   ReadPatent
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add by Lydia 2014/10/22 發明人輸入比對兼自動代入(模糊比對)
    ' 刪除串列結構
    If m_InventorListCount > 0 Then
       Erase m_InventorList
    End If
    m_InventorListCount = 0
    
    PUB_SendMailCache 'Added by Lydia 2022/08/11
    
   Set frm06010605_2 = Nothing
End Sub

'Add By Sindy 2014/11/11
Private Sub SetGrd(tmpGrd As MSHFlexGrid)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   '                        0    1             2           3           4           5       6
   arrGridHeadText = Array("V", "發明人編號", "中文名稱", "英文名稱", "日文名稱", "國籍", "申請人1")
   arrGridHeadWidth = Array(200, 1100, 2200, 2200, 2200, 0, 0)
   tmpGrd.Visible = False
   tmpGrd.Cols = UBound(arrGridHeadText) + 1
   tmpGrd.Rows = 2
   For iRow = 0 To tmpGrd.Cols - 1
      tmpGrd.row = 0
      tmpGrd.col = iRow
      tmpGrd.Text = arrGridHeadText(iRow)
      tmpGrd.ColWidth(iRow) = arrGridHeadWidth(iRow)
      tmpGrd.CellAlignment = flexAlignCenterCenter
   Next
   tmpGrd.Visible = True
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, i As Integer, strTempName As String, rsTemp1 As New ADODB.Recordset
Dim j As Integer
'Add by Morgan 2004/8/5
Dim varPA72 As Variant
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   For Each Lbl In Label2
      Lbl = ""
   Next
 '20140306START REMARK By eric   發明人顯示方式變更
 '  For Each lbl In Label4
 '     lbl = ""
 '  Next
 '  For Each lbl In Label6
 '     lbl = ""
 '  Next
 '  For Each lbl In Label8
 '     lbl = ""
 '  Next
 '20140306END
   ReDim pa(TF_PA) 'Add by Morgan 2006/10/20
   Text9 = ""
   pa(1) = Text2
   pa(2) = Text3
   pa(3) = Text4
   pa(4) = Text5
   Label2(1) = frm06010605_1.Text5
   
   'Add By Sindy 2014/11/11
   cmdAddRow.Tag = ""
   cmdDelRow.Tag = ""
   GRD1.Clear
   Call SetGrd(GRD1)
   '2014/11/11 END
   
   'Modify by Morgan 2006/10/20 改不Call Dll
   'If objPublicData.ReadPatentDatabase(pa(), intWhere) Then
   If PUB_ReadPatentDatabase(pa(), intWhere) Then
      Text1 = pa(11)
      'Added by Morgan 2012/1/4
      If Mid(pa(11), 4, 1) = "5" Then m_bolIs117 = True
      
      If pa(8) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, pA(8), strTempName, False, 台灣國家代號) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, pa(8), strTempName, False, 台灣國家代號) = 1 Then
            Label2(0) = strTempName
         End If
      End If
      For i = 2 To 6
         If pa(i + 24) <> "" Then ChgType (i)
      Next
      Text6 = pa(22)
      Text7 = pa(21)
        '專用期起日預設為公告日
'      Text8(0) = pa(24)
      'Modify by Morgan 2004/10/7 專用期改西元
      'Text8(0) = pa(14)
      'Text8(1) = pa(25)
      'Modified by Morgan 2012/1/6 積體電路佈局例外
      'Text8(0) = TransDate(pa(14), 2)
      If m_bolIs117 = True Then
         Text8(0) = TransDate(pa(24), 2)
      Else
         Text8(0) = TransDate(pa(14), 2)
      End If
      Text8(1) = TransDate(pa(25), 2)
      
      Text9 = pa(48)
      For i = 0 To 2
         Text10(i) = pa(i + 5)
      Next
      
      Combo4.AddItem ""           '20140306 MODIFY By eric
      strTempName = ""
      For i = 26 To 30
            'Modify By Cheng 2003/03/24
            '申請人代號只抓8碼否則補滿
'         If pa(i) <> "" Then strTempName = strTempName & "'" & Left(pa(i), 8) & _
'            String(8 - Len(pa(i)), "0") & "',"
         If pa(i) <> "" Then strTempName = strTempName & "'" & Left(pa(i) & "000000000", 8) & "',"
      Next
      If strTempName <> "" Then strTempName = Left(strTempName, Len(strTempName) - 1)

      'Modified by Morgan 2014/5/2 發明人要帶出名稱
      'strExc(0) = "SELECT IN01||IN02 FROM INVENTOR WHERE IN01 IN (" & strTempName & ")"
      
      'Add by Lydia 2014/10/22 發明人輸入比對兼自動代入(模糊比對)
'      strExc(0) = "SELECT NVL(IN05,NVL(IN04,IN06))||'('||IN01||IN02||')' FROM INVENTOR WHERE IN01 IN (" & strTempName & ")"
       strExc(0) = "SELECT NVL(IN05,NVL(IN04,IN06))||'('||IN01||IN02||')' as sort,IN01, IN02, IN04, IN05, IN06 " & _
                  "FROM INVENTOR WHERE IN01 IN (" & strTempName & ") order by sort"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      With RsTemp
         If intI = 1 Then
            Do While Not .EOF
               Combo4.AddItem .Fields(0)           '20140306 MODIFY By eric
               'Add by Lydia 2014/10/22 發明人輸入比對兼自動代入(模糊比對)
               If RsTemp.AbsolutePosition = 1 Then
                  Erase m_InventorList '清空陣列
                  ReDim m_InventorList(RsTemp.RecordCount - 1) '定義陣列
                  m_InventorListCount = 0
               End If
                  strExc(1) = "" & RsTemp.Fields("IN01")
                  strExc(2) = "" & RsTemp.Fields("IN02")
                  strExc(4) = "" & RsTemp.Fields("IN04")
                  strExc(5) = "" & RsTemp.Fields("IN05")
                  strExc(6) = "" & RsTemp.Fields("IN06")
                  AddInventor strExc(1), strExc(2), strExc(4), strExc(5), strExc(6)
               'Add by Lydia 2014/10/22 .end
               .MoveNext
            Loop
         End If
      End With
      'Add By Sindy 2014/11/10
      'Modify By Sindy 2015/12/4 + na03
      StrSQLa = "SELECT '' as V,pi06 as 發明人編號,in04 as 中文名稱,in05 as 英文名稱,in06 as 日文名稱,na03 as 國籍,'' as 申請人1" & _
                " from PatentInventor,Inventor,nation" & _
                " where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
                " and substr(pi06,1,8)=in01(+) and substr(pi06,9,2)=in02(+)" & _
                " and in11=na01(+)" & _
                " order by pi05 asc"
      If rsA.State <> adStateClosed Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         Set GRD1.Recordset = rsA
      End If
      '2014/11/10 END
'      '20140306START Modify By eric
'      For i = 0 To 9
'         If pa(i + 60) <> "" Then
'            For j = 0 To Combo4(i).ListCount - 1
'               Combo4(i).ListIndex = j
'               'Modified by Morgan 2014/5/2
'               'If Combo4(i).Text = pa(i + 60) Then
'               If Replace(Right(Combo4(i).Text, 11), ")", "") = pa(i + 60) Then
'               'end 2014/5/2
'                  bolFormLoad = True
'                  ChgType i + 10
'                  Exit For
'               End If
'            Next
'            If bolFormLoad = False Then Combo4(i).ListIndex = 0
'            bolFormLoad = False
'         End If
'      Next
'      'For i = 0 To 9
'      '   If pa(i + 60) <> "" Then
'      '      For j = 0 To Combo1(i).ListCount - 1
'      '         Combo1(i).ListIndex = j
'      '         If Combo1(i).Text = pa(i + 60) Then
'      '            bolFormLoad = True
'      '            ChgType (i + 10)
'      '            Exit For
'      '         End If
'      '      Next
'      '      If bolFormLoad = False Then Combo1(i).ListIndex = 0
'      '      bolFormLoad = False
'      '   End If
'      'Next
'      '20140306END
            
      bolFormLoad = True
      For i = 0 To 5
         Text33(i) = pa(i + 51)
      Next
      Text33(6) = pa(139) 'Add by Morgan 2006/10/20
      If pa(75) <> "" Then
         
         Select Case pa(85)
            Case 1
               strExc(0) = "FA07,FA52"
            Case 2
               strExc(0) = "FA08,FA53"
            Case 3
               strExc(0) = "FA09,FA54"
            Case Else
               strExc(0) = "FA08,FA53"
         End Select
         
         strExc(0) = "SELECT " & strExc(0) & " FROM FAGENT WHERE " & ChgFagent(pa(75))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If IsNull(RsTemp.Fields(0)) Then
               strExc(0) = ""
            Else
               strExc(0) = "-" & RsTemp.Fields(0)
            End If
            Combo2(0).AddItem pa(75) & "-1" & strExc(0)
            Combo2(1).AddItem pa(75) & "-1" & strExc(0)
            If IsNull(RsTemp.Fields(1)) Then
               strExc(0) = ""
            Else
               strExc(0) = "-" & RsTemp.Fields(1)
            End If
            Combo2(0).AddItem pa(75) & "-2" & strExc(0)
            Combo2(1).AddItem pa(75) & "-2" & strExc(0)
         End If
      Else
         For i = 26 To 30
            If pa(i) <> "" Then
               Select Case pa(85)
                  Case 1
                     strExc(0) = "CU58,CU61"
                  Case 2
                     strExc(0) = "CU59,CU62"
                  Case 3
                     strExc(0) = "CU60,CU63"
                  Case Else
                     strExc(0) = "CU59,CU62"
               End Select
               strExc(0) = "SELECT " & strExc(0) & " FROM CUSTOMER WHERE " & ChgCustomer(pa(i))
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  For j = 1 To 2
                     If IsNull(RsTemp.Fields(j - 1)) Then
                        strExc(0) = ""
                     Else
                        strExc(0) = "-" & RsTemp.Fields(j - 1)
                     End If
                     Combo2(0).AddItem pa(i) & "-" & j & strExc(0)
                     Combo2(1).AddItem pa(i) & "-" & j & strExc(0)
                  Next
               End If
            End If
         Next
      End If
   End If
   
   If Combo2(0).ListCount > 0 And Text33(0) = "" Then Combo2(0).ListIndex = 1
   If Combo2(1).ListCount > 0 And Text33(3) = "" Then Combo2(1).ListIndex = 1
    'Add By Cheng 2003/04/01
    '顯示聯絡人
    For i = 0 To 5
        Text33(i) = pa(i + 51)
        If i >= 0 And i <= 2 Then
            If Me.Text33(i).Text <> "" Then Me.Combo2(0).Enabled = False
        Else
            If Me.Text33(i).Text <> "" Then Me.Combo2(1).Enabled = False
        End If
    Next
   
   strExc(0) = "SELECT NP09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND NP07=" & 年費 & " AND NP06 IS NULL"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then If Not IsNull(RsTemp.Fields(0)) Then Label2(7) = TransDate(RsTemp.Fields(0), 1)
   '93.3.6 ADD BY SONIA
   If Label2(7) = "" Then
      'Modified by Morgan 2012/1/4 積體電路案件例外--David
      If m_bolIs117 = False Then
         MsgBox "此案下一程序無年費期限, 請確認 !", vbCritical
      End If
   End If
   '93.3.6 END
   
   'Add by Morgan 2004/8/4
   '取得專用期起訖
   Dim strYear As String
   
   
   Call GetMoneyDate(Val(pa(8)) + 10, pa(9), pa, DATE1, strYear, DATE2)
   
   '台灣無公告日或93.7.1以後要輸入公告日
   m_bolNew = False
   lblPA14.Visible = False
   txtPA14.Visible = False
   
   'Modified by Morgan 2012/1/4 積體電路案件例外--David
   'If pa(9) = 台灣國家代號 Then
   If pa(9) = 台灣國家代號 And m_bolIs117 = False Then
   
      '公告日<93.7.1以前的新型專用期為12年
      If Val(pa(14)) > 0 And Val(pa(14)) < 930701 Then
         If pa(8) = "2" Then
            DATE2 = CompDate(2, -1, CompDate(0, 12, Val(pa(10)) + 19110000))
         End If
      '公告日93.7.1以後要輸入公告日
      Else
         m_bolNew = True
         
         lblPA14.Visible = True
         txtPA14.Visible = True
         If pa(72) <> "" Then
            varPA72 = Split(pa(72), ",")
            m_intLastYear = Val(varPA72(UBound(varPA72)))
         Else
            m_intLastYear = 0
         End If
         'Add by Morgan 2004/8/17
         '台灣新法公報會回存證書號故若無發證日時要做雙重檢查
         If pa(21) = "" Then
            Select Case pa(8)
               Case "1": Text6.Text = "I"
               Case "2": Text6.Text = "M"
               Case "3": Text6.Text = "D"
            End Select
            Text6.SelStart = 1
            Text6.SelLength = 0
         '若已發證則不用
         Else
            txtPA14.Text = pa(14)
         End If
      End If
   End If
   
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTxt(1 To 5) As String
   ChgType = False
   Select Case i
      Case 2, 3, 4, 5, 6
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCustomer(pA(i + 24), strTxt(1)) Then
         If ClsPDGetCustomer(pa(i + 24), strTxt(1)) Then
            Label2(i) = strTxt(1)
            ChgType = True
         End If
'      '20140306START MODIFY By eric
'      Case 10, 11, 12, 13, 14, 15, 16, 17, 18, 19
'         'Modified by Morgan 2014/5/2
'         'If ClsLawGetInventor(Combo4(i - 10).Text, strTxt) = True Then
'         If ClsLawGetInventor(Replace(Right(Combo4(i - 10).Text, 11), ")", ""), strTxt) = True Then
'            txtInvField((i - 10) * 3) = strTxt(1)
'            txtInvField((i - 10) * 3 + 1) = strTxt(2)
'            txtInvField((i - 10) * 3 + 2) = strTxt(3)
'            ChgType = True
'         Else
'            txtInvField((i - 10) * 3) = ""
'            txtInvField((i - 10) * 3 + 1) = ""
'            txtInvField((i - 10) * 3 + 2) = ""
'         End If
'      Case 10, 11, 12, 13, 14, 15, 16, 17, 18, 19
'         'edit by nickc 2007/02/05 不用 dll 了
'         'If objLawDll.GetInventor(Combo1(i - 10).Text, strTxt) = True Then
'         If ClsLawGetInventor(Combo1(i - 10).Text, strTxt) = True Then
'            Label4(i - 10) = strTxt(1)
'            Label6(i - 10) = strTxt(2)
'            Label8(i - 10) = strTxt(3)
'            ChgType = True
'         Else
'            Label4(i - 10) = ""
'            Label6(i - 10) = ""
'            Label8(i - 10) = ""
'         End If
'      '2040306END
   End Select
End Function

'Add By Sindy 2014/11/11
Private Sub Grd1_Click()
Dim nCol As Integer, nRow As Integer
Dim iCol As Integer
   
   With GRD1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow > 0 Then 'And .TextMatrix(nRow, 1) <> "" Then
      nCol = .col
      If pPrevRow > 0 Then
         If pPrevRow <> nRow Then
            .row = pPrevRow
            .TextMatrix(pPrevRow, 0) = ""
            If .FixedCols > 0 Then
               .col = .FixedCols - 1
               .CellBackColor = .BackColorFixed
               .CellForeColor = .ForeColor
            End If
            For iCol = .FixedCols To .Cols - 1
               .col = iCol
               .CellBackColor = .BackColor
            Next
         End If
      End If
   
      If nRow > 0 Then
         .row = nRow
         .TextMatrix(nRow, 0) = "V"
         If .FixedCols > 0 Then
            .col = .FixedCols - 1
            .CellBackColor = .BackColorSel
            .CellForeColor = .ForeColorSel
         End If
         For iCol = .FixedCols To .Cols - 1
           .col = iCol
           .CellBackColor = &HFFC0C0
         Next
      End If
      .col = nCol
      pPrevRow = nRow
      Call SetCombo4Data(.TextMatrix(nRow, 1))
      'Add By Sindy 2015/3/5
      If .TextMatrix(nRow, 1) = "" Then
         txtInvField(0) = .TextMatrix(nRow, 2)
         txtInvField(1) = .TextMatrix(nRow, 3)
         txtInvField(2) = .TextMatrix(nRow, 4)
         txtIN11 = .TextMatrix(nRow, 5)
         cmdUpdRow.Enabled = True
         cmdAddRow.Enabled = False
      End If
      '2015/3/5 END
   End If
   .Visible = True
   End With
End Sub

'Add By Sindy 2014/11/11
Private Sub SetCombo4Data(ByVal strData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To Combo4.ListCount - 1
      'Modify By Sindy 2015/12/4
      'If Combo4.List(nPos) = strData Then
      If InStr(Combo4.List(nPos), strData) > 0 Then
      '2015/12/4 END
         bFind = True
         Exit For
      End If
   Next nPos
   If Not bFind Then
      Combo4.AddItem strData
      Combo4.Refresh
      Combo4.ListIndex = Combo4.ListCount - 1
   Else
      Combo4.ListIndex = nPos
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If PreviousTab = 0 Then
   'Add by Lydia 2014/10/27 日文定稿頁籤切換提示發明人語系
    If m_LetterLanguage = "3" Then
       MsgBox "本案為日文定稿，請輸入日文發明人名稱。"
    End If
End If
End Sub

Private Sub Text10_GotFocus(Index As Integer)
  TextInverse Text10(Index)
End Sub

Private Sub Text10_Validate(Index As Integer, Cancel As Boolean)
   If Index = 2 Then
      If Text10(0) = "" And Text10(1) = "" And Text10(2) = "" Then
         MsgBox "案件名稱不可同時空白 !", vbCritical
      End If
   End If
End Sub

Private Sub Text33_GotFocus(Index As Integer)
  TextInverse Text33(Index)
End Sub

Private Sub Text33_Validate(Index As Integer, Cancel As Boolean)
   'Added by Lydia 2017/06/14 設欄位長度
    Dim iLen As Integer
    Select Case Index
    Case 0, 3 '專利-聯絡人中文
         iLen = 30
    Case 1, 4 '聯絡人英文
         iLen = 35
    Case 2, 5, 6 '聯絡人日文
         iLen = 60
    Case Else
         iLen = Text33(Index).MaxLength
    End Select
    'end 2017/06/14
    
   'Modified by Lydia 2017/06/14
   'If Not CheckLengthIsOK(Text33(Index), Text33(Index).MaxLength) Then
   If Not CheckLengthIsOK(Text33(Index), iLen) Then
      Cancel = True
   End If
End Sub

Private Sub Text6_GotFocus()
   If Len(Text6) > 1 Then
      TextInverse Text6
   End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = "" Then
      MsgBox "專利號數不可空白 !"
      Cancel = True
   'Modified by Morgan 2012/1/4 積體電路案件例外--David
   ElseIf m_bolIs117 = False Then
   
      'Modify by Morgan 2004/8/4
      '新法證書號改7碼
      If m_bolNew = True Then
         If Not Len(Text6.Text) = 7 Then
            MsgBox "輸入的專利號數錯誤"
            Cancel = True
            
         '與公報輸入做雙重檢查
         ElseIf Text6.Text <> pa(22) And pa(22) <> "" Then
            MsgBox "專利號數應為【" & pa(22) & "】！", vbCritical
            Cancel = True
         'Add by Morgan 2004/8/18 帶出公告日--靜芳
         ElseIf txtPA14.Text = "" Then
            txtPA14.Text = pa(14)
            'Add by Morgan 2005/3/1 帶出公告日--靜芳
            Text7.Text = txtPA14.Text
         End If
      ElseIf Len(Text6) <> 6 Then
         MsgBox "專利號數必須為六位數字 !"
         Cancel = True
      End If
   End If
   If Cancel = True Then Text6_GotFocus
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 = "" Then
      MsgBox "發證日不可空白，請重新輸入 !", vbCritical
      Cancel = True
   Else
      If Not ChkDate(Text7) Or Val(Text7) > Val(strSrvDate(2)) Then
         MsgBox "發證日不可大於系統日 !", vbCritical
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text7
End Sub

Private Sub Text8_GotFocus(Index As Integer)
  TextInverse Text8(Index)
End Sub


Private Sub Text8_Validate(Index As Integer, Cancel As Boolean)
 Dim strTmp(0 To 4) As String, strTmp1(0 To 5) As String, i As Integer
   
   'Add By Cheng 2001/12/31
   m_bln_keyinValidate = False

   If Not ChkDate(Text8(Index)) Then
      DoEvents
      Text8(Index).SetFocus
      Text8_GotFocus (Index)
      Cancel = True
      Exit Sub
   Else
   
      'Modify by Morgan 2004/8/4
      '改在一開始時抓一次就好
'      If Index = 1
'         For i = 1 To 4
'            strTmp1(i) = pa(i)
'         Next
'         If GetMoneyDate(Val(pa(8)) + 10, pa(9), strTmp1, strTmp(1), strTmp(2), strTmp(3)) = True Then
'            If Text8(0) <> TransDate(strTmp(1), 1) Then
'            MsgBox "專利期間輸入錯誤 !", vbCritical
      
      If Index = 0 Then
         If DATE1 <> "" Then
            'Modify by Morgan 2004/10/7 專用期改西元
            'If Text8(0) <> TransDate(DATE1, 1) Then
            '   MsgBox "專利期間應為【" & TransDate(DATE1, 1) & "】!", vbCritical
            If Text8(0) <> TransDate(DATE1, 2) Then
               'Modified by Morgan 2012/1/4 積體電路案件例外--David
               If m_bolIs117 = True Then
                  If MsgBox("確定專用期限起日日期不為 " & TransDate(DATE1, 2) & " (申請日起算)？", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                     Text8(0).SetFocus
                     Cancel = True
                     Exit Sub
                  Else
                     DATE1 = TransDate(Text8(0), 2)
                  End If
               Else
                  MsgBox "專利期間應為【" & TransDate(DATE1, 2) & "】!", vbCritical
                  Text8(0).SetFocus
                  Cancel = True
                  Exit Sub
               End If
            End If
         End If
         
      Else
      
         'Modify By Cheng 2001/12/31
         'Modify by Morgan 2012/1/4 申請案號改碼數
         'If Len(Me.Text1.Text) <= 8 Then
         If Len(Me.Text1.Text) <= 9 Then
            'Modify by Morgan 2004/8/4
            'If Text8(1) <> TransDate(strTmp(3), 1) Then
            '  MsgBox "專利期間輸入錯誤 !", vbCritical
            
            'Modify by Morgan 2004/10/7 專用期改西元
            'If Text8(1) <> TransDate(DATE2, 1) Then
            '    MsgBox "專利期間應為【" & TransDate(DATE2, 1) & "】!", vbCritical
            If Text8(1) <> TransDate(DATE2, 2) Then
            
               'Modified by Morgan 2012/1/4 積體電路案件例外--David
               If m_bolIs117 = True Then
                  If MsgBox("確定專用期限迄日日期不為 " & TransDate(DATE2, 2) & " (申請日起算)？", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                     Text8(1).SetFocus
                     Cancel = True
                     Exit Sub
                  Else
                     DATE2 = TransDate(Text8(1), 2)
                  End If
               Else
            
                  MsgBox "專利期間應為【" & TransDate(DATE2, 2) & "】!", vbCritical
                  
                  Text8(1).SetFocus
                  Text8_GotFocus 1
                  Cancel = True
                  Exit Sub
                  
               End If
               
             End If
             
         'Add By Cheng 2001/12/31
         '若申請案為追加, 聯合案(A01,U01)其專案截止日必須等於母案的專案截止日
         'Modify by Morgan 2010/12/27 申請案號改碼數
         'ElseIf Len(Me.Text1.Text) > 8 Then
         ElseIf Len(Me.Text1.Text) > 9 Then
            'Modify by Morgan 2010/12/27 申請案號改碼數
            strExc(0) = "SELECT PA25,PA01||'-'||PA02||decode(PA03||PA04,'000','','-'||PA03||'-'||PA04) MNo FROM PATENT " & _
                        " WHERE PA11='" & Left(pa(11), 9) & "' " & _
                        " AND PA09 = '" & pa(9) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Not RsTemp.EOF Then
                  'Added by Morgan 2018/5/14 檢查若母案無專用期時提醒 EX.FCP-056630
                  If IsNull(RsTemp.Fields(0).Value) Then
                     MsgBox "母案 ( " & RsTemp("MNo") & " ) 尚無專利期間，請先輸入!!", vbCritical
                     Text8(1).SetFocus
                     Text8_GotFocus 1
                     If RsTemp.State <> adStateClosed Then RsTemp.Close
                     Set RsTemp = Nothing
                     Cancel = True
                     Exit Sub
                  
                  'end 2018/5/14
                  'Modify by Morgan 2004/10/7 專用期改西元
                  'If Me.Text8(1).Text <> TransDate(rsTemp.Fields(0).Value, 1) Then
                  ElseIf Me.Text8(1).Text <> TransDate(RsTemp.Fields(0).Value, 2) Then
                     MsgBox "專利期間輸入錯誤 !", vbCritical
                     Text8(1).SetFocus
                     Text8_GotFocus 1
                     If RsTemp.State <> adStateClosed Then RsTemp.Close
                     Set RsTemp = Nothing
                     Cancel = True
                     Exit Sub
                  End If
               End If
            End If
            If RsTemp.State <> adStateClosed Then RsTemp.Close
            Set RsTemp = Nothing
         End If
         
      End If
   End If
   
   'Add By Cheng 2001/12/31
   m_bln_keyinValidate = True

End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
'20140306START ADD By eric 增加發明人
Dim strNo As String, strSql As String
Dim strk1(1 To 10) As String, strk2(1 To 10) As String, strk3(1 To 10) As String
Dim j As Integer, k1 As Integer, k2 As Integer, k3 As Integer
Dim rsTmp  As New ADODB.Recordset
Dim strInv(0 To 10) As String
'20140306END
'Add by Lydia 2014/10/27 日文定稿頁籤切換提示發明人語系
'移到表頭
'Dim m_LetterLanguage As String   'add by sonia 2014/4/22

   TxtValidate = False
   For Each objTxt In Text10
      If objTxt.Enabled = True Then
         Cancel = False
         Text10_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   If Me.Text6.Enabled = True Then
      Cancel = False
      Text6_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.Text7.Enabled = True Then
      Cancel = False
      Text7_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   For Each objTxt In Text8
      If objTxt.Enabled = True Then
         Cancel = False
         Text8_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   If txtPA14.Visible = True And txtPA14.Enabled = True Then
      Cancel = False
      txtPA14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   strNo = pa(26)
   If pa(27) <> "" Then strNo = strNo & "," & pa(27)
   If pa(28) <> "" Then strNo = strNo & "," & pa(28)
   If pa(29) <> "" Then strNo = strNo & "," & pa(29)
   If pa(30) <> "" Then strNo = strNo & "," & pa(30)
   If Text2 = "FCP" Or Text2 = "P" Or Text2 = "CFP" Then
'      Cancel = False
      'add by sonia 2014/4/22
'      m_LetterLanguage = PUB_GetLanguage(Text2, Text3, Text4, Text5)
'      If txtInvField(0) = "" And txtInvField(1) = "" And txtInvField(2) = "" Then
      'Modify By Sindy 2014/12/23
      'If GRD1.Rows = 2 And GRD1.TextMatrix(1, 1) = "" Then
      If GRD1.Rows = 2 And GRD1.TextMatrix(1, 1) = "" And GRD1.TextMatrix(1, 6) = "" Then
      '2014/12/23 END
         MsgBox "發明人不可空白 !", vbCritical
         Combo4.SetFocus
         Exit Function
      End If
      '2014/4/22 end
      'Modify by Sindy 2014/11/11 檢查發明人
      For ii = 1 To GRD1.Rows - 1
         strExc(0) = Trim(GRD1.TextMatrix(ii, 1))
         If strExc(0) <> "" Then
            If PUB_ChkInventor(strExc(0), strNo) = False Then
                If MsgBox("發明人(" & strExc(0) & ")資料與目前申請人不符，是否要繼續！", vbYesNo + vbDefaultButton2) = vbNo Then
                   SSTab1.Tab = 1
                   Exit Function
                End If
            End If
         Else
            If Trim(GRD1.TextMatrix(ii, 6)) <> "" Then
               If Trim(GRD1.TextMatrix(ii, 6)) <> pa(26) & String(8 - Len(Left(pa(26), 8)), "0") Then
                  If MsgBox("第 " & ii & " 筆發明人資料與目前申請人不符，是否要繼續！", vbYesNo + vbDefaultButton2) = vbNo Then
                     SSTab1.Tab = 1
                     Exit Function
                  End If
               End If
            End If
         End If
      Next ii
      '2014/11/11 END
      
'      j = 1: k1 = 1: k2 = 1: k3 = 1
'      For ii = 0 To 9
'         strExc(0) = Replace(Right(Combo4(ii).Text, 11), ")", "")
'         If PUB_ChkInventor(strExc(0), strNo) = False Then
'            If MsgBox("發明人資料與目前申請人不符，是否要繼續！", vbYesNo + vbDefaultButton2) = vbNo Then
'               SSTab1.Tab = 2
'               Combo4(ii).SetFocus
'               Exit Function
'            End If
'         End If
'
'         If strExc(0) <> "" Then
'            strExc(j) = strExc(0)
'            j = j + 1
'            'add by sonia 2014/4/22
'            If m_LetterLanguage = "3" And txtInvField(ii * 3 + 2) = "" Then
'               If MsgBox("定稿語文為日文, 發明人 " & ii + 1 & " 的日文空白, 若需要日文請選 是 並自行至客戶發明人資料維護補輸日文, 不需要日文請選 否 !", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                  txtInvField(ii * 3 + 2).SetFocus
'                  TextInverse txtInvField(ii * 3 + 2)
'                  Exit Function
'               End If
'            End If
'            '2014/4/22 end
'         ElseIf strExc(0) = "" And (txtInvField(ii * 3) <> "" Or txtInvField(ii * 3 + 1) <> "" Or txtInvField(ii * 3 + 2) <> "") Then
'            If txtInvField(ii * 3) <> "" Then strk1(k1) = Trim(txtInvField(ii * 3)): k1 = k1 + 1
'            If txtInvField(ii * 3 + 1) <> "" Then strk2(k2) = Trim(txtInvField(ii * 3 + 1)): k2 = k2 + 1
'            If txtInvField(ii * 3 + 2) <> "" Then strk3(k3) = Trim(txtInvField(ii * 3 + 2)): k3 = k3 + 1
'
'            '判斷客戶發明人檔是否有重覆資料
'            'Modified by Morgan 2014/4/22
'            'strNo = strNo & String(8 - Len(pa(26)), "0")
'            strNo = strNo & String(8 - Len(Left(pa(26), 8)), "0")
'            'Modified by Morgan 2014/5/2 發明人會有造字無法存檔時會加空白,所以改在語法內trim
'            'strSql = "Select * From Inventor Where IN01=" & CNULL(strNo) & " and (IN04='" + ChgSQL(Trim$(txtInvField(ii * 3))) & "'" & _
'                    " OR IN05='" & ChgSQL(Trim$(txtInvField(ii * 3 + 1))) & "' OR IN06='" & ChgSQL(Trim$(txtInvField(ii * 3 + 2))) & "')"
'            strSql = "Select * From Inventor Where IN01=" & CNULL(strNo) & " and (rtrim(IN04)=rtrim('" + ChgSQL(txtInvField(ii * 3)) & "')" & _
'                     " OR rtrim(IN05)=rtrim('" & ChgSQL(txtInvField(ii * 3 + 1)) & "') OR rtrim(IN06)=rtrim('" & ChgSQL(txtInvField(ii * 3 + 2)) & "'))"
'            'end 2014/5/2
'            Set rsTmp = ClsPDReadRst(strSql)
'
'            If Not rsTmp.EOF Then
'               If Trim(txtInvField(ii * 3)) = Trim(rsTmp.Fields("IN04")) Then
'                  If MsgBox("發明人名稱中文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                     rsTmp.Close
'                     txtInvField(ii * 3).SetFocus
'                     TextInverse txtInvField(ii * 3)
'                     Exit Function
'                  End If
'               End If
'               If Trim(txtInvField(ii * 3 + 1)) = Trim(rsTmp.Fields("IN05")) Then
'                  If MsgBox("發明人名稱英文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                     rsTmp.Close
'                     txtInvField(ii * 3 + 1).SetFocus
'                     TextInverse txtInvField(ii * 3 + 1)
'                     Exit Function
'                  End If
'               End If
'               If Trim(txtInvField(ii * 3 + 2)) = Trim(rsTmp.Fields("IN06")) Then
'                  If MsgBox("發明人名稱日文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                     rsTmp.Close
'                     txtInvField(ii * 3 + 2).SetFocus
'                     TextInverse txtInvField(ii * 3 + 2)
'                     Exit Function
'                  End If
'               End If
'
'               '判斷國籍是否有輸入
'               If txtIN11(ii) = "" Then
'                  MsgBox ("請輸入國藉")
'                  If ii > 4 Then
'                     SSTab1.Tab = 2
'                  Else
'                     SSTab1.Tab = 1
'                  End If
'                  txtIN11(ii).SetFocus
'                  Exit Function
'               Else
'                  Cancel = False
'                  txtIN11_Validate ii, Cancel
'                  If Cancel = True Then
'                     If ii > 4 Then
'                        SSTab1.Tab = 2
'                     Else
'                        SSTab1.Tab = 1
'                     End If
'                     txtIN11(ii).SetFocus
'                     TextInverse txtIN11(ii)
'                     Exit Function
'                  End If
'               End If
'
'            End If
'
'            'add by sonia 2014/4/22
'            If m_LetterLanguage = "3" And txtInvField(ii * 3 + 2) = "" Then
'               If MsgBox("定稿語文為日文, 發明人 " & ii + 1 & " 的日文空白, 是否要輸入日文名稱 ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                  txtInvField(ii * 3 + 2).SetFocus
'                  TextInverse txtInvField(ii * 3 + 2)
'                  Exit Function
'               End If
'            End If
'            '2014/4/22 end
'         End If
'      Next
'
'      Sort strExc, j - 2
'      For ii = 1 To j - 2
'         If strExc(ii) = strExc(ii + 1) Then
'            Cancel = True
'            Exit For
'         End If
'      Next
'      If Cancel Then
'         MsgBox "發明人不可重覆 !", vbCritical
'         Combo4(0).SetFocus
'         Exit Function
'      End If
'
'      Sort strk1, k1 - 2
'      For ii = 1 To k1 - 2
'         If strk1(ii) = strk1(ii + 1) Then
'            Cancel = True
'            Exit For
'         End If
'      Next
'      If Cancel Then
'         MsgBox "發明人中文名稱不可重覆 !", vbCritical
'         Combo4(0).SetFocus
'         Exit Function
'      End If
'
'      Sort strk2, k2 - 2
'      For ii = 1 To k2 - 2
'         If strk2(ii) = strk2(ii + 1) Then
'            Cancel = True
'            Exit For
'         End If
'      Next
'      If Cancel Then
'         MsgBox "發明人英文名稱不可重覆 !", vbCritical
'         Combo4(0).SetFocus
'         Exit Function
'      End If
'
'      Sort strk3, k3 - 2
'      For ii = 1 To k3 - 2
'         If strk3(ii) = strk3(ii + 1) Then
'            Cancel = True
'            Exit For
'         End If
'      Next
'      If Cancel Then
'         MsgBox "發明人日文名稱不可重覆 !", vbCritical
'         Combo4(0).SetFocus
'         Exit Function
'      End If
   End If
   '20140306END
   
   'Added by Morgan 2023/2/15 電子證書另存至 \\TYPING2\FCP_workflow\Patent Certificate
   If m_DocNo <> "" Then
      strExc(1) = "$" & m_DocNo & ".CERT.pdf"
      'Modified by Lydia 2024/07/22 改用變數
      'strExc(2) = "\\TYPING2\FCP_workflow\Patent Certificate\" & PUB_CaseNo2FileName(pa(1), pa(2), pa(3), pa(4)) & "Patent Certificate.pdf"
      strExc(2) = "\\" & strTyping2Path & "\FCP_workflow\Patent Certificate\" & PUB_CaseNo2FileName(pa(1), pa(2), pa(3), pa(4)) & "Patent Certificate.pdf"
      intI = 1
      If Dir(strExc(2)) <> "" Then
         'Modified by Lydia 2024/07/22 改用變數
         'If MsgBox("[ \\TYPING2\FCP_workflow\Patent Certificate ]電子證書已存在！" & vbCrLf & "是否要重新下載？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         If MsgBox("[ \\" & strTyping2Path & "\FCP_workflow\Patent Certificate ]電子證書已存在！" & vbCrLf & "是否要重新下載？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
            intI = 0
         End If
      End If
      If intI = 1 Then
         If PUB_GetAttachFile_CPP(m_DocNo, strExc(1), strExc(2), True) = False Then
            'Modified by Lydia 2024/07/22 改用變數
            'If MsgBox("電子證書下載至[ \\TYPING2\FCP_workflow\Patent Certificate ]失敗！" & vbCrLf & "是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
            If MsgBox("電子證書下載至[ \\" & strTyping2Path & "\FCP_workflow\Patent Certificate ]失敗！" & vbCrLf & "是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            End If
         End If
      End If
   End If
   'end 2023/2/15
         
   TxtValidate = True

End Function

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

'Private Sub txtInvField_Change(Index As Integer)
'If Len(Combo4.Text) = 0 Then
'   nResponse = Empty 'Modified by Lydia 2015/01/06
'End If
'End Sub

Private Sub txtPA14_GotFocus()
   TextInverse txtPA14
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtPA14.IMEMode = 2
   CloseIme
End Sub

Private Sub txtPA14_KeyPress(KeyAscii As Integer)
   '只能key倒退鍵和數字
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtPA14_Validate(Cancel As Boolean)
   If txtPA14.Text = "" Then
      MsgBox "公告日不可空白！", vbCritical
      Cancel = True
   ElseIf Not ChkDate(txtPA14) Then
      MsgBox "日期格式錯誤！", vbCritical
      Cancel = True
   ElseIf pa(14) <> "" Then
      If Val(txtPA14) <> Val(pa(14)) Then
         MsgBox "公告日應為【 " & ChangeTStringToTDateString(pa(14)) & " 】！", vbCritical
         txtPA14_GotFocus
         Cancel = True
      End If
   End If
   '若用新法則專用期起日=公告日
   If Cancel = False And m_bolNew = True Then
      'Modify by Morgan 2004/10/7 專用期改西元
      'Text8(0).Text = txtPA14.Text
      Text8(0).Text = TransDate(txtPA14.Text, 2)
      
      DATE1 = TransDate(Text8(0).Text, 2)
      m_strNextDueDate = CompDate(0, m_intLastYear, DATE1)
      m_strNextDueDate = CompDate(2, -1, m_strNextDueDate)
      'Modified by Morgan 2014/11/20 外專改回舊規則
      ''Added by Morgan 2014/10/29
      'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
      '   m_strNextFeeDate = PUB_GetOurDeadline(m_strNextDueDate)
      'Else
      ''end 2014/10/29
      
      'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
      If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
         'Modify By Sindy 2021/8/17 + , , m_strAgreeOnDate
         m_strNextFeeDate = PUB_GetFCPOurDeadline(m_strNextDueDate, 2, , m_strAgreeOnDate)
      Else
      'end 2019/7/11
      
         m_strNextFeeDate = CompDate(2, -2, m_strNextDueDate)
         
      End If 'Added by Morgan 2019/7/11
      
      'End If 'Added by Morgan 2014/10/29
      'end 2014/11/20
      
      '下次繳費日
      Label2(7).Caption = TransDate(m_strNextDueDate, 1)
   End If
End Sub

'20140306START ADD By eric
Private Sub txtIN11_Validate(Cancel As Boolean)
   If txtIN11.Text = "" Then Exit Sub 'Add By Sindy 2015/12/4
   If Val(txtIN11) >= 1 And Val(txtIN11) <= 8 Then
      MsgBox ("發明人國籍不可輸入 001 - 008")
      Me.Lb_IN11N.Caption = ""
      Cancel = True
   Else
      If ClsPDGetNation(txtIN11, strName) Then
         Me.Lb_IN11N.Caption = strName
      Else
         Me.Lb_IN11N.Caption = ""
         Cancel = True
      End If
   End If
End Sub

'20140306START ADD By eric
Private Sub txtInvField_GotFocus(Index As Integer)
    If Combo4 <> "" Then
'        Lb_IN11.Visible = False
'        txtIN11.Visible = False
'        Lb_IN11N.Visible = False
        Combo4.SetFocus
    Else
'        Lb_IN11.Visible = True
'        txtIN11.Visible = True
'        Lb_IN11N.Visible = True
        If txtInvField(0) = "" And txtInvField(1) = "" And txtInvField(2) = "" Then
            txtIN11.Text = ""
            Lb_IN11N.Caption = ""
        End If
    End If
End Sub

'20140306START ADD By eric
Private Sub txtInvField_LostFocus(Index As Integer)
    If Combo4 = "" And txtInvField(0) = "" And txtInvField(1) = "" And txtInvField(2) = "" Then
'        Lb_IN11.Visible = False
'        txtIN11.Visible = False
'        Lb_IN11N.Visible = False
    End If
'Add by Lydia 2014/10/22 控制移到 txtInvField_Validate
'    Select Case Index Mod 3
'        Case 2
'            If idx <= 3 Then
'                Combo4(idx + 1).SetFocus
'            End If
'    End Select
End Sub

'Add by Lydia 2014/10/22 發明人輸入比對兼自動代入(模糊比對)
Private Sub txtInvField_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim tRec As Integer, tSearch As Boolean
   Dim tInx As Integer, tSno As Integer 'tInx =combo4(index), tSno=List編號
   'Dim nResponse 'Modified by Lydia 2015/01/06
   Cancel = False
'   Dim menuStr() As String 'Modified by Lydia 2015/01/06
   If IsEmptyText(txtInvField(Index)) = False Then
      If StrLength(txtInvField(Index)) > txtInvField(Index).MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "發明人名稱太長"
        ' nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         MsgBox strMsg, vbOKOnly + vbCritical, strTit
      Else
            'Modified by Lydia 2015/01/06 改為選擇即有發明人或新增發明人->淑華表示用自動帶,若遇到名字相同唯寫法不同,到維護畫面採人工新增
'         If nResponse = Empty Then
'          ReDim menuStr(0 To 1)
         For tRec = 0 To m_InventorListCount - 1
        '    tSno = 0 'Modified by Lydia 2015/01/06
   '            Select Case Index Mod 3
            If Index = 0 Then '(發明人)中文名稱
               If InStr(m_InventorList(tRec).iN04, txtInvField(Index)) > 0 Then
                 tSearch = True: tInx = Index \ 3: tSno = tRec
                 Cancel = True
                 Exit For
               End If
               
            ElseIf Index = 1 Then '(發明人)英文名稱
               If InStr(UCase(m_InventorList(tRec).IN05), UCase(txtInvField(Index))) > 0 Then
                 tSearch = True: tInx = Index \ 3: tSno = tRec
                 Cancel = True
                 Exit For
               End If
               
            ElseIf Index = 2 Then '(發明人)日文名稱
               If InStr(m_InventorList(tRec).IN06, txtInvField(Index)) > 0 Then
                 tSearch = True: tInx = Index \ 3: tSno = tRec
                 Cancel = True
                 Exit For
               End If
            End If
            
'            If tSno <> 0 Then  'Modified by Lydia 2015/01/06
'                 menuStr(0) = menuStr(0) & Format(tSno, "000") & "," 'List編號
'                 menuStr(1) = menuStr(1) & Trim(m_InventorList(tRec).IN02) & "  "
'                 menuStr(1) = menuStr(1) & Trim(m_InventorList(tRec).IN04) & "  "
'                 menuStr(1) = menuStr(1) & Trim(m_InventorList(tRec).IN05) & vbCrLf
'            End If
            
         Next tRec
'         End If
      End If
   End If
   
   If Cancel = False Then
      CloseIme
'      If Index Mod 3 = 2 And Index <> 14 Then
'        Combo4((Index \ 3) + 1).SetFocus '跳到下一個發明人combo List
'      End If
   Else
      If tSearch = True Then
       'Modified by Lydia 2015/01/06
'       menuStr(1) = menuStr(1) & "0      新增發明人"
'       nResponse = InputBox(menuStr(1), "請選擇輸入", "0")
'       If nResponse <> "0" And Not (nResponse = Empty) Then
         Combo4.ListIndex = tSno + 1 '讀發明人List=>call combo4_click
        ' Combo4.ListIndex = Val(nResponse)
         Combo4.SetFocus  '移到比對出的發明人combo List
       'End If
      End If
   End If
End Sub
'Added by Morgan 2018/5/28 整合
'提示訊息
Private Sub ShowPrompt()
   'Added by Morgan 2016/2/3
   If Left(pa(75), 8) = "Y4829203" Then
      'Modified by Morgan 2018/11/19 --Winfrey
      'MsgBox "至 HP 平台輸入相關資料!!", vbExclamation
      MsgBox "證書請早上優先輸入平台資料後，當天上傳證書。", vbExclamation
      'end 2018/11/19
   'Added by Morgan 2016/9/30--鄭詠心
   'Modify By Sindy 2018/3/2 取消 Y4580901, 新增 X2166000
'   ElseIf Left(pa(75), 8) = "Y4580901" Then
   'Modified by Morgan 2018/5/18 +Y4905300
   'Modified by Morgan 2018/5/28 +FCP056723 --何淑華
   'Modified by Morgan 2018/6/28 +FCP056721 --何淑華
   'Modified by Morgan 2018/7/27 +FCP051418 --何淑華
   'Modified by Morgan 2018/8/13 +FCP056314 --何淑華
   'Modified by Morgan 2018/11/16 +FCP050921 --何淑華
   ElseIf InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X2166000") > 0 Or _
      Left(ChangeCustomerL(pa(75)), 8) = "Y4905300" Or _
      pa(1) & pa(2) = "FCP056723" Then
      MsgBox "優先退Winfrey處理證書(卷上別條子註明)", vbExclamation
   ElseIf InStr("FCP056721,FCP051418,FCP056314,FCP050921", pa(1) & pa(2)) > 0 Then
      MsgBox "證書及公報優先處理", vbExclamation
      
   'Added by Morgan 2018/11/30--陳亭妙
   'Modified by Morgan 2018/12/19 +FCP-058259--陳亭妙
   'Modified by Morgan 2019/1/9 +FCP055018--陳亭妙
   'Modified by Morgan 2019/1/18 +FCP058272--陳亭妙
   ElseIf pa(1) & pa(2) = "FCP052005" Or pa(1) & pa(2) = "FCP058259" Or pa(1) & pa(2) = "FCP055018" Or pa(1) & pa(2) = "FCP058272" Then
      MsgBox "「請證書優先交Winfrey寄出，謝謝。」", vbExclamation
   End If
   'end 2016/2/3
   
   'add by sonia 2016/9/22
   If pa(75) = "Y52643" Then
      MsgBox "優先寄證書(不可延誤寄出)，且證書須E+寄！"
   'Add By Sindy 2017/9/14
   ElseIf pa(75) = "Y34210" Or pa(75) = "Y3421003" Then
      MsgBox "請直接交Winfrey並寫紙條：7日內優先寄出！"
   'Added by Morgan 2018/11/13 --Winfrey
   ElseIf pa(1) & pa(2) = "FCP050443" Then
      MsgBox pa(91), vbExclamation
   'end 2018/11/13
   'Add By Sindy 2018/11/21
   '當代理人為Metis IP (Beijing) LLC (Y54339B10), Metis IP (Suzhou) LLC (Y54339B20), METIS IP LLC (Y54339)的案件：
   '當程序在證書號數輸入完成時，請設定彈跳訊息「注意：專利證書正本需寄蘇州分部Metis IP (Suzhou) LLC (Y54339B20)」。
   ElseIf Left(pa(75), 6) = "Y54339" Then
      MsgBox "注意：專利證書正本需寄蘇州分部Metis IP (Suzhou) LLC (Y54339B20)"
   End If
End Sub
