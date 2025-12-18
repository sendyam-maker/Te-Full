VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_7 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務基本資料（條碼）"
   ClientHeight    =   6090
   ClientLeft      =   780
   ClientTop       =   980
   ClientWidth     =   8070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8070
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   6
      Left            =   2803
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "各項指示"
      Height          =   400
      Index           =   5
      Left            =   780
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "相關卷號"
      Height          =   400
      Index           =   4
      Left            =   1744
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人資料"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   3627
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   0
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6275
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請人資料"
      Height          =   400
      Index           =   0
      Left            =   4951
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   0
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7200
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   0
      Width           =   800
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   5280
      Left            =   90
      TabIndex        =   7
      Top             =   450
      Width           =   7935
      _ExtentX        =   13988
      _ExtentY        =   9313
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   " 基本資料"
      TabPicture(0)   =   "frm100101_7.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label92"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label91"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label84"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label89"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label88"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label19"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label17"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label12"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label8"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label10"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label9"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl1(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl1(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lbl1(4)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbl1(5)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl1(6)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbl1(7)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lbl1(10)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lbl1(11)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lbl1(8)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lbl1(9)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label13"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lbl1(27)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lbl1(85)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label113"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label112"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lbl1(87)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label22"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lbl1(86)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label25"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txt1(0)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txt1(1)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt1(2)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt1(3)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txt1(4)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txt1(5)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt1(7)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txt1(8)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).ControlCount=   47
      TabCaption(1)   =   "代理人相關資料"
      TabPicture(1)   =   "frm100101_7.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label29"
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(4)=   "Label18"
      Tab(1).Control(5)=   "Label20"
      Tab(1).Control(6)=   "Label24"
      Tab(1).Control(7)=   "Label28"
      Tab(1).Control(8)=   "Label31"
      Tab(1).Control(9)=   "lbl1(13)"
      Tab(1).Control(10)=   "lbl1(14)"
      Tab(1).Control(11)=   "lbl1(15)"
      Tab(1).Control(12)=   "lbl1(16)"
      Tab(1).Control(13)=   "lbl1(17)"
      Tab(1).Control(14)=   "lbl1(18)"
      Tab(1).Control(15)=   "lbl1(19)"
      Tab(1).Control(16)=   "lbl1(12)"
      Tab(1).Control(17)=   "lbl1(23)"
      Tab(1).Control(18)=   "Label21"
      Tab(1).Control(19)=   "lbl1(84)"
      Tab(1).Control(20)=   "Label14"
      Tab(1).Control(21)=   "txt1(6)"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "正片號碼"
      TabPicture(2)   =   "frm100101_7.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdDataList1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "銷卷資料"
      TabPicture(3)   =   "frm100101_7.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label81"
      Tab(3).Control(1)=   "Label80"
      Tab(3).Control(2)=   "Label79"
      Tab(3).Control(3)=   "Label78"
      Tab(3).Control(4)=   "lbl1(22)"
      Tab(3).Control(5)=   "lbl1(24)"
      Tab(3).Control(6)=   "lbl1(25)"
      Tab(3).Control(7)=   "lbl1(26)"
      Tab(3).ControlCount=   8
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList1 
         Height          =   4560
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   6492
         _ExtentX        =   11448
         _ExtentY        =   8043
         _Version        =   393216
         Rows            =   101
         FixedCols       =   0
         HighLight       =   0
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
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   8
         Left            =   1050
         TabIndex        =   81
         Top             =   1950
         Width           =   1785
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3149;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   7
         Left            =   1290
         TabIndex        =   80
         Top             =   270
         Width           =   2265
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3995;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   1110
         Index           =   6
         Left            =   -74880
         TabIndex        =   14
         Top             =   3600
         Width           =   7695
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13573;1958"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   525
         Index           =   5
         Left            =   1050
         TabIndex        =   13
         Top             =   4380
         Width           =   6735
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11880;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   525
         Index           =   4
         Left            =   1050
         TabIndex        =   12
         Top             =   3840
         Width           =   6735
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11880;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   3
         Left            =   4890
         TabIndex        =   11
         Top             =   2527
         Width           =   2895
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5106;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   2
         Left            =   1290
         TabIndex        =   10
         Top             =   1290
         Width           =   6495
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11456;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   1
         Left            =   1290
         TabIndex        =   9
         Top             =   930
         Width           =   6495
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11456;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   0
         Left            =   1290
         TabIndex        =   8
         Top             =   570
         Width           =   6495
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11456;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "國內副本收件人："
         Height          =   255
         Left            =   60
         TabIndex        =   88
         Top             =   3570
         Width           =   1440
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   86
         Left            =   1545
         TabIndex        =   87
         Top             =   3570
         Width           =   3255
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "國內副本接洽人："
         Height          =   255
         Left            =   4935
         TabIndex        =   86
         Top             =   3570
         Width           =   1440
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   87
         Left            =   6450
         TabIndex        =   85
         Top             =   3570
         Width           =   1365
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2408;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         Caption         =   "(J:智權公司 空白:系統預設)"
         Height          =   180
         Left            =   1680
         TabIndex        =   84
         Top             =   4920
         Width           =   2115
      End
      Begin VB.Label Label113 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司："
         Height          =   180
         Left            =   60
         TabIndex        =   83
         Top             =   4920
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   85
         Left            =   1350
         TabIndex        =   82
         Top             =   4920
         Width           =   225
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "397;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID："
         Height          =   255
         Left            =   -74910
         TabIndex        =   79
         Top             =   960
         Width           =   1860
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   84
         Left            =   -72990
         TabIndex        =   78
         Top             =   960
         Width           =   5400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "9525;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   27
         Left            =   5745
         TabIndex        =   77
         Top             =   3300
         Width           =   2085
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3678;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   255
         Left            =   4950
         TabIndex        =   76
         Top             =   3300
         Width           =   720
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   26
         Left            =   -73560
         TabIndex        =   75
         Top             =   1320
         Width           =   6255
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11033;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   25
         Left            =   -73560
         TabIndex        =   74
         Top             =   1020
         Width           =   1005
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1773;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   24
         Left            =   -73560
         TabIndex        =   73
         Top             =   720
         Width           =   1000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1764;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   22
         Left            =   -73560
         TabIndex        =   72
         Top             =   420
         Width           =   1000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1764;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷備註："
         Height          =   255
         Left            =   -74850
         TabIndex        =   71
         Top             =   1320
         Width           =   1260
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   255
         Left            =   -74850
         TabIndex        =   70
         Top             =   1020
         Width           =   1110
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   255
         Left            =   -74850
         TabIndex        =   69
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   255
         Left            =   -74850
         TabIndex        =   68
         Top             =   420
         Width           =   1110
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "D/N固定列印對象："
         Height          =   255
         Left            =   -74910
         TabIndex        =   67
         Top             =   2970
         Width           =   1545
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   -73290
         TabIndex        =   66
         Top             =   2970
         Width           =   5715
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10081;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   12
         Left            =   -73920
         TabIndex        =   65
         Top             =   390
         Width           =   6690
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11800;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   4890
         TabIndex        =   64
         Top             =   1950
         Width           =   2745
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4842;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   8
         Left            =   4890
         TabIndex        =   63
         Top             =   1680
         Width           =   2805
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4948;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   19
         Left            =   -73920
         TabIndex        =   60
         Top             =   2700
         Width           =   6225
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10980;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   18
         Left            =   -73920
         TabIndex        =   59
         Top             =   2430
         Width           =   6435
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11351;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   17
         Left            =   -73080
         TabIndex        =   58
         Top             =   2115
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   16
         Left            =   -73620
         TabIndex        =   57
         Top             =   1830
         Width           =   6090
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10742;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   -73920
         TabIndex        =   56
         Top             =   1530
         Width           =   1155
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2037;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   14
         Left            =   -73920
         TabIndex        =   55
         Top             =   1230
         Width           =   6255
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11033;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   13
         Left            =   -73920
         TabIndex        =   54
         Top             =   690
         Width           =   6390
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11271;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   1425
         TabIndex        =   53
         Top             =   3300
         Width           =   3390
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5980;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   10
         Left            =   4890
         TabIndex        =   52
         Top             =   2280
         Width           =   1575
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2778;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   51
         Top             =   3030
         Width           =   2775
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4895;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   50
         Top             =   2790
         Width           =   735
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1296;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   49
         Top             =   2550
         Width           =   2775
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4895;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   4
         Left            =   2370
         TabIndex        =   48
         Top             =   2280
         Width           =   1455
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2566;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   47
         Top             =   2280
         Width           =   1095
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1931;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   810
         TabIndex        =   46
         Top             =   1680
         Width           =   2985
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5265;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "(1.中文  2.英文  3.日文）"
         Height          =   255
         Left            =   1920
         TabIndex        =   45
         Top             =   2790
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y：閉卷）"
         Height          =   255
         Left            =   6690
         TabIndex        =   44
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "是否閉卷："
         Height          =   255
         Left            =   3930
         TabIndex        =   43
         Top             =   2550
         Width           =   900
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "代理人備註："
         Height          =   255
         Left            =   -74910
         TabIndex        =   42
         Top             =   3330
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人："
         Height          =   255
         Left            =   -74910
         TabIndex        =   41
         Top             =   2700
         Width           =   1080
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人："
         Height          =   255
         Left            =   -74910
         TabIndex        =   40
         Top             =   2430
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人：               （Y：印）"
         Height          =   255
         Left            =   -74910
         TabIndex        =   39
         Top             =   2115
         Width           =   3240
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "固定請款對象："
         Height          =   255
         Left            =   -74910
         TabIndex        =   38
         Top             =   1830
         Width           =   1260
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人："
         Height          =   255
         Left            =   -74910
         TabIndex        =   37
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "折扣："
         Height          =   255
         Left            =   -74910
         TabIndex        =   36
         Top             =   1530
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "客戶備註："
         Height          =   180
         Left            =   60
         TabIndex        =   35
         Top             =   4410
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "申請日："
         Height          =   255
         Left            =   60
         TabIndex        =   34
         Top             =   1973
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "廠商號碼："
         Height          =   255
         Left            =   3930
         TabIndex        =   33
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "使用期限：           "
         Height          =   255
         Left            =   60
         TabIndex        =   32
         Top             =   2280
         Width           =   1395
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   255
         Left            =   60
         TabIndex        =   31
         Top             =   1680
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   2250
         X2              =   2370
         Y1              =   2370
         Y2              =   2370
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請國家："
         Height          =   180
         Left            =   3930
         TabIndex        =   30
         Top             =   1710
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "案件備註："
         Height          =   180
         Left            =   60
         TabIndex        =   29
         Top             =   3870
         Width           =   900
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(日)："
         Height          =   255
         Left            =   60
         TabIndex        =   28
         Top             =   1343
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(英)："
         Height          =   255
         Left            =   60
         TabIndex        =   27
         Top             =   983
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(中)："
         Height          =   255
         Left            =   60
         TabIndex        =   26
         Top             =   623
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   255
         Left            =   60
         TabIndex        =   25
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "FC代理人："
         Height          =   255
         Left            =   -74910
         TabIndex        =   24
         Top             =   390
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號："
         Height          =   255
         Left            =   -74910
         TabIndex        =   23
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         Caption         =   "閉卷日期："
         Height          =   255
         Left            =   60
         TabIndex        =   22
         Top             =   2550
         Width           =   900
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         Caption         =   "閉卷原因："
         Height          =   255
         Left            =   3930
         TabIndex        =   21
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文："
         Height          =   255
         Left            =   60
         TabIndex        =   20
         Top             =   2790
         Width           =   900
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "分所案號："
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   3030
         Width           =   900
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號："
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Top             =   3300
         Width           =   1260
      End
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   21
      Left            =   5130
      TabIndex        =   62
      Top             =   5820
      Width           =   2895
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "5106;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   20
      Left            =   1050
      TabIndex        =   61
      Top             =   5820
      Width           =   3075
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "5424;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label49 
      Caption         =   "Create ID："
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   5820
      Width           =   855
   End
   Begin VB.Label Label51 
      Caption         =   "Update ID："
      Height          =   255
      Left            =   4170
      TabIndex        =   16
      Top             =   5820
      Width           =   975
   End
End
Attribute VB_Name = "frm100101_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/20 改成Form2.0 ; lbl1(index)、txt1(index)、grdDataList1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
Option Explicit

Dim StrTag As String, StrTag1 As String, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim m_pub_QL05 As String 'Add By Sindy 2025/8/28 只記錄於此Form


'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag1 ' StrTag    傳申請人代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 1
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 2
     fnCloseAllFrm100
Case 3
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_10.Show
     frm100101_10.Tag = StrTag
     frm100101_10.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_10.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'add by nickc 2005/05/31
Case 4
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100108_3.Show
     frm100108_3.Tag = txt1(7).Text
     frm100108_3.Caption = "相關卷號"
     frm100108_3.StrMenu2
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'Added by Lydia 2016/11/23
Case 5 '各項指示
    'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
    If PUB_CheckFormExist("frm12040159") Then
        MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
        Exit Sub
    End If
    'end 2020/05/05
    
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm12040159.SetParent "Q", Trim(Replace(txt1(7), "-", "")), Me
     frm12040159.Show
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'end 2016/11/23
'Add By Sindy 2020/7/15
Case 6 '進度
   cmdState = -1
   Me.Enabled = False
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   frm100101_2.Show
   frm100101_2.Tag = txt1(7)
   frm100101_2.StrMenu
   Screen.MousePointer = vbDefault
   Me.Enabled = True
Case Else
End Select
End Sub


Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
End Sub

Sub StrMenu()
Dim strSql  As String
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String
'edit by nickc 2006/07/12
'Dim strArr(T_SP) As String, i As Integer, StrOk(23) As String, StrOkTxt(6) As String
Dim strArr() As String, i As Integer, StrOk(23) As String, StrOkTxt(6) As String
'add by nickc 2006/07/12
ReDim strArr(tf_SP) As String
'ADD by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
Dim tmp01 As String, tmp02 As String

'Add By Cheng 2002/07/08
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strSK03 As String

'add by Toni 20080926 控制跨部門權限訊息
Dim strTit As String
Dim strMsg As String
Dim nResponse
'End by Toni 20080926

Str01 = ""
Str02 = ""
Str03 = ""
Str04 = ""
If Left(Me.Tag, 1) = "N" Then
   strSql = Right(Me.Tag, Len(Me.Tag) - 1)
Else
   strSql = Me.Tag
End If
Str01 = SystemNumber(strSql, 1)
Str02 = SystemNumber(strSql, 2)
Str03 = SystemNumber(strSql, 3)
Str04 = SystemNumber(strSql, 4)

'add by Toni 20080926 控制跨部門權限 for 條碼
'2008/10/2 modify by sonia
'If IsUserHasRightOfSystem(strUserNum, Str01) = False Then
'   If IsUserHasRightOfFunction("frm100101_1", strCrossDept, False) = False Then
'      strTit = "檢核資料"
'      strMsg = "您沒有使用該系統類別的權限"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Exit Sub
'   End If
'End If
If CheckSR09(strUserNum, Str01, "Y", , Str01, Str02, Str03, Str04) = False Then
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
'2008/10/2 end
'End 20080926

pub_QL05 = ";本所案號：" & Str01 & "-" & Str02 & "-" & Str03 & "-" & Str04 & _
           "(基本資料)" 'Add By Sindy 2025/8/7

'Add By Cheng 2002/07/08
strSK03 = ""
StrSQLa = "Select SK03 From SystemKind Where SK01='" & Str01 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic
If rsA.RecordCount > 0 Then
   strSK03 = "" & rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

'欲搜尋的SQL字串
strSql = "SELECT * FROM SERVICEPRACTICE WHERE SP01='" & Str01 & "' AND SP02='" & Str02 & "' AND SP03='" & Str03 & "' AND SP04='" & Str04 & "'"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28 記錄此Form的查詢條件
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/16
   'For i = 0 To 61
   For i = 0 To (tf_SP - 1) 'edit by nickc 2006/07/12 (tf_sp-1 'edit by nickc 2006/07/12 T_SP - 1)
      Select Case i
      Case 9, 11, 15, 19, 20, 30, 38, 39, 52, 53, 55, 56
           If IsNull(adoRecordset.Fields(i)) Then
               strArr(i + 1) = ""
           Else
               strArr(i + 1) = str(adoRecordset.Fields(i))
           End If
      Case Else
           If IsNull(adoRecordset.Fields(i)) Then
                strArr(i + 1) = ""
           Else
                strArr(i + 1) = adoRecordset.Fields(i)
           End If
      End Select
      DoEvents
   Next i
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2010/11/16
   ShowNoData
   Screen.MousePointer = vbDefault
       '920416 nick
     'Me.Hide
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
CheckOC
Dim strTemp As Variant  '暫存
Dim strTemp1 As Variant, strTemp2 As Variant, strTemp3 As Variant
Dim j As Integer
intK = 62
'For i = 0 To 62
For i = 1 To tf_SP 'edit by nickc 2006/07/12 T_SP
    Select Case i
    Case 1
         If Left(Me.Tag, 1) = "N" Then
            StrOk(0) = Right(Me.Tag, Len(Me.Tag) - 1)
         Else
            StrOk(0) = Me.Tag
         End If
         txt1(7) = StrOk(0) 'Add By Sindy 2013/1/31
    Case 8
         If Len(strArr(i)) = 9 Then
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
         Else
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(1) = strArr(i) + ""
'                     Else
'                          StrOk(1) = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     StrOk(1) = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  StrOk(1) = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
            If IsNull(adoRecordset.Fields("CU04")) = False Then
               StrOk(1) = strArr(i) + "  " + adoRecordset.Fields("CU04")
            ElseIf IsNull(adoRecordset.Fields("CU05")) = False Then
               StrOk(1) = strArr(i) + "  " + adoRecordset.Fields("CU05")
            ElseIf IsNull(adoRecordset.Fields("CU06")) = False Then
               StrOk(1) = strArr(i) + "  " + adoRecordset.Fields("CU06")
            End If
             If IsNull(adoRecordset.Fields(3)) Then
                  StrOkTxt(5) = ""
             Else
                  StrOkTxt(5) = adoRecordset.Fields(3)
             End If
            'Add by Morgan 2004/1/14
            Lbl1(1).ForeColor = vbBlack
         Else
            StrOk(1) = ""
            'Add by Morgan 2004/1/14
            Lbl1(1).ForeColor = vbRed
             StrOk(1) = strArr(i)
             StrOkTxt(5) = ""
         End If
         CheckOC
    Case 10
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(2) = ""
         Else
             StrOk(2) = ChangeWStringToTString(strArr(i))
         End If
         txt1(8) = StrOk(2) 'Add By Sindy 2013/1/31
    Case 20
         StrOk(3) = strArr(i)
    Case 21
         StrOk(4) = strArr(i)
    Case 16
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(5) = ""
         Else
             StrOk(5) = ChangeWStringToTString(strArr(i))
         End If
    Case 34
         StrOk(6) = strArr(i)
    Case 28
         StrOk(7) = strArr(i)
    Case 9
         strSql = "SELECT NA03 FROM NATION WHERE NA01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(8) = strArr(i) + ""
              Else
                  StrOk(8) = strArr(i) + "  " + adoRecordset.Fields(0)
              End If
         Else
              StrOk(8) = ""
         End If
         CheckOC
    Case 19
         StrOk(9) = strArr(i)
    Case 15
         StrOk(10) = strArr(i)
    Case 29
         StrOk(11) = strArr(i)
    Case 26
         If Len(strArr(i)) = 9 Then
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
              strSql = "SELECT FA05,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
         Else
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
              strSql = "SELECT FA05,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            '2005/9/15 MODIFY BY SONIA
            'If IsNull(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))) Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'            If Trim(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))) = "" Then
'            '2005/9/14 END
'               'Modify By Cheng 2002/07/08
''               If IsNull(adoRecordset.Fields(1)) Then
'               If IsNull(adoRecordset.Fields(IIf(strSK03 = "0", 0, 1))) Then
'                   If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(12) = strArr(i) + ""
'                   Else
'                         StrOk(12) = strArr(i) + "  " + adoRecordset.Fields(2)
'                   End If
'               Else
'                  'Modify By Cheng 2002/07/08
''                   StrOk(12) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                   StrOk(12) = strArr(i) + "  " + adoRecordset.Fields(IIf(strSK03 = "0", 0, 1))
'               End If
'            Else
'               'Modify By Cheng 2002/07/08
''               StrOk(12) = StrArr(i) + "  " + adoRecordset.Fields(0)
'               StrOk(12) = strArr(i) + "  " + adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))
'
'            End If
            If IsNull(adoRecordset.Fields("FA05")) = False Then
               StrOk(12) = strArr(i) + "  " + adoRecordset.Fields("FA05")
            ElseIf IsNull(adoRecordset.Fields("FA04")) = False Then
               StrOk(12) = strArr(i) + "  " + adoRecordset.Fields("FA04")
            ElseIf IsNull(adoRecordset.Fields("FA06")) = False Then
               StrOk(12) = strArr(i) + "  " + adoRecordset.Fields("FA06")
            End If
            
            If IsNull(adoRecordset.Fields(3)) Then
                StrOkTxt(6) = ""
            Else
                StrOkTxt(6) = adoRecordset.Fields(3)
            End If
            'Add by Morgan 2004/1/16
            Lbl1(12).ForeColor = vbBlack
         Else
            StrOk(12) = ""
            'Add by Morgan 2004/1/6
            Lbl1(12).ForeColor = vbRed
            StrOk(12) = strArr(i)
            
            StrOkTxt(6) = ""
         End If
         CheckOC
    Case 27
         StrOk(13) = strArr(i)
    Case 30
         StrOk(14) = strArr(i)
    Case 31
         StrOk(15) = strArr(i)
    Case 37
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
            StrOk(16) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
            If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
               StrOk(16) = strArr(i) + "  " + tmp02
            Else
               StrOk(16) = strArr(i)
            End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Modify By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'               'Modify By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(16) = strArr(i) + ""
'                    Else
'                        StrOk(16) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                  'Modify By Cheng 2002/07/08
''                    StrOk(16) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(16) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(16) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(16) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(16) <> strArr(i) Then
            'Add by Morgan 2004/1/16
            Lbl1(16).ForeColor = vbBlack
         Else
            StrOk(16) = ""
            'Add by Morgan 2004/1/16
            Lbl1(16).ForeColor = vbRed
            StrOk(16) = ""
         End If
         CheckOC
    Case 33
         StrOk(17) = strArr(i)
    Case 35
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
            StrOk(18) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
            If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
               StrOk(18) = strArr(i) + "  " + tmp02
            Else
               StrOk(18) = strArr(i)
            End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Modify By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'               'Modify By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(18) = strArr(i) + ""
'                    Else
'                        StrOk(18) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                  'Modify By Cheng 2002/07/08
''                    StrOk(18) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(18) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(18) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(18) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(18) <> strArr(i) Then
            'Add by Morgan 2004/1/16
            Lbl1(18).ForeColor = vbBlack
         Else
            StrOk(18) = ""
            'Add by Morgan 2004/1/16
            Lbl1(18).ForeColor = vbRed
            StrOk(18) = strArr(i)
         End If
         CheckOC
    Case 36
         StrOk(19) = strArr(i)
    Case 52
         'edit by nick 2004/10/05
         'StrOk(20) = strArr(i)
         StrOk(20) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(53))) & " " & Format(strArr(54), "##:##")
    Case 55
         'edit by nick 2004/10/05
         'StrOk(21) = strArr(i)
         StrOk(21) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(56))) & " " & Format(strArr(57), "##:##")
    Case 61
         'edit by nickc 2006/07/12
         'StrOk(22) = strArr(i)
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(22) = ""
         Else
             StrOk(22) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 23
         '911016 nick 更新成 table
         'strTemp = Split(StrArr(i), ",")
         'grdDataList1.col = 1
         'For j = 1 To Len(StrArr(i))
         '   grdDataList1.Row = j
         '   grdDataList1.Text = Mid(StrArr(i), j, 1)
         'Next j
    Case 22
         '911016 nick 更新成 table
         'strTemp = Split(StrArr(i), ",")
         'grdDataList1.col = 0
         'For j = 0 To UBound(strTemp)
         '   grdDataList1.Row = j + 1
         '   grdDataList1.Text = strTemp(j)
         'Next j
    Case 5
         StrOkTxt(0) = strArr(i)
    Case 6
         StrOkTxt(1) = strArr(i)
    Case 7
         StrOkTxt(2) = strArr(i)
    Case 17
         strSql = "SELECT ROR02 FROM REASONOFRELIEF WHERE ROR01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
             If IsNull(adoRecordset.Fields(0)) Then
                     StrOkTxt(3) = strArr(i) + ""
             Else
                     StrOkTxt(3) = strArr(i) + "  " + adoRecordset.Fields(0)
             End If
         Else
             StrOkTxt(3) = ""
         End If
         CheckOC
    Case 18
         StrOkTxt(4) = strArr(i)
    Case 67 'D/N固定列印對象
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(23) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
            If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
               StrOk(23) = strArr(i) + "  " + tmp02
            Else
               StrOk(23) = strArr(i)
            End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(23) = strArr(i) + ""
'                    Else
'                        StrOk(23) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(23) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'                StrOk(23) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(23) <> strArr(i) Then
            'Add by Morgan 2004/1/16
            Lbl1(23).ForeColor = vbBlack
         Else
            StrOk(23) = ""
            'Add by Morgan 2004/1/16
            Lbl1(23).ForeColor = vbRed
            StrOk(23) = strArr(i)
         End If
         CheckOC
    'add by nickc 2006/07/12
    Case 68
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             Lbl1(24) = ""
         Else
             Lbl1(24) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 69
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               Lbl1(25) = strArr(i) + ""
            Else
               Lbl1(25) = strArr(i) + "  " + adoRecordset.Fields(0)
            End If
         Else
            Lbl1(25) = ""
         End If
         CheckOC
    Case 70
         Lbl1(26) = strArr(i)
    'Add by Morgan 2008/8/5
    Case 78
         Lbl1(27) = PUB_GetContact(strArr(8), strArr(i))
         
    Case 84 'Add by Morgan 2010/11/8
         Lbl1(i) = strArr(i)
    Case 85 'Add by Sindy 2014/2/10
         Lbl1(i) = strArr(i)
    
    'Added by Morgan 2016/12/8
    Case 86 '國內副本收件人
         Lbl1(i) = strArr(i)
         If strArr(i) <> "" Then
            If ClsLawLawGetName(strArr(i), strExc(9)) = True Then
               Lbl1(i) = Lbl1(i) + "  " + strExc(9)
            End If
         End If
    Case 87 '國內副本接洽人
         If strArr(86) <> "" And strArr(i) <> "" Then
            Lbl1(i) = PUB_GetContact(strArr(86), strArr(i))
         Else
            Lbl1(i) = ""
         End If
    'end 2016/12/8
    Case Else
    End Select
    DoEvents
Next i
For i = 0 To 23                  '2006/07/12 加備註，以後新增欄位，直接在上面修改，此2段迴圈
   If i <> 0 And i <> 2 Then 'Add By Sindy 2013/1/31 +if
      Lbl1(i) = StrOk(i)           '不可修改，不然會影響資料顯現，而且陣列的宣告也不用一直的修改
   End If
Next i
For i = 0 To 6
   txt1(i) = StrOkTxt(i)
Next i
'傳參數　　　代理人
StrTag = strArr(26)
'傳參數　　　申請人
StrTag1 = strArr(8)
'911016 nick 讀取正片號碼檔
'***** start
Dim nick911016rs As New ADODB.Recordset
Dim nickstrsql As String
   nickstrsql = "select bc02,bc03,bc01 from barcode,caseprogress where cp01='" & strArr(1) & "' " & _
            " and cp02='" & strArr(2) & "' and cp03='" & strArr(3) & "' " & _
            " and cp04='" & strArr(4) & "' and cp09=bc01 order by bc01,bc02,bc03 "
Set nick911016rs = New ADODB.Recordset
nick911016rs.CursorLocation = adUseClient
nick911016rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
If nick911016rs.RecordCount <> 0 Then
    Set grdDataList1.Recordset = nick911016rs
End If
GRIDHEAND
'***** end
'add by nickc 2005/05/31  檢查有無分割或相關卷號
     cmdok(4).Visible = ChkDataByCR(txt1(7).Text)
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/28 還原此Form的查詢條件記錄
End Sub

'edit by nickc 2005/05/30 改成與我們現在共用相同
'Private Sub cmdRef_Click()
'    Dim stTmp As String
'    stTmp = Right(Space(2) & txt1(7), 15)
'    Where1103ComeFrom Me, Trim(Left(stTmp, 3)), Mid(stTmp, 5, 6), Mid(stTmp, 12, 1), Mid(stTmp, 14, 2)
'End Sub

Private Sub Form_Load()
Dim Lbl As Object

For Each Lbl In Me.Lbl1
    Lbl.BackColor = &H8000000F
Next
bolToEndByNick = False

SSTab3.Tab = 0 'Added by Lydia 2016/11/23

   MoveFormToCenter Me
   
If bolFNation = False Then
    SSTab3.TabVisible(1) = False
    cmdok(3).Visible = False
End If
Call GRIDHEAND
'92.04.16 nick
cmdState = -1

'Added by Lydia 2020/05/05 各項指示：顯示按鈕
If strSrvDate(1) >= 各項指示啟用日 Then
   cmdok(5).Visible = True
Else
   cmdok(5).Visible = False
End If
'end 2020/05/05
    
End Sub

Private Function GRIDHEAND()
With grdDataList1
'911016 nick
.Cols = 3
.row = 0
.col = 0
.ColWidth(0) = 1800
.Text = "正片號碼"
.col = 1
.ColWidth(1) = 1200
.Text = "是否合格"
'911016 nick 新增
.col = 2
.ColWidth(2) = 1500
.Text = "收文號"
End With
End Function

Private Sub Form_Unload(Cancel As Integer)
pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
Set frm100101_7 = Nothing
End Sub
'add by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
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
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
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
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
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

'Added by Lydia 2016/10/27 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index) 'Added by Lydia 2016/12/6
   OpenIme
End Sub
