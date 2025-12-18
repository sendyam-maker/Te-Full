VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_27 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "客戶端平台帳號管理作業"
   ClientHeight    =   5670
   ClientLeft      =   180
   ClientTop       =   990
   ClientWidth     =   8960
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   8100
      TabIndex        =   49
      Top             =   50
      Width           =   800
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   0
      Left            =   6840
      TabIndex        =   48
      Top             =   50
      Width           =   1230
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm100101_27.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm100101_27.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm100101_27.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm100101_27.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm100101_27.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm100101_27.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm100101_27.frx":0234
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm100101_27.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm100101_27.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm100101_27.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm100101_27.frx":03AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   30
      TabIndex        =   1
      Top             =   480
      Width           =   8895
      _ExtentX        =   15699
      _ExtentY        =   8484
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "平台資訊"
      TabPicture(0)   =   "frm100101_27.frx":040A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textCW(18)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textCW(17)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCW(15)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textCW(14)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textCW(16)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textCW(13)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCW(12)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCW(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lstAtt"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCW(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCW(5)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCW(1)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lstUsers(1)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "CommonDialog1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cboCW19"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Frame5"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Frame4"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cboCW03"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmdSelect(0)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdSaveAtt(0)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdOpenAtt(0)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "CmdOK(2)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).ControlCount=   36
      TabCaption(1)   =   "帳號資料"
      TabPicture(1)   =   "frm100101_27.frx":0426
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(2)"
      Tab(1).Control(1)=   "lstUsers(2)"
      Tab(1).Control(2)=   "grd1"
      Tab(1).Control(3)=   "Text2"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).ControlCount=   5
      Begin VB.CommandButton CmdOK 
         Caption         =   "匯出客戶名單"
         Height          =   345
         Index           =   2
         Left            =   7170
         TabIndex        =   29
         Top             =   2070
         Width           =   1260
      End
      Begin VB.Frame Frame1 
         Caption         =   "複製區"
         ForeColor       =   &H000000C0&
         Height          =   465
         Left            =   -74790
         TabIndex        =   50
         Top             =   4250
         Width           =   7170
         Begin VB.TextBox textCD03 
            BackColor       =   &H8000000F&
            Height          =   264
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   150
            Width           =   1815
         End
         Begin VB.TextBox textCD04 
            BackColor       =   &H8000000F&
            Height          =   264
            Left            =   4350
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   150
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "帳　號："
            Height          =   180
            Index           =   5
            Left            =   750
            TabIndex        =   54
            Top             =   150
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "密　碼："
            Height          =   180
            Index           =   4
            Left            =   3630
            TabIndex        =   53
            Top             =   150
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   255
         Index           =   0
         Left            =   8070
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3450
         Width           =   735
      End
      Begin VB.CommandButton cmdSaveAtt 
         Caption         =   "下載"
         Height          =   255
         Index           =   0
         Left            =   8070
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3990
         Width           =   735
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "全選"
         Height          =   255
         Index           =   0
         Left            =   8070
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3720
         Width           =   735
      End
      Begin VB.ComboBox cboCW03 
         Height          =   260
         ItemData        =   "frm100101_27.frx":0442
         Left            =   1050
         List            =   "frm100101_27.frx":0444
         Style           =   1  '組合式
         TabIndex        =   23
         Text            =   "cboCW03"
         Top             =   630
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   -74850
         TabIndex        =   20
         Text            =   $"frm100101_27.frx":0446
         Top             =   870
         Width           =   2235
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   4620
         TabIndex        =   12
         Top             =   2700
         Width           =   3105
         Begin VB.OptionButton Option1 
            Caption         =   "本所管理"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   14
            Top             =   0
            Width           =   1065
         End
         Begin VB.OptionButton Option1 
            Caption         =   "客戶核准"
            Height          =   255
            Index           =   1
            Left            =   2010
            TabIndex        =   13
            Top             =   0
            Width           =   1065
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "管理方式："
            Height          =   180
            Left            =   30
            TabIndex        =   15
            Top             =   30
            Width           =   900
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   30
         TabIndex        =   7
         Top             =   2700
         Width           =   3735
         Begin VB.CheckBox Check2 
            Caption         =   "帳號密碼"
            Height          =   315
            Index           =   0
            Left            =   1020
            TabIndex        =   10
            Top             =   -30
            Width           =   1035
         End
         Begin VB.CheckBox Check2 
            Caption         =   "憑證"
            Height          =   315
            Index           =   1
            Left            =   2070
            TabIndex        =   9
            Top             =   -30
            Width           =   705
         End
         Begin VB.CheckBox Check2 
            Caption         =   "網址"
            Height          =   315
            Index           =   2
            Left            =   2790
            TabIndex        =   8
            Top             =   -30
            Width           =   705
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "驗證方式："
            Height          =   180
            Left            =   60
            TabIndex        =   11
            Top             =   30
            Width           =   900
         End
      End
      Begin VB.ComboBox cboCW19 
         Height          =   260
         ItemData        =   "frm100101_27.frx":0463
         Left            =   6480
         List            =   "frm100101_27.frx":0465
         Style           =   1  '組合式
         TabIndex        =   2
         Text            =   "cboCW19"
         Top             =   630
         Width           =   1305
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   2000
         Left            =   -74790
         TabIndex        =   32
         Top             =   2190
         Width           =   8490
         _ExtentX        =   14975
         _ExtentY        =   3528
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   9
         FixedCols       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "客戶編號 |客戶名稱 |帳號 |密碼 |身份別 |使用者 |建置日期 |下次更新日期 |註解"
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
         _Band(0).Cols   =   9
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   510
         Top             =   4110
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   930
         Index           =   1
         Left            =   1050
         TabIndex        =   31
         Top             =   1770
         Width           =   5970
         VariousPropertyBits=   746586139
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "10530;1640"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW 
         Height          =   285
         Index           =   1
         Left            =   1050
         TabIndex        =   26
         Top             =   360
         Width           =   615
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1085;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW 
         Height          =   480
         Index           =   5
         Left            =   1050
         TabIndex        =   25
         Top             =   2970
         Width           =   7755
         VariousPropertyBits=   -1466941413
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "13679;847"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW 
         Height          =   285
         Index           =   2
         Left            =   1050
         TabIndex        =   24
         Top             =   930
         Width           =   6990
         VariousPropertyBits=   671105051
         Size            =   "12330;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstAtt 
         Height          =   1320
         Left            =   1050
         TabIndex        =   22
         Top             =   3450
         Width           =   6990
         VariousPropertyBits=   746586139
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "12330;2328"
         MatchEntry      =   0
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   1680
         Index           =   2
         Left            =   -72570
         TabIndex        =   21
         Top             =   480
         Width           =   6225
         VariousPropertyBits=   746586139
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "10980;2963"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW 
         Height          =   285
         Index           =   4
         Left            =   30
         TabIndex        =   19
         Top             =   2310
         Width           =   345
         VariousPropertyBits=   671105051
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW 
         Height          =   285
         Index           =   12
         Left            =   4050
         TabIndex        =   18
         Top             =   360
         Width           =   4560
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "8043;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW 
         Height          =   285
         Index           =   13
         Left            =   3330
         TabIndex        =   17
         Top             =   660
         Width           =   795
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1402;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW 
         Height          =   285
         Index           =   16
         Left            =   5070
         TabIndex        =   16
         Top             =   660
         Width           =   405
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "714;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW 
         Height          =   285
         Index           =   14
         Left            =   360
         TabIndex        =   6
         Top             =   2310
         Width           =   345
         VariousPropertyBits=   671105051
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW 
         Height          =   285
         Index           =   15
         Left            =   690
         TabIndex        =   5
         Top             =   2310
         Width           =   345
         VariousPropertyBits=   671105051
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW 
         Height          =   285
         Index           =   17
         Left            =   1050
         TabIndex        =   4
         Top             =   1200
         Width           =   6990
         VariousPropertyBits=   671105051
         Size            =   "12330;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCW 
         Height          =   285
         Index           =   18
         Left            =   1050
         TabIndex        =   3
         Top             =   1470
         Width           =   6990
         VariousPropertyBits=   671105051
         Size            =   "12330;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客　　戶："
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   47
         Top             =   1770
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "平台編號："
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   46
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "平台類別："
         Height          =   180
         Left            =   90
         TabIndex        =   45
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "備　　註："
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   44
         Top             =   3000
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "網　址 1 ："
         Height          =   180
         Left            =   90
         TabIndex        =   43
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(系統自動給號)"
         Height          =   180
         Index           =   7
         Left            =   1740
         TabIndex        =   42
         Top             =   390
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "附件或憑證："
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   41
         Top             =   3480
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客　　戶："
         Height          =   180
         Index           =   2
         Left            =   -73440
         TabIndex        =   40
         Top             =   510
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "平台名稱："
         Height          =   180
         Left            =   3120
         TabIndex        =   39
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "建置日期："
         Height          =   180
         Left            =   2415
         TabIndex        =   38
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "更新週期：           (月)"
         Height          =   180
         Left            =   4150
         TabIndex        =   37
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "網　址 2 ："
         Height          =   180
         Left            =   90
         TabIndex        =   36
         Top             =   1230
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "網　址 3 ："
         Height          =   180
         Left            =   90
         TabIndex        =   35
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label13 
         Caption         =   "註：網址欄位（快按二下）即可進入網站"
         ForeColor       =   &H000000C0&
         Height          =   915
         Left            =   8070
         TabIndex        =   34
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "性質："
         Height          =   180
         Left            =   5920
         TabIndex        =   33
         Top             =   690
         Width           =   540
      End
   End
   Begin MSForms.Label Label23 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   5370
      Width           =   8580
      VariousPropertyBits=   27
      Caption         =   "CREATE : 　　　  101/09/03  13:54:00          UPDATE : 　　　  101/09/04  09:21:44"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm100101_27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/06 Form2.0已修改
'Create by Amy 2015/04/16 與frm100101_26 類似-以客戶編號為主(只能顯示一個客戶)
Option Explicit

'附件
Dim m_FilesRemoved() As String
Dim ii As Integer, jj As Integer
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
'
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

' 變數宣告區
Dim RsQ As New ADODB.Recordset
Dim strSql As String
Dim MyArr As Variant
Dim m_AttachPath As String
Dim i As Integer, j As Integer
Dim oText As Object 'TextBox
Dim oCheck As Object 'CheckBox
Dim oOption As Object 'OptionButton
Dim strCW01 As String
Dim bolMuchCust As Boolean '是否為多筆客戶

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
      Case 2
         Call RunExcelFile 'Add By Sindy 2021/4/7
      Case Else
   End Select
End Sub

'Add By Sindy 2021/4/7
'匯出客戶名單
Private Sub RunExcelFile()
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim intItem As Integer, intCounter As Integer
Dim varTemp As Variant
   
   Set xlsAnnuity = New Excel.Application
   
   xlsAnnuity.Visible = True
   xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   xlsAnnuity.ActiveWindow.Zoom = 75 '畫面比例100%太大了,調整為75%
   'wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
   wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity.PageSetup.LeftMargin = 28.34
   wksAnnuity.PageSetup.RightMargin = 28.34
   wksAnnuity.PageSetup.TopMargin = 42.51
   wksAnnuity.PageSetup.BottomMargin = 42.51
   wksAnnuity.PageSetup.HeaderMargin = 28.34
   wksAnnuity.PageSetup.FooterMargin = 28.34
   '設定各欄位長度
   wksAnnuity.Columns("A:A").ColumnWidth = 10
   wksAnnuity.Columns("B:B").ColumnWidth = 30
   '標題
   intCounter = 1
   xlsAnnuity.Range("A" & intCounter).Value = "編號"
   xlsAnnuity.Range("B" & intCounter).Value = "客戶名稱"
   
   For intItem = 0 To lstUsers(1).ListCount - 1
      varTemp = Split(lstUsers(1).List(intItem), "@")
      intCounter = intCounter + 1
      xlsAnnuity.Range("A" & intCounter).Value = Left(varTemp(0), 9)
      xlsAnnuity.Range("B" & intCounter).Value = Trim(Mid(varTemp(0), 10))
   Next intItem
   
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
End Sub

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
    Dim hLocalFile As Long
    Dim stFileName As String
    Dim strAtt As String, strType As String
    
    Screen.MousePointer = vbHourglass
    
    If Index = 0 Then
       strAtt = lstAtt.List(lstAtt.ListIndex)
    End If
    
    If strAtt = "" Then
       MsgBox "請選擇欲開啟的附件！"
    Else
       For ii = 0 To lstAtt.ListCount - 1
          If lstAtt.Selected(ii) Then
             stFileName = lstAtt.List(ii)
             If InStrRev(stFileName, " (") > 0 Then
                stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
             End If
             
             If InStr(stFileName, "\") = 0 Then
                If GetAttachFile(stFileName) = False Then
                   Exit Sub
                End If
             End If
             
             ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
          End If
       Next ii
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSaveAtt_Click(Index As Integer)
    Dim stFileName As String, stFolderPath As String, stFullName As String
    Dim bMultiFile As Boolean
    Dim ii As Integer, oList As Object
    
    Screen.MousePointer = vbHourglass
    
    If Index = 0 Then
       Set oList = lstAtt
    End If
    
    stFileName = ""
    bMultiFile = False
    For ii = 0 To oList.ListCount - 1
       If oList.Selected(ii) Then
          stFileName = oList.List(ii)
          If stFileName <> "" Then
             bMultiFile = True
             Exit For
          Else
             stFileName = oList.List(ii)
          End If
       End If
    Next
    
    If stFileName = "" Then
       MsgBox "請選擇欲存檔的附件！"
    Else
       '多選
       If bMultiFile Then
          stFolderPath = BrowseForFolder()
          If stFolderPath <> "" Then
             For ii = 0 To oList.ListCount - 1
                If oList.Selected(ii) Then
                   stFileName = oList.List(ii)
                   If InStrRev(stFileName, " (") > 0 Then
                      stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
                   End If
                   stFullName = stFolderPath & stFileName
                   If stFullName <> "" Then
                      If Dir(stFullName) <> "" Then
                         If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                            stFullName = ""
                         End If
                      End If
                      If stFullName <> "" Then
                         If GetAttachFile(stFileName, stFullName) = False Then
                            MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                         End If
                      End If
                   End If
                End If
             Next
          End If
       
       Else
          stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
          stFullName = GetSaveName(stFileName)
          If stFullName <> "" Then
             If Dir(stFullName) <> "" Then
                If MsgBox("檔案[ " & stFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                   stFullName = ""
                End If
             End If
             If stFullName <> "" Then
                If GetAttachFile(stFileName, stFullName) = False Then
                   MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                End If
             End If
          End If
       End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim ii As Integer, oList As Object
    If Index = 0 Then
       Set oList = lstAtt
    End If
    
    For ii = 0 To oList.ListCount - 1
       lstAtt.Selected(ii) = True
    Next
End Sub

Private Sub Form_Load()
    bolToEndByNick = False
    MoveFormToCenter Me
    SetLock
    SSTab1.Tab = 0
    ClearField
    'Modify By Sindy 2021\5\19
    'm_AttachPath = App.path & "\SeminarAttach"
    m_AttachPath = App.path & "\SeminarAttach\" & strUserNum
    '2021\5\19 END
    textCW(4).Visible = False
    textCW(14).Visible = False
    textCW(15).Visible = False
    
    Pub_Can_Copy_Pic = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Pub_Can_Copy_Pic = False
    KillAttach
    Set frm100101_27 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next
   'Modify By Sindy 2021/2/3 刪不掉,改用函數
'   If Dir(m_AttachPath & "\.") <> "" Then
'      Kill m_AttachPath & "\*.*"
'   End If
   'Modify By Sindy 2021\5\19
   'PUB_KillTempFile "SeminarAttach\*.*"
   PUB_KillTempFile "SeminarAttach\" & strUserNum & "\*.*"
   '2021\5\19 END
   '2021/2/3 END
End Sub

Private Sub UpdateCUID()
    Dim strTemp As String
    Dim strCName As String, strCDate As String, strCTime As String
    Dim strUName As String, strUDate As String, strUTime As String
    With RsQ
        If IsNull(.Fields("cw06")) = False Then
            If IsEmptyText(.Fields("cw06")) = False Then
                strCName = GetStaffName(.Fields("cw06"), True)
            End If
        End If
        If IsNull(.Fields("cw07")) = False Then
            If IsEmptyText(.Fields("cw07")) = False Then
                strTemp = TAIWANDATE(.Fields("cw07"))
                strCDate = Format(strTemp, "###/##/##")
            End If
        End If
        If IsNull(.Fields("cw08")) = False Then
            If IsEmptyText(.Fields("cw08")) = False Then
                strTemp = .Fields("cw08")
                strCTime = Format(strTemp, "##:##")
            End If
        End If
        If IsNull(.Fields("cw09")) = False Then
            If IsEmptyText(.Fields("cw09")) = False Then
                strUName = GetStaffName(.Fields("cw09"), True)
            End If
        End If
        If IsNull(.Fields("cw10")) = False Then
            If IsEmptyText(.Fields("cw10")) = False Then
                strTemp = TAIWANDATE(.Fields("cw10"))
                strUDate = Format(strTemp, "###/##/##")
            End If
        End If
        If IsNull(.Fields("cw11")) = False Then
            If IsEmptyText(.Fields("cw11")) = False Then
                strTemp = .Fields("cw11")
                strUTime = Format(strTemp, "##:##")
            End If
        End If
    End With
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Sub SetLock(Optional ByVal bolYN As Boolean = True)
    For Each oText In textCW
        oText.Locked = bolYN
    Next
    Frame4.Enabled = Not bolYN '管理方式
    Frame5.Enabled = Not bolYN '驗證方式
    cboCW03.Locked = bolYN
    cboCW19.Locked = bolYN
End Sub

Public Sub StrMenu()
    Dim strCD02 As String
    
    strCW01 = Me.Tag
    '抓取平台主檔資料
    strSql = "Select * From CustWeb Where cw01='" & strCW01 & "' "
                  
    If RsQ.State = adStateOpen Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        FormShow
        SetlstUsers 1, textCW(4)
        UpdateCUID
    End If
    RsQ.Close
    
    '抓取平台帳號資料
    If textCW(4) <> MsgText(601) Then
        SetlstUsers 2, textCW(4) '取得客戶名稱
        '先預設為單筆客戶
        Text2.Visible = False
        lstUsers(2).Enabled = False
        bolMuchCust = False
        '檢查全部客戶設定同一帳號 (cw01=cd02)
        strSql = "Select Distinct(cd02) From CustWebId Where cd01 = '" & textCW(1) & "' And cd02<>cd01 Order by cd02 asc"
        RsQ.CursorLocation = adUseClient
        RsQ.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If RsQ.RecordCount > 0 Then
            lstUsers(2).Selected(0) = True
            MyArr = Split(lstUsers(2).List(lstUsers(2).ListIndex), "@")
            strCD02 = Trim(MyArr(1))
            bolMuchCust = True
            Text2.Visible = True
            lstUsers(2).Enabled = True
        Else
            strCD02 = strCW01
        End If
        QueryCustWebId (strCD02)
    End If
    RsQ.Close
    
    '附件檔
    strSql = "Select cf02,cf03 From CustWebFile Where cf01=" & strCW01 & " Order by cf02"
    
    If RsQ.State = adStateOpen Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        With RsQ
            Do While Not .EOF
                lstAtt.AddItem .Fields("cf02") & " (" & Round(.Fields("cf03") / 1024, 2) & " KB)", 0
                'lstAtt.ItemData(0) = 1
                .MoveNext
            Loop
        End With
        cmdOpenAtt(0).Enabled = True
        cmdSelect(0).Enabled = True
        cmdSaveAtt(0).Enabled = True
    End If
    'If lstAtt.ListCount > 0 Then SetListScroll lstAtt
End Sub

Private Sub SetGrd(bolMuchCust As Boolean)
    Dim arrGridHeadText, arrGridHeadWidth
    Dim iRow As Integer
   
    arrGridHeadText = Array("V", "客戶編號", "客戶名稱", "帳號", "密碼", "身份別", "使用者", "建置日期", "下次更新日期", "註解", "CD06")
    If bolMuchCust = True Then
       arrGridHeadWidth = Array(0, 1000, 1000, 1000, 1000, 800, 1000, 1000, 1200, 2000, 0)
    Else
       arrGridHeadWidth = Array(0, 0, 0, 1000, 1000, 800, 1000, 1000, 1200, 2000, 0)
    End If
    grd1.Visible = False
    grd1.Cols = UBound(arrGridHeadText) + 1
    For iRow = 0 To grd1.Cols - 1
       grd1.row = 0
       grd1.col = iRow
       grd1.Text = arrGridHeadText(iRow)
       grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
    Next
    grd1.Visible = True
End Sub

Private Sub FormShow()
    With RsQ
        If IsNull(.Fields("cw01")) = False Then: textCW(1) = .Fields("cw01")
        If IsNull(.Fields("cw02")) = False Then: textCW(2) = .Fields("cw02")
        If IsNull(.Fields("cw03")) = False Then: SetCombo cboCW03.Name, .Fields("cw03")
        If IsNull(.Fields("cw04")) = False Then: textCW(4) = .Fields("cw04") '客戶編號
        If IsNull(.Fields("cw05")) = False Then: textCW(5) = .Fields("cw05")
        If IsNull(.Fields("cw12")) = False Then: textCW(12) = .Fields("cw12")
        If IsNull(.Fields("cw13")) = False Then: textCW(13) = ChangeWStringToTString(.Fields("cw13"))
        If IsNull(.Fields("cw14")) = False Then
            MyArr = Split(.Fields("cw14"), ",")
            For j = 0 To UBound(MyArr)
                If MyArr(j) = "1" Then Check2(0).Value = 1
                If MyArr(j) = "2" Then Check2(1).Value = 1
                If MyArr(j) = "3" Then Check2(2).Value = 1
            Next j
            textCW(14) = .Fields("cw14")
        End If
        If IsNull(.Fields("cw15")) = False Then
            If .Fields("cw15") = "1" Then Option1(0).Value = True
            If .Fields("cw15") = "2" Then Option1(1).Value = True
            textCW(15) = .Fields("cw15")
        End If
        If IsNull(.Fields("cw16")) = False Then: textCW(16) = .Fields("cw16")
        If IsNull(.Fields("cw17")) = False Then: textCW(17) = .Fields("cw17")
        If IsNull(.Fields("cw18")) = False Then: textCW(18) = .Fields("cw18")
        cboCW19 = "" & .Fields("cw19") '性質
    End With
End Sub

Private Sub SetCombo(CboName As String, strValue As String)
    Select Case CboName
        Case "cboCW03"
            'Modified by Morgan 2017/10/24
            'Select Case strValue
            '    Case "1"
            '       cboCW03 = "1 IP管理"
            '    Case "2"
            '        cboCW03 = "2 檔案存取"
            '    Case "3"
            '        cboCW03 = "3 電子帳單"
            '    Case "4"
            '        cboCW03 = "4 憑證"
            '    Case Else
            'End Select
            cboCW03 = strValue & " " & PUB_GetCW03Name(strValue)
            'end 2017/10/24
    End Select
End Sub

Private Sub ClearField()
  
    '*** 平台資訊 ***
    For Each oText In textCW
        oText = Empty
    Next
    textCW(13) = strSrvDate(2)
    Label23.Caption = Empty
     
    '驗證方式
    For Each oCheck In Check2
        oCheck.Value = 0
    Next
    '管理方式
    For Each oOption In Option1
        oOption.Value = False
    Next
    
    lstUsers(1).Clear '客戶
    lstAtt.Clear '操作手冊
    Erase m_FilesRemoved
    ReDim m_FilesRemoved(0) As String
    cmdOpenAtt(0).Enabled = False
    cmdSelect(0).Enabled = False
    cmdSaveAtt(0).Enabled = False
    
    '*** 帳號資料 ***
    lstUsers(2).Clear
    Text2.Visible = False
    grd1.Clear
    grd1.Rows = 2
    Call SetGrd(False)
    'Add by Amy 2015/04/27
    textCD03.Text = Empty
    textCD04.Text = Empty
End Sub
   
Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   lstUsers(p_idx).Clear
   If p_stNums <> "" Then
      Select Case p_idx
         Case 0 '員工編號
            strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               arrID = Split(p_stNums, ",")
               With RsTemp
               '照原順序排
               For intI = UBound(arrID) To LBound(arrID) Step -1
                  .MoveFirst
                  Do While Not .EOF
                     If .Fields("st01") = arrID(intI) Then
                        lstUsers(p_idx).AddItem "" & .Fields(1) & "                                                            @" & .Fields(0), 0
                        .MoveLast
                     End If
                     .MoveNext
                  Loop
               Next
               End With
            End If
         Case 1, 2 '客戶編號
'            strExc(0) = "select cu01||cu02,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) from customer where cu01>' ' and instr('" & p_stNums & "',cu01||cu02)>0" & _
'                        " union" & _
'                        " select fa01||fa02,NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)) from fagent where fa01>' ' and instr('" & p_stNums & "',fa01||fa02)>0"
            strExc(0) = "select cu01||cu02,NVL(CU05||CU88||CU89||CU90,NVL(CU04,CU06)) from customer where cu01>' ' and instr('" & p_stNums & "',cu01||cu02)>0" & _
                        " union" & _
                        " select fa01||fa02,NVL(FA05||FA63||FA64||FA65,NVL(FA04,FA06)) from fagent where fa01>' ' and instr('" & p_stNums & "',fa01||fa02)>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               arrID = Split(p_stNums, ",")
               With RsTemp
               '照原順序排
               For intI = UBound(arrID) To LBound(arrID) Step -1
                  .MoveFirst
                  Do While Not .EOF
                     If .Fields(0) = arrID(intI) Then
                        'Modify By Sindy 2021/3/25 + 客戶編號
                        lstUsers(p_idx).AddItem .Fields(0) & " " & .Fields(1) & "                                                            @" & .Fields(0), 0
                        .MoveLast
                     End If
                     .MoveNext
                  Loop
               Next
               End With
            End If
      End Select
   End If
End Sub

Private Sub SetListScroll(oList As Object)
    Dim ii As Integer
    Dim lWnow As Long, lWmax As Long
     
     lWmax = 0
     For ii = 0 To oList.ListCount - 1
        lWnow = TextWidth(oList.List(ii) & " ")
        If lWnow > lWmax Then
           lWmax = lWnow
        End If
     Next
    
     If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
     SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Function GetAttachFile(ByRef pFileName As String, Optional pSavePath As String) As Boolean
    Dim stAttPath As String
    Dim lngSize As Long
    Dim iFileNo As Integer
    Dim bytes() As Byte
   
On Error GoTo ErrHnd
   
    If pSavePath = "" Then
        If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
        End If
        stAttPath = m_AttachPath & "\" & pFileName
        '檔案已存在時不必重新下載
        If Dir(stAttPath) <> "" Then
            pFileName = stAttPath
            GetAttachFile = True
            Exit Function
        End If
    Else
        stAttPath = pSavePath
    End If
       
    strExc(0) = "Select * From CustWebFile b Where cf01=" & textCW(1) & " and cf02='" & ChgSQL(pFileName) & "'"
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
       If Dir(stAttPath) <> "" Then Kill stAttPath
       
       'Add By Sindy 2017/5/25
      If "" & RsTemp.Fields("cf08") <> "" Then
         GetAttachFile = PUB_GetFtpFile(RsTemp.Fields("cf08"), stAttPath, UCase("custwebfile"))
      Else
      '2017/5/25 END
         With RsTemp
         lngSize = Val(.Fields("cf03").Value)
         ReDim bytes(lngSize)
         If lngSize > 0 Then bytes() = .Fields("cf04").GetChunk(lngSize)
         End With
         iFileNo = FreeFile
         Open stAttPath For Binary Access Write As #iFileNo
         If lngSize > 0 Then Put #iFileNo, , bytes()
         Close #iFileNo
       End If
       
       pFileName = stAttPath
       GetAttachFile = True
    End If
    Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

'查詢帳號資料
Private Sub QueryCustWebId(strCD02 As String)
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    Dim strWhere As String, strID As String, strText As String
    
    grd1.Clear
    grd1.Rows = 2
    Call SetGrd(bolMuchCust)
    If strCD02 > "" Then
        strWhere = " And cd02='" & strCD02 & "'"
    End If
    If Pub_StrUserSt03 <> "M51" Then
        strWhere = strWhere & "And (InStr(cd06,'" & strUserNum & "')>0 or cd06 is null) "
    End If
    strQ = "Select ' ' as V,cd02 as 客戶編號,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as 客戶名稱,cd03 as 帳號,cd04 as 密碼,Decode(cd05,'1','1 管理者','2','2 使用者') as 身份別,cd06 as 使用者,sqldatet(cd09) as 建置日期,sqldatet(cd07) as 下次更新日期,cd08 as 註解,cd06 as CD06 From CustWebId,Customer Where cd01='" & textCW(1) & "'" & strWhere & " And substr(cd02,1,1)='X' And cu01>' ' And cd02=cu01||cu02 " & _
     "Union Select ' ' as V,cd02 as 客戶編號,NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06)) as 客戶名稱,cd03 as 帳號,cd04 as 密碼,Decode(cd05,'1','1 管理者','2','2 使用者') as 身份別,cd06 as 使用者,sqldatet(cd09) as 建置日期,sqldatet(cd07) as 下次更新日期,cd08 as 註解,cd06 as CD06 From CustWebId,Fagent Where cd01='" & textCW(1) & "'" & strWhere & " And substr(cd02,1,1)='Y' And fa01>' ' And cd02=fa01||fa02 " & _
     "Union Select ' ' as V,cd02 as 客戶編號,cd02 as 客戶名稱,cd03 as 帳號,cd04 as 密碼,Decode(cd05,'1','1 管理者','2','2 使用者') as 身份別,cd06 as 使用者,sqldatet(cd09) as 建置日期,sqldatet(cd07) as 下次更新日期,cd08 as 註解,cd06 as CD06 From CustWebId Where cd01='" & textCW(1) & "'" & strWhere & " And substr(cd02,1,1)<>'X' And substr(cd02,1,1)<>'Y' " & _
     "Order by 客戶編號,身份別,帳號 asc"
     intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        Set grd1.Recordset = RsQ
        For i = 1 To grd1.Rows - 1
            '轉換使用者姓名
            strID = grd1.TextMatrix(i, 6)
            If ChangeCD06CN(strID, strText) = True Then
                grd1.TextMatrix(i, 6) = strText
            Else
                grd1.TextMatrix(i, 6) = strID
            End If
        Next i
    End If
    RsQ.Close
    Set RsQ = Nothing
End Sub

'轉換使用者姓名
Private Function ChangeCD06CN(strID As String, ByRef strText As String) As Boolean
    Dim strTempName As String
   
    ChangeCD06CN = False
    If strID <> "" Then
       MyArr = Split(strID, ",")
       strText = ""
       For j = 0 To UBound(MyArr)
          If ClsPDGetStaff(MyArr(j), strTempName) = True Then
             strText = strText & "," & strTempName
          Else
             strText = strText & "," & MyArr(j)
          End If
       Next j
       strText = Mid(strText, 2, Len(strText))
       ChangeCD06CN = True
    End If
End Function

'Add by Amy 2015/04/27 +複製區
Private Sub grd1_SelChange()
    Dim tmpMouseRow
    
    grd1.Visible = False
    tmpMouseRow = grd1.row
    grd1.Visible = True
    
    If tmpMouseRow <> 0 And grd1.TextMatrix(tmpMouseRow, 3) <> "" Then
        grd1.col = 0
        grd1.Visible = False
        For j = 1 To grd1.Rows - 1
            grd1.row = j
            If tmpMouseRow = j Then
                grd1.Text = "V"
                For i = 0 To grd1.Cols - 1
                    grd1.col = i
                    grd1.CellBackColor = &HFFC0C0
                Next i
            Else
                If grd1.Text = "V" Then grd1.Text = ""
                For i = 0 To grd1.Cols - 1
                    grd1.col = i
                    grd1.CellBackColor = QBColor(15)
                Next i
            End If
         Next j
     
         textCD03 = grd1.TextMatrix(tmpMouseRow, 3) '帳號
         textCD04 = grd1.TextMatrix(tmpMouseRow, 4) '密碼
         grd1.Visible = True
   End If
End Sub

Private Sub lstUsers_Click(Index As Integer)
    Dim m_CurrKeyCD02 As String
    
    Select Case Index
        Case 2
            If lstUsers(Index).ListIndex < 0 Then
                Call QueryCustWebId(textCW(1)) '查詢共同帳號資料
            Else
               MyArr = Split(lstUsers(Index).List(lstUsers(Index).ListIndex), "@")
               m_CurrKeyCD02 = Trim(MyArr(1))
               Call QueryCustWebId(m_CurrKeyCD02) '查詢該客戶的帳號資料
            End If
            'Add by Amy 2015/04/27 +複製區
            textCD03.Text = Empty
            textCD04.Text = Empty
            Exit Sub
    End Select
End Sub

Private Function BrowseForFolder(Optional sCaption As String = "請選擇欲儲存的位置", Optional sDefault As String) As String
    Const BIF_RETURNONLYFSDIRS = 1
    Const MAX_PATH = 260
    Dim lPos As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, tBrowse As BrowseInfo

    With tBrowse
        'Set the owner window
        .hwndOwner = GetActiveWindow        'Me.hWnd in VB
        .lpszTitle = sCaption
        .ulFlags = BIF_RETURNONLYFSDIRS     'Return only if the user selected a directory
    End With

    'Show the dialog
    lpIDList = SHBrowseForFolder(tBrowse)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        lPos = InStr(sPath, vbNullChar)
        If lPos Then
            BrowseForFolder = Left$(sPath, lPos - 1)
            If Right$(BrowseForFolder, 1) <> "\" Then
                BrowseForFolder = BrowseForFolder & "\"
            End If
        End If
    Else
        'User cancelled, return default path
        BrowseForFolder = sDefault
    End If
End Function

Private Function GetSaveName(ByVal pFileName As String) As String
   
On Error GoTo ErrHnd

   With CommonDialog1
      .CancelError = True
      .FileName = pFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowSave
      If .FileName <> "" Then
         GetSaveName = .FileName
      End If
   End With
   
   Exit Function
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function

'進入網站
Private Sub OpenIE(strWebAddr As String)
   Dim myweb As Object
   Dim hLocalFile As Long
   
   'Modify By Sindy 2021/8/25 薛:平台資料0041,請修改為由chrome 開始。（網站已不支援ＩＥ）
   '  Tymetrix360 指定用Chrome開啟--經理
   If textCW(1) = "0026" Or textCW(1) = "0041" Then
      PUB_OpenURL strWebAddr, 1
   Else
      'Modify By Sindy 2020/12/29
      ShellExecute hLocalFile, "open", strWebAddr, vbNullString, vbNullString, 1
   End If
   Exit Sub
   '2020/12/29 END
   
   Set myweb = CreateObject("InternetExplorer.Application")
   Screen.MousePointer = vbHourglass
   With myweb
      .Toolbar = 0
      .Visible = True ' 顯示IE
      .Navigate strWebAddr ' 瀏覽網址 www.lativ.com.tw/Home/Login
   End With
   Set myweb = Nothing ' 釋放IE 物件
   Screen.MousePointer = vbDefault
End Sub

Private Sub textCW_DblClick(Index As Integer, Cancel As MSForms.ReturnBoolean)
   If Index <> 2 And Index <> 17 And Index <> 18 Then Exit Sub
   If Trim(textCW(Index).Text) <> "" Then
      Call OpenIE(Trim(textCW(Index).Text))
   End If
End Sub
