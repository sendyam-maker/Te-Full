VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090634 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人衍生工作記錄維護"
   ClientHeight    =   5460
   ClientLeft      =   6096
   ClientTop       =   1548
   ClientWidth     =   9132
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9132
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   30
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090634.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090634.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090634.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090634.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090634.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090634.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090634.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090634.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090634.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090634.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090634.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4605
      Left            =   180
      TabIndex        =   26
      Top             =   720
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   8128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm090634.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCaseCnt"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(155)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(154)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(55)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(6)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(7)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblSANo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Combo1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblSAName"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtEH(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtEH(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtEH(8)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtEH(9)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtEH(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtEH(5)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtEH(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtEH(4)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtEH(14)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Check1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdSelCp09"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Check2"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm090634.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(4)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(5)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(8)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtEH(10)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtEH(11)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtEH(12)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtEH(13)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(9)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtEH(15)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label5"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "grdList"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdQuery(0)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdQuery(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Check3"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtCode(0)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtCode(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtCode(2)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtSystem"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      Begin VB.TextBox txtSystem 
         Height          =   300
         Left            =   -74040
         MaxLength       =   3
         TabIndex        =   18
         Top             =   900
         Width           =   732
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   2
         Left            =   -71610
         MaxLength       =   2
         TabIndex        =   21
         Top             =   900
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   1
         Left            =   -72030
         MaxLength       =   1
         TabIndex        =   20
         Top             =   900
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   0
         Left            =   -73275
         MaxLength       =   6
         TabIndex        =   19
         Top             =   900
         Width           =   1212
      End
      Begin VB.CheckBox Check3 
         Caption         =   "僅查管制中案件未輸入本所案號或收文號"
         Height          =   195
         Left            =   -71835
         TabIndex        =   17
         Top             =   630
         Width           =   3600
      End
      Begin VB.CheckBox Check2 
         Caption         =   "管制案件"
         Height          =   225
         Left            =   2520
         TabIndex        =   12
         Top             =   3705
         Width           =   1605
      End
      Begin VB.CommandButton cmdSelCp09 
         Caption         =   "選擇收文號"
         Height          =   300
         Left            =   3300
         TabIndex        =   9
         Top             =   1590
         Width           =   1200
      End
      Begin VB.CheckBox Check1 
         Caption         =   "主管核可"
         Height          =   315
         Left            =   540
         TabIndex        =   11
         Top             =   3660
         Width           =   1785
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "列印(&P)"
         Height          =   400
         Index           =   1
         Left            =   -67230
         TabIndex        =   24
         Top             =   390
         Width           =   912
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Default         =   -1  'True
         Height          =   400
         Index           =   0
         Left            =   -68190
         TabIndex        =   23
         Top             =   390
         Width           =   912
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   3240
         Left            =   -74820
         TabIndex        =   45
         Top             =   1290
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   5715
         _Version        =   393216
         Cols            =   13
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
         _Band(0).Cols   =   13
      End
      Begin MSForms.Label Label5 
         Height          =   255
         Left            =   -69000
         TabIndex        =   47
         Top             =   960
         Width           =   1305
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2302;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   15
         Left            =   -69990
         TabIndex        =   22
         Top             =   900
         Width           =   885
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作人員:"
         Height          =   180
         Index           =   9
         Left            =   -70830
         TabIndex        =   46
         Top             =   960
         Width           =   765
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   14
         Left            =   1410
         TabIndex        =   4
         Top             =   1620
         Width           =   1305
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "2302;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   4
         Left            =   1410
         TabIndex        =   3
         Top             =   1290
         Width           =   945
         VariousPropertyBits=   671105051
         MaxLength       =   5
         Size            =   "1667;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   0
         Left            =   1410
         TabIndex        =   1
         Top             =   570
         Width           =   945
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1667;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   5
         Left            =   1410
         TabIndex        =   5
         Top             =   1950
         Width           =   525
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   3
         Left            =   4650
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   570
         Width           =   525
         VariousPropertyBits=   671105049
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   825
         Index           =   9
         Left            =   1380
         TabIndex        =   10
         Top             =   2670
         Width           =   5805
         VariousPropertyBits=   -1466941413
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "10239;1455"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   8
         Left            =   3450
         TabIndex        =   8
         Top             =   1950
         Width           =   435
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "767;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   7
         Left            =   3030
         TabIndex        =   7
         Top             =   1950
         Width           =   315
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "556;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   6
         Left            =   2040
         TabIndex        =   6
         Top             =   1950
         Width           =   915
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1614;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   13
         Left            =   -70050
         TabIndex        =   16
         Top             =   360
         Width           =   525
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   12
         Left            =   -70680
         TabIndex        =   15
         Top             =   360
         Width           =   525
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   11
         Left            =   -72990
         TabIndex        =   14
         Top             =   450
         Width           =   945
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1667;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEH 
         Height          =   300
         Index           =   10
         Left            =   -74040
         TabIndex        =   13
         Top             =   450
         Width           =   945
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1667;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblSAName 
         Height          =   255
         Left            =   2130
         TabIndex        =   44
         Top             =   2340
         Width           =   1845
         VariousPropertyBits=   27
         Caption         =   "lblSAName"
         Size            =   "3254;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   300
         Left            =   1410
         TabIndex        =   2
         Top             =   930
         Width           =   1800
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3175;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Left            =   540
         TabIndex        =   43
         Top             =   4080
         Width           =   3300
         VariousPropertyBits=   27
         Caption         =   "lblFM2-Create :"
         Size            =   "5821;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label4 
         Height          =   255
         Left            =   4020
         TabIndex        =   42
         Top             =   4080
         Width           =   3300
         VariousPropertyBits=   27
         Caption         =   "lblFM2-Update :"
         Size            =   "5821;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblSANo 
         AutoSize        =   -1  'True
         Caption         =   "lblSANo"
         Height          =   255
         Left            =   1440
         TabIndex        =   41
         Top             =   2340
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   8
         Left            =   -74850
         TabIndex        =   40
         Top             =   945
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收文號:"
         Height          =   255
         Index           =   7
         Left            =   510
         TabIndex        =   39
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "序號:"
         Height          =   180
         Index           =   6
         Left            =   3810
         TabIndex        =   38
         Top             =   630
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作日期:"
         Height          =   255
         Index           =   55
         Left            =   510
         TabIndex        =   35
         Top             =   630
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   255
         Index           =   154
         Left            =   510
         TabIndex        =   34
         Top             =   2010
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作人員:"
         Height          =   255
         Index           =   155
         Left            =   510
         TabIndex        =   33
         Top             =   990
         Width           =   765
      End
      Begin VB.Line Line1 
         X1              =   1590
         X2              =   3720
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員:"
         Height          =   255
         Index           =   0
         Left            =   510
         TabIndex        =   32
         Top             =   2340
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作時數:"
         Height          =   255
         Index           =   1
         Left            =   510
         TabIndex        =   31
         Top             =   1350
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註:"
         Height          =   255
         Index           =   2
         Left            =   510
         TabIndex        =   30
         Top             =   2700
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "折算基數:"
         Height          =   255
         Index           =   3
         Left            =   3810
         TabIndex        =   29
         Top             =   1350
         Width           =   765
      End
      Begin VB.Label lblCaseCnt 
         AutoSize        =   -1  'True
         Caption         =   "lblCaseCnt"
         Height          =   255
         Left            =   4650
         TabIndex        =   28
         Top             =   1350
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(EX : 123.5) "
         Height          =   255
         Left            =   2430
         TabIndex        =   27
         Top             =   1350
         Width           =   930
      End
      Begin VB.Line Line3 
         X1              =   -70380
         X2              =   -69810
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作人員部門:"
         Height          =   180
         Index           =   5
         Left            =   -71850
         TabIndex        =   37
         Top             =   390
         Width           =   1125
      End
      Begin VB.Line Line2 
         X1              =   -73260
         X2              =   -72840
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作日期:"
         Height          =   180
         Index           =   4
         Left            =   -74850
         TabIndex        =   36
         Top             =   480
         Width           =   765
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   9135
      _ExtentX        =   16108
      _ExtentY        =   1080
      ButtonWidth     =   1138
      ButtonHeight    =   1032
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
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
End
Attribute VB_Name = "frm090634"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/030 改成Form2.0 ; grdList改字型=新細明體-ExtB(MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉)、Combo1、Label3、Label4、lblSAName、txtEH(index); Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Copied from frm090623 by Morgan 2011/7/29
Option Explicit

Dim EH(1 To 19) As String
Dim strRsStart1 As String, strRsStart2 As String, strRsStart4 As String, strRsEnd1 As String, strRsEnd2 As String, strRsEnd4 As String
Dim rsDefineSize As New ADODB.Recordset
Dim intWhere As Integer
Dim ActionEdit As Integer
Dim intRow As Integer
'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_CurrSel As Integer
Dim PLeft(0 To 8) As Integer
Dim iPrint As Integer
Dim Page As Integer
Dim StrSQLa As String
Dim Seek_Now_Cp09 As String
Dim intTemp As Boolean
Public p_iRtn As Integer
Dim bolMail2Boss As Boolean
Dim bolMail2Boss2 As Boolean
Dim m_blnColOrderAsc As Boolean 'Added by Lydia 2016/08/11 欄位資料由小到大排序
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限

Public Sub SelectToolbarButtom()
   Dim btn
   '設定為按下查詢鈕扭
   Set btn = Me.TBar1.Buttons(4)
   Tbar1_ButtonClick btn
End Sub

Private Sub Check1_Click()
    '若從個人進入, 若已核可的資料其他欄位不可更改
    If ProState = "1" Then
        '個人進入時, 只可修改個人輸入的資料, 且不可改工作人員
        'Modify By Sindy 2014/7/31
        'If Me.Check1.Value = vbChecked Or (txtEH(0) <> "" And txtEH(1) <> strUserNum) Then
        If Me.Check1.Value = vbChecked Or (txtEH(0) <> "" And Trim(Left(Combo1.Text, 6)) <> strUserNum) Then
        '2014/7/31 END
            Me.txtEH(0).Enabled = False
            'Me.txtEH(1).Enabled = False
            Combo1.Enabled = False
            Me.txtEH(3).Enabled = False
            Me.txtEH(4).Enabled = False
            Me.txtEH(5).Enabled = False
            Me.txtEH(6).Enabled = False
            Me.txtEH(7).Enabled = False
            Me.txtEH(8).Enabled = False
            Me.txtEH(9).Enabled = False
            Me.txtEH(14).Enabled = False
            cmdSelCp09.Enabled = False
        Else
            Me.txtEH(0).Enabled = True
            'Me.txtEH(1).Enabled = True
            Combo1.Enabled = True
            Me.txtEH(3).Enabled = False
            Me.txtEH(4).Enabled = True
            Me.txtEH(5).Enabled = True
            Me.txtEH(6).Enabled = True
            Me.txtEH(7).Enabled = True
            Me.txtEH(8).Enabled = True
            Me.txtEH(9).Enabled = True
            Me.txtEH(14).Enabled = True
            cmdSelCp09.Enabled = True
        End If
    End If
   
End Sub

'是否衍生工作時數累計超過 3 小時
Private Function ChkTotalOver() As Boolean
   Dim strTot As String
   
   strSql = "select sum(eh05) from extendhour where eh12='" & txtEH(14) & "'"
   
   If ActionEdit = 1 Then
      strSql = strSql & " and not (eh01=" & DBDATE(txtEH(0).Tag) & _
         " and eh02='" & Trim(Left(Combo1.Tag, 6)) & "' and eh04='" & txtEH(3).Tag & "')"
   End If
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strTot = "" & RsTemp.Fields(0)
      If Val(strTot) > 0 Then
         strTot = Val(strTot) + Val(txtEH(4))
         If Val(strTot) > 3 Then
            ChkTotalOver = True
         End If
      End If
   End If
End Function

Private Sub cmdQuery_Click(Index As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
    
    '若頁籤在 基本資料，就不管
    If SSTab1.Tab = 0 Then Exit Sub
    
    If Me.txtEH(10).Text = "" And txtSystem.Text = "" And txtCode(0).Text = "" Then
        MsgBox "請輸入工作日期範圍或本所案號條件!!!", vbExclamation + vbOKOnly
        Me.txtEH(10).SetFocus
        Exit Sub
    End If
    If Me.txtEH(10).Text <> "" Then
       If CheckIsTaiwanDate(Me.txtEH(10).Text) = False Then
          Me.txtEH(10).SetFocus
          txtEH_GotFocus 10
          Exit Sub
       End If
    End If
    If Me.txtEH(10).Text <> "" And Me.txtEH(11).Text = "" Then
        MsgBox "請輸入工作迄日!!!", vbExclamation + vbOKOnly
        Me.txtEH(11).SetFocus
        Exit Sub
    End If
    If Me.txtEH(11).Text <> "" Then
       If CheckIsTaiwanDate(Me.txtEH(11).Text) = False Then
          Me.txtEH(11).SetFocus
          Exit Sub
       End If
    End If
    If Val(Me.txtEH(10).Text) > Val(Me.txtEH(11).Text) Then
        MsgBox "工作日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
        Me.txtEH(10).SetFocus
        txtEH_GotFocus 10
        Exit Sub
    End If
    If Me.txtEH(12).Text <> "" And Me.txtEH(13).Text <> "" Then
        If Me.txtEH(12).Text > Me.txtEH(13).Text Then
            MsgBox "工作人員部門範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.txtEH(12).SetFocus
            txtEH_GotFocus 12
            Exit Sub
        End If
    End If
    
    If txtSystem.Text <> "" Or txtCode(0).Text <> "" Then
      If txtSystem.Text = "" Or txtCode(0).Text = "" Then
         MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
         If txtSystem.Text = "" Then
            Me.txtSystem.SetFocus
         End If
         If txtCode(0).Text = "" Then
            Me.txtCode(0).SetFocus
         End If
         Exit Sub
      End If
    End If
    
    '查詢
    If Index = 0 Then
        m_blnColOrderAsc = True 'Added by Lydia 2016/08/11 欄位資料由小到大排序
        Screen.MousePointer = vbHourglass
        Me.grdList.MousePointer = flexHourglass
        If QueryData() = False Then
            strTit = "查詢資料"
            strMsg = "無資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        End If
        Me.grdList.MousePointer = flexDefault
        Screen.MousePointer = vbDefault
    '列印
    Else
        If Val(Me.grdList.Rows) > 1 And Trim(grdList.TextMatrix(1, 1)) <> "" Then
            Screen.MousePointer = vbHourglass
            PrintData
            ShowPrintOk
            Screen.MousePointer = vbDefault
        Else
            ShowNoData
        End If
    End If
End Sub

Private Sub cmdSelCp09_Click()
   If Trim(txtEH(5)) <> "" And Trim(txtEH(6)) <> "" Then
      Load frm090634_1
      frm090634_1.Hide
      frm090634_1.oCP01 = txtEH(5).Text
      frm090634_1.oCP02 = txtEH(6).Text
      frm090634_1.oCP03 = IIf(Trim(txtEH(7).Text) = "", "0", txtEH(7).Text)
      frm090634_1.oCP04 = IIf(Trim(txtEH(8).Text) = "", "00", txtEH(8).Text)
      If frm090634_1.Process = False Then ShowNoData: Exit Sub
      frm090634_1.Show vbModal
      Unload frm090634_1
   Else
      MsgBox "請先輸入本所案號！", vbExclamation, "警告！"
      If Me.txtEH(5).Enabled = True Then Me.txtEH(5).SetFocus
   End If
End Sub

'Add By Sindy 2014/7/31
'Modified by Lydia 2022/01/030 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Combo1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo1_LostFocus()
   If Combo1 <> "" Then
      Combo1 = Trim(Left(Combo1, 6)) & " " & GetPrjSalesNM(Trim(Left(Combo1, 6)))
   End If
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
Dim strEmp As String
   
   If Combo1 <> "" Then
      strEmp = GetStaffName(Trim(Left(Combo1, 6)))
      If strEmp = "" And Check1.Value = vbUnchecked Then
         MsgBox "工作人員輸入錯誤!!!", vbExclamation + vbOKOnly
         Cancel = True
      End If
   End If
End Sub
'2014/7/31 END

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5
         If ActionEdit = 3 Then
            Select Case KeyCode
               Case vbKeyF2
                  RsSitu 0
               Case vbKeyF3
                  RsSitu 1
               Case vbKeyF5
                  RsSitu 2
               Case vbKeyF4
                  RsSitu 5
            End Select
            KeyCode = 0
         End If
      'Modified by Lydia 2022/01/03 去掉Enter鍵vbKeyReturn
      'Case vbKeyF9, vbKeyF10, vbKeyReturn
      Case vbKeyF9, vbKeyF10
         If ActionEdit <> 3 Then
            Select Case KeyCode
               'Modified by Lydia 2022/01/03 去掉Enter鍵vbKeyReturn
               'Case vbKeyF9, vbKeyReturn
               Case vbKeyF9
                  RsSitu 3
               Case vbKeyF10
                  RsSitu 4
            End Select
            KeyCode = 0
         End If
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd
         If ActionEdit = 3 Then
            Select Case KeyCode
               Case vbKeyHome
                  RsAction 0
               Case vbKeyPageUp
                  RsAction 1
               Case vbKeyPageDown
                  RsAction 2
               Case vbKeyEnd
                  RsAction 3
            End Select
            KeyCode = 0
         End If
    Case vbKeyEscape
        If MsgBox("是否確定結束?", vbYesNo + vbCritical) = vbYes Then Unload Me
    Case Else
        Exit Sub
    End Select
   If KeyCode <> vbKeyF2 And KeyCode <> vbKeyF3 And KeyCode <> vbKeyF4 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyEscape Then
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
         'Added by Lydia 2022/05/18
         If m_bQuery Then
             TBar1.Buttons(4).Enabled = True
         Else
             TBar1.Buttons(4).Enabled = False
         End If
         'end 2022/05/18
   End If
End Sub

Private Sub Form_Load()
   Dim i As Integer
    
    m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
    MoveFormToCenter Me
    '取得使用者執行各項功能的權限
    '由個人進入
    If ProState = "1" Then
        m_bInsert = IsUserHasRightOfFunction("frm090634P", strAdd, False)
        m_bUpdate = IsUserHasRightOfFunction("frm090634P", strEdit, False)
        m_bDelete = IsUserHasRightOfFunction("frm090634P", strDel, False)
        m_bQuery = IsUserHasRightOfFunction("frm090634P", strFind, False)
        cmdQuery(1).Enabled = IsUserHasRightOfFunction("frm090634P", strPrint, False)
        Check1.Enabled = False
        Check2.Enabled = False
        
    '由管理進入
    Else
        m_bInsert = IsUserHasRightOfFunction("frm090634M", strAdd, False)
        m_bUpdate = IsUserHasRightOfFunction("frm090634M", strEdit, False)
        m_bDelete = IsUserHasRightOfFunction("frm090634M", strDel, False)
        m_bQuery = IsUserHasRightOfFunction("frm090634M", strFind, False)
        cmdQuery(1).Enabled = IsUserHasRightOfFunction("frm090634M", strPrint, False)
        Check1.Enabled = True
        Check2.Enabled = True
        
    End If
    Call SetCombo1 'Add By Sindy 2014/7/31
    If Val(strSrvDate(1)) >= 20140401 Then m_bInsert = False 'Added by Morgan 2014/3/19 4/1起取消新增功能

    strExc(0) = "SELECT * FROM ExtendHour WHERE ROWNUM<1"
    intI = 1
    Set rsDefineSize = ClsLawReadRstMsg(intI, strExc(0))
   strRsStart1 = Empty: strRsStart2 = Empty: strRsStart4 = Empty
   strRsEnd1 = Empty: strRsEnd2 = Empty: strRsEnd4 = Empty
   strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour Order By EH01, EH02, EH04 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
   If intI = 1 Then
        RsTemp.MoveFirst
      strRsStart1 = "" & RsTemp.Fields("EH01").Value
      strRsStart2 = "" & RsTemp.Fields("EH02").Value
      strRsStart4 = "" & RsTemp.Fields("EH04").Value
        RsTemp.MoveLast
      strRsEnd1 = "" & RsTemp.Fields("EH01").Value
      strRsEnd2 = "" & RsTemp.Fields("EH02").Value
      strRsEnd4 = "" & RsTemp.Fields("EH04").Value
      RsAction 0
   End If
   ActionEdit = 3
   CmdSitu True
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
   'Added by Lydia 2022/05/18
   If m_bQuery Then
       TBar1.Buttons(4).Enabled = True
   Else
       TBar1.Buttons(4).Enabled = False
   End If
   'end 2022/05/18
   TxtLock 3
   
   SSTab1.Tab = 0
   
   Label5.Caption = "" 'Added by Lydia 2022/05/18
End Sub

'Add By Sindy 2014/7/31
Private Sub SetCombo1()
Dim strTemp As String, arrData As Variant, i As Integer
   Combo1.Clear
   Combo1.AddItem strUserNum & " " & strUserName
   '檢查當時是否需要為他人職代
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)

   '開放部份智權同仁的資料給彥葶操作
   If Pub_GetSpecMan("A8") = strUserNum Then
      strTemp = Pub_GetSpecMan("A7")
      arrData = Split(strTemp, ";")
      For i = 0 To UBound(arrData)
         Combo1.AddItem arrData(i) & " " & GetPrjSalesNM(CStr(arrData(i)))
      Next
   End If
   Combo1.Text = Combo1.List(0)
End Sub

Private Function ReadExtendHour(ByRef tsTmp() As String) As Boolean
   Dim i As Integer, j As Integer, Lbl As Label, txt As TextBox, strTmp As String
   Dim strTxt(0 To 4) As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
   strTxt(1) = tsTmp(1): strTxt(2) = tsTmp(2): strTxt(4) = tsTmp(4)
   EH(1) = strTxt(1): EH(2) = strTxt(2):  EH(4) = strTxt(4)
   For i = 0 To 10
       If i = 10 Then
           Me.Check1.Value = vbUnchecked
       ElseIf i <> 1 And i <> 2 Then 'Modify By Sindy 2014/7/31 +i <> 1 and
           Me.txtEH(i).Text = ""
       End If
   Next i
   Me.txtEH(14).Text = ""
   Me.lblCaseCnt.Caption = ""
   Me.lblSANo.Caption = ""
   Me.lblSAName.Caption = ""
   'Me.lblSupName.Caption = ""
   Me.Label3.Caption = "Create : "
   Me.Label4.Caption = "Update : "
   Check2.Value = vbUnchecked
   If EH(1) = "" Then Exit Function
   StrSQLa = "Select * From ExtendHour,caseprogress Where EH01=" & EH(1) & " And EH02='" & EH(2) & "' And EH04='" & EH(4) & "' and cp09(+)=EH12 Order By EH01, EH02, EH04 "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
        EH(1) = "" & rsA.Fields("EH01").Value
        EH(2) = "" & rsA.Fields("EH02").Value
        EH(4) = "" & rsA.Fields("EH04").Value
        EH(5) = "" & rsA.Fields("EH05").Value
        EH(6) = "" & rsA.Fields("EH06").Value
        EH(7) = "" & rsA.Fields("EH07").Value
        EH(8) = "" & rsA.Fields("EH08").Value
        EH(9) = "" & rsA.Fields("EH09").Value
        EH(10) = "" & rsA.Fields("EH10").Value
        EH(11) = "" & rsA.Fields("EH11").Value
        EH(12) = "" & rsA.Fields("EH12").Value
        EH(13) = "" & rsA.Fields("EH13").Value
        EH(14) = "" & rsA.Fields("EH14").Value
        EH(15) = "" & rsA.Fields("EH15").Value
        EH(16) = "" & rsA.Fields("EH16").Value
        EH(17) = "" & rsA.Fields("EH17").Value
        EH(18) = "" & rsA.Fields("EH18").Value
        lblSANo = "" & rsA.Fields("CP13").Value
        EH(19) = "" & rsA.Fields("EH19").Value
    Else
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Function
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   Me.txtEH(0).Text = ChangeWStringToTString(EH(1))
   Combo1.Text = EH(2) & " " & GetStaffName(EH(2), True) 'Add By Sindy 2014/7/31
'   Me.txtEH(1).Text = EH(2)
'   Me.lblSupName.Caption = GetStaffName(EH(2), True)
   Me.lblSAName.Caption = GetStaffName(lblSANo, True)
   Me.txtEH(3).Text = EH(4)
   Me.txtEH(4).Text = EH(5)
   Me.txtEH(5).Text = EH(6)
   Me.txtEH(6).Text = EH(7)
   Me.txtEH(7).Text = EH(8)
   Me.txtEH(8).Text = EH(9)
   Me.txtEH(9).Text = EH(10)
   Me.lblCaseCnt.Caption = Format(Val(Me.txtEH(4).Text) * 0.25, "0.00")
   Me.Check1.Value = IIf(EH(11) <> "", vbChecked, vbUnchecked)
   Me.txtEH(14).Text = EH(12)
   Call Check1_Click
   Me.Check2.Value = IIf(EH(19) <> "", vbChecked, vbUnchecked)
   
   Me.txtEH(0).Tag = Me.txtEH(0).Text
   'Me.txtEH(1).Tag = Me.txtEH(1).Text
   Combo1.Tag = Combo1.Text
   Me.txtEH(3).Tag = Me.txtEH(3).Text
   Me.txtEH(4).Tag = Me.txtEH(4).Text
   Me.txtEH(5).Tag = Me.txtEH(5).Text
   Me.txtEH(6).Tag = Me.txtEH(6).Text
   Me.txtEH(7).Tag = Me.txtEH(7).Text
   Me.txtEH(8).Tag = Me.txtEH(8).Text
   Me.txtEH(9).Tag = Me.txtEH(9).Text
   Me.Check1.Tag = Me.Check1.Value
   Me.Check2.Tag = Me.Check2.Value
   Me.txtEH(14).Tag = Me.txtEH(14).Text
   If EH(13) <> "" Then
       Me.Label3.Caption = Me.Label3.Caption & GetStaffName(EH(13))
   End If
   If EH(14) <> "" Then
       Me.Label3.Caption = Me.Label3.Caption & " " & ChangeTStringToTDateString(Val(EH(14)) - 19110000)
   End If
   If EH(15) <> "" Then
       Me.Label3.Caption = Me.Label3.Caption & " " & Format(EH(15), "##:##")
   End If
   If EH(16) <> "" Then
       Me.Label4.Caption = Me.Label4.Caption & GetStaffName(EH(16))
   End If
   If EH(17) <> "" Then
       Me.Label4.Caption = Me.Label4.Caption & " " & ChangeTStringToTDateString(Val(EH(17)) - 19110000)
   End If
   If EH(18) <> "" Then
       Me.Label4.Caption = Me.Label4.Caption & " " & Format(EH(18), "##:##")
   End If
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm090634 = Nothing
End Sub

Private Sub RsSitu(ByVal Situ As Integer)
   Dim i As Integer, St1 As String, St2 As String
   Dim TBmk As Variant
   Dim StrSQLa As String
   Dim EH04 As String
 
 On Error GoTo CheckingErr
 
 Static TmpEH(4) As String
   Select Case Situ
      Case 0 '按下新增add
        TmpEH(1) = ChangeTStringToWString(Me.txtEH(0).Text)
        'TmpEH(2) = Me.txtEH(1).Text
        TmpEH(2) = Trim(Left(Combo1.Text, 6))
        TmpEH(4) = Me.txtEH(3).Text
        Me.lblCaseCnt.Caption = ""
        Me.lblSANo.Caption = ""
        Me.lblSAName.Caption = ""
        Me.lblCaseCnt.Caption = ""
        Me.Label3.Caption = "Create : "
        Me.Label4.Caption = "Update : "
        CmdSitu False
        TxtLock 0
        ActionEdit = 0
        If Me.txtEH(0).Enabled = True Then Me.txtEH(0).SetFocus
        txtEH_GotFocus 0
        'Modify By Sindy 2014/7/31
        Combo1.ListIndex = 0
        'Combo1.Locked = True
        '2014/7/31 END
'        Me.txtEH(1).Text = strUserNum
'        Me.txtEH(1).Locked = True
'        Me.lblSupName.Caption = GetStaffName(Me.txtEH(1).Text)
        Seek_Now_Cp09 = ""
        Call Check1_Click
        
      Case 1 '按下修改modi
         CmdSitu False
         TxtLock 1
         ActionEdit = 1
        TmpEH(1) = ChangeTStringToWString(Me.txtEH(0).Text)
        'TmpEH(2) = Me.txtEH(1).Text
        TmpEH(2) = Trim(Left(Combo1.Text, 6))
        TmpEH(4) = Me.txtEH(3).Text
        Seek_Now_Cp09 = txtEH(14).Text
      Case 2 '按下刪除delete
         '若從個人進入, 若已核可的資料不可刪除
         '個人進入時, 只可刪除個人輸入的資料
         'Modify By Sindy 2014/7/31
         'If ProState = "1" And (Me.Check1.Value = vbChecked Or (txtEH(0) <> "" And txtEH(1) <> strUserNum)) Then
         If ProState = "1" And (Me.Check1.Value = vbChecked Or (txtEH(0) <> "" And Trim(Left(Combo1.Text, 6)) <> strUserNum)) Then
         '2014/7/31 END
         Else
             If Me.txtEH(0).Text = "" Then
                 MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
                 Exit Sub
             End If
             If DelMsg Then
                 'Modify By Sindy 2014/7/31
                 'StrSQLa = "Delete From ExtendHour Where EH01=" & ChangeTStringToWString(Me.txtEH(0).Text) & " And EH02='" & Me.txtEH(1).Text & "' And EH04='" & Me.txtEH(3).Text & "' "
                 StrSQLa = "Delete From ExtendHour Where EH01=" & ChangeTStringToWString(Me.txtEH(0).Text) & " And EH02='" & Trim(Left(Combo1.Text, 6)) & "' And EH04='" & Me.txtEH(3).Text & "' "
                 '2014/7/31 END
                 cnnConnection.Execute StrSQLa
                 strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour WHERE EH01||EH02||EH04>='" & EH(1) & EH(2) & EH(4) & "' Order By EH01, EH02, EH04 "
                  intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                 If intI = 1 Then
                    strExc(1) = "" & RsTemp.Fields("EH01").Value
                    strExc(2) = "" & RsTemp.Fields("EH02").Value
                    strExc(4) = "" & RsTemp.Fields("EH04").Value
                    ReadExtendHour strExc
                 Else
                     strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour WHERE EH01||EH02||EH04<='" & EH(1) & EH(2) & EH(4) & "' Order By EH01 Desc , EH02 Desc, EH04 Desc "
                      intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strExc(1) = "" & RsTemp.Fields("EH01").Value
                        strExc(2) = "" & RsTemp.Fields("EH02").Value
                        strExc(4) = "" & RsTemp.Fields("EH04").Value
                        ReadExtendHour strExc
                     Else
                        RsAction 0
                     End If
                 End If
                 strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour Order By EH01, EH02, EH04 "
                 intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
                 If intI = 1 Then
                      RsTemp.MoveFirst
                    strRsStart1 = "" & RsTemp.Fields("EH01").Value
                    strRsStart2 = "" & RsTemp.Fields("EH02").Value
                    strRsStart4 = "" & RsTemp.Fields("EH04").Value
                      RsTemp.MoveLast
                    strRsEnd1 = "" & RsTemp.Fields("EH01").Value
                    strRsEnd2 = "" & RsTemp.Fields("EH02").Value
                    strRsEnd4 = "" & RsTemp.Fields("EH04").Value
                 End If
             End If
         End If
      Case 3 'update
         If ActionEdit = 0 Then '在新增狀態按Enter鍵
            If Not GetData Then Exit Sub
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            '檢查重複
            'Modify By Sindy 2014/7/31
            'strExc(0) = "SELECT EH04 FROM ExtendHour where EH01=" & ChangeTStringToWString(txtEH(0).Text) & " and EH02='" & txtEH(1).Text & "' and EH12='" & txtEH(14).Text & "'"
            strExc(0) = "SELECT EH04 FROM ExtendHour where EH01=" & ChangeTStringToWString(txtEH(0).Text) & " and EH02='" & Trim(Left(Combo1.Text, 6)) & "' and EH12='" & txtEH(14).Text & "'"
            '2014/7/31 END
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               If MsgBox("當天該收文號已經有衍生工作的紀錄，是否繼續？", vbYesNo + vbQuestion, "警告！") = vbNo Then
                  Exit Sub
               End If
            End If
            
            If Me.txtEH(5).Text = "" Or Me.txtEH(6).Text = "" Then
                Me.txtEH(5).Text = "": Me.txtEH(6).Text = "": Me.txtEH(7).Text = "": Me.txtEH(8).Text = ""
            End If
            'Modify By Sindy 2014/7/31
            'Me.txtEH(3).Text = GetSerialNo(ChangeTStringToWString(Me.txtEH(0).Text), Me.txtEH(1).Text)
            Me.txtEH(3).Text = GetSerialNo(ChangeTStringToWString(Me.txtEH(0).Text), Trim(Left(Combo1.Text, 6)))
            '2014/7/31 END
            'Modify By Sindy 2014/7/31
            'StrSQLa = "Insert Into ExtendHour (EH01, EH02, EH04, EH05, EH06, EH07, EH08, EH09, EH10, EH11, EH12) Values(" & ChangeTStringToWString(Me.txtEH(0).Text) & ",'" & Me.txtEH(1).Text & "','" & txtEH(3).Text & "'," & Val(Me.txtEH(4).Text) & ",'" & Me.txtEH(5).Text & "','" & Me.txtEH(6).Text & "','" & Me.txtEH(7).Text & "','" & Me.txtEH(8).Text & "','" & ChgSQL(Me.txtEH(9).Text) & "','" & IIf(Me.Check1.Value = vbChecked, "V", "") & "'," & CNULL(Me.txtEH(14).Text) & ")"
            StrSQLa = "Insert Into ExtendHour (EH01, EH02, EH04, EH05, EH06, EH07, EH08, EH09, EH10, EH11, EH12) Values(" & ChangeTStringToWString(Me.txtEH(0).Text) & ",'" & Trim(Left(Combo1.Text, 6)) & "','" & txtEH(3).Text & "'," & Val(Me.txtEH(4).Text) & ",'" & Me.txtEH(5).Text & "','" & Me.txtEH(6).Text & "','" & Me.txtEH(7).Text & "','" & Me.txtEH(8).Text & "','" & ChgSQL(Me.txtEH(9).Text) & "','" & IIf(Me.Check1.Value = vbChecked, "V", "") & "'," & CNULL(Me.txtEH(14).Text) & ")"
            '2014/7/31 END
            cnnConnection.Execute StrSQLa, intI
         
            If bolMail2Boss = True Then subAddMail
            If bolMail2Boss2 = True Then subAddMail2
            
            ActionEdit = 3
            TxtLock 3
            'Modify By Sindy 2014/7/31
            'If ChangeTStringToWString(Me.txtEH(0).Text) & Me.txtEH(1).Text & Me.txtEH(3).Text < strRsStart1 & strRsStart2 & strRsStart4 Then
            If ChangeTStringToWString(Me.txtEH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtEH(3).Text < strRsStart1 & strRsStart2 & strRsStart4 Then
            '2014/7/31 END
                strRsStart1 = ChangeTStringToWString(Me.txtEH(0).Text)
                'strRsStart2 = Me.txtEH(1).Text
                strRsStart2 = Trim(Left(Combo1.Text, 6))
                strRsStart4 = Me.txtEH(3).Text
            End If
            'Modify By Sindy 2014/7/31
            'If ChangeTStringToWString(Me.txtEH(0).Text) & Me.txtEH(1).Text & Me.txtEH(3).Text > strRsEnd1 & strRsEnd2 & strRsEnd4 Then
            If ChangeTStringToWString(Me.txtEH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtEH(3).Text > strRsEnd1 & strRsEnd2 & strRsEnd4 Then
            '2014/7/31 END
                strRsEnd1 = ChangeTStringToWString(Me.txtEH(0).Text)
                'strRsEnd2 = Me.txtEH(1).Text
                strRsEnd2 = Trim(Left(Combo1.Text, 6))
                strRsEnd4 = Me.txtEH(3).Text
            End If
            strExc(1) = ChangeTStringToWString(Me.txtEH(0).Text)
            'strExc(2) = Me.txtEH(1).Text
            strExc(2) = Trim(Left(Combo1.Text, 6))
            strExc(4) = Me.txtEH(3).Text
            ReadExtendHour strExc
            
         ElseIf ActionEdit = 1 Then '在修改狀態按Enter鍵
            If Not GetData Then Exit Sub
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            '檢查重複
            'Modify By Sindy 2014/7/31
            'If txtEH(0).Tag <> txtEH(0).Text Or txtEH(1).Tag <> txtEH(1).Text Or txtEH(14).Tag <> txtEH(14).Text Then
            If txtEH(0).Tag <> txtEH(0).Text Or Trim(Left(Combo1.Tag, 6)) <> Trim(Left(Combo1.Text, 6)) Or txtEH(14).Tag <> txtEH(14).Text Then
               'strExc(0) = "SELECT EH04 FROM ExtendHour where EH01=" & ChangeTStringToWString(txtEH(0).Text) & " and EH02='" & txtEH(1).Text & "' and EH12='" & txtEH(14).Text & "'"
               strExc(0) = "SELECT EH04 FROM ExtendHour where EH01=" & ChangeTStringToWString(txtEH(0).Text) & " and EH02='" & Trim(Left(Combo1.Text, 6)) & "' and EH12='" & txtEH(14).Text & "'"
            '2014/7/31 END
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
               If intI = 1 Then
                  If MsgBox("當天該收文號已經有衍生工作的紀錄，是否繼續？", vbYesNo + vbQuestion, "警告！") = vbNo Then
                     Exit Sub
                  End If
               End If
            End If
            If Me.txtEH(5).Text = "" Or Me.txtEH(6).Text = "" Then
                Me.txtEH(5).Text = "": Me.txtEH(6).Text = "": Me.txtEH(7).Text = "": Me.txtEH(8).Text = ""
            End If
            StrSQLa = ""
            If Me.txtEH(0).Text <> Me.txtEH(0).Tag Then
                StrSQLa = StrSQLa & " EH01=" & ChangeTStringToWString(Me.txtEH(0).Text) & ","
            End If
            'Modify By Sindy 2014/7/31
            'If Me.txtEH(1).Text <> Me.txtEH(1).Tag Then
            If Trim(Left(Combo1.Text, 6)) <> Trim(Left(Combo1.Tag, 6)) Then
                'StrSQLa = StrSQLa & " EH02='" & Val(Me.txtEH(1).Text) & "',"
                StrSQLa = StrSQLa & " EH02='" & Trim(Left(Combo1.Text, 6)) & "',"
            '2014/7/31 END
            End If
            'Modify By Sindy 2014/7/31
            'If txtEH(0).Tag <> txtEH(0).Text Or txtEH(1).Tag <> txtEH(1).Text Then
            If txtEH(0).Tag <> txtEH(0).Text Or Trim(Left(Combo1.Tag, 6)) <> Trim(Left(Combo1.Text, 6)) Then
               'EH04 = GetSerialNo(ChangeTStringToWString(Me.txtEH(0).Text), Me.txtEH(1).Text)
               EH04 = GetSerialNo(ChangeTStringToWString(Me.txtEH(0).Text), Trim(Left(Combo1.Text, 6)))
            '2014/7/31 END
               StrSQLa = StrSQLa & " EH04='" & EH04 & "',"
            Else
               EH04 = txtEH(3)
            End If
            
            If Me.txtEH(4).Text <> Me.txtEH(4).Tag Then
                StrSQLa = StrSQLa & " EH05=" & Val(Me.txtEH(4).Text) & ","
            End If
            If Me.txtEH(5).Text <> Me.txtEH(5).Tag Then
                StrSQLa = StrSQLa & " EH06='" & Me.txtEH(5).Text & "',"
            End If
            If Me.txtEH(6).Text <> Me.txtEH(6).Tag Then
                StrSQLa = StrSQLa & " EH07='" & Me.txtEH(6).Text & "',"
            End If
            If Me.txtEH(7).Text <> Me.txtEH(7).Tag Then
                StrSQLa = StrSQLa & " EH08='" & Me.txtEH(7).Text & "',"
            End If
            If Me.txtEH(8).Text <> Me.txtEH(8).Tag Then
                StrSQLa = StrSQLa & " EH09='" & Me.txtEH(8).Text & "',"
            End If
            If Me.txtEH(9).Text <> Me.txtEH(9).Tag Then
                StrSQLa = StrSQLa & " EH10='" & Me.txtEH(9).Text & "',"
            End If
            If Me.Check1.Value <> Me.Check1.Tag Then
                StrSQLa = StrSQLa & " EH11='" & IIf(Me.Check1.Value = vbChecked, "V", "") & "',"
            End If
            If Me.txtEH(14).Text <> Me.txtEH(14).Tag Then
                StrSQLa = StrSQLa & " EH12='" & Me.txtEH(14).Text & "',"
            End If

            If Me.Check2.Value <> Check2.Tag Then
                StrSQLa = StrSQLa & " EH19='" & IIf(Me.Check2.Value = vbChecked, "V", "") & "',"
            End If
            
            If StrSQLa <> "" Then
                StrSQLa = Left(StrSQLa, Len(StrSQLa) - 1)
                
            Else
                GoTo NoUpdate
                
            End If
            
   On Error GoTo flgRollback
   
            cnnConnection.BeginTrans
            
            If StrSQLa <> "" Then
               'Modify By Sindy 2014/7/31
               'StrSQLa = "Update ExtendHour Set " & StrSQLa & " Where EH01=" & Val(ChangeTStringToWString(Me.txtEH(0).Tag)) & " And EH02='" & Me.txtEH(1).Tag & "' And EH04='" & Me.txtEH(3).Tag & "' "
               StrSQLa = "Update ExtendHour Set " & StrSQLa & " Where EH01=" & Val(ChangeTStringToWString(Me.txtEH(0).Tag)) & " And EH02='" & Trim(Left(Combo1.Tag, 6)) & "' And EH04='" & Me.txtEH(3).Tag & "' "
               '2014/7/31 END
               cnnConnection.Execute StrSQLa
            End If
            If bolMail2Boss = True Then subAddMail
            If bolMail2Boss2 = True Then subAddMail2
            
            cnnConnection.CommitTrans
            
   On Error GoTo CheckingErr
   
            txtEH(3) = EH04

NoUpdate:
            ActionEdit = 3
            TxtLock 3
            
            strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour Order By EH01, EH02, EH04 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
                 RsTemp.MoveFirst
               strRsStart1 = "" & RsTemp.Fields("EH01").Value
               strRsStart2 = "" & RsTemp.Fields("EH02").Value
               strRsStart4 = "" & RsTemp.Fields("EH04").Value
                 RsTemp.MoveLast
               strRsEnd1 = "" & RsTemp.Fields("EH01").Value
               strRsEnd2 = "" & RsTemp.Fields("EH02").Value
               strRsEnd4 = "" & RsTemp.Fields("EH04").Value
            End If
            strExc(1) = ChangeTStringToWString(Me.txtEH(0).Text)
            'strExc(2) = Me.txtEH(1).Text
            strExc(2) = Trim(Left(Combo1.Text, 6))
            strExc(4) = Me.txtEH(3).Text
            ReadExtendHour strExc
            
         ElseIf ActionEdit = 2 Then '在查詢狀態按下Enter鍵
            If Me.txtEH(0).Text = "" Then
               MsgBox "工作日期不可空白，請重新輸入 !", vbCritical
               If Me.txtEH(0).Enabled = True Then Me.txtEH(0).SetFocus
               txtEH_GotFocus 0
               Exit Sub
            End If
            intI = 1
            'Modify By Sindy 2014/7/31
            'strExc(0) = "SELECT COUNT(*) FROM ExtendHour WHERE EH01=" & ChangeTStringToWString(Me.txtEH(0).Text) & " And EH02='" & Me.txtEH(1).Text & "' And EH04= '" & Me.txtEH(3).Text & "'"
            strExc(0) = "SELECT COUNT(*) FROM ExtendHour WHERE EH01=" & ChangeTStringToWString(Me.txtEH(0).Text) & " And EH02='" & Trim(Left(Combo1.Text, 6)) & "' And EH04= '" & Me.txtEH(3).Text & "'"
            '2014/7/31 END
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) = 0 Then
                  MsgBox "查無此衍生工作記錄 !", vbCritical
                    strExc(1) = TmpEH(1)
                    strExc(2) = TmpEH(2)
                    strExc(4) = TmpEH(4)
               Else
                    strExc(1) = ChangeTStringToWString(Me.txtEH(0).Text)
                    'strExc(2) = Me.txtEH(1).Text
                    strExc(2) = Trim(Left(Combo1.Text, 6))
                    strExc(4) = Me.txtEH(3).Text
               End If
            End If
            ReadExtendHour strExc
         End If
         
         '發信
         PUB_SendMailCache
         CmdSitu True
         
      Case 4 'cancel
         If ActionEdit <> 2 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
         End If
         CmdSitu True
        If TmpEH(1) = "" Then TmpEH(1) = strRsStart1
        If TmpEH(2) = "" Then TmpEH(2) = strRsStart2
        If TmpEH(4) = "" Then TmpEH(4) = strRsStart4
        strExc(1) = TmpEH(1)
        strExc(2) = TmpEH(2)
        strExc(4) = TmpEH(4)
         ActionEdit = 3
         ReadExtendHour strExc
         TxtLock 3
      Case 5 'query
        TmpEH(1) = ChangeTStringToWString(Me.txtEH(0).Text)
        'TmpEH(2) = Me.txtEH(1).Text
        TmpEH(2) = Trim(Left(Combo1.Text, 6))
        TmpEH(4) = Me.txtEH(3).Text
         CmdSitu False
         TxtLock 2
         ActionEdit = 2
         If Me.txtEH(0).Enabled = True Then Me.txtEH(0).SetFocus
         txtEH_GotFocus 0
   End Select
   
   Exit Sub
   
flgRollback:
   cnnConnection.RollbackTrans
   
CheckingErr:
   MsgBox Err.Description
End Sub

Private Sub RsAction(ByVal Sty As Integer)
 Dim i As Integer
 
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case Sty
      Case 0 '第一筆
         strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour WHERE EH01=" & strRsStart1 & " And EH02 ='" & strRsStart2 & "' And EH04= '" & strRsStart4 & "' "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "" & RsTemp.Fields("EH01").Value
            strExc(2) = "" & RsTemp.Fields("EH02").Value
            strExc(4) = "" & RsTemp.Fields("EH04").Value
        Else
            strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour WHERE EH01||EH02||EH04>='" & strRsStart1 & strRsStart2 & strRsStart4 & "' Order By EH01, EH02, EH04 "
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                strExc(1) = "" & RsTemp.Fields("EH01").Value
                strExc(2) = "" & RsTemp.Fields("EH02").Value
                strExc(4) = "" & RsTemp.Fields("EH04").Value
                strRsStart1 = strExc(1)
                strRsStart2 = strExc(2)
                strRsStart4 = strExc(4)
            End If
         End If
      Case 1 '前一筆
         'Modify By Sindy 2014/7/31
         'If ChangeTStringToWString(Me.txtEH(0).Text) & Me.txtEH(1).Text & Me.txtEH(3).Text = strRsStart1 & strRsStart2 & strRsStart4 Then
         If ChangeTStringToWString(Me.txtEH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtEH(3).Text = strRsStart1 & strRsStart2 & strRsStart4 Then
         '2014/7/31 END
            Beep
            Screen.MousePointer = vbDefault
            DataErrorMessage 6
            Exit Sub
         Else
            intI = 1
            'Modify By Sindy 2014/7/31
            'strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour WHERE EH01||EH02||EH04<'" & ChangeTStringToWString(Me.txtEH(0).Text) & Me.txtEH(1).Text & Me.txtEH(3).Text & "' Order By EH01 Desc, EH02 Desc, EH04 Desc "
            strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour WHERE EH01||EH02||EH04<'" & ChangeTStringToWString(Me.txtEH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtEH(3).Text & "' Order By EH01 Desc, EH02 Desc, EH04 Desc "
            '2014/7/31 END
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               strExc(1) = "" & RsTemp.Fields("EH01").Value
               strExc(2) = "" & RsTemp.Fields("EH02").Value
               strExc(4) = "" & RsTemp.Fields("EH04").Value
            End If
         End If
      Case 2 '後一筆
         'Modify By Sindy 2014/7/31
         'If ChangeTStringToWString(Me.txtEH(0).Text) & Me.txtEH(1).Text & Me.txtEH(3).Text = strRsEnd1 & strRsEnd2 & strRsEnd4 Then
         If ChangeTStringToWString(Me.txtEH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtEH(3).Text = strRsEnd1 & strRsEnd2 & strRsEnd4 Then
         '2014/7/31 END
            Beep
            Screen.MousePointer = vbDefault
            DataErrorMessage 7
            Exit Sub
         Else
            'Modify By Sindy 2014/7/31
            'strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour WHERE EH01||EH02||EH04>'" & ChangeTStringToWString(Me.txtEH(0).Text) & Me.txtEH(1).Text & Me.txtEH(3).Text & "' Order By EH01, EH02, EH04 "
            strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour WHERE EH01||EH02||EH04>'" & ChangeTStringToWString(Me.txtEH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtEH(3).Text & "' Order By EH01, EH02, EH04 "
            '2014/7/31 END
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               strExc(1) = "" & RsTemp.Fields("EH01").Value
               strExc(2) = "" & RsTemp.Fields("EH02").Value
               strExc(4) = "" & RsTemp.Fields("EH04").Value
            End If
         End If
      Case 3 '最後筆
         strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour WHERE EH01=" & strRsEnd1 & " And EH02='" & strRsEnd2 & "' And EH04='" & strRsEnd4 & "'"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "" & RsTemp.Fields("EH01").Value
            strExc(2) = "" & RsTemp.Fields("EH02").Value
            strExc(4) = "" & RsTemp.Fields("EH04").Value
        Else
            strExc(0) = "SELECT EH01, EH02, EH04 FROM ExtendHour WHERE EH01||EH02||EH04<='" & strRsEnd1 & strRsEnd2 & strRsEnd4 & "' Order By EH01 Desc, EH02 Desc, EH04 Desc "
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                strExc(1) = "" & RsTemp.Fields("EH01").Value
                strExc(2) = "" & RsTemp.Fields("EH02").Value
                strExc(4) = "" & RsTemp.Fields("EH04").Value
                strRsEnd1 = strExc(1)
                strRsEnd2 = strExc(2)
                strRsEnd4 = strExc(4)
            End If
         End If
   End Select
   ReadExtendHour strExc
   Screen.MousePointer = vbDefault
   Exit Sub
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

Private Sub CmdSitu(ByVal TF As Boolean)
 Dim i As Integer, txt As TextBox
   If TF = True Then
'      TxtLock 0
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = True
         If Not IsEmptyText(strRsStart1) And Not IsEmptyText(strRsEnd1) Then
            TBar1.Buttons(i + 5).Enabled = True
         Else
            TBar1.Buttons(i + 5).Enabled = False
         End If
      Next
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
   Else
'      TxtLock 1
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
   End If
End Sub

Private Sub TxtLock(ByVal Lt As Integer)

Select Case Lt
Case 0 '新增
    Me.txtEH(0).Locked = False
    'Me.txtEH(1).Locked = False
    Combo1.Locked = False
    Me.txtEH(4).Locked = False
    Me.txtEH(5).Locked = False
    Me.txtEH(6).Locked = False
    Me.txtEH(7).Locked = False
    Me.txtEH(8).Locked = False
    Me.txtEH(9).Locked = False
    Me.txtEH(14).Locked = False
    Me.txtEH(0).Text = ""
    'Me.txtEH(1).Text = ""
    Me.txtEH(3).Text = ""
    Me.txtEH(4).Text = ""
    Me.txtEH(5).Text = ""
    Me.txtEH(6).Text = ""
    Me.txtEH(7).Text = ""
    Me.txtEH(8).Text = ""
    Me.txtEH(9).Text = ""
    Me.txtEH(14).Text = ""
    Me.lblCaseCnt.Caption = ""
    Me.lblSANo.Caption = ""
    Me.lblSAName.Caption = ""
    Combo1.Text = ""
    'Me.lblSupName.Caption = ""
    If ProState = "2" Then
        Me.Check1.Enabled = True
        Me.Check1.Value = vbUnchecked
        Check2.Value = vbUnchecked
        Check2.Enabled = True
    Else
        Me.Check1.Enabled = False
        Me.Check1.Value = vbUnchecked
        Check2.Value = vbUnchecked
        Check2.Enabled = False
    End If
    cmdSelCp09.Enabled = True
    
Case 1 '修改
    Me.txtEH(0).Locked = False
    'Me.txtEH(1).Locked = True
    Combo1.Locked = True
    Me.txtEH(4).Locked = False
    Me.txtEH(5).Locked = False
    Me.txtEH(6).Locked = False
    Me.txtEH(7).Locked = False
    Me.txtEH(8).Locked = False
    Me.txtEH(9).Locked = False
    Me.txtEH(14).Locked = False
    If ProState = "2" Then
         Me.Check1.Enabled = True
         cmdSelCp09.Enabled = True
         Check2.Enabled = True
    Else
         Me.Check1.Enabled = False
         cmdSelCp09.Enabled = False
         Check2.Enabled = False
    End If

Case 2 '查詢
    Me.txtEH(0).Locked = False
    'Me.txtEH(1).Locked = False
    Combo1.Locked = False
    Me.txtEH(4).Locked = True
    Me.txtEH(5).Locked = True
    Me.txtEH(6).Locked = True
    Me.txtEH(7).Locked = True
    Me.txtEH(8).Locked = True
    Me.txtEH(9).Locked = True
    Me.txtEH(14).Locked = True
    Me.txtEH(0).Text = ""
    'Me.txtEH(1).Text = ""
    Me.txtEH(3).Text = ""
    Me.txtEH(4).Text = ""
    Me.txtEH(5).Text = ""
    Me.txtEH(6).Text = ""
    Me.txtEH(7).Text = ""
    Me.txtEH(8).Text = ""
    Me.txtEH(9).Text = ""
    Me.txtEH(14).Text = ""
    Me.lblCaseCnt.Caption = ""
    Me.lblSANo.Caption = ""
    Me.lblSAName.Caption = ""
    Combo1.Text = ""
    'Me.lblSupName.Caption = ""
    Me.Check1.Enabled = False
    Me.Check1.Value = vbUnchecked
    cmdSelCp09.Enabled = False
    Check2.Enabled = False
    Check2.Value = vbUnchecked
    
Case 3 '按下取消後的狀態
    Me.txtEH(0).Locked = True
    'Me.txtEH(1).Locked = True
    Combo1.Locked = True
    Me.txtEH(4).Locked = True
    Me.txtEH(5).Locked = True
    Me.txtEH(6).Locked = True
    Me.txtEH(7).Locked = True
    Me.txtEH(8).Locked = True
    Me.txtEH(9).Locked = True
    Me.txtEH(14).Locked = True
    Me.Check1.Enabled = False
    cmdSelCp09.Enabled = False
    Check2.Enabled = False
End Select
End Sub

Private Sub grdList_Click()
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

Private Sub grdList_DblClick()
   SSTab1.Tab = 0
End Sub

Private Sub grdList_SelChange()
   Dim nRow As Integer
    grdList_ShowSelection

    If grdList.row > 0 And grdList.row <= grdList.Rows - 1 Then
        nRow = grdList.row
        strExc(1) = DBDATE(Me.grdList.TextMatrix(nRow, 1))
        strExc(2) = Me.grdList.TextMatrix(nRow, 2)
        strExc(4) = Me.grdList.TextMatrix(nRow, 6)
        ReadExtendHour strExc
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
On Error Resume Next
    Select Case Me.SSTab1.Tab
    Case 0
        Me.txtEH(0).SetFocus
        txtEH_GotFocus 0
        Me.cmdQuery(0).Default = False
    Case 1
        Me.txtEH(10).SetFocus
        txtEH_GotFocus 12
        Me.cmdQuery(0).Default = True
    End Select
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHand
   SSTab1.Tab = 0 'Add by Morgan 2011/10/19
   Select Case Button.Index
      Case 1 '按下新增
         SSTab1.TabEnabled(1) = False 'Add by Morgan 2011/10/19
         RsSitu 0
      Case 2 '按下修改
         SSTab1.TabEnabled(1) = False 'Add by Morgan 2011/10/19
         RsSitu 1
      Case 3 '按下刪除
         RsSitu 2
      Case 4 '按下查詢
         SSTab1.TabEnabled(1) = False 'Add by Morgan 2011/10/19
         RsSitu 5
      Case 6 '第一筆
         RsAction 0
      Case 7 '前一筆
         RsAction 1
      Case 8 '後一筆
         RsAction 2
      Case 9 '最後筆
         RsAction 3
      Case 11 '按下確定
         RsSitu 3
      Case 12 '按下取消
         RsSitu 4
      Case 14 '結束
         Unload Me
         Exit Sub
   End Select

   If ActionEdit = 3 Then
      SSTab1.TabEnabled(1) = True 'Add by Morgan 2011/10/19
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
         'Added by Lydia 2022/05/18
         If m_bQuery Then
             TBar1.Buttons(4).Enabled = True
         Else
             TBar1.Buttons(4).Enabled = False
         End If
         'end 2022/05/18
      End If
   End If
   Exit Sub
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

Private Function CheckRule() As Boolean
   Dim i As Integer, bolChk As Boolean, j As Integer
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
      
   CheckRule = False
   If Me.txtEH(0).Text = "" Then
      MsgBox "工作日期不可空白 !", vbCritical
      Me.txtEH(0).SetFocus
      txtEH_GotFocus 0
      Exit Function
   End If
   'Modify By Sindy 2014/7/31
   'If Me.txtEH(1).Text = "" Then
   If Trim(Combo1.Text) = "" Then
   '2014/7/31 END
      MsgBox "工作人員不可空白 !", vbCritical
      Combo1.SetFocus
      Exit Function
   End If

   If Me.txtEH(4).Text = "" Then
      MsgBox "工作數時不可空白 !", vbCritical
      Me.txtEH(4).SetFocus
      txtEH_GotFocus 4
      Exit Function
   End If
    If Me.txtEH(5).Text <> "" And Me.txtEH(6).Text <> "" Then
        '案號補滿
        If Me.txtEH(7).Text = "" Then Me.txtEH(7).Text = "0"
        If Me.txtEH(8).Text = "" Then Me.txtEH(8).Text = "00"
        StrSQLa = "Select PA01 From Patent Where PA01='" & Me.txtEH(5).Text & "' And PA02='" & Me.txtEH(6).Text & "' And PA03='" & Me.txtEH(7).Text & "' And PA04='" & Me.txtEH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select TM01 From Trademark Where TM01='" & Me.txtEH(5).Text & "' And TM02='" & Me.txtEH(6).Text & "' And TM03='" & Me.txtEH(7).Text & "' And TM04='" & Me.txtEH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select LC01 From Lawcase Where LC01='" & Me.txtEH(5).Text & "' And LC02='" & Me.txtEH(6).Text & "' And LC03='" & Me.txtEH(7).Text & "' And LC04='" & Me.txtEH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select HC01 From Hirecase Where HC01='" & Me.txtEH(5).Text & "' And HC02='" & Me.txtEH(6).Text & "' And HC03='" & Me.txtEH(7).Text & "' And HC04='" & Me.txtEH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select SP01 From Servicepractice Where SP01='" & Me.txtEH(5).Text & "' And SP02='" & Me.txtEH(6).Text & "' And SP03='" & Me.txtEH(7).Text & "' And SP04='" & Me.txtEH(8).Text & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount <= 0 Then
            MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
            Me.txtEH(5).SetFocus
            txtEH_GotFocus 5
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Function
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
    If Me.txtEH(14).Text <> "" Then
        StrSQLa = "Select * From Caseprogress Where CP09='" & Me.txtEH(14).Text & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic
        If rsA.RecordCount <= 0 Then
            MsgBox "無此收文號資料!!!", vbExclamation + vbOKOnly
            Me.txtEH(14).SetFocus
            txtEH_GotFocus 14
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Function
        Else
            If Me.txtEH(5).Text <> "" Or Me.txtEH(6).Text <> "" Then
                If Me.txtEH(5).Text <> "" & rsA.Fields(0).Value Or Me.txtEH(6).Text <> "" & rsA.Fields(1).Value Or Me.txtEH(7).Text <> "" & rsA.Fields(2).Value Or Me.txtEH(8).Text <> "" & rsA.Fields(3).Value Then
                    MsgBox "此收文號對應的本所案號錯誤!!!", vbExclamation + vbOKOnly
                    Me.txtEH(14).SetFocus
                    txtEH_GotFocus 14
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                    Exit Function
                End If
            End If
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    
    End If
    'End
   CheckRule = True
End Function

Private Function GetData() As Boolean
   Dim i As Integer
   GetData = False
   If CheckRule = False Then Exit Function
   EH(1) = ChangeTStringToWString(Me.txtEH(0).Text)
   'EH(2) = Me.txtEH(1).Text
   EH(2) = Trim(Left(Combo1.Text, 6))
   EH(5) = Me.txtEH(4).Text
   EH(6) = Me.txtEH(5).Text
   EH(7) = Me.txtEH(6).Text
   EH(8) = Me.txtEH(7).Text
   EH(9) = Me.txtEH(8).Text
   EH(10) = Me.txtEH(9).Text
   EH(11) = IIf(Me.Check1.Value = vbChecked, "V", "")
   EH(12) = Me.txtEH(14).Text
   GetData = True
End Function

Private Function TxtValidate() As Boolean
   Dim objTxt As Object
   Dim ii As Integer
   Dim Cancel As Boolean
   
   
   TxtValidate = False
   For Each objTxt In Me.txtEH
       If objTxt.Enabled = True Then
          Cancel = False
          txtEH_Validate objTxt.Index, Cancel
          If Cancel = True Then
             Exit Function
          End If
       End If
   Next
   If Me.txtEH(0).Text = "" Then MsgBox "工作日期不可空白！", vbCritical, "嚴重錯誤！": txtEH(0).SetFocus: Exit Function
   'Modify By Sindy 2014/7/31
   'If Me.txtEH(1).Text = "" Then MsgBox "工作人員不可空白！", vbCritical, "嚴重錯誤！": txtEH(1).SetFocus: Exit Function
   If Trim(Combo1.Text) = "" Then MsgBox "工作人員不可空白！", vbCritical, "嚴重錯誤！": Combo1.SetFocus: Exit Function
   '2014/7/31 END
   bolMail2Boss = False
   bolMail2Boss2 = False
   If ProState = "1" Then
      If (ActionEdit = 0 Or Val(txtEH(4)) > Val(txtEH(4).Tag)) Then
         If Val(txtEH(4)) > 2 Then
            frm090633_2.p_Choice = 1
            Set frm090633_2.p_Parent = Me
            frm090633_2.lblAlert = "本次衍生工作時數超過 2 小時"
            frm090633_2.Show vbModal
            If p_iRtn = 0 Then
               Exit Function
            ElseIf p_iRtn = 1 Then
               bolMail2Boss = True
            ElseIf p_iRtn = 2 Then
               MsgBox "因未呈報時數將自動改為 2 小時！"
               txtEH(4) = 2
            End If
         End If
      End If
      
      '有收文號且收文號有改或時數增加時檢查
      If txtEH(14) <> "" And (txtEH(14) <> txtEH(14).Tag Or (ActionEdit = 0 Or Val(txtEH(4)) > Val(txtEH(4).Tag))) Then
         If ChkTotalOver = True Then
            bolMail2Boss2 = True
         End If
      End If
   End If
   
   'Add By Sindy 2014/7/31
   Cancel = False
   Call Combo1_Validate(Cancel)
   If Cancel = True Then Exit Function
   '檢查是否有增修刪權限 P10.專利處主管
   'Modified by Morgan 2022/9/13 改判斷個人權限才檢查(不要限定P10因還會設個人Ex:99050)
   'If Pub_StrUserSt03 <> "P10" And Pub_StrUserSt03 <> "M51" Then
   If m_ProState = "1" Then
   'end 2022/9/13
      Cancel = False
      For ii = 0 To Combo1.ListCount - 1
         If Combo1.List(ii) = Combo1.Text Then
            Cancel = True
            Exit For
         End If
      Next ii
      If Cancel = False Then
         MsgBox "無權限維護該人員資料！", vbExclamation
         Combo1.SetFocus
         Exit Function
      End If
   End If
   '2014/7/31 END
   
    'Added by Lydia 2022/01/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If
    
   TxtValidate = True
End Function

Private Sub txtEH_Change(Index As Integer)
   Select Case Index
   Case 4 '工作時數
      If Me.txtEH(Index).Text <> "" Then
         Me.lblCaseCnt.Caption = Format(Val(Me.txtEH(Index).Text) * 0.25, "0.00")
      Else
         Me.lblCaseCnt.Caption = ""
      End If
      
   Case 14 '收文號
      Me.lblSANo.Caption = ""
      Me.lblSAName.Caption = ""
      If Me.txtEH(Index).Text <> "" Then
         strExc(0) = "select cp13,st02 from caseprogress,staff where cp09='" & txtEH(Index) & "' and st01(+)=cp13"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Me.lblSANo.Caption = "" & RsTemp("cp13")
            Me.lblSAName.Caption = "" & RsTemp("st02")
         End If
      End If
   End Select
End Sub

Private Sub txtEH_GotFocus(Index As Integer)
    TextInverse Me.txtEH(Index)
End Sub

'Modified by Lydia 2022/01/030 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txtEH_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
    'Modified by Lydia 2022/05/18 + 15 (工作人員)
    Case 1, 2, 7, 5, 12, 13, 14, 15 '系統類別, 工作人員部門別, 收文號
        KeyAscii = UpperCase(KeyAscii)
    Case 0
        If KeyAscii = 47 Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub txtEH_LostFocus(Index As Integer)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

    Select Case Index
    Case 8 '本所案號
        If Me.txtEH(5).Text <> "" And Me.txtEH(6).Text <> "" Then
            'Add By Cheng 2003/08/01
            '案號補滿
            If Me.txtEH(7).Text = "" Then Me.txtEH(7).Text = "0"
            If Me.txtEH(8).Text = "" Then Me.txtEH(8).Text = "00"
            StrSQLa = "Select PA01 From Patent Where " & ChgPatent(Me.txtEH(5).Text & Me.txtEH(6).Text & Me.txtEH(7).Text & Me.txtEH(8).Text)
            StrSQLa = StrSQLa & " Union Select TM01 From Trademark Where " & ChgTradeMark(Me.txtEH(5).Text & Me.txtEH(6).Text & Me.txtEH(7).Text & Me.txtEH(8).Text)
            StrSQLa = StrSQLa & " Union Select LC01 From Lawcase Where " & ChgLawcase(Me.txtEH(5).Text & Me.txtEH(6).Text & Me.txtEH(7).Text & Me.txtEH(8).Text)
            StrSQLa = StrSQLa & " Union Select HC01 From Hirecase Where " & ChgHirecase(Me.txtEH(5).Text & Me.txtEH(6).Text & Me.txtEH(7).Text & Me.txtEH(8).Text)
            StrSQLa = StrSQLa & " Union Select SP01 From Servicepractice Where " & ChgService(Me.txtEH(5).Text & Me.txtEH(6).Text & Me.txtEH(7).Text & Me.txtEH(8).Text)
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount <= 0 Then
                MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
                Me.txtEH(5).SetFocus
                txtEH_GotFocus 5
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
    Case 11 '工作日期
        If Me.txtEH(10).Text <> "" And Me.txtEH(11).Text <> "" Then
            If Val(Me.txtEH(10).Text) > Val(Me.txtEH(11).Text) Then
                MsgBox "工作日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                Me.txtEH(10).SetFocus
                txtEH_GotFocus 10
                Exit Sub
            End If
        End If
    Case 12 '工作人員部門
        If Me.txtEH(12).Text <> "" And Me.txtEH(13).Text <> "" Then
            If Me.txtEH(12).Text > Me.txtEH(13).Text Then
                MsgBox "工作人員部門範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                Me.txtEH(12).SetFocus
                txtEH_GotFocus 12
                Exit Sub
            End If
        End If
    'Added by Lydia 2022/05/18
    Case 15 '工作人員
        Label5.Caption = ""
        If Trim(txtEH(Index)) <> "" Then
            StrSQLa = GetStaffName(Trim(txtEH(Index)), True)
            If StrSQLa = "" Then
                MsgBox "工作人員編號輸入錯誤!!!", vbExclamation + vbOKOnly
                Me.txtEH(Index).SetFocus
                txtEH_GotFocus Index
                Exit Sub
            Else
                Label5.Caption = StrSQLa
            End If
        End If
    'end 2022/05/18
    End Select
End Sub

Private Sub txtEH_Validate(Index As Integer, Cancel As Boolean)
   If Me.txtEH(Index).Text = "" Then Exit Sub
   Select Case Index
   Case 0 '工作日期
       If CheckIsTaiwanDate(Me.txtEH(Index).Text) = False Then
           Cancel = True
       End If
'   Case 1 '工作人員
'       Me.lblSupName.Caption = GetStaffName(Me.txtEH(Index).Text)
'       If Me.lblSupName.Caption = "" And Check1.Value = vbUnchecked Then
'           MsgBox "工作人員輸入錯誤!!!", vbExclamation + vbOKOnly
'           Cancel = True
'       End If

   Case 4 '工作時數
      If Val(Me.txtEH(4)) < 0.5 Then
         MsgBox "衍生工作時數未達 [ 0.5 ] 小時!!!", vbExclamation + vbOKOnly
         Cancel = True
      End If
      
   Case 9 '備註
       If CheckLengthIsOK(Me.txtEH(Index).Text, 200) = False Then
          Cancel = True
       End If
   Case 10, 11 '工作日期區間
       If CheckIsTaiwanDate(Me.txtEH(Index).Text) = False Then
           Cancel = True
       End If
   Case 14 '收文號
       Cancel = Not GetOurCaseNo(Me.txtEH(14).Text)
   End Select
   If Cancel = True Then txtEH_GotFocus Index
End Sub

Private Function QueryData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nRow As Integer
   
    QueryData = False
    InitialGridList
    strSql = ""
    If Me.txtEH(10).Text <> "" Then
        strSql = strSql & " And EH01>=" & DBDATE(Me.txtEH(10).Text) & " "
    End If
    If Me.txtEH(11).Text <> "" Then
        strSql = strSql & " And EH01<=" & DBDATE(Me.txtEH(11).Text) & " "
    End If
    If Me.txtEH(12).Text <> "" Then
        strSql = strSql & " And S1.ST03>='" & ChgSQL(Me.txtEH(12).Text) & "' "
    End If
    If Me.txtEH(13).Text <> "" Then
        strSql = strSql & " And S1.ST03<='" & ChgSQL(Me.txtEH(13).Text) & "' "
    End If
    If Check3.Value = vbChecked Then
        strSql = strSql & " and EH19='V' and ((eh06 is null and eh07 is null) or  eh12 is null ) "
    End If
    
    If Trim(txtSystem.Text) <> "" And Trim(txtCode(0).Text) <> "" Then
        strSql = strSql & " and EH06='" & Trim(txtSystem.Text) & "' and EH07='" & Trim(txtCode(0).Text) & "' "
        If Trim(txtCode(1).Text) <> "" Then
            strSql = strSql & " and EH08='" & Trim(txtCode(1).Text) & "' "
        Else
            strSql = strSql & " and EH08='0' "
        End If
        If Trim(txtCode(2).Text) <> "" Then
            strSql = strSql & " and EH09='" & Trim(txtCode(2).Text) & "' "
        Else
            strSql = strSql & " and EH09='00' "
        End If
    End If
    'Added by Lydia 2022/05/18 工作人員
    If Trim(txtEH(15)) <> "" Then
        strSql = strSql & " and EH02='" & Trim(txtEH(15)) & "'"
    End If
    'end 2022/05/18
    
    strSql = "SELECT Decode(EH01, Null, Null, EH01-19110000), EH02, S1.ST02, CP13, S2.ST02, EH04, EH12, Decode(EH06, Null, '', EH06||Decode(EH07, Null, '','-'||EH07||Decode(EH08, Null, '','-'||EH08||Decode(EH09, Null, '','-'||EH09)))), EH05, EH05 * 0.25, EH11, EH10 FROM ExtendHour,caseprogress, Staff S1, Staff S2 " & _
         "WHERE EH02=S1.ST01(+) AND cp09(+)=EH12 and CP13=S2.ST01(+) " & strSql & " Order By 1, 2, 4, 5 "
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        QueryData = True
        UpdateGridList rsTmp
        'Added by Lydia 2022/01/11 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
        If grdList.Rows > 1 Then
           grdList.FixedRows = 1
        End If
    End If
    rsTmp.Close
    Set rsTmp = Nothing
End Function

' 初始化列表
Public Sub InitialGridList()
    grdList.Clear
    grdList.Rows = 1
    grdList.Cols = 13
    grdList.ColWidth(0) = 300
    grdList.row = 0
    grdList.col = 0
    grdList.ColAlignment(0) = flexAlignCenterCenter
    grdList.col = 1
    grdList.Text = "工作日期"
    grdList.ColWidth(1) = 800
    grdList.ColAlignment(1) = flexAlignCenterCenter
    grdList.col = 2
    'grdList.Text = "工作人員代號"
    grdList.ColWidth(2) = 0
    grdList.ColAlignment(2) = flexAlignCenterCenter
    grdList.col = 3
    grdList.Text = "工作人員"
    grdList.ColWidth(3) = 800
    grdList.ColAlignment(3) = flexAlignCenterCenter
    grdList.col = 4
    'grdList.Text = "智權人員代號"
    grdList.ColWidth(4) = 0
    grdList.ColAlignment(4) = flexAlignCenterCenter
    grdList.col = 5
    grdList.Text = "智權人員"
    grdList.ColWidth(5) = 800
    grdList.ColAlignment(5) = flexAlignCenterCenter
    grdList.col = 6
    grdList.Text = "序號"
    grdList.ColWidth(6) = 500
    grdList.ColAlignment(6) = flexAlignCenterCenter
    grdList.col = 7
    grdList.Text = "收文號"
    grdList.ColWidth(7) = 1400
    grdList.ColAlignment(7) = flexAlignCenterCenter
    grdList.col = 8
    grdList.Text = "本所案號"
    grdList.ColWidth(8) = 1400
    grdList.ColAlignment(8) = flexAlignCenterCenter
    grdList.col = 9
    grdList.Text = "工作時數"
    grdList.ColWidth(9) = 800
    grdList.ColAlignment(9) = flexAlignCenterCenter
    grdList.col = 10
    grdList.Text = "折算件數"
    grdList.ColWidth(10) = 800
    grdList.ColAlignment(10) = flexAlignCenterCenter
    grdList.col = 11
    grdList.Text = "主管核可"
    grdList.ColWidth(11) = 800
    grdList.ColAlignment(11) = flexAlignCenterCenter
    grdList.col = 12
    grdList.Text = "備註"
    grdList.ColWidth(12) = 3000
    grdList.ColAlignment(12) = flexAlignCenterCenter
End Sub

Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset)
Dim nRow As Integer
    rsTmp.MoveFirst
    Do While rsTmp.EOF = False
        grdList.Rows = grdList.Rows + 1
        nRow = grdList.Rows - 1
        grdList.TextMatrix(nRow, 1) = "" & rsTmp.Fields(0).Value
        grdList.TextMatrix(nRow, 2) = "" & rsTmp.Fields(1).Value
        Me.grdList.row = nRow
        Me.grdList.col = 3
        Me.grdList.CellAlignment = flexAlignLeftCenter
        grdList.TextMatrix(nRow, 3) = "" & rsTmp.Fields(2).Value
        grdList.TextMatrix(nRow, 4) = "" & rsTmp.Fields(3).Value
        Me.grdList.row = nRow
        Me.grdList.col = 5
        Me.grdList.CellAlignment = flexAlignLeftCenter
        grdList.TextMatrix(nRow, 5) = "" & rsTmp.Fields(4).Value
        grdList.TextMatrix(nRow, 6) = "" & rsTmp.Fields(5).Value
        Me.grdList.row = nRow
        Me.grdList.col = 7
        Me.grdList.CellAlignment = flexAlignLeftCenter
        grdList.TextMatrix(nRow, 7) = "" & rsTmp.Fields(6).Value
        Me.grdList.row = nRow
        Me.grdList.col = 8
        Me.grdList.CellAlignment = flexAlignLeftCenter
        grdList.TextMatrix(nRow, 8) = "" & rsTmp.Fields(7).Value
        Me.grdList.row = nRow
        Me.grdList.col = 9
        Me.grdList.CellAlignment = flexAlignRightCenter
        grdList.TextMatrix(nRow, 9) = "" & rsTmp.Fields(8).Value
        Me.grdList.row = nRow
        Me.grdList.col = 10
        Me.grdList.CellAlignment = flexAlignRightCenter
        grdList.TextMatrix(nRow, 10) = "" & rsTmp.Fields(9).Value
        Me.grdList.row = nRow
        Me.grdList.col = 11
        Me.grdList.CellAlignment = flexAlignCenterCenter
        grdList.TextMatrix(nRow, 11) = "" & rsTmp.Fields(10).Value
        Me.grdList.col = 12
        Me.grdList.CellAlignment = flexAlignLeftCenter
        grdList.TextMatrix(nRow, 12) = "" & rsTmp.Fields(11).Value
        rsTmp.MoveNext
    Loop
End Sub

'取得序號
Private Function GetSerialNo(strEH01 As String, strEH02 As String) As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   'Modified by Morgan 2012/7/13 沒資料改預設 0 否則 win7 會有錯誤(資料提供者或其他服務傳回E_FAIL狀態)
   StrSQLa = "Select nvl(max(EH04),0) as EH04 From ExtendHour Where EH01=" & strEH01 & " And EH02='" & strEH02 & "'  Order By EH04 Desc "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsA.EOF And Not rsA.BOF Then
       GetSerialNo = Format(Val(rsA("EH04").Value) + 1, "000")
   Else
       GetSerialNo = "001"
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

Private Sub PrintData()
Dim ii As Integer
    
    Page = 1
    PrintTitle
    For ii = 1 To Me.grdList.Rows - 1
        '工作日期
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 1)
        '工作人員
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 3)
        '智權人員
        Printer.CurrentX = PLeft(2)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 5)
        '收文號
        Printer.CurrentX = PLeft(3)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 7)
        '本所案號
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 8)
        '工作時數
        Printer.CurrentX = PLeft(5) + Printer.TextWidth("工作時數") - Printer.TextWidth(Format(Me.grdList.TextMatrix(ii, 9), "0.0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(Me.grdList.TextMatrix(ii, 9), "0.0")
        '折算件數
        Printer.CurrentX = PLeft(6) + Printer.TextWidth("折算件數") - Printer.TextWidth(Format(Me.grdList.TextMatrix(ii, 10), "0.00"))
        Printer.CurrentY = iPrint
        Printer.Print Format(Me.grdList.TextMatrix(ii, 10), "0.00")
        '主管核可
        Printer.CurrentX = PLeft(7)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 11)
        '備註
        Printer.CurrentX = PLeft(8)
        Printer.CurrentY = iPrint
        Printer.Print Replace(Me.grdList.TextMatrix(ii, 12), vbCrLf, "")
        iPrint = iPrint + 300
        If iPrint > 10000 And ii <> Me.grdList.Rows - 1 Then
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print String(200, "-")
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
    Next ii
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    Printer.EndDoc
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "衍生工作記錄明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "工作日期：" & Format(ChangeTStringToTDateString(Me.txtEH(10).Text) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Me.txtEH(11).Text)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "工作人員部門別：" & Me.txtEH(12).Text & " " & IIf(Me.txtEH(12).Text <> "" Or Me.txtEH(13).Text <> "", "－", "") & " " & Me.txtEH(13).Text
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "工作日期"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "工作人員"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "收文號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "工作時數"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "折算件數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "核可"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "備　　註"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = PLeft(0) + 1250
PLeft(2) = PLeft(1) + 1250
PLeft(3) = PLeft(2) + 1250
PLeft(4) = PLeft(3) + 1250
PLeft(5) = PLeft(4) + 2300
PLeft(6) = PLeft(5) + 1250
PLeft(7) = PLeft(6) + 1250
PLeft(8) = PLeft(7) + 625
End Sub

Private Function GetOurCaseNo(strCP09) As Boolean
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
   StrSQLa = "Select * From Caseprogress Where CP09='" & strCP09 & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic
   If rsA.RecordCount > 0 Then
      If rsA.Fields("cp01") <> "P" And rsA.Fields("cp01") <> "CFP" Then
         MsgBox "請輸入 P 或 CFP 案的收文號資料!!!", vbExclamation + vbOKOnly, "收文號輸入錯誤"
      ElseIf (Len(rsA.Fields("cp10")) = 3 And ((rsA.Fields("cp10") >= "101" And rsA.Fields("cp10") <= "140") Or (rsA.Fields("cp10") >= "301" And rsA.Fields("cp10") <= "315"))) Then
         Me.txtEH(5).Text = "" & rsA("CP01").Value
         Me.txtEH(6).Text = "" & rsA("CP02").Value
         Me.txtEH(7).Text = "" & rsA("CP03").Value
         Me.txtEH(8).Text = "" & rsA("CP04").Value
         GetOurCaseNo = True
      Else
          MsgBox "請輸入案件性質為 申請[1XX] 或 改請[3XX] 的收文號資料!!!", vbExclamation + vbOKOnly, "收文號輸入錯誤"
      End If
   Else
       MsgBox "無此收文號資料!!!", vbExclamation + vbOKOnly
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

Private Sub txtSystem_GotFocus()
   CloseIme
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   CloseIme
   intTemp = True
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index))
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub subAddMail()
   Dim strMod As String
   Dim strSubject As String, strContent As String
   
   If ActionEdit = 0 Then
      strMod = "新增"
   Else
      strMod = "修改"
   End If
   
   strSubject = strMod & "<<衍生工作記錄>>輸入時數超過 2 小時通知"
   strContent = "工作日期：" & ChangeTStringToTDateString(Me.txtEH(0).Text) & vbCrLf & _
         "工作人員：" & Combo1.Text & vbCrLf & _
         "工作時數：" & Me.txtEH(4).Text & vbCrLf & _
         "折算件數：" & Me.lblCaseCnt.Caption & vbCrLf & _
         "收 文 號：" & Me.txtEH(14).Text & vbCrLf & _
         "本所案號：" & Me.txtEH(5).Text & "-" & Me.txtEH(6).Text & "-" & Me.txtEH(7).Text & "-" & Me.txtEH(8).Text & vbCrLf & _
         "智權人員：" & Me.lblSANo.Caption & " " & Me.lblSAName.Caption & vbCrLf & _
         "備　　註：" & Me.txtEH(9).Text & vbCrLf & _
         "是否主管核可：" & IIf(Me.Check1.Value = vbChecked, "是", "否") & vbCrLf & vbCrLf & _
         strMod & "資料人員：" & strUserNum & " " & GetStaffName(strUserNum)
   
   StrSQLa = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " select '" & strUserNum & "',OMAN,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & strSubject & "','" & ChgSQL(strContent) & "','71011'" & _
            " from staff,SetSpecMan where st01='" & Trim(Left(Combo1.Text, 6)) & "' and OCODE(+)=decode(st06,'1','A2','2','A3','3','A4','4','A5')"
   cnnConnection.Execute StrSQLa, intI
End Sub

Private Sub subAddMail2()
   Dim strMod As String
   Dim strSubject As String, strContent As String
   
   If ActionEdit = 0 Then
      strMod = "新增"
   Else
      strMod = "修改"
   End If
   
   strSubject = strMod & "<<衍生工作記錄>>有異常的衍生工作問題，請瞭解處理。"
   strContent = "工作日期：" & ChangeTStringToTDateString(Me.txtEH(0).Text) & vbCrLf & _
         "工作人員：" & Combo1.Text & vbCrLf & _
         "工作時數：" & Me.txtEH(4).Text & vbCrLf & _
         "折算件數：" & Me.lblCaseCnt.Caption & vbCrLf & _
         "收 文 號：" & Me.txtEH(14).Text & vbCrLf & _
         "本所案號：" & Me.txtEH(5).Text & "-" & Me.txtEH(6).Text & "-" & Me.txtEH(7).Text & "-" & Me.txtEH(8).Text & vbCrLf & _
         "智權人員：" & Me.lblSANo.Caption & " " & Me.lblSAName.Caption & vbCrLf & _
         "備　　註：" & Me.txtEH(9).Text & vbCrLf & _
         "是否主管核可：" & IIf(Me.Check1.Value = vbChecked, "是", "否") & vbCrLf & vbCrLf & _
         strMod & "資料人員：" & strUserNum & " " & GetStaffName(strUserNum)
   
   StrSQLa = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " select '" & strUserNum & "',OMAN,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & strSubject & "','" & ChgSQL(strContent) & "'" & _
            " from staff,SetSpecMan where st01='" & Trim(Left(Combo1.Text, 6)) & "' and OCODE(+)=decode(st06,'1','A2','2','A3','3','A4','4','A5')"
   cnnConnection.Execute StrSQLa, intI
End Sub

'Added by Lydia 2016/08/11 點選欄位進行排序
Private Sub grdList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   'Modified by Lydia 2022/01/11
   'Pub_MSFGrdColRow grdList, x, y, nCol, nRow
   getGrdColRow grdList, x, y, nCol, nRow
   
   If nCol < 0 Or nRow < 0 Then Exit Sub
   
   grdList.col = nCol
   grdList.row = nRow
   If Me.grdList.row < 1 Then
      If InStr("工作日期,序號,工作時數,折算件數", Me.grdList.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.grdList.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.grdList.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.grdList.Sort = 5 '字串昇冪
            
            m_blnColOrderAsc = False
         Else
            Me.grdList.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub
'end 2016/08/11

