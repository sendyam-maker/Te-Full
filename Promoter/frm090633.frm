VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090633 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人修改記錄維護"
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
            Picture         =   "frm090633.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090633.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090633.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090633.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090633.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090633.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090633.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090633.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090633.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090633.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090633.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4605
      Left            =   180
      TabIndex        =   25
      Top             =   720
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   8128
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm090633.frx":20F4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "lblCaseCnt"
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(5)=   "Label1(0)"
      Tab(0).Control(6)=   "Line1"
      Tab(0).Control(7)=   "Label1(155)"
      Tab(0).Control(8)=   "Label1(154)"
      Tab(0).Control(9)=   "Label1(55)"
      Tab(0).Control(10)=   "Label1(6)"
      Tab(0).Control(11)=   "Label1(7)"
      Tab(0).Control(12)=   "lblSANo"
      Tab(0).Control(13)=   "Label3"
      Tab(0).Control(14)=   "Label4"
      Tab(0).Control(15)=   "lblSAName"
      Tab(0).Control(16)=   "Combo1"
      Tab(0).Control(17)=   "txtMH(6)"
      Tab(0).Control(18)=   "txtMH(7)"
      Tab(0).Control(19)=   "txtMH(8)"
      Tab(0).Control(20)=   "txtMH(9)"
      Tab(0).Control(21)=   "txtMH(3)"
      Tab(0).Control(22)=   "txtMH(5)"
      Tab(0).Control(23)=   "txtMH(0)"
      Tab(0).Control(24)=   "txtMH(4)"
      Tab(0).Control(25)=   "txtMH(14)"
      Tab(0).Control(26)=   "Check1"
      Tab(0).Control(27)=   "cmdSelCp09"
      Tab(0).Control(28)=   "Check2"
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm090633.frx":2110
      Tab(1).ControlEnabled=   -1  'True
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
      Tab(1).Control(5)=   "txtMH(10)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtMH(11)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtMH(12)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtMH(13)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "grdList"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdQuery(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdQuery(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Check3"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtCode(0)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtCode(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtCode(2)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtSystem"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      Begin VB.TextBox txtSystem 
         Height          =   300
         Left            =   960
         MaxLength       =   3
         TabIndex        =   18
         Top             =   780
         Width           =   732
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   2
         Left            =   3390
         MaxLength       =   2
         TabIndex        =   21
         Top             =   780
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   1
         Left            =   2970
         MaxLength       =   1
         TabIndex        =   20
         Top             =   780
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   0
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   19
         Top             =   780
         Width           =   1212
      End
      Begin VB.CheckBox Check3 
         Caption         =   "僅查管制中案件"
         Height          =   195
         Left            =   4770
         TabIndex        =   17
         Top             =   870
         Width           =   1620
      End
      Begin VB.CheckBox Check2 
         Caption         =   "管制中"
         Height          =   225
         Left            =   -72480
         TabIndex        =   12
         Top             =   3705
         Width           =   1605
      End
      Begin VB.CommandButton cmdSelCp09 
         Caption         =   "選擇收文號"
         Height          =   300
         Left            =   -71700
         TabIndex        =   9
         Top             =   1590
         Width           =   1200
      End
      Begin VB.CheckBox Check1 
         Caption         =   "主管核可"
         Height          =   315
         Left            =   -74460
         TabIndex        =   11
         Top             =   3660
         Width           =   1785
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "列印(&P)"
         Height          =   400
         Index           =   1
         Left            =   7770
         TabIndex        =   23
         Top             =   390
         Width           =   912
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Default         =   -1  'True
         Height          =   400
         Index           =   0
         Left            =   6810
         TabIndex        =   22
         Top             =   390
         Width           =   912
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   3330
         Left            =   180
         TabIndex        =   44
         Top             =   1170
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   5884
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
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   14
         Left            =   -73590
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
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   4
         Left            =   -73590
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
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   0
         Left            =   -73590
         TabIndex        =   0
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
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   5
         Left            =   -73590
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
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   3
         Left            =   -70350
         TabIndex        =   1
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
      Begin MSForms.TextBox txtMH 
         Height          =   855
         Index           =   9
         Left            =   -73590
         TabIndex        =   10
         Top             =   2670
         Width           =   5805
         VariousPropertyBits=   -1466941413
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "10239;1508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   8
         Left            =   -71550
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
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   7
         Left            =   -71970
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
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   6
         Left            =   -72960
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
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   13
         Left            =   5700
         TabIndex        =   16
         Top             =   480
         Width           =   525
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   12
         Left            =   5070
         TabIndex        =   15
         Top             =   480
         Width           =   525
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   11
         Left            =   2010
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
      Begin MSForms.TextBox txtMH 
         Height          =   300
         Index           =   10
         Left            =   960
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
      Begin MSForms.ComboBox Combo1 
         Height          =   300
         Left            =   -73590
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
      Begin MSForms.Label lblSAName 
         Height          =   255
         Left            =   -72810
         TabIndex        =   43
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
      Begin MSForms.Label Label4 
         Height          =   255
         Left            =   -71010
         TabIndex        =   42
         Top             =   4050
         Width           =   3300
         VariousPropertyBits=   27
         Caption         =   "lblFM2-Update :"
         Size            =   "5821;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Left            =   -74490
         TabIndex        =   41
         Top             =   4050
         Width           =   3300
         VariousPropertyBits=   27
         Caption         =   "lblFM2-Create :"
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
         Left            =   -73560
         TabIndex        =   40
         Top             =   2340
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   8
         Left            =   150
         TabIndex        =   39
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收文號:"
         Height          =   180
         Index           =   7
         Left            =   -74490
         TabIndex        =   38
         Top             =   1656
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "序號:"
         Height          =   180
         Index           =   6
         Left            =   -71190
         TabIndex        =   37
         Top             =   630
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "修改日期:"
         Height          =   180
         Index           =   55
         Left            =   -74490
         TabIndex        =   34
         Top             =   630
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   154
         Left            =   -74490
         TabIndex        =   33
         Top             =   1998
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "修改人員:"
         Height          =   180
         Index           =   155
         Left            =   -74490
         TabIndex        =   32
         Top             =   972
         Width           =   765
      End
      Begin VB.Line Line1 
         X1              =   -73410
         X2              =   -71280
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員:"
         Height          =   180
         Index           =   0
         Left            =   -74490
         TabIndex        =   31
         Top             =   2340
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "修改時數:"
         Height          =   180
         Index           =   1
         Left            =   -74490
         TabIndex        =   30
         Top             =   1314
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註:"
         Height          =   180
         Index           =   2
         Left            =   -74490
         TabIndex        =   29
         Top             =   2700
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "折算基數:"
         Height          =   180
         Index           =   3
         Left            =   -71190
         TabIndex        =   28
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label lblCaseCnt 
         AutoSize        =   -1  'True
         Caption         =   "lblCaseCnt"
         Height          =   180
         Left            =   -70350
         TabIndex        =   27
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(EX : 123.5) "
         Height          =   180
         Left            =   -72570
         TabIndex        =   26
         Top             =   1290
         Width           =   930
      End
      Begin VB.Line Line3 
         X1              =   5370
         X2              =   5940
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "修改人員部門:"
         Height          =   180
         Index           =   5
         Left            =   3900
         TabIndex        =   36
         Top             =   510
         Width           =   1125
      End
      Begin VB.Line Line2 
         X1              =   1740
         X2              =   2160
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "修改日期:"
         Height          =   180
         Index           =   4
         Left            =   150
         TabIndex        =   35
         Top             =   480
         Width           =   765
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   24
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
Attribute VB_Name = "frm090633"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ; grdList改字型=新細明體-ExtB(MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉)、Combo1、Label3、Label4、lblSAName、txtMH(index)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Create by Morgan 2011/7/22 參考 frm090623
Option Explicit

Dim MH(1 To 19) As String
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
Dim bolMail2Sales As Boolean
Dim iUnlockChoice As Integer
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
        '個人進入時, 只可修改個人輸入的資料, 且不可改修改人員
        'Modify By Sindy 2014/7/31
        'If Me.Check1.Value = vbChecked Or (txtMH(0) <> "" And txtMH(1) <> strUserNum) Then
        If Me.Check1.Value = vbChecked Or (txtMH(0) <> "" And Trim(Left(Combo1.Text, 6)) <> strUserNum) Then
        '2014/7/31 END
            Me.txtMH(0).Enabled = False
            'Me.txtMH(1).Enabled = False
            Combo1.Enabled = False
            Me.txtMH(3).Enabled = False
            Me.txtMH(4).Enabled = False
            Me.txtMH(5).Enabled = False
            Me.txtMH(6).Enabled = False
            Me.txtMH(7).Enabled = False
            Me.txtMH(8).Enabled = False
            Me.txtMH(9).Enabled = False
            Me.txtMH(14).Enabled = False
            cmdSelCp09.Enabled = False
        Else
            Me.txtMH(0).Enabled = True
            'Me.txtMH(1).Enabled = True
            Combo1.Enabled = True
            Me.txtMH(3).Enabled = False
            Me.txtMH(4).Enabled = True
            Me.txtMH(5).Enabled = True
            Me.txtMH(6).Enabled = True
            Me.txtMH(7).Enabled = True
            Me.txtMH(8).Enabled = True
            Me.txtMH(9).Enabled = True
            Me.txtMH(14).Enabled = True
            cmdSelCp09.Enabled = True
        End If
    End If
    If ActionEdit = 1 Then
      If Check1.Value = vbChecked And Check1.Value <> Check1.Tag And Check2.Value = 1 Then
         MsgBox "本收文號目前已在管制中！", vbInformation, "提醒"
      End If
    End If
End Sub
'是否管制中
Private Function ChkControl() As Boolean
   'Modified by Morgan 2011/12/12 + B類收文控管,智權人員會收A類
   strSql = "select cp09 from caseprogress where cp43='" & txtMH(14) & "' and cp09>'B' and cp10='944' and cp27||cp57 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ChkControl = True
   End If
End Function
'是否修改時數累計超過 4 小時
Private Function ChkTotalOver() As Boolean
   Dim strTot As String
   strSql = "select sum(mh05) from modifyhour where mh12='" & txtMH(14) & "' and not (mh01='" & txtMH(0) & "'" & _
   " and mh02='" & Trim(Left(Combo1.Text, 6)) & "' and mh04='" & txtMH(3) & "')"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strTot = "" & RsTemp.Fields(0)
      If Val(strTot) > 0 Then
         strTot = Val(strTot) + Val(txtMH(4))
         If Val(strTot) > 4 Then
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
    
    If Me.txtMH(10).Text = "" And txtSystem.Text = "" And txtCode(0).Text = "" Then
        MsgBox "請輸入修改日期範圍或本所案號條件!!!", vbExclamation + vbOKOnly
        Me.txtMH(10).SetFocus
        Exit Sub
    End If
    If Me.txtMH(10).Text <> "" Then
       If CheckIsTaiwanDate(Me.txtMH(10).Text) = False Then
          Me.txtMH(10).SetFocus
          txtMH_GotFocus 10
          Exit Sub
       End If
    End If
    If Me.txtMH(10).Text <> "" And Me.txtMH(11).Text = "" Then
        MsgBox "請輸入修改迄日!!!", vbExclamation + vbOKOnly
        Me.txtMH(11).SetFocus
        Exit Sub
    End If
    If Me.txtMH(11).Text <> "" Then
       If CheckIsTaiwanDate(Me.txtMH(11).Text) = False Then
          Me.txtMH(11).SetFocus
          Exit Sub
       End If
    End If
    If Val(Me.txtMH(10).Text) > Val(Me.txtMH(11).Text) Then
        MsgBox "修改日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
        Me.txtMH(10).SetFocus
        txtMH_GotFocus 10
        Exit Sub
    End If
    If Me.txtMH(12).Text <> "" And Me.txtMH(13).Text <> "" Then
        If Me.txtMH(12).Text > Me.txtMH(13).Text Then
            MsgBox "修改人員部門範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.txtMH(12).SetFocus
            txtMH_GotFocus 12
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
   If Trim(txtMH(5)) <> "" And Trim(txtMH(6)) <> "" Then
      Load frm090633_1
      frm090633_1.Hide
      frm090633_1.oCP01 = txtMH(5).Text
      frm090633_1.oCP02 = txtMH(6).Text
      frm090633_1.oCP03 = IIf(Trim(txtMH(7).Text) = "", "0", txtMH(7).Text)
      frm090633_1.oCP04 = IIf(Trim(txtMH(8).Text) = "", "00", txtMH(8).Text)
      If frm090633_1.Process = False Then ShowNoData: Exit Sub
      frm090633_1.Show vbModal
      Unload frm090633_1
   Else
      MsgBox "請先輸入本所案號！", vbExclamation, "警告！"
      If Me.txtMH(5).Enabled = True Then Me.txtMH(5).SetFocus
   End If
End Sub

'Add By Sindy 2014/7/31
'Modified by Lydia 2022/01/03 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
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
         MsgBox "修改人員輸入錯誤!!!", vbExclamation + vbOKOnly
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
        m_bInsert = IsUserHasRightOfFunction("frm090633P", strAdd, False)
        m_bUpdate = IsUserHasRightOfFunction("frm090633P", strEdit, False)
        m_bDelete = IsUserHasRightOfFunction("frm090633P", strDel, False)
        m_bQuery = IsUserHasRightOfFunction("frm090633P", strFind, False)
        cmdQuery(1).Enabled = IsUserHasRightOfFunction("frm090633P", strPrint, False)
        Check1.Enabled = False
        Check2.Enabled = False
        
    '由管理進入
    Else
        m_bInsert = IsUserHasRightOfFunction("frm090633M", strAdd, False)
        m_bUpdate = IsUserHasRightOfFunction("frm090633M", strEdit, False)
        m_bDelete = IsUserHasRightOfFunction("frm090633M", strDel, False)
        m_bQuery = IsUserHasRightOfFunction("frm090633M", strFind, False)
        cmdQuery(1).Enabled = IsUserHasRightOfFunction("frm090633M", strPrint, False)
        Check1.Enabled = True
        Check2.Enabled = True
        
    End If
    Call SetCombo1 'Add By Sindy 2014/7/31
    If Val(strSrvDate(1)) >= 20140401 Then m_bInsert = False 'Added by Morgan 2014/3/19 4/1起取消新增功能
    
    strExc(0) = "SELECT * FROM ModifyHour WHERE ROWNUM<1"
    intI = 1
    Set rsDefineSize = ClsLawReadRstMsg(intI, strExc(0))
   strRsStart1 = Empty: strRsStart2 = Empty: strRsStart4 = Empty
   strRsEnd1 = Empty: strRsEnd2 = Empty: strRsEnd4 = Empty
   strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour Order By MH01, MH02, MH04 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
   If intI = 1 Then
        RsTemp.MoveFirst
      strRsStart1 = "" & RsTemp.Fields("MH01").Value
      strRsStart2 = "" & RsTemp.Fields("MH02").Value
      strRsStart4 = "" & RsTemp.Fields("MH04").Value
        RsTemp.MoveLast
      strRsEnd1 = "" & RsTemp.Fields("MH01").Value
      strRsEnd2 = "" & RsTemp.Fields("MH02").Value
      strRsEnd4 = "" & RsTemp.Fields("MH04").Value
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

Private Function ReadModifyHour(ByRef tsTmp() As String) As Boolean
   Dim i As Integer, j As Integer, Lbl As Label, txt As TextBox, strTmp As String
   Dim strTxt(0 To 4) As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
   strTxt(1) = tsTmp(1): strTxt(2) = tsTmp(2): strTxt(4) = tsTmp(4)
   MH(1) = strTxt(1): MH(2) = strTxt(2):  MH(4) = strTxt(4)
   For i = 0 To 10
       If i = 10 Then
           Me.Check1.Value = vbUnchecked
       ElseIf i <> 1 And i <> 2 Then 'Modify By Sindy 2014/7/31 +i <> 1 And
           Me.txtMH(i).Text = ""
       End If
   Next i
   Me.txtMH(14).Text = ""
   Me.lblCaseCnt.Caption = ""
   Me.lblSANo.Caption = ""
   Me.lblSAName.Caption = ""
'   Me.lblSupName.Caption = ""
   Me.Label3.Caption = "Create : "
   Me.Label4.Caption = "Update : "
   Check2.Value = vbUnchecked
   If MH(1) = "" Then Exit Function
   StrSQLa = "Select * From ModifyHour,caseprogress Where MH01=" & MH(1) & " And MH02='" & MH(2) & "' And MH04='" & MH(4) & "' and cp09(+)=MH12 Order By MH01, MH02, MH04 "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
        MH(1) = "" & rsA.Fields("MH01").Value
        MH(2) = "" & rsA.Fields("MH02").Value
        MH(4) = "" & rsA.Fields("MH04").Value
        MH(5) = "" & rsA.Fields("MH05").Value
        MH(6) = "" & rsA.Fields("MH06").Value
        MH(7) = "" & rsA.Fields("MH07").Value
        MH(8) = "" & rsA.Fields("MH08").Value
        MH(9) = "" & rsA.Fields("MH09").Value
        MH(10) = "" & rsA.Fields("MH10").Value
        MH(11) = "" & rsA.Fields("MH11").Value
        MH(12) = "" & rsA.Fields("MH12").Value
        MH(13) = "" & rsA.Fields("MH13").Value
        MH(14) = "" & rsA.Fields("MH14").Value
        MH(15) = "" & rsA.Fields("MH15").Value
        MH(16) = "" & rsA.Fields("MH16").Value
        MH(17) = "" & rsA.Fields("MH17").Value
        MH(18) = "" & rsA.Fields("MH18").Value
        lblSANo = "" & rsA.Fields("CP13").Value
    Else
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Function
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   Me.txtMH(0).Text = ChangeWStringToTString(MH(1))
   Combo1.Text = MH(2) & " " & GetStaffName(MH(2), True) 'Add By Sindy 2014/7/31
   'Me.txtMH(1).Text = MH(2)
   'Me.lblSupName.Caption = GetStaffName(MH(2), True)
   Me.lblSAName.Caption = GetStaffName(lblSANo, True)
   Me.txtMH(3).Text = MH(4)
   Me.txtMH(4).Text = MH(5)
   Me.txtMH(5).Text = MH(6)
   Me.txtMH(6).Text = MH(7)
   Me.txtMH(7).Text = MH(8)
   Me.txtMH(8).Text = MH(9)
   Me.txtMH(9).Text = MH(10)
   Me.lblCaseCnt.Caption = Format(Val(Me.txtMH(4).Text) * 0.2, "0.00")
   Me.Check1.Value = IIf(MH(11) <> "", vbChecked, vbUnchecked)
   Me.txtMH(14).Text = MH(12)
   Call Check1_Click
   
   Me.txtMH(0).Tag = Me.txtMH(0).Text
   'Me.txtMH(1).Tag = Me.txtMH(1).Text
   Combo1.Tag = Combo1.Text
   Me.txtMH(3).Tag = Me.txtMH(3).Text
   Me.txtMH(4).Tag = Me.txtMH(4).Text
   Me.txtMH(5).Tag = Me.txtMH(5).Text
   Me.txtMH(6).Tag = Me.txtMH(6).Text
   Me.txtMH(7).Tag = Me.txtMH(7).Text
   Me.txtMH(8).Tag = Me.txtMH(8).Text
   Me.txtMH(9).Tag = Me.txtMH(9).Text
   Me.Check1.Tag = Me.Check1.Value
   If ChkControl = True Then
      Check2.Value = 1
   End If
   Check2.Tag = Check2.Value
   Me.txtMH(14).Tag = Me.txtMH(14).Text
   If MH(13) <> "" Then
       Me.Label3.Caption = Me.Label3.Caption & GetStaffName(MH(13))
   End If
   If MH(14) <> "" Then
       Me.Label3.Caption = Me.Label3.Caption & " " & ChangeTStringToTDateString(Val(MH(14)) - 19110000)
   End If
   If MH(15) <> "" Then
       Me.Label3.Caption = Me.Label3.Caption & " " & Format(MH(15), "##:##")
   End If
   If MH(16) <> "" Then
       Me.Label4.Caption = Me.Label4.Caption & GetStaffName(MH(16))
   End If
   If MH(17) <> "" Then
       Me.Label4.Caption = Me.Label4.Caption & " " & ChangeTStringToTDateString(Val(MH(17)) - 19110000)
   End If
   If MH(18) <> "" Then
       Me.Label4.Caption = Me.Label4.Caption & " " & Format(MH(18), "##:##")
   End If
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm090633 = Nothing
End Sub

Private Sub RsSitu(ByVal Situ As Integer)
   Dim i As Integer, St1 As String, St2 As String
   Dim TBmk As Variant
   Dim StrSQLa As String
   Dim MH04 As String
 
 On Error GoTo CheckingErr
 
 Static TmpMH(4) As String
   Select Case Situ
      Case 0 '按下新增add
        TmpMH(1) = ChangeTStringToWString(Me.txtMH(0).Text)
        'TmpMH(2) = Me.txtMH(1).Text
        TmpMH(2) = Trim(Left(Combo1.Text, 6))
        TmpMH(4) = Me.txtMH(3).Text
        Me.lblCaseCnt.Caption = ""
        Me.lblSANo.Caption = ""
        Me.lblSAName.Caption = ""
        Me.lblCaseCnt.Caption = ""
        Me.Label3.Caption = "Create : "
        Me.Label4.Caption = "Update : "
        CmdSitu False
        TxtLock 0
        ActionEdit = 0
        If Me.txtMH(0).Enabled = True Then Me.txtMH(0).SetFocus
        txtMH_GotFocus 0
        'Modify By Sindy 2014/7/31
        Combo1.ListIndex = 0
'        Combo1.Locked = True
        '2014/7/31 END
'        Me.txtMH(1).Text = strUserNum
'        Me.txtMH(1).Locked = True
'        Me.lblSupName.Caption = GetStaffName(Me.txtMH(1).Text)
        Seek_Now_Cp09 = ""
        Call Check1_Click
        
      Case 1 '按下修改modi
         CmdSitu False
         TxtLock 1
         ActionEdit = 1
        TmpMH(1) = ChangeTStringToWString(Me.txtMH(0).Text)
        'TmpMH(2) = Me.txtMH(1).Text
        TmpMH(2) = Trim(Left(Combo1.Text, 6))
        TmpMH(4) = Me.txtMH(3).Text
        Seek_Now_Cp09 = txtMH(14).Text
      Case 2 '按下刪除delete
         '若從個人進入, 若已核可的資料不可刪除
         '個人進入時, 只可刪除個人輸入的資料
         'Modify By Sindy 2014/7/31
         'If ProState = "1" And (Me.Check1.Value = vbChecked Or (txtMH(0) <> "" And txtMH(1) <> strUserNum)) Then
         If ProState = "1" And (Me.Check1.Value = vbChecked Or (txtMH(0) <> "" And Trim(Left(Combo1.Text, 6)) <> strUserNum)) Then
         '2014/7/31 END
         Else
             If Me.txtMH(0).Text = "" Then
                 MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
                 Exit Sub
             End If
             If DelMsg Then
                 'Modify By Sindy 2014/7/31
                 'StrSQLa = "Delete From ModifyHour Where MH01=" & ChangeTStringToWString(Me.txtMH(0).Text) & " And MH02='" & Me.txtMH(1).Text & "' And MH04='" & Me.txtMH(3).Text & "' "
                 StrSQLa = "Delete From ModifyHour Where MH01=" & ChangeTStringToWString(Me.txtMH(0).Text) & " And MH02='" & Trim(Left(Combo1.Text, 6)) & "' And MH04='" & Me.txtMH(3).Text & "' "
                 '2014/7/31 END
                 cnnConnection.Execute StrSQLa
                 strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour WHERE MH01||MH02||MH04>='" & MH(1) & MH(2) & MH(4) & "' Order By MH01, MH02, MH04 "
                  intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                 If intI = 1 Then
                    strExc(1) = "" & RsTemp.Fields("MH01").Value
                    strExc(2) = "" & RsTemp.Fields("MH02").Value
                    strExc(4) = "" & RsTemp.Fields("MH04").Value
                    ReadModifyHour strExc
                 Else
                     strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour WHERE MH01||MH02||MH04<='" & MH(1) & MH(2) & MH(4) & "' Order By MH01 Desc , MH02 Desc, MH04 Desc "
                      intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strExc(1) = "" & RsTemp.Fields("MH01").Value
                        strExc(2) = "" & RsTemp.Fields("MH02").Value
                        strExc(4) = "" & RsTemp.Fields("MH04").Value
                        ReadModifyHour strExc
                     Else
                        RsAction 0
                     End If
                 End If
                 strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour Order By MH01, MH02, MH04 "
                 intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
                 If intI = 1 Then
                      RsTemp.MoveFirst
                    strRsStart1 = "" & RsTemp.Fields("MH01").Value
                    strRsStart2 = "" & RsTemp.Fields("MH02").Value
                    strRsStart4 = "" & RsTemp.Fields("MH04").Value
                      RsTemp.MoveLast
                    strRsEnd1 = "" & RsTemp.Fields("MH01").Value
                    strRsEnd2 = "" & RsTemp.Fields("MH02").Value
                    strRsEnd4 = "" & RsTemp.Fields("MH04").Value
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
            'strExc(0) = "SELECT MH04 FROM ModifyHour where MH01=" & ChangeTStringToWString(txtMH(0).Text) & " and MH02='" & txtMH(1).Text & "' and MH12='" & txtMH(14).Text & "'"
            strExc(0) = "SELECT MH04 FROM ModifyHour where MH01=" & ChangeTStringToWString(txtMH(0).Text) & " and MH02='" & Trim(Left(Combo1.Text, 6)) & "' and MH12='" & txtMH(14).Text & "'"
            '2014/7/31 END
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               If MsgBox("當天已經有修改該收文號的紀錄，是否繼續？", vbYesNo + vbQuestion, "警告！") = vbNo Then
                  Exit Sub
               End If
            End If
            
            If Me.txtMH(5).Text = "" Or Me.txtMH(6).Text = "" Then
                Me.txtMH(5).Text = "": Me.txtMH(6).Text = "": Me.txtMH(7).Text = "": Me.txtMH(8).Text = ""
            End If
            'Modify By Sindy 2014/7/31
            'Me.txtMH(3).Text = GetSerialNo(ChangeTStringToWString(Me.txtMH(0).Text), Me.txtMH(1).Text)
            Me.txtMH(3).Text = GetSerialNo(ChangeTStringToWString(Me.txtMH(0).Text), Trim(Left(Combo1.Text, 6)))
            '2014/7/31 END
            
   On Error GoTo flgRollback
   
            cnnConnection.BeginTrans
            'Modify By Sindy 2014/7/31
            'StrSQLa = "Insert Into ModifyHour (MH01, MH02, MH04, MH05, MH06, MH07, MH08, MH09, MH10, MH11, MH12) Values(" & ChangeTStringToWString(Me.txtMH(0).Text) & ",'" & Me.txtMH(1).Text & "','" & txtMH(3).Text & "'," & Val(Me.txtMH(4).Text) & ",'" & Me.txtMH(5).Text & "','" & Me.txtMH(6).Text & "','" & Me.txtMH(7).Text & "','" & Me.txtMH(8).Text & "','" & ChgSQL(Me.txtMH(9).Text) & "','" & IIf(Me.Check1.Value = vbChecked, "V", "") & "'," & CNULL(Me.txtMH(14).Text) & ")"
            StrSQLa = "Insert Into ModifyHour (MH01, MH02, MH04, MH05, MH06, MH07, MH08, MH09, MH10, MH11, MH12) Values(" & ChangeTStringToWString(Me.txtMH(0).Text) & ",'" & Trim(Left(Combo1.Text, 6)) & "','" & txtMH(3).Text & "'," & Val(Me.txtMH(4).Text) & ",'" & Me.txtMH(5).Text & "','" & Me.txtMH(6).Text & "','" & Me.txtMH(7).Text & "','" & Me.txtMH(8).Text & "','" & ChgSQL(Me.txtMH(9).Text) & "','" & IIf(Me.Check1.Value = vbChecked, "V", "") & "'," & CNULL(Me.txtMH(14).Text) & ")"
            '2014/7/31 END
            cnnConnection.Execute StrSQLa, intI
         
            If bolMail2Boss = True Then subAddMail
            
            cnnConnection.CommitTrans
            
   On Error GoTo CheckingErr
            
            ActionEdit = 3
            TxtLock 3
            'Modify By Sindy 2014/7/31
            'If ChangeTStringToWString(Me.txtMH(0).Text) & Me.txtMH(1).Text & Me.txtMH(3).Text < strRsStart1 & strRsStart2 & strRsStart4 Then
            If ChangeTStringToWString(Me.txtMH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtMH(3).Text < strRsStart1 & strRsStart2 & strRsStart4 Then
            '2014/7/31 END
                strRsStart1 = ChangeTStringToWString(Me.txtMH(0).Text)
                'strRsStart2 = Me.txtMH(1).Text
                strRsStart2 = Trim(Left(Combo1.Text, 6))
                strRsStart4 = Me.txtMH(3).Text
            End If
            'Modify By Sindy 2014/7/31
            'If ChangeTStringToWString(Me.txtMH(0).Text) & Me.txtMH(1).Text & Me.txtMH(3).Text > strRsEnd1 & strRsEnd2 & strRsEnd4 Then
            If ChangeTStringToWString(Me.txtMH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtMH(3).Text > strRsEnd1 & strRsEnd2 & strRsEnd4 Then
            '2014/7/31 END
                strRsEnd1 = ChangeTStringToWString(Me.txtMH(0).Text)
                'strRsEnd2 = Me.txtMH(1).Text
                strRsEnd2 = Trim(Left(Combo1.Text, 6))
                strRsEnd4 = Me.txtMH(3).Text
            End If
            strExc(1) = ChangeTStringToWString(Me.txtMH(0).Text)
            'strExc(2) = Me.txtMH(1).Text
            strExc(2) = Trim(Left(Combo1.Text, 6))
            strExc(4) = Me.txtMH(3).Text
            ReadModifyHour strExc
            
         ElseIf ActionEdit = 1 Then '在修改狀態按Enter鍵
            If Not GetData Then Exit Sub
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            '檢查重複
            'Modify By Sindy 2014/7/31
            'If txtMH(0).Tag <> txtMH(0).Text Or txtMH(1).Tag <> txtMH(1).Text Or txtMH(14).Tag <> txtMH(14).Text Then
            If txtMH(0).Tag <> txtMH(0).Text Or Trim(Left(Combo1.Tag, 6)) <> Trim(Left(Combo1.Text, 6)) Or txtMH(14).Tag <> txtMH(14).Text Then
            '2014/7/31 END
               'Modify By Sindy 2014/7/31
               'strExc(0) = "SELECT MH04 FROM ModifyHour where MH01=" & ChangeTStringToWString(txtMH(0).Text) & " and MH02='" & txtMH(1).Text & "' and MH12='" & txtMH(14).Text & "'"
               strExc(0) = "SELECT MH04 FROM ModifyHour where MH01=" & ChangeTStringToWString(txtMH(0).Text) & " and MH02='" & Trim(Left(Combo1.Text, 6)) & "' and MH12='" & txtMH(14).Text & "'"
               '2014/7/31 END
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
               If intI = 1 Then
                  If MsgBox("當天已經有修改該收文號的紀錄，是否繼續？", vbYesNo + vbQuestion, "警告！") = vbNo Then
                     Exit Sub
                  End If
               End If
            End If
            If Me.txtMH(5).Text = "" Or Me.txtMH(6).Text = "" Then
                Me.txtMH(5).Text = "": Me.txtMH(6).Text = "": Me.txtMH(7).Text = "": Me.txtMH(8).Text = ""
            End If
            StrSQLa = ""
            If Me.txtMH(0).Text <> Me.txtMH(0).Tag Then
                StrSQLa = StrSQLa & " MH01=" & ChangeTStringToWString(Me.txtMH(0).Text) & ","
            End If
            'Modify By Sindy 2014/7/31
            'If Me.txtMH(1).Text <> Me.txtMH(1).Tag Then
            If Trim(Left(Combo1.Text, 6)) <> Trim(Left(Combo1.Tag, 6)) Then
               'StrSQLa = StrSQLa & " MH02='" & Val(Me.txtMH(1).Text) & "',"
               StrSQLa = StrSQLa & " MH02='" & Trim(Left(Combo1.Text, 6)) & "',"
            '2014/7/31 END
            End If
            'Modify By Sindy 2014/7/31
            'If txtMH(0).Tag <> txtMH(0).Text Or txtMH(1).Tag <> txtMH(1).Text Then
            If txtMH(0).Tag <> txtMH(0).Text Or Trim(Left(Combo1.Tag, 6)) <> Trim(Left(Combo1.Text, 6)) Then
               'MH04 = GetSerialNo(ChangeTStringToWString(Me.txtMH(0).Text), Me.txtMH(1).Text)
               MH04 = GetSerialNo(ChangeTStringToWString(Me.txtMH(0).Text), Trim(Left(Combo1.Text, 6)))
            '2014/7/31 END
               StrSQLa = StrSQLa & " MH04='" & MH04 & "',"
            Else
               MH04 = txtMH(3)
            End If
            
            If Me.txtMH(4).Text <> Me.txtMH(4).Tag Then
                StrSQLa = StrSQLa & " MH05=" & Val(Me.txtMH(4).Text) & ","
            End If
            If Me.txtMH(5).Text <> Me.txtMH(5).Tag Then
                StrSQLa = StrSQLa & " MH06='" & Me.txtMH(5).Text & "',"
            End If
            If Me.txtMH(6).Text <> Me.txtMH(6).Tag Then
                StrSQLa = StrSQLa & " MH07='" & Me.txtMH(6).Text & "',"
            End If
            If Me.txtMH(7).Text <> Me.txtMH(7).Tag Then
                StrSQLa = StrSQLa & " MH08='" & Me.txtMH(7).Text & "',"
            End If
            If Me.txtMH(8).Text <> Me.txtMH(8).Tag Then
                StrSQLa = StrSQLa & " MH09='" & Me.txtMH(8).Text & "',"
            End If
            If Me.txtMH(9).Text <> Me.txtMH(9).Tag Then
                StrSQLa = StrSQLa & " MH10='" & Me.txtMH(9).Text & "',"
            End If
            If Me.Check1.Value <> Me.Check1.Tag Then
                StrSQLa = StrSQLa & " MH11='" & IIf(Me.Check1.Value = vbChecked, "V", "") & "',"
            End If
            If Me.txtMH(14).Text <> Me.txtMH(14).Tag Then
                StrSQLa = StrSQLa & " MH12='" & Me.txtMH(14).Text & "',"
            End If

            If StrSQLa <> "" Then
                StrSQLa = Left(StrSQLa, Len(StrSQLa) - 1)
                
            ElseIf iUnlockChoice = 0 Then
                GoTo NoUpdate
                
            End If
            
   On Error GoTo flgRollback
   
            cnnConnection.BeginTrans
            
            If StrSQLa <> "" Then
               'Modify By Sindy 2014/7/31
               'StrSQLa = "Update ModifyHour Set " & StrSQLa & " Where MH01=" & Val(ChangeTStringToWString(Me.txtMH(0).Tag)) & " And MH02='" & Me.txtMH(1).Tag & "' And MH04='" & Me.txtMH(3).Tag & "' "
               StrSQLa = "Update ModifyHour Set " & StrSQLa & " Where MH01=" & Val(ChangeTStringToWString(Me.txtMH(0).Tag)) & " And MH02='" & Trim(Left(Combo1.Tag, 6)) & "' And MH04='" & Me.txtMH(3).Tag & "' "
               '2014/7/31 END
               cnnConnection.Execute StrSQLa
            End If
            If bolMail2Boss = True Then subAddMail
            If bolMail2Sales = True Then subAddRec
            If iUnlockChoice <> 0 Then subUnlock iUnlockChoice
            
            cnnConnection.CommitTrans
            
   On Error GoTo CheckingErr
   
            txtMH(3) = MH04

NoUpdate:
            ActionEdit = 3
            TxtLock 3
            
            strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour Order By MH01, MH02, MH04 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
                 RsTemp.MoveFirst
               strRsStart1 = "" & RsTemp.Fields("MH01").Value
               strRsStart2 = "" & RsTemp.Fields("MH02").Value
               strRsStart4 = "" & RsTemp.Fields("MH04").Value
                 RsTemp.MoveLast
               strRsEnd1 = "" & RsTemp.Fields("MH01").Value
               strRsEnd2 = "" & RsTemp.Fields("MH02").Value
               strRsEnd4 = "" & RsTemp.Fields("MH04").Value
            End If
            strExc(1) = ChangeTStringToWString(Me.txtMH(0).Text)
            'strExc(2) = Me.txtMH(1).Text
            strExc(2) = Trim(Left(Combo1.Text, 6))
            strExc(4) = Me.txtMH(3).Text
            ReadModifyHour strExc
            
         ElseIf ActionEdit = 2 Then '在查詢狀態按下Enter鍵
            If Me.txtMH(0).Text = "" Then
               MsgBox "修改日期不可空白，請重新輸入 !", vbCritical
               If Me.txtMH(0).Enabled = True Then Me.txtMH(0).SetFocus
               txtMH_GotFocus 0
               Exit Sub
            End If
            intI = 1
            'Modify By Sindy 2014/7/31
            'strExc(0) = "SELECT COUNT(*) FROM ModifyHour WHERE MH01=" & ChangeTStringToWString(Me.txtMH(0).Text) & " And MH02='" & Me.txtMH(1).Text & "' And MH04= '" & Me.txtMH(3).Text & "'"
            strExc(0) = "SELECT COUNT(*) FROM ModifyHour WHERE MH01=" & ChangeTStringToWString(Me.txtMH(0).Text) & " And MH02='" & Trim(Left(Combo1.Text, 6)) & "' And MH04= '" & Me.txtMH(3).Text & "'"
            '2014/7/31 END
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) = 0 Then
                  MsgBox "查無此修改記錄 !", vbCritical
                    strExc(1) = TmpMH(1)
                    strExc(2) = TmpMH(2)
                    strExc(4) = TmpMH(4)
               Else
                    strExc(1) = ChangeTStringToWString(Me.txtMH(0).Text)
                    'strExc(2) = Me.txtMH(1).Text
                    strExc(2) = Trim(Left(Combo1.Text, 6))
                    strExc(4) = Me.txtMH(3).Text
               End If
            End If
            ReadModifyHour strExc
         End If
         
         '發信
         PUB_SendMailCache
         CmdSitu True
      Case 4 'cancel
         If ActionEdit <> 2 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
         End If
         CmdSitu True
        If TmpMH(1) = "" Then TmpMH(1) = strRsStart1
        If TmpMH(2) = "" Then TmpMH(2) = strRsStart2
        If TmpMH(4) = "" Then TmpMH(4) = strRsStart4
        strExc(1) = TmpMH(1)
        strExc(2) = TmpMH(2)
        strExc(4) = TmpMH(4)
         ActionEdit = 3
         ReadModifyHour strExc
         TxtLock 3
      Case 5 'query
        TmpMH(1) = ChangeTStringToWString(Me.txtMH(0).Text)
        'TmpMH(2) = Me.txtMH(1).Text
        TmpMH(2) = Trim(Left(Combo1.Text, 6))
        TmpMH(4) = Me.txtMH(3).Text
         CmdSitu False
         TxtLock 2
         ActionEdit = 2
         If Me.txtMH(0).Enabled = True Then Me.txtMH(0).SetFocus
         txtMH_GotFocus 0
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
         strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour WHERE MH01=" & strRsStart1 & " And MH02 ='" & strRsStart2 & "' And MH04= '" & strRsStart4 & "' "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "" & RsTemp.Fields("MH01").Value
            strExc(2) = "" & RsTemp.Fields("MH02").Value
            strExc(4) = "" & RsTemp.Fields("MH04").Value
        Else
            strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour WHERE MH01||MH02||MH04>='" & strRsStart1 & strRsStart2 & strRsStart4 & "' Order By MH01, MH02, MH04 "
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                strExc(1) = "" & RsTemp.Fields("MH01").Value
                strExc(2) = "" & RsTemp.Fields("MH02").Value
                strExc(4) = "" & RsTemp.Fields("MH04").Value
                strRsStart1 = strExc(1)
                strRsStart2 = strExc(2)
                strRsStart4 = strExc(4)
            End If
         End If
      Case 1 '前一筆
         'Modify By Sindy 2014/7/31
         'If ChangeTStringToWString(Me.txtMH(0).Text) & Me.txtMH(1).Text & Me.txtMH(3).Text = strRsStart1 & strRsStart2 & strRsStart4 Then
         If ChangeTStringToWString(Me.txtMH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtMH(3).Text = strRsStart1 & strRsStart2 & strRsStart4 Then
         '2014/7/31 END
            Beep
            Screen.MousePointer = vbDefault
            DataErrorMessage 6
            Exit Sub
         Else
            intI = 1
            'Modify By Sindy 2014/7/31
            'strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour WHERE MH01||MH02||MH04<'" & ChangeTStringToWString(Me.txtMH(0).Text) & Me.txtMH(1).Text & Me.txtMH(3).Text & "' Order By MH01 Desc, MH02 Desc, MH04 Desc "
            strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour WHERE MH01||MH02||MH04<'" & ChangeTStringToWString(Me.txtMH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtMH(3).Text & "' Order By MH01 Desc, MH02 Desc, MH04 Desc "
            '2014/7/31 END
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               strExc(1) = "" & RsTemp.Fields("MH01").Value
               strExc(2) = "" & RsTemp.Fields("MH02").Value
               strExc(4) = "" & RsTemp.Fields("MH04").Value
            End If
         End If
      Case 2 '後一筆
         'Modify By Sindy 2014/7/31
         'If ChangeTStringToWString(Me.txtMH(0).Text) & Me.txtMH(1).Text & Me.txtMH(3).Text = strRsEnd1 & strRsEnd2 & strRsEnd4 Then
         If ChangeTStringToWString(Me.txtMH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtMH(3).Text = strRsEnd1 & strRsEnd2 & strRsEnd4 Then
         '2014/7/31 END
            Beep
            Screen.MousePointer = vbDefault
            DataErrorMessage 7
            Exit Sub
         Else
            'Modify By Sindy 2014/7/31
            'strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour WHERE MH01||MH02||MH04>'" & ChangeTStringToWString(Me.txtMH(0).Text) & Me.txtMH(1).Text & Me.txtMH(3).Text & "' Order By MH01, MH02, MH04 "
            strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour WHERE MH01||MH02||MH04>'" & ChangeTStringToWString(Me.txtMH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtMH(3).Text & "' Order By MH01, MH02, MH04 "
            '2014/7/31 END
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               strExc(1) = "" & RsTemp.Fields("MH01").Value
               strExc(2) = "" & RsTemp.Fields("MH02").Value
               strExc(4) = "" & RsTemp.Fields("MH04").Value
            End If
         End If
      Case 3 '最後筆
         strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour WHERE MH01=" & strRsEnd1 & " And MH02='" & strRsEnd2 & "' And MH04='" & strRsEnd4 & "'"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "" & RsTemp.Fields("MH01").Value
            strExc(2) = "" & RsTemp.Fields("MH02").Value
            strExc(4) = "" & RsTemp.Fields("MH04").Value
        Else
            strExc(0) = "SELECT MH01, MH02, MH04 FROM ModifyHour WHERE MH01||MH02||MH04<='" & strRsEnd1 & strRsEnd2 & strRsEnd4 & "' Order By MH01 Desc, MH02 Desc, MH04 Desc "
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                strExc(1) = "" & RsTemp.Fields("MH01").Value
                strExc(2) = "" & RsTemp.Fields("MH02").Value
                strExc(4) = "" & RsTemp.Fields("MH04").Value
                strRsEnd1 = strExc(1)
                strRsEnd2 = strExc(2)
                strRsEnd4 = strExc(4)
            End If
         End If
   End Select
   ReadModifyHour strExc
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
    Me.txtMH(0).Locked = False
    'Me.txtMH(1).Locked = False
    Combo1.Locked = False
    Me.txtMH(4).Locked = False
    Me.txtMH(5).Locked = False
    Me.txtMH(6).Locked = False
    Me.txtMH(7).Locked = False
    Me.txtMH(8).Locked = False
    Me.txtMH(9).Locked = False
    Me.txtMH(14).Locked = False
    Me.txtMH(0).Text = ""
    'Me.txtMH(1).Text = ""
    Me.txtMH(3).Text = ""
    Me.txtMH(4).Text = ""
    Me.txtMH(5).Text = ""
    Me.txtMH(6).Text = ""
    Me.txtMH(7).Text = ""
    Me.txtMH(8).Text = ""
    Me.txtMH(9).Text = ""
    Me.txtMH(14).Text = ""
    Me.lblCaseCnt.Caption = ""
    Me.lblSANo.Caption = ""
    Me.lblSAName.Caption = ""
    Combo1.Text = ""
    'Me.lblSupName.Caption = ""
    If ProState = "2" Then
        Me.Check1.Enabled = True
        Me.Check1.Value = vbUnchecked
        Check2.Value = vbUnchecked
        'Check2.Enabled = True
        Check2.Enabled = False
    Else
        Me.Check1.Enabled = False
        Me.Check1.Value = vbUnchecked
        Check2.Value = vbUnchecked
        Check2.Enabled = False
    End If
    cmdSelCp09.Enabled = True
    
Case 1 '修改
    Me.txtMH(0).Locked = False
    'Me.txtMH(1).Locked = True
    Combo1.Locked = True
    Me.txtMH(4).Locked = False
    Me.txtMH(5).Locked = False
    Me.txtMH(6).Locked = False
    Me.txtMH(7).Locked = False
    Me.txtMH(8).Locked = False
    Me.txtMH(9).Locked = False
    Me.txtMH(14).Locked = False
    If ProState = "2" Then
         Me.Check1.Enabled = True
         cmdSelCp09.Enabled = True
         If Check1.Value = 1 And Check2.Value = 1 Then
            Check2.Enabled = True
         End If
    Else
         Me.Check1.Enabled = False
         cmdSelCp09.Enabled = False
         Check2.Enabled = False
    End If

Case 2 '查詢
    Me.txtMH(0).Locked = False
    'Me.txtMH(1).Locked = False
    Combo1.Locked = False
    Me.txtMH(4).Locked = True
    Me.txtMH(5).Locked = True
    Me.txtMH(6).Locked = True
    Me.txtMH(7).Locked = True
    Me.txtMH(8).Locked = True
    Me.txtMH(9).Locked = True
    Me.txtMH(14).Locked = True
    Me.txtMH(0).Text = ""
    'Me.txtMH(1).Text = ""
    Me.txtMH(3).Text = ""
    Me.txtMH(4).Text = ""
    Me.txtMH(5).Text = ""
    Me.txtMH(6).Text = ""
    Me.txtMH(7).Text = ""
    Me.txtMH(8).Text = ""
    Me.txtMH(9).Text = ""
    Me.txtMH(14).Text = ""
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
    Me.txtMH(0).Locked = True
    'Me.txtMH(1).Locked = True
    Combo1.Locked = True
    Me.txtMH(4).Locked = True
    Me.txtMH(5).Locked = True
    Me.txtMH(6).Locked = True
    Me.txtMH(7).Locked = True
    Me.txtMH(8).Locked = True
    Me.txtMH(9).Locked = True
    Me.txtMH(14).Locked = True
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
        ReadModifyHour strExc
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
On Error Resume Next
    Select Case Me.SSTab1.Tab
    Case 0
        Me.txtMH(0).SetFocus
        txtMH_GotFocus 0
        Me.cmdQuery(0).Default = False
    Case 1
        Me.txtMH(10).SetFocus
        txtMH_GotFocus 12
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
   If Me.txtMH(0).Text = "" Then
      MsgBox "修改日期不可空白 !", vbCritical
      Me.txtMH(0).SetFocus
      txtMH_GotFocus 0
      Exit Function
   End If
   'Modify By Sindy 2014/7/31
   'If Me.txtMH(1).Text = "" Then
   If Trim(Combo1.Text) = "" Then
   '2014/7/31 END
      MsgBox "修改人員不可空白 !", vbCritical
      Combo1.SetFocus
      Exit Function
   End If

   If Me.txtMH(4).Text = "" Then
      MsgBox "修改數時不可空白 !", vbCritical
      Me.txtMH(4).SetFocus
      txtMH_GotFocus 4
      Exit Function
   End If
    If Me.txtMH(5).Text <> "" And Me.txtMH(6).Text <> "" Then
        '案號補滿
        If Me.txtMH(7).Text = "" Then Me.txtMH(7).Text = "0"
        If Me.txtMH(8).Text = "" Then Me.txtMH(8).Text = "00"
        StrSQLa = "Select PA01 From Patent Where PA01='" & Me.txtMH(5).Text & "' And PA02='" & Me.txtMH(6).Text & "' And PA03='" & Me.txtMH(7).Text & "' And PA04='" & Me.txtMH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select TM01 From Trademark Where TM01='" & Me.txtMH(5).Text & "' And TM02='" & Me.txtMH(6).Text & "' And TM03='" & Me.txtMH(7).Text & "' And TM04='" & Me.txtMH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select LC01 From Lawcase Where LC01='" & Me.txtMH(5).Text & "' And LC02='" & Me.txtMH(6).Text & "' And LC03='" & Me.txtMH(7).Text & "' And LC04='" & Me.txtMH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select HC01 From Hirecase Where HC01='" & Me.txtMH(5).Text & "' And HC02='" & Me.txtMH(6).Text & "' And HC03='" & Me.txtMH(7).Text & "' And HC04='" & Me.txtMH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select SP01 From Servicepractice Where SP01='" & Me.txtMH(5).Text & "' And SP02='" & Me.txtMH(6).Text & "' And SP03='" & Me.txtMH(7).Text & "' And SP04='" & Me.txtMH(8).Text & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount <= 0 Then
            MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
            Me.txtMH(5).SetFocus
            txtMH_GotFocus 5
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Function
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
    If Me.txtMH(14).Text <> "" Then
        StrSQLa = "Select * From Caseprogress Where CP09='" & Me.txtMH(14).Text & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic
        If rsA.RecordCount <= 0 Then
            MsgBox "無此收文號資料!!!", vbExclamation + vbOKOnly
            Me.txtMH(14).SetFocus
            txtMH_GotFocus 14
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Function
        Else
            If Me.txtMH(5).Text <> "" Or Me.txtMH(6).Text <> "" Then
                If Me.txtMH(5).Text <> "" & rsA.Fields(0).Value Or Me.txtMH(6).Text <> "" & rsA.Fields(1).Value Or Me.txtMH(7).Text <> "" & rsA.Fields(2).Value Or Me.txtMH(8).Text <> "" & rsA.Fields(3).Value Then
                    MsgBox "此收文號對應的本所案號錯誤!!!", vbExclamation + vbOKOnly
                    Me.txtMH(14).SetFocus
                    txtMH_GotFocus 14
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
   MH(1) = ChangeTStringToWString(Me.txtMH(0).Text)
   'MH(2) = Me.txtMH(1).Text
   MH(2) = Trim(Left(Combo1.Text, 6))
   MH(5) = Me.txtMH(4).Text
   MH(6) = Me.txtMH(5).Text
   MH(7) = Me.txtMH(6).Text
   MH(8) = Me.txtMH(7).Text
   MH(9) = Me.txtMH(8).Text
   MH(10) = Me.txtMH(9).Text
   MH(11) = IIf(Me.Check1.Value = vbChecked, "V", "")
   MH(12) = Me.txtMH(14).Text
   GetData = True
End Function

Private Function TxtValidate() As Boolean
   Dim objTxt As Object
   Dim ii As Integer
   Dim Cancel As Boolean
   
   
   TxtValidate = False
   For Each objTxt In Me.txtMH
       If objTxt.Enabled = True Then
          Cancel = False
          txtMH_Validate objTxt.Index, Cancel
          If Cancel = True Then
             Exit Function
          End If
       End If
   Next
   If Me.txtMH(0).Text = "" Then MsgBox "修改日期不可空白！", vbCritical, "嚴重錯誤！": txtMH(0).SetFocus: Exit Function
   'If Me.txtMH(1).Text = "" Then MsgBox "修改人員不可空白！", vbCritical, "嚴重錯誤！": txtMH(1).SetFocus: Exit Function
   If Trim(Combo1.Text) = "" Then MsgBox "修改人員不可空白！", vbCritical, "嚴重錯誤！": Combo1.SetFocus: Exit Function
   If Me.txtMH(14).Text = "" Then MsgBox "收文號不可空白！", vbCritical, "嚴重錯誤！": txtMH(14).SetFocus: Exit Function
   
   bolMail2Boss = False
   bolMail2Sales = False
   iUnlockChoice = 0
   If ProState = "1" Then
      If Val(txtMH(4)) > 3 And (ActionEdit = 0 Or Val(txtMH(4)) > Val(txtMH(4).Tag)) Then
         frm090633_2.p_Choice = 1
         Set frm090633_2.p_Parent = Me
         frm090633_2.lblAlert = "本次修改時數超過 3 小時"
         frm090633_2.Show vbModal
         If p_iRtn = 0 Then
            Exit Function
         ElseIf p_iRtn = 1 Then
            bolMail2Boss = True
         ElseIf p_iRtn = 2 Then
            MsgBox "因未呈報時數將自動改為 3 小時！"
            txtMH(4) = 3
         End If
      End If
      
   Else
      If Check2.Value = 0 Then
         If Check1.Value = vbChecked And Check1.Value <> Check1.Tag Then
            strExc(1) = ""
            If Val(txtMH(4)) > 3 Then
               strExc(1) = "1"
               strExc(2) = "本次修改時數超過 3 小時"
            ElseIf ChkTotalOver = True Then
               strExc(1) = "2"
               strExc(2) = "該收文號修改累計時數超過 4 小時"
            End If
            If strExc(1) <> "" Then
               frm090633_2.p_Choice = 2
               Set frm090633_2.p_Parent = Me
               frm090633_2.lblAlert = strExc(2)
               frm090633_2.Show vbModal
               If p_iRtn = 0 Then
                  Exit Function
               ElseIf p_iRtn = 1 Then
                  bolMail2Sales = True
               Else
                  'MsgBox strExc(2) & ",您已選擇不列管!!"
               End If
            End If
         End If
      End If
      
      If ActionEdit = 1 And Check2.Value = vbUnchecked And Check2.Value <> Check2.Tag Then
         frm090633_2.p_Choice = 3
         frm090633_2.lblAlert = "解除管制"
         Set frm090633_2.p_Parent = Me
         frm090633_2.Show vbModal
         iUnlockChoice = p_iRtn
         If p_iRtn = 0 Then
            Exit Function
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

Private Sub txtMH_Change(Index As Integer)
   Select Case Index
   Case 4 '修改時數
      If Me.txtMH(Index).Text <> "" Then
         Me.lblCaseCnt.Caption = Format(Val(Me.txtMH(Index).Text) * 0.2, "0.00")
      Else
         Me.lblCaseCnt.Caption = ""
      End If
      
   Case 14 '收文號
      Me.lblSANo.Caption = ""
      Me.lblSAName.Caption = ""
      If Me.txtMH(Index).Text <> "" Then
         strExc(0) = "select cp13,st02 from caseprogress,staff where cp09='" & txtMH(Index) & "' and st01(+)=cp13"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Me.lblSANo.Caption = "" & RsTemp("cp13")
            Me.lblSAName.Caption = "" & RsTemp("st02")
         End If
      End If
   End Select
End Sub

Private Sub txtMH_GotFocus(Index As Integer)
    TextInverse Me.txtMH(Index)
End Sub

'Modified by Lydia 2022/01/03 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txtMH_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
    Case 1, 2, 7, 5, 12, 13, 14 '系統類別, 修改人員部門別, 收文號
        KeyAscii = UpperCase(KeyAscii)
    Case 0
        If KeyAscii = 47 Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub txtMH_LostFocus(Index As Integer)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

    Select Case Index
    Case 8 '本所案號
        If Me.txtMH(5).Text <> "" And Me.txtMH(6).Text <> "" Then
            'Add By Cheng 2003/08/01
            '案號補滿
            If Me.txtMH(7).Text = "" Then Me.txtMH(7).Text = "0"
            If Me.txtMH(8).Text = "" Then Me.txtMH(8).Text = "00"
            StrSQLa = "Select PA01 From Patent Where " & ChgPatent(Me.txtMH(5).Text & Me.txtMH(6).Text & Me.txtMH(7).Text & Me.txtMH(8).Text)
            StrSQLa = StrSQLa & " Union Select TM01 From Trademark Where " & ChgTradeMark(Me.txtMH(5).Text & Me.txtMH(6).Text & Me.txtMH(7).Text & Me.txtMH(8).Text)
            StrSQLa = StrSQLa & " Union Select LC01 From Lawcase Where " & ChgLawcase(Me.txtMH(5).Text & Me.txtMH(6).Text & Me.txtMH(7).Text & Me.txtMH(8).Text)
            StrSQLa = StrSQLa & " Union Select HC01 From Hirecase Where " & ChgHirecase(Me.txtMH(5).Text & Me.txtMH(6).Text & Me.txtMH(7).Text & Me.txtMH(8).Text)
            StrSQLa = StrSQLa & " Union Select SP01 From Servicepractice Where " & ChgService(Me.txtMH(5).Text & Me.txtMH(6).Text & Me.txtMH(7).Text & Me.txtMH(8).Text)
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount <= 0 Then
                MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
                Me.txtMH(5).SetFocus
                txtMH_GotFocus 5
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
    Case 11 '修改日期
        If Me.txtMH(10).Text <> "" And Me.txtMH(11).Text <> "" Then
            If Val(Me.txtMH(10).Text) > Val(Me.txtMH(11).Text) Then
                MsgBox "修改日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                Me.txtMH(10).SetFocus
                txtMH_GotFocus 10
                Exit Sub
            End If
        End If
    Case 12 '修改人員部門
        If Me.txtMH(12).Text <> "" And Me.txtMH(13).Text <> "" Then
            If Me.txtMH(12).Text > Me.txtMH(13).Text Then
                MsgBox "修改人員部門範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                Me.txtMH(12).SetFocus
                txtMH_GotFocus 12
                Exit Sub
            End If
        End If
    End Select
End Sub

Private Sub txtMH_Validate(Index As Integer, Cancel As Boolean)
   If Me.txtMH(Index).Text = "" Then Exit Sub
   Select Case Index
   Case 0 '修改日期
       If CheckIsTaiwanDate(Me.txtMH(Index).Text) = False Then
           Cancel = True
       End If
'   Case 1 '修改人員
'       Me.lblSupName.Caption = GetStaffName(Me.txtMH(Index).Text)
'       If Me.lblSupName.Caption = "" And Check1.Value = vbUnchecked Then
'           MsgBox "修改人員輸入錯誤!!!", vbExclamation + vbOKOnly
'           Cancel = True
'       End If

   Case 4 '修改時數
      If Val(Me.txtMH(4)) < 0.5 Then
         MsgBox "修改時數未達 [ 0.5 ] 小時!!!", vbExclamation + vbOKOnly
         Cancel = True
      End If
      
   Case 9 '備註
       If CheckLengthIsOK(Me.txtMH(Index).Text, 200) = False Then
          Cancel = True
       End If
   Case 10, 11 '修改日期區間
       If CheckIsTaiwanDate(Me.txtMH(Index).Text) = False Then
           Cancel = True
       End If
   Case 14 '收文號
       Cancel = Not GetOurCaseNo(Me.txtMH(14).Text)
   End Select
   If Cancel = True Then txtMH_GotFocus Index
End Sub

Private Function QueryData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nRow As Integer
   
    QueryData = False
    InitialGridList
    strSql = ""
    If Me.txtMH(10).Text <> "" Then
        strSql = strSql & " And MH01>=" & DBDATE(Me.txtMH(10).Text) & " "
    End If
    If Me.txtMH(11).Text <> "" Then
        strSql = strSql & " And MH01<=" & DBDATE(Me.txtMH(11).Text) & " "
    End If
    If Me.txtMH(12).Text <> "" Then
        strSql = strSql & " And S1.ST03>='" & ChgSQL(Me.txtMH(12).Text) & "' "
    End If
    If Me.txtMH(13).Text <> "" Then
        strSql = strSql & " And S1.ST03<='" & ChgSQL(Me.txtMH(13).Text) & "' "
    End If
    If Check3.Value = vbChecked Then
        'Modified by Morgan 2011/12/12 + B類收文控管,智權人員會收A類
        strSql = strSql & " and exists(select * from caseprogress where cp43=MH12 and cp09>'B' and  cp10='944' and cp27||cp57 is null)"
    End If
    
    If Trim(txtSystem.Text) <> "" And Trim(txtCode(0).Text) <> "" Then
        strSql = strSql & " and MH06='" & Trim(txtSystem.Text) & "' and MH07='" & Trim(txtCode(0).Text) & "' "
        If Trim(txtCode(1).Text) <> "" Then
            strSql = strSql & " and MH08='" & Trim(txtCode(1).Text) & "' "
        Else
            strSql = strSql & " and MH08='0' "
        End If
        If Trim(txtCode(2).Text) <> "" Then
            strSql = strSql & " and MH09='" & Trim(txtCode(2).Text) & "' "
        Else
            strSql = strSql & " and MH09='00' "
        End If
    End If
    
    strSql = "SELECT Decode(MH01, Null, Null, MH01-19110000), MH02, S1.ST02, CP13, S2.ST02, MH04, MH12, Decode(MH06, Null, '', MH06||Decode(MH07, Null, '','-'||MH07||Decode(MH08, Null, '','-'||MH08||Decode(MH09, Null, '','-'||MH09)))), MH05, MH05 * 0.2, MH11, MH10 FROM ModifyHour,caseprogress, Staff S1, Staff S2 " & _
         "WHERE MH02=S1.ST01(+) AND cp09(+)=MH12 and CP13=S2.ST01(+) " & strSql & " Order By 1, 2, 4, 5 "
    
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
    grdList.Text = "修改日期"
    grdList.ColWidth(1) = 800
    grdList.ColAlignment(1) = flexAlignCenterCenter
    grdList.col = 2
    'grdList.Text = "修改人員代號"
    grdList.ColWidth(2) = 0
    grdList.ColAlignment(2) = flexAlignCenterCenter
    grdList.col = 3
    grdList.Text = "修改人員"
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
    grdList.Text = "修改時數"
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
Private Function GetSerialNo(strMH01 As String, strMH02 As String) As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   'Modified by Morgan 2012/7/13 沒資料改預設 0 否則 win7 會有錯誤(資料提供者或其他服務傳回E_FAIL狀態)
   StrSQLa = "Select nvl(max(MH04),0) as MH04 From ModifyHour Where MH01=" & strMH01 & " And MH02='" & strMH02 & "'  Order By MH04 Desc "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsA.EOF And Not rsA.BOF Then
       GetSerialNo = Format(Val(rsA("MH04").Value) + 1, "000")
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
        '修改日期
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 1)
        '修改人員
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
        '修改時數
        Printer.CurrentX = PLeft(5) + Printer.TextWidth("修改時數") - Printer.TextWidth(Format(Me.grdList.TextMatrix(ii, 9), "0.0"))
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
Printer.Print "修改記錄明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "修改日期：" & Format(ChangeTStringToTDateString(Me.txtMH(10).Text) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Me.txtMH(11).Text)
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
Printer.Print "修改人員部門別：" & Me.txtMH(12).Text & " " & IIf(Me.txtMH(12).Text <> "" Or Me.txtMH(13).Text <> "", "－", "") & " " & Me.txtMH(13).Text
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
Printer.Print "修改日期"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "修改人員"
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
Printer.Print "修改時數"
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
         Me.txtMH(5).Text = "" & rsA("CP01").Value
         Me.txtMH(6).Text = "" & rsA("CP02").Value
         Me.txtMH(7).Text = "" & rsA("CP03").Value
         Me.txtMH(8).Text = "" & rsA("CP04").Value
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

Private Sub subAddRec()
   Dim strBCP09 As String
   
   strBCP09 = AutoNo("B", 6) 'B類總收文號
   StrSQLa = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13" & _
      ",cp14,cp20,cp43) " & _
      " select cp01,cp02,cp03,cp04," & strSrvDate(1) & ",'" & strBCP09 & "','944',cp12,cp13" & _
      ",'" & Trim(Left(Combo1.Text, 6)) & "','N','" & txtMH(14) & "' from caseprogress where cp09='" & txtMH(14) & "'"
   cnnConnection.Execute StrSQLa, intI
      
   StrSQLa = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
      " select '" & strUserNum & "',cp13,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'案('||cp09||')的修改幅度過大，" & _
      "已內部收文會稿修改，請處理加收修改費用事宜或回覆說明。','如旨'" & _
      " from caseprogress where cp09='" & txtMH(14) & "'"
      
   cnnConnection.Execute StrSQLa, intI
End Sub

Private Sub subAddMail()
   Dim strMod As String
   Dim strSubject As String, strContent As String
   
   If ActionEdit = 0 Then
      strMod = "新增"
   Else
      strMod = "修改"
   End If
   
   strSubject = strMod & "<<修改記錄>>輸入時數超過 3 小時通知"
   strContent = "修改日期：" & ChangeTStringToTDateString(Me.txtMH(0).Text) & vbCrLf & _
         "修改人員：" & Combo1.Text & vbCrLf & _
         "修改時數：" & Me.txtMH(4).Text & vbCrLf & _
         "折算件數：" & Me.lblCaseCnt.Caption & vbCrLf & _
         "收 文 號：" & Me.txtMH(14).Text & vbCrLf & _
         "本所案號：" & Me.txtMH(5).Text & "-" & Me.txtMH(6).Text & "-" & Me.txtMH(7).Text & "-" & Me.txtMH(8).Text & vbCrLf & _
         "智權人員：" & Me.lblSANo.Caption & " " & Me.lblSAName.Caption & vbCrLf & _
         "備　　註：" & Me.txtMH(9).Text & vbCrLf & _
         "是否主管核可：" & IIf(Me.Check1.Value = vbChecked, "是", "否") & vbCrLf & vbCrLf & _
         strMod & "資料人員：" & strUserNum & " " & GetStaffName(strUserNum)
   
   StrSQLa = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " select '" & strUserNum & "',OMAN,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & strSubject & "','" & ChgSQL(strContent) & "','71011'" & _
            " from staff,SetSpecMan where st01='" & Trim(Left(Combo1.Text, 6)) & "' and OCODE(+)=decode(st06,'1','A2','2','A3','3','A4','4','A5')"
   cnnConnection.Execute StrSQLa, intI
End Sub

Private Sub subUnlock(pChoice As Integer)
   'Modified by Morgan 2011/12/12 + B類收文控管,智權人員會收A類
   If pChoice = 1 Then
      StrSQLa = "update caseprogress set cp27=" & strSrvDate(1) & " where cp43='" & txtMH(14) & "' and cp09>'B' and cp10='944' and cp27||cp57 is null"
      cnnConnection.Execute StrSQLa
      
   ElseIf pChoice = 2 Then
      StrSQLa = "delete from caseprogress where CP01='" & txtMH(5) & "' AND CP02='" & txtMH(6) & "' AND CP03='" & txtMH(7) & "'" & _
         " AND CP04='" & txtMH(8) & "' AND cp43='" & txtMH(14) & "' and cp09>'B' and cp10='944' and cp27||cp57 is null"
      Pub_SeekTbLog StrSQLa
      cnnConnection.Execute StrSQLa, intI
      
   End If
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
      If InStr("修改日期,序號,修改時數,折算件數", Me.grdList.Text) > 0 Then
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
