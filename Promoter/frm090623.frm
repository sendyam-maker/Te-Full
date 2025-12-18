VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090623 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人/繪圖人員支援記錄維護"
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
            Picture         =   "frm090623.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090623.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090623.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090623.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090623.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090623.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090623.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090623.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090623.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090623.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090623.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   150
      TabIndex        =   26
      Top             =   690
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   8276
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm090623.frx":20F4
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
      Tab(0).Control(12)=   "lblSAName"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Combo1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label4"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtSH09"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(10)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtSH(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtSH(7)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtSH(8)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtSH(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtSH(5)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtSH(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtSH(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtSH(4)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Check1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtSH(14)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdSelCp09"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Check2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Check4"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Check5"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtCnt"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm090623.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check6"
      Tab(1).Control(1)=   "txtSH(15)"
      Tab(1).Control(2)=   "txtSystem"
      Tab(1).Control(3)=   "txtCode(2)"
      Tab(1).Control(4)=   "txtCode(1)"
      Tab(1).Control(5)=   "txtCode(0)"
      Tab(1).Control(6)=   "Check3"
      Tab(1).Control(7)=   "txtSH(13)"
      Tab(1).Control(8)=   "txtSH(12)"
      Tab(1).Control(9)=   "txtSH(11)"
      Tab(1).Control(10)=   "txtSH(10)"
      Tab(1).Control(11)=   "cmdQuery(1)"
      Tab(1).Control(12)=   "cmdQuery(0)"
      Tab(1).Control(13)=   "grdList"
      Tab(1).Control(14)=   "Label5"
      Tab(1).Control(15)=   "Label1(9)"
      Tab(1).Control(16)=   "Label1(8)"
      Tab(1).Control(17)=   "Line3"
      Tab(1).Control(18)=   "Label1(5)"
      Tab(1).Control(19)=   "Line2"
      Tab(1).Control(20)=   "Label1(4)"
      Tab(1).ControlCount=   21
      Begin VB.TextBox txtCnt 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   264
         Left            =   5688
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2388
         Width           =   324
      End
      Begin VB.CheckBox Check6 
         Caption         =   "僅查不收文"
         Height          =   225
         Left            =   -71850
         TabIndex        =   50
         Top             =   480
         Width           =   1275
      End
      Begin VB.CheckBox Check5 
         Caption         =   "不收文"
         Height          =   225
         Left            =   6060
         TabIndex        =   49
         Top             =   3705
         Width           =   975
      End
      Begin VB.CheckBox Check4 
         Caption         =   "計算支援"
         Height          =   225
         Left            =   4440
         TabIndex        =   43
         Top             =   3705
         Width           =   1605
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   15
         Left            =   -70020
         MaxLength       =   6
         TabIndex        =   41
         Top             =   1110
         Width           =   945
      End
      Begin VB.TextBox txtSystem 
         Height          =   300
         Left            =   -74040
         MaxLength       =   3
         TabIndex        =   19
         Top             =   1110
         Width           =   732
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   2
         Left            =   -71610
         MaxLength       =   2
         TabIndex        =   22
         Top             =   1110
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   1
         Left            =   -72030
         MaxLength       =   1
         TabIndex        =   21
         Top             =   1110
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   0
         Left            =   -73275
         MaxLength       =   6
         TabIndex        =   20
         Top             =   1110
         Width           =   1212
      End
      Begin VB.CheckBox Check3 
         Caption         =   "僅查管制中案件未輸入本所案號或收文號"
         Height          =   195
         Left            =   -72360
         TabIndex        =   18
         Top             =   810
         Width           =   3555
      End
      Begin VB.CheckBox Check2 
         Caption         =   "管制案件"
         Height          =   225
         Left            =   2520
         TabIndex        =   13
         Top             =   3705
         Width           =   1605
      End
      Begin VB.CommandButton cmdSelCp09 
         Caption         =   "選擇收文號"
         Height          =   300
         Left            =   3300
         TabIndex        =   10
         Top             =   1980
         Width           =   1200
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   14
         Left            =   1410
         MaxLength       =   9
         TabIndex        =   5
         Top             =   1980
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "主管核可"
         Height          =   315
         Left            =   540
         TabIndex        =   12
         Top             =   3660
         Width           =   1785
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   4
         Left            =   1410
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1650
         Width           =   945
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   2
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1290
         Width           =   945
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   0
         Left            =   1410
         MaxLength       =   7
         TabIndex        =   0
         Top             =   570
         Width           =   945
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   5
         Left            =   1410
         MaxLength       =   3
         TabIndex        =   6
         Top             =   2370
         Width           =   525
      End
      Begin VB.TextBox txtSH 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   4650
         TabIndex        =   1
         Top             =   570
         Width           =   525
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   8
         Left            =   3420
         MaxLength       =   2
         TabIndex        =   9
         Top             =   2370
         Width           =   435
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   7
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   8
         Top             =   2370
         Width           =   315
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   6
         Left            =   2010
         MaxLength       =   6
         TabIndex        =   7
         Top             =   2370
         Width           =   915
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   13
         Left            =   -73020
         MaxLength       =   3
         TabIndex        =   17
         Top             =   780
         Width           =   525
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   12
         Left            =   -73650
         MaxLength       =   3
         TabIndex        =   16
         Top             =   780
         Width           =   525
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   11
         Left            =   -72990
         MaxLength       =   7
         TabIndex        =   15
         Top             =   450
         Width           =   945
      End
      Begin VB.TextBox txtSH 
         Height          =   300
         Index           =   10
         Left            =   -74040
         MaxLength       =   7
         TabIndex        =   14
         Top             =   450
         Width           =   945
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
         Height          =   3120
         Left            =   -74820
         TabIndex        =   48
         Top             =   1470
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   5503
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "已支援次數:         (含相關案)"
         Height          =   180
         Index           =   10
         Left            =   4656
         TabIndex        =   51
         Top             =   2430
         Width           =   2220
      End
      Begin MSForms.TextBox txtSH09 
         Height          =   870
         Left            =   1410
         TabIndex        =   11
         Top             =   2730
         Width           =   5805
         VariousPropertyBits=   -1466941413
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "10239;1535"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label4 
         Height          =   255
         Left            =   3990
         TabIndex        =   47
         Top             =   4110
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
         Left            =   510
         TabIndex        =   46
         Top             =   4110
         Width           =   3300
         VariousPropertyBits=   27
         Caption         =   "lblFM2-Create :"
         Size            =   "5821;450"
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
      Begin MSForms.Label lblSAName 
         Height          =   255
         Left            =   2430
         TabIndex        =   45
         Top             =   1313
         Width           =   1875
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3307;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label5 
         Height          =   255
         Left            =   -68970
         TabIndex        =   44
         Top             =   1140
         Width           =   1815
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3201;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "支援人員:"
         Height          =   180
         Index           =   9
         Left            =   -70920
         TabIndex        =   42
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   8
         Left            =   -74850
         TabIndex        =   40
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收文號:"
         Height          =   180
         Index           =   7
         Left            =   510
         TabIndex        =   39
         Top             =   2040
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
         Caption         =   "支援日期:"
         Height          =   180
         Index           =   55
         Left            =   510
         TabIndex        =   35
         Top             =   630
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   154
         Left            =   510
         TabIndex        =   34
         Top             =   2430
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "支援人員:"
         Height          =   180
         Index           =   155
         Left            =   510
         TabIndex        =   33
         Top             =   990
         Width           =   765
      End
      Begin VB.Line Line1 
         X1              =   1560
         X2              =   3690
         Y1              =   2490
         Y2              =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員:"
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   32
         Top             =   1350
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "支援時數:"
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   31
         Top             =   1710
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註:"
         Height          =   180
         Index           =   2
         Left            =   510
         TabIndex        =   30
         Top             =   2760
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "折算基數:"
         Height          =   180
         Index           =   3
         Left            =   3810
         TabIndex        =   29
         Top             =   1710
         Width           =   765
      End
      Begin VB.Label lblCaseCnt 
         AutoSize        =   -1  'True
         Caption         =   "lblCaseCnt"
         Height          =   180
         Left            =   4650
         TabIndex        =   28
         Top             =   1710
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(EX : 123.5) "
         Height          =   180
         Left            =   2430
         TabIndex        =   27
         Top             =   1710
         Width           =   930
      End
      Begin VB.Line Line3 
         X1              =   -73260
         X2              =   -72690
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "支援人員部門:"
         Height          =   180
         Index           =   5
         Left            =   -74850
         TabIndex        =   37
         Top             =   840
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
         Caption         =   "支援日期:"
         Height          =   180
         Index           =   4
         Left            =   -74850
         TabIndex        =   36
         Top             =   510
         Width           =   765
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
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
Attribute VB_Name = "frm090623"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Lydia 2022/01/03 改成Form2.0 ; grdList改字型=新細明體-ExtB(MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉)、Label3、Label4、Label5、lblSAName、txtSH(9)改為txtSH09; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

'Dim SH(1 To 12) As String
'edit by nickc 2006/01/18
'Dim SH(1 To 18) As String
'Modidified by Lydia 2016/01/25
'Dim SH(1 To 19) As String
Dim sH(1 To 21) As String 'Modified by Morgan 2022/7/1 20->21
Dim strRsStart1 As String, strRsStart2 As String, strRsStart3 As String, strRsStart4 As String, strRsEnd1 As String, strRsEnd2 As String, strRsEnd3 As String, strRsEnd4 As String
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
'add by nickc 2005/07/11
Dim StrSQLa As String
'add by nickc 2005/10/25
Dim Seek_Now_Cp09 As String
'add by nickc 2006/12/29   紀錄 mail 資料，在 trans 後發
Dim skMail() As SeekMails
Dim intTemp As Boolean
Dim adoPrint As ADODB.Recordset 'Added by Morgan 2012/7/9
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
        'Modify By Sindy 2011/1/31 個人進入時, 只可修改個人輸入的資料, 且不可改支援人員
        'If Me.Check1.Value = vbChecked Then
        'Modify By Sindy 2014/7/31
        'If Me.Check1.Value = vbChecked Or (txtSH(0) <> "" And txtSH(1) <> strUserNum) Then
        If Me.Check1.Value = vbChecked Or (txtSH(0) <> "" And Trim(Left(Combo1.Text, 6)) <> strUserNum) Then
            Me.txtSH(0).Enabled = False
            'Me.txtSH(1).Enabled = False
            Combo1.Enabled = False
            Me.txtSH(2).Enabled = False
            Me.txtSH(3).Enabled = False
            Me.txtSH(4).Enabled = False
            Me.txtSH(5).Enabled = False
            Me.txtSH(6).Enabled = False
            Me.txtSH(7).Enabled = False
            Me.txtSH(8).Enabled = False
            Me.txtSH09.Enabled = False
            Me.txtSH(14).Enabled = False
            'add by nickc 2005/10/25
            cmdSelCp09.Enabled = False
        Else
            Me.txtSH(0).Enabled = True
            'Me.txtSH(1).Enabled = True
            Combo1.Enabled = True
            Me.txtSH(2).Enabled = True
            Me.txtSH(3).Enabled = False
            Me.txtSH(4).Enabled = True
            Me.txtSH(5).Enabled = True
            Me.txtSH(6).Enabled = True
            Me.txtSH(7).Enabled = True
            Me.txtSH(8).Enabled = True
            Me.txtSH09.Enabled = True
            Me.txtSH(14).Enabled = True
            'add by nickc 2005/10/25
            cmdSelCp09.Enabled = True
        End If
    End If
End Sub
'Added by Morgan 2022/7/1
Private Sub Check2_Click()
   If Check2.Value = vbChecked Then
      If Check5.Value = vbChecked Then
         Check5.Value = vbUnchecked
      End If
   End If
End Sub

'Added by Morgan 2022/7/1
Private Sub Check5_Click()
   If Check5.Value = vbChecked Then
      If Check2.Value = vbChecked Then
         Check2.Value = vbUnchecked
      End If
   End If
End Sub

Private Sub cmdQuery_Click(Index As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
    
    'add by nickc 2005/07/06 加入判斷，若頁籤在 基本資料，就不管
    If SSTab1.Tab = 0 Then Exit Sub
    
    '2009/12/17 modify by sonia
    'If Me.txtSH(10).Text = "" Then
    '    MsgBox "請輸入支援起日!!!", vbExclamation + vbOKOnly
    '2014/2/19 modify by sonia 加支援人員條件
    If Me.txtSH(10).Text = "" And txtSystem.Text = "" And txtCode(0).Text = "" And txtSH(15).Text = "" Then
        MsgBox "請輸入支援日期範圍或本所案號或支援人員條件!!!", vbExclamation + vbOKOnly
    '2009/12/17 end
        Me.txtSH(10).SetFocus
        Exit Sub
    End If
    If Me.txtSH(10).Text <> "" Then
       If CheckIsTaiwanDate(Me.txtSH(10).Text) = False Then
          Me.txtSH(10).SetFocus
          txtSH_GotFocus 10
          Exit Sub
       End If
    End If
    If Me.txtSH(10).Text <> "" And Me.txtSH(11).Text = "" Then
        MsgBox "請輸入支援迄日!!!", vbExclamation + vbOKOnly
        Me.txtSH(11).SetFocus
        Exit Sub
    End If
    If Me.txtSH(11).Text <> "" Then
       If CheckIsTaiwanDate(Me.txtSH(11).Text) = False Then
          Me.txtSH(11).SetFocus
          Exit Sub
       End If
    End If
    If Val(Me.txtSH(10).Text) > Val(Me.txtSH(11).Text) Then
        MsgBox "支援日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
        Me.txtSH(10).SetFocus
        txtSH_GotFocus 10
        Exit Sub
    End If
    If Me.txtSH(12).Text <> "" And Me.txtSH(13).Text <> "" Then
        If Me.txtSH(12).Text > Me.txtSH(13).Text Then
            MsgBox "支援人員部門範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.txtSH(12).SetFocus
            txtSH_GotFocus 12
            Exit Sub
        End If
    End If
    
    'Add By Sindy 2009/10/28
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
    '2009/10/28 End
    
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
            Screen.MousePointer = vbDefault
            ShowPrintOk
        Else
            ShowNoData
        End If
    End If
End Sub

Private Sub cmdSelCp09_Click()
If Trim(txtSH(5)) <> "" And Trim(txtSH(6)) <> "" Then
   Load frm090623_1
   frm090623_1.Hide
   frm090623_1.oCP01 = txtSH(5).Text
   frm090623_1.oCP02 = txtSH(6).Text
   frm090623_1.oCP03 = IIf(Trim(txtSH(7).Text) = "", "0", txtSH(7).Text)
   frm090623_1.oCP04 = IIf(Trim(txtSH(8).Text) = "", "00", txtSH(8).Text)
   If frm090623_1.Process = False Then ShowNoData: Exit Sub
   frm090623_1.Show vbModal
   Unload frm090623_1
Else
   MsgBox "請先輸入本所案號！", vbExclamation, "警告！"
   If Me.txtSH(5).Enabled = True Then Me.txtSH(5).SetFocus
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
      'edit by nickc 2008/03/28 若是協理已經核可的，協理說不要控制，因為他要修改，但是支援人員已經離職了
      'If Me.lblSupName.Caption = "" Then
      If strEmp = "" And Check1.Value = vbUnchecked Then
         MsgBox "支援人員輸入錯誤!!!", vbExclamation + vbOKOnly
         Cancel = True
      End If
   End If
End Sub
'2014/7/31 END

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'Add By Sindy 2021/2/5 按enter鍵維持換行功能而不是存檔功能
   If KeyCode = vbKeyReturn Then
      Exit Sub
   End If
   '2021/2/5 END
   
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
 
'add by nickc 2006/12/29   紀錄 mail 資料，在 trans 後發
ReDim skMail(0) As SeekMails
   
   m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
   MoveFormToCenter Me
   
   SSTab1.Tab = 0 '設定顯示在第一頁
   Call SetCombo1 'Add By Sindy 2014/7/31
   Label5.Caption = "" 'Added by Lydia 2022/01/03
   
    '取得使用者執行各項功能的權限
'    m_bInsert = IsUserHasRightOfFunction("frm090623", strAdd, False)
'    m_bUpdate = IsUserHasRightOfFunction("frm090623", strEdit, False)
'    m_bDelete = IsUserHasRightOfFunction("frm090623", strDel, False)
'    m_bQuery = IsUserHasRightOfFunction("frm090623", strFind, False)
'    Me.cmdQuery(1).Enabled = IsUserHasRightOfFunction("frm090623", strPrint, False)
    '由個人進入
    If ProState = "1" Then
        m_bInsert = IsUserHasRightOfFunction("frm090623P", strAdd, False)
        m_bUpdate = IsUserHasRightOfFunction("frm090623P", strEdit, False)
        m_bDelete = IsUserHasRightOfFunction("frm090623P", strDel, False)
        m_bQuery = IsUserHasRightOfFunction("frm090623P", strFind, False)
        Me.cmdQuery(1).Enabled = IsUserHasRightOfFunction("frm090623P", strPrint, False)
        'Add By Cheng 2003/12/15
        Me.Check1.Enabled = False
        'End
        'add by nickc 2006/01/18
        Check2.Enabled = False
        'Added by Lydia 2016/01/25
        Check4.Enabled = False
        'Added by Morgan 2022/7/1
        Check5.Enabled = False
    '由管理進入
    Else
        m_bInsert = IsUserHasRightOfFunction("frm090623M", strAdd, False)
        m_bUpdate = IsUserHasRightOfFunction("frm090623M", strEdit, False)
        m_bDelete = IsUserHasRightOfFunction("frm090623M", strDel, False)
        m_bQuery = IsUserHasRightOfFunction("frm090623M", strFind, False)
        Me.cmdQuery(1).Enabled = IsUserHasRightOfFunction("frm090623M", strPrint, False)
        'Add By Cheng 2003/12/15
        Me.Check1.Enabled = True
        'End
        'add by nickc 2006/01/18
        Check2.Enabled = True
        'Added by Lydia 2016/01/25
        Check4.Enabled = True
        'Added by Morgan 2022/7/1
        Check5.Enabled = True
    End If
    strExc(0) = "SELECT * FROM SupportHour WHERE ROWNUM<1"
    intI = 1
    Set rsDefineSize = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   strRsStart1 = Empty: strRsStart2 = Empty: strRsStart3 = Empty: strRsStart4 = Empty
   strRsEnd1 = Empty: strRsEnd2 = Empty: strRsEnd3 = Empty: strRsEnd4 = Empty
   strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour Order By SH01, SH02, SH03, SH04 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ReadRstMsg(intI, strExc(0), True)
   If intI = 1 Then
      RsTemp.MoveFirst
      strRsStart1 = "" & RsTemp.Fields(0).Value
      strRsStart2 = "" & RsTemp.Fields(1).Value
      strRsStart3 = "" & RsTemp.Fields(2).Value
      strRsStart4 = "" & RsTemp.Fields(3).Value
      RsTemp.MoveLast
      strRsEnd1 = "" & RsTemp.Fields(0).Value
      strRsEnd2 = "" & RsTemp.Fields(1).Value
      strRsEnd3 = "" & RsTemp.Fields(2).Value
      strRsEnd4 = "" & RsTemp.Fields(3).Value
      'Modified by Morgan 2019/3/27 改預設在最後一筆--柄佑
      'RsAction 0
      RsAction 3
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
   'add by nickc 2005/07/11
   TxtLock 3
'   If Not IsEmptyText(m_CurrTS(1)) Then
'      ActionEdit = 3
'      CmdSitu True
'   Else
        'Modify By Cheng 2003/08/29
        'Begin
'      CmdSitu False
'      TxtLock 2
'      ActionEdit = 2
        'End
'   End If
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

Private Function ReadSupportHour(ByRef tsTmp() As String) As Boolean
Dim i As Integer, j As Integer, Lbl As LABEL, txt As TextBox, strTmp As String
Dim strTxt(0 To 4) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
    strTxt(1) = tsTmp(1): strTxt(2) = tsTmp(2): strTxt(3) = tsTmp(3): strTxt(4) = tsTmp(4)
    sH(1) = strTxt(1): sH(2) = strTxt(2): sH(3) = strTxt(3): sH(4) = strTxt(4)
    For i = 0 To 10
        If i = 10 Then
            Me.Check1.Value = vbUnchecked
        'Added by Lydia 2022/01/03
        ElseIf i = 9 Then
             Me.txtSH09.Text = ""
        'end 2022/01/03
        ElseIf i <> 1 Then  'Modify By Sindy 2014/7/31 +If i <> 1 Then
            Me.txtSH(i).Text = ""
        End If
    Next i
    Me.txtSH(14).Text = ""
    Me.lblCaseCnt.Caption = ""
    Me.lblSAName.Caption = ""
    'Me.lblSupName.Caption = ""
    Me.Label3.Caption = "Create : "
    Me.Label4.Caption = "Update : "
    Check2.Value = vbUnchecked 'add by nickc 2006/01/18
    Check5.Value = vbUnchecked 'Added by Morgan 2022/7/1
   If sH(1) = "" Then Exit Function
   StrSQLa = "Select * From SupportHour Where SH01=" & sH(1) & " And SH02=" & IIf(sH(2) <> "", "'" & sH(2) & "'", "SH02") & " And SH03=" & IIf(sH(3) <> "", "'" & sH(3) & "'", "SH03") & " And SH04=" & IIf(sH(4) <> "", "'" & sH(4) & "'", "SH04") & " Order By SH01, SH02, SH03, SH04 "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
        sH(1) = "" & rsA.Fields(0).Value
        sH(2) = "" & rsA.Fields(1).Value
        sH(3) = "" & rsA.Fields(2).Value
        sH(4) = "" & rsA.Fields(3).Value
        sH(5) = "" & rsA.Fields(4).Value
        sH(6) = "" & rsA.Fields(5).Value
        sH(7) = "" & rsA.Fields(6).Value
        sH(8) = "" & rsA.Fields(7).Value
        sH(9) = "" & rsA.Fields(8).Value
        sH(10) = "" & rsA.Fields(9).Value
        sH(11) = "" & rsA.Fields(10).Value
        sH(12) = "" & rsA.Fields(11).Value
        'Add By Cheng 2004/03/04
        sH(13) = "" & rsA.Fields(12).Value
        sH(14) = "" & rsA.Fields(13).Value
        sH(15) = "" & rsA.Fields(14).Value
        sH(16) = "" & rsA.Fields(15).Value
        sH(17) = "" & rsA.Fields(16).Value
        sH(18) = "" & rsA.Fields(17).Value
        'End
        'add by nickc 2006/01/18
        sH(19) = "" & rsA.Fields(18).Value
        'Added by Lydia 2016/01/25
        sH(20) = "" & rsA.Fields(19).Value
        sH(21) = "" & rsA.Fields("SH21").Value 'Added by Morgan 2022/7/1
    Else
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Function
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   Me.txtSH(0).Text = ChangeWStringToTString(sH(1))
   Combo1 = sH(2) & " " & GetPrjSalesNM(sH(2)) 'Add By Sindy 2014/7/31
'   Me.txtSH(1).Text = sH(2)
'   Me.lblSupName.Caption = GetStaffName(sH(2), True)
   Me.txtSH(2).Text = sH(3)
   Me.lblSAName.Caption = GetStaffName(sH(3), True)
   Me.txtSH(3).Text = sH(4)
   Me.txtSH(4).Text = sH(5)
   Me.txtSH(5).Text = sH(6)
   Me.txtSH(6).Text = sH(7)
   Me.txtSH(7).Text = sH(8)
   Me.txtSH(8).Text = sH(9)
   Me.txtSH09.Text = sH(10)
    'Modify By Cheng 2003/12/15
'    Me.lblCaseCnt.Caption = IIf(SH(4) <> "", Format(Val(SH(4)) / 4, "0.00"), "")
    'Modify By Cheng 2004/03/01
'    Me.lblCaseCnt.Caption = IIf(Me.txtSH(4).Text <> "", Format(Val(Me.txtSH(4).Text) / 4, "0.00"), "")

'Removed by Morgan 2014/4/28 txtSH_Change 事件會算,此處可免
'    Select Case Me.txtSH(5).Text
'    Case "CFP"
'        Me.lblCaseCnt.Caption = IIf(Me.txtSH(4).Text <> "", Format(Val(Me.txtSH(4).Text) / 6, "0.00"), "")
'    Case Else
'        Me.lblCaseCnt.Caption = IIf(Me.txtSH(4).Text <> "", Format(Val(Me.txtSH(4).Text) / 4, "0.00"), "")
'    End Select
'end 2014/4/29
    'End
    Me.Check1.Value = IIf(sH(11) <> "", vbChecked, vbUnchecked)
    Me.txtSH(14).Text = sH(12)
    Call Check1_Click 'Add By Sindy 2011/1/31
    
    'add by nickc 2005/10/25
    Me.txtSH(0).Tag = Me.txtSH(0).Text
    'Me.txtSH(1).Tag = Me.txtSH(1).Text
    Me.Combo1.Tag = Me.Combo1.Text
    Me.txtSH(2).Tag = Me.txtSH(2).Text
    Me.txtSH(3).Tag = Me.txtSH(3).Text
    'Add By Cheng 2004/03/04
    Me.txtSH(4).Tag = Me.txtSH(4).Text
    Me.txtSH(5).Tag = Me.txtSH(5).Text
    Me.txtSH(6).Tag = Me.txtSH(6).Text
    Me.txtSH(7).Tag = Me.txtSH(7).Text
    Me.txtSH(8).Tag = Me.txtSH(8).Text
    Me.txtSH09.Tag = Me.txtSH09.Text
    Me.Check1.Tag = Me.Check1.Value
    Me.txtSH(14).Tag = Me.txtSH(14).Text
    If sH(13) <> "" Then
        'Modified by Morgan 2019/8/12 離職也要顯示
        Me.Label3.Caption = Me.Label3.Caption & GetStaffName(sH(13), True)
    End If
    If sH(14) <> "" Then
        Me.Label3.Caption = Me.Label3.Caption & " " & ChangeTStringToTDateString(Val(sH(14)) - 19110000)
    End If
    If sH(15) <> "" Then
        Me.Label3.Caption = Me.Label3.Caption & " " & Format(sH(15), "##:##")
    End If
    If sH(16) <> "" Then
        'Modified by Morgan 2019/8/12 離職也要顯示
        Me.Label4.Caption = Me.Label4.Caption & GetStaffName(sH(16), True)
    End If
    If sH(17) <> "" Then
        Me.Label4.Caption = Me.Label4.Caption & " " & ChangeTStringToTDateString(Val(sH(17)) - 19110000)
    End If
    If sH(18) <> "" Then
        Me.Label4.Caption = Me.Label4.Caption & " " & Format(sH(18), "##:##")
    End If
    'End
    'add by nickc 2006/01/18
    Me.Check2.Value = IIf(sH(19) <> "", vbChecked, vbUnchecked)
    Check2.Tag = Check2.Value
    'Added by Lydia 2016/01/25
    Me.Check4.Value = IIf(sH(20) <> "", vbChecked, vbUnchecked)
    Check4.Tag = Check4.Value
    'Added by Morgan 2022/7/1
    Me.Check5.Value = IIf(sH(21) <> "", vbChecked, vbUnchecked)
    Check5.Tag = Check5.Value
    SetSupCount 'Added by Morgan 2024/12/23
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm090623 = Nothing
End Sub

Private Sub RsSitu(ByVal Situ As Integer)
Dim i As Integer, St1 As String, St2 As String
Dim TBmk As Variant
Dim StrSQLa As String
Dim bolP4SH As Boolean '是否已扣支援點數
Dim strReceiptNo As String '已收款收據號 Added by Morgan 2011/11/21
Dim bolInTrans As Boolean, bolClearRefCase As Boolean 'Added by Morgan 2014/7/30

 '911106 nick
 On Error GoTo CheckingErr
 
 Static TmpSH(4) As String
   Select Case Situ
      Case 0 '按下新增add
        TmpSH(1) = ChangeTStringToWString(Me.txtSH(0).Text)
        'TmpSH(2) = Me.txtSH(1).Text
        TmpSH(2) = Trim(Left(Me.Combo1.Text, 6))
        TmpSH(3) = Me.txtSH(2).Text
        TmpSH(4) = Me.txtSH(3).Text
        Me.lblCaseCnt.Caption = ""
        Me.lblSAName.Caption = ""
        Me.lblCaseCnt.Caption = ""
        Me.Label3.Caption = "Create : "
        Me.Label4.Caption = "Update : "
        CmdSitu False
        TxtLock 0
        ActionEdit = 0
        If Me.txtSH(0).Enabled = True Then Me.txtSH(0).SetFocus
        txtSH_GotFocus 0
        'Add By Sindy 2014/7/31
        Combo1.ListIndex = 0
        'Combo1.Locked = True
        '2014/7/31 END
'        Me.txtSH(1).Text = strUserNum
'        Me.txtSH(1).Locked = True 'Add By Sindy 2011/1/31
'        Me.lblSupName.Caption = GetStaffName(Me.txtSH(1).Text)
        'add by nickc 2005/10/25 紀錄收文號
        Seek_Now_Cp09 = ""
        Call Check1_Click 'Add By Sindy 2011/1/31
        
      Case 1 '按下修改modi
         CmdSitu False
         TxtLock 1
         ActionEdit = 1
        TmpSH(1) = ChangeTStringToWString(Me.txtSH(0).Text)
        'TmpSH(2) = Me.txtSH(1).Text
        TmpSH(2) = Trim(Left(Me.Combo1.Text, 6))
        TmpSH(3) = Me.txtSH(2).Text
        TmpSH(4) = Me.txtSH(3).Text
        'add by nickc 2005/10/25 紀錄收文號
        Seek_Now_Cp09 = txtSH(14).Text
      Case 2 '按下刪除delete
         'Add By Sindy 2011/1/31
         '若從個人進入, 若已核可的資料不可刪除
         '個人進入時, 只可刪除個人輸入的資料
         'Modify By Sindy 2014/7/31
         'If ProState = "1" And (Me.Check1.Value = vbChecked Or (txtSH(0) <> "" And txtSH(1) <> strUserNum)) Then
         If ProState = "1" And (Me.Check1.Value = vbChecked Or (txtSH(0) <> "" And Trim(Left(Combo1.Text, 6)) <> strUserNum)) Then
         '2014/7/31 END
         '2011/1/31 End
         Else
             If Me.txtSH(0).Text = "" Then
                 MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
                 Exit Sub
             End If
             If DelMsg Then
                 'Modify By Sindy 2014/7/31
                 'StrSQLa = "Delete From SupportHour Where SH01=" & ChangeTStringToWString(Me.txtSH(0).Text) & " And SH02='" & Me.txtSH(1).Text & "' And SH03='" & Me.txtSH(2).Text & "' And SH04='" & Me.txtSH(3).Text & "' "
                 StrSQLa = "Delete From SupportHour Where SH01=" & ChangeTStringToWString(Me.txtSH(0).Text) & " And SH02='" & Trim(Left(Combo1.Text, 6)) & "' And SH03='" & Me.txtSH(2).Text & "' And SH04='" & Me.txtSH(3).Text & "' "
                 '2014/7/31 END
                 cnnConnection.Execute StrSQLa
                 strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour WHERE SH01||SH02||SH03||SH04>='" & sH(1) & sH(2) & sH(3) & sH(4) & "' Order By SH01, SH02, SH03, SH04 "
                  intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                 If intI = 1 Then
                    strExc(1) = "" & RsTemp.Fields(0).Value
                    strExc(2) = "" & RsTemp.Fields(1).Value
                    strExc(3) = "" & RsTemp.Fields(2).Value
                    strExc(4) = "" & RsTemp.Fields(3).Value
                    ReadSupportHour strExc
                 Else
                     strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour WHERE SH01||SH02||SH03||SH04<='" & sH(1) & sH(2) & sH(3) & sH(4) & "' Order By SH01 Desc , SH02 Desc, SH03 Desc, SH04 Desc "
                      intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strExc(1) = "" & RsTemp.Fields(0).Value
                        strExc(2) = "" & RsTemp.Fields(1).Value
                        strExc(3) = "" & RsTemp.Fields(2).Value
                        strExc(4) = "" & RsTemp.Fields(3).Value
                        ReadSupportHour strExc
                     Else
                        RsAction 0
                     End If
                 End If
                 strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour Order By SH01, SH02, SH03, SH04 "
                 intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ReadRstMsg(intI, strExc(0), True)
                 If intI = 1 Then
                      RsTemp.MoveFirst
                    strRsStart1 = "" & RsTemp.Fields(0).Value
                    strRsStart2 = "" & RsTemp.Fields(1).Value
                    strRsStart3 = "" & RsTemp.Fields(2).Value
                    strRsStart4 = "" & RsTemp.Fields(3).Value
                      RsTemp.MoveLast
                    strRsEnd1 = "" & RsTemp.Fields(0).Value
                    strRsEnd2 = "" & RsTemp.Fields(1).Value
                    strRsEnd3 = "" & RsTemp.Fields(2).Value
                    strRsEnd4 = "" & RsTemp.Fields(3).Value
                 End If
             End If
         End If
      Case 3 'update
         If ActionEdit = 0 Then '在新增狀態按Enter鍵
            'Modified by Lydia 2025/10/08
            'If Not GetData Then Exit Sub
            If Not GetData Then GoTo EXITSUB
            '重新檢查欄位有效性
            'Modified by Lydia 2025/10/08
            'If TxtValidate = False Then Exit Sub
            If TxtValidate = False Then GoTo EXITSUB
            If Me.txtSH(5).Text = "" Or Me.txtSH(6).Text = "" Then
                Me.txtSH(5).Text = "": Me.txtSH(6).Text = "": Me.txtSH(7).Text = "": Me.txtSH(8).Text = ""
            End If
            'Modify By Sindy 2014/7/31
            'Me.txtSH(3).Text = GetSerialNo(ChangeTStringToWString(Me.txtSH(0).Text), Me.txtSH(1).Text, Me.txtSH(2).Text)
            Me.txtSH(3).Text = GetSerialNo(ChangeTStringToWString(Me.txtSH(0).Text), Trim(Left(Combo1.Text, 6)), Me.txtSH(2).Text)
            '2014/7/31 END
            'Modify By Cheng 2003/12/15
'            strSQLA = "Insert Into SupportHour Values(" & ChangeTStringToWString(Me.txtSH(0).Text) & ",'" & Me.txtSH(1).Text & "','" & Me.txtSH(2).Text & "','" & Me.txtSH(3).Text & "'," & Val(Me.txtSH(4).Text) & ",'" & Me.txtSH(5).Text & "','" & Me.txtSH(6).Text & "','" & Me.txtSH(7).Text & "','" & Me.txtSH(8).Text & "','" & ChgSQL(Me.txtsh09.Text) & "')"
            'Modify By Cheng 2004/02/23
            '若勾選主管核可則用"V"存入資料庫
'            strSQLA = "Insert Into SupportHour (SH01, SH02, SH03, SH04, SH05, SH06, SH07, SH08, SH09, SH10, SH11, SH12) Values(" & ChangeTStringToWString(Me.txtSH(0).Text) & ",'" & Me.txtSH(1).Text & "','" & Me.txtSH(2).Text & "','" & Me.txtSH(3).Text & "'," & Val(Me.txtSH(4).Text) & ",'" & Me.txtSH(5).Text & "','" & Me.txtSH(6).Text & "','" & Me.txtSH(7).Text & "','" & Me.txtSH(8).Text & "','" & ChgSQL(Me.txtsh09.Text) & "','" & IIf(Me.Check1.Value = vbChecked, "Y", "") & "','" & Me.txtSH(14).Text & "' )"
            'edit by nickc 2006/01/18
            'StrSQLa = "Insert Into SupportHour (SH01, SH02, SH03, SH04, SH05, SH06, SH07, SH08, SH09, SH10, SH11, SH12) Values(" & ChangeTStringToWString(Me.txtSH(0).Text) & ",'" & Me.txtSH(1).Text & "','" & Me.txtSH(2).Text & "','" & Me.txtSH(3).Text & "'," & Val(Me.txtSH(4).Text) & ",'" & Me.txtSH(5).Text & "','" & Me.txtSH(6).Text & "','" & Me.txtSH(7).Text & "','" & Me.txtSH(8).Text & "','" & ChgSQL(Me.txtsh09.Text) & "','" & IIf(Me.Check1.Value = vbChecked, "V", "") & "'," & CNULL(Me.txtSH(14).Text) & " )"
            'Modify By Sindy 2014/7/31
            'StrSQLa = "Insert Into SupportHour (SH01, SH02, SH03, SH04, SH05, SH06, SH07, SH08, SH09, SH10, SH11, SH12,sh19) Values(" & ChangeTStringToWString(Me.txtSH(0).Text) & ",'" & Me.txtSH(1).Text & "','" & Me.txtSH(2).Text & "','" & Me.txtSH(3).Text & "'," & Val(Me.txtSH(4).Text) & ",'" & Me.txtSH(5).Text & "','" & Me.txtSH(6).Text & "','" & Me.txtSH(7).Text & "','" & Me.txtSH(8).Text & "','" & ChgSQL(Me.txtsh09.Text) & "','" & IIf(Me.Check1.Value = vbChecked, "V", "") & "'," & CNULL(Me.txtSH(14).Text) & ",'" & IIf(Me.Check2.Value = vbChecked, "V", "") & "')"
            'Modified by Lydia 2016/01/25 +計算支援
            'StrSQLa = "Insert Into SupportHour (SH01, SH02, SH03, SH04, SH05, SH06, SH07, SH08, SH09, SH10, SH11, SH12,sh19) Values(" & ChangeTStringToWString(Me.txtSH(0).Text) & ",'" & Trim(Left(Combo1.Text, 6)) & "','" & Me.txtSH(2).Text & "','" & Me.txtSH(3).Text & "'," & Val(Me.txtSH(4).Text) & ",'" & Me.txtSH(5).Text & "','" & Me.txtSH(6).Text & "','" & Me.txtSH(7).Text & "','" & Me.txtSH(8).Text & "','" & ChgSQL(Me.txtsh09.Text) & "','" & IIf(Me.Check1.Value = vbChecked, "V", "") & "'," & CNULL(Me.txtSH(14).Text) & ",'" & IIf(Me.Check2.Value = vbChecked, "V", "") & "')"
            StrSQLa = "Insert Into SupportHour (SH01, SH02, SH03, SH04, SH05, SH06, SH07, SH08, SH09, SH10, SH11, SH12,sh19,sh20) Values(" & ChangeTStringToWString(Me.txtSH(0).Text) & ",'" & Trim(Left(Combo1.Text, 6)) & "','" & Me.txtSH(2).Text & "','" & Me.txtSH(3).Text & "'," & Val(Me.txtSH(4).Text) & ",'" & Me.txtSH(5).Text & "','" & Me.txtSH(6).Text & "','" & Me.txtSH(7).Text & "','" & Me.txtSH(8).Text & "','" & ChgSQL(Me.txtSH09.Text) & "','" & IIf(Me.Check1.Value = vbChecked, "V", "") & "'," & CNULL(Me.txtSH(14).Text) & ",'" & IIf(Me.Check2.Value = vbChecked, "V", "") & "','" & IIf(Me.Check4.Value = vbChecked, "V", "") & "')"
            '2014/7/31 END
            'End
            cnnConnection.Execute StrSQLa
            '寄E-Mail
            SendMail "新增"
            'Modify By Sindy 2014/7/31
            'If ChangeTStringToWString(Me.txtSH(0).Text) & Me.txtSH(1).Text & Me.txtSH(2).Text & Me.txtSH(3).Text < strRsStart1 & strRsStart2 & strRsStart3 & strRsStart4 Then
            If ChangeTStringToWString(Me.txtSH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtSH(2).Text & Me.txtSH(3).Text < strRsStart1 & strRsStart2 & strRsStart3 & strRsStart4 Then
            '2014/7/31 END
                strRsStart1 = ChangeTStringToWString(Me.txtSH(0).Text)
                'strRsStart2 = Me.txtSH(1).Text
                strRsStart2 = Trim(Left(Combo1.Text, 6))
                strRsStart3 = Me.txtSH(2).Text
                strRsStart4 = Me.txtSH(3).Text
            End If
            'Modify By Sindy 2014/7/31
            'If ChangeTStringToWString(Me.txtSH(0).Text) & Me.txtSH(1).Text & Me.txtSH(2).Text & Me.txtSH(3).Text > strRsEnd1 & strRsEnd2 & strRsEnd3 & strRsEnd4 Then
            If ChangeTStringToWString(Me.txtSH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtSH(2).Text & Me.txtSH(3).Text > strRsEnd1 & strRsEnd2 & strRsEnd3 & strRsEnd4 Then
            '2014/7/31 END
                strRsEnd1 = ChangeTStringToWString(Me.txtSH(0).Text)
                'strRsEnd2 = Me.txtSH(1).Text
                strRsEnd2 = Trim(Left(Combo1.Text, 6))
                strRsEnd3 = Me.txtSH(2).Text
                strRsEnd4 = Me.txtSH(3).Text
            End If
            strExc(1) = ChangeTStringToWString(Me.txtSH(0).Text)
            'strExc(2) = Me.txtSH(1).Text
            strExc(2) = Trim(Left(Combo1.Text, 6))
            strExc(3) = Me.txtSH(2).Text
            strExc(4) = Me.txtSH(3).Text
            ReadSupportHour strExc
         ElseIf ActionEdit = 1 Then '在修改狀態按Enter鍵
            'Modified by Lydia 2025/10/08
            'If Not GetData Then Exit Sub
            If Not GetData Then GoTo EXITSUB
            '重新檢查欄位有效性
            'Modified by Lydia 2025/10/08
            'If TxtValidate = False Then Exit Sub
            If TxtValidate = False Then GoTo EXITSUB
            'add by nickc 2005/10/25 檢查重複
            'Modify By Sindy 2014/7/31
            'If txtSH(0).Tag <> txtSH(0).Text Or txtSH(1).Tag <> txtSH(1).Text Or txtSH(2).Tag <> txtSH(2).Text Then
            If txtSH(0).Tag <> txtSH(0).Text Or Trim(Left(Combo1.Tag, 6)) <> Trim(Left(Combo1.Text, 6)) Or txtSH(2).Tag <> txtSH(2).Text Then
            '2014/7/31 END
               'Modify By Sindy 2014/7/31
               'strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour where sh01=" & ChangeTStringToWString(txtSH(0).Text) & " and sh02='" & txtSH(1).Text & "' and sh03='" & txtSH(2).Text & "' Order By SH01, SH02, SH03, SH04 "
               strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour where sh01=" & ChangeTStringToWString(txtSH(0).Text) & " and sh02='" & Trim(Left(Combo1.Text, 6)) & "' and sh03='" & txtSH(2).Text & "' Order By SH01, SH02, SH03, SH04 "
               '2014/7/31 END
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ReadRstMsg(intI, strExc(0), True)
               If intI = 1 Then
                  If MsgBox("當天已經有支援該智權人員的紀錄，是否繼續？", vbYesNo + vbQuestion, "警告！") = vbNo Then
                     Exit Sub
                  End If
               End If
            End If
            If Me.txtSH(5).Text = "" Or Me.txtSH(6).Text = "" Then
                Me.txtSH(5).Text = "": Me.txtSH(6).Text = "": Me.txtSH(7).Text = "": Me.txtSH(8).Text = ""
            End If
'            strSQLA = "Update SupportHour Set SH05=" & Val(Me.txtSH(4).Text) & ", SH06='" & Me.txtSH(5).Text & "', SH07='" & Me.txtSH(6).Text & "', SH08='" & Me.txtSH(7).Text & "', SH09='" & Me.txtSH(8).Text & "', SH10='" & ChgSQL(Me.txtsh09.Text) & "' " & _
'                            " Where SH01=" & ChangeTStringToWString(Me.txtSH(0).Text) & " And SH02='" & Me.txtSH(1).Text & "' And SH03='" & Me.txtSH(2).Text & "' And SH04='" & Me.txtSH(3).Text & "' "
'            cnnConnection.Execute strSQLA
'            strSQLA = "Delete From SupportHour Where SH01=" & TmpSH(1) & " And SH02='" & TmpSH(2) & "' And SH03='" & TmpSH(3) & "' And SH04='" & TmpSH(4) & "' "
'            cnnConnection.Execute strSQLA
'            Me.txtSH(3).Text = GetSerialNo(ChangeTStringToWString(Me.txtSH(0).Text), Me.txtSH(1).Text, Me.txtSH(2).Text)
'            strSQLA = "Insert Into SupportHour Values(" & ChangeTStringToWString(Me.txtSH(0).Text) & ",'" & Me.txtSH(1).Text & "','" & Me.txtSH(2).Text & "','" & Me.txtSH(3).Text & "'," & Val(Me.txtSH(4).Text) & ",'" & Me.txtSH(5).Text & "','" & Me.txtSH(6).Text & "','" & Me.txtSH(7).Text & "','" & Me.txtSH(8).Text & "','" & ChgSQL(Me.txtsh09.Text) & "')"
            'Modify By Cheng 2004/02/23
            '若勾選主管核可則用"V"存入資料庫
'            strSQLA = "Insert Into SupportHour (SH01, SH02, SH03, SH04, SH05, SH06, SH07, SH08, SH09, SH10, SH11, SH12) Values(" & ChangeTStringToWString(Me.txtSH(0).Text) & ",'" & Me.txtSH(1).Text & "','" & Me.txtSH(2).Text & "','" & Me.txtSH(3).Text & "'," & Val(Me.txtSH(4).Text) & ",'" & Me.txtSH(5).Text & "','" & Me.txtSH(6).Text & "','" & Me.txtSH(7).Text & "','" & Me.txtSH(8).Text & "','" & ChgSQL(Me.txtsh09.Text) & "','" & IIf(Me.Check1.Value = vbChecked, "Y", "") & "','" & Me.txtSH(14).Text & "')"
'            strSQLA = "Insert Into SupportHour (SH01, SH02, SH03, SH04, SH05, SH06, SH07, SH08, SH09, SH10, SH11, SH12) Values(" & ChangeTStringToWString(Me.txtSH(0).Text) & ",'" & Me.txtSH(1).Text & "','" & Me.txtSH(2).Text & "','" & Me.txtSH(3).Text & "'," & Val(Me.txtSH(4).Text) & ",'" & Me.txtSH(5).Text & "','" & Me.txtSH(6).Text & "','" & Me.txtSH(7).Text & "','" & Me.txtSH(8).Text & "','" & ChgSQL(Me.txtsh09.Text) & "','" & IIf(Me.Check1.Value = vbChecked, "V", "") & "','" & Me.txtSH(14).Text & "')"
            StrSQLa = ""
            'add by nickc 2005/10/25
            If Me.txtSH(0).Text <> Me.txtSH(0).Tag Then
                StrSQLa = StrSQLa & " SH01=" & ChangeTStringToWString(Me.txtSH(0).Text) & ","
            End If
            'Modify By Sindy 2014/7/31
            'If Me.txtSH(1).Text <> Me.txtSH(1).Tag Then
            If Trim(Left(Combo1.Text, 6)) <> Trim(Left(Combo1.Tag, 6)) Then
                'StrSQLa = StrSQLa & " SH02='" & Val(Me.txtSH(1).Text) & "',"
                StrSQLa = StrSQLa & " SH02='" & Trim(Left(Combo1.Text, 6)) & "',"
            End If
            '2014/7/31 END
            If Me.txtSH(2).Text <> Me.txtSH(2).Tag Then
                'Modified by Morgan 2017/6/8 非數字員工號會變0
                'StrSQLa = StrSQLa & " SH03='" & Val(Me.txtSH(2).Text) & "',"
                StrSQLa = StrSQLa & " SH03='" & Me.txtSH(2).Text & "',"
                'end 2017/6/8
            End If
            'add by nickc 2005/10/25
            'Modify By Sindy 2014/7/31
            'If txtSH(0).Tag <> txtSH(0).Text Or txtSH(1).Tag <> txtSH(1).Text Or txtSH(2).Tag <> txtSH(2).Text Then
            If txtSH(0).Tag <> txtSH(0).Text Or Trim(Left(Combo1.Tag, 6)) <> Trim(Left(Combo1.Text, 6)) Or txtSH(2).Tag <> txtSH(2).Text Then
               'StrSQLa = StrSQLa & " sh04='" & GetSerialNo(ChangeTStringToWString(Me.txtSH(0).Text), Me.txtSH(1).Text, Me.txtSH(2).Text) & "',"
               StrSQLa = StrSQLa & " sh04='" & GetSerialNo(ChangeTStringToWString(Me.txtSH(0).Text), Trim(Left(Combo1.Text, 6)), Me.txtSH(2).Text) & "',"
            End If
            '2014/7/31 END
            If Me.txtSH(4).Text <> Me.txtSH(4).Tag Then
                StrSQLa = StrSQLa & " SH05=" & Val(Me.txtSH(4).Text) & ","
            End If
            If Me.txtSH(5).Text <> Me.txtSH(5).Tag Then
                StrSQLa = StrSQLa & " SH06='" & Me.txtSH(5).Text & "',"
            End If
            If Me.txtSH(6).Text <> Me.txtSH(6).Tag Then
                StrSQLa = StrSQLa & " SH07='" & Me.txtSH(6).Text & "',"
            End If
            If Me.txtSH(7).Text <> Me.txtSH(7).Tag Then
                StrSQLa = StrSQLa & " SH08='" & Me.txtSH(7).Text & "',"
            End If
            If Me.txtSH(8).Text <> Me.txtSH(8).Tag Then
                StrSQLa = StrSQLa & " SH09='" & Me.txtSH(8).Text & "',"
            End If
            If Me.txtSH09.Text <> Me.txtSH09.Tag Then
                'Modified by Morgan 2018/12/14
                'StrSQLa = StrSQLa & " SH10='" & Me.txtsh09.Text & "',"
                StrSQLa = StrSQLa & " SH10='" & ChgSQL(Me.txtSH09.Text) & "',"
            End If
            If Me.Check1.Value <> Me.Check1.Tag Then
                StrSQLa = StrSQLa & " SH11='" & IIf(Me.Check1.Value = vbChecked, "V", "") & "',"
            End If
            If Me.txtSH(14).Text <> Me.txtSH(14).Tag Then
                StrSQLa = StrSQLa & " SH12='" & Me.txtSH(14).Text & "',"
            End If
            'add by nickc 2006/01/18
            If Me.Check2.Value <> Check2.Tag Then
                StrSQLa = StrSQLa & " SH19='" & IIf(Me.Check2.Value = vbChecked, "V", "") & "',"
            End If
            'Added by Lydia 2016/01/25
            If Me.Check4.Value <> Check4.Tag Then
                StrSQLa = StrSQLa & " SH20='" & IIf(Me.Check4.Value = vbChecked, "V", "") & "',"
            End If
            
            'Added by Morgan 2022/7/1
            If Me.Check5.Value <> Check5.Tag Then
                StrSQLa = StrSQLa & " SH21='" & IIf(Me.Check5.Value = vbChecked, "V", "") & "',"
            End If
            'end 2022/7/1
            
            If StrSQLa <> "" Then
                StrSQLa = Left(StrSQLa, Len(StrSQLa) - 1)
            Else
                GoTo NoUpdate
            End If
            
            'Added by Morgan 2014/7/30
            cnnConnection.BeginTrans
            bolInTrans = True
            
            '取消扣支援點數的收文號時e-mail通知智權人員同時將該筆所有多國案之支援記錄的總收文號取消
            bolClearRefCase = False
            If Check1.Tag = vbChecked And txtSH(14).Tag <> "" And txtSH(14).Text = "" Then
               strExc(2) = PUB_GetSameCaseSQL(txtSH(14).Tag)
               
               strExc(1) = ""
               strExc(0) = "select a0j13 from acc0j0 where a0j01='" & txtSH(14).Tag & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(1) = "；收據號碼：" & Trim(RsTemp.GetString(, , , " "))
               End If
               
               ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
               skMail(UBound(skMail)).fiSender = strUserNum
               'Modified by Morgan 2024/11/28
               'skMail(UBound(skMail)).fiReceiver = txtSH(2).Text
               skMail(UBound(skMail)).fiReceiver = txtSH(2).Tag
               'end 2024/11/28
               'Added by Morgan 2022/5/24 智權若已離職則改發目前案件管制的智權
               'Removed by Morgan 2022/5/27 改在 pub_sendmail 控制
               'If txtSH(5).Text <> "" Then
               '   If GetStaffName(txtSH(2).Text) = "" Then
               '      skMail(UBound(skMail)).fiReceiver = PUB_GetAKindSalesNo(txtSH(5).Text, txtSH(6).Text, txtSH(7).Text, txtSH(8).Text)
               '   End If
               'End If
               'end 2022/5/27
               'end 2022/5/24
               
               'Modified by Morgan 2024/11/29 比照給財務處的通知,增加客戶及案件資訊--瑞婷
               'skMail(UBound(skMail)).fiContent = "如旨"
               skMail(UBound(skMail)).fiContent = GetCaseInfo(txtSH(14).Tag)
               'end 2024/11/29
               
               'Modified by Morgan 2024/11/28 案號也可能會被拿掉,改抓修改前資料
               'skMail(UBound(skMail)).fiSubject = "本所案號：" & txtSH(5).Text & "-" & txtSH(6).Text & "-" & txtSH(7).Text & "-" & txtSH(8).Text & strExc(1) & "；已取消扣支援點數的記錄！"
               skMail(UBound(skMail)).fiSubject = "本所案號：" & txtSH(5).Tag & "-" & txtSH(6).Tag & "-" & txtSH(7).Tag & "-" & txtSH(8).Tag & strExc(1) & "；已取消扣支援點數的記錄！"
               'end 2024/11/28
               skMail(UBound(skMail)).fiRecriverNo = GetRecNo 'Modified by Morgan 2022/5/27  "" -> GetRecNo
               
               '若相關案收據已收款, 則e-mail給財務處
               strExc(1) = ""
               strExc(0) = "select distinct a1u02 from supporthour,acc1u0 where (sh06||sh07||sh08||sh09) in (" & strExc(2) & ") and sh12 is not null and a1u03(+)=sh12 and nvl(a1u04,0)+nvl(a1u04,0)>0"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(1) = "；收據號碼：" & Trim(RsTemp.GetString(, , , " "))
                  
                  ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
                  skMail(UBound(skMail)).fiSender = strUserNum
                  skMail(UBound(skMail)).fiReceiver = Pub_GetSpecMan("財務處總帳人員")
                  If RsTemp.RecordCount > 1 Then
                     skMail(UBound(skMail)).fiContent = "請逐一檢查主旨內所列收據號碼，若未扣支援點數則不可再扣，若已扣支援點數則請取消！謝謝！"
                  Else
                     skMail(UBound(skMail)).fiContent = "請檢查若未扣支援點數則不可再扣，若已扣支援點數則請取消！謝謝！"
                  End If
                  'Modified by Morgan 2024/11/28 案號可能會被拿掉,改抓修改前資料
                  'skMail(UBound(skMail)).fiSubject = "本所案號：" & txtSH(5).Text & "-" & txtSH(6).Text & "-" & txtSH(7).Text & "-" & txtSH(8).Text & strExc(1) & "；已收款，現已取消扣支援點數的記錄！"
                  skMail(UBound(skMail)).fiSubject = "本所案號：" & txtSH(5).Tag & "-" & txtSH(6).Tag & "-" & txtSH(7).Tag & "-" & txtSH(8).Tag & strExc(1) & "；已收款，現已取消扣支援點數的記錄！"
                  'end 2024/11/28
                  skMail(UBound(skMail)).fiRecriverNo = GetRecNo 'Modified by Morgan 2022/5/27  "" -> GetRecNo
               End If
               
               bolClearRefCase = True
            End If
            'end 2014/7/30
            
            'edit by nickc 2005/10/25
            'StrSQLa = "Update SupportHour Set " & StrSQLa & " Where SH01=" & Val(ChangeTStringToWString(Me.txtSH(0).Text)) & " And SH02='" & Me.txtSH(1).Text & "' And SH03='" & Me.txtSH(2).Text & "' And SH04='" & Me.txtSH(3).Text & "' "
            'Modify By Sindy 2014/7/31
            'StrSQLa = "Update SupportHour Set " & StrSQLa & " Where SH01=" & Val(ChangeTStringToWString(Me.txtSH(0).Tag)) & " And SH02='" & Me.txtSH(1).Tag & "' And SH03='" & Me.txtSH(2).Tag & "' And SH04='" & Me.txtSH(3).Tag & "' "
            StrSQLa = "Update SupportHour Set " & StrSQLa & " Where SH01=" & Val(ChangeTStringToWString(Me.txtSH(0).Tag)) & " And SH02='" & Trim(Left(Combo1.Tag, 6)) & "' And SH03='" & Me.txtSH(2).Tag & "' And SH04='" & Me.txtSH(3).Tag & "' "
            '2014/7/31 END
            'End
            
            Pub_SeekTbLog StrSQLa 'Added by Morgan 2014/7/30
            cnnConnection.Execute StrSQLa
            
            'Added by Morgan 2014/7/30
            If bolClearRefCase = True Then
               '清除所有相關案的支援記錄的收文號
               strExc(0) = "update supporthour set sh12=null where (sh06||sh07||sh08||sh09) in (" & strExc(2) & ") and sh12 is not null"
               cnnConnection.Execute strExc(0), intI
            End If
               
            cnnConnection.CommitTrans
            bolInTrans = False
            'end 2014/7/30
            
            'Add By Cheng 2004/03/18
            '修改也要寄E-Mail(若從管理進入者不發)
            If ProState <> "2" Then SendMail "修改"
            'add by nickc 2005/10/25 核可時，收文號不同，要發 mail
            '2007/11/1 MODIFY BY SONIA 改為核可時有收文號就發MAIL, 不管是否修改收文號
            'If Check1.Value = vbChecked And txtSH(14).Text <> "" And txtSH(14).Tag <> txtSH(14).Text Then
            If Check1.Value = vbChecked And txtSH(14).Text <> "" Then
'edit by nickc 2006/12/29 改在 trans 後發
'               PUB_SendMail strUserNum, txtSH(2).Text, "", "本所案號：" & txtSH(5).Text & "-" & txtSH(6).Text & "-" & txtSH(7).Text & "-" & txtSH(8).Text & "；支援日期：" & ChangeTStringToTDateString(txtSH(0).Text) & "；扣業務點數 5 點", strUserName

               'Add by Morgan 2007/11/29
               '是否有相關案已扣過支援點數
               bolP4SH = PUB_ChkP4SH(txtSH(14).Text)
               '沒扣過點數也沒有其他支援紀錄已核可才發給智權人員
               If bolP4SH = False Then
                  'Modify By Sindy 2014/7/31
                  'If PUB_ChkSH(txtSH(14).Text, txtSH(0).Text, txtSH(1).Text, txtSH(2).Text, txtSH(3).Text) = False Then
                  If PUB_ChkSH(txtSH(14).Text, txtSH(0).Text, Trim(Left(Combo1.Text, 6)), txtSH(2).Text, txtSH(3).Text) = False Then
                  '2014/7/31 END
               'end 2007/11/29
                     ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
                     skMail(UBound(skMail)).fiSender = strUserNum
                     skMail(UBound(skMail)).fiReceiver = txtSH(2).Text
                     'Added by Morgan 2022/5/24 智權若已離職則改發目前案件管制的智權
                     'Removed by Morgan 2022/5/27 改在 pub_sendmail 控制
                     'If txtSH(5) <> "" Then
                     '   If GetStaffName(txtSH(2).Text) = "" Then
                     '      skMail(UBound(skMail)).fiReceiver = PUB_GetAKindSalesNo(txtSH(5).Text, txtSH(6).Text, txtSH(7).Text, txtSH(8).Text)
                     '   End If
                     'End If
                     'end 2022/5/27
                     'end 2022/5/24
                     
                     skMail(UBound(skMail)).fiContent = strUserName
                     'Modified by Morgan 2025/10/13 + 申請人,支援人員
                     'skMail(UBound(skMail)).fiSubject = "本所案號：" & txtSH(5).Text & "-" & txtSH(6).Text & "-" & txtSH(7).Text & "-" & txtSH(8).Text & "；支援日期：" & ChangeTStringToTDateString(txtSH(0).Text) & "；扣業務點數 5 點"
                     strExc(1) = "本所案號：" & txtSH(5).Text & "-" & txtSH(6).Text & "-" & txtSH(7).Text & "-" & txtSH(8).Text & "；支援日期：" & ChangeTStringToTDateString(txtSH(0).Text) & "；扣業務點數 5 點 ("
                     strExc(0) = "select NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) cu04 from patent,customer where pa01='" & txtSH(5) & "' and pa02='" & txtSH(6) & "' and pa03='" & txtSH(7) & "' and pa04='" & txtSH(8) & "' and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
                     If intI = 1 Then
                        strExc(1) = strExc(1) & "申請人：" & RsTemp(0) & "；"
                     End If
                     strExc(1) = strExc(1) & "工程師：" & Mid(Me.Combo1, 7) & ")"
                     skMail(UBound(skMail)).fiSubject = strExc(1)
                     'end 2025/10/13
                     skMail(UBound(skMail)).fiRecriverNo = GetRecNo 'Modified by Morgan 2022/5/27  "" -> GetRecNo
                     
                     'Modify by Morgan 2009/7/9 從判斷式外面移進來,通知財務處也要判斷沒有其他支援紀錄已核可才發,以免重複通知(當有兩件以上相同案已收款未扣點數時會發生)
                     
                     '檢查是否已經收款
                     'Modified by Morgan 2011/11/21 考慮拆收據情形
                     'strExc(0) = "SELECT * FROM caseprogress where cp09='" & txtSH(14).Text & "' "
                     strExc(0) = "SELECT distinct cp73,a1u02 FROM caseprogress,acc1u0 where cp09='" & txtSH(14).Text & "' and a1u03(+)=cp09 and substr(a1u01,1,1)='F'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ReadRstMsg(intI, strExc(0), True)
                     If intI = 1 Then
                        If Val(CheckStr(RsTemp.Fields("cp73"))) > 0 Then
                           If bolP4SH = False Then 'Add by Morgan 2007/11/29 '相關案都沒扣過點數才發給財務處
                              
                              'Modified by Morgan 2011/11/21 考慮拆收據情形
                              'strReceiptNo = "" & RsTemp("cp60")
                              strReceiptNo = RsTemp("a1u02")
                              RsTemp.MoveNext
                              Do While Not RsTemp.EOF
                                 strReceiptNo = strReceiptNo & "," & RsTemp("a1u02")
                                 RsTemp.MoveNext
                              Loop
                              'end 2011/11/21
                              
                              'edit by nickc 2006/12/29 改在 trans 後發
                              '2006/6/19 MODIFY BY SONIA 71006->71005
                              'Modified by Morgan 2013/6/17 71005->71006
                              'PUB_SendMail strUserNum, "71005", "", "本所案號：" & txtSH(5).Text & "-" & txtSH(6).Text & "-" & txtSH(7).Text & "-" & txtSH(8).Text & "；收據號碼：" & CheckStr(rsTemp.Fields("cp60")) & "；已收款，但為專業部支援，請確認是否扣業務點數 5 點！", strUserName
                              ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
                              skMail(UBound(skMail)).fiSender = strUserNum
                              skMail(UBound(skMail)).fiReceiver = Pub_GetSpecMan("財務處總帳人員")
                              skMail(UBound(skMail)).fiContent = strUserName
                              skMail(UBound(skMail)).fiSubject = "本所案號：" & txtSH(5).Text & "-" & txtSH(6).Text & "-" & txtSH(7).Text & "-" & txtSH(8).Text & "；收據號碼：" & strReceiptNo & "；已收款，但為專業部支援，請確認是否扣業務點數 5 點！"
                              skMail(UBound(skMail)).fiRecriverNo = GetRecNo 'Modified by Morgan 2022/5/27  "" -> GetRecNo
                              
                              'Added by Lydia 2016/01/28 支援記錄已收款, 通知財務處之E-MAIL加智權人員的代碼 .簡稱.收據抬頭.國別及案件性質
                              'Modified by Morgan 2024/11/29 改寫共用
                              'intI = 1
                              'strExc(0) = " select cp13,sn01,cp60,a0k04,decode(pa09,'000',cpm03,cpm04) casetype,na03" & _
                              '            " From caseprogress, salesNo, patent, acc0k0, casepropertymap,nation where cp09='" & txtSH(14).Text & "' " & _
                              '            " and cp13=sn02(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp60=a0k01(+) and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)"
                              'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              'If intI = 1 Then
                              '   skMail(UBound(skMail)).fiContent = vbCrLf & RsTemp.Fields("cp13") & " " & RsTemp.Fields("sn01") & "/" & RsTemp.Fields("a0k04") & " " & RsTemp.Fields("na03") & " " & RsTemp.Fields("casetype") & _
                              '                                      vbCrLf & vbCrLf & strUserName
                              'End If
                              skMail(UBound(skMail)).fiContent = GetCaseInfo(txtSH(14).Text)
                              'end 2016/01/28
                           End If
                        End If
                     End If
                  End If
               End If
            End If
            'End
            
NoUpdate:
            strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour Order By SH01, SH02, SH03, SH04 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
                 RsTemp.MoveFirst
               strRsStart1 = "" & RsTemp.Fields(0).Value
               strRsStart2 = "" & RsTemp.Fields(1).Value
               strRsStart3 = "" & RsTemp.Fields(2).Value
               strRsStart4 = "" & RsTemp.Fields(3).Value
                 RsTemp.MoveLast
               strRsEnd1 = "" & RsTemp.Fields(0).Value
               strRsEnd2 = "" & RsTemp.Fields(1).Value
               strRsEnd3 = "" & RsTemp.Fields(2).Value
               strRsEnd4 = "" & RsTemp.Fields(3).Value
            End If
            strExc(1) = ChangeTStringToWString(Me.txtSH(0).Text)
            'strExc(2) = Me.txtSH(1).Text
            strExc(2) = Trim(Left(Combo1.Text, 6))
            strExc(3) = Me.txtSH(2).Text
            strExc(4) = Me.txtSH(3).Text
            ReadSupportHour strExc
         ElseIf ActionEdit = 2 Then '在查詢狀態按下Enter鍵
            If Me.txtSH(0).Text = "" Then
               MsgBox "支援日期不可空白，請重新輸入 !", vbCritical
               If Me.txtSH(0).Enabled = True Then Me.txtSH(0).SetFocus
               txtSH_GotFocus 0
               Exit Sub
            End If
            intI = 1
            'Modify By Sindy 2014/7/31
            'strExc(0) = "SELECT COUNT(*) FROM SupportHour WHERE SH01=" & ChangeTStringToWString(Me.txtSH(0).Text) & " And SH02=" & IIf(Me.txtSH(1).Text <> "", "'" & Me.txtSH(1).Text & "'", "SH02") & " And SH03= " & IIf(Me.txtSH(2).Text <> "", "'" & Me.txtSH(2).Text & "'", "SH03") & " And SH04= " & IIf(Me.txtSH(3).Text <> "", "'" & Me.txtSH(3).Text & "'", "SH04")
            strExc(0) = "SELECT COUNT(*) FROM SupportHour WHERE SH01=" & ChangeTStringToWString(Me.txtSH(0).Text) & " And SH02=" & IIf(Trim(Left(Combo1.Text, 6)) <> "", "'" & Trim(Left(Combo1.Text, 6)) & "'", "SH02") & " And SH03= " & IIf(Me.txtSH(2).Text <> "", "'" & Me.txtSH(2).Text & "'", "SH03") & " And SH04= " & IIf(Me.txtSH(3).Text <> "", "'" & Me.txtSH(3).Text & "'", "SH04")
            '2014/7/31 END
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) = 0 Then
                  MsgBox "查無此支援記錄 !", vbCritical
                    strExc(1) = TmpSH(1)
                    strExc(2) = TmpSH(2)
                    strExc(3) = TmpSH(3)
                    strExc(4) = TmpSH(4)
               Else
                    strExc(1) = ChangeTStringToWString(Me.txtSH(0).Text)
                    'strExc(2) = Me.txtSH(1).Text
                    strExc(2) = Trim(Left(Combo1.Text, 6))
                    strExc(3) = Me.txtSH(2).Text
                    strExc(4) = Me.txtSH(3).Text
               End If
            End If
            ReadSupportHour strExc
         End If
         'add by nickc 2006/12/29 集中發信
         For i = 1 To UBound(skMail)
            PUB_SendMail skMail(i).fiSender, skMail(i).fiReceiver, skMail(i).fiRecriverNo, skMail(i).fiSubject, skMail(i).fiContent
         Next i
         ReDim skMail(0) As SeekMails
         CmdSitu True
         ActionEdit = 3
         TxtLock 3
      Case 4 'cancel
         If ActionEdit <> 2 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
         End If
         CmdSitu True
        If TmpSH(1) = "" Then TmpSH(1) = strRsStart1
        If TmpSH(2) = "" Then TmpSH(2) = strRsStart2
        If TmpSH(3) = "" Then TmpSH(3) = strRsStart3
        If TmpSH(4) = "" Then TmpSH(4) = strRsStart4
        strExc(1) = TmpSH(1)
        strExc(2) = TmpSH(2)
        strExc(3) = TmpSH(3)
        strExc(4) = TmpSH(4)
         ActionEdit = 3
         ReadSupportHour strExc
         TxtLock 3
      Case 5 'query
        TmpSH(1) = ChangeTStringToWString(Me.txtSH(0).Text)
        'TmpSH(2) = Me.txtSH(1).Text
        TmpSH(2) = Trim(Left(Combo1.Text, 6))
        TmpSH(3) = Me.txtSH(2).Text
        TmpSH(4) = Me.txtSH(3).Text
         CmdSitu False
         TxtLock 2
         ActionEdit = 2
         If Me.txtSH(0).Enabled = True Then Me.txtSH(0).SetFocus
         txtSH_GotFocus 0
   End Select
   
   Exit Sub
CheckingErr:
   If bolInTrans = True Then cnnConnection.RollbackTrans 'Added by Morgan 2014/7/30
   MsgBox Err.Description

EXITSUB: 'Added by Lydia 2025/10/08 檢查資料不過，不用多彈一次訊息
End Sub

Private Sub RsAction(ByVal Sty As Integer)
 Dim i As Integer
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case Sty
      Case 0 '第一筆
         strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour WHERE SH01=" & strRsStart1 & " And SH02 ='" & strRsStart2 & "' And SH03= '" & strRsStart3 & "' And SH04= '" & strRsStart4 & "' "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "" & RsTemp.Fields(0).Value
            strExc(2) = "" & RsTemp.Fields(1).Value
            strExc(3) = "" & RsTemp.Fields(2).Value
            strExc(4) = "" & RsTemp.Fields(3).Value
        Else
            strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour WHERE SH01||SH02||SH03||SH04>='" & strRsStart1 & strRsStart2 & strRsStart3 & strRsStart4 & "' Order By SH01, SH02, SH03, SH04 "
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                strExc(1) = "" & RsTemp.Fields(0).Value
                strExc(2) = "" & RsTemp.Fields(1).Value
                strExc(3) = "" & RsTemp.Fields(2).Value
                strExc(4) = "" & RsTemp.Fields(3).Value
                strRsStart1 = strExc(1)
                strRsStart2 = strExc(2)
                strRsStart3 = strExc(3)
                strRsStart4 = strExc(4)
            End If
         End If
      Case 1 '前一筆
         'Modify By Sindy 2014/7/31
         'If ChangeTStringToWString(Me.txtSH(0).Text) & Me.txtSH(1).Text & Me.txtSH(2).Text & Me.txtSH(3).Text = strRsStart1 & strRsStart2 & strRsStart3 & strRsStart4 Then
         If ChangeTStringToWString(Me.txtSH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtSH(2).Text & Me.txtSH(3).Text = strRsStart1 & strRsStart2 & strRsStart3 & strRsStart4 Then
         '2014/7/31 END
            Beep
            Screen.MousePointer = vbDefault
            DataErrorMessage 6
            Exit Sub
         Else
            intI = 1
            'Modify By Sindy 2014/7/31
            'strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour WHERE SH01||SH02||SH03||SH04<'" & ChangeTStringToWString(Me.txtSH(0).Text) & Me.txtSH(1).Text & Me.txtSH(2).Text & Me.txtSH(3).Text & "' Order By SH01 Desc, SH02 Desc, SH03 Desc, SH04 Desc "
            strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour WHERE SH01||SH02||SH03||SH04<'" & ChangeTStringToWString(Me.txtSH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtSH(2).Text & Me.txtSH(3).Text & "' Order By SH01 Desc, SH02 Desc, SH03 Desc, SH04 Desc "
            '2014/7/31 END
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               strExc(1) = "" & RsTemp.Fields(0).Value
                strExc(2) = "" & RsTemp.Fields(1).Value
                strExc(3) = "" & RsTemp.Fields(2).Value
                strExc(4) = "" & RsTemp.Fields(3).Value
            End If
         End If
      Case 2 '後一筆
         'Modify By Sindy 2014/7/31
         'If ChangeTStringToWString(Me.txtSH(0).Text) & Me.txtSH(1).Text & Me.txtSH(2).Text & Me.txtSH(3).Text = strRsEnd1 & strRsEnd2 & strRsEnd3 & strRsEnd4 Then
         If ChangeTStringToWString(Me.txtSH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtSH(2).Text & Me.txtSH(3).Text = strRsEnd1 & strRsEnd2 & strRsEnd3 & strRsEnd4 Then
         '2014/7/31 END
            Beep
            Screen.MousePointer = vbDefault
            DataErrorMessage 7
            Exit Sub
         Else
            'Modify By Sindy 2014/7/31
            'strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour WHERE SH01||SH02||SH03||SH04>'" & ChangeTStringToWString(Me.txtSH(0).Text) & Me.txtSH(1).Text & Me.txtSH(2).Text & Me.txtSH(3).Text & "' Order By SH01, SH02, SH03, SH04 "
            strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour WHERE SH01||SH02||SH03||SH04>'" & ChangeTStringToWString(Me.txtSH(0).Text) & Trim(Left(Combo1.Text, 6)) & Me.txtSH(2).Text & Me.txtSH(3).Text & "' Order By SH01, SH02, SH03, SH04 "
            '2014/7/31 END
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               strExc(1) = "" & RsTemp.Fields(0).Value
                strExc(2) = "" & RsTemp.Fields(1).Value
                strExc(3) = "" & RsTemp.Fields(2).Value
                strExc(4) = "" & RsTemp.Fields(3).Value
            End If
         End If
      Case 3 '最後筆
         strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour WHERE SH01=" & strRsEnd1 & " And SH02='" & strRsEnd2 & "' And SH03='" & strRsEnd3 & "' And SH04='" & strRsEnd4 & "'"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "" & RsTemp.Fields(0).Value
            strExc(2) = "" & RsTemp.Fields(1).Value
            strExc(3) = "" & RsTemp.Fields(2).Value
            strExc(4) = "" & RsTemp.Fields(3).Value
        Else
            strExc(0) = "SELECT SH01, SH02, SH03, SH04 FROM SupportHour WHERE SH01||SH02||SH03||SH04<='" & strRsEnd1 & strRsEnd2 & strRsEnd3 & strRsEnd4 & "' Order By SH01 Desc, SH02 Desc, SH03 Desc, SH04 Desc "
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                strExc(1) = "" & RsTemp.Fields(0).Value
                strExc(2) = "" & RsTemp.Fields(1).Value
                strExc(3) = "" & RsTemp.Fields(2).Value
                strExc(4) = "" & RsTemp.Fields(3).Value
                strRsEnd1 = strExc(1)
                strRsEnd2 = strExc(2)
                strRsEnd3 = strExc(3)
                strRsEnd4 = strExc(4)
            End If
         End If
   End Select
   ReadSupportHour strExc
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
    Me.txtSH(0).Locked = False
    'Modify By Sindy 2014/7/31
    Combo1.Text = ""
    Combo1.Locked = False
    '2014/7/31 END
'    Me.txtSH(1).Locked = False
    Me.txtSH(2).Locked = False
    Me.txtSH(4).Locked = False
    Me.txtSH(5).Locked = False
    Me.txtSH(6).Locked = False
    Me.txtSH(7).Locked = False
    Me.txtSH(8).Locked = False
    Me.txtSH09.Locked = False
    Me.txtSH(14).Locked = False
    Me.txtSH(0).Text = ""
    'Me.txtSH(1).Text = ""
    Me.txtSH(2).Text = ""
    Me.txtSH(3).Text = ""
    Me.txtSH(4).Text = ""
    Me.txtSH(5).Text = ""
    Me.txtSH(6).Text = ""
    Me.txtSH(7).Text = ""
    Me.txtSH(8).Text = ""
    Me.txtSH09.Text = ""
    Me.txtSH(14).Text = ""
    Me.lblCaseCnt.Caption = ""
    Me.lblSAName.Caption = ""
'    Me.lblSupName.Caption = ""
    'Add By Cheng 2003/12/15
    If ProState = "2" Then
        Me.Check1.Enabled = True
        Me.Check1.Value = vbUnchecked
        'add by nickc 2006/01/18
        Check2.Enabled = True
        Check2.Value = vbUnchecked
        'Added by Lydia 2016/01/25
        Check4.Enabled = True
        Check4.Value = vbUnchecked
        'Added by Morgan 2022/7/1
        Check5.Enabled = True
        Check5.Value = vbUnchecked
    Else
        Me.Check1.Enabled = False
        Me.Check1.Value = vbUnchecked
        'add by nickc 2006/01/18
        Check2.Enabled = False
        Check2.Value = vbUnchecked
        'Added by Lydia 2016/01/25
        Check4.Enabled = False
        Check4.Value = vbUnchecked
        'Added by Morgan 2022/7/1
        Check5.Enabled = False
        Check5.Value = vbUnchecked
    End If
    'End
    'add by nickc 2005/07/11
    cmdSelCp09.Enabled = True
Case 1 '修改
    'edit by nickc 2005/10/25
    'Me.txtSH(0).Locked = True
    Me.txtSH(0).Locked = False
    'Me.txtSH(1).Locked = True
    Combo1.Locked = True
    'edit by nickc 2005/10/25
    'Me.txtSH(2).Locked = True
    Me.txtSH(2).Locked = False
    Me.txtSH(4).Locked = False
    Me.txtSH(5).Locked = False
    Me.txtSH(6).Locked = False
    Me.txtSH(7).Locked = False
    Me.txtSH(8).Locked = False
    Me.txtSH09.Locked = False
    Me.txtSH(14).Locked = False
    'Add By Cheng 2003/12/15
    If ProState = "2" Then
        Me.Check1.Enabled = True
      'add by nickc 2005/07/11
      cmdSelCp09.Enabled = True
      'add by nickc 2006/01/18
      Check2.Enabled = True
      'Added by Lydia 2016/01/25
      Check4.Enabled = True
      'Added by Morgan 2022/7/1
      Check5.Enabled = True
    Else
        Me.Check1.Enabled = False
      'add by nickc 2005/07/11
      cmdSelCp09.Enabled = False
      'add by nickc 2006/01/18
      Check2.Enabled = False
      'Added by Lydia 2016/01/25
      Check4.Enabled = False
      'Added by Morgan 2022/7/1
      Check5.Enabled = False
    End If
    'End

Case 2 '查詢
    Me.txtSH(0).Locked = False
    'Modify By Sindy 2014/7/31
    Combo1.Text = ""
    Combo1.Locked = False
    '2014/7/31 END
'    Me.txtSH(1).Locked = False
    Me.txtSH(2).Locked = False
    Me.txtSH(4).Locked = True
    Me.txtSH(5).Locked = True
    Me.txtSH(6).Locked = True
    Me.txtSH(7).Locked = True
    Me.txtSH(8).Locked = True
    Me.txtSH09.Locked = True
    Me.txtSH(14).Locked = True
    Me.txtSH(0).Text = ""
'    Me.txtSH(1).Text = ""
    Me.txtSH(2).Text = ""
    Me.txtSH(3).Text = ""
    Me.txtSH(4).Text = ""
    Me.txtSH(5).Text = ""
    Me.txtSH(6).Text = ""
    Me.txtSH(7).Text = ""
    Me.txtSH(8).Text = ""
    Me.txtSH09.Text = ""
    Me.txtSH(14).Text = ""
    Me.lblCaseCnt.Caption = ""
    Me.lblSAName.Caption = ""
'    Me.lblSupName.Caption = ""
    'Add By Cheng 2003/12/15
    Me.Check1.Enabled = False
    Me.Check1.Value = vbUnchecked
    'End
    'add by nickc 2005/07/11
    cmdSelCp09.Enabled = False
    'add by nickc 2006/01/18
    Check2.Enabled = False
    Check2.Value = vbUnchecked
    'Added by Lydia 2016/01/25
    Check4.Enabled = False
    Check4.Value = vbUnchecked
    'Added by Morgan 2022/7/1
    Check5.Enabled = False
    Check5.Value = vbUnchecked
    
Case 3 '按下取消後的狀態
    Me.txtSH(0).Locked = True
    'Me.txtSH(1).Locked = True
    Me.Combo1.Locked = True
    Me.txtSH(2).Locked = True
    Me.txtSH(4).Locked = True
    Me.txtSH(5).Locked = True
    Me.txtSH(6).Locked = True
    Me.txtSH(7).Locked = True
    Me.txtSH(8).Locked = True
    Me.txtSH09.Locked = True
    Me.txtSH(14).Locked = True
    'Add By Cheng 2003/12/15
    Me.Check1.Enabled = False
    'End
    'add by nickc 2005/07/11
    cmdSelCp09.Enabled = False
    'add by nickc 2006/01/18
    Check2.Enabled = False
    'Added by Lydia 2016/01/26
    Check4.Enabled = False
    'Added by Morgan 2022/7/1
    Check5.Enabled = False
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
'Add by Morgan 2003/12/26
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
        strExc(3) = Me.grdList.TextMatrix(nRow, 4)
        strExc(4) = Me.grdList.TextMatrix(nRow, 6)
        ReadSupportHour strExc
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
On Error Resume Next
    Select Case Me.SSTab1.Tab
    Case 0
        Me.txtSH(0).SetFocus
        txtSH_GotFocus 0
        Me.cmdQuery(0).Default = False
    Case 1
        Me.txtSH(10).SetFocus
        txtSH_GotFocus 12
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
   
   If ActionEdit <> 3 Then Exit Sub 'Add by Morgan 2011/10/19
   SSTab1.TabEnabled(1) = True 'Add by Morgan 2011/10/19
   
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
      'Added by Lydia 2022/05/18
      If m_bQuery Then
          TBar1.Buttons(4).Enabled = True
      Else
          TBar1.Buttons(4).Enabled = False
      End If
      'end 2022/05/18
   End If
   ' Ken 90.07.16 -- End
   
   Exit Sub
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

Private Function CheckRule() As Boolean
Dim i As Integer, bolChk As Boolean, j As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   CheckRule = False
   If Me.txtSH(0).Text = "" Then
      MsgBox "支援日期不可空白 !", vbCritical
      Me.txtSH(0).SetFocus
      txtSH_GotFocus 0
      Exit Function
   End If
   'If Me.txtSH(1).Text = "" Then
   If Trim(Me.Combo1.Text) = "" Then
      MsgBox "支援人員不可空白 !", vbCritical
      Me.Combo1.SetFocus
      'txtSH_GotFocus 1
      Exit Function
   End If
   If Me.txtSH(2).Text = "" Then
      MsgBox "智權人員不可空白 !", vbCritical
      Me.txtSH(2).SetFocus
      txtSH_GotFocus 2
      Exit Function
   End If
   If Me.txtSH(4).Text = "" Then
      MsgBox "支援數時不可空白 !", vbCritical
      Me.txtSH(4).SetFocus
      txtSH_GotFocus 4
      Exit Function
   End If
    If Me.txtSH(5).Text <> "" And Me.txtSH(6).Text <> "" Then
        'Add By Cheng 2003/08/01
        '案號補滿
        If Me.txtSH(7).Text = "" Then Me.txtSH(7).Text = "0"
        If Me.txtSH(8).Text = "" Then Me.txtSH(8).Text = "00"
        'Modify By Cheng 2004/03/01
'        strSQLA = "Select PA01 From Patent Where " & ChgPatent(Me.txtSH(5).Text & Me.txtSH(6).Text & Me.txtSH(7).Text & Me.txtSH(8).Text)
'        strSQLA = strSQLA & " Union Select TM01 From Trademark Where " & ChgTradeMark(Me.txtSH(5).Text & Me.txtSH(6).Text & Me.txtSH(7).Text & Me.txtSH(8).Text)
'        strSQLA = strSQLA & " Union Select LC01 From Lawcase Where " & ChgLawcase(Me.txtSH(5).Text & Me.txtSH(6).Text & Me.txtSH(7).Text & Me.txtSH(8).Text)
'        strSQLA = strSQLA & " Union Select HC01 From Hirecase Where " & ChgHirecase(Me.txtSH(5).Text & Me.txtSH(6).Text & Me.txtSH(7).Text & Me.txtSH(8).Text)
'        strSQLA = strSQLA & " Union Select SP01 From Servicepractice Where " & ChgService(Me.txtSH(5).Text & Me.txtSH(6).Text & Me.txtSH(7).Text & Me.txtSH(8).Text)
        StrSQLa = "Select PA01 From Patent Where PA01='" & Me.txtSH(5).Text & "' And PA02='" & Me.txtSH(6).Text & "' And PA03='" & Me.txtSH(7).Text & "' And PA04='" & Me.txtSH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select TM01 From Trademark Where TM01='" & Me.txtSH(5).Text & "' And TM02='" & Me.txtSH(6).Text & "' And TM03='" & Me.txtSH(7).Text & "' And TM04='" & Me.txtSH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select LC01 From Lawcase Where LC01='" & Me.txtSH(5).Text & "' And LC02='" & Me.txtSH(6).Text & "' And LC03='" & Me.txtSH(7).Text & "' And LC04='" & Me.txtSH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select HC01 From Hirecase Where HC01='" & Me.txtSH(5).Text & "' And HC02='" & Me.txtSH(6).Text & "' And HC03='" & Me.txtSH(7).Text & "' And HC04='" & Me.txtSH(8).Text & "' "
        StrSQLa = StrSQLa & " Union Select SP01 From Servicepractice Where SP01='" & Me.txtSH(5).Text & "' And SP02='" & Me.txtSH(6).Text & "' And SP03='" & Me.txtSH(7).Text & "' And SP04='" & Me.txtSH(8).Text & "' "
        'End
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount <= 0 Then
            MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
            Me.txtSH(5).SetFocus
            txtSH_GotFocus 5
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Function
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
    'Add By Cheng 2003/12/22
    If Me.txtSH(14).Text <> "" Then
        StrSQLa = "Select * From Caseprogress Where CP09='" & Me.txtSH(14).Text & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic
        If rsA.RecordCount <= 0 Then
            MsgBox "無此收文號資料!!!", vbExclamation + vbOKOnly
            Me.txtSH(14).SetFocus
            txtSH_GotFocus 14
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Function
        Else
            If Me.txtSH(5).Text <> "" Or Me.txtSH(6).Text <> "" Then
                If Me.txtSH(5).Text <> "" & rsA.Fields(0).Value Or Me.txtSH(6).Text <> "" & rsA.Fields(1).Value Or Me.txtSH(7).Text <> "" & rsA.Fields(2).Value Or Me.txtSH(8).Text <> "" & rsA.Fields(3).Value Then
                    MsgBox "此收文號對應的本所案號錯誤!!!", vbExclamation + vbOKOnly
                    Me.txtSH(14).SetFocus
                    txtSH_GotFocus 14
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                    Exit Function
                End If
            End If
            'Added by Lydia 2025/10/08 在點選收文號或輸入收文號時，程式檢查不可為沒有收費的收文號。Ex. P-134497
            If ActionEdit = 0 Or ActionEdit = 1 Then
               If Val("" & rsA.Fields("cp16")) = 0 Or "" & rsA.Fields("CP20") = "N" Then
                  MsgBox "不可為沒有收費的收文號!!!", vbExclamation + vbOKOnly
                  Me.txtSH(14).SetFocus
                  txtSH_GotFocus 14
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
                  Exit Function
               End If
            End If
            'end 2025/10/08
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
    sH(1) = ChangeTStringToWString(Me.txtSH(0).Text)
    'sH(2) = Me.txtSH(1).Text
    sH(2) = Trim(Left(Me.Combo1.Text, 6))
    sH(3) = Me.txtSH(2).Text
    sH(5) = Me.txtSH(4).Text
    sH(6) = Me.txtSH(5).Text
    sH(7) = Me.txtSH(6).Text
    sH(8) = Me.txtSH(7).Text
    sH(9) = Me.txtSH(8).Text
    sH(10) = Me.txtSH09.Text
    'Modify By Cheng 2004/02/23
'    SH(11) = IIf(Me.Check1.Value = vbChecked, "Y", "")
    sH(11) = IIf(Me.Check1.Value = vbChecked, "V", "")
    'End
    sH(12) = Me.txtSH(14).Text
    'add by nickc 2006/01/18
    sH(19) = IIf(Me.Check2.Value = vbChecked, "V", "")
    'Added by Lydia 2016/01/25
    sH(20) = IIf(Me.Check4.Value = vbChecked, "V", "")
    GetData = True
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
   For Each objTxt In Me.txtSH
       If objTxt.Enabled = True Then
          Cancel = False
          txtSH_Validate objTxt.Index, Cancel
          If Cancel = True Then Exit Function
       End If
   Next
   
   'add by nickc  2005/10/25
   If Me.txtSH(0).Text = "" Then MsgBox "支援日期不可空白！", vbCritical, "嚴重錯誤！": txtSH(0).SetFocus: Exit Function
   'If Me.txtSH(1).Text = "" Then MsgBox "支援人員不可空白！", vbCritical, "嚴重錯誤！": txtSH(1).SetFocus: Exit Function
   If Trim(Me.Combo1.Text) = "" Then MsgBox "支援人員不可空白！", vbCritical, "嚴重錯誤！": Combo1.SetFocus: Exit Function
   If Me.txtSH(2).Text = "" Then MsgBox "智權人員不可空白！", vbCritical, "嚴重錯誤！": txtSH(2).SetFocus: Exit Function
   
   'Added by Morgan 2019/8/12 單獨檢查( 因若在 Validate 彈訊息且最後駐點是收文號時會離不開 )
   If txtSH(14) <> "" Then
      If GetOurCaseNo(txtSH(14), True) = False Then
         If txtSH(2).Enabled Then txtSH(2).SetFocus
         Exit Function
      End If
   End If
   'end 2019/8/12
   
   'Add By Sindy 2025/6/9
   If DBDATE(Me.txtSH(0).Text) < DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))) Then
      If MsgBox("此筆支援距今已超過一個月，請確認日期是否有誤？" & vbCrLf & vbCrLf & _
                "若日期有誤，請按＜是＞並修改日期。" & vbCrLf & _
                "若日期無誤，請按＜否＞。", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
         If txtSH(0).Enabled Then txtSH(0).SetFocus
         Exit Function
      End If
   End If
   '2025/6/9 End
   
'   'add by nickc 2005/07/11 若是主管核可，收文號必須要有值
'   'edit by nickc 2005/07/22 不是所有支援的都要扣，且智權人員掛協理或郭雅娟的也不管
'   If Check1.Value = 1 Then
'      If Trim(txtSH(14)) = "" Then
'         MsgBox "收文號必須輸入！", vbCritical, "警告！"
'         txtSH(14).SetFocus
'         Exit Function
'      End If
'   End If
   
   'Add By Sindy 2014/7/31
   Cancel = False
   Call Combo1_Validate(Cancel)
   If Cancel = True Then Exit Function
   '檢查是否有增修刪權限 P10.專利處主管
   'Modify By Sindy 2022/9/12 + And CheckUse("frm090623M", strExec) = False
   'Modified by Morgan 2022/9/13 改判斷個人權限才檢查(不要限定P10因還會設個人Ex:99050)
   'If Pub_StrUserSt03 <> "P10" And Pub_StrUserSt03 <> "M51" And CheckUse("frm090623M", strExec) = False Then
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
         If Combo1.Enabled Then Combo1.SetFocus 'Modified by Morgan 2024/11/4 +Enabled判斷
         Exit Function
      End If
   End If
   '2014/7/31 END
   
   'Added by Morgan 2022/7/1
   If Check5.Enabled Then
      If txtSH(14) <> "" And Me.Check5.Value = vbChecked Then
         MsgBox "已有收文號不應勾不收文！", vbExclamation
         Check5.SetFocus
         Exit Function
      End If
   End If
   'end 2022/7/1
   
    'Added by Lydia 2022/01/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If
   
   TxtValidate = True
End Function

Private Sub txtSH_Change(Index As Integer)
    Select Case Index
'Modified by Morgan 2014/3/20 2014/4/1 起支援改每小時折計0.2基數
'    Case 4 '支援時數
'        If Me.txtSH(Index).Text <> "" Then
'            If Me.txtSH(5).Text = "CFP" Then
'                Me.lblCaseCnt.Caption = Format(Val(Me.txtSH(Index).Text) / 8, "0.00")
'            Else
'                Me.lblCaseCnt.Caption = Format(Val(Me.txtSH(Index).Text) / 4, "0.00")
'            End If
'        Else
'            Me.lblCaseCnt.Caption = ""
'        End If
'
'    Case 5 '系統類別
'        If Me.txtSH(5).Text = "CFP" Then
'            If Me.txtSH(4).Text <> "" Then
'                Me.lblCaseCnt.Caption = Format(Val(Me.txtSH(4).Text) / 8, "0.00")
'            Else
'                Me.lblCaseCnt.Caption = ""
'            End If
'        Else
'            If Me.txtSH(4).Text <> "" Then
'                Me.lblCaseCnt.Caption = Format(Val(Me.txtSH(4).Text) / 4, "0.00")
'            Else
'                Me.lblCaseCnt.Caption = ""
'            End If
'        End If
      Case 0, 4, 5
         SetPoint
'end 2014/3/20
    End Select
    
    'Added by Morgan 2024/12/23
    If ActionEdit <> 3 Then
      If Index = 5 Or Index = 6 Or Index = 7 Or Index = 8 Then
        SetSupCount
      End If
   End If
    'end 2024/12/23
End Sub

'Added by Morgan 2014/3/20
'支援時數折算基數
Private Sub SetPoint()
   If Me.txtSH(4).Text <> "" Then
       If Val(DBDATE(txtSH(0))) >= 20140401 Then
         Me.lblCaseCnt.Caption = Format(Val(Me.txtSH(4).Text) * 0.2, "0.00")
       ElseIf Me.txtSH(5).Text = "CFP" Then
         Me.lblCaseCnt.Caption = Format(Val(Me.txtSH(4).Text) / 3, "0.00")
       Else
         Me.lblCaseCnt.Caption = Format(Val(Me.txtSH(4).Text) / 4, "0.00")
       End If
   Else
       Me.lblCaseCnt.Caption = ""
   End If
End Sub

Private Sub txtSH_GotFocus(Index As Integer)
    TextInverse Me.txtSH(Index)
End Sub

Private Sub txtSH_KeyPress(Index As Integer, KeyAscii As Integer)
    'Added by Morgan 2018/12/14
    '除備註欄位外一律不可輸入單引號,否則語法會出錯
    If Index <> 9 Then
       If KeyAscii = 39 Then
         KeyAscii = 0
         Exit Sub
       End If
    End If
    'end 2018/12/14
    
    Select Case Index
    Case 1, 2, 7, 5, 12, 13, 14, 15 '系統類別, 支援人員部門別, 收文號
        KeyAscii = UpperCase(KeyAscii)
    Case 0
        If KeyAscii = 47 Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub txtSH_LostFocus(Index As Integer)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

    Select Case Index
    Case 8 '本所案號
        If Me.txtSH(5).Text <> "" And Me.txtSH(6).Text <> "" Then
            'Add By Cheng 2003/08/01
            '案號補滿
            If Me.txtSH(7).Text = "" Then Me.txtSH(7).Text = "0"
            If Me.txtSH(8).Text = "" Then Me.txtSH(8).Text = "00"
            StrSQLa = "Select PA01 From Patent Where " & ChgPatent(Me.txtSH(5).Text & Me.txtSH(6).Text & Me.txtSH(7).Text & Me.txtSH(8).Text)
            StrSQLa = StrSQLa & " Union Select TM01 From Trademark Where " & ChgTradeMark(Me.txtSH(5).Text & Me.txtSH(6).Text & Me.txtSH(7).Text & Me.txtSH(8).Text)
            StrSQLa = StrSQLa & " Union Select LC01 From Lawcase Where " & ChgLawcase(Me.txtSH(5).Text & Me.txtSH(6).Text & Me.txtSH(7).Text & Me.txtSH(8).Text)
            StrSQLa = StrSQLa & " Union Select HC01 From Hirecase Where " & ChgHirecase(Me.txtSH(5).Text & Me.txtSH(6).Text & Me.txtSH(7).Text & Me.txtSH(8).Text)
            StrSQLa = StrSQLa & " Union Select SP01 From Servicepractice Where " & ChgService(Me.txtSH(5).Text & Me.txtSH(6).Text & Me.txtSH(7).Text & Me.txtSH(8).Text)
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount <= 0 Then
                MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
                Me.txtSH(5).SetFocus
                txtSH_GotFocus 5
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
    Case 11 '支援日期
        If Me.txtSH(10).Text <> "" And Me.txtSH(11).Text <> "" Then
            If Val(Me.txtSH(10).Text) > Val(Me.txtSH(11).Text) Then
                MsgBox "支援日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                Me.txtSH(10).SetFocus
                txtSH_GotFocus 10
                Exit Sub
            End If
        End If
    Case 12 '支援人員部門
        If Me.txtSH(12).Text <> "" And Me.txtSH(13).Text <> "" Then
            If Me.txtSH(12).Text > Me.txtSH(13).Text Then
                MsgBox "支援人員部門範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                Me.txtSH(12).SetFocus
                txtSH_GotFocus 12
                Exit Sub
            End If
        End If
    End Select
End Sub

Private Sub txtSH_Validate(Index As Integer, Cancel As Boolean)
   If Me.txtSH(Index).Text = "" Then
      'Add By Sindy 2015/3/19
      If Index = 2 Then Me.lblSAName.Caption = ""
      If Index = 15 Then Me.Label5.Caption = ""
      '2015/3/19 END
      Exit Sub
   End If
    Select Case Index
    Case 0 '支援日期
        If CheckIsTaiwanDate(Me.txtSH(Index).Text) = False Then
            Cancel = True
'edit by nickc 2006/11/27 王協裡打電話來說，不用一定要工作天，因為林建志在 95/11/25 有去之支援
'        ElseIf ChkWorkDay(ChangeTStringToWString(Me.txtSH(Index).Text)) = False Then
'            MsgBox "輸入的日期非工作天!!!", vbExclamation + vbOKOnly
'            Cancel = True
        End If

'    Case 1 '支援人員
'        Me.lblSupName.Caption = GetStaffName(Me.txtSH(Index).Text)
'        'edit by nickc 2008/03/28 若是協理已經核可的，協理說不要控制，因為他要修改，但是支援人員已經離職了
'        'If Me.lblSupName.Caption = "" Then
'        If Me.lblSupName.Caption = "" And Check1.Value = vbUnchecked Then
'            MsgBox "支援人員輸入錯誤!!!", vbExclamation + vbOKOnly
'            Cancel = True
'        End If
    Case 2 '智權人員
        '2014/1/24 modify by sonia
        'Me.lblSAName.Caption = GetStaffName(Me.txtSH(Index).Text)
        'edit by nickc 2008/03/28 若是協理已經核可的，協理說不要控制，因為他要修改，但是支援人員已經離職了
        'If Me.lblSAName.Caption = "" Then
        'If Me.lblSAName.Caption = "" And Check1.Value = vbUnchecked Then
        '    MsgBox "智權人員輸入錯誤!!!", vbExclamation + vbOKOnly
        '    Cancel = True
        'End If
        'Modify By Sindy 2014/4/18
        'If ClsPDGetStaff(Me.txtSH(Index).Text, Me.lblSAName.Caption) = False And Check1.Value = vbUnchecked Then
        
        'Modify By Sindy 2015/3/19
'        If ClsPDGetStaff(Me.txtSH(Index).Text, strExc(0)) = False And Check1.Value = vbUnchecked Then
'           Cancel = True
'        End If
'        Me.lblSAName.Caption = strExc(0) 'Add By Sindy 2014/4/18
        If ByInputGetST01or02(Me.txtSH(Index).Text, strExc(0), strExc(1)) = False And Check1.Value = vbUnchecked Then
            Cancel = True
            Me.txtSH(Index).SetFocus
        End If
        Me.txtSH(Index).Text = strExc(0)
        Me.lblSAName.Caption = strExc(1)
        '2015/3/19 END
        '2014/1/24 end
    '2010/5/12 ADD BY SONIA
    'Mark by Lydia 2022/01/03 改成txtSH09
    'Case 9
    '    If CheckLengthIsOK(Me.txtSH(Index).Text, 200) = False Then
    '       Cancel = True
    '    End If
    'end 2022/01/03
    '2010/5/12 END
    Case 10, 11 '支援日期區間
        If CheckIsTaiwanDate(Me.txtSH(Index).Text) = False Then
            Cancel = True
        End If
        
    Case 14 '收文號
      'Modified by Morgan 2019/8/12
      'GetOurCaseNo Me.txtSH(14).Text
      If ActionEdit <> 3 Then GetOurCaseNo Me.txtSH(14).Text
      'end 2019/8/12
        
    '2014/2/19 add by sonia
    Case 15 '支援人員(多筆查詢)
         'Modify By Sindy 2015/3/19
'         Me.Label5.Caption = GetStaffName(Me.txtSH(Index).Text)
'         If Me.Label5.Caption = "" And Me.txtSH(Index).Text <> "" Then
'             MsgBox "支援人員條件錯誤!!!", vbExclamation + vbOKOnly
'             Cancel = True
'         End If
         If ByInputGetST01or02(Me.txtSH(Index).Text, strExc(0), strExc(1)) = False Then
            Cancel = True
            Me.txtSH(Index).SetFocus
         End If
         Me.txtSH(Index).Text = strExc(0)
         Me.Label5.Caption = strExc(1)
         '2015/3/19 END
   '2014/2/19 end
    End Select
    If Cancel = True Then txtSH_GotFocus Index
End Sub

Private Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim nRow As Integer
   
    QueryData = False
    InitialGridList
    strSql = ""
    If Me.txtSH(10).Text <> "" Then
        strSql = strSql & " And SH01>=" & DBDATE(Me.txtSH(10).Text) & " "
    End If
    If Me.txtSH(11).Text <> "" Then
        strSql = strSql & " And SH01<=" & DBDATE(Me.txtSH(11).Text) & " "
    End If
    If Me.txtSH(12).Text <> "" Then
        strSql = strSql & " And S1.ST03>='" & ChgSQL(Me.txtSH(12).Text) & "' "
    End If
    If Me.txtSH(13).Text <> "" Then
        strSql = strSql & " And S1.ST03<='" & ChgSQL(Me.txtSH(13).Text) & "' "
    End If
    '2014/2/19 add by sonia 加多筆查詢支援人員條件
    If Me.txtSH(15).Text <> "" Then
        strSql = strSql & " And SH02='" & ChgSQL(Me.txtSH(15).Text) & "' "
    End If
    'add by nickc 2006/01/18
    If Check3.Value = vbChecked Then
        strSql = strSql & " and SH19='V' and ((sh06 is null and sh07 is null) or  sh12 is null ) "
    End If
    
    'Added by Morgan 2022/7/1
    If Check6.Value = vbChecked Then
        strSql = strSql & " and SH21='V' "
    End If
    'end 2022/7/1
    
    'Add By Sindy 2009/10/28
    If Trim(txtSystem.Text) <> "" And Trim(txtCode(0).Text) <> "" Then
        strSql = strSql & " and SH06='" & Trim(txtSystem.Text) & "' and SH07='" & Trim(txtCode(0).Text) & "' "
        If Trim(txtCode(1).Text) <> "" Then
            strSql = strSql & " and SH08='" & Trim(txtCode(1).Text) & "' "
        Else
            strSql = strSql & " and SH08='0' "
        End If
        If Trim(txtCode(2).Text) <> "" Then
            strSql = strSql & " and SH09='" & Trim(txtCode(2).Text) & "' "
        Else
            strSql = strSql & " and SH09='00' "
        End If
    End If
    '2009/10/28 End
    
    'Modify By Cheng 2004/03/01
'    strSQL = "SELECT Decode(SH01, Null, Null, SH01-19110000), SH02, S1.ST02, SH03, S2.ST02, SH04, SH12, Decode(SH06, Null, '', SH06||Decode(SH07, Null, '','-'||SH07||Decode(SH08, Null, '','-'||SH08||Decode(SH09, Null, '','-'||SH09)))), SH05, Decode(SH05, Null, '', Round(SH05/4,2)), SH11, SH10 FROM SupportHour, Staff S1, Staff S2 " & _
'        "WHERE SH02=S1.ST01 AND SH03=S2.ST01 " & strSQL & " Order By 1, 2, 4, 5 "
    'Modified by Morgan 2012/7/6 +所別,改排序為所別,工程師,日期 --王副總
    'Modified by Morgan 2012/7/9 +查詢排序改回來--王副總
    '2014/2/19 MODIFY BY SONIA 改日期格式
    'strSql = "SELECT Decode(SH01, Null, Null, SH01-19110000) Dt, SH02, S1.ST02, SH03, S2.ST02, SH04, SH12, Decode(SH06, Null, '', SH06||Decode(SH07, Null, '','-'||SH07||Decode(SH08, Null, '','-'||SH08||Decode(SH09, Null, '','-'||SH09)))), SH05, Decode(SH05, Null, '', Decode(SH06,'CFP',Round(SH05/8,2),Round(SH05/4,2))), SH11, SH10,S1.ST06 FROM SupportHour, Staff S1, Staff S2 " & _
         " WHERE SH02=S1.ST01 AND SH03=S2.ST01 " & strSql & " Order By 1, 2, 4, 6 "
    'Modified by Morgan 2014/11/3
    'strSql = "SELECT SQLDATET(SH01) DT, SH02, S1.ST02, SH03, S2.ST02, SH04, SH12, Decode(SH06, Null, '', SH06||Decode(SH07, Null, '','-'||SH07||Decode(SH08, Null, '','-'||SH08||Decode(SH09, Null, '','-'||SH09)))), SH05, Decode(SH05, Null, '', Decode(SH06,'CFP',Round(SH05/8,2),Round(SH05/4,2))), SH11, SH10,S1.ST06 FROM SupportHour, Staff S1, Staff S2 " & _
         " WHERE SH02=S1.ST01 AND SH03=S2.ST01 " & strSql & " Order By 1, 2, 4, 6 "
    strSql = "SELECT substr(SQLDATET(SH01),1,10) DT, SH02, S1.ST02, SH03, S2.ST02, SH04, SH12, Decode(SH06, Null, '', SH06||Decode(SH07, Null, '','-'||SH07||Decode(SH08, Null, '','-'||SH08||Decode(SH09, Null, '','-'||SH09)))), SH05, Decode(SH05, Null, '', Decode(sign(SH01-20140400),1,SH05*0.2, Decode(SH06,'CFP',Round(SH05/3,2),Round(SH05/4,2)))), SH11, SH10,S1.ST06 FROM SupportHour, Staff S1, Staff S2 " & _
         " WHERE SH02=S1.ST01 AND SH03=S2.ST01 " & strSql & " Order By 1, 2, 4, 6 "
    'end 2014/11/3
    'End
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    Set adoPrint = rsTmp.Clone 'Added by Morgan 2012/7/9
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
    grdList.Text = "支援日期"
    grdList.ColWidth(1) = 800
    grdList.ColAlignment(1) = flexAlignCenterCenter
    grdList.col = 2
    grdList.Text = "支援人員代號"
    grdList.ColWidth(2) = 0
    grdList.ColAlignment(2) = flexAlignRightCenter
    grdList.col = 3
    grdList.Text = "支援人員"
    grdList.ColWidth(3) = 800
    grdList.ColAlignment(3) = flexAlignCenterCenter
    grdList.col = 4
    grdList.Text = "智權人員代號"
    grdList.ColWidth(4) = 0
    grdList.ColAlignment(4) = flexAlignRightCenter
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
    grdList.Text = "支援時數"
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
      If InStr("序號,支援時數,折算件數", Me.grdList.Text) > 0 Then
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
Private Function GetSerialNo(strSH01 As String, strSH02 As String, strSH03 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

'edit by nickc 2005/09/16
'StrSQLa = "Select * From SupportHour Where SH01=" & strSH01 & " And SH02='" & strSH02 & "' And SH03='" & strSH03 & "' Order By SH04 Desc "
'Modified by Morgan 2012/7/13 沒資料改預設 0
StrSQLa = "Select nvl(max(sh04),0) as sh04 From SupportHour Where SH01=" & strSH01 & " And SH02='" & strSH02 & "'  Order By SH04 Desc "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'edit by nickc 2005/09/16
'If rsA.RecordCount > 0 Then
If Not rsA.EOF And Not rsA.BOF Then
    GetSerialNo = Format(Val(rsA("SH04").Value) + 1, "000")
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
    'Modified by Morgan 2012/7/9 列印排序不同改抓 adoPrint
    'For ii = 1 To Me.grdList.Rows - 1
    With adoPrint
    .Sort = "st06,SH02,Dt,SH03,SH04" 'Added by Morgan 2012/7/9
    .MoveFirst
    Do While Not .EOF
        '支援人員
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        'Printer.Print Me.grdList.TextMatrix(ii, 3)
        Printer.Print "" & .Fields(2).Value
        
        '支援日期
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iPrint
        'Printer.Print Me.grdList.TextMatrix(ii, 1)
        Printer.Print "" & .Fields(0).Value
        
        '智權人員
        Printer.CurrentX = PLeft(2)
        Printer.CurrentY = iPrint
        'Printer.Print Me.grdList.TextMatrix(ii, 5)
        Printer.Print "" & .Fields(4).Value
        
        '收文號
        Printer.CurrentX = PLeft(3)
        Printer.CurrentY = iPrint
        'Printer.Print Me.grdList.TextMatrix(ii, 7)
        Printer.Print "" & .Fields(6).Value
        
        '本所案號
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iPrint
        'Printer.Print Me.grdList.TextMatrix(ii, 8)
        Printer.Print "" & .Fields(7).Value
        
        '支援時數
        Printer.CurrentX = PLeft(5) + Printer.TextWidth("支援時數") - Printer.TextWidth(Format(Me.grdList.TextMatrix(ii, 9), "0.0"))
        Printer.CurrentY = iPrint
        'Printer.Print Format(Me.grdList.TextMatrix(ii, 9), "0.0")
        Printer.Print Format("" & .Fields(8).Value, "0.0")
        
        '折算件數
        Printer.CurrentX = PLeft(6) + Printer.TextWidth("折算件數") - Printer.TextWidth(Format(Me.grdList.TextMatrix(ii, 10), "0.00"))
        Printer.CurrentY = iPrint
        'Printer.Print Format(Me.grdList.TextMatrix(ii, 10), "0.00")
        Printer.Print Format("" & .Fields(9).Value, "0.00")
        
        '主管核可
        Printer.CurrentX = PLeft(7)
        Printer.CurrentY = iPrint
        'Printer.Print Me.grdList.TextMatrix(ii, 11)
        Printer.Print "" & .Fields(10).Value
        
        '備註
        Printer.CurrentX = PLeft(8)
        Printer.CurrentY = iPrint
        'Printer.Print Replace(Me.grdList.TextMatrix(ii, 12), vbCrLf, "")
        Printer.Print Replace("" & .Fields(11).Value, vbCrLf, "")
        
        iPrint = iPrint + 300
        If iPrint > 10000 And ii <> Me.grdList.Rows - 1 Then
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print String(200, "-")
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
        .MoveNext
    'Next ii
    Loop
    End With

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
Printer.Print "支援記錄明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "支援日期：" & Format(ChangeTStringToTDateString(Me.txtSH(10).Text) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Me.txtSH(11).Text)
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
Printer.Print "支援人員部門別：" & Me.txtSH(12).Text & " " & IIf(Me.txtSH(12).Text <> "" Or Me.txtSH(13).Text <> "", "－", "") & " " & Me.txtSH(13).Text
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
Printer.Print "支援人員"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "支援日期"
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
Printer.Print "支援時數"
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

Private Sub SendMail(strMod As String)
'add by nickc 2006/12/29   紀錄 mail 資料，在 trans 後發
ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
skMail(UBound(skMail)).fiSender = strUserNum

    If strUserNum = "71011" Or strUserNum = "67002" Then
        'edit by nickc 2006/12/29 改在 trans 後發
        'frm880005.txtEmail(0).Text = "68001"
        'modify by sonia 2014/9/9 改68001為94007
        skMail(UBound(skMail)).fiReceiver = "94007"
    Else
        Select Case Left(GetStaffDepartment(strUserNum), 2)
        Case "P1"
            'edit by nickc 2006/12/29 改在 trans 後發
            'frm880005.txtEmail(0).Text = "71011"
            'Added by Lydia 2023/04/24 修改王副總退休之相關控制
            If strSrvDate(1) >= "20230511" Then
                skMail(UBound(skMail)).fiReceiver = "99050"
            ElseIf strSrvDate(1) >= "20230501" Then
                skMail(UBound(skMail)).fiReceiver = "71011;99050"
            Else
            'end 2023/04/24
               skMail(UBound(skMail)).fiReceiver = "71011"
            End If 'Added by Lydia 2023/04/24
        Case "P2"
            'edit by nickc 2006/12/29 改在 trans 後發
            'frm880005.txtEmail(0).Text = "67002"
            'skMail(UBound(skMail)).fiReceiver = "67002"  'cancel by sonia 2020/5/5
        Case Else
            Exit Sub
        End Select
    End If
    '若使用者為北所人員, 則E-Mail後面不加@taie.com.tw
    If PUB_GetST06(strUserNum) = "1" Then
        '無動作
    '若使用者非北所人員, 則E-Mail後面加@taie.com.tw
    Else
        'edit by nickc 2006/12/29 改在 trans 後發
        'frm880005.txtEmail(0).Text = frm880005.txtEmail(0).Text & "@taie.com.tw"
        skMail(UBound(skMail)).fiReceiver = skMail(UBound(skMail)).fiReceiver & "@taie.com.tw"
    End If
    ''edit by nickc 2006/12/29 改在 trans 後發
    'frm880005.txtEmail(1).Text = "<<支援記錄>>" & strMod & "記錄通知"
    'frm880005.txtEmail(2).Text = "支援日期：" & ChangeTStringToTDateString(Me.txtSH(0).Text) & vbCrLf & _
                                                "支援人員：" & Me.txtSH(1).Text & " " & Me.lblSupName.Caption & vbCrLf & _
                                                "智權人員：" & Me.txtSH(2).Text & " " & Me.lblSAName.Caption & vbCrLf & _
                                                "支援時數：" & Me.txtSH(4).Text & vbCrLf & _
                                                "折算件數：" & Me.lblCaseCnt.Caption & vbCrLf & _
                                                "本所案號：" & Me.txtSH(5).Text & "-" & Me.txtSH(6).Text & "-" & Me.txtSH(7).Text & "-" & Me.txtSH(8).Text & vbCrLf & _
                                                "備　　註：" & Me.txtsh09.Text & vbCrLf & _
                                                "是否主管核可：" & IIf(Me.Check1.Value = vbChecked, "是", "否") & vbCrLf & vbCrLf & _
                                                strMod & "資料人員：" & strUserNum & " " & GetStaffName(strUserNum)
    'frm880005.Form_Activate: DoEvents
    'frm880005.cmdok_Click 0: DoEvents
    'Add By Sindy 2025/6/11
    If DBDATE(Me.txtSH(0).Text) < DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(strSrvDate(1)))) Then
      skMail(UBound(skMail)).fiSubject = "[異常!日期超過一個月] <<支援記錄>>" & strMod & "記錄通知"
    Else
    '2025/6/11 END
      skMail(UBound(skMail)).fiSubject = "<<支援記錄>>" & strMod & "記錄通知"
    End If
    skMail(UBound(skMail)).fiContent = "支援日期：" & ChangeTStringToTDateString(Me.txtSH(0).Text) & vbCrLf & _
                                                "支援人員：" & Combo1.Text & vbCrLf & _
                                                "智權人員：" & Me.txtSH(2).Text & " " & Me.lblSAName.Caption & vbCrLf & _
                                                "支援時數：" & Me.txtSH(4).Text & vbCrLf & _
                                                "折算件數：" & Me.lblCaseCnt.Caption & vbCrLf & _
                                                "本所案號：" & Me.txtSH(5).Text & "-" & Me.txtSH(6).Text & "-" & Me.txtSH(7).Text & "-" & Me.txtSH(8).Text & vbCrLf & _
                                                "備　　註：" & Me.txtSH09.Text & vbCrLf & _
                                                "是否主管核可：" & IIf(Me.Check1.Value = vbChecked, "是", "否") & vbCrLf & vbCrLf & _
                                                strMod & "資料人員：" & strUserNum & " " & GetStaffName(strUserNum)
   skMail(UBound(skMail)).fiRecriverNo = GetRecNo 'Modified by Morgan 2022/5/27  "" -> GetRecNo
End Sub

'Modified by Morgan 2019/8/12 支援智權與收文智權不同時提醒
Private Function GetOurCaseNo(strCP09, Optional bolCheck As Boolean = False) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select * From Caseprogress,staff Where CP09='" & strCP09 & "' and st01(+)=cp13"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic
If rsA.RecordCount > 0 Then
    Me.txtSH(5).Text = "" & rsA("CP01").Value
    Me.txtSH(6).Text = "" & rsA("CP02").Value
    Me.txtSH(7).Text = "" & rsA("CP03").Value
    Me.txtSH(8).Text = "" & rsA("CP04").Value
    
    'Added by Morgan 2019/8/12
    If txtSH(2) = "" Then
      txtSH(2) = rsA.Fields("cp13")
      lblSAName = rsA.Fields("st02")
      GetOurCaseNo = True
    ElseIf bolCheck = True Then
      If txtSH(2) <> rsA.Fields("cp13") Then
         'Modified by Morgan 2022/5/24 MsgBox->UniMsgBox
         If UniMsgBox("收文號智權人員為【" & rsA.Fields("st02") & "】，與支援紀錄智權人員【" & lblSAName & "】不同！是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
            GetOurCaseNo = True
         End If
      Else
         GetOurCaseNo = True
      End If
    End If
    'end 2019/8/12
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

'Added by Lydia 2022/01/03
Private Sub txtSH09_GotFocus()
    TextInverse txtSH09
End Sub

'Added by Morgan 2022/5/27
Private Function GetRecNo() As String
   If txtSH(14) <> "" Then
      GetRecNo = txtSH(14)
      
   ElseIf txtSH(5) <> "" Then
      GetRecNo = PUB_GetLastABKindCP09(txtSH(5), txtSH(6), txtSH(7), txtSH(8))
      
   'Added by Morgan 2024/11/29 可能會連本所號也拿掉
   ElseIf txtSH(14).Tag <> "" Then
      GetRecNo = txtSH(14).Tag
      
   End If
End Function

'Added by Morgan 2024/11/28
'收據相關資料(智權,抬頭,國家,性質)
Private Function GetCaseInfo(pCP09 As String) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   intQ = 1
   stSQL = " select a0k20,sn01,cp60,a0k04,getcp10desc(cp01,cp10,a0j04) casetype,na03" & _
               " From caseprogress,acc0j0,nation, acc0k0, salesNo, casepropertymap" & _
               " where cp09='" & pCP09 & "' and na01(+)=a0j04 and a0j01(+)=cp09 and a0k01(+)=a0j13 and sn02(+)=a0k20 and cp01=cpm01(+) and cp10=cpm02(+)"
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      GetCaseInfo = .Fields("a0k20") & " " & .Fields("sn01") & "/" & .Fields("a0k04") & " " & .Fields("na03") & " " & .Fields("casetype")
      End With
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2024/12/23
'設定已支援次數
Private Sub SetSupCount()
   txtCnt = ""
   If txtSH(5) <> "" And Len(txtSH(6)) = 6 Then
      strExc(0) = "select cp09 from caseprogress where cp01='" & txtSH(5) & "' and cp02='" & txtSH(6) & "' and cp03='" & IIf(Trim(txtSH(7).Text) = "", "0", txtSH(7).Text) & "' and cp04='" & IIf(Trim(txtSH(8).Text) = "", "00", txtSH(8).Text) & "' order by cp09 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
      If intI = 1 Then
         strExc(2) = PUB_GetSameCaseSQL(RsTemp("cp09"))
         strExc(0) = "select count(*) from supporthour where (sh06||sh07||sh08||sh09) in (" & strExc(2) & ") and sh12 is not null and upper(sh11)='V'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
         If intI = 1 Then
            txtCnt = RsTemp(0)
         End If
      End If
   End If
End Sub
