VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010610 
   BorderStyle     =   1  '單線固定
   Caption         =   "行事曆資料維護"
   ClientHeight    =   5760
   ClientLeft      =   420
   ClientTop       =   4420
   ClientWidth     =   8460
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8460
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7425
      Top             =   30
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
            Picture         =   "frm06010610.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010610.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010610.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010610.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010610.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010610.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010610.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010610.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010610.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010610.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010610.frx":1DD8
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
      Width           =   8460
      _ExtentX        =   14923
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   60
      TabIndex        =   31
      Top             =   765
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   8696
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm06010610.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(10)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(13)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblFC(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(11)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblCancel(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblCancel(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(12)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblCancel(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(14)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblFC(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Line1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(9)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(15)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtSC(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtSC(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtSC(5)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtSC(6)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtSC(7)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtSC(8)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtSC(10)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtSC(3)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtSC(9)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtSC(4)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lstUsers(0)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lstSC04"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lstUsers(1)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cboSCT02"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textCUID"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtSC(20)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cmdRemSC04"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmdAddSC04"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Frame2"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Cmb1"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "lstSCT03"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Frame1"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmdCancel"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).ControlCount=   45
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm06010610.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(0)"
      Tab(1).Control(1)=   "Label2(1)"
      Tab(1).Control(2)=   "lblName(2)"
      Tab(1).Control(3)=   "Label2(2)"
      Tab(1).Control(4)=   "lblName(3)"
      Tab(1).Control(5)=   "Line2"
      Tab(1).Control(6)=   "Label2(3)"
      Tab(1).Control(7)=   "Txt1(0)"
      Tab(1).Control(8)=   "Txt1(1)"
      Tab(1).Control(9)=   "Txt1(2)"
      Tab(1).Control(10)=   "Txt1(3)"
      Tab(1).Control(11)=   "Check1"
      Tab(1).Control(12)=   "cmdSearch"
      Tab(1).Control(13)=   "GRD1"
      Tab(1).Control(14)=   "Txt1(4)"
      Tab(1).Control(15)=   "Txt1(5)"
      Tab(1).Control(16)=   "Txt1(6)"
      Tab(1).Control(17)=   "Txt1(7)"
      Tab(1).ControlCount=   18
      Begin VB.TextBox Txt1 
         Height          =   285
         Index           =   7
         Left            =   -72300
         MaxLength       =   2
         TabIndex        =   29
         Top             =   1200
         Width           =   435
      End
      Begin VB.TextBox Txt1 
         Height          =   285
         Index           =   6
         Left            =   -72619
         MaxLength       =   1
         TabIndex        =   28
         Top             =   1200
         Width           =   315
      End
      Begin VB.TextBox Txt1 
         Height          =   285
         Index           =   5
         Left            =   -73422
         MaxLength       =   6
         TabIndex        =   27
         Top             =   1200
         Width           =   800
      End
      Begin VB.TextBox Txt1 
         Height          =   285
         Index           =   4
         Left            =   -73920
         MaxLength       =   4
         TabIndex        =   26
         Top             =   1200
         Width           =   495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   62
         Top             =   1560
         Width           =   8055
         _ExtentX        =   14199
         _ExtentY        =   5733
         _Version        =   393216
         FixedCols       =   0
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
      Begin VB.CommandButton cmdCancel 
         Caption         =   "解除管制"
         Height          =   285
         Left            =   5760
         TabIndex        =   19
         Top             =   4208
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   5760
         TabIndex        =   51
         Top             =   3360
         Width           =   1815
         Begin VB.CommandButton cmdRemove 
            Caption         =   "移除 ->"
            Height          =   285
            Index           =   1
            Left            =   45
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   420
            Width           =   735
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "<- 新增"
            Height          =   285
            Index           =   1
            Left            =   45
            TabIndex        =   17
            Top             =   120
            Width           =   735
         End
         Begin VB.TextBox txtUserNo 
            Height          =   264
            Index           =   1
            Left            =   810
            MaxLength       =   6
            TabIndex        =   16
            Top             =   120
            Width           =   945
         End
         Begin MSForms.Label lblName 
            Height          =   180
            Index           =   1
            Left            =   840
            TabIndex        =   52
            Top             =   480
            Width           =   855
            VariousPropertyBits=   27
            Caption         =   "lblName"
            Size            =   "1508;317"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.ListBox lstSCT03 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         ItemData        =   "frm06010610.frx":212C
         Left            =   4560
         List            =   "frm06010610.frx":2133
         Style           =   1  '項目包含核取方塊
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1965
      End
      Begin VB.ComboBox Cmb1 
         Height          =   260
         Left            =   1080
         TabIndex        =   7
         Text            =   "Cmb1"
         Top             =   1170
         Width           =   5655
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   2280
         TabIndex        =   37
         Top             =   2730
         Width           =   1815
         Begin VB.CommandButton cmdRemove 
            Caption         =   "移除 ->"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   420
            Width           =   735
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "<- 新增"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   14
            Top             =   120
            Width           =   735
         End
         Begin VB.TextBox txtUserNo 
            Height          =   264
            Index           =   0
            Left            =   810
            MaxLength       =   6
            TabIndex        =   13
            Top             =   120
            Width           =   945
         End
         Begin MSForms.Label lblName 
            Height          =   180
            Index           =   0
            Left            =   810
            TabIndex        =   38
            Top             =   450
            Width           =   915
            VariousPropertyBits=   27
            Caption         =   "lblName"
            Size            =   "1614;317"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.CommandButton cmdAddSC04 
         Caption         =   "新增↑"
         Height          =   285
         Left            =   6540
         TabIndex        =   12
         Top             =   2430
         Width           =   735
      End
      Begin VB.CommandButton cmdRemSC04 
         Caption         =   "移除↓"
         Height          =   285
         Left            =   6540
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "查詢(&Q)"
         Height          =   330
         Left            =   -69480
         TabIndex        =   25
         Top             =   870
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "包含已解除資料"
         Height          =   255
         Left            =   -71280
         TabIndex        =   24
         Top             =   915
         Width           =   1695
      End
      Begin VB.TextBox Txt1 
         Height          =   285
         Index           =   3
         Left            =   -73920
         MaxLength       =   6
         TabIndex        =   23
         Top             =   877
         Width           =   800
      End
      Begin VB.TextBox Txt1 
         Height          =   285
         Index           =   2
         Left            =   -70320
         MaxLength       =   6
         TabIndex        =   22
         Top             =   555
         Width           =   800
      End
      Begin VB.TextBox Txt1 
         Height          =   285
         Index           =   1
         Left            =   -72900
         MaxLength       =   7
         TabIndex        =   21
         Top             =   555
         Width           =   900
      End
      Begin VB.TextBox Txt1 
         Height          =   285
         Index           =   0
         Left            =   -73920
         MaxLength       =   7
         TabIndex        =   20
         Top             =   555
         Width           =   900
      End
      Begin MSForms.TextBox txtSC 
         Height          =   276
         Index           =   20
         Left            =   6984
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   504
         Visible         =   0   'False
         Width           =   792
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1799;487"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   90
         TabIndex        =   69
         Top             =   4560
         Width           =   6735
         VariousPropertyBits=   671107099
         Size            =   "11880;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboSCT02 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   2370
         Width           =   2715
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "4789;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   780
         Index           =   1
         Left            =   4590
         TabIndex        =   68
         Top             =   3450
         Width           =   1125
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "1984;1376"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstSC04 
         Height          =   600
         Left            =   1050
         TabIndex        =   67
         Top             =   1740
         Width           =   5445
         VariousPropertyBits=   746586139
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "9604;1058"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   1140
         Index           =   0
         Left            =   1080
         TabIndex        =   66
         Top             =   2820
         Width           =   1125
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "1984;2011"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   270
         Index           =   4
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   795
         VariousPropertyBits=   671105051
         MaxLength       =   300
         Size            =   "1799;487"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   264
         Index           =   9
         Left            =   3840
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   720
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "1799;487"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   264
         Index           =   3
         Left            =   120
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   720
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "1799;487"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   10
         Left            =   4560
         TabIndex        =   6
         Top             =   810
         Width           =   280
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "494;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   8
         Left            =   2880
         TabIndex        =   5
         Top             =   810
         Width           =   420
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "741;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   7
         Left            =   2520
         TabIndex        =   4
         Top             =   810
         Width           =   300
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "529;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   6
         Left            =   1680
         TabIndex        =   3
         Top             =   810
         Width           =   720
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1270;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   2
         Top             =   810
         Width           =   480
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "847;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   0
         Top             =   480
         Width           =   1020
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1799;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSC 
         Height          =   285
         Index           =   2
         Left            =   4560
         TabIndex        =   1
         Top             =   480
         Width           =   525
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "926;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "本所案號："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   65
         Top             =   1245
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "(可複選)"
         Height          =   180
         Index           =   15
         Left            =   6600
         TabIndex        =   64
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "(可由管制分類選取)"
         Height          =   180
         Index           =   9
         Left            =   6600
         TabIndex        =   63
         Top             =   1680
         Width           =   1620
      End
      Begin VB.Line Line2 
         X1              =   -72480
         X2              =   -73200
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         X1              =   1320
         X2              =   3000
         Y1              =   960
         Y2              =   960
      End
      Begin MSForms.Label lblFC 
         Height          =   180
         Index           =   1
         Left            =   1920
         TabIndex        =   61
         Top             =   1515
         Width           =   4815
         VariousPropertyBits=   27
         Caption         =   "lblFC"
         Size            =   "1535;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "時間："
         Height          =   180
         Index           =   14
         Left            =   4320
         TabIndex        =   58
         Top             =   4260
         Width           =   540
      End
      Begin MSForms.Label lblCancel 
         Height          =   180
         Index           =   2
         Left            =   4920
         TabIndex        =   57
         Top             =   4260
         Width           =   675
         VariousPropertyBits=   27
         Caption         =   "lblCan(2)"
         Size            =   "1535;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "日期："
         Height          =   180
         Index           =   12
         Left            =   2400
         TabIndex        =   56
         Top             =   4260
         Width           =   540
      End
      Begin MSForms.Label lblCancel 
         Height          =   180
         Index           =   1
         Left            =   3000
         TabIndex        =   55
         Top             =   4260
         Width           =   870
         VariousPropertyBits=   27
         Caption         =   "lblCancel(1)"
         Size            =   "1535;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCancel 
         Height          =   180
         Index           =   0
         Left            =   1080
         TabIndex        =   54
         Top             =   4260
         Width           =   870
         VariousPropertyBits=   27
         Caption         =   "lblCancel(0)"
         Size            =   "1535;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "可解除人員：(輸入人員亦可解除)"
         Height          =   180
         Index           =   11
         Left            =   4560
         TabIndex        =   53
         Top             =   3192
         Width           =   2640
      End
      Begin VB.Label Label1 
         Caption         =   "管制分類："
         Height          =   180
         Index           =   8
         Left            =   135
         TabIndex        =   50
         Top             =   2460
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "解除人員："
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   49
         Top             =   4260
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "事　　由："
         Height          =   180
         Index           =   6
         Left            =   135
         TabIndex        =   48
         Top             =   1800
         Width           =   900
      End
      Begin MSForms.Label lblFC 
         Height          =   180
         Index           =   0
         Left            =   1080
         TabIndex        =   47
         Top             =   1515
         Width           =   735
         VariousPropertyBits=   27
         Caption         =   "lblFC"
         Size            =   "1535;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "FC代理人："
         Height          =   180
         Index           =   4
         Left            =   135
         TabIndex        =   46
         Top             =   1530
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "週　　期：          (1.單次 2.每週 3.每月 4.每3個月 5. 每年)"
         Height          =   180
         Index           =   3
         Left            =   3600
         TabIndex        =   45
         Top             =   840
         Width           =   4440
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   855
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "管制日期："
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   43
         Top             =   525
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱："
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   42
         Top             =   1170
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "細類："
         Height          =   180
         Index           =   5
         Left            =   3960
         TabIndex        =   41
         Top             =   2460
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "流  水  號：              (自動編號)"
         Height          =   180
         Index           =   13
         Left            =   3600
         TabIndex        =   40
         Top             =   525
         Width           =   2370
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "提醒人員："
         Height          =   180
         Index           =   10
         Left            =   135
         TabIndex        =   39
         Top             =   2880
         Width           =   900
      End
      Begin MSForms.Label lblName 
         Height          =   180
         Index           =   3
         Left            =   -73035
         TabIndex        =   36
         Top             =   945
         Width           =   945
         VariousPropertyBits=   27
         Caption         =   "lblName"
         Size            =   "1667;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "輸入人員："
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   35
         Top             =   922
         Width           =   975
      End
      Begin MSForms.Label lblName 
         Height          =   180
         Index           =   2
         Left            =   -69435
         TabIndex        =   34
         Top             =   600
         Width           =   885
         VariousPropertyBits=   27
         Caption         =   "lblName"
         Size            =   "1561;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "提醒人員："
         Height          =   180
         Index           =   1
         Left            =   -71280
         TabIndex        =   33
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "管制日期："
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   32
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm06010610"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/04/21 改成Form2.0 ; txtSC(index)、lstUsers(index)、lblCancel(index)、lblFC(index)、lblName(index)、textCUID、GRD1改字型=新細明體-ExtB
'Memo by Lydia 2020/01/15 更名為「行事曆資料維護」
'Created by Lydia 2015/12/23 國外部行事曆資料維護
Option Explicit

Dim m_EditMode As Integer '0:瀏覽 1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type

Dim m_FieldList() As FIELDITEM

Dim TF_SC As Integer
Dim strTmp As String
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

'Modified by Lydia 2021/04/21
'Dim oText As TextBox
'Dim oLabel As LABEL
Dim oText As Control, oLabel As Control
Dim idx As Integer
Dim iType As String '適用部門
Dim mSC11 As String '輸入人員
Dim bolUpdate As Boolean '是否能修改
Dim strSCT02 As String '暫存點選分類
Dim strSCT03() As String  '細類編號+說明
Dim bolClose As Boolean '北所銷卷案的處理
'Dim mESeqNo As String '暫存TB編號 'Remove by Lydia 2020/09/14
Dim bSC01 As String, bSC02 As String '新增前的編號
Dim strCase(1 To 4) As String 'Added by Lydia 2023/07/28 本所案號
Dim m_PA177 As String 'Added by Lydia 2023/07/28 FCP專利連結通知

Private Sub cboSCT02_Click()
Dim tmpX As Integer
  tmpX = cboSCT02.ListIndex
  If tmpX > 0 Then
    If strSCT03(tmpX) <> "" Then
       Call SetlstSCT03(tmpX)
    'Added by Lydia 2016/01/21 清除細類
    Else
       lstSCT03.Clear
    End If
    strSCT02 = tmpX
  Else
    lstSCT03.Clear
  End If
End Sub

'新增同仁
Private Sub cmdAdd_Click(Index As Integer)

   If ChkUsersNum(Index) = False Then Exit Sub
   
   AddlstUsers Index
   If Index = 0 Then
      txtSC(3) = ComposeListX(Index)
      '可解除人員：預設為提醒人員1，可修改且至少輸入一名；
      If txtSC(9) = "" And Len(txtSC(3)) <= 6 Then
         txtUserNo(1) = txtSC(3)
         Call cmdAdd_Click(1)
      End If
   ElseIf Index = 1 Then
      txtSC(9) = ComposeListX(Index)
   End If
   txtUserNo(Index).SetFocus
   txtUserNo_GotFocus Index
End Sub

Private Sub cmdCancel_Click()
Dim Sdate As String
Dim SNo  As Integer
Dim m_InputDate As String 'Added by Lydia 2023/07/28 輸入勘誤日期

   If m_EditMode = 1 Or m_EditMode = 2 Then
      MsgBox "解除管制前,請先存檔!", vbCritical
      Exit Sub
   End If
   If lblCancel(1).Caption <> "" Then
       MsgBox "本記錄已解除管制!", vbCritical, "解除管制"
   Else
       'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：行事曆解除期限檢查
       If m_PA177 = "Y" And txtSC(20) <> "" And InStr(txtSC(4), "請程序確認") > 0 And InStr(txtSC(4), "公報刊載日期") > 0 Then
          m_InputDate = PUB_ChkFCPlinkSC(DBDATE(txtSC(1)), txtSC(2))
          If m_InputDate = "" Then
             Exit Sub
          End If
       End If
       'end 2023/07/28
       
On Error GoTo ErrorHand
       strExc(2) = Mid(Right("000000" & ServerTime, 6), 1, 4)
       'Added by Lydia 2016/02/25  +案號顯示
       strExc(3) = IIf(Trim(txtSC(5) & txtSC(6)) <> "", " ，案號: " & txtSC(5) & Val(Trim(txtSC(6))) & IIf(txtSC(7) & txtSC(8) = "000", "", txtSC(7) & txtSC(8)), "")
       cnnConnection.BeginTrans
          'Modified by Lydia 2017/07/18 + chgsql 去除單引號
          If PUB_AddFCPStaffCalendar(IIf(txtSC(10) = "1", "", DBDATE(txtSC(1))), txtSC(10), txtSC(3), ChgSQL(txtSC(4)), txtSC(9), txtSC(10), txtSC(5), txtSC(6), txtSC(7), txtSC(8), Sdate, SNo, mSC11) Then
             'Modified by Lydia 2016/02/25
             'MsgBox "下次行事曆的管制日期: " & ChangeWStringToTString(Sdate) & "　流水號: " & sNO, vbInformation, "解除管制"
             MsgBox "下次行事曆的管制日期: " & ChangeWStringToTString(Sdate) & "　流水號: " & SNo & strExc(3), vbInformation, "解除管制"
          Else
             If txtSC(10) <> "1" Then MsgBox "下次行事曆新增失敗!", vbCritical, "解除管制"
          End If
          strSql = "UPDATE staff_calendar SET sc17='" & strUserNum & "',sc18=" & strSrvDate(1) & ",sc19=" & CNULL(strExc(2), True) & _
                   " where sc01=" & DBDATE(txtSC(1)) & " and sc02=" & CNULL(txtSC(2), True)
          cnnConnection.Execute strSql, intI
          'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：解除行事曆=>當程序確認公報刊載日期後解除行事曆自動收文「通知資訊變更961」,發一封Email給承辦工程師
          If m_InputDate <> "" Then
             strExc(0) = "select c2.cp09 as oldCP09,c2.cp10 as oldCP10,c1.cp09,c1.cp10,c1.cp12,c1.cp13,c1.cp14 from caseprogress c1, caseprogress c2 where c1.cp09='" & txtSC(20) & "' and c1.cp43=c2.cp09(+)"
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
                'Modified by Lydia 2023/12/19 傳入值拿掉CP14 ,改在模組內抓最後一道承辦工程師
                If PUB_GetFCPlinkMC("2A", m_InputDate, strCase, "" & RsTemp.Fields("CP09"), "" & RsTemp.Fields("oldCP10"), "" & RsTemp.Fields("CP10"), "" & RsTemp.Fields("CP12"), "" & RsTemp.Fields("CP13")) = True Then
                End If
                'Added by Lydia 2024/04/10 當程序解除行事曆期限時，系統會彈視窗輸入公告日，請自動將公報刊載日期一併掛在核准那道的承辦期限。 ----請參考frm06010602_3
                strSql = "Update Caseprogress Set cp48='" & DBDATE(m_InputDate) & "' where cp09='" & txtSC(20) & "' and cp158=0 and cp159=0 "
                cnnConnection.Execute strSql
                'end 2024/04/10
             End If
          End If
          'end 2023/07/28
       cnnConnection.CommitTrans
       lblCancel(0) = strUserName
       lblCancel(1) = strSrvDate(2)
       lblCancel(2) = Format(strExc(2), "##:##")
       '解除人非輸入人員時,mail通知輸入人
       'Modified by Lydia 2016/07/18 改成模組判斷
       'If strUserNum <> mSC11 Then
         'Modified by Lydia 2016/02/25 +案號顯示
          'Modified by Lydia 2020/01/15
          'strExc(1) = "國外部行事曆：管制日期: " & txtSC(1) & " 流水號: " & txtSC(2) & strExc(3) & " ， 已被解除管制!"
          strExc(1) = "行事曆：管制日期: " & txtSC(1) & " 流水號: " & txtSC(2) & strExc(3) & " ， 已被解除管制!"
          'Modified by Lydia 2016/03/01 +行事曆內容
          strExc(4) = "本所案號: " & IIf(Trim(txtSC(5) & txtSC(6)) <> "", txtSC(5) & "-" & txtSC(6) & "-" & txtSC(7) & "-" & txtSC(8), "") & vbCrLf
          strExc(4) = strExc(4) & "案件名稱: " & Trim(IIf(Cmb1.Text <> "", Mid(Cmb1.Text, InStr(Cmb1.Text, ":") + 1), "")) & vbCrLf
          strExc(4) = strExc(4) & "事　　由: " & Replace(txtSC(4), vbCrLf, vbCrLf & "　　　　  ") & vbCrLf
       '   PUB_SendMail strUserNum, mSC11, "", strExc(1), strExc(4)
       'End If
       Call PUB_CancelFCPStaffCalendar(strUserNum, mSC11, strExc(1), strExc(4), txtSC(5), txtSC(6), txtSC(7), txtSC(8))
       'end 2016/07/18
       
       '多筆資料-整理
       If txt1(3) = strUserNum Then
          If QueryData(False) = False Then
          End If
       End If
       Call PUB_SendMailCache 'Added by Lydia 2023/08/25
   End If
   
   Exit Sub
   
ErrorHand:
   If Err.Number > 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Sub

'移除同仁
Private Sub cmdRemove_Click(Index As Integer)
   RemovelstUsers Index
   If Index = 0 Then
      txtSC(3) = ComposeListX(Index)
   ElseIf Index = 1 Then
      txtSC(9) = ComposeListX(Index)
   End If
   txtUserNo(Index).SetFocus
End Sub
Private Function ComposeListX(p_index As Integer) As String
   'Modified by Lydia 2021/04/21
   'strExc(1) = ""
   'If lstUsers(p_index).ListCount > 0 Then
   '   strExc(1) = PUB_Num2Id(lstUsers(p_index).ItemData(0))
   '   For intI = 1 To lstUsers(p_index).ListCount - 1
   '      strExc(1) = strExc(1) & "," & PUB_Num2Id(lstUsers(p_index).ItemData(intI))
   '   Next
   'End If
   'ComposeListX = strExc(1)
   ComposeListX = lstUsers(p_index).Tag
   'end 2021/04/21
End Function
Private Function ChkUsersNum(ByVal aID As Integer, Optional ByVal inA As Integer = 0) As Boolean
'inA = 存檔檢查+1
Dim tmpB As Boolean
ChkUsersNum = False
   Select Case aID
      Case 0
          idx = txtSC(3).MaxLength \ 6
          If lstUsers(aID).ListCount >= idx + inA Then
             MsgBox "提醒人員最多" & idx & "人!", vbInformation, "輸入限制"
             txtUserNo(aID).SetFocus
             txtUserNo_GotFocus aID
             Exit Function
          End If
      Case 1
          'Modified by Lydia 2017/06/22 解除人員放到4名(2=>4)
          idx = 4
          txtUserNo_Validate aID, tmpB
          If tmpB = True Then Exit Function
          
          If lstUsers(aID).ListCount >= idx + inA Then
              'Modified by Lydia 2017/06/22 解除人員放到4名
              MsgBox "可解除人員最多4人!", vbInformation, "輸入限制"
              txtUserNo(aID).SetFocus
              txtUserNo_GotFocus aID
              Exit Function
          End If
      Case Else
          Exit Function
   End Select
   
ChkUsersNum = True
End Function
Private Sub cmdAddSC04_Click()
   If AddList(cboSCT02, lstSCT03, lstSC04) = True Then
      txtSC(4) = ComposeList(lstSC04)
      Call GetDefUsers(iType, strSCT02)
      cboSCT02 = ""
      lstSCT03.Clear
   End If
   cboSCT02.SetFocus
End Sub

Private Sub cmdRemSC04_Click()
   If RemoveList(lstSC04) = True Then
      txtSC(4) = ComposeList(lstSC04)
      cboSCT02.SetFocus
   End If
End Sub
'預設提醒人員:畫面上之提醒人員欄若為空白且有輸入本所案號時，再依下述規則預設提醒人員
Private Sub GetDefUsers(ByVal sTyp As String, ByVal CNo As String)
Dim strMid As String 'Added by Lydia 2017/06/20
Dim tmpArr As Variant 'Added by Lydia 2025/05/16

   'Modified by Lydia 2016/01/21 不限制案號
   'If CNo <> "" And txtSC(5) <> "" And txtSC(6) <> "" And Cmb1.ListCount > 0 And txtSC(3) = "" Then
      'strExc(5) = txtSC(5): strExc(6) = txtSC(6): strExc(7) = Right("0" & txtSC(7), 1): strExc(8) = Right("00" & txtSC(8), 2)
   'Modified by Lydia 2016/01/26
   'If CNo <> "" And txtSC(3) = "" Then
   If txtSC(3) = "" Then
      'Memo by Lydia 2023/07/28 strCase(1~4)取掉strExc(5~8)
        '輸入人員為外專程序組時：
        If sTyp = "1" Then
           'Added by Lydia 2016/01/21 無案號時設定為輸入者
           If strCase(1) = "" And strCase(2) = "" Then
              txtUserNo(0) = strUserNum
              Call cmdAdd_Click(0)
              Exit Sub
           End If
           'end 2016/01/21
           Select Case Val(CNo)
              '點選分類1至4項時，提醒人員1預設為該本所案號之FCP承辦業務員(以PUB_GetFCPSalesNo抓取人員)，提醒人員2預設為該本所案號之FCP管制人；
              Case 1, 2, 3, 4
                   txtUserNo(0).Text = PUB_GetFCPSalesNo(strCase(1), strCase(2), strCase(3), strCase(4))
                   If txtUserNo(0).Text <> "" Then
                      Call cmdAdd_Click(0)
                   End If
                   txtUserNo(0).Text = PUB_GetFCPHandler(strCase(1), strCase(2), strCase(3), strCase(4)) 'Modified by Lydia 2016/01/21
                   strMid = txtUserNo(0).Text 'Added by Lydia 2017/06/20
                   If txtUserNo(0).Text <> "" Then
                      Call cmdAdd_Click(0)
                   End If
              '點選分類5時，提醒人員1預設為該本所案號之FCP管制人(以PUB_GetFCPHandler抓取人員)；再以本所案號抓新案翻譯201、檢視中說209、製作中說210的承辦人，若其ST03為”F51”時則提醒人員2預設為系統特殊人員(M)，非該部門時則不預設提醒人員2；
              Case 5
                   txtUserNo(0).Text = PUB_GetFCPHandler(strCase(1), strCase(2), strCase(3), strCase(4))
                   strMid = txtUserNo(0).Text 'Added by Lydia 2017/06/20
                   If txtUserNo(0).Text <> "" Then
                      Call cmdAdd_Click(0)
                   End If
                   'Modified by Lydia 2016/09/21 cp57 is null => CP159=0
                   strExc(0) = "select cp14,st04,st02,st03 from caseprogress,staff where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "' " & _
                               "and st01(+)=cp14 and cp10 in('201','209','210') and CP159=0 order by cp05 desc,cp09 desc"
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                      If Trim("" & RsTemp("st03")) = "F51" Then
                         txtUserNo(0).Text = Pub_GetSpecMan("M") '
                      End If
                      If txtUserNo(0).Text <> "" Then
                          Call cmdAdd_Click(0)
                      End If
                   End If
              '點選分類6至11項時，提醒人員1預設為該本所案號之FCP管制人；'點選分類13-17項時，提醒人員1預設為該本所案號之FCP管制人；
              Case 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17
                   txtUserNo(0).Text = PUB_GetFCPHandler(strCase(1), strCase(2), strCase(3), strCase(4))
                   strMid = txtUserNo(0).Text 'Added by Lydia 2017/06/20
                   If txtUserNo(0).Text <> "" Then
                      Call cmdAdd_Click(0)
                   End If
              '點選分類12時，提醒人員1預設為該本所案號之FCP承辦業務員，提醒人員2預設為該案之最後工程師；
              Case 12 '追蹤會稿結果
                   txtUserNo(0).Text = PUB_GetFCPSalesNo(strCase(1), strCase(2), strCase(3), strCase(4))
                   If txtUserNo(0).Text <> "" Then
                      Call cmdAdd_Click(0)
                   End If
                   'Modified by Lydia 2016/09/21 cp57 is null => CP159=0
                   'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
                   strExc(0) = "select cp14,st04,st02 from caseprogress,staff where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "' " & _
                               "and st01(+)=cp14 and cp14<>'F4102' and st03='F21' and CP159=0 order by cp05 desc,cp09 desc"
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                      If Trim("" & RsTemp("st04")) = "1" Then
                         txtUserNo(0).Text = "" & RsTemp("cp14")
                         If txtUserNo(0).Text <> "" Then
                            Call cmdAdd_Click(0)
                         End If
                      End If
                   End If
                   'Added by Lydia 2017/10/18 追蹤會稿結果:通知人員和解除人員,增加FCP管制人員
                   strExc(9) = PUB_GetFCPHandler(strCase(1), strCase(2), strCase(3), strCase(4))
                   txtUserNo(0).Text = strExc(9)
                   strMid = txtUserNo(0).Text
                   If txtUserNo(0).Text <> "" Then
                      Call cmdAdd_Click(0)
                      txtUserNo(1).Text = strExc(9)
                      Call cmdAdd_Click(1)
                   End If
                   'end 2017/10/18
              Case Else
                  'Added by Lydia 2016/01/26 預設FCP管制人
                   txtUserNo(0).Text = PUB_GetFCPHandler(strCase(1), strCase(2), strCase(3), strCase(4))
                   strMid = txtUserNo(0).Text 'Added by Lydia 2017/06/20
                   If txtUserNo(0).Text <> "" Then
                      Call cmdAdd_Click(0)
                   End If
              End Select
              
              'Added by Lydia 2017/06/20 國外部行事曆資料維護若有輸入案號時，可解除人員自動多掛FCP程序人員的案件第一職代
              If strMid <> "" Then
                 'Modified by Lydia 2025/05/16 改抓全部職代; Ex. Winfrey第一職代Teresa不處理案件
                 'txtUserNo(1).Text = GetABS001_17(strMid)
                 'If txtUserNo(1).Text <> "" Then
                 '   Call cmdAdd_Click(1)
                 'End If
                 Call GetABS001_1(strMid, strExc(1), strExc(2), strExc(3))
                 If strExc(1) & strExc(2) & strExc(3) <> "" Then
                    tmpArr = Split(strExc(1) & IIf(strExc(2) <> "", "," & strExc(2), "") & IIf(strExc(3) <> "", "," & strExc(3), ""), ",")
                    For intI = 0 To UBound(tmpArr)
                       If Trim(tmpArr(intI)) <> "" Then
                          txtUserNo(1).Text = Trim(tmpArr(intI))
                          Call cmdAdd_Click(1)
                       End If
                    Next
                 End If
                 'end 2025/05/16
              End If
              'end 2017/06/20
        ElseIf sTyp = "2" Then
            txtUserNo(0).Text = PUB_GetFCPSalesNo(strCase(1), strCase(2), strCase(3), strCase(4))
            If txtUserNo(0).Text <> "" Then
               Call cmdAdd_Click(0)
            End If
        End If
   End If
End Sub

Private Sub cmdSearch_Click()
    If QueryData(True) = False Then
    End If
End Sub

Private Function QueryData(Optional ByRef bolM As Boolean = True) As Boolean
Dim rsRead As New ADODB.Recordset
Dim strS1 As String, inX As Integer
Dim stSQL As String, strTempName As String
Dim tmpArr As Variant
Dim strUsers As String

QueryData = False

   '管制日期
   If txt1(0) <> "" And txt1(1) <> "" Then
      strS1 = strS1 & " and sc01>=" & DBDATE(txt1(0)) & " and sc01<=" & DBDATE(txt1(1))
   ElseIf txt1(0) <> "" Then
         strS1 = strS1 & " and sc01>=" & DBDATE(txt1(0))
       ElseIf txt1(1) <> "" Then
         strS1 = strS1 & " and sc01<=" & DBDATE(txt1(1))
   End If
   '提醒人員
   If txt1(2) <> "" Then
      strS1 = strS1 & " and instr(sc03," & CNULL(txt1(2)) & ") > 0 "
   End If
   '輸入人員
   If txt1(3) <> "" Then
      strS1 = strS1 & " and sc11=" & CNULL(txt1(3))
   End If
   '本所案號
   For intI = 4 To 7
      If txt1(intI) <> "" Then
         strS1 = strS1 & " and sc" & Format(intI + 1, "00") & "=" & CNULL(txt1(intI))
      End If
   Next intI
   
   If strS1 = "" Then
      MsgBox "查詢條件至少輸入一項!", vbCritical, "查詢錯誤"
      Exit Function
   End If
   'Modified by Lydia 2020/10/29 debug
   'If Check1.Value = False Then
   If Check1.Value = 0 Then
      strS1 = strS1 & " and sc18 is null "
   End If

'Remove by Lydia 2020/09/14 直接用DB的函數
'   strTmp = "select sc01,sc02,sc03 from staff_calendar where 1=1 " & strS1
'   strTmp = strTmp & " order by 1,2,3 "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strTmp)
'On Error GoTo ErrHnd:
'   If intI = 1 Then
'       '提醒人員從員工編號轉姓名
'       Set rsRead = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
'
'       cnnConnection.BeginTrans
'       rsRead.MoveFirst
'       With rsRead
'          Do While Not .EOF
'             tmpArr = Empty: strUsers = ""
'             tmpArr = Split(.Fields("SC03"), ",")
'             For inX = 0 To UBound(tmpArr)
'                 If tmpArr(inX) <> "" Then
'                    'Modified by Lydia 2016/08/16 遇到離職人員不彈訊息
'                    'If ClsPDGetStaff(TmpArr(inX), strTempName) = True Then
'                    strTempName = GetStaffName(tmpArr(inX), True)
'                    If strTempName <> "" Then
'                      strUsers = strUsers & IIf(Len(strUsers) > 0, ",", "") & strTempName
'                    End If
'                 End If
'             Next
'             If strUsers <> "" Then
'                strExc(1) = "update rdatafactory set r004=" & CNULL(strUsers) & " where FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "' and rowseq=" & CNULL(rsRead.AbsolutePosition)
'                cnnConnection.Execute strExc(1), intI
'             End If
'             .MoveNext
'          Loop
'       End With
'       cnnConnection.CommitTrans
'end 2020/09/14

       'Modified by Lydia 2016/06/28 + 4.每3個月
       'Modified by Lydia 2020/09/14 直接用DB的函數
       'strTmp = "select (sc01-19110000) 管制日期,sc02,r004 提醒人員,sc04,decode(sc05||sc06,null,'',sc05||'-'||sc06||decode(sc07||sc08,'000','','-'||sc07||'-'||sc08)) sc0508," & _
                "sc09,decode(sc10,'1','單次','2','每週','3','每月','4','每3個月',sc10) sc10,decode(sc18,null,'','Y') 解除,sc11,(st02) sc11n from staff_calendar,staff,RDataFactory where sc11=st01(+) " & _
                strS1 & " and to_char(sc01)=R001(+) and to_char(sc02)=R002(+) and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "'"
On Error GoTo ErrHnd
       'Modified by Lydia 2020/10/29
       'strTmp = "select (sc01-19110000) 管制日期,sc02,getstaffnamelist(sc03) 提醒人員,sc04, decode(sc05||sc06,null,'',sc05||'-'||sc06||decode(sc07||sc08,'000','','-'||sc07||'-'||sc08)) sc0508,sc09," & _
                     "decode(sc10,'1','單次','2','每週','3','每月','4','每3個月',sc10) sc10,decode(sc18,null,'','y') 解除,sc11,(st02) sc11n from staff_calendar,staff where sc11=st01(+) " & strS1 & " and sc18 is null "
       strTmp = "select (sc01-19110000) 管制日期,sc02,getstaffnamelist(sc03) 提醒人員,sc04, decode(sc05||sc06,null,'',sc05||'-'||sc06||decode(sc07||sc08,'000','','-'||sc07||'-'||sc08)) sc0508,sc09," & _
                     "decode(sc10,'1','單次','2','每週','3','每月','4','每3個月',sc10) sc10,decode(sc18,null,'','Y') 解除,sc11,(st02) sc11n from staff_calendar,staff where sc11=st01(+) " & strS1
       strTmp = strTmp & " order by sc01,sc02 " 'Added by Lydia 2020/09/29
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strTmp)
       'Added by Lydia 2020/09/14
       If intI = 0 Then
            If bolM = True Then MsgBox "查無資料!!"
            GRD1.Clear
            Call SetGrd
       Else
       'end 2020/09/14
            GRD1.FixedCols = 0
            Set GRD1.Recordset = RsTemp
            Call SetGrd(RsTemp.RecordCount + 1)
            GRD1.FixedCols = 2
            QueryData = True
       End If 'Added by Lydia 2020/09/14
'Remove by Lydia 2020/09/14
'   Else
'       If bolM = True Then MsgBox "查無資料!!"
'       GRD1.Clear
'       Call SetGrd
'   End If
'end 2020/09/14

   Exit Function

ErrHnd:
   If Err.Number > 0 Then
      'cnnConnection.RollbackTrans 'Remove by Lydia 2020/09/14
      MsgBox Err.Description
   End If
End Function

Private Sub Form_Initialize()
   strExc(0) = "select * from Staff_Calendar where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   TF_SC = RsTemp.Fields.Count
   ReDim m_FieldList(1 To TF_SC) As FIELDITEM
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   If Pub_StrUserSt03 = "M51" Then
       iType = InputBox("請輸入欲操作的適用部門？" & vbCrLf & "(1:外專程序 2:外專承辦)")
       If iType <> "1" And iType <> "2" Then
          iType = "1"
       End If
   'Modified by Lydia 2016/06/30
   'ElseIf InStr("31,33,34,32", Pub_strUserST05) > 0 Then
   ElseIf Pub_StrUserSt03 = "F22" Then
       iType = "1"
   'Modified by Lydia 2016/06/30
   'ElseIf InStr("35,36,37", Pub_strUserST05) > 0 Then
   Else
       iType = "2"
   End If
   
   SetcboSCT02 (IIf(iType = "0", "1", iType))
 
   textCUID.BackColor = &H8000000F
   ClearField
   InitialField
   m_EditMode = 0
   ShowRecord 9, False
   SetInputEntry
   UpdateToolbarState
   '清除查詢
   For Each oText In txt1
      oText.Text = Empty
   Next
   lblName(2).Caption = ""
   '輸入人員：預設為操作者；
   txt1(3) = strUserNum
   lblName(3).Caption = strUserName
   Check1.Value = False
   If QueryData(False) = False Then
   End If
   
   Me.SSTab1.Tab = 0
   
End Sub
Private Sub SetGrd(Optional ByRef iR As Integer = 2)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("管制日期", "流水號", "提醒人員", "事由", "本所案號", "SC09", "週期", "解除", "SC11", "輸入人員")
   'Modified by Lydia 2016/06/28
   'arrGridHeadWidth = Array(840, 640, 1500, 2000, 1320, 0, 500, 500, 0, 840)
   arrGridHeadWidth = Array(840, 600, 1500, 2000, 1320, 0, 720, 500, 0, 840)
   
   With GRD1
       .Visible = False
       .Cols = UBound(arrGridHeadText) + 1
       .Rows = iR
       For iRow = 0 To .Cols - 1
          .row = 0
          .col = iRow
          .Text = arrGridHeadText(iRow)
          .ColWidth(iRow) = arrGridHeadWidth(iRow)
          .CellAlignment = flexAlignCenterCenter
       Next

       For intI = 1 To iR - 1
         .row = intI
         For iRow = 0 To .Cols - 1
           .col = iRow
           If iRow < 2 Then
              .CellBackColor = QBColor(15)
           End If
           '流水號
           If iRow = 1 Then
              .CellAlignment = flexAlignRightCenter
           '週期，解除
           ElseIf iRow = 6 Or iRow = 7 Then
              .CellAlignment = flexAlignCenterCenter
           End If

         Next iRow
       Next intI


       .Visible = True
   End With
End Sub
'設定管制分類項目+儲存細類項目
Private Sub SetcboSCT02(ByVal aKind As String)
Dim idR As Integer
Dim rsAD As New ADODB.Recordset
Dim strGrp As String 'Added by Lydia 2018/02/21 SCT02分類

    cboSCT02.Clear
    cboSCT02.AddItem "", 0
    If Val(aKind) > 0 Then
       strTmp = "select * from staff_calendar_type where sct01='" & aKind & "' order by 1,2,3 "
       idx = 1
       Set rsAD = ClsLawReadRstMsg(idx, strTmp)
       If idx = 1 Then
          ReDim strSCT03(1 To rsAD.RecordCount)
          idx = 1 'Added by Lydia 2018/02/21
          rsAD.MoveFirst
          Do While Not rsAD.EOF
             'Modified by Lydia 2018/02/21 SCT02分類第一筆不是SCT03=00
             'If Val(rsAD.Fields("SCT03")) = 0 Then
             '   cboSCT02.AddItem Trim(rsAD.Fields("SCT04")), Val(rsAD.Fields("SCT02"))
             If strGrp <> "" & rsAD.Fields("SCT02") Then
                 cboSCT02.AddItem Trim(rsAD.Fields("SCT04")), idx
                 strGrp = "" & rsAD.Fields("SCT02")
                 idx = idx + 1
             'end 2018/02/21
             Else
                idR = rsAD.Fields("SCT02")
                strSCT03(idR) = strSCT03(idR) & Trim(rsAD.Fields("SCT04")) & ","
             End If
             rsAD.MoveNext
          Loop
       Else
          MsgBox "查無管制分類項目!", vbCritical
       End If
    End If
End Sub
'設定管制細類
Private Sub SetlstSCT03(ByVal rID As Integer)
Dim idR As Integer
Dim tmpArr As Variant

    lstSCT03.Clear
    tmpArr = Empty
    If strSCT03(rID) <> "" Then
       tmpArr = Split(strSCT03(rID), ",")
       For idx = 0 To UBound(tmpArr)
           If tmpArr(idx) <> "" Then
              lstSCT03.AddItem Trim(tmpArr(idx)), idx
           End If
       Next
    End If
End Sub
' 開始輸入資料
Private Sub SetInputEntry()
   If Me.Visible = True Then
      txtSC(2).TabStop = False
      Select Case m_EditMode
         Case 1 '新增
            txtSC(1).SetFocus
            txtSC_GotFocus 1
         Case 2 '修改
            txtSC(1).SetFocus
            txtSC_GotFocus 1
         Case 4 '查詢
            txtSC(2).TabStop = True
            txtSC(1).SetFocus
            txtSC_GotFocus 1
         Case Else
      End Select
   End If
End Sub

'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         'Modifeid by Lydia 2016/01/28 改到OnAction控制
         'If m_bUpdate And txtSC(1) <> "" Then
         If m_bUpdate Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         'Modifeid by Lydia 2016/01/28 改到OnAction控制
         'If m_bDelete And txtSC(1) <> "" Then
         If m_bDelete Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtSC(1) <> "" Then
            TBar1.Buttons(6).Enabled = True
            TBar1.Buttons(7).Enabled = True
            TBar1.Buttons(8).Enabled = True
            TBar1.Buttons(9).Enabled = True
         Else
            TBar1.Buttons(6).Enabled = False
            TBar1.Buttons(7).Enabled = False
            TBar1.Buttons(8).Enabled = False
            TBar1.Buttons(9).Enabled = False
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
      
      Case 1, 2, 3, 4 '維護
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010610 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   For nIndex = 1 To TF_SC
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex).fiName = "SC" & strTmp
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
      m_FieldList(nIndex).fiType = 0
      '定義數字
      Select Case nIndex
         Case 1, 2, 12, 13, 15, 16, 18, 19
            m_FieldList(nIndex).fiType = 1
      End Select
   Next nIndex
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 1 To TF_SC
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

' 更新欄位的內容
Private Sub UpdateFieldNewData()
Dim sC02 As Integer
  
   For Each oText In txtSC
      If oText.Index = 1 Then
         SetFieldNewData "SC01", DBDATE(oText.Text)
         If m_EditMode = 2 And oText.Text <> oText.Tag Then
            intI = 1
            strExc(1) = "select nvl(max(sc02),0) from staff_calendar where sc01=" & CNULL(DBDATE(oText.Text), True)
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
            If intI = 1 Then
               sC02 = RsTemp.Fields(0) + 1
            End If
            txtSC(2).Text = sC02
         End If
      Else
         SetFieldNewData "SC" & Right("0" & oText.Index, 2), oText.Text
      End If
   Next
   '新增
   If m_EditMode = 1 Then
      SetFieldNewData "SC11", strUserNum
      SetFieldNewData "SC12", strSrvDate(1)
      SetFieldNewData "SC13", Mid(Right("000000" & ServerTime, 6), 1, 4)
   '修改
   ElseIf m_EditMode = 2 Then
      SetFieldNewData "SC14", strUserNum
      SetFieldNewData "SC15", strSrvDate(1)
      SetFieldNewData "SC16", Mid(Right("000000" & ServerTime, 6), 1, 4)
   End If

End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String

   For nIndex = 1 To TF_SC
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
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0, Optional ByVal bolMsg As Boolean = True) As Boolean
   
   Dim adoRst As New ADODB.Recordset
   Dim stCon As String
   
   If p_iWay = -1 Or p_iWay = 1 Or p_iWay = 0 Then
      If txtSC(1) <> "" Then stCon = stCon & " and SC01=" & CNULL(DBDATE(txtSC(1)), True)
      If txtSC(2) <> "" Then stCon = stCon & " and SC02=" & CNULL(txtSC(2), True)
   End If

   strExc(0) = "SELECT * FROM Staff_Calendar where "
   Select Case p_iWay
      '尋找
      Case 0
          strExc(0) = strExc(0) & " 1=1 " & stCon
      '首筆
      Case -2
          strExc(0) = strExc(0) & "SC01||lpad(SC02,4,'0')=(select min(SC01||lpad(SC02,4,'0')) from staff_calendar) "
      '前一筆
      Case -1
          strExc(0) = strExc(0) & "SC01||lpad(SC02,4,'0')=(select max(SC01||lpad(SC02,4,'0')) from staff_calendar where SC01||lpad(SC02,4,'0') <'" & DBDATE(txtSC(1)) & Format(txtSC(2), "0000") & "') "
      '後一筆
      Case 1
          strExc(0) = strExc(0) & "SC01||lpad(SC02,4,'0')=(select min(SC01||lpad(SC02,4,'0')) from staff_calendar where SC01||lpad(SC02,4,'0') >'" & DBDATE(txtSC(1)) & Format(txtSC(2), "0000") & "') "
      '末筆
      Case 2
          strExc(0) = strExc(0) & "SC01||lpad(SC02,4,'0')=(select max(SC01||lpad(SC02,4,'0')) from staff_calendar) "
      'Form_Load 時,預設使用者未解除期限的最近行事曆提醒
      Case 9
          strExc(0) = strExc(0) & "SC01||lpad(SC02,4,'0')=(select min(SC01||lpad(SC02,4,'0')) from staff_calendar where SC17||SC18||SC19 IS NULL AND SC01||lpad(SC02,4,'0') >'" & strSrvDate(1) & "0000" & "' " & _
                      "AND INSTR(SC03||','||SC11,'" & strUserNum & "') > 0 ) "
   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      UpdateFieldOldData adoRst
      UpdateCtrlData adoRst
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         If bolMsg = True Then MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         If bolMsg = True Then MsgBox "已經是最後筆！", vbInformation
      Else
         If bolMsg = True Then MsgBox "查無資料！", vbInformation
      End If
   End If
   'Modified by Lydia 2016/01/28
   'If m_EditMode = 0 Then
   If m_EditMode <> "1" And m_EditMode <> "2" Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing

End Function
' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
Dim CUID(1 To 6) As String
Dim tmpArr As Variant, tmpBol As Boolean
   ClearField
   With p_Rst
      If .RecordCount > 0 Then
         For Each oText In txtSC
            idx = oText.Index
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
            m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
            If idx = 1 Then
               oText.Text = ChangeWStringToTString(m_FieldList(idx).fiOldData)
            Else
               oText.Text = m_FieldList(idx).fiOldData
            End If
            oText.Tag = oText.Text
         Next
         If txtSC(5).Text <> "" And txtSC(6).Text <> "" Then
            If GetPdata(txtSC(5).Text, txtSC(6).Text, txtSC(7).Text, txtSC(8).Text) Then
            End If
         End If
         If txtSC(4) <> "" Then
            tmpArr = Empty: lstSC04.Clear
            tmpArr = Split(txtSC(4), vbCrLf)
            For idx = 0 To UBound(tmpArr)
               If tmpArr(idx) <> "" Then
                  lstSC04.AddItem Trim(tmpArr(idx))
               End If
            Next
         End If
         
         CUID(1) = "" & .Fields("SC11")
         CUID(2) = "" & .Fields("SC12")
         CUID(3) = "" & .Fields("SC13")
         CUID(4) = "" & .Fields("SC14")
         CUID(5) = "" & .Fields("SC15")
         CUID(6) = "" & .Fields("SC16")
         mSC11 = "" & .Fields("SC11")
         
         If Not IsNull(.Fields("SC17")) Then
            'Modified by Lydia 2016/08/16 遇到離職人員不彈訊息
            'If ClsPDGetStaff(.Fields("SC17"), strExc(1)) = True Then
            '   lblCancel(0) = strExc(1)
            'End If
            lblCancel(0) = GetStaffName(.Fields("SC17"), True)
         End If
         If Not IsNull(.Fields("SC18")) Then
            lblCancel(1) = ChangeWStringToTDateString(.Fields("SC18"))
         End If
         If Not IsNull(.Fields("SC19")) Then
            lblCancel(2) = Format(.Fields("SC19"), "##:##")
         End If
         
         Call JudgeRight("R")

         If txtSC(3) <> "" Then
            SetlstUsers 0, txtSC(3)
         End If
         If txtSC(9) <> "" Then
            SetlstUsers 1, txtSC(9)
         End If
      End If
   End With
   UpdateCUID CUID, textCUID

End Sub
'判斷是否能修改或解除
Private Sub JudgeRight(ByVal iMode As String)
Dim tmpArr As Variant
Dim idR As Integer
Dim tmpBol As Boolean
    
    bolUpdate = False
    cmdCancel.Visible = False
    If iMode = "R" Then
        bolUpdate = False
        cmdCancel.Visible = False
        
        '個人輸入資料僅個人及各級帶人主管及M51人員可修改；輸入人員不可刪除只可解除；
        'Modified by Lydia 2017/06/20 解除人員可修改（解除人員的主管不用）。
        'If UCase(strUserNum) = UCase(mSC11) Or Pub_StrUserSt03 = "M51" Or PUB_GetST52(mSC11, strUserNum) Then
        'Modified by Lydia 2018/09/07 行事曆的可修改人員請增加程序、承辦的管制人員
        'If UCase(strUserNum) = UCase(mSC11) Or Pub_StrUserSt03 = "M51" Or PUB_GetST52(mSC11, strUserNum) Or (txtSC(9) <> "" And InStr(txtSC(9), strUserNum) > 0) Then
        If txtSC(5) <> "" And txtSC(6) <> "" Then
             strExc(1) = PUB_GetFCPHandler(txtSC(5), txtSC(6), txtSC(7), txtSC(8)) '程序管制
             strExc(2) = PUB_GetFCPSalesNo(txtSC(5), txtSC(6), txtSC(7), txtSC(8)) '承辦管制
        Else
             strExc(1) = "": strExc(2) = ""
        End If
        If UCase(strUserNum) = UCase(mSC11) Or Pub_StrUserSt03 = "M51" Or PUB_GetST52(mSC11, strUserNum) Or (txtSC(9) <> "" And InStr(txtSC(9), strUserNum) > 0) _
                Or strExc(1) = UCase(strUserNum) Or strExc(2) = UCase(strUserNum) Then
        'end 2018/09/07
           bolUpdate = True
           cmdCancel.Visible = True
        End If
        '輸入人員或其ST52，或可解除人員及其ST52查詢此筆時才可解除
        tmpArr = Empty
        tmpArr = Split(txtSC(9).Text, ",")
        tmpBol = False
        For idR = 0 To UBound(tmpArr)
            If tmpArr(idR) <> "" Then
               If UCase(tmpArr(idR)) = strUserNum Then
                  tmpBol = True
               ElseIf PUB_GetST52(tmpArr(idR), strUserNum) = True Then
                  tmpBol = True
               End If
            End If
            If tmpBol = True Then Exit For
        Next
        If tmpBol = True Then
           cmdCancel.Visible = True
        End If
    ElseIf iMode = "A" Then
           bolUpdate = True
           cmdCancel.Visible = True
    End If
End Sub

Private Sub ClearField()
   
   For Each oText In txtSC
      oText.Text = Empty
      oText.Tag = Empty
   Next
   If m_EditMode = 1 Then
      txtSC(1) = Left(strSrvDate(2), 5)
      txtSC(10) = "1"
      mSC11 = strUserNum
   Else
      mSC11 = ""
   End If
   For intI = 1 To TF_SC
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   bolUpdate = False: bolClose = False
   cmdCancel.Visible = False
   strSCT02 = ""
   textCUID = ""
   cboSCT02.Text = ""
   lstSCT03.Clear
   lstSC04.Clear
   Cmb1.Text = "": Cmb1.Clear
   txtUserNo(0) = "":    txtUserNo(1) = ""
   For Each oLabel In lblName
      If oLabel.Index < 2 Then oLabel.Caption = ""
   Next
   lstUsers(0).Clear
   lstUsers(1).Clear
   lblFC(0) = "": lblFC(1) = ""
   lblCancel(0) = "": lblCancel(1) = "": lblCancel(2) = ""
   'Added by Lydia 2021/04/21
   lstUsers(0).Tag = ""
   lstUsers(1).Tag = ""
   Erase strCase 'Added by Morgan 2024/3/15
End Sub

'Modified by Lydia 2021/04/21
'Private Function ComposeList(oList As ListBox) As String
Private Function ComposeList(oList As Control) As String
   Dim iPos As Integer, stItem As String
   strExc(1) = ""
   If oList.ListCount > 0 Then
      For intI = 0 To oList.ListCount - 1
         iPos = InStr(oList.List(intI), Chr(1))
         If iPos > 0 Then
            stItem = Left(oList.List(intI), iPos - 1)
         Else
            stItem = oList.List(intI)
         End If
         If intI = 0 Then
            strExc(1) = stItem
         Else
            strExc(1) = strExc(1) & vbCrLf & stItem
         End If
      Next
   End If
   ComposeList = strExc(1)
End Function

Private Sub GRD1_DblClick()
   If GRD1.MouseRow > 0 And GRD1.TextMatrix(GRD1.row, 0) <> "" Then
      txtSC(1) = GRD1.TextMatrix(GRD1.row, 0)
      txtSC(2) = GRD1.TextMatrix(GRD1.row, 1)
      If ShowRecord(0, True) Then
         SSTab1.Tab = 0
      End If
   End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If InStr("管制日期,流水號", Me.GRD1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

'Added by Lydia 2017/01/04 雙擊事由,將事由代入管制分類,方便修改內容
'Modified by Lydia 2021/04/21 改成Form 2.0
'Private Sub lstSC04_DblClick()
Private Sub lstSC04_DblClick(Cancel As MSForms.ReturnBoolean)
Dim tmpPos As Integer
tmpPos = lstSC04.ListIndex

   If lstSC04.List(tmpPos) <> "" And (m_EditMode = 1 Or m_EditMode = 2) Then
      If cboSCT02.Text <> "" Then
         If MsgBox("是否取代現有的管制分類內容 ？", vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
      End If
      cboSCT02.Text = lstSC04.List(tmpPos)
   End If
   
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   'Added by Lydia 2019/09/24 依照目前頁籤，區分Enter鍵的動作
   If PreviousTab = 0 Then '在"多筆查詢"頁籤按Enter鍵，可自動執行查詢。
       cmdSearch.Default = True
       cmdSearch.SetFocus
   Else          '在"單筆資料"頁籤按Enter鍵，狀態為修改或單筆查詢，可自動執行確定作業。
       cmdSearch.Default = False
   End If
End Sub


Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
       Case 0, 1
            KeyAscii = Pub_NumAscii(KeyAscii)
       Case Else
            KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Dim iLen As Integer
   Select Case Index

      Case 0, 1
         If txt1(Index) <> "" Then
            If CheckIsTaiwanDate(txt1(Index)) = False Then
                GoTo JumpCancel
            Else
                If txt1(0) <> "" And txt1(1) <> "" And txt1(0) > txt1(1) Then
                   MsgBox "管制日期止不可小於管制日期起!", vbCritical, "輸入錯誤"
                   GoTo JumpCancel
                End If
            End If
         End If

      Case 2, 3
         If txt1(Index) <> "" Then
            If Len(txt1(Index)) = 5 Then
               'Modified by Lydia 2016/08/16 遇到離職人員不彈訊息
               'If ClsPDGetStaff(Txt1(Index), strExc(1)) = True Then
               '   lblName(Index) = strExc(1)
               'Else
               '   lblName(Index) = ""
               'End If
               lblName(Index) = GetStaffName(txt1(Index), True)
            End If
            If lblName(Index) = "" Then
               MsgBox "員工編號輸入錯誤！", vbExclamation
               Cancel = True
            End If
         Else
            lblName(Index) = ""
         End If
   End Select
   
   If Cancel = False Then
      If txt1(Index).MaxLength > 0 Then
         If Not CheckLengthIsOK(txt1(Index), iLen) Then
            GoTo JumpCancel
         End If
      End If
   End If
   Exit Sub
   
JumpCancel:
   txt1_GotFocus Index
   Cancel = True
End Sub
Private Sub lblCancel_Change(Index As Integer)
   If Index = 1 Then
      If lblCancel(1) <> "" Then
         cmdCancel.Enabled = False
      Else
         cmdCancel.Enabled = True
      End If
   End If
End Sub

Private Sub txtSC_GotFocus(Index As Integer)
    TextInverse txtSC(Index)
End Sub

'Modified by Lydia 2021/04/21
'Private Sub txtSC_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtSC_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
       Case 1, 2, 10
            KeyAscii = Pub_NumAscii(KeyAscii)
       Case Else
            KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txtSC_Validate(Index As Integer, Cancel As Boolean)
   Dim iLen As Integer
   Select Case Index
      Case 2
         If txtSC(Index) <> "" Then
            txtSC(Index) = Trim(Val(txtSC(Index)))
         End If
      Case 6, 7, 8
         If txtSC(5) = "" And txtSC(Index) <> "" Then
            MsgBox "請輸入系統別!", vbCritical, "輸入錯誤"
            txtSC(5).SetFocus
            txtSC_GotFocus 5
            Cancel = True
            Exit Sub
         ElseIf txtSC(5) <> "" And txtSC(Index).Text <> "" And txtSC(Index).Text <> txtSC(Index).Tag Then
               If txtSC(7).Text = "" Then txtSC(7).Text = "0"
               If txtSC(8).Text = "" Then txtSC(8).Text = "00"
            If GetPdata(txtSC(5).Text, txtSC(6).Text, txtSC(7).Text, txtSC(8).Text) Then
               For intI = 5 To 8
                   txtSC(intI).Tag = txtSC(intI).Text
               Next
            Else
               For intI = 5 To 8
                   txtSC(intI).Tag = ""
               Next
               txtSC(Index).SetFocus
               txtSC_GotFocus Index
               Cancel = True
               Exit Sub
            End If
         End If
      Case 10
         If m_EditMode <> 4 And txtSC(Index) = "" Then
            MsgBox "週期不可空白!", vbCritical, "輸入錯誤"
            GoTo JumpCancel
         'Modified by Lydia 2016/06/28 +4
         'Modified by Lydia 2017/11/14 +5
         ElseIf txtSC(Index) <> "" And InStr("1,2,3,4,5", txtSC(Index).Text) = 0 Then
            MsgBox "請輸入1-5!", vbCritical, "輸入錯誤"
            GoTo JumpCancel
         End If
   End Select
   
   If Cancel = False Then
      If txtSC(Index).MaxLength > 0 Then
         If Not CheckLengthIsOK(txtSC(Index), iLen) Then
            GoTo JumpCancel
         End If
      End If
   End If
   Exit Sub
   
JumpCancel:
   txtSC_GotFocus Index
   Cancel = True
End Sub

Private Function GetPdata(ByVal Cc01 As String, Cc02 As String, Optional ByVal Cc03 As String, Optional ByVal Cc04 As String) As Boolean
Dim inX As Integer
Dim Str01 As String, Str02 As String, Str03 As String, sDate01, sDate02 As String

GetPdata = False
bolClose = False 'Added by Lydia 2016/01/21
Cmb1.Clear
lblFC(0) = "": lblFC(1) = ""

If Cc03 = "" Then Cc03 = "0"
If Cc04 = "" Then Cc04 = "00"
'Added by Lydia 2023/07/28
strCase(1) = "": strCase(2) = "": strCase(3) = "": strCase(4) = ""
m_PA177 = ""
'end 2023/07/28

Dim strSql As String, intCaseKind As Integer

If ClsPDGetSystemKind(Cc01, intCaseKind) Then

   Select Case intCaseKind
      Case 專利
         'Modified by Lydia 2023/07/28 +PA177
         strSql = "select pa05,pa06,pa07,pa108,pa136,pa75,NVL(FA05,NVL(FA04,FA06)),pa57,PA177 from patent,fagent " & _
            "where pa01=" & CNULL(Cc01) & " and pa02=" & CNULL(Cc02) & " and pa03=" & CNULL(Cc03) & " and pa04=" & CNULL(Cc04) & _
            " and substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) "
      Case 商標
         'Modified by Lydia 2023/07/28 +'' as PA177
         strSql = "select tm05,tm06,tm07,tm57,tm73,tm44,NVL(FA05,NVL(FA04,FA06)),tm29,'' as PA177 from trademark,fagent " & _
            "where tm01=" & CNULL(Cc01) & " and tm02=" & CNULL(Cc02) & " and tm03=" & CNULL(Cc03) & " and tm04=" & CNULL(Cc04) & _
            " and substr(TM44,1,8)=FA01(+) And substr(TM44,9,1)=FA02(+) "
      Case 法務
         'Modified by Lydia 2023/07/28 +'' as PA177
         strSql = "select lc05,lc06,lc07,lc34,lc36,lc22,NVL(FA05,NVL(FA04,FA06)),lc08,'' as PA177 from lawcase,FAGENT " & _
            "where lc01=" & CNULL(Cc01) & " and lc02=" & CNULL(Cc02) & " and lc03=" & CNULL(Cc03) & " and lc04=" & CNULL(Cc04) & _
            " and substr(LC22,1,8)=FA01(+) And substr(LC22,9,1)=FA02(+) "
      Case 顧問
         'Modified by Lydia 2023/07/28 +'' as PA177
         strSql = "select hc06,'','',hc19,hc20,'','',hc09,'' as PA177 from hirecase " & _
            "where hc01=" & CNULL(Cc01) & " and hc02=" & CNULL(Cc02) & " and hc03=" & CNULL(Cc03) & " and hc04=" & CNULL(Cc04)
      Case Else
         'Modified by Lydia 2023/07/28 +'' as PA177
         strSql = "select sp05,sp06,sp07,sp61,sp68,sp26,NVL(FA05,NVL(FA04,FA06)),sp15,'' as PA177 from servicepractice,FAGENT " & _
            "where sp01=" & CNULL(Cc01) & " and sp02=" & CNULL(Cc02) & " and sp03=" & CNULL(Cc03) & " and sp04=" & CNULL(Cc04) & _
            " and substr(SP26,1,8)=FA01(+) And substr(SP26,9,1)=FA02(+) "
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
       If "" & RsTemp(3) & RsTemp(7) <> "" Then
          bolClose = True
       End If
       txtSC(7) = Cc03: txtSC(8) = Cc04
       Cmb1.AddItem "中 : " & Trim(RsTemp(0))
       Cmb1.AddItem "英 : " & Trim(RsTemp(1))
       'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
       Cmb1.AddItem "外 : " & Trim(RsTemp(2))
       Cmb1.Text = "中 : " & Trim(RsTemp(0))
       lblFC(0) = ChangeCustomerS("" & RsTemp(5))
       lblFC(1) = "" & Trim(RsTemp(6))
       'Added by Lydia 2023/07/28
       strCase(1) = Cc01
       strCase(2) = Cc02
       strCase(3) = Cc03
       strCase(4) = Cc04
       m_PA177 = "" & RsTemp.Fields("PA177") 'FCP專利連結通知
       'end 2023/07/28
       GetPdata = True
   Else
       ShowMsg MsgText(9141)
   End If
End If

End Function

'Modified by Lydia 2021/04/21
'Private Function AddList(ByRef iCbo1 As ComboBox, ByRef iLst1 As ListBox, Optional ByRef oList As ListBox) As Boolean
Private Function AddList(ByRef iCbo1 As Control, ByRef iLst1 As Control, Optional ByRef oList As Control) As Boolean
   Dim idx As Integer, bFound As Boolean, stNewItem As String
   Dim strT As String, iPos As Integer
   
   If iCbo1.Text = "" Then
      Exit Function
   End If

   If iLst1.ListCount > 0 Then
      For idx = 0 To iLst1.ListCount - 1
         If iLst1.Selected(idx) = True Then
            strT = strT & "、" & iLst1.List(idx)
         End If
      Next
   End If
   'Modified by Lydia 2017/06/13 + chgsql 去除單引號
   'Remove by Lydia 2017/07/18 -chgsql
   stNewItem = iCbo1.Text & IIf(strT <> "", ": " & IIf(Left(strT, 1) = "、", Mid(strT, 2), strT), "")
   '若有控制字元時後面為說明文字不抓
   iPos = InStr(stNewItem, Chr(1))
   If iPos > 0 Then
      stNewItem = Left(stNewItem, iPos - 1)
   Else
      stNewItem = stNewItem
   End If
   If InStr(stNewItem, ";") > 0 Then
      MsgBox "分號[;]為系統保留字，請改用其他符號！", vbExclamation
      iCbo1.SetFocus
      Exit Function
   End If

   If stNewItem <> "" Then
      oList.AddItem stNewItem, 0
      AddList = True
   End If
End Function

'Modified by Lydia 2021/04/21
'Private Function RemoveList(ByRef oList As ListBox) As Boolean
Private Function RemoveList(ByRef oList As Control) As Boolean
   Dim ii As Integer
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            RemoveList = True
            oList.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
   End If
End Function

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: OnAction vbKeyF5
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub
' 執行指令
'Private Sub OnAction(ByVal KeyCode As Integer)
Public Sub OnAction(ByVal KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         m_EditMode = 1
         bSC01 = txtSC(1)
         bSC02 = txtSC(2)
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SSTab1.Tab = 0 'Added by Lydia 2017/11/14
         SetInputEntry
         
      Case vbKeyF3 ' 修改
         If bolUpdate = False Then
            MsgBox "無權限修改本記錄!!", vbCritical
            Exit Sub
         End If
         'Added by Lydia 2016/01/28 不在UpdateToolbarState控制
         If txtSC(1) = "" Then
            MsgBox "無記錄可修改!!", vbCritical
            Exit Sub
         End If
         'end 2016/01/28
         'Added by Lydia 2016/02/02 必須切換到單筆維護介面
         If SSTab1.Tab = 1 Then
            MsgBox "請先選擇記錄並切換到單筆資料!", vbCritical
            Exit Sub
         End If
         'end 2016/02/22
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF5 ' 刪除
         'Added by Lydia 2016/01/28 不在UpdateToolbarState控制
         If txtSC(1) = "" Then
            MsgBox "無記錄可刪除!!", vbCritical
            Exit Sub
         End If
         'end 2016/01/28
         If DelMsg() Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         m_EditMode = 4
         SetCtrlReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
         
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
         
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
         
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
         
      Case vbKeyF9 ' 確定
         If m_EditMode = "1" Then
            Call JudgeRight("A")
         End If
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
       '  SetCtrlReadOnly False 'Added by Lydia 2016/01/28
         SetInputEntry
         
      Case vbKeyF10 ' 取消
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  If m_EditMode = 1 Then
                     m_EditMode = 0
                     ClearField
                     SetInputEntry
                     If bSC01 <> "" Then
                        txtSC(1).Text = bSC01
                        txtSC(2).Text = bSC02
                     End If
                  Else
                     m_EditMode = 0
                     txtSC(1).Text = txtSC(1).Tag
                     txtSC(2).Text = txtSC(2).Tag
                  End If
                  ShowRecord 0, False
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               SetInputEntry
               ShowRecord 2, False
               UpdateToolbarState
         End Select
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

Private Function OnWork() As Boolean
Dim bolR  As Boolean

   bolR = False
   
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If AddRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord 0, False
               bolR = True
            End If
         End If
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord 0, False
               bolR = True
            End If
         End If
      Case 3: '刪除
         If DelRecord = True Then
            OnWork = True
            m_EditMode = 0
            ShowRecord 2, False
         End If
      Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord(0, True) = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtSC(1).SetFocus
               txtSC_GotFocus 1
            End If
         End If
   End Select
   
   '多筆資料-整理
   If bolR And txt1(3) = strUserNum Then
      If QueryData(False) = False Then
      End If
   End If
End Function

Private Function TxtValidate() As Boolean
Dim tmpArr As Variant
   Dim Cancel As Boolean, ii As Integer, jj As Integer

TxtValidate = False

   If txtSC(1) = "" Then
       MsgBox "管制日期不可空白!", vbCritical, "輸入錯誤"
       txtSC(1).SetFocus
       txtSC_GotFocus 1
       Cancel = True
       Exit Function
   Else
       If CheckIsTaiwanDate(txtSC(1)) = False Then
            txtSC(1).SetFocus
            txtSC_GotFocus 1
            Cancel = True
            Exit Function
       'Modified by Morgan 2025/2/19 +修改也要檢查工作天
       'ElseIf m_EditMode = 1 Then
       ElseIf (m_EditMode = 1 Or m_EditMode = 2) Then
          If ChkWorkDay(DBDATE(txtSC(1))) = False Then
              MsgBox "管制日期應為工作天!", vbCritical, "輸入錯誤"
              txtSC(1).SetFocus
              txtSC_GotFocus 1
              Cancel = True
              Exit Function
          ElseIf txtSC(1) < strSrvDate(2) Then
              MsgBox "管制日期不可小於系統日!", vbCritical, "輸入錯誤"
              txtSC(1).SetFocus
              txtSC_GotFocus 1
              Cancel = True
              Exit Function
          End If
       End If
   End If
   
   For Each oText In txtSC
       Cancel = False
       If oText.Index < 5 Or oText.Index > 9 Then
          txtSC_Validate oText.Index, Cancel
          If Cancel = True Then
             oText.SetFocus
             txtSC_GotFocus oText.Index
             Exit Function
          End If
       End If
   Next

   If m_EditMode > 0 And m_EditMode < 4 Then
       If txtSC(5) = "" And txtSC(6) & txtSC(7) & txtSC(8) <> "" Then
            MsgBox "請輸入系統別!", vbCritical, "輸入錯誤"
            txtSC_GotFocus 5
            Cancel = True
            Exit Function
       ElseIf txtSC(5) <> "" And txtSC(6) = "" Then
            MsgBox "請輸入完整本所案號!", vbCritical, "輸入錯誤"
            txtSC_GotFocus 6
            Cancel = True
            Exit Function
       ElseIf txtSC(5) <> "" And txtSC(6) <> "" Then
          If txtSC(5).Text <> txtSC(5).Tag Or txtSC(6).Text <> txtSC(6).Tag Or txtSC(7).Text <> txtSC(7).Tag Or txtSC(8).Text <> txtSC(8).Tag Then
            If GetPdata(txtSC(5).Text, txtSC(6).Text, txtSC(7).Text, txtSC(8).Text) = False Then
               txtSC(6).SetFocus
               txtSC_GotFocus 6
               Exit Function
            End If
          End If
          If bolClose Then
              If MsgBox("本案有北所銷卷日或已閉卷,確定繼續存檔?", vbInformation + vbYesNo) = vbNo Then
                 Cancel = False
                 txtSC(6).SetFocus
                 txtSC_GotFocus 6
                 Exit Function
              Else
                 bolClose = False
              End If
          End If
       End If
       '事由
       If txtSC(4).Text = "" Then
           MsgBox "事由不可空白!", vbCritical, "輸入錯誤"
           cboSCT02.SetFocus
           Cancel = True
           Exit Function
       End If
       '提醒人員
       If txtSC(3).Text = "" Then
           MsgBox "提醒人員不可空白!", vbCritical, "輸入人員"
           txtUserNo(0).SetFocus
           txtUserNo_GotFocus 0
           Cancel = True
           Exit Function
       Else
           tmpArr = Empty
           tmpArr = Split(txtSC(3), ",")
           For intI = 0 To UBound(tmpArr)
               If tmpArr(intI) <> "" Then
                  If Mid(tmpArr(intI), 4, 1) = "9" Then
                     MsgBox "虛建編號不可為提醒人員!", vbCritical, "輸入人員"
                     lstUsers(0).SetFocus
                     Cancel = True
                  End If
               End If
           Next
       End If
       
       '可解除人員
       If txtSC(9).Text = "" Then
           MsgBox "可解除人員不可空白!", vbCritical, "輸入人員"
           txtUserNo(1).SetFocus
           txtUserNo_GotFocus 0
           Cancel = True
           Exit Function
       End If
   End If
 
   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         Cancel = True
         Exit Function
   End If
   
   TxtValidate = True
   
End Function

' 新增記錄
Private Function AddRecord() As Boolean
   Dim stSQL As String, stCols As String, stValues As String
   Dim stMem As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans

   If txtSC(2) = "" Then
      intI = 1
      strSql = "select nvl(max(sc02),0) from staff_calendar where sc01=" & CNULL(DBDATE(txtSC(1)), True)
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         m_FieldList(2).fiNewData = RsTemp.Fields(0) + 1
      End If
   End If

   '畫面有的欄位才更新
   For idx = 1 To TF_SC
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
         End If
         'Added by Lydia 2020/01/15 記錄提醒人員
         If m_FieldList(idx).fiName = "SC03" Then
              stMem = ChgSQL(m_FieldList(idx).fiNewData)
         End If
         'end 2020/01/15
      End If
   Next

   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO Staff_Calendar (" & stCols & ") Values (" & stValues & ")"
   
   cnnConnection.Execute stSQL
   
   'Added by Lydia 2020/01/15 新增存檔時若提醒人員有外專以外的人員，若不存在於非外專有行事曆提醒人員檔，再新增此檔資料。
   stSQL = "select st01, " & strSrvDate(1) & " from staff where st04='1' and st01 in (" & GetAddStr(stMem) & ") and st03 not like 'F2%' and st01 not in (select scm01 from staff_calendar_member) "
   stSQL = "Insert Into staff_calendar_member (scm01,scm02) " & stSQL
   cnnConnection.Execute stSQL, intI
   'end 2020/01/15
   
   cnnConnection.CommitTrans
   AddRecord = True
   
   txtSC(2) = m_FieldList(2).fiNewData
   txtSC(2).Tag = txtSC(2)
   Exit Function
   
ErrHand:
   If Err.Number > 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If

End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
      stSQL = "delete from Staff_Calendar where SC01=" & CNULL(DBDATE(txtSC(1)), True) & " and SC02=" & CNULL(txtSC(2), True)
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL
   
   cnnConnection.CommitTrans
   
   DelRecord = True

   Exit Function
   
ErrHand:
   If Err.Number > 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand

   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE Staff_Calendar SET "
   stSet = ""

   For idx = 1 To TF_SC
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
          '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where SC01=" & CNULL(DBDATE(txtSC(1).Tag), True) & " and SC02=" & CNULL(txtSC(2).Tag, True) & "; end; "
      cnnConnection.Execute stSQL, intI
   End If
   
   cnnConnection.CommitTrans
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtSC
      If (m_EditMode = 1 Or m_EditMode = 2) And oText.Index = 2 Then
         oText.Locked = Not bLocked
      Else
         oText.Locked = bLocked
      End If
   Next
   
   Me.SSTab1.TabEnabled(1) = bLocked

   cboSCT02.Locked = bLocked
   lstSCT03.Enabled = Not bLocked
   cmdAddSC04.Enabled = Not bLocked
   cmdRemSC04.Enabled = Not bLocked

   lstUsers(0).Enabled = Not bLocked
   lstUsers(1).Enabled = Not bLocked
   Frame1.Enabled = Not bLocked
   Frame2.Enabled = Not bLocked
End Sub

' 更新 Create 及 Update 的人
'Modified by Lydia 2021/04/21
'Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As TextBox)
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Control)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
         
      Case vbKeyEscape:
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction KeyCode
            End If
         End If
         
      Case vbKeyReturn
        '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
        KeyCode = 0
        If m_EditMode <> 0 Then
           OnAction vbKeyF9
        End If
   End Select
End Sub


Private Sub AddlstUsers(p_idx As Integer)
   Dim idx As Integer, bFound As Boolean
   Dim tmpX As Integer
   If txtUserNo(p_idx) <> "" And lblName(p_idx) <> "" Then
      '員工編號已可非數字需做轉換
      For idx = 0 To lstUsers(p_idx).ListCount - 1
         'Modified by Lydia 2021/04/21
         'If lstUsers(p_idx).ItemData(idx) = PUB_Id2Num(txtUserNo(p_idx)) Then
         If InStr(lstUsers(p_idx).Tag, txtUserNo(p_idx)) > 0 Then
            MsgBox "員工已存在於清單中！"
            txtUserNo(p_idx).SetFocus
            txtUserNo_GotFocus p_idx
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         tmpX = IIf(lstUsers(p_idx).ListCount = -1, 0, lstUsers(p_idx).ListCount)
         lstUsers(p_idx).AddItem lblName(p_idx), tmpX
         'Modified by Lydia 2021/04/21
         'lstUsers(p_idx).ItemData(tmpX) = PUB_Id2Num(txtUserNo(p_idx))
         lstUsers(p_idx).Tag = lstUsers(p_idx).Tag & IIf(lstUsers(p_idx).Tag <> "", ",", "") & txtUserNo(p_idx)
         txtUserNo(p_idx) = ""
         lblName(p_idx) = ""
      End If
   End If
End Sub

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   Dim arrName 'Added by Lydia 2021/04/23
   
   lstUsers(p_idx).Clear
   lstUsers(p_idx).Tag = "" 'Added by Lydia 2021/04/23
   
   If p_stNums <> "" Then
      'Modified by Lydia 2021/04/23 改寫法
'      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         arrID = Split(p_stNums, ",")
'         With RsTemp
'         '照原順序排
'         For intI = UBound(arrID) To LBound(arrID) Step -1
'            .MoveFirst
'            Do While Not .EOF
'               If .Fields("st01") = arrID(intI) Then
'                  lstUsers(p_idx).AddItem "" & .Fields(1), 0
'                  '員工編號已可非數字需做轉換
'                  lstUsers(p_idx).ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
'                  .MoveLast
'               End If
'               .MoveNext
'            Loop
'         Next
'         End With
'      End If
      strExc(0) = "select getstaffnamelist('" & p_stNums & "') from dual"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
          arrID = Split(p_stNums, ",")
          arrName = Split("" & RsTemp.Fields(0), ",")
          For intI = UBound(arrID) To LBound(arrID) Step -1
               lstUsers(p_idx).AddItem arrName(intI), 0
               'Form 2.0的Listbox沒有ItemData,改放在.Tag; 讀取用PUB_GetItemData
               lstUsers(p_idx).Tag = arrID(intI) & IIf(lstUsers(p_idx).Tag <> "", ",", "") & lstUsers(p_idx).Tag
          Next intI
      End If
   End If
End Sub

Private Sub RemovelstUsers(p_idx As Integer)
   Dim idx As Integer, ii As Integer
   Dim strTmp As String 'Added by Lydia 2021/04/23
   
   If lstUsers(p_idx).ListCount > 0 Then
      ii = 0
      strTmp = "," & lstUsers(p_idx).Tag 'Added by Lydia 2021/04/23
      For idx = 0 To lstUsers(p_idx).ListCount - 1
         'Modified by Lydia 2021/04/23
         'If lstUsers(p_idx).Selected(ii) = True Then
         '   lstUsers(p_idx).RemoveItem ii
         '   ii = ii - 1
         'End If
         'ii = ii + 1
         If ii >= 0 Then
             If lstUsers(p_idx).Selected(ii) = True Then
                strTmp = Replace(strTmp, "," & PUB_GetItemData(lstUsers(p_idx).Tag, ii), "")
                lstUsers(p_idx).RemoveItem ii
                ii = ii - 1
             Else
                ii = ii + 1
             End If
         End If
         'end 2021/04/23
      Next
      lstUsers(p_idx).Tag = Mid(strTmp, 2) 'Added by Lydia 2021/04/23
   End If
End Sub

Private Sub txtUserNo_Change(Index As Integer)
   Dim strTempName As String
   If Len(txtUserNo(Index)) = 5 Then
      If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
         lblName(Index) = strTempName
      End If
   Else
      lblName(Index) = ""
   End If
End Sub

Private Sub txtUserNo_GotFocus(Index As Integer)
   TextInverse txtUserNo(Index)
End Sub

Private Sub txtUserNo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUserNo_Validate(Index As Integer, Cancel As Boolean)
   Dim strTempName As String
   
   If m_EditMode > 0 And m_EditMode < 4 Then
       If Index = 0 And Mid(txtUserNo(Index), 4, 1) = "9" Then
           MsgBox "虛建編號不可為提醒人員!", vbCritical, "輸入人員"
           Cancel = True
           Exit Sub
       End If
   End If
   
   If txtUserNo(Index).Visible = True Then
      If txtUserNo(Index) <> "" And lblName(Index) = "" Then
         If Len(txtUserNo(Index)) = 5 Then
            If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
               lblName(Index) = strTempName
            End If
         End If
         If lblName(Index) = "" Then
            MsgBox "員工編號輸入錯誤！", vbExclamation
            Cancel = True
         End If
      End If
   End If
End Sub

