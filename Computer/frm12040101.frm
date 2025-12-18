VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040101 
   BorderStyle     =   1  '單線固定
   Caption         =   "國家基本檔維護"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7992
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   7992
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7380
      Top             =   240
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
            Picture         =   "frm12040101.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040101.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040101.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040101.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040101.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040101.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040101.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040101.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040101.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040101.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040101.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   70
      Top             =   0
      Width           =   7992
      _ExtentX        =   14097
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
      Height          =   4812
      Left            =   0
      TabIndex        =   72
      Top             =   648
      Width           =   7908
      _ExtentX        =   13949
      _ExtentY        =   8488
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   529
      TabCaption(0)   =   "商標"
      TabPicture(0)   =   "frm12040101.frx":20F4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text1(88)"
      Tab(0).Control(1)=   "Text1(70)"
      Tab(0).Control(2)=   "Text1(71)"
      Tab(0).Control(3)=   "Text1(86)"
      Tab(0).Control(4)=   "Text1(87)"
      Tab(0).Control(5)=   "Text1(32)"
      Tab(0).Control(6)=   "Text1(84)"
      Tab(0).Control(7)=   "Text1(77)"
      Tab(0).Control(8)=   "Text1(76)"
      Tab(0).Control(9)=   "Text1(59)"
      Tab(0).Control(10)=   "Text1(69)"
      Tab(0).Control(11)=   "Text1(68)"
      Tab(0).Control(12)=   "Text1(58)"
      Tab(0).Control(13)=   "Text1(54)"
      Tab(0).Control(14)=   "Text1(51)"
      Tab(0).Control(15)=   "Combo1"
      Tab(0).Control(16)=   "Text1(0)"
      Tab(0).Control(17)=   "Text1(1)"
      Tab(0).Control(18)=   "Text1(2)"
      Tab(0).Control(19)=   "Text1(3)"
      Tab(0).Control(20)=   "Text1(4)"
      Tab(0).Control(21)=   "Text1(29)"
      Tab(0).Control(22)=   "Text1(30)"
      Tab(0).Control(23)=   "Text1(31)"
      Tab(0).Control(24)=   "Text1(34)"
      Tab(0).Control(25)=   "Text1(35)"
      Tab(0).Control(26)=   "Text1(36)"
      Tab(0).Control(27)=   "Text1(37)"
      Tab(0).Control(28)=   "Text1(38)"
      Tab(0).Control(29)=   "Text1(49)"
      Tab(0).Control(30)=   "Label1(64)"
      Tab(0).Control(31)=   "Label1(90)"
      Tab(0).Control(32)=   "Label1(89)"
      Tab(0).Control(33)=   "Label1(88)"
      Tab(0).Control(34)=   "Label1(82)"
      Tab(0).Control(35)=   "Label1(81)"
      Tab(0).Control(36)=   "Label1(61)"
      Tab(0).Control(37)=   "Label1(76)"
      Tab(0).Control(38)=   "Label1(75)"
      Tab(0).Control(39)=   "Label1(72)"
      Tab(0).Control(40)=   "Label1(71)"
      Tab(0).Control(41)=   "Label2(9)"
      Tab(0).Control(42)=   "Label1(70)"
      Tab(0).Control(43)=   "Label1(60)"
      Tab(0).Control(44)=   "Label1(56)"
      Tab(0).Control(45)=   "Label2(7)"
      Tab(0).Control(46)=   "Label1(53)"
      Tab(0).Control(47)=   "Label1(51)"
      Tab(0).Control(48)=   "Label3"
      Tab(0).Control(49)=   "Label1(38)"
      Tab(0).Control(50)=   "Label1(37)"
      Tab(0).Control(51)=   "Label1(36)"
      Tab(0).Control(52)=   "Label1(0)"
      Tab(0).Control(53)=   "Label1(1)"
      Tab(0).Control(54)=   "Label1(2)"
      Tab(0).Control(55)=   "Label1(3)"
      Tab(0).Control(56)=   "Label1(4)"
      Tab(0).Control(57)=   "Label1(29)"
      Tab(0).Control(58)=   "Label1(30)"
      Tab(0).Control(59)=   "Label1(31)"
      Tab(0).Control(60)=   "Label1(32)"
      Tab(0).Control(61)=   "Label1(34)"
      Tab(0).Control(62)=   "Label1(35)"
      Tab(0).ControlCount=   63
      TabCaption(1)   =   "專利年費設定"
      TabPicture(1)   =   "frm12040101.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Text1(78)"
      Tab(1).Control(2)=   "Text1(55)"
      Tab(1).Control(3)=   "Text1(56)"
      Tab(1).Control(4)=   "Text1(57)"
      Tab(1).Control(5)=   "Text1(5)"
      Tab(1).Control(6)=   "Text1(6)"
      Tab(1).Control(7)=   "Text1(8)"
      Tab(1).Control(8)=   "Text1(9)"
      Tab(1).Control(9)=   "Text1(10)"
      Tab(1).Control(10)=   "Text1(12)"
      Tab(1).Control(11)=   "Text1(13)"
      Tab(1).Control(12)=   "Text1(14)"
      Tab(1).Control(13)=   "Text1(39)"
      Tab(1).Control(14)=   "Text1(40)"
      Tab(1).Control(15)=   "Text1(16)"
      Tab(1).Control(16)=   "Text1(41)"
      Tab(1).Control(17)=   "Text1(33)"
      Tab(1).Control(18)=   "Text1(50)"
      Tab(1).Control(19)=   "Label2(12)"
      Tab(1).Control(20)=   "Label1(83)"
      Tab(1).Control(21)=   "Label1(74)"
      Tab(1).Control(22)=   "Label1(57)"
      Tab(1).Control(23)=   "Label1(58)"
      Tab(1).Control(24)=   "Label1(59)"
      Tab(1).Control(25)=   "Label2(6)"
      Tab(1).Control(26)=   "Label1(50)"
      Tab(1).Control(27)=   "Label2(3)"
      Tab(1).Control(28)=   "Label1(33)"
      Tab(1).Control(29)=   "Label2(2)"
      Tab(1).Control(30)=   "Label2(1)"
      Tab(1).Control(31)=   "Label2(0)"
      Tab(1).Control(32)=   "Label1(16)"
      Tab(1).Control(33)=   "Label1(5)"
      Tab(1).Control(34)=   "Label1(6)"
      Tab(1).Control(35)=   "Label1(8)"
      Tab(1).Control(36)=   "Label1(9)"
      Tab(1).Control(37)=   "Label1(10)"
      Tab(1).Control(38)=   "Label1(12)"
      Tab(1).Control(39)=   "Label1(13)"
      Tab(1).Control(40)=   "Label1(14)"
      Tab(1).Control(41)=   "Label1(39)"
      Tab(1).Control(42)=   "Label1(40)"
      Tab(1).Control(43)=   "Label1(41)"
      Tab(1).Control(44)=   "Label1(42)"
      Tab(1).Control(45)=   "Label2(4)"
      Tab(1).ControlCount=   46
      TabCaption(2)   =   "專利專用期間/實體/公開設定"
      TabPicture(2)   =   "frm12040101.frx":212C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(15)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(11)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(7)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(22)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(21)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label1(20)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label1(19)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label1(18)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label1(17)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label1(23)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label1(28)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label1(27)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label1(26)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label1(25)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label1(24)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label1(44)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label1(45)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label1(46)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label1(47)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Label1(48)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Label1(49)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Label1(43)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Label2(5)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Label1(52)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Label1(54)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Label1(55)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Label1(73)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Label1(85)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Label1(86)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Label1(87)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Text1(83)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Text1(82)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Text1(81)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Text1(47)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Text1(46)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Text1(45)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Text1(44)"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Text1(43)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Text1(42)"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Text1(15)"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Text1(11)"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Text1(7)"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Text1(28)"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Text1(27)"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "Text1(26)"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "Text1(25)"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "Text1(24)"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "Text1(23)"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "Text1(22)"
      Tab(2).Control(48).Enabled=   0   'False
      Tab(2).Control(49)=   "Text1(21)"
      Tab(2).Control(49).Enabled=   0   'False
      Tab(2).Control(50)=   "Text1(20)"
      Tab(2).Control(50).Enabled=   0   'False
      Tab(2).Control(51)=   "Text1(19)"
      Tab(2).Control(51).Enabled=   0   'False
      Tab(2).Control(52)=   "Text1(18)"
      Tab(2).Control(52).Enabled=   0   'False
      Tab(2).Control(53)=   "Text1(17)"
      Tab(2).Control(53).Enabled=   0   'False
      Tab(2).Control(54)=   "Text1(53)"
      Tab(2).Control(54).Enabled=   0   'False
      Tab(2).Control(55)=   "Text1(52)"
      Tab(2).Control(55).Enabled=   0   'False
      Tab(2).Control(56)=   "Text1(48)"
      Tab(2).Control(56).Enabled=   0   'False
      Tab(2).ControlCount=   57
      TabCaption(3)   =   "其他"
      TabPicture(3)   =   "frm12040101.frx":2148
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2(8)"
      Tab(3).Control(1)=   "Label1(63)"
      Tab(3).Control(2)=   "Label1(84)"
      Tab(3).Control(3)=   "Label1(65)"
      Tab(3).Control(4)=   "Label1(66)"
      Tab(3).Control(5)=   "Text1(79)"
      Tab(3).Control(6)=   "Frame2"
      Tab(3).Control(7)=   "Text1(62)"
      Tab(3).Control(8)=   "Text1(63)"
      Tab(3).ControlCount=   9
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   63
         Left            =   -72240
         MaxLength       =   2
         TabIndex        =   183
         Top             =   2784
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   62
         Left            =   -72240
         MaxLength       =   2
         TabIndex        =   182
         Top             =   2448
         Width           =   330
      End
      Begin VB.Frame Frame2 
         Caption         =   "專利"
         Height          =   1380
         Left            =   -74904
         TabIndex        =   178
         Top             =   912
         Width           =   7692
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   74
            Left            =   1896
            MaxLength       =   3
            TabIndex        =   181
            Top             =   648
            Width           =   330
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   75
            Left            =   1896
            MaxLength       =   3
            TabIndex        =   186
            Top             =   948
            Width           =   330
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   61
            Left            =   3252
            MaxLength       =   3
            TabIndex        =   180
            Top             =   276
            Width           =   330
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   60
            Left            =   1896
            MaxLength       =   3
            TabIndex        =   179
            Top             =   264
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PCT進國家階段月數"
            Height          =   180
            Index           =   79
            Left            =   144
            TabIndex        =   188
            Top             =   696
            Width           =   1572
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "恢復原狀延長月數"
            Height          =   180
            Index           =   80
            Left            =   144
            TabIndex        =   187
            Top             =   996
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "提申管制週期(天)                   , 底限(天)"
            Height          =   192
            Index           =   62
            Left            =   144
            TabIndex        =   184
            Top             =   312
            Width           =   3156
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   88
         Left            =   -67584
         MaxLength       =   2
         TabIndex        =   177
         Top             =   960
         Width           =   450
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   70
         Left            =   -69708
         MaxLength       =   3
         TabIndex        =   21
         Top             =   4020
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   71
         Left            =   -72888
         MaxLength       =   40
         TabIndex        =   22
         Top             =   4275
         Width           =   3000
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   86
         Left            =   -68592
         MaxLength       =   3
         TabIndex        =   174
         Top             =   4530
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   87
         Left            =   -68292
         MaxLength       =   1
         TabIndex        =   23
         Top             =   4275
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   705
         Left            =   -70500
         TabIndex        =   166
         Top             =   3240
         Visible         =   0   'False
         Width           =   3165
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   72
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   168
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   73
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   167
            Top             =   270
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CFP程序－單數"
            Height          =   180
            Index           =   77
            Left            =   0
            TabIndex        =   172
            Top             =   52
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CFP程序－雙數"
            Height          =   180
            Index           =   78
            Left            =   0
            TabIndex        =   171
            Top             =   315
            Width           =   1200
         End
         Begin MSForms.Label Label2 
            Height          =   255
            Index           =   10
            Left            =   2130
            TabIndex        =   170
            Top             =   15
            Width           =   900
            VariousPropertyBits=   27
            Caption         =   "Label2"
            Size            =   "1587;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label Label2 
            Height          =   255
            Index           =   11
            Left            =   2130
            TabIndex        =   169
            Top             =   315
            Width           =   900
            VariousPropertyBits=   27
            Caption         =   "Label2"
            Size            =   "1587;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   32
         Left            =   -68508
         MaxLength       =   2
         TabIndex        =   11
         Top             =   2010
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   84
         Left            =   -68508
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1770
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   79
         Left            =   -72975
         MaxLength       =   50
         TabIndex        =   185
         Top             =   3864
         Width           =   4890
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   78
         Left            =   -73080
         MaxLength       =   6
         TabIndex        =   41
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   77
         Left            =   -67968
         MaxLength       =   2
         TabIndex        =   157
         Top             =   3270
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   76
         Left            =   -71172
         MaxLength       =   1
         TabIndex        =   10
         Top             =   2010
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   59
         Left            =   -69228
         MaxLength       =   2
         TabIndex        =   154
         Top             =   960
         Width           =   450
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   69
         Left            =   -70092
         MaxLength       =   50
         TabIndex        =   20
         Top             =   3750
         Width           =   2685
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   68
         Left            =   -72888
         MaxLength       =   6
         TabIndex        =   19
         Top             =   4005
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   58
         Left            =   -68508
         MaxLength       =   1
         TabIndex        =   5
         Top             =   1500
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   55
         Left            =   -72540
         MaxLength       =   1
         TabIndex        =   37
         Top             =   3210
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   56
         Left            =   -72540
         MaxLength       =   1
         TabIndex        =   38
         Top             =   3450
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   57
         Left            =   -72540
         MaxLength       =   1
         TabIndex        =   39
         Top             =   3690
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   54
         Left            =   -72888
         MaxLength       =   6
         TabIndex        =   18
         Top             =   3750
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   48
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   55
         Top             =   3864
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   52
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   56
         Top             =   4104
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   53
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   57
         Top             =   4344
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   51
         Left            =   -69192
         MaxLength       =   4
         TabIndex        =   135
         Top             =   510
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         ItemData        =   "frm12040101.frx":2164
         Left            =   -69228
         List            =   "frm12040101.frx":2166
         Style           =   2  '單純下拉式
         TabIndex        =   24
         Top             =   510
         Width           =   1812
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   -72888
         MaxLength       =   4
         TabIndex        =   0
         Top             =   510
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   -72888
         MaxLength       =   3
         TabIndex        =   1
         Top             =   750
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   2
         Left            =   -72888
         MaxLength       =   20
         TabIndex        =   2
         Top             =   984
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   3
         Left            =   -72888
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1230
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   4
         Left            =   -72888
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1470
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   17
         Left            =   6600
         MaxLength       =   1
         TabIndex        =   58
         Top             =   525
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   18
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   59
         Top             =   765
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   19
         Left            =   6600
         MaxLength       =   1
         TabIndex        =   62
         Top             =   1635
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   20
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   63
         Top             =   1875
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   21
         Left            =   6600
         MaxLength       =   1
         TabIndex        =   66
         Top             =   2745
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   22
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   67
         Top             =   2985
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   23
         Left            =   6600
         MaxLength       =   1
         TabIndex        =   60
         Top             =   1005
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   24
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   61
         Top             =   1245
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   25
         Left            =   6600
         MaxLength       =   1
         TabIndex        =   64
         Top             =   2115
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   26
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   65
         Top             =   2355
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   27
         Left            =   6600
         MaxLength       =   1
         TabIndex        =   68
         Top             =   3225
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   28
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   69
         Top             =   3465
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   29
         Left            =   -72888
         MaxLength       =   1
         TabIndex        =   6
         Top             =   1740
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   30
         Left            =   -71172
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1770
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   31
         Left            =   -72888
         MaxLength       =   2
         TabIndex        =   9
         Top             =   2010
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   34
         Left            =   -72888
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   12
         Top             =   2280
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   35
         Left            =   -72888
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2760
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   36
         Left            =   -72888
         MaxLength       =   50
         TabIndex        =   14
         Top             =   3000
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   37
         Left            =   -72888
         MaxLength       =   2
         TabIndex        =   15
         Top             =   3270
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   38
         Left            =   -70548
         MaxLength       =   2
         TabIndex        =   16
         Top             =   3270
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   5
         Left            =   -73080
         MaxLength       =   4
         TabIndex        =   25
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   6
         Left            =   -68820
         MaxLength       =   1
         TabIndex        =   26
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   8
         Left            =   -73080
         MaxLength       =   80
         TabIndex        =   27
         Top             =   1080
         Width           =   5385
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   9
         Left            =   -73080
         MaxLength       =   4
         TabIndex        =   29
         Top             =   1590
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   10
         Left            =   -68820
         MaxLength       =   1
         TabIndex        =   30
         Top             =   1590
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   12
         Left            =   -73080
         MaxLength       =   80
         TabIndex        =   31
         Top             =   1830
         Width           =   5385
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   13
         Left            =   -73080
         MaxLength       =   4
         TabIndex        =   33
         Top             =   2340
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   14
         Left            =   -68820
         MaxLength       =   1
         TabIndex        =   34
         Top             =   2340
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   7
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   43
         Top             =   765
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   11
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   47
         Top             =   1785
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   15
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   51
         Top             =   2805
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   39
         Left            =   -73080
         MaxLength       =   1
         TabIndex        =   28
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   40
         Left            =   -73080
         MaxLength       =   1
         TabIndex        =   32
         Top             =   2070
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   16
         Left            =   -73080
         MaxLength       =   80
         TabIndex        =   35
         Top             =   2580
         Width           =   5385
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   41
         Left            =   -73080
         MaxLength       =   1
         TabIndex        =   36
         Top             =   2820
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   33
         Left            =   -73080
         MaxLength       =   6
         TabIndex        =   40
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   42
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   44
         Top             =   1005
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   43
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   45
         Top             =   1245
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   44
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   48
         Top             =   2025
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   45
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   49
         Top             =   2265
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   46
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   52
         Top             =   3045
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   47
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   53
         Top             =   3285
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   50
         Left            =   -69180
         MaxLength       =   6
         TabIndex        =   42
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   49
         Left            =   -72888
         MaxLength       =   1
         TabIndex        =   17
         Top             =   3510
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   81
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   46
         Top             =   1470
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   82
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   50
         Top             =   2490
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   83
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   54
         Top             =   3510
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CFT公告日更新商申催審月數"
         Height          =   180
         Index           =   66
         Left            =   -74784
         TabIndex        =   190
         Top             =   2808
         Width           =   2292
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CFT審查報告更新商申催審月數"
         Height          =   180
         Index           =   65
         Left            =   -74784
         TabIndex        =   189
         Top             =   2472
         Width           =   2568
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DHL國家代碼"
         Height          =   180
         Index           =   64
         Left            =   -68712
         TabIndex        =   176
         Top             =   984
         Width           =   1068
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "電子證書專利種類             (Ex:123)"
         Height          =   180
         Index           =   90
         Left            =   -70032
         TabIndex        =   175
         Top             =   4560
         Width           =   2652
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "商標是否電子證書         (Y:是)"
         Height          =   180
         Index           =   89
         Left            =   -69732
         TabIndex        =   173
         Top             =   4320
         Width           =   2316
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "計算商標專用期是否減1天         (Y/N)"
         Height          =   180
         Index           =   88
         Left            =   -70572
         TabIndex        =   165
         Top             =   1812
         Width           =   2880
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "計算設計專用期是否減1天                 (Y/N)"
         Height          =   180
         Index           =   87
         Left            =   240
         TabIndex        =   164
         Top             =   3510
         Width           =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "計算新型專用期是否減1天                 (Y/N)"
         Height          =   180
         Index           =   86
         Left            =   240
         TabIndex        =   163
         Top             =   2490
         Width           =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "計算發明專用期是否減1天                 (Y/N)"
         Height          =   180
         Index           =   85
         Left            =   240
         TabIndex        =   162
         Top             =   1470
         Width           =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "華銀結匯國家名稱"
         Height          =   180
         Index           =   84
         Left            =   -74784
         TabIndex        =   161
         Top             =   3864
         Width           =   1440
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   12
         Left            =   -72300
         TabIndex        =   160
         Top             =   4335
         Width           =   900
         VariousPropertyBits=   27
         Caption         =   "Label2"
         Size            =   "1587;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FMP 管制人"
         Height          =   180
         Index           =   83
         Left            =   -74760
         TabIndex        =   159
         Top             =   4365
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展後使用宣誓年度"
         Height          =   180
         Index           =   82
         Left            =   -69648
         TabIndex        =   158
         Top             =   3276
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展是否可延期          (N:不可延期)"
         Height          =   180
         Index           =   81
         Left            =   -72468
         TabIndex        =   156
         Top             =   2040
         Width           =   2712
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "央行國家代碼"
         Height          =   180
         Index           =   61
         Left            =   -70368
         TabIndex        =   155
         Top             =   996
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(公報地區名稱若有多個請以,區隔)"
         Height          =   180
         Index           =   76
         Left            =   -70176
         TabIndex        =   153
         Top             =   3540
         Width           =   2688
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   $"frm12040101.frx":2168
         Height          =   156
         Index           =   75
         Left            =   -74808
         TabIndex        =   152
         Top             =   4272
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "改繳費起始年度要同時改證書定稿內容！"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   74
         Left            =   -71160
         TabIndex        =   151
         Top             =   540
         Width           =   3240
      End
      Begin VB.Label Label1 
         Caption         =   "改為自動發證時，記得證書定稿要加證書費的文字段落！"
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   73
         Left            =   4200
         TabIndex        =   150
         Top             =   3885
         Width           =   3420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公報地區名稱"
         Height          =   180
         Index           =   72
         Left            =   -71172
         TabIndex        =   149
         Top             =   3756
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利公報地區代號"
         Height          =   180
         Index           =   71
         Left            =   -71172
         TabIndex        =   148
         Top             =   4056
         Width           =   1440
      End
      Begin MSForms.Label Label2 
         Height          =   252
         Index           =   9
         Left            =   -72084
         TabIndex        =   147
         Top             =   4008
         Width           =   864
         VariousPropertyBits=   27
         Caption         =   "Label2"
         Size            =   "1517;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CFT承辦人"
         Height          =   180
         Index           =   70
         Left            =   -74808
         TabIndex        =   146
         Top             =   4008
         Width           =   852
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國名(中)"
         Height          =   195
         Index           =   63
         Left            =   -74730
         TabIndex        =   145
         Top             =   540
         Width           =   690
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   8
         Left            =   -73830
         TabIndex        =   144
         Top             =   540
         Width           =   2565
         VariousPropertyBits=   27
         Size            =   "4524;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否英語系國家          (N:否)"
         Height          =   180
         Index           =   60
         Left            =   -69792
         TabIndex        =   143
         Top             =   1536
         Width           =   2172
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CFP發明是否准後繳年費                 (Y:准後繳年費)"
         Height          =   180
         Index           =   57
         Left            =   -74760
         TabIndex        =   142
         Top             =   3240
         Width           =   3870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CFP新型是否准後繳年費                 (Y:准後繳年費)"
         Height          =   180
         Index           =   58
         Left            =   -74760
         TabIndex        =   141
         Top             =   3480
         Width           =   3870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CFP設計是否准後繳年費                 (Y:准後繳年費)"
         Height          =   180
         Index           =   59
         Left            =   -74760
         TabIndex        =   140
         Top             =   3720
         Width           =   3870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCT承辦智權人員"
         Height          =   180
         Index           =   56
         Left            =   -74808
         TabIndex        =   139
         Top             =   3756
         Width           =   1392
      End
      Begin MSForms.Label Label2 
         Height          =   252
         Index           =   7
         Left            =   -72084
         TabIndex        =   138
         Top             =   3756
         Width           =   864
         VariousPropertyBits=   27
         Caption         =   "Label2"
         Size            =   "1517;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利設計是否自動發證                      (Y:自動發證)"
         Height          =   180
         Index           =   55
         Left            =   240
         TabIndex        =   137
         Top             =   4365
         Width           =   3795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利新型是否自動發證                      (Y:自動發證)"
         Height          =   180
         Index           =   54
         Left            =   240
         TabIndex        =   136
         Top             =   4128
         Width           =   3864
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "帳單請款幣別"
         Height          =   180
         Index           =   53
         Left            =   -70368
         TabIndex        =   134
         Top             =   516
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利發明是否自動發證                      (Y:自動發證)"
         Height          =   180
         Index           =   52
         Left            =   240
         TabIndex        =   133
         Top             =   3888
         Width           =   3864
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "商標是否自動發證                     (Y:自動發證)"
         Height          =   180
         Index           =   51
         Left            =   -74808
         TabIndex        =   132
         Top             =   3516
         Width           =   3456
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   6
         Left            =   -68385
         TabIndex        =   131
         Top             =   4095
         Width           =   900
         VariousPropertyBits=   27
         Caption         =   "Label2"
         Size            =   "1587;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP承辦智權人員"
         Height          =   180
         Index           =   50
         Left            =   -70680
         TabIndex        =   130
         Top             =   4125
         Width           =   1380
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   1140
         TabIndex        =   129
         Top             =   528
         Width           =   2565
         VariousPropertyBits=   27
         Size            =   "4524;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國名(中)"
         Height          =   180
         Index           =   43
         Left            =   240
         TabIndex        =   128
         Top             =   528
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "設計專用期截止日起算日"
         Height          =   180
         Index           =   49
         Left            =   240
         TabIndex        =   127
         Top             =   3285
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "設計專用期起始日起算日"
         Height          =   180
         Index           =   48
         Left            =   240
         TabIndex        =   126
         Top             =   3045
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新型專用期截止日起算日"
         Height          =   180
         Index           =   47
         Left            =   240
         TabIndex        =   125
         Top             =   2265
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新型專用期起始日起算日"
         Height          =   180
         Index           =   46
         Left            =   240
         TabIndex        =   124
         Top             =   2025
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發明專用期截止日起算日"
         Height          =   180
         Index           =   45
         Left            =   240
         TabIndex        =   123
         Top             =   1245
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發明專用期起始日起算日"
         Height          =   180
         Index           =   44
         Left            =   240
         TabIndex        =   122
         Top             =   1005
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公開發明月份"
         Height          =   180
         Index           =   24
         Left            =   4560
         TabIndex        =   121
         Top             =   1245
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公開新型起算日"
         Height          =   180
         Index           =   25
         Left            =   4560
         TabIndex        =   120
         Top             =   2115
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公開新型月份"
         Height          =   180
         Index           =   26
         Left            =   4560
         TabIndex        =   119
         Top             =   2355
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公開設計起算日"
         Height          =   180
         Index           =   27
         Left            =   4560
         TabIndex        =   118
         Top             =   3225
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公開設計月份"
         Height          =   180
         Index           =   28
         Left            =   4560
         TabIndex        =   117
         Top             =   3480
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公開發明起算日"
         Height          =   180
         Index           =   23
         Left            =   4560
         TabIndex        =   116
         Top             =   1005
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體審查發明起算日"
         Height          =   180
         Index           =   17
         Left            =   4560
         TabIndex        =   115
         Top             =   525
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體審查發明月份"
         Height          =   180
         Index           =   18
         Left            =   4560
         TabIndex        =   114
         Top             =   765
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體審查新型起算日"
         Height          =   180
         Index           =   19
         Left            =   4560
         TabIndex        =   113
         Top             =   1635
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體審查新型月份"
         Height          =   180
         Index           =   20
         Left            =   4560
         TabIndex        =   112
         Top             =   1875
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體審查設計起算日"
         Height          =   180
         Index           =   21
         Left            =   4560
         TabIndex        =   111
         Top             =   2745
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體審查設計月份"
         Height          =   180
         Index           =   22
         Left            =   4560
         TabIndex        =   110
         Top             =   2985
         Width           =   1440
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   -72300
         TabIndex        =   108
         Top             =   4065
         Width           =   900
         VariousPropertyBits=   27
         Caption         =   "Label2"
         Size            =   "1587;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         Caption         =   "(AXX:台灣 BXX:大陸 CXX:國外)"
         Height          =   252
         Left            =   -72192
         TabIndex        =   109
         Top             =   756
         Width           =   2772
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP 管制人"
         Height          =   180
         Index           =   33
         Left            =   -74760
         TabIndex        =   107
         Top             =   4110
         Width           =   885
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   -72240
         TabIndex        =   106
         Top             =   2340
         Width           =   1605
         VariousPropertyBits=   27
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   -72240
         TabIndex        =   105
         Top             =   1590
         Width           =   1605
         VariousPropertyBits=   27
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   -72240
         TabIndex        =   104
         Top             =   840
         Width           =   1605
         VariousPropertyBits=   27
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "使用宣誓年度"
         Height          =   180
         Index           =   38
         Left            =   -74808
         TabIndex        =   103
         Top             =   3276
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "下次使用宣誓年度"
         Height          =   180
         Index           =   37
         Left            =   -72108
         TabIndex        =   102
         Top             =   3276
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "未使用撤銷年度"
         Height          =   180
         Index           =   36
         Left            =   -74808
         TabIndex        =   101
         Top             =   3000
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國家代碼"
         Height          =   180
         Index           =   0
         Left            =   -74808
         TabIndex        =   100
         Top             =   528
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "區域別"
         Height          =   180
         Index           =   1
         Left            =   -74808
         TabIndex        =   99
         Top             =   756
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國名(中)"
         Height          =   180
         Index           =   2
         Left            =   -74808
         TabIndex        =   98
         Top             =   996
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國名(英)"
         Height          =   180
         Index           =   3
         Left            =   -74808
         TabIndex        =   97
         Top             =   1236
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "使用語文"
         Height          =   180
         Index           =   4
         Left            =   -74808
         TabIndex        =   96
         Top             =   1476
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "商標專用年度起算日"
         Height          =   180
         Index           =   29
         Left            =   -74808
         TabIndex        =   95
         Top             =   1752
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "商標專用年度"
         Height          =   180
         Index           =   30
         Left            =   -72312
         TabIndex        =   94
         Top             =   1788
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展年度"
         Height          =   180
         Index           =   31
         Left            =   -74808
         TabIndex        =   93
         Top             =   2016
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展時間(月)"
         Height          =   180
         Index           =   32
         Left            =   -69588
         TabIndex        =   92
         Top             =   2076
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展文件"
         Height          =   180
         Index           =   34
         Left            =   -74808
         TabIndex        =   91
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展證書附件"
         Height          =   180
         Index           =   35
         Left            =   -74808
         TabIndex        =   90
         Top             =   2760
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "設計繳費年度"
         Height          =   180
         Index           =   16
         Left            =   -74760
         TabIndex        =   89
         Top             =   2580
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發明年費案件性質"
         Height          =   180
         Index           =   5
         Left            =   -74760
         TabIndex        =   88
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發明年費起算日"
         Height          =   180
         Index           =   6
         Left            =   -70500
         TabIndex        =   87
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發明繳費年度"
         Height          =   180
         Index           =   8
         Left            =   -74760
         TabIndex        =   86
         Top             =   1110
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新型年費案件性質"
         Height          =   180
         Index           =   9
         Left            =   -74760
         TabIndex        =   85
         Top             =   1590
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新型年費起算日"
         Height          =   180
         Index           =   10
         Left            =   -70500
         TabIndex        =   84
         Top             =   1590
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新型繳費年度"
         Height          =   180
         Index           =   12
         Left            =   -74760
         TabIndex        =   83
         Top             =   1830
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "設計年費案件性質"
         Height          =   180
         Index           =   13
         Left            =   -74760
         TabIndex        =   82
         Top             =   2340
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "設計年費起算日"
         Height          =   180
         Index           =   14
         Left            =   -70500
         TabIndex        =   81
         Top             =   2340
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發明專用年度"
         Height          =   180
         Index           =   7
         Left            =   240
         TabIndex        =   80
         Top             =   765
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新型專用年度"
         Height          =   180
         Index           =   11
         Left            =   240
         TabIndex        =   79
         Top             =   1785
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "設計專用年度"
         Height          =   180
         Index           =   15
         Left            =   240
         TabIndex        =   78
         Top             =   2805
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發明繳費年起算日"
         Height          =   180
         Index           =   39
         Left            =   -74760
         TabIndex        =   77
         Top             =   1320
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新型繳費年起算日"
         Height          =   180
         Index           =   40
         Left            =   -74760
         TabIndex        =   76
         Top             =   2070
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "設計繳費年起算日"
         Height          =   180
         Index           =   41
         Left            =   -74760
         TabIndex        =   75
         Top             =   2820
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國名(中)"
         Height          =   180
         Index           =   42
         Left            =   -74760
         TabIndex        =   74
         Top             =   540
         Width           =   660
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   -73860
         TabIndex        =   73
         Top             =   540
         Width           =   2565
         VariousPropertyBits=   27
         Size            =   "4524;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "起算日：(1.收文日 2.申請日 3.發文日 4.准駁日 5.公告日 6.發證日 7.公開日)"
      Height          =   180
      Left            =   960
      TabIndex        =   71
      Top             =   5550
      Width           =   5835
   End
End
Attribute VB_Name = "frm12040101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/15 改成Form2.0 ; Label2(index)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

'Modify by Morgan 2006/8/31
'Const iFieldTotal = 57
'Modify by Morgan 2008/8/27
'Const iFieldTotal = 61
'Const iFieldTotal = 67
'Modify by Sindy 2010/3/1
'Const iFieldTotal = 68
'Modify by Sindy 2012/3/1
'MODIFY BY SONIA 2013/5/24
'Modify By Sindy 2013/9/30
'Modified by Lydia 2014/11/24 73->75
'Modified by Lydia 2017/02/13 77->78
'Modified by Lydia 2017/09/12 78->79
'Modified by Morgan 2019/11/4 79->83
'Modified by Morgan 2019/11/4 83->85
'Modified by Lydia 2025/09/09 85=>88
Const iFieldTotal = 88

Dim RcMain As New ADODB.Recordset, cp As New ADODB.Recordset
Dim TmpField(0 To iFieldTotal) As String, ActionEdit As Integer
Dim Bmk As Variant

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim strTrigList As String 'Added by Lydia 2018/01/17 已彈過"更新XXX會觸發Trigger"訊息的欄位textbox

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2
         If m_bInsert Then
            Text1(0).SetFocus
            RcEdit 0
         End If
      Case vbKeyF3
         If m_bUpdate Then
            RcEdit 1
         End If
      Case vbKeyF5
         If m_bDelete Then
            RcEdit 2
         End If
      Case vbKeyF4
         If m_bQuery Then
            RcEdit 5
         End If
      Case vbKeyHome
         If m_bQuery Then
            If Not (ActionEdit = 0 Or ActionEdit = 1) Then
               ActionRc 0
            End If
         End If
      Case vbKeyPageUp
         If m_bQuery Then
            If Not (ActionEdit = 0 Or ActionEdit = 1) Then
               ActionRc 1
            End If
         End If
      Case vbKeyPageDown
         If m_bQuery Then
            If Not (ActionEdit = 0 Or ActionEdit = 1) Then
               ActionRc 2
            End If
         End If
      Case vbKeyEnd
         If m_bQuery Then
            If Not (ActionEdit = 0 Or ActionEdit = 1) Then
               ActionRc 3
            End If
         End If
      Case vbKeyF9
         Text1(0).Locked = False
         If Text1(0) = "" Then MsgBox "國家代碼不可為空值", vbInformation: Text1(0).SetFocus: Exit Sub
         'edit by nickc 2006/11/14
         If ActionEdit <> 2 Then
            If Text1(1) = "" Then MsgBox "區域別不可為空值", vbInformation: Text1(1).SetFocus: Exit Sub
            If Text1(2) = "" Then MsgBox "國名(中)不可為空值", vbInformation: Text1(2).SetFocus: Exit Sub
            '2006/12/26 CANCEL BY SONIA 因為046柬埔寨
            'If Val(Text1(37)) + Val(Text1(38)) > Val(Text1(30)) Then
            '   MsgBox "使用宣誓年度+下次使用宣誓年度必須<=商標專用年度", vbInformation
            '   Text1(37).SetFocus
            '   Exit Sub
            'End If
            '2006/12/26 END
         End If
         RcEdit 3
         RcMain.ReQuery
         RcMain.Find "na01='" & Text1(0) & "'", 0, adSearchForward, 1
      Case vbKeyF10
         Text1(0).Locked = False
         If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
            RcEdit 4
         End If
      Case vbKeyEscape
         Unload Me
         Set frm12040101 = Nothing
   End Select
End Sub
'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If ActionEdit <> 3 Then
            KeyAscii = 0
            Form_KeyDown vbKeyF9, 0
         End If
    End Select
End Sub
Private Sub Form_Load()
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040101", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040101", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040101", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040101", strFind, False)
   
   '910625 Sieg
   Label2(3).Caption = ""
   Label2(6).Caption = ""
   Label2(7).Caption = ""
   Label2(9).Caption = "" 'Add By Sindy 2010/3/1
   Combo1.AddItem ""
   strExc(0) = "SELECT A1Y01||'-'||A1Y02 FROM ACC1Y0"
   intI = 1
   'edit by nickc 2007/02/09 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Do While Not RsTemp.EOF
      Combo1.AddItem RsTemp.Fields(0)
      RsTemp.MoveNext
   Loop
   
   MoveFormToCenter Me
   cp.CursorLocation = adUseClient
   'Modify by Morgan 2008/8/27 +NA63~NA68
   'Modify by Sindy 2010/3/1 +NA69
   'Modify by Sindy 2012/3/1 +NA70,71
   'modify by sonia 2013/5/24 +NA72
   'Modify by Sindy 2013/9/30 +NA73,NA74
   'Modified by Lydia 2014/11/24 +NA75,NA76
   'Modify by Sindy 2015/3/31 +NA77
   'MODIFY BY SONIA 2015/12/4 +NA78
   'Modified by Lydia 2017/02/13 +NA79
   'Modified by Lydia 2017/09/13 +NA80
   'Modified by Morgan 2019/11/4 +NA82~85
   'Modified by Morgan 2021/11/19 +NA87,NA88
   'Modified by Lydia 2025/09/09 +NA81,NA86,DHL國家代號(NA89)
   strExc(0) = "SELECT NA01,NA02,NA03,NA04,NA05,NA20,NA06,NA07,NA21,NA22,NA08" & _
               ",NA09,NA23,NA24,NA10,NA11,NA25,NA26,NA27,NA28,NA29,NA30,NA31,NA32,NA33" & _
               ",NA34,NA35,NA36,NA37,NA12,NA13,NA14,NA15,NA16,NA17,NA18,NA19,NA38,NA39" & _
               ",NA42,NA45,NA48,NA40,NA41,NA43,NA44,NA46,NA47,NA49,NA50,NA51,NA52,NA53" & _
               ",NA54,NA55,NA56,NA57,NA58,NA59,NA60,NA61,NA62,NA63,NA64,NA65,NA66,NA67" & _
               ",NA68,NA69,NA70,NA71,NA72,NA73,NA74,NA75,NA76,NA77,NA78,NA79,NA80,NA81" & _
               ",NA82,NA83,NA84,NA85,NA86,NA87,NA88,NA89 FROM NATION ORDER BY NA01"
   RcMain.CursorType = adOpenDynamic
   RcMain.CursorLocation = adUseClient
   RcMain.LockType = adLockBatchOptimistic
   RcMain.Open strExc(0), cnnConnection
   If Not RcMain.BOF Then ActionRc 0
   TxtSitu True
   ActionEdit = 3
   SSTab1.Tab = 0 'Add By Sindy 2013/9/30
End Sub

Private Sub ActionRc(ByVal Sty As Integer)
   TxtLock 2
   If RcMain.EOF And RcMain.BOF Then MsgBox "資料庫內無資料 !", vbInformation: Exit Sub
   With RcMain
      Select Case Sty
         Case 0
            .MoveFirst
         Case 1
            .MovePrevious
            If .BOF Then
               Beep
               MsgBox "巳是第一筆了 ! ", vbInformation
               .MoveFirst
            End If
         Case 2
            .MoveNext
            If .EOF Then
               Beep
               MsgBox "巳是最後一筆了 ! ", vbInformation
               .MoveLast
            End If
         Case 3
            .MoveLast
      End Select
   End With
   SetTxtValue
End Sub

Private Sub SetTxtValue()
 Dim i As Integer, j As Integer
 Dim oLbl As Object 'Added by Lydia 2019/05/17
 Dim obj As Object 'Added by Morgan 2022/5/6
   
   'Modified by Lydia 2019/05/17
   'For i = 0 To 3
   '   Label2(i).Caption = ""
   'Next
   'Label2(6).Caption = ""
   'Label2(7).Caption = ""
   'Label2(9).Caption = "" 'Add By Sindy 2010/3/1
   For Each oLbl In Label2
        oLbl.Caption = ""
   Next
   'end 2019/05/17
   
   'Modified by Morgan 2022/5/6 改以物件索引讀資料(可相容於欄位保留情形)
   'For i = 0 To iFieldTotal
   For Each obj In Text1
      i = obj.Index
   'end 2022/5/6

      If IsNull(RcMain.Fields(i).Value) = False Then
         Text1(i).Text = RcMain.Fields(i).Value
         If i = 68 Then Text1(i).Tag = RcMain.Fields(i).Value 'Add By Sindy 2014/9/11 CFT承辦人
         'Added By Lydia 2016/01/21 FCP管制人和FCP承辦業務員
         'Modified by Lydia 2017/02/13 FMP管制人(+78)
         'Modify By Sindy 2019/4/17 + 國家名稱(+2)
         If i = 33 Or i = 50 Or i = 78 Or i = 2 Then
            Text1(i).Tag = RcMain.Fields(i).Value
         End If
         If i = 2 Then
            Label2(4) = Text1(i)
            Label2(5) = Text1(i)
            Label2(8) = Text1(i)
         End If
         If Text1(i).Text <> "" Then
            Select Case i
               Case 5
                  Label2(0).Caption = ChgType(0, Text1(i).Text)
               Case 9
                  Label2(1).Caption = ChgType(0, Text1(i).Text)
               Case 13
                  Label2(2).Caption = ChgType(0, Text1(i).Text)
               Case 33
                  Label2(3).Caption = ChgType(1, Text1(i).Text)
               Case 50
                  Label2(6).Caption = ChgType(1, Text1(i).Text)
               Case 54
                  Label2(7).Caption = ChgType(1, Text1(i).Text)
               'Add By Sindy 2010/3/1
               Case 68
                  Label2(9).Caption = ChgType(1, Text1(i).Text)
               '2010/3/1 End
               'Add By Sindy 2013/9/30
               Case 72
                  Label2(10).Caption = ChgType(1, Text1(i).Text)
               Case 73
                  Label2(11).Caption = ChgType(1, Text1(i).Text)
               '2013/9/30 END
               'Added by Lydia 2017/02/13
               Case 78
                  Label2(12).Caption = ChgType(1, Text1(i).Text)
            End Select
         End If
      End If
   Next
   If Text1(51) = "" Then
      Combo1.ListIndex = 0
   Else
      For i = 0 To Combo1.ListCount - 1
         Combo1.ListIndex = i
         If InStr(Combo1.Text, Text1(51)) > 0 Then
            Exit For
         End If
      Next
   End If
End Sub

Private Sub RcEdit(Situ As Integer)
 Dim i As Integer
 Dim oLbl As Object 'Added by Lydia 2019/05/17
 Dim obj As Object 'Added by Morgan 2022/5/6
 
   Select Case Situ
      Case 0 'add
         TxtSitu False
         ActionEdit = 0
         TxtLock 2
      Case 1 'modi
         TxtSitu False
         Text1(0).Locked = True
         ActionEdit = 1
         'Modified by Morgan 2022/5/6 改以物件索引對應資料欄位(可相容於欄位保留情形)
         'For i = 0 To iFieldTotal
         For Each obj In Text1
            i = obj.Index
         'end 2022/5/6
            TmpField(i) = Text1(i).Text
         Next
      Case 2 'delete
         If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            RcMain.Delete
            RcMain.UpdateBatch
            If RcMain.EOF = True Then
               ActionRc 1
            Else
               ActionRc 2
            End If
         End If
      Case 3 'update
         If ActionEdit = 0 Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            RcMain.AddNew
            If GetVal = False Then Exit Sub
            ActionRc 3
         ElseIf ActionEdit = 1 Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            If GetVal = False Then Exit Sub
            'Modify By Sindy 2019/4/17 + 國家名稱(+2)
            If Text1(2).Tag <> "" And Text1(2).Text <> Text1(2).Tag Then
               'Modified by Lydia 2025/02/24
               'MsgBox "請更新商標公告資料使用的【公報地區名稱】"
               MsgBox "請更新商標公告資料使用的【公報地區名稱】；" & vbCrLf & "請以語法更新NA81外商國名，例如ＸＸ商；"
            End If
            '2019/4/17 END
            'Add By Sindy 2014/9/11
            'CFT承辦人有異動時,更新CFT該國案件之下一程序未處理催審期限之智權人員
            'Mark by Lydia 2022/08/05 改寫成Trigger: NATION_BEFORE 和 SETSPECMAN_BEFORE
            'If Text1(68).Text <> "" And Text1(68).Text <> Text1(68).Tag Then
               'If Left(Trim(Text1(0)), 3) = "011" Then
               '   MsgBox "請更新【系統特殊設定】的CFT承辦人日本(中南高所、北所)管制人"
               ''Added by Lydia 2016/11/16
               'ElseIf Left(Trim(Text1(0)), 3) = "101" Or Left(Trim(Text1(0)), 3) = "239" Then
               '   'Modified by Lydia 2018/10/03
               '   'MsgBox "請更新【系統特殊設定】的CFT承辦人美國歐盟(南高所、北中所)管制人"
               '   MsgBox "請更新【系統特殊設定】的CFT承辦人美國歐盟(南高所、北所和中所)管制人"
               ''end 2016/11/16
               'Else
               '   If PUB_UpdNpCFT305Np10(Text1(68).Text, Left(Trim(Text1(0).Text), 3)) = False Then Exit Sub
               'End If
            'End If
            'end 2022/08/05
            Text1(68).Tag = Text1(68).Text
            '2014/9/11 END
            'Added by Lydia 2016/01/21
            Text1(33).Tag = Text1(33).Text
            Text1(50).Tag = Text1(50).Text
            Text1(78).Tag = Text1(78).Text 'Added by Lydia 2017/02/13 FMP管制人
            Text1(2).Tag = Text1(2).Text 'Add By Sindy 2020/7/14
            strTrigList = "" 'Added by Lydia 2018/01/17
         ElseIf ActionEdit = 2 Then
            RcMain.Find "NA01='" & Text1(0).Text & "'", 0, adSearchForward, 1
            If RcMain.EOF Then
               MsgBox "無此記錄之資料 !", vbCritical
               RcMain.Bookmark = Bmk
            End If
            SetTxtValue
         End If
         ActionEdit = 3
         TxtSitu True
      Case 4 'cancel
         TxtSitu True
         If ActionEdit = 0 Then
            ActionRc 3
         ElseIf ActionEdit = 1 Then
            'Modified by Morgan 2022/5/6 改以物件索引對應資料欄位(可相容於欄位保留情形)
            'For i = 0 To iFieldTotal
            For Each obj In Text1
               i = obj.Index
            'end 2022/5/6
               Text1(i).Text = TmpField(i)
            Next
         ElseIf ActionEdit = 2 Then
            RcMain.Bookmark = Bmk
            SetTxtValue
         End If
         ActionEdit = 3
      Case 5 'query
         Bmk = RcMain.Bookmark
         TxtSitu False
         TxtLock 2
         'Modified by Morgan 2022/5/6 改以物件索引對應資料欄位(可相容於欄位保留情形)
         'For i = 1 To iFieldTotal
         '   Text1(i).Locked = True
         For Each obj In Text1
            i = obj.Index
            If i > 0 Then Text1(i).Locked = True
            'end 2022/5/6
         Next
         'Added by Lydia 2019/05/17
         For Each oLbl In Label2
              oLbl.Caption = ""
         Next
         'end 2019/05/17
         ActionEdit = 2
         SSTab1.Tab = 0
         Text1(0).SetFocus
   End Select
End Sub

Private Function GetVal() As Boolean
 Dim i As Integer
 Dim obj As Object 'Added by Morgan 2022/5/6
 
On Error GoTo ErrHand
   'Modified by Morgan 2022/5/6 改以物件索引對應資料欄位(可相容於欄位保留情形)
   'For i = 0 To iFieldTotal
   For Each obj In Text1
      i = obj.Index
   'end 2022/5/6
      If Text1(i).Text <> "" Then
         RcMain.Fields(i).Value = Text1(i).Text
      Else
         RcMain.Fields(i).Value = Null
      End If
   Next
      
   If Combo1.Text = "" Then
      RcMain.Fields(51).Value = Null
   Else
      RcMain.Fields(51).Value = Left(Combo1.Text, InStr(Combo1.Text, "-") - 1)
   End If
   
   RcMain.UpdateBatch
   GetVal = True
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
 Dim txt As TextBox, i As Integer
   Select Case Lt
      Case 0
         For Each txt In frm12040101.Text1
            txt.Locked = True
         Next
         Combo1.Locked = True
      Case 1
         For Each txt In frm12040101.Text1
            txt.Locked = False
         Next
         Combo1.Locked = False
      Case 2
         For Each txt In frm12040101.Text1
            txt.Text = ""
         Next
         For i = 0 To 3
            Label2(i).Caption = ""
         Next
         Label2(6).Caption = ""
         Label2(7).Caption = ""
         Label2(9).Caption = "" 'Add By Sindy 2010/3/1
         Combo1.ListIndex = 0
   End Select
End Sub

Private Sub TxtSitu(ByVal TF As Boolean)
 Dim i As Integer, txt As TextBox
   If TF = True Then
      TxtLock 0
      ' 90.07.13 modify by louis (更新TOOLBAR按紐的狀態)
      'For i = 1 To 4
      '   TBar1.Buttons(i).Enabled = True
      '   TBar1.Buttons(i + 5).Enabled = True
      'Next
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
      If m_bQuery Then
         TBar1.Buttons(4).Enabled = True
      Else
         TBar1.Buttons(4).Enabled = False
      End If
      If m_bQuery Then
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
   Else
      TxtLock 1
      ' 90.07.13 modify by louis (更新TOOLBAR按紐的狀態)
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      'If m_bInsert Then
      '   TBar1.Buttons(1).Enabled = True
      'Else
      '   TBar1.Buttons(1).Enabled = False
      'End If
      'If m_bUpdate Then
      '   TBar1.Buttons(2).Enabled = True
      'Else
      '   TBar1.Buttons(2).Enabled = False
      'End If
      'If m_bDelete Then
      '   TBar1.Buttons(3).Enabled = True
      'Else
      '   TBar1.Buttons(3).Enabled = False
      'End If
      'If m_bQuery Then
      '   TBar1.Buttons(4).Enabled = True
      'Else
      '   TBar1.Buttons(4).Enabled = False
      'End If
      'If m_bQuery Then
      '   TBar1.Buttons(6).Enabled = True
      '   TBar1.Buttons(7).Enabled = True
      '   TBar1.Buttons(8).Enabled = True
      '   TBar1.Buttons(9).Enabled = True
      'Else
      '   TBar1.Buttons(6).Enabled = False
      '   TBar1.Buttons(7).Enabled = False
      '   TBar1.Buttons(8).Enabled = False
      '   TBar1.Buttons(9).Enabled = False
      'End If
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm12040101 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
         Text1(0).SetFocus
         RcEdit 0
      Case 2
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
         Text1(0).Locked = False
         If Text1(0) = "" Then MsgBox "國家代碼不可為空值", vbInformation: Exit Sub
         ' 90.10.15 modify by louis
         If ActionEdit <> 2 Then
            If Text1(1) = "" Then MsgBox "區域別不可為空值", vbInformation: Exit Sub
            If Text1(2) = "" Then MsgBox "國名(中)不可為空值", vbInformation: Exit Sub
            '2006/12/26 CANCEL BY SONIA 因為046柬埔寨
            'If Not (Val(Text1(37)) + Val(Text1(38)) <= Val(Text1(30))) Then
            '   MsgBox "使用宣誓年度+下次使用宣誓年度必須<=商標專用年度", vbInformation
            '   Text1(37).SetFocus
            '   Exit Sub
            'End If
            '2006/12/26 END
         End If
         RcEdit 3
         RcMain.ReQuery
         RcMain.Find "na01='" & Text1(0) & "'", 0, adSearchForward, 1
      Case 12
         Text1(0).Locked = False
          If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
          RcEdit 4
          End If
      Case 14
         Unload Me
         Set frm12040101 = Nothing
   End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   Select Case Index
      Case 2, 4, 34, 35, 69
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(Index).IMEMode = 1
         OpenIme
      Case Else
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(Index).IMEMode = 2
         CloseIme
   End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      'Modify By Sindy 2013/9/30 +, 72, 73
      'Modified by Lydia 2015/02/16 NA60為央行國家代碼(+59)
      'Modified by Lydia 2017/02/13 FMP管制人(+78)
      'Modified by Lydia 2025/09/09 +DHL國家代號(+88)
      Case 0, 1, 3, 33, 50, 51, 54, 59, 68, 70, 72, 73, 78, 88
         KeyAscii = UpperCase(KeyAscii)

      'Modify By Cheng 2002/10/01
'      Case 48, 49
      'Modified by Morgan 2021/11/19 +84,85
      Case 48, 49, 52, 53, 55, 56, 57, 85
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
         
      'Add by Morgan 2006/8/31
      'Modify by Morgan 2008/8/27 +62~67
      'Modified by Lydia 2014/11/24 +74,75
      'Modified by Lydia 2015/02/16 NA60為央行國家代碼(-59)
      Case 60, 61, 62, 63, 64, 65, 66, 67, 74, 75
         If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Add by Morgan 2009/4/16
      'Add By Sindy 2015/3/31 +76
      Case 58, 76
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Added by Morgan 2019/11/4 +80~83(NA82~NA85)
      Case 80, 81, 82, 83
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 78 And KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Added by Morgan 2021/12/16
      Case 84
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
   If Index = 0 And KeyAscii = 13 And ActionEdit = 2 Then
      RcEdit 3
   End If
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTmp As String, i As Integer

   If ActionEdit = 3 Then Exit Sub
   Select Case Index
      Case 33
         Label2(3).Caption = ""
      Case 50
         Label2(6).Caption = ""
      Case 54
         Label2(7).Caption = ""
      'Add By Sindy 2010/3/1
      Case 68
         Label2(9).Caption = ""
      '2010/3/1 End
      'Add By Sindy 2013/9/30
      Case 72
         Label2(10).Caption = ""
      Case 73
         Label2(11).Caption = ""
      '2013/9/30 END
      'Added by Lydia 2017/02/13
      Case 78
         Label2(12).Caption = ""
   End Select
   
   If Text1(Index).Text = "" Then Exit Sub
   Cancel = False
   Select Case Index
      Case 0
         ' 90.10.15 modify by louis
         If ActionEdit = 2 Then
         ElseIf ActionEdit <> 1 Then
            If cp.State = adStateOpen Then cp.Close
            strExc(0) = "select count(na01) from nation where na01='" & Text1(0) & "'"
            cp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
            If cp.Fields(0) <> "0" Then
               MsgBox "此國家代碼已存在", vbInformation
               Cancel = True
            End If
         End If
      Case 1
         If Not (Mid(Text1(1), 1, 1) >= "A" And Mid(Text1(1), 1, 1) <= "C") Then
            MsgBox "區域別輸入錯誤", vbInformation
            Cancel = True
         End If
      Case 2, 69
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(2).IMEMode = 2
         CloseIme
         Label2(4) = Text1(2)
         Label2(5) = Text1(2)
      Case 4
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(4).IMEMode = 4
         CloseIme
      Case 5, 9, 13
         If Index = 5 Then
            i = 0
         ElseIf Index = 9 Then
            i = 1
         ElseIf Index = 13 Then
            i = 2
         End If
         Label2(i).Caption = ChgType(0, Text1(Index).Text)
         If Label2(i).Caption = "" Then Cancel = True
      Case 33
         Label2(3).Caption = ChgType(1, Text1(Index).Text)
         If Label2(3).Caption = "" Then Cancel = True
         'Add By Lydia 2016/01/21
         'Modified by Lydai 2018/01/17 +判斷記錄
         'If Label2(3).Caption <> "" And Text1(33).Tag <> Text1(33).Text Then
         If Label2(3).Caption <> "" And Text1(33).Tag <> Text1(33).Text And (strTrigList = "" Or (strTrigList <> "" And InStr(strTrigList, "033") = 0)) Then
            MsgBox "更新FCP管制人會觸發Triggers因此會花些時間更新下一程序期限!!"
            strTrigList = strTrigList & "033," 'Added by Lydia 2018/01/17
         End If
         'end 2016/01/21
      Case 50
         Label2(6).Caption = ChgType(1, Text1(Index).Text)
         If Label2(6).Caption = "" Then Cancel = True
         'Add By Lydia 2016/01/21
          'Modified by Lydai 2018/01/17 +判斷記錄
         'If Label2(6).Caption <> "" And Text1(50).Tag <> Text1(50).Text Then
         If Label2(6).Caption <> "" And Text1(50).Tag <> Text1(50).Text And (strTrigList = "" Or (strTrigList <> "" And InStr(strTrigList, "050") = 0)) Then
            MsgBox "更新FCP承辦智權人員會觸發Triggers因此會花些時間更新下一程序期限!!"
            strTrigList = strTrigList & "050," 'Added by Lydia 2018/01/17
         End If
         'end 2016/01/21
      'Add By Sindy 2013/9/30
      Case 72
         Label2(10).Caption = ChgType(1, Text1(Index).Text)
         If Label2(10).Caption = "" Then Cancel = True
      Case 73
         Label2(11).Caption = ChgType(1, Text1(Index).Text)
         If Label2(11).Caption = "" Then Cancel = True
      '2013/9/30 END
      Case 34
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(34).IMEMode = 2
         CloseIme
      Case 35
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text1(35).IMEMode = 2
         CloseIme
      Case 6, 10, 14, 17, 19, 21, 23, 25, 27, 29, 39, 40, 41, 42, 43, 44, 45, 46, 47
         If Not (Text1(Index) >= "1" And Text1(Index) <= "7") Then
            MsgBox "起算日輸入錯誤", vbInformation
            Cancel = True
         End If
      Case 36
         If Val(Text1(36)) > Val(Text1(30)) Then
            MsgBox "未使用撤消年度必須小於商標專用年度", vbInformation
            Cancel = True
         End If
      Case 37
         If Val(Text1(37)) > Val(Text1(30)) Then
            MsgBox "使用宣誓年度必須小於商標專用年度", vbInformation
            Cancel = True
         End If
      Case 38
         If Val(Text1(38)) > Val(Text1(30)) Then
            MsgBox "下次使用宣誓年度必須小於商標專用年度", vbInformation
            Cancel = True
         End If
      'Add By Cheng 2003/08/27
      Case 54 'FCT承辦智權人員
         Label2(7).Caption = ChgType(1, Text1(Index).Text)
         If Label2(7).Caption = "" Then Cancel = True
      'Add By Sindy 2010/3/1
      Case 68 'CFT承辦人
         Label2(9).Caption = ChgType(1, Text1(Index).Text)
         If Label2(9).Caption = "" Then Cancel = True
         'Add By Sindy 2015/6/2
          'Modified by Lydai 2018/01/17 +判斷記錄
         'If Label2(9).Caption <> "" And Text1(68).Tag <> Text1(68).Text Then
         If Label2(9).Caption <> "" And Text1(68).Tag <> Text1(68).Text And (strTrigList = "" Or (strTrigList <> "" And InStr(strTrigList, "068") = 0)) Then
            'Modified by Lydia 2016/03/11 +提示訊息(非離職人員)
            'Modified by Lydia 2022/08/05 改提示
            'MsgBox "更新CFT承辦人會觸發Triggers因此會花些時間更新下一程序期限!!" & vbCrLf & _
                   "若原承辦人在職,則下一程序的NP10不變."
            MsgBox "更新CFT承辦人會觸發Triggers因此會花些時間更新下一程序期限!!"
            strTrigList = strTrigList & "068," 'Added by Lydia 2018/01/17
         End If
         '2015/6/2 END
      '2010/3/1 End
      'Add By Sindy 2012/3/1
      Case 70
         If ActionEdit = 2 Then '刪除
         ElseIf ActionEdit <= 1 Then
            If ActionEdit = 0 And Text1(0) = Text1(Index) Then
            Else
               If cp.State = adStateOpen Then cp.Close
               strExc(0) = "select count(na01) from nation where na01='" & Text1(Index) & "'"
               cp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
               If Val(cp.Fields(0)) = 0 Then
                  MsgBox "此專利公報地區代號不存在！", vbInformation
                  Cancel = True
               End If
            End If
         End If
      '2012/3/1 End
      'ADD BY SONIA 2013/5/24
      Case 71
         If Not CheckLengthIsOK(Text1(Index).Text, 40) Then
            Cancel = True
         End If
      '2013/5/24 END
      'add by sonia 2015/12/4
      Case 77
         If Val(Text1(77)) > Val(Text1(31)) Then
            MsgBox "延展後使用宣誓年度必須小於延展年度", vbInformation
            Cancel = True
         End If
      'end 2015/12/4
      
      'Added by Lydia 2017/02/13
      Case 78
         Label2(12).Caption = ChgType(1, Text1(Index).Text)
         If Label2(12).Caption = "" Then Cancel = True
         'Added By Lydia 2017/02/13
         'Modified by Lydai 2018/01/17 +判斷記錄
         'If Label2(12).Caption <> "" And Text1(78).Tag <> Text1(78).Text Then
         If Label2(12).Caption <> "" And Text1(78).Tag <> Text1(78).Text And (strTrigList = "" Or (strTrigList <> "" And InStr(strTrigList, "078") = 0)) Then
            MsgBox "更新FMP管制人會觸發Triggers因此會花些時間更新下一程序期限!!"
            strTrigList = strTrigList & "078," 'Added by Lydia 2018/01/17
         End If
         'end 2017/02/13
   End Select
   If Cancel Then TextInverse Text1(Index)
 End Sub

Private Function ChgType(ByVal Sty As Integer, ByVal txt As String) As String
 Dim strTmp As String
   If Sty = 0 Then
      'edit by nickc 2007/02/09 不用 dll 了
      'If objPublicData.GetCaseProperty("P", txt, strTmp, , False) = True Then
      If ClsPDGetCaseProperty("P", txt, strTmp, , False) = True Then
         ChgType = strTmp
      Else
      'edit by nickc 2007/02/09 不用 dll 了
      'If objPublicData.GetCaseProperty("CFP", txt, strTmp) = True Then
      If ClsPDGetCaseProperty("CFP", txt, strTmp) = True Then
         ChgType = strTmp
      Else
         ChgType = ""
      End If
      End If
   Else
      'edit by nickc 2007/02/09 不用 dll 了
      'If objPublicData.GetStaff(txt, strTmp) = True Then
      '2010/4/13 modify by sonia
      'If ClsPDGetStaff(txt, strTmp) = True Then
      If ClsPDGetStaffN(txt, strTmp) = True Then
         ChgType = strTmp
      Else
         ChgType = ""
      End If
   End If
End Function

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Text1
   If objTxt.Enabled = True Then
      Cancel = False
      Text1_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function
