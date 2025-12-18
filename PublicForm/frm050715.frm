VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050715 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶減免身分維護"
   ClientHeight    =   5748
   ClientLeft      =   4248
   ClientTop       =   3000
   ClientWidth     =   9924
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9924
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3780
      Top             =   1380
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
            Picture         =   "frm050715.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050715.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050715.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050715.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050715.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050715.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050715.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050715.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050715.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050715.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050715.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9924
      _ExtentX        =   17505
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5085
      Left            =   90
      TabIndex        =   14
      Top             =   660
      Width           =   9795
      _ExtentX        =   17272
      _ExtentY        =   8975
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm050715.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(55)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(154)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(155)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblNationName"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblAD10"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label8"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblCaseName"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "SSTab2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtAD(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtAD(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtAD(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtAD(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtAD(3)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtAD(7)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtAD(6)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtAD(5)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm050715.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdList"
      Tab(1).Control(1)=   "cmdQuery"
      Tab(1).Control(2)=   "txtFn(0)"
      Tab(1).Control(3)=   "txtFn(1)"
      Tab(1).Control(4)=   "txtFn(2)"
      Tab(1).Control(5)=   "txtFn(3)"
      Tab(1).Control(6)=   "Label1(4)"
      Tab(1).Control(7)=   "Line2"
      Tab(1).Control(8)=   "Label1(5)"
      Tab(1).Control(9)=   "Line3"
      Tab(1).ControlCount=   10
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   3888
         Left            =   -74928
         TabIndex        =   87
         Top             =   888
         Width           =   8208
         _ExtentX        =   14478
         _ExtentY        =   6858
         _Version        =   393216
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
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   400
         Left            =   -67770
         TabIndex        =   12
         Top             =   390
         Width           =   912
      End
      Begin VB.TextBox txtFn 
         Height          =   270
         Index           =   0
         Left            =   -74040
         MaxLength       =   8
         TabIndex        =   8
         Top             =   450
         Width           =   945
      End
      Begin VB.TextBox txtFn 
         Height          =   270
         Index           =   1
         Left            =   -72990
         MaxLength       =   8
         TabIndex        =   9
         Top             =   450
         Width           =   945
      End
      Begin VB.TextBox txtFn 
         Height          =   270
         Index           =   2
         Left            =   -70260
         MaxLength       =   3
         TabIndex        =   10
         Top             =   450
         Width           =   525
      End
      Begin VB.TextBox txtFn 
         Height          =   270
         Index           =   3
         Left            =   -69630
         MaxLength       =   3
         TabIndex        =   11
         Top             =   450
         Width           =   525
      End
      Begin VB.TextBox txtAD 
         Height          =   270
         Index           =   5
         Left            =   1875
         MaxLength       =   6
         TabIndex        =   5
         Top             =   2850
         Width           =   915
      End
      Begin VB.TextBox txtAD 
         Height          =   270
         Index           =   6
         Left            =   2865
         MaxLength       =   1
         TabIndex        =   6
         Top             =   2850
         Width           =   315
      End
      Begin VB.TextBox txtAD 
         Height          =   270
         Index           =   7
         Left            =   3315
         MaxLength       =   2
         TabIndex        =   7
         Top             =   2850
         Width           =   435
      End
      Begin VB.TextBox txtAD 
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   1185
         MaxLength       =   1
         TabIndex        =   3
         Top             =   2250
         Width           =   375
      End
      Begin VB.TextBox txtAD 
         Height          =   270
         Index           =   4
         Left            =   1245
         MaxLength       =   3
         TabIndex        =   4
         Top             =   2850
         Width           =   525
      End
      Begin VB.TextBox txtAD 
         Height          =   270
         Index           =   0
         Left            =   1185
         MaxLength       =   8
         TabIndex        =   0
         Top             =   570
         Width           =   945
      End
      Begin VB.TextBox txtAD 
         Height          =   270
         Index           =   1
         Left            =   1185
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1320
         Width           =   945
      End
      Begin VB.TextBox txtAD 
         Height          =   270
         Index           =   2
         Left            =   1185
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1650
         Width           =   375
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   5025
         Left            =   4590
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   5145
         _ExtentX        =   9081
         _ExtentY        =   8869
         _Version        =   393216
         Tabs            =   7
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "中小企業"
         TabPicture(0)   =   "frm050715.frx":212C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblNotice(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "chkAD15JP1(14)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkAD15JP1(13)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkAD15JP1(12)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkAD15JP1(11)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkAD15JP1(10)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "chkAD15JP1(9)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "chkAD15JP1(8)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "chkAD15JP1(7)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "chkAD15JP1(6)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "chkAD15JP1(5)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "chkAD15JP1(4)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "chkAD15JP1(3)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "chkAD15JP1(2)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "chkAD15JP1(1)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "獨資企業"
         TabPicture(1)   =   "frm050715.frx":2148
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkAD15JP2(6)"
         Tab(1).Control(1)=   "chkAD15JP2(7)"
         Tab(1).Control(2)=   "chkAD15JP2(5)"
         Tab(1).Control(3)=   "chkAD15JP2(4)"
         Tab(1).Control(4)=   "chkAD15JP2(3)"
         Tab(1).Control(5)=   "chkAD15JP2(2)"
         Tab(1).Control(6)=   "chkAD15JP2(1)"
         Tab(1).Control(7)=   "lblNotice(2)"
         Tab(1).Control(8)=   "lblNotice(0)"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "小企業"
         TabPicture(2)   =   "frm050715.frx":2164
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkAD15JP3(2)"
         Tab(2).Control(1)=   "chkAD15JP3(1)"
         Tab(2).Control(2)=   "lblNotice(4)"
         Tab(2).Control(3)=   "lblNotice(3)"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "新興企業"
         TabPicture(3)   =   "frm050715.frx":2180
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "chkAD15JP4(2)"
         Tab(3).Control(1)=   "chkAD15JP4(1)"
         Tab(3).Control(2)=   "lblNotice(6)"
         Tab(3).Control(3)=   "lblNotice(5)"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "大學"
         TabPicture(4)   =   "frm050715.frx":219C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "chkAD15JP5(1)"
         Tab(4).Control(1)=   "lblNotice(7)"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "個人"
         TabPicture(5)   =   "frm050715.frx":21B8
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "chkAD15JP6(3)"
         Tab(5).Control(1)=   "chkAD15JP6(2)"
         Tab(5).Control(2)=   "chkAD15JP6(1)"
         Tab(5).Control(3)=   "lblNotice(9)"
         Tab(5).Control(4)=   "lblNotice(8)"
         Tab(5).ControlCount=   5
         TabCaption(6)   =   "台灣"
         TabPicture(6)   =   "frm050715.frx":21D4
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Frame2"
         Tab(6).Control(1)=   "Frame1"
         Tab(6).ControlCount=   2
         Begin VB.Frame Frame2 
            Caption         =   "台灣專利中小企業符合減免之資格"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   -74820
            TabIndex        =   82
            Top             =   750
            Width           =   3885
            Begin VB.TextBox txtAD16 
               Alignment       =   1  '靠右對齊
               Enabled         =   0   'False
               Height          =   270
               Index           =   6
               Left            =   810
               MaxLength       =   3
               TabIndex        =   86
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox txtAD16 
               Alignment       =   1  '靠右對齊
               Enabled         =   0   'False
               Height          =   270
               Index           =   5
               Left            =   2280
               MaxLength       =   9
               TabIndex        =   84
               Top             =   450
               Width           =   1005
            End
            Begin VB.CheckBox chkAD15 
               Caption         =   "依法辦理公司登記或商業登記，實收資本額在新臺幣1億元以下：                        元"
               Height          =   555
               Index           =   5
               Left            =   120
               TabIndex        =   83
               Top             =   180
               Width           =   3540
            End
            Begin VB.CheckBox chkAD15 
               Caption         =   "經常僱用員工數未滿200人之事業：員工數          人"
               Height          =   555
               Index           =   6
               Left            =   120
               TabIndex        =   85
               Top             =   690
               Width           =   3390
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "台灣專利中小企業符合減免之資格(舊)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2625
            Left            =   -74820
            TabIndex        =   73
            Top             =   2220
            Visible         =   0   'False
            Width           =   3915
            Begin VB.TextBox txtAD16 
               Alignment       =   1  '靠右對齊
               Enabled         =   0   'False
               Height          =   270
               Index           =   1
               Left            =   2160
               MaxLength       =   8
               TabIndex        =   77
               Top             =   420
               Width           =   1005
            End
            Begin VB.CheckBox chkAD15 
               Caption         =   "製造業、營造業、礦業及土石採取業實收資本額八千萬以下：                        元"
               Height          =   375
               Index           =   1
               Left            =   90
               TabIndex        =   81
               Top             =   240
               Width           =   3390
            End
            Begin VB.TextBox txtAD16 
               Alignment       =   1  '靠右對齊
               Enabled         =   0   'False
               Height          =   270
               Index           =   2
               Left            =   1080
               MaxLength       =   9
               TabIndex        =   76
               Top             =   900
               Width           =   1005
            End
            Begin VB.TextBox txtAD16 
               Alignment       =   1  '靠右對齊
               Enabled         =   0   'False
               Height          =   270
               Index           =   3
               Left            =   2835
               MaxLength       =   3
               TabIndex        =   75
               Top             =   1560
               Width           =   375
            End
            Begin VB.TextBox txtAD16 
               Alignment       =   1  '靠右對齊
               Enabled         =   0   'False
               Height          =   270
               Index           =   4
               Left            =   1575
               MaxLength       =   2
               TabIndex        =   74
               Top             =   2250
               Width           =   375
            End
            Begin VB.CheckBox chkAD15 
               Caption         =   "前項除外之其他行業前一年營業額一億元以下：                        元"
               Height          =   375
               Index           =   2
               Left            =   90
               TabIndex        =   80
               Top             =   720
               Width           =   3390
            End
            Begin VB.CheckBox chkAD15 
               Caption         =   "我國前項除外之其他行業前一年營業額一億元以上者但經常僱用員工數未滿100人：員工數          人"
               Height          =   555
               Index           =   4
               Left            =   90
               TabIndex        =   78
               Top             =   1890
               Width           =   3390
            End
            Begin VB.CheckBox chkAD15 
               Caption         =   "我國製造業、營造業、礦業及土石採取業實收資本額新台幣八千萬以上但經常僱用員工數未滿200人：員工數          人"
               Height          =   765
               Index           =   3
               Left            =   90
               TabIndex        =   79
               Top             =   1110
               Width           =   3390
            End
         End
         Begin VB.CheckBox chkAD15JP6 
            Caption         =   "獨資企業房屋土地交易所得額及營利事業所得額合計未滿日幣290萬"
            Height          =   360
            Index           =   3
            Left            =   -74760
            TabIndex        =   71
            Top             =   2190
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP6 
            Caption         =   "年所得合計未滿日幣250萬"
            Height          =   240
            Index           =   2
            Left            =   -74760
            TabIndex        =   70
            Top             =   1920
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP6 
            Caption         =   "年所得合計未滿日幣150萬"
            Height          =   240
            Index           =   1
            Left            =   -74760
            TabIndex        =   68
            Top             =   870
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP5 
            Caption         =   "大學"
            Height          =   240
            Index           =   1
            Left            =   -74760
            TabIndex        =   66
            Top             =   870
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP4 
            Caption         =   "獨資企業：公司成立未滿10年"
            Height          =   240
            Index           =   2
            Left            =   -74760
            TabIndex        =   64
            Top             =   2250
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP4 
            Caption         =   "中小型企業：公司成立未滿10年且總資本額3億日圓以下"
            Height          =   405
            Index           =   1
            Left            =   -74760
            TabIndex        =   62
            Top             =   870
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP3 
            Caption         =   $"frm050715.frx":21F0
            Height          =   405
            Index           =   2
            Left            =   -74760
            TabIndex        =   60
            Top             =   1950
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP3 
            Caption         =   "一般企業：員工20人以下（貿易或服務業公司員工5人以下）"
            Height          =   405
            Index           =   1
            Left            =   -74760
            TabIndex        =   58
            Top             =   870
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP2 
            Caption         =   "軟體或資料處理業（員工300人以下）"
            Height          =   200
            Index           =   6
            Left            =   -74760
            TabIndex        =   55
            Top             =   2520
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP2 
            Caption         =   "旅館業（員工200人以下）"
            Height          =   200
            Index           =   7
            Left            =   -74760
            TabIndex        =   54
            Top             =   2730
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP2 
            Caption         =   "橡膠製造業（汽車和飛機輪胎、內胎和工業用皮帶的製造業除外）（員工900人以下）"
            Height          =   375
            Index           =   5
            Left            =   -74760
            TabIndex        =   53
            Top             =   2130
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP2 
            Caption         =   "零售業（員工50人以下）"
            Height          =   195
            Index           =   4
            Left            =   -74760
            TabIndex        =   52
            Top             =   1920
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP2 
            Caption         =   "服務業（員工100人以下）"
            Height          =   195
            Index           =   3
            Left            =   -74760
            TabIndex        =   51
            Top             =   1710
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP2 
            Caption         =   "批發業（員工100人以下）"
            Height          =   195
            Index           =   2
            Left            =   -74760
            TabIndex        =   50
            Top             =   1500
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP2 
            Caption         =   "製造業、建築業、運輸業（員工300人以下）"
            Height          =   195
            Index           =   1
            Left            =   -74760
            TabIndex        =   49
            Top             =   1290
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "製造業、建築業、運輸業（員工300人以下）"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   47
            Top             =   690
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "製造業、建築業、運輸業（總資本額3億日圓以下）"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   46
            Top             =   900
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "批發業（員工100人以下）"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   45
            Top             =   1110
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "批發業（總資本額1億日圓以下）"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   44
            Top             =   1320
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "服務業（員工100人以下）"
            Height          =   200
            Index           =   5
            Left            =   240
            TabIndex        =   43
            Top             =   1530
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "服務業（總資本額5,000萬日圓以下）"
            Height          =   200
            Index           =   6
            Left            =   240
            TabIndex        =   42
            Top             =   1740
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "零售業（員工50人以下）"
            Height          =   200
            Index           =   7
            Left            =   240
            TabIndex        =   41
            Top             =   1950
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "零售業（總資本額5,000萬日圓以下）"
            Height          =   200
            Index           =   8
            Left            =   240
            TabIndex        =   40
            Top             =   2160
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "橡膠製造業（汽車和飛機輪胎、內胎和工業用皮帶的製造業除外）（員工900人以下）"
            Height          =   375
            Index           =   9
            Left            =   240
            TabIndex        =   39
            Top             =   2370
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "橡膠製造業（汽車和飛機輪胎、內胎和工業用皮帶的製造業除外）（總資本額3億日圓以下）"
            Height          =   375
            Index           =   10
            Left            =   240
            TabIndex        =   38
            Top             =   2760
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "軟體或資料處理業（員工300人以下）"
            Height          =   200
            Index           =   11
            Left            =   240
            TabIndex        =   37
            Top             =   3150
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "軟體或資料處理業（總資本額3億日圓以下）"
            Height          =   200
            Index           =   12
            Left            =   240
            TabIndex        =   36
            Top             =   3360
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "旅館業（員工200人以下）"
            Height          =   200
            Index           =   13
            Left            =   240
            TabIndex        =   35
            Top             =   3570
            Width           =   4500
         End
         Begin VB.CheckBox chkAD15JP1 
            Caption         =   "旅館業（總資本額5,000萬日圓以下）"
            Height          =   200
            Index           =   14
            Left            =   240
            TabIndex        =   34
            Top             =   3780
            Width           =   4500
         End
         Begin VB.Label lblNotice 
            Caption         =   "減免額度：實體審查規費減免50%，年費（第1-10年）減免50%。"
            Height          =   435
            Index           =   9
            Left            =   -74760
            TabIndex        =   72
            Top             =   2640
            Width           =   4605
         End
         Begin VB.Label lblNotice 
            Caption         =   "減免額度：免繳實體審查規費及第1-3年年費，第4-10年年費減免50%。"
            Height          =   435
            Index           =   8
            Left            =   -74760
            TabIndex        =   69
            Top             =   1170
            Width           =   4605
         End
         Begin VB.Label lblNotice 
            Caption         =   $"frm050715.frx":222B
            Height          =   675
            Index           =   7
            Left            =   -74760
            TabIndex        =   67
            Top             =   1170
            Width           =   4605
         End
         Begin VB.Label lblNotice 
            Caption         =   "減免額度：實體審查規費減免66%，年費（第1-10年）規費減免66%。"
            Height          =   405
            Index           =   6
            Left            =   -74760
            TabIndex        =   65
            Top             =   2550
            Width           =   4605
         End
         Begin VB.Label lblNotice 
            Caption         =   $"frm050715.frx":229D
            Height          =   735
            Index           =   5
            Left            =   -74760
            TabIndex        =   63
            Top             =   1320
            Width           =   4605
         End
         Begin VB.Label lblNotice 
            Caption         =   "減免額度：實體審查規費減免66%，年費（第1-10年）規費減免66%。"
            Height          =   405
            Index           =   4
            Left            =   -74760
            TabIndex        =   61
            Top             =   2430
            Width           =   4605
         End
         Begin VB.Label lblNotice 
            Caption         =   $"frm050715.frx":2333
            Height          =   735
            Index           =   3
            Left            =   -74760
            TabIndex        =   59
            Top             =   1290
            Width           =   4605
         End
         Begin VB.Label lblNotice 
            Caption         =   "減免資格的行業別中小企業，惟僅限制各行業別的員工數，不審核資本額"
            Height          =   465
            Index           =   2
            Left            =   -74760
            TabIndex        =   57
            Top             =   780
            Width           =   4485
         End
         Begin VB.Label lblNotice 
            Caption         =   "減免額度：實體審查規費減免50%，年費（第1-10年）規費減免50%。"
            Height          =   465
            Index           =   0
            Left            =   -74760
            TabIndex        =   56
            Top             =   3060
            Width           =   4485
         End
         Begin VB.Label lblNotice 
            Caption         =   $"frm050715.frx":23C9
            Height          =   945
            Index           =   1
            Left            =   270
            TabIndex        =   48
            Top             =   4020
            Width           =   4545
         End
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Left            =   2175
         TabIndex        =   32
         Top             =   570
         Width           =   1395
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2461;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCaseName 
         Height          =   285
         Left            =   1245
         TabIndex        =   31
         Top             =   3240
         Width           =   4725
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "8334;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(Y/N)"
         Height          =   180
         Left            =   1665
         TabIndex        =   30
         Top             =   2310
         Width           =   405
      End
      Begin VB.Label lblAD10 
         AutoSize        =   -1  'True
         Caption         =   "(1：自然人  2：學校  3：中小企業)"
         Height          =   540
         Left            =   1665
         TabIndex        =   29
         Top             =   1680
         Width           =   2730
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   180
         Left            =   1185
         TabIndex        =   28
         Top             =   1080
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "個人"
         Height          =   180
         Left            =   1185
         TabIndex        =   27
         Top             =   870
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號:"
         Height          =   180
         Index           =   4
         Left            =   -74850
         TabIndex        =   26
         Top             =   480
         Width           =   765
      End
      Begin VB.Line Line2 
         X1              =   -73260
         X2              =   -72840
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請國家:"
         Height          =   180
         Index           =   5
         Left            =   -71430
         TabIndex        =   25
         Top             =   480
         Width           =   765
      End
      Begin VB.Line Line3 
         X1              =   -69960
         X2              =   -69390
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "減免證明"
         Height          =   180
         Index           =   3
         Left            =   195
         TabIndex        =   24
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱:"
         Height          =   180
         Index           =   2
         Left            =   195
         TabIndex        =   23
         Top             =   3270
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "減免身份:"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   22
         Top             =   1680
         Width           =   765
      End
      Begin VB.Line Line1 
         X1              =   1395
         X2              =   3525
         Y1              =   3000
         Y2              =   3000
      End
      Begin MSForms.Label lblNationName 
         Height          =   285
         Left            =   2175
         TabIndex        =   21
         Top             =   1320
         Width           =   1395
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2461;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請國家:"
         Height          =   180
         Index           =   155
         Left            =   195
         TabIndex        =   20
         Top             =   1380
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   154
         Left            =   195
         TabIndex        =   19
         Top             =   2910
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號:"
         Height          =   180
         Index           =   55
         Left            =   195
         TabIndex        =   18
         Top             =   630
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否可減免:"
         Height          =   180
         Index           =   6
         Left            =   195
         TabIndex        =   17
         Top             =   2280
         Width           =   945
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Left            =   195
         TabIndex        =   16
         Top             =   3570
         Width           =   3615
         VariousPropertyBits=   27
         Caption         =   "Create : "
         Size            =   "6376;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label4 
         Height          =   285
         Left            =   195
         TabIndex        =   15
         Top             =   3900
         Width           =   3615
         VariousPropertyBits=   27
         Caption         =   "Update : "
         Size            =   "6376;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
End
Attribute VB_Name = "frm050715"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo By Sindy 2021/12/07 Form2.0已修改
'Memo By Morgan 2012/12/12 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'Create by Morgan 2003/12/24
Option Explicit

'前次紀錄KEY
Dim lst_AD01 As String
'本次紀錄KEY
Dim cur_AD01 As String
'目前狀態
Dim iCurState As Integer
'使用者權限設定
Dim bolInsert As Boolean
Dim bolUpdate As Boolean
Dim bolDelete As Boolean
Dim bolSelect As Boolean
'列印控制
Dim PLeft(0 To 8) As Integer
Dim iPrint As Integer
Dim Page As Integer
Dim strSql As String
Dim bolIsSameDept As Boolean 'Add by Morgan 2010/6/28 使用者是否與建立者為同部門
Dim strCU15 As String 'Added by Morgan 2014/7/15


'檢查查詢條件
Private Function CheckQueryData() As Boolean

   Dim bolCancel As Boolean, i As Integer
   
   If txtFn(0).Text = "" Then
        MsgBox "請輸入客戶編號起!!!", vbExclamation + vbOKOnly
        txtFn(0).SetFocus
        Exit Function
   End If
   If txtFn(1).Text = "" Then
        MsgBox "請輸入客戶編號迄!!!", vbExclamation + vbOKOnly
        txtFn(1).SetFocus
        Exit Function
   End If
   
   For i = 0 To 3
      Call txtFn_Validate(i, bolCancel)
      If bolCancel = True Then
         txtFn(i).SetFocus
         Exit Function
      End If
   Next
   CheckQueryData = True
   
End Function

Private Sub InitGrid()

   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer

   arrGridHeadText = Array("", "客戶編號", "名　　稱", "個人/公司", "ID" _
                     , "申請國家", "減免", "身份", "減免證明存卷案號", "", "", "")

   arrGridHeadWidth = Array(200, 900, 1600, 800 _
                     , 900, 900, 500, 1300, 1500, 0, 0, 0)

   With GrdList
      .row = 0
      .Cols = UBound(arrGridHeadText) + 1
      For iCol = 0 To .Cols - 1
         .col = iCol
         .Text = arrGridHeadText(iCol)
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .CellAlignment = flexAlignCenterCenter
      Next
      .Rows = 1
   End With
   
   
End Sub

Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset)

   Dim iRow As Integer, iCol As Integer
   GrdList.Rows = 1
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      With GrdList
         .Rows = .Rows + 1
         iRow = .Rows - 1
         For iCol = 1 To GrdList.Cols - 1
            .TextMatrix(iRow, iCol) = "" & rsTmp.Fields(iCol - 1).Value
         Next iCol
      End With
      rsTmp.MoveNext
   Loop
   GrdList.FixedRows = 1 'Added by Lydia 2023/10/16
End Sub

Private Function QueryData() As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset
   Dim strCon As String
   
On Error GoTo ErrHand

   strCon = ""
   If txtFn(0) <> "" Then
      strCon = strCon & " AND ad01>='" & Mid(GetNewFagent2(txtFn(0)), 1, 8) & "' "
   End If
   If txtFn(1) <> "" Then
      strCon = strCon & " AND ad01<='" & Mid(GetNewFagent2(txtFn(1)), 1, 8) & "' "
   End If
   If txtFn(2) <> "" Then
      strCon = strCon & " AND ad02>='" & txtFn(2) & "'"
   End If
   If txtFn(3) <> "" Then
      strCon = strCon & " AND ad02<='" & txtFn(3) & "'"
   End If
   
   'Modify By Sindy 2012/5/24 decode(cu15,'0','個人','1','公司','')==>decode(cu15,'0','個人','1','公司','2','學校','3','特殊機構','')
   'Modified by Morgan 2019/4/23 +日本減免身分
   strSql = "select ad01,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(cu15,'0','個人','1','公司','2','學校','3','特殊機構',''),cu11,na03,decode(ad03,'Y','是','N','否',''),decode(ad02,'011',decode(ad10,'1','中小企業','2','獨資企業','3','小企業','4','新興企業','5','大學','6','個人'),decode(ad10,'1','自然人','2','學校','3','中小企業')),replace(ad11||'-'||ad12||'-'||ad13||'-'||ad14,'---',''),na01,ad03,ad10" & _
            " from ApplicantDiscount,customer,nation" & _
            " where  ad01=cu01(+) and '0'=cu02(+) and ad02=na01(+)  " & strCon & " ORDER BY ad01,ad02"
            
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   InitGrid
   If rsQuery.RecordCount > 0 Then
      QueryData = True
      Call UpdateGridList(rsQuery)
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:

   MsgBox Err.Description, vbCritical
            
End Function

Private Sub chkAD15_Click(Index As Integer)
   Dim oCheck As CheckBox
   
   If SSTab2.Enabled = True Then
      If chkAD15(Index).Value = 1 Then
         txtAD16(Index).Enabled = True
         If txtAD16(Index).Enabled Then txtAD16(Index).SetFocus
         For Each oCheck In chkAD15
            If oCheck.Index <> Index And oCheck.Value = 1 Then
               oCheck.Value = 0
            End If
         Next
      Else
         txtAD16(Index).Text = ""
         txtAD16(Index).Enabled = False
      End If
   End If
End Sub

Private Sub chkAD15JP1_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP1(Index).Value = 1 Then
      For Each oCheck In chkAD15JP1
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub chkAD15JP2_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP2(Index).Value = 1 Then
      For Each oCheck In chkAD15JP2
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub chkAD15JP3_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP3(Index).Value = 1 Then
      For Each oCheck In chkAD15JP3
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub chkAD15JP4_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP4(Index).Value = 1 Then
      For Each oCheck In chkAD15JP4
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub chkAD15JP5_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP5(Index).Value = 1 Then
      For Each oCheck In chkAD15JP5
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub chkAD15JP6_Click(Index As Integer)
   Dim oCheck As CheckBox
   If chkAD15JP6(Index).Value = 1 Then
      For Each oCheck In chkAD15JP6
         If oCheck.Index <> Index And oCheck.Value = 1 Then
            oCheck.Value = 0
         End If
      Next
   End If
End Sub

Private Sub cmdQuery_Click()
   
   If TxtValidate(1) = False Then Exit Sub
   '查詢
      GrdList.Rows = 1
      If CheckQueryData = True Then
         Screen.MousePointer = vbHourglass
         GrdList.MousePointer = flexHourglass
         If QueryData() = False Then
             MsgBox "無資料", vbOKOnly, "查詢資料"
             txtFn(0).SetFocus
         End If
         GrdList.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
      End If
     
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2
      '新增
         If SSTab1.Tab = 0 And TBar1.Buttons(1).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(1))
         End If
      Case vbKeyF3
      '修改
         If SSTab1.Tab = 0 And TBar1.Buttons(2).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(2))
         End If
      Case vbKeyF5
      '刪除
         If SSTab1.Tab = 0 And TBar1.Buttons(3).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(3))
         End If
      Case vbKeyF4
      '查詢
         If SSTab1.Tab = 0 And TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
         End If
      Case vbKeyHome
      '第一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(6).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(6))
         End If

      Case vbKeyPageUp
      '上一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(7).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(7))
         End If
      Case vbKeyPageDown
      '下一筆
         If SSTab1.Tab = 0 And TBar1.Buttons(8).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(8))
         End If
      Case vbKeyEnd
      '最後筆
         If SSTab1.Tab = 0 And TBar1.Buttons(9).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(9))
         End If
      Case vbKeyF9
      '存檔
         If SSTab1.Tab = 0 And TBar1.Buttons(11).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(11))
         End If
      Case vbKeyReturn
      '確定
         If SSTab1.Tab = 1 Then
            Call cmdQuery_Click
         ElseIf SSTab1.Tab = 0 And TBar1.Buttons(11).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(11))
         End If
      Case vbKeyF10
      '取消
         If SSTab1.Tab = 0 And TBar1.Buttons(12).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(12))
         End If
      Case vbKeyEscape
      '結束
        If TBar1.Buttons(14).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(14))
         End If
    End Select
End Sub

Private Sub Form_Load()
   
   MoveFormToCenter Me
   Me.Show
   setAuthority
   Call FormReset(0)
   Call InitGrid
   '預設為瀏覽
   If doQuery(6) = True Then
      iCurState = 0
   Else
      iCurState = 9
   End If
   Call SetToolBar(iCurState)
   Call SetInputs(iCurState)
   SSTab1.Tab = 0 'Added by Lydia 2023/10/16
End Sub
'使用者權限設定
Private Sub setAuthority()
      bolInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
      bolUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
      bolDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
      bolSelect = IsUserHasRightOfFunction(Me.Name, strFind, False)
End Sub
'檢查本所案號
Private Function CheckCaseNo() As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset
   
On Error GoTo ErrHnd

   CheckCaseNo = False
   
   strSql = "Select PA01 From Patent Where PA01='" & txtAD(4) & "' AND PA02='" & txtAD(5) & "' AND PA03='" & txtAD(6) & "' AND PA04='" & txtAD(7) & "'"
   strSql = strSql & " Union Select TM01 From Trademark Where TM01='" & txtAD(4) & "' AND TM02='" & txtAD(5) & "' AND TM03='" & txtAD(6) & "' AND TM04='" & txtAD(7) & "'"
   strSql = strSql & " Union Select LC01 From Lawcase Where LC01='" & txtAD(4) & "' AND LC02='" & txtAD(5) & "' AND LC03='" & txtAD(6) & "' AND LC04='" & txtAD(7) & "'"
   strSql = strSql & " Union Select HC01 From Hirecase Where HC01='" & txtAD(4) & "' AND HC02='" & txtAD(5) & "' AND HC03='" & txtAD(6) & "' AND HC04='" & txtAD(7) & "'"
   strSql = strSql & " Union Select SP01 From Servicepractice Where SP01='" & txtAD(4) & "' AND SP02='" & txtAD(5) & "' AND SP03='" & txtAD(6) & "' AND SP04='" & txtAD(7) & "'"
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      CheckCaseNo = True
   End If
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   Exit Function
   
ErrHnd:

   MsgBox Err.Description
   
End Function

Private Function TxtValidate(Optional ByVal iTab As Integer = 0) As Boolean

   Dim oText As TextBox, bolCancel As Boolean, arrText, oMaskEdBox As MaskEdBox
   Dim oCheck As CheckBox
   
   TxtValidate = False
    
   Select Case iTab
   Case 0
       For Each oText In txtAD
         If oText.Locked = False Then
            txtAD_Validate oText.Index, bolCancel
            If bolCancel = True Then
               oText.SetFocus
               TextInverse oText
               Exit For
            End If
         End If
      Next
      'Added by Morgan 2013/3/22
      '台灣符合減免資格的中小企業需勾選資格
      'If bolCancel = False And txtAD(1) = "000" And txtAD(2) = "3" And txtAD(3) = "Y" Then
      If bolCancel = False And SSTab2.Enabled = True Then
         If txtAD(1) = "000" Then
            bolCancel = True
            intI = 0
            For Each oCheck In chkAD15
               If oCheck.Value = 1 Then
                  'Modified by Morgan 2020/7/24
                  'If Val(txtAD16(oCheck.Index).Text) > 0 Then
                  '   bolCancel = False
                  'Else
                  '   MsgBox "請輸入資本額/員工數！", vbInformation
                  '   intI = 2
                  '   If txtAD16(oCheck.Index).Enabled = True Then
                  '      txtAD16(oCheck.Index).SetFocus
                  '   End If
                  'End If
                  intI = 1
                  bolCancel = False
                  txtAD16_Validate oCheck.Index, bolCancel
                  'end 2020/7/24
                  Exit For
               End If
            Next
            If bolCancel = True And intI = 0 Then
               '暫不強制
               'If txtAD(4) = "P" Then
               '   MsgBox "台灣符合減免資格的中小企業請勾選資格！", vbExclamation
               ''FCP不用--靜芳
               'Else
               '   bolCancel = False
               'End If
               bolCancel = False
            End If
         End If
      End If
      'end 2013/3/22
   Case 1
       For Each oText In txtFn
         If oText.Locked = False Then
            txtFn_Validate oText.Index, bolCancel
            If bolCancel = True Then
               oText.SetFocus
               TextInverse oText
               Exit For
            End If
         End If
      Next
   Case Else
   End Select
   If bolCancel = False Then TxtValidate = True
   
End Function

Private Function CheckConfirm() As Boolean
   
   CheckConfirm = False
   
   Select Case iCurState
      '1:新增;2:修改
      Case 1, 2
      
         If TxtValidate = False Then Exit Function
         

         If txtAD(0) = "" Then
            MsgBox "客戶編號不可空白！", vbCritical
            txtAD(0).SetFocus
            Call txtAD_GotFocus(0)
            Exit Function
         ElseIf txtAD(1) = "" Then
            MsgBox "申請國家不可空白！", vbCritical
            txtAD(1).SetFocus
            Call txtAD_GotFocus(1)
            Exit Function
         'add by nick 2004/07/16
         ElseIf txtAD(1) >= "001" And txtAD(1) <= "008" Then
            MsgBox "申請國家不可為 001 ～008！", vbCritical
            txtAD(1).SetFocus
            Call txtAD_GotFocus(1)
            Exit Function
         ElseIf Trim(txtAD(3)) = "" Then
            MsgBox "是否可減免不可空白！", vbCritical
            txtAD(3).SetFocus
            Call txtAD_GotFocus(3)
            Exit Function
         'edit by nick 2004/07/16 個人可以不打案號
         '非台灣不用輸，沒輸入身份也不用輸
         'ElseIf txtAD(3) = "Y" And (txtAD(4) = "" Or txtAD(5) = "" Or txtAD(6) = "" Or txtAD(7) = "") And Trim(Label5.Caption) <> "個人" Then
         'Modify by Morgan 2010/6/28 使用者為國外部程序不用
         'ElseIf txtAD(3) = "Y" And (txtAD(4) = "" Or txtAD(5) = "" Or txtAD(6) = "" Or txtAD(7) = "") And Trim(Label5.Caption) <> "個人" And txtAD(1).Text = "000" And Trim(txtAD(2).Text) <> "" Then
         ElseIf Pub_StrUserSt03 <> "F22" And txtAD(3) = "Y" And (txtAD(4) = "" Or txtAD(5) = "" Or txtAD(6) = "" Or txtAD(7) = "") And Trim(Label5.Caption) <> "個人" And txtAD(1).Text = "000" And Trim(txtAD(2).Text) <> "" Then
                MsgBox "可減免必須輸入案號！", vbCritical
               txtAD(4).SetFocus
               Call txtAD_GotFocus(4)
               Exit Function
         '沒有打本所案號第一碼
         ElseIf txtAD(4) = "" And (txtAD(5) <> "" Or txtAD(6) <> "" Or txtAD(7) <> "") Then
               MsgBox "本所案號錯誤！", vbCritical
               txtAD(4).SetFocus
               Call txtAD_GotFocus(4)
               Exit Function
         '有打本所案號第一碼
         ElseIf txtAD(4) <> "" Then
               If CheckCaseNo() = False Then
                  MsgBox "查無此本所案號!!!", vbExclamation + vbOKOnly
                  txtAD(4).SetFocus
                  Call txtAD_GotFocus(4)
                  Exit Function
               End If
         End If
         If iCurState = 1 And CheckDBData = False Then
                MsgBox "資料已存在不允許新增！", vbCritical
                txtAD(0).SetFocus
                Call txtAD_GotFocus(0)
                Exit Function
         End If
         If iCurState = 2 And CheckDBData = True Then
                MsgBox "資料不存在不允許修改！", vbCritical
                txtAD(2).SetFocus
                Call txtAD_GotFocus(2)
                Exit Function
         End If
      '查詢
      Case 4
         If txtAD(0) = "" Then
            MsgBox "客戶編號不可空白！", vbCritical
            txtAD(0).SetFocus
            Call txtAD_GotFocus(0)
            Exit Function
         'Added by Morgan 2019/9/5
         ElseIf txtAD(1) = "" Then
            MsgBox "申請國家不可空白！", vbCritical
            txtAD(1).SetFocus
            Call txtAD_GotFocus(1)
            Exit Function
         End If
   End Select
   CheckConfirm = True
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm050715 = Nothing
End Sub

Private Sub grdList_DblClick()

   Dim lRow As Long, lCurRow As Long, iCol As Integer

   lCurRow = GrdList.row
   '呼叫查詢
   If lCurRow > 0 Then
      If TBar1.Buttons(4).Enabled = True Then
         Call Tbar1_ButtonClick(TBar1.Buttons(4))
         If txtAD(1).Locked = False Then
            txtAD(0).Text = GrdList.TextMatrix(lCurRow, 1)
            txtAD(1).Text = GrdList.TextMatrix(lCurRow, 9)
            If TBar1.Buttons(11).Enabled = True Then
               Call Tbar1_ButtonClick(TBar1.Buttons(11))
            End If
         End If
      End If
   End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Select Case PreviousTab
      Case 0
         InitGrid
         cmdQuery.Default = True
         If iCurState = 0 Then txtFn(0).SetFocus
       Case 1
         cmdQuery.Default = False
   End Select
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
      '新增
         iCurState = 1
      Case 2
      '修改
         iCurState = 2
      Case 3
      '刪除
         If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            If DeleteData = True Then
               If doQuery(8, False) = True Then
                  iCurState = 0
               ElseIf doQuery(9) = True Then
                  iCurState = 0
               Else
                  cur_AD01 = ""
                  iCurState = 9
               End If
            End If
         End If
      Case 4
      '查詢
         iCurState = 4
      Case 6
      '第一筆
         Call doQuery(6)
      Case 7
      '上一筆
         Call doQuery(7)
      Case 8
      '下一筆
         Call doQuery(8)
      Case 9
      '最後筆
         Call doQuery(9)
      Case 11
      '確定
         If CheckConfirm = False Then Exit Sub
         'Added by Lydia 2021/06/01 檢查減免身份
         'Modified by Morgan 2021/7/12
         'If iCurState = 1 Or iCurState = 2 Then
         If (iCurState = 1 Or iCurState = 2) And SSTab2.Visible = True And SSTab2.TabVisible(6) = False Then
            intI = GetAD15
            If intI = 0 Then
                 MsgBox "請在右側勾選符合的減免身份！", vbCritical
                 Exit Sub
            End If
         End If
         'end 2021/06/01
         Select Case iCurState
            '新增
            Case 1
               If insertdata() = False Then
                  Exit Sub
               End If
            '查詢
            Case 4
               cur_AD01 = txtAD(0) & txtAD(1)
               
            '修改
            Case 2
               If UpdateData() = False Then
                  Exit Sub
               End If
               
         End Select
         '重新查詢
         If doQuery(4) = True Then
            Call SetToolBar(0)
            Call SetInputs
         Else
            If iCurState = 4 Then
               txtAD(1).SetFocus
               Call txtAD_GotFocus(1)
            End If
            Exit Sub
         End If
         iCurState = 0
      Case 12
      '取消
         Select Case iCurState
            
            '1:新增
            Case 1
               If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                  Exit Sub
               ElseIf cur_AD01 = "" Then
                  If doQuery(6) = True Then
                     iCurState = 0
                  Else
                     iCurState = 9
                  End If
               ElseIf doQuery(4) = True Then
                  iCurState = 0
               Else
                  Exit Sub
               End If
            '2:修改
            Case 2
               If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                  Exit Sub
               ElseIf doQuery(4) = True Then
                  iCurState = 0
               Else
                  Exit Sub
               End If
            '查詢
            Case 4
               cur_AD01 = lst_AD01
               If cur_AD01 = "" Then
                  If doQuery(6) = True Then
                     iCurState = 0
                  Else
                     iCurState = 9
                  End If
               ElseIf doQuery(4) = True Then
                  iCurState = 0
               Else
                  Exit Sub
               End If
         End Select
      Case 14
      '結束
         If iCurState = 2 Or iCurState = 1 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
               Unload Me
               Exit Sub
            End If
         Else
            Unload Me
            Exit Sub
         End If
         
   End Select
   Call SetToolBar(iCurState)
   Call SetInputs(iCurState)
   lst_AD01 = cur_AD01
   
End Sub
'清除畫面
Private Sub FormReset(Optional ByVal iTab As Integer = 0)

   Dim oText As TextBox
   
   Select Case iTab
   
      Case 0
      '頁籤0
         For Each oText In txtAD
            oText.Text = ""
         Next
         txtFn(0).Text = ""
         txtFn(1).Text = ""
         txtFn(2).Text = ""
         txtFn(3).Text = ""
         Label5.Caption = ""
         Label6.Caption = ""
         Label2.Caption = ""
         lblNationName.Caption = ""
         lblCaseName.Caption = ""
         Label3.Caption = "Create:"
         Label4.Caption = "Update:"
         'Add by Morgan 2010/6/28
         '國外部程序新增時預設
         If Pub_StrUserSt03 = "F22" And iCurState = 1 Then
            txtAD(1) = "000"
            lblNationName = GetPrjNationName(txtAD(1))
            txtAD(2) = "3"
            txtAD(3) = "Y"
         End If
         'end 2010/6/28
         
         'Added by Morgan 2013/3/22
         'Mopdified by Morgan 2019/4/10 加日本案減免資格，改寫函數
         'For Each oCheck In chkAD15
         '   oCheck.Value = 0
         '   If oCheck.Index < 4 Then
         '      txtAD16(oCheck.Index).Text = ""
         '      txtAD16(oCheck.Index).Enabled = False
         '   End If
         'Next
         ResetTab2
         'end 2019/4/10
         
      Case 1
      '頁籤1
      
   End Select
End Sub
'工具列控制
Private Sub SetToolBar(Optional ByVal iStatus As Integer)

   Dim i As Integer
   For i = 1 To 13
      TBar1.Buttons(i).Enabled = False
   Next
   TBar1.Buttons(14).Enabled = True
   
   Select Case iStatus
   
      Case 0
      '瀏覽
         If bolInsert Then
            TBar1.Buttons(1).Enabled = True
         End If
         If bolUpdate And bolIsSameDept Then
            TBar1.Buttons(2).Enabled = True
         End If
         If bolDelete And bolIsSameDept Then
            TBar1.Buttons(3).Enabled = True
         End If
         If bolSelect Then
            TBar1.Buttons(4).Enabled = True
         End If
         TBar1.Buttons(6).Enabled = True
         TBar1.Buttons(7).Enabled = True
         TBar1.Buttons(8).Enabled = True
         TBar1.Buttons(9).Enabled = True
         
      Case 1, 2, 4
      '1:新增  '2:修改  '4查詢
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
               
      Case 9
      '無資料
         If bolInsert Then
            TBar1.Buttons(1).Enabled = True
         End If
         
   End Select
   
End Sub
'設定文字框
Private Sub SetInputs(Optional ByVal iStatus As Integer = 0)

   Dim oText As TextBox, oLabel As LABEL, oMaskEdBox As MaskEdBox
   
   Select Case iStatus
      
      Case 0
      '瀏覽
         For Each oText In txtAD
            oText.Enabled = True
            oText.Locked = True
         Next
         SSTab2.Enabled = False 'Added by Morgan 2013/3/22
         txtAD(0).SetFocus
      Case 1
      '新增
         SSTab1.Tab = 0
         For Each oText In txtAD
            oText.Text = ""
            oText.Locked = False
            oText.Enabled = True
         Next
         Call FormReset(0)
         txtAD(0).SetFocus
         Call txtAD_GotFocus(0)
      Case 2
      '修改
         SSTab1.Tab = 0
         For Each oText In txtAD
            oText.Locked = False
            oText.Enabled = True
         Next
         txtAD(0).Locked = True
         txtAD(1).Locked = True
         txtAD(2).SetFocus
         Call txtAD_GotFocus(2)
         SetOption 'Added by Morgan 2013/3/22
      Case 4
      '查詢
         SSTab1.Tab = 0
         For Each oText In txtAD
            oText.Locked = False
            oText.Enabled = False
         Next
         Call FormReset(0)
         txtAD(0).Locked = False
         txtAD(0).Enabled = True
         txtAD(1).Locked = False
         txtAD(1).Enabled = True
         txtAD(0).SetFocus
         
      Case 9
      '無資料
         For Each oText In txtAD
            oText.Enabled = False
            oText.Locked = True
         Next
         Call FormReset(0)
   End Select
   
End Sub
'讀取資料
Private Function doQuery(ByVal iAct As Integer, Optional ByVal bolMsg As Boolean = True) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, stMessage As String
   
   rsQuery.MaxRecords = 2
   rsQuery.CursorLocation = adUseClient
   doQuery = False
   
   Select Case iAct
      Case 4
      '查詢
         strSql = "Select AD01||ad02 From ApplicantDiscount where AD01||ad02='" & cur_AD01 & "'"
         stMessage = "查無資料！"
   
      Case 6
      '第一筆
         strSql = "Select AD01||ad02 From ApplicantDiscount ORDER BY 1 ASC"
         stMessage = "無減免客戶！"
      Case 7
      '上一筆
         strSql = "Select AD01||ad02 From ApplicantDiscount where AD01||ad02<'" & cur_AD01 & "'" & _
            " ORDER BY 1 DESC"
         stMessage = "已是第一筆了！"

      Case 8
      '下一筆
         strSql = "Select AD01||ad02 From ApplicantDiscount where AD01||ad02>'" & cur_AD01 & "'" & _
            " ORDER BY 1 ASC"
         stMessage = "已是最後一筆了！"

      Case 9
      '最後筆
         strSql = "Select AD01||ad02 From ApplicantDiscount" & _
            " ORDER BY 1 DESC"
         stMessage = "無減免客戶！"
        
   End Select
   
On Error GoTo ErrHand

   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
         lst_AD01 = cur_AD01
         cur_AD01 = "" & rsQuery.Fields(0).Value
         If ReQuery() = True Then doQuery = True
   ElseIf bolMsg Then
      MsgBox stMessage, vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:

   MsgBox Err.Description, vbCritical
   
End Function

Private Sub txtAD_Change(Index As Integer)
   'Added by Morgan 2013/3/22
   If Index = 1 Or Index = 2 Or Index = 3 Then
      SetOption
   End If
End Sub

Private Sub txtAD_GotFocus(Index As Integer)
   If txtAD(Index).Locked = False Then
          TextInverse txtAD(Index)
            'edit by nickc 2007/07/11 切換輸入法改用API
            'txtAD(Index).IMEMode = 2
            CloseIme
   End If
End Sub
'完整資料查詢
Private Function ReQuery(Optional ByVal bolMsg As Boolean = True) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, intI As Integer
   
On Error GoTo ErrHand

   Screen.MousePointer = vbHourglass
   
   ReQuery = False
   'Modify By Sindy 2012/5/24 decode(cu15,'0','個人','1','公司','')==>decode(cu15,'0','個人','1','公司','2','學校','3','特殊機構','')
   'Modified by Morgan 2013/3/22 +ad15,ad16
   'Modified by Morgan 2014/7/15 +cu15
   strSql = "SELECT ad01,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),decode(cu15,'0','個人','1','公司','2','學校','3','特殊機構',''),cu11,na03,nvl(ad03,' ') ||' '|| decode(ad03,'1','是','2','否',''),nvl(ad10,' ')||' '||decode(ad10,'1','自然人','2','學校','3','中小企業',''),ad11||'-'||ad12||'-'||ad13||'-'||ad14,ad02,ad04,ad05,ad06,ad07,ad08,ad09,s1.st02 ad04Name,s2.st02 ad07Name,s1.st03 " & _
            ",ad15,ad16,cu15 from ApplicantDiscount, customer , nation,staff s1,staff s2 " & _
            " where  ad04=s1.st01(+) and ad07=s2.st01(+) and ad01=cu01(+) and '0'=cu02(+) and ad02=na01(+)  and ad01||ad02='" & cur_AD01 & "' ORDER BY ad01,ad02"
            
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      FormReset 'Added by Morgan 2013/3/22
      strCU15 = "" & rsQuery("cu15") 'Added by Morgan 2014/7/15
        txtAD(0) = CheckStr(rsQuery.Fields(0).Value)
'        txtAD(0).Locked = False
'        txtAD_Validate 0, False
'        txtAD(0).Locked = True
        Label2.Caption = CheckStr(rsQuery.Fields(1).Value)
        Label5.Caption = CheckStr(rsQuery.Fields(2).Value)
        Label6.Caption = CheckStr(rsQuery.Fields(3).Value)
        txtAD(1) = CheckStr(rsQuery.Fields(8).Value)
        lblNationName.Caption = CheckStr(rsQuery.Fields(4).Value)
        txtAD(2) = Mid(CheckStr(rsQuery.Fields(6).Value), 1, 1)
        txtAD(3) = CheckStr(rsQuery.Fields(5).Value)
        If rsQuery.Fields(7) <> "---" Then 'Added by Morgan 2019/6/14
            txtAD(4) = SystemNumber(CheckStr(rsQuery.Fields(7).Value), 1)
            txtAD(5) = SystemNumber(CheckStr(rsQuery.Fields(7).Value), 2)
            txtAD(6) = SystemNumber(CheckStr(rsQuery.Fields(7).Value), 3)
            txtAD(7) = SystemNumber(CheckStr(rsQuery.Fields(7).Value), 4)
        End If 'Added by Morgan 2019/6/14
        lblCaseName = GetPrjName(txtAD(4) & "-" & txtAD(5) & "-" & txtAD(6) & "-" & txtAD(7))
        Label3.Caption = "Create:    " & CheckStr(rsQuery.Fields("ad04Name").Value) & "      " & ChangeWStringToWDateString(CheckStr(rsQuery.Fields("ad05").Value)) & "      " & Format(CheckStr(rsQuery.Fields("ad06").Value), "##:##")
        Label4.Caption = "Update:   " & CheckStr(rsQuery.Fields("ad07Name").Value) & "      " & ChangeWStringToWDateString(CheckStr(rsQuery.Fields("ad08").Value)) & "      " & Format(CheckStr(rsQuery.Fields("ad09").Value), "##:##")
        
        'Added by Morgan 2013/3/22
        If Not IsNull(rsQuery.Fields("ad15")) Then
            'Added by Lydia 2021/06/01 排除未勾選AD15=0的問題
            If Val("" & rsQuery.Fields("ad15")) = 0 Then
                MsgBox "請在右側勾選符合的減免身份！", vbCritical
            Else
            'end 2021/06/01
                'Added by Morgan 2019/4/10 +日本
                If txtAD(1) = "011" Then
                   If txtAD(2) = "1" Then
                      chkAD15JP1(rsQuery.Fields("ad15")).Value = 1
                   ElseIf txtAD(2) = "2" Then
                      chkAD15JP2(rsQuery.Fields("ad15")).Value = 1
                   ElseIf txtAD(2) = "3" Then
                      chkAD15JP3(rsQuery.Fields("ad15")).Value = 1
                   ElseIf txtAD(2) = "4" Then
                      chkAD15JP4(rsQuery.Fields("ad15")).Value = 1
                   ElseIf txtAD(2) = "5" Then
                      chkAD15JP5(rsQuery.Fields("ad15")).Value = 1
                   ElseIf txtAD(2) = "6" Then
                      chkAD15JP6(rsQuery.Fields("ad15")).Value = 1
                   End If
                Else
                'end 2019/4/10
                   chkAD15(rsQuery.Fields("ad15")).Value = 1
                   txtAD16(rsQuery.Fields("ad15")).Text = rsQuery.Fields("ad16")
                End If 'Added by Morgan 2019/4/10
            End If 'Added by Lydia 2021/06/01
        End If
        'end 2013/3/22
        
        ReQuery = True
        
      'Add by Morgan 2010/6/28
      'Modified by Morgan 2012/10/1 分國內外部門就好
      'If Pub_StrUserSt03 = "M51" Or rsQuery.Fields("st03") = Pub_StrUserSt03 Then
      If Pub_StrUserSt03 = "M51" Or (Left(Pub_StrUserSt03, 1) <> "F" And Left(rsQuery.Fields("st03"), 1) <> "F") Or (Left(Pub_StrUserSt03, 1) = "F" And Left(rsQuery.Fields("st03"), 1) = "F") Then
         bolIsSameDept = True
      Else
         bolIsSameDept = False
      End If
      'end 2010/6/28
   ElseIf bolMsg Then
        MsgBox "客戶編號及申請國家〔" & cur_AD01 & "〕已被刪除！", vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Screen.MousePointer = vbDefault
   
   Exit Function
   
ErrHand:
   MsgBox Err.Description, vbCritical
   Screen.MousePointer = vbDefault
   
End Function

Private Sub txtAD_KeyPress(Index As Integer, KeyAscii As Integer)
   If txtAD(Index).Locked = False And KeyAscii <> 8 Then
      KeyAscii = UpperCase(KeyAscii)
      Select Case Index
      Case 1 '申請國家
         If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
         Else
            txtAD(2) = ""
         End If
      Case 2 '減免身份
         If txtAD(1) = "011" Then
            If Chr(KeyAscii) < "1" Or Chr(KeyAscii) > "6" Then
               KeyAscii = 0
            End If
         Else
            If Chr(KeyAscii) < "1" Or Chr(KeyAscii) > "3" Then
               KeyAscii = 0
            End If
         End If
      Case 3 '是否可減免
         If Chr(KeyAscii) <> "Y" And Chr(KeyAscii) <> "N" Then
            KeyAscii = 0
         'Added by Morgan 2019/4/12
         ElseIf Chr(KeyAscii) = "N" Then
            txtAD(2) = ""
         End If
      Case 4 '系統別
         If Chr(KeyAscii) < "A" Or Chr(KeyAscii) > "Z" Then
            KeyAscii = 0
         End If
      '本所案號
      Case 5, 6, 7
         If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
         End If
      End Select
   End If
End Sub

Private Sub txtAD_LostFocus(Index As Integer)
   If SSTab1.Tab = 1 Then
      Dim bolCancel As Boolean
      bolCancel = False
      Call txtAD_Validate(Index, bolCancel)
      If bolCancel = True Then
         SSTab1.Tab = 0
         txtAD(Index).SetFocus
      End If
   
'Removed by Morgan 2022/9/13 統一在 txtAD_Validate 檢查 (因點存檔時駐點所在的欄位不會觸發LostFocus事件)
'   Else
'            If txtAD(Index).Locked = False Then
'      Select Case Index
'         Case 0
'            If txtAD(Index) <> "" Then
'               txtAD(Index).Text = Mid(GetNewFagent2(txtAD(Index).Text), 1, 8)
'               CheckOC
'               'Modify By Sindy 2012/5/24 decode(cu15,'0','個人','1','公司','')==>decode(cu15,'0','個人','1','公司','2','學校','3','特殊機構','')
'               strSql = "select decode(cu15,'0','個人','1','公司','2','學校','3','特殊機構',''),cu11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),cu15 from customer where cu01='" & txtAD(Index).Text & "' and cu02='0' "
'               adoRecordset.CursorLocation = adUseClient
'               adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'               If adoRecordset.RecordCount <> 0 Then
'                  strCU15 = "" & adoRecordset("cu15") 'Added by Morgan 2014/7/15
'                       Label5.Caption = CheckStr(adoRecordset.Fields(0).Value)
'                       Label6.Caption = CheckStr(adoRecordset.Fields(1).Value)
'                       Label2.Caption = CheckStr(adoRecordset.Fields(2).Value)
'                       'add by nick 2004/07/16 若是個人，減免帶1是否可減免為 Y
'                       If Trim(Label5.Caption) = "個人" Then
'                             txtAD(2).Text = "1"
'                             txtAD(3).Text = "Y"
'                       End If
'               Else
'                        Label5.Caption = ""
'                        Label6.Caption = ""
'                        Label2.Caption = ""
'                        MsgBox "查無此客戶！", vbCritical
'               End If
'               CheckOC
'            End If
'         Case 1
'            If txtAD(Index) <> "" Then
'                If txtAD(Index) >= "001" And txtAD(Index) <= "008" Then
'                    lblNationName.Caption = ""
'                    MsgBox "申請國家不可以 001 ～008！", vbCritical
'                Else
'                        CheckOC
'                        strSql = "select na03 from nation where na01='" & txtAD(Index).Text & "' "
'                        adoRecordset.CursorLocation = adUseClient
'                        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                        If adoRecordset.RecordCount <> 0 Then
'                                lblNationName.Caption = CheckStr(adoRecordset.Fields(0).Value)
'                        Else
'                                 lblNationName.Caption = ""
'                                 MsgBox "查無此國家！", vbCritical
'                        End If
'                        CheckOC
'               End If
'            End If
'        Case 7
'                    If txtAD(4) <> "" And Index = 7 Then
'                        If lblCaseName.Caption = "" Then
'                             MsgBox "無此本所案號！", vbCritical
'                             txtAD(4).SetFocus
'                             Call txtAD_GotFocus(4)
'                        End If
'                    End If
'        Case Else
'        End Select
'        End If
   End If
End Sub

Private Sub txtAD_Validate(Index As Integer, Cancel As Boolean)

   If txtAD(Index).Locked = False And txtAD(Index).Enabled = True Then
      Select Case Index
         'Added by Morgan 2022/9/13 從 txtAD_LostFocus 移來
         Case 0
            If txtAD(Index) <> "" Then
               txtAD(Index).Text = Mid(GetNewFagent2(txtAD(Index).Text), 1, 8)
               CheckOC
               'Modify By Sindy 2012/5/24 decode(cu15,'0','個人','1','公司','')==>decode(cu15,'0','個人','1','公司','2','學校','3','特殊機構','')
               strSql = "select decode(cu15,'0','個人','1','公司','2','學校','3','特殊機構',''),cu11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),cu15 from customer where cu01='" & txtAD(Index).Text & "' and cu02='0' "
               adoRecordset.CursorLocation = adUseClient
               adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset.RecordCount <> 0 Then
                  strCU15 = "" & adoRecordset("cu15") 'Added by Morgan 2014/7/15
                       Label5.Caption = CheckStr(adoRecordset.Fields(0).Value)
                       Label6.Caption = CheckStr(adoRecordset.Fields(1).Value)
                       Label2.Caption = CheckStr(adoRecordset.Fields(2).Value)
                       'add by nick 2004/07/16 若是個人，減免帶1是否可減免為 Y
                       'Modified by Lydia 2023/06/09 排除新增和修改
                       If Trim(Label5.Caption) = "個人" And iCurState <> 1 And iCurState <> 2 Then
                             txtAD(2).Text = "1"
                             txtAD(3).Text = "Y"
                       End If
               Else
                        Label5.Caption = ""
                        Label6.Caption = ""
                        Label2.Caption = ""
                        MsgBox "查無此客戶！", vbCritical
               End If
               CheckOC
            End If
         'end 2022/9/13
         Case 1
            If txtAD(Index) <> "" Then
                If txtAD(Index) >= "001" And txtAD(Index) <= "008" Then
                    lblNationName.Caption = ""
                    MsgBox "申請國家不可以 001 ～008！", vbCritical
                    txtAD_GotFocus (Index)
                    Cancel = True
               End If
            End If
       Case 2
            
            If Trim(txtAD(Index).Text) = "" Then
               If txtAD(1).Text = "000" Or txtAD(1).Text = "011" Then 'Added by Morgan 2024/4/2 非台灣非日本時不預設 Ex:X51103000美國
                  txtAD(3).Text = "N"
                  Exit Sub 'Added by Morgan 2019/6/14
               End If
            Else
                txtAD(3).Text = "Y"
            End If
            
            
            'add by nick 2004/09/03 不是台灣的不管
            If txtAD(1).Text = "000" Then
               If iCurState <> 4 Then
                  If txtAD(Index).Text = "1" And strCU15 <> "0" Then
                     MsgBox "減免身份為自然人時只能是個人！", vbCritical
                     txtAD_GotFocus (Index)
                     Cancel = True
                     Exit Sub
                  ElseIf txtAD(Index).Text = "2" And strCU15 <> "2" Then
                     MsgBox "減免身份為學校時只能是學校！", vbCritical
                     txtAD_GotFocus (Index)
                     Cancel = True
                     Exit Sub
                  ElseIf txtAD(Index).Text = "3" And (strCU15 = "0" Or strCU15 = "2") Then
                     MsgBox "減免身份為中小企業時不能是個人或學校！", vbCritical
                     txtAD_GotFocus (Index)
                     Cancel = True
                     Exit Sub
                  End If
               End If
            'Added by Morgan 2019/4/10
            ElseIf txtAD(1).Text = "011" Then
               If iCurState <> 4 Then
                  If txtAD(Index).Text = "6" And strCU15 <> "0" Then
                     MsgBox "減免身份為個人時只能是個人！", vbCritical
                     txtAD_GotFocus (Index)
                     Cancel = True
                     Exit Sub
                     
                  ElseIf txtAD(Index).Text = "5" And Not (strCU15 = "0" Or strCU15 = "2") Then
                     MsgBox "減免身份為大學只能是學校或個人！", vbCritical
                     txtAD_GotFocus (Index)
                     Cancel = True
                     Exit Sub
                  ElseIf txtAD(Index).Text < "5" And (strCU15 = "0" Or strCU15 = "2") Then
                     MsgBox "減免身份為企業時不能是個人或學校！", vbCritical
                     txtAD_GotFocus (Index)
                     Cancel = True
                     Exit Sub
                  End If
               End If
            End If
         Case 3
            If InStr(1, "YN", UCase(txtAD(Index).Text)) = 0 Then
                    MsgBox "是否可減免僅可輸入 Y 或 N ！", vbCritical
                    txtAD_GotFocus (Index)
                    Cancel = True
            Else
                'edit by nick 2004/09/03
                'If UCase(txtAD(Index).Text) = "N" And Trim(txtAD(2).Text) <> "" Then
                If UCase(txtAD(Index).Text) = "N" And Trim(txtAD(2).Text) <> "" And txtAD(1).Text = "000" Then
                    MsgBox "若不符合減免，不可輸入減免身份！", vbCritical
                     txtAD_GotFocus (Index)
                    Cancel = True
                Else
                    'edit by nick 2004/09/03
                    'If UCase(txtAD(Index).Text) = "Y" And Trim(txtAD(2).Text) = "" Then
                    If UCase(txtAD(Index).Text) = "Y" And Trim(txtAD(2).Text) = "" And txtAD(1).Text = "000" Then
                        MsgBox "若符合減免，應要輸入減免身份！", vbCritical
                        Cancel = True
                    Else
                        Cancel = False
                    End If
                End If
            End If
         Case 4
         '本所案號
            txtAD(Index) = Trim(txtAD(Index))
            If CheckSysKind(txtAD(4)) = False Then
               MsgBox "系統代碼輸入錯誤！", vbCritical
               txtAD_GotFocus (Index)
               Cancel = True
            End If
         Case 5
         '本所案號
            If txtAD(4) <> "" Then
               txtAD(Index) = UCase(Right("000000" & txtAD(Index).Text, 6))
            End If
         Case 6
         '本所案號
            If txtAD(4) <> "" Then
               txtAD(Index) = UCase(Right("0" & txtAD(Index).Text, 1))
            End If
         Case 7
         '本所案號
            If txtAD(4) <> "" Then
               txtAD(Index) = UCase(Right("00" & txtAD(Index).Text, 2))
               lblCaseName = GetPrjName(txtAD(4) & "-" & txtAD(5) & "-" & txtAD(6) & "-" & txtAD(7))
            End If
      End Select
   End If
   If Cancel = True Then
      Call txtAD_GotFocus(Index)
   End If
End Sub

Private Function CheckSysKind(ByVal stSys As String, Optional ByVal bolMsg As Boolean) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, stMessage As String
   Dim i As Integer, arrSys() As String
   
On Error GoTo ErrHand

   CheckSysKind = False
   strSql = "Select SK01 from systemkind"
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly

   If rsQuery.RecordCount > 0 Then
      arrSys = Split(rsQuery.GetString, Chr(13))
      For i = 0 To UBound(arrSys)
         If arrSys(i) = stSys Then
            CheckSysKind = True
            Exit For
         End If
      Next i
   Else
      If bolMsg Then
         MsgBox "無法取得系統代碼！", vbCritical
      End If
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:

   MsgBox Err.Description, vbCritical
   
End Function

Private Function DeleteData() As Boolean
   Dim strSql As String, lngEffRec As Long
   
   strSql = "Delete ApplicantDiscount Where AD01||ad02='" & cur_AD01 & "'"
   
   DeleteData = False
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   cnnConnection.Execute strSql, lngEffRec
   cnnConnection.CommitTrans
   DeleteData = True
   
   Exit Function
   
ErrHnd:

   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Function UpdateData() As Boolean

   Dim strSql As String, intI As Integer, strSNo As String, OG(3 To 11) As String
   Dim rsQuery As New ADODB.Recordset, strUpdSQL As String, lngEffRec As Long
   Dim oCheck As CheckBox
   
   OG(3) = "ad10='" & txtAD(2).Text & "'"
   OG(4) = "ad03='" & txtAD(3).Text & "'"
   OG(5) = "ad11='" & txtAD(4).Text & "'"
   OG(6) = "ad12='" & txtAD(5).Text & "'"
   OG(7) = "ad13='" & txtAD(6).Text & "'"
   OG(8) = "ad14='" & txtAD(7).Text & "'"
   OG(9) = "ad07='" & strUserNum & "'"
   OG(10) = "ad08=TO_NUMBER(TO_CHAR(SYSDATE,'YYYYMMDD'))"
   OG(11) = "ad09=TO_NUMBER(TO_CHAR(SYSDATE,'HH24MI'))"
   
   strUpdSQL = Join(OG, ",")
   
   'Added by Morgan 2013/3/22
   'Modified by Morgan 2019/4/10 加日本案減免資格
   If SSTab2.Enabled = True Then
      '台灣
      If SSTab2.TabVisible(6) = True Then
         'Modified by Morgan 2020/7/24
         'For intI = 1 To 4
         For intI = 1 To 6
            If chkAD15(intI) = 1 Then
               strUpdSQL = strUpdSQL & ",ad15='" & intI & "'"
               strUpdSQL = strUpdSQL & ",ad16=" & Val(Format(txtAD16(intI).Text))
               Exit For
            End If
         Next
         If intI > 6 Then
            strUpdSQL = strUpdSQL & ",ad15=null,ad16=null"
         End If
      '日本
      Else
         strUpdSQL = strUpdSQL & ",ad15='" & GetAD15() & "'"
      End If
   Else
      strUpdSQL = strUpdSQL & ",ad15=null,ad16=null"
   End If
   'end 2013/3/22
   
   strSql = "Update applicantdiscount Set " & strUpdSQL & " Where AD01='" & txtAD(0).Text & "' and ad02='" & txtAD(1).Text & "' "
         
   UpdateData = False
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   cnnConnection.Execute strSql, lngEffRec
   cnnConnection.CommitTrans
   UpdateData = True
   
   Exit Function
   
ErrHnd:

   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Function insertdata() As Boolean

   Dim strSql As String, intI As Integer, strSNo As String, OG(1 To 16) As String
   Dim strCols As String, strValues As String, lngEffRec As Long
   Dim rsQuery As New ADODB.Recordset, oCheck As CheckBox
   
   strCols = "ad01"
   For intI = 2 To 16
      strCols = strCols & ",ad" & Format(intI, "00")
   Next intI
   
   OG(1) = "'" & txtAD(0).Text & "'"
   OG(2) = "'" & txtAD(1).Text & "'"
   OG(3) = "'" & txtAD(3).Text & "'"
   OG(4) = "'" & strUserNum & "'"
   OG(5) = "TO_NUMBER(TO_CHAR(SYSDATE,'YYYYMMDD'))"
   OG(6) = "TO_NUMBER(TO_CHAR(SYSDATE,'HH24MI'))"
   OG(7) = "null"
   OG(8) = "null"
   OG(9) = "null"
   OG(10) = "'" & txtAD(2).Text & "'"
   OG(11) = "'" & txtAD(4).Text & "'"
   OG(12) = "'" & txtAD(5).Text & "'"
   OG(13) = "'" & txtAD(6).Text & "'"
   OG(14) = "'" & txtAD(7).Text & "'"
   
   'Added by Morgan 2013/3/22
   OG(15) = "null"
   OG(16) = "null"
   'Modified by Morgan 2019/4/10 加日本案減免資格
   If SSTab2.Enabled = True Then
      '台灣
      If SSTab2.TabVisible(6) = True Then
         'Modified by Morgan 2020/7/24
         'For intI = 1 To 4
         For intI = 1 To 6
            If chkAD15(intI).Value = 1 Then
               OG(15) = "'" & intI & "'"
               OG(16) = Val(txtAD16(intI).Text)
               Exit For
            End If
         Next
      '日本
      Else
         OG(15) = "'" & GetAD15() & "'"
      End If
   End If
   'end 2013/3/22
   
   strValues = Join(OG, ",")
   
   strSql = " INSERT INTO applicantdiscount (" & strCols & ") VALUES(" & strValues & ") "
         
         
   insertdata = False
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   cnnConnection.Execute strSql, lngEffRec
   cnnConnection.CommitTrans
   cur_AD01 = txtAD(0).Text & txtAD(1).Text
   insertdata = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub txtAD16_GotFocus(Index As Integer)
   TextInverse txtAD16(Index)
   CloseIme
End Sub

Private Sub txtAD16_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtAD16_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1
         If Val(txtAD16(Index)) > 80000000 Then
            MsgBox "金額輸入錯誤，不可高於八千萬！", vbExclamation
            Cancel = True
         End If
      Case 2
         If Val(txtAD16(Index)) > 100000000 Then
            MsgBox "金額輸入錯誤，不可高於一億！", vbExclamation
            Cancel = True
         End If
      Case 3
         If Val(txtAD16(Index)) >= 200 Then
            MsgBox "員工數輸入錯誤，必須小於200！", vbExclamation
            Cancel = True
         End If
      Case 4
         If Val(txtAD16(Index)) >= 100 Then
            MsgBox "員工數輸入錯誤，必須小於100！", vbExclamation
            Cancel = True
         End If
      
      'Added by Morgan 2020/7/24
      Case 5
         If Val(txtAD16(Index)) > 100000000 Then
            MsgBox "金額輸入錯誤，不可高於一億！", vbExclamation
            Cancel = True
         ElseIf Val(Format(txtAD16(Index))) < 1000 Then
            MsgBox "金額輸入錯誤，不可低於1000！", vbExclamation
            Cancel = True
         End If
      Case 6
         If Val(txtAD16(Index)) >= 200 Then
            MsgBox "員工數輸入錯誤，必須小於200！", vbExclamation
            Cancel = True
         ElseIf Val(txtAD16(Index)) = 0 Then
            MsgBox "員工數輸入錯誤，必須至少1人！", vbExclamation
            Cancel = True
         End If
      'end 2020/7/24
   End Select
   If Cancel = True Then txtAD16_GotFocus Index
End Sub

Private Sub txtFn_GotFocus(Index As Integer)
   If txtFn(Index).Locked = False Then
      TextInverse txtFn(Index)
      If txtFn(Index).Locked = False Then
         'edit by nickc 2007/07/11 切換輸入法改用API
         'txtFn(Index).IMEMode = 2
         CloseIme
      End If
   End If
End Sub

Private Sub txtFn_KeyPress(Index As Integer, KeyAscii As Integer)
   If txtFn(Index).Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      Select Case Index
         Case 0, 1
         '只可為文數字
            If Not (KeyAscii = 8 Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 47 And KeyAscii < 58)) Then
               KeyAscii = 0
            End If
         Case 2, 3
         '數字
            If Not (KeyAscii = 8 Or (KeyAscii > 47 And KeyAscii < 58)) Then
                KeyAscii = 0
            End If
      End Select
   End If
End Sub

Private Sub txtFn_LostFocus(Index As Integer)
   If SSTab1.Tab = 0 Then
      Dim bolCancel As Boolean
      bolCancel = False
      Call txtFn_Validate(Index, bolCancel)
      If bolCancel = True Then
         SSTab1.Tab = 1
         txtFn(Index).SetFocus
      End If
   End If
End Sub

Private Sub txtFn_Validate(Index As Integer, Cancel As Boolean)
   If txtFn(Index).Locked = False Then
      Select Case Index
         Case 0
         Case 1
               If Mid(UCase(txtFn(0)), 1, 6) <> Mid(UCase(txtFn(1)), 1, 6) And (txtFn(1) < txtFn(0)) And txtFn(1) <> "" Then
                  MsgBox "客戶編號迄值必需大於起值，且前 6 碼要相同！", vbCritical
                  Cancel = True
               End If
         Case 3
            If txtFn(2) <> "" And txtFn(3) < txtFn(2) Then
               MsgBox "申請國家迄值必需大於起值！", vbCritical
               Cancel = True
            End If
      End Select
      If Cancel = True Then txtFn_GotFocus (Index)
   End If
End Sub

Private Function CheckDBData() As Boolean
CheckOC
strSql = "select count(*) from applicantdiscount where ad01='" & txtAD(0).Text & "' and ad02='" & txtAD(1) & "' "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
    If adoRecordset.Fields(0).Value = 0 Then
        CheckDBData = True
    Else
        CheckDBData = False
    End If
End If
CheckOC
End Function
'Added by Morgan 2013/3/22
'Modified by Morgan 2019/4/9 增加日本,改用頁籤並將台灣加入
'減免資格
Private Sub SetOption()
   SSTab2.Visible = False
   SSTab2.Enabled = False
   SSTab2.TabVisible(6) = True '要先顯示最後一頁否則頁籤內容不會更新
   lblAD10 = "(1：自然人  2：學校  3：中小企業)"
   If txtAD(1) = "000" And txtAD(2) = "3" And txtAD(3) = "Y" Then
      For intI = 0 To SSTab2.Tabs - 1
         If intI = 6 Then
            SSTab2.TabVisible(intI) = True
         Else
            SSTab2.TabVisible(intI) = False
         End If
      Next
      SSTab2.Visible = True
      SSTab2.Enabled = True
   ElseIf txtAD(1) = "011" Then
      lblAD10 = "(1：中小企業　2：獨資企業" & vbCrLf & "  3：小企業　　4：新興企業" & vbCrLf & "  5：大學　　　6：個人)"
      If txtAD(3) = "Y" Then
         If txtAD(2) <> "" Then
            For intI = 0 To SSTab2.Tabs - 1
               If Val(txtAD(2)) - 1 = intI Then
                  SSTab2.TabVisible(intI) = True
                  If txtAD(2) = "5" Then chkAD15JP5(1).Value = 1 '大學只有一個選項自動勾選
               Else
                  SSTab2.TabVisible(intI) = False
               End If
            Next
            SSTab2.Visible = True
            SSTab2.Enabled = True
         End If
      End If
   End If
End Sub

Private Function GetAD15() As Integer
   Dim oCheck As CheckBox
   '中小企業
   If SSTab2.Tab = 0 Then
      For Each oCheck In chkAD15JP1
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   '獨資企業
   ElseIf SSTab2.Tab = 1 Then
      For Each oCheck In chkAD15JP2
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   '小企業
   ElseIf SSTab2.Tab = 2 Then
      For Each oCheck In chkAD15JP3
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   '新興企業
   ElseIf SSTab2.Tab = 3 Then
      For Each oCheck In chkAD15JP4
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   '大學
   ElseIf SSTab2.Tab = 4 Then
      For Each oCheck In chkAD15JP5
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   '個人
   ElseIf SSTab2.Tab = 5 Then
      For Each oCheck In chkAD15JP6
         If oCheck.Value = 1 Then
            GetAD15 = oCheck.Index
            Exit For
         End If
      Next
   End If
End Function

Private Sub ResetTab2()
   Dim oCheck As CheckBox
   
   For Each oCheck In chkAD15
      oCheck.Value = 0
      txtAD16(oCheck.Index).Text = ""
      txtAD16(oCheck.Index).Enabled = False
   Next
         
   For Each oCheck In chkAD15JP1
      oCheck.Value = 0
   Next
   For Each oCheck In chkAD15JP2
      oCheck.Value = 0
   Next
   For Each oCheck In chkAD15JP3
      oCheck.Value = 0
   Next
   For Each oCheck In chkAD15JP4
      oCheck.Value = 0
   Next
   For Each oCheck In chkAD15JP5
      oCheck.Value = 0
   Next
   For Each oCheck In chkAD15JP6
      oCheck.Value = 0
   Next
End Sub
