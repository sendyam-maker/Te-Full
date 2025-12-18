VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040160 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外部關聯企業資料維護"
   ClientHeight    =   5796
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   8496
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5796
   ScaleWidth      =   8496
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7425
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
            Picture         =   "frm12040160.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040160.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040160.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040160.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040160.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040160.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040160.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040160.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040160.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040160.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040160.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8496
      _ExtentX        =   14986
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
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   795
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   8700
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm12040160.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(21)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(22)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Lbl3(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Lbl3(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Lbl3(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Lbl3(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Lbl3(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Lbl3(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Lbl3(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Lbl3(7)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Lbl3(8)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Lbl3(9)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Lbl3(10)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Lbl3(11)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(12)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(13)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(14)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblPS"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblCUID"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtData(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtData(2)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtData(3)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtData(4)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdRemove"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdAdd"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "List1"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Combo1"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmdCopy(1)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cmdCopy(2)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm12040160.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(0)"
      Tab(1).Control(1)=   "Line2"
      Tab(1).Control(2)=   "Txt1(0)"
      Tab(1).Control(3)=   "Txt1(1)"
      Tab(1).Control(4)=   "cmdSearch"
      Tab(1).Control(5)=   "GRD1"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdCopy 
         Caption         =   "財務共通指示複製"
         Height          =   285
         Index           =   2
         Left            =   5760
         TabIndex        =   4
         Top             =   3300
         Width           =   1900
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "案件共通指示複製"
         Height          =   285
         Index           =   1
         Left            =   3800
         TabIndex        =   3
         Top             =   3300
         Width           =   1900
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   4440
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   2880
         Width           =   2895
      End
      Begin VB.ListBox List1 
         Height          =   768
         Left            =   1080
         TabIndex        =   41
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "←"
         Height          =   285
         Left            =   2620
         TabIndex        =   23
         Top             =   3000
         Width           =   600
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "→"
         Height          =   285
         Left            =   2620
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3300
         Width           =   600
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   14
         Top             =   960
         Width           =   8055
         _ExtentX        =   14203
         _ExtentY        =   6795
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         AllowUserResizing=   3
         FormatString    =   "編號|名稱|國籍|狀態|關聯編號|名稱|關聯關係|關聯說明"
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
         _Band(0).Cols   =   8
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "查詢"
         Height          =   330
         Left            =   -70920
         TabIndex        =   11
         Top             =   525
         Width           =   735
      End
      Begin VB.TextBox Txt1 
         Height          =   270
         Index           =   1
         Left            =   -72000
         MaxLength       =   8
         TabIndex        =   10
         Top             =   555
         Width           =   1020
      End
      Begin VB.TextBox Txt1 
         Height          =   270
         Index           =   0
         Left            =   -73200
         MaxLength       =   8
         TabIndex        =   8
         Top             =   555
         Width           =   1020
      End
      Begin MSForms.TextBox txtData 
         Height          =   950
         Index           =   4
         Left            =   1080
         TabIndex        =   5
         Top             =   3870
         Width           =   7020
         VariousPropertyBits=   -1466941413
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "12382;1676"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtData 
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Visible         =   0   'False
         Width           =   660
         VariousPropertyBits=   671105051
         MaxLength       =   18
         Size            =   "1164;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtData 
         Height          =   300
         Index           =   2
         Left            =   1830
         TabIndex        =   1
         Top             =   1680
         Width           =   1020
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1799;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtData 
         Height          =   300
         Index           =   1
         Left            =   1830
         TabIndex        =   0
         Top             =   477
         Width           =   1020
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1799;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCUID 
         Height          =   255
         Left            =   2310
         TabIndex        =   45
         Top             =   30
         Width           =   5895
         BackColor       =   12648384
         Size            =   "10398;450"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblPS 
         Caption         =   "P.S: 共通指示複製是從來源關聯企業複製指示到申請人/代理人編號，連同紀錄日期一併複製	"
         ForeColor       =   &H00FF0000&
         Height          =   228
         Left            =   912
         TabIndex        =   44
         Top             =   3672
         Width           =   7224
      End
      Begin VB.Label Label1 
         Caption         =   "關聯代號："
         Height          =   180
         Index           =   14
         Left            =   3480
         TabIndex        =   43
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "關聯說明："
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   42
         Top             =   3930
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "關       聯："
         Height          =   180
         Index           =   12
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Width           =   900
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   11
         Left            =   1020
         TabIndex        =   39
         Top             =   2520
         Width           =   7000
         VariousPropertyBits=   27
         Caption         =   "Lbl3(11)"
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   10
         Left            =   1020
         TabIndex        =   38
         Top             =   2280
         Width           =   7000
         VariousPropertyBits=   27
         Caption         =   "Lbl3(10)"
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   9
         Left            =   1020
         TabIndex        =   37
         Top             =   2055
         Width           =   7000
         VariousPropertyBits=   27
         Caption         =   "Lbl3(9)"
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   8
         Left            =   6000
         TabIndex        =   36
         Top             =   1725
         Width           =   1515
         VariousPropertyBits=   27
         Caption         =   "Lbl3(8)"
         Size            =   "2672;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   7
         Left            =   4200
         TabIndex        =   35
         Top             =   1728
         Width           =   1065
         VariousPropertyBits=   27
         Caption         =   "Lbl3(7)"
         Size            =   "1879;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   6
         Left            =   3720
         TabIndex        =   34
         Top             =   1728
         Width           =   495
         VariousPropertyBits=   27
         Caption         =   "Lbl3(6)"
         Size            =   "873;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   5
         Left            =   1020
         TabIndex        =   33
         Top             =   1320
         Width           =   7005
         VariousPropertyBits=   27
         Caption         =   "Lbl3(5)"
         Size            =   "12356;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   4
         Left            =   1020
         TabIndex        =   32
         Top             =   1080
         Width           =   7005
         VariousPropertyBits=   27
         Caption         =   "Lbl3(4)"
         Size            =   "12356;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   3
         Left            =   1020
         TabIndex        =   31
         Top             =   840
         Width           =   7005
         VariousPropertyBits=   27
         Caption         =   "Lbl3(3)"
         Size            =   "12356;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   2
         Left            =   6000
         TabIndex        =   30
         Top             =   525
         Width           =   1305
         VariousPropertyBits=   27
         Caption         =   "Lbl3(2)"
         Size            =   "2302;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   29
         Top             =   525
         Width           =   1060
         VariousPropertyBits=   27
         Caption         =   "Lbl3(1)"
         Size            =   "1870;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Lbl3 
         Height          =   285
         Index           =   0
         Left            =   3720
         TabIndex        =   28
         Top             =   525
         Width           =   495
         VariousPropertyBits=   27
         Caption         =   "Lbl3(0)"
         Size            =   "873;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "狀態："
         Height          =   180
         Index           =   11
         Left            =   5400
         TabIndex        =   27
         Top             =   1728
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "狀態："
         Height          =   180
         Index           =   10
         Left            =   5400
         TabIndex        =   26
         Top             =   525
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "國籍："
         Height          =   180
         Index           =   9
         Left            =   3120
         TabIndex        =   25
         Top             =   1728
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "國籍："
         Height          =   180
         Index           =   8
         Left            =   3120
         TabIndex        =   24
         Top             =   525
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "　　(日)："
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "　　(英)："
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "來源關聯企業編號："
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   19
         Top             =   1725
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "名稱(中)："
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   2055
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "　　(日)："
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "　　(英)："
         Height          =   180
         Index           =   22
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   900
      End
      Begin VB.Line Line2 
         X1              =   -71880
         X2              =   -72600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         Caption         =   "名稱(中)："
         Height          =   180
         Index           =   21
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "申請人/代理人編號："
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   525
         Width           =   1710
      End
      Begin VB.Label Label2 
         Caption         =   "申請人/代理人編號："
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   9
         Top             =   600
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frm12040160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/23 改成Form2.0 ; GRD1改字型=新細明體-ExtB、txtData(index)、Lbl3(index)、lblCUID
'Created by Lydia 2016/11/24 國外部關聯企業資料維護
Option Explicit

Dim m_EditMode As Integer '0:瀏覽 1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim mESeqNo As String '暫存TB編號

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Dim ii As Integer, jj As Integer
Dim idx As Integer

'Modified by Lydia 2021/09/23
'Dim oText As TextBox
'Dim oLabel As LABEL
Dim oText As Control
Dim oLabel As Control
'end 2021/09/23
Dim colFR01 As Integer 'Grid中FR01的欄位位置
Dim colFR02 As Integer 'Grid中FR02的欄位位置
Dim m_insCopy1 As Boolean '是否執行(各項指示)案件共通指示複製
Dim m_insCopy2 As Boolean '是否執行(各項指示)財務共通指示複製
Dim m_basCopy1 As Boolean '是否執行(指示=基本檔)案件共通指示複製
Dim m_basCopy2 As Boolean '是否執行(指示=基本檔)財務共通指示複製
'Added by Lydia 2020/07/27
Dim m_insDup1 As String '案件共通指示:檢查是否有同一分類的有效指示
Dim m_insDup2 As String '財務共通指示:檢查是否有同一分類的有效指示

Private Sub cmdAdd_Click()
Dim strAns As String

  If Combo1.ListIndex >= 0 And (m_EditMode = 1 Or m_EditMode = 2) Then
     Call Pub_AddFRelationList(Me.Combo1, Me.List1, strAns)
     If strAns <> "" Then
        txtData(3).Text = txtData(3).Text & strAns & ","
        'Mark by Lydia 2020/08/14 David: 先不開放複製功能
        'Remove Mark by Lydia 2021/08/18 開放複製功能
        If strAns = "1" Then
            cmdCopy(1).Visible = True
        ElseIf strAns = "2" Then
            cmdCopy(2).Visible = True
        End If
        If InStr("1,2", strAns) > 0 Then
            lblPS.Visible = True
        End If
        'end 2020/08/14
     End If
  End If
End Sub

Private Sub cmdCopy_Click(Index As Integer)
Dim strTitle As String, strKind As String

  If Me.txtData(1).Text = "" Then
     MsgBox "請輸入申請人/代理人編號！", vbInformation
     Exit Sub
  ElseIf Me.txtData(2).Text = "" Then
     MsgBox "請輸入關聯企業編號！", vbInformation
     Exit Sub
  End If
           
  '各項指示：檢查表單是否開啟中
  If PUB_CheckFormExist("frm12040159") Then
      MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
      Exit Sub
  End If
  
  If Index = 1 Then
     strTitle = "案件"
     strKind = "A"
  Else
     strTitle = "財務"
     strKind = "F"
  End If
  
  If Pub_GetInstructions(Me.Name, txtData(2), strExc(1), strKind, , , , "1") = False Then
     'Modified by Lydia 2020/08/14 David: 不用開維護
     'MsgBox "關聯企業無" & strTitle & "指示！" & vbCrLf & "請自行維護" & txtData(1) & "的各項指示。", vbInformation
     'frm12040159.SetParent "E", Trim(Me.txtData(1).Text), Me
     'frm12040159.Show
     'Exit Sub
     MsgBox "關聯企業無" & strTitle & "指示！", vbInformation
  Else
     'Memo by Lydia 2020/06/12 關聯企業資料維護的共通指示複製功能：按下功能按鈕時先檢查二編號的指示=基本檔的欄位是否有不同的欄位設定
                                     '，若有則提醒"二編號基本檔的指示欄位設定值不同，是否確定要覆蓋？(是)覆蓋 (否)取消複製"，讓使用者可選擇，若覆蓋時同時記錄修改記錄DML_LOG。
                                     '再檢查維護編號的各項指示檔若有資料再提醒"X(Y)編號已有各項指示的設定，是否確定要覆蓋？(是)覆蓋 (否)取消" ，
                                     '讓使用者可選擇，若覆蓋時則將原各項指示都改為無效指示以便留下記錄，覆蓋過來的指示都設定為系統日。
     If Pub_GetInstructions(Me.Name, txtData(2), strExc(1), strKind, , , , "2") = True Then
        If Pub_GetInstructions(Me.Name, txtData(1), strExc(1), strKind, , , , "2") = True Then
            If MsgBox(strTitle & "共通指示複製：" & vbCrLf & "二編號基本檔的指示欄位設定值不同，是否確定要覆蓋？" & vbCrLf & "是：覆蓋　　否：取消覆蓋", vbInformation + vbYesNo + vbDefaultButton2, strTitle & "共通指示複製 ") = vbYes Then
                If Index = 1 Then m_basCopy1 = True
                If Index = 2 Then m_basCopy2 = True
            End If
        Else
            If Index = 1 Then m_basCopy1 = True
            If Index = 2 Then m_basCopy2 = True
        End If
     End If
     
     '檢查是否有重複關聯
     'Mark by Lydia 2022/11/25 複製對象編號ITS13：將來源X / Y編號各項指示複製至目的X / Y編號，連同"紀錄日期"一併複製
'     strExc(0) = "select count(*) cnt from instructions where its02='" & txtData(1) & "' and its04=" & CompWorkDay(2, strSrvDate(1)) & _
'                      " AND SUBSTR(ITS03,1,1) " & IIf(strKind = "F", "='F'", "<>'F'") & " and its03 in (select its03 from instructions where its02='" & txtData(2) & "' and  nvl(its05,'Y') <> 'N')"
'     intI = 1
'     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'     If intI = 1 Then
'         If Val("" & RsTemp.Fields("cnt")) > 0 Then
'              MsgBox strTitle & "共通指示複製：" & vbCrLf & txtData(1) & "已有各項指示的關聯設定，請先到各項指示進行維護！", vbExclamation, strTitle & "共通指示複製"
'              Exit Sub
'         End If
'     End If
     'end 2022/11/25
     If Pub_GetInstructions(Me.Name, txtData(1), strExc(1), strKind, , , , "1") = True Then
         If MsgBox(strTitle & "共通指示複製：" & vbCrLf & txtData(1) & "已有各項指示的設定，是否確定要複製？" & vbCrLf & "是：複製　　否：取消複製", vbInformation + vbYesNo + vbDefaultButton2, strTitle & "共通指示複製") = vbYes Then
             If Index = 1 Then m_insCopy1 = True
             If Index = 2 Then m_insCopy2 = True
         End If
     Else
         If Index = 1 Then m_insCopy1 = True
         If Index = 2 Then m_insCopy2 = True
     End If
     'end 2020/06/12
     
     'Added by Lydia 2020/07/27 因為覆蓋過去的指示都預設為系統日，所以要先檢查是否有不同記錄日期的同一分類指示
     'Mark by Lydia 2022/11/25 複製對象編號ITS13：將來源X / Y編號各項指示複製至目的X / Y編號，連同"紀錄日期"一併複製(而非現行的系統日+1工作天)
'     If (Index = 1 And m_insCopy1 = True) Or (Index = 2 And m_insCopy2 = True) Then
'         strExc(0) = "select its03,count(*) cnt from instructions where its02='" & txtData(2) & "' and its05 is null "
'         If strKind = "F" Then
'             strExc(0) = strExc(0) & "AND SUBSTR(ITS03,1,1)='F' "
'         Else
'             strExc(0) = strExc(0) & "AND SUBSTR(ITS03,1,1)<>'F' "
'         End If
'         strExc(0) = strExc(0) & "group by its03 having count(*) > 1 "
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'             strExc(1) = ""
'             RsTemp.MoveFirst
'             Do While Not RsTemp.EOF
'                 If Val("" & RsTemp.Fields("cnt")) > 1 Then
'                     strExc(1) = strExc(1) & "," & RsTemp.Fields("its03")
'                 End If
'                 RsTemp.MoveNext
'             Loop
'             If strExc(1) <> "" Then
'                MsgBox strTitle & "共通指示複製：" & vbCrLf & txtData(2) & "的分類代號：" & Mid(strExc(1), 2) & vbCrLf & "有不同記錄日期的指示，不可進行覆蓋，請先到各項指示進行維護！", vbExclamation, strTitle & "共通指示複製"
'                If Index = 1 Then m_insDup1 = Mid(strExc(1), 2)
'                If Index = 2 Then m_insDup2 = Mid(strExc(1), 2)
'                Exit Sub
'             End If
'         End If
'     End If
'     'end 2020/07/27
     'end 2022/11/25
  End If
End Sub

Private Sub cmdRemove_Click()
  If m_EditMode = 1 Or m_EditMode = 2 Then
     Call Pub_RemSelectList(List1, strExc(1))
     If strExc(1) <> "" Then
        txtData(3).Text = Replace(txtData(3).Text, strExc(1) & ",", "")
        If strExc(1) = "1" Then
           m_insCopy1 = False
           m_basCopy1 = False
           m_insDup1 = ""
           'Remove Mark by Lydia 2021/08/18 開放複製功能
           cmdCopy(1).Visible = False 'Mark by Lydia 2020/08/14 David: 先不開放複製功能
        End If
        If strExc(1) = "2" Then
           m_insCopy2 = False
           m_basCopy2 = False
           m_insDup2 = ""
           'Remove Mark by Lydia 2021/08/18 開放複製功能
           cmdCopy(2).Visible = False  'Mark by Lydia 2020/08/14 David: 先不開放複製功能
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
Dim strS1 As String
Dim stSQL As String
Dim tmpArr As Variant
Dim strType As String
Dim strTypeName As String

QueryData = False
   
   If SSTab1.Tab = 1 Then '多筆查詢
        If Txt1(0) <> "" And Txt1(1) <> "" Then
           If Txt1(1) < Txt1(0) Then
              MsgBox "申請人/代理人終止編號不可小於起始編號!"
              Exit Function
           Else
              strS1 = strS1 & "AND FR01>=" & CNULL(Txt1(0)) & " AND FR01<=" & CNULL(Txt1(1))
           End If
        End If
   Else   '單筆維護
      If Txt1(0) <> "" Then
         strS1 = strS1 & "AND FR01>=" & CNULL(Txt1(0))
      ElseIf Txt1(1) <> "" Then
         strS1 = strS1 & "AND FR01<=" & CNULL(Txt1(1))
      End If
   End If
   
   m_blnColOrderAsc = True
   QueryData = True

   strSql = "select fr01,fr02,fr03 from frelation where 1=1 " & IIf(strS1 <> "", strS1, "")
   strSql = strSql & " order by 1,2,3 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   
On Error GoTo ErrHnd:
   If intI = 1 Then
       '暫存db
       Set rsRead = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)

       cnnConnection.BeginTrans
       rsRead.MoveFirst
       With rsRead
          Do While Not .EOF
             stSQL = ""
             For jj = 1 To 2
                If jj = 1 Then
                   strExc(2) = "" & .Fields("fr01")
                Else
                   strExc(2) = "" & .Fields("fr02")
                End If
                '申請人/代理人資料
                If Left(strExc(2), 1) = "X" Then
                   strExc(0) = "select cu01 No,replace(cu04,'&','＆') name1,replace(cu05||' '||cu88||' '||cu89||' '||cu90,'&','＆') name2,replace(cu06,'&','＆') name3,cu10 na01,na03,cu80 status" & _
                      " from customer,nation where cu01 = '" & strExc(2) & "' and cu02='0' and cu10=na01(+) "
                Else
                   strExc(0) = "select fa01 No,replace(fa04,'&','＆') name1,replace(fa05||' '||fa63||' '||fa64||' '||fa65,'&','＆') name2,replace(fa06,'&','＆') name3,fa10 na01,na03,fa69 status" & _
                      " from fagent,nation where fa01 = '" & strExc(2) & "' and fa02='0' and fa10=na01(+) "
                End If
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                   If jj = 1 Then
                      stSQL = stSQL & ", r004=" & CNULL(ChgSQL(Trim("" & RsTemp.Fields("name1")))) 'FR01Name1
                      stSQL = stSQL & ", r005=" & CNULL(ChgSQL(Trim("" & RsTemp.Fields("name2")))) 'FR01Name2
                      stSQL = stSQL & ", r006=" & CNULL(ChgSQL(Trim("" & RsTemp.Fields("name3")))) 'FR01Name3
                      stSQL = stSQL & ", r007=" & CNULL(Trim(Mid("" & RsTemp.Fields("na01"), 1, 3))) 'FR01Na01
                      stSQL = stSQL & ", r008=" & CNULL(Trim("" & RsTemp.Fields("na03")))  'FR01Na03
                      stSQL = stSQL & ", r009=" & CNULL(Trim("" & RsTemp.Fields("status"))) 'FR01status
                   Else
                      stSQL = stSQL & ", r010=" & CNULL(ChgSQL(Trim("" & RsTemp.Fields("name1")))) 'FR02Name1
                      stSQL = stSQL & ", r011=" & CNULL(ChgSQL(Trim("" & RsTemp.Fields("name2")))) 'FR02Name2
                      stSQL = stSQL & ", r012=" & CNULL(ChgSQL(Trim("" & RsTemp.Fields("name3")))) 'FR02Name3
                      stSQL = stSQL & ", r013=" & CNULL(Trim(Mid("" & RsTemp.Fields("na01"), 1, 3))) 'FR02Na01
                      stSQL = stSQL & ", r014=" & CNULL(Trim("" & RsTemp.Fields("na03")))  'FR02Na03
                      stSQL = stSQL & ", r015=" & CNULL(Trim("" & RsTemp.Fields("status"))) 'FR02status
                   End If
                End If
                
                '關聯代號->說明
                If jj = 2 Then
                    tmpArr = Empty: strType = "": strTypeName = ""
                    tmpArr = Split(.Fields("fr03"), ",")
                    For ii = 0 To UBound(tmpArr)
                        If Trim(tmpArr(ii)) <> "" Then
                           strExc(0) = "select ft02 from ftype where ft01=" & CNULL(Trim(tmpArr(ii)))
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              strTypeName = strTypeName & IIf(strTypeName <> "", ";", "") & "" & RsTemp(0)
                           Else
                              strTypeName = strTypeName & IIf(strTypeName <> "", ";", "") & "" & Trim(tmpArr(ii))
                           End If
                        End If
                    Next
                    
                    If strTypeName <> "" Then stSQL = stSQL & ", r016=" & CNULL(strTypeName)
                    
                End If

                If jj = 2 And stSQL <> "" Then
                   strExc(1) = "update rdatafactory set " & Trim(IIf(Left(stSQL, 1) = ",", Mid(stSQL, 2), stSQL)) & " where FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "' and rowseq=" & CNULL(rsRead.AbsolutePosition)
                   cnnConnection.Execute strExc(1), intI
                End If
             Next jj
             .MoveNext
          Loop
       End With
       cnnConnection.CommitTrans

       strSql = "select fr01,nvl(r004,nvl(r005,r006)) fr01name,r007 fr01na01,r008 fr01na03,r009 fr01status," & _
                "fr02,nvl(r010,nvl(r011,r012)) fr02name,r013 fr02na01,r014 fr02na03,r015 fr02status," & _
                "fr03,r016 fr03desc,fr04,replace(fr04,chr(13)||chr(10),' & ') fr04r from frelation,RDataFactory where 1=1 " & strS1 & " and fr01=r001(+) and fr02=r002(+) " & _
                " and FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno='" & mESeqNo & "'"
       strSql = strSql & "order by fr01,fr02 "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strSql)
       GRD1.FixedCols = 0
       Set GRD1.Recordset = RsTemp
       Call SetGrd(RsTemp.RecordCount + 1)
       GRD1.FixedCols = 2
       QueryData = True
   Else
       If bolM = True Then
          MsgBox "查無資料!!"
       End If
       GRD1.Clear
       Call SetGrd
   End If
   
   Set rsRead = Nothing
   Exit Function
   
ErrHnd:
   If Err.Number > 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function

'Added by Lydia 2021/10/21
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
'Memo by Lydia 2021/10/21 從Form_KeyDown搬來
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
      'Remove by Lydia 2021/11/22 取消以ENTER控制為換行的功能 (Form2.0修改之維護資料功能Toolbar之修改統一)
      'Case vbKeyReturn
      '   If m_EditMode <> 0 Then
      '      '說明欄位可允許Enter(換行)
      '      If Me.ActiveControl <> txtData(4) Then
      '          OnAction vbKeyF9
      '      End If
      '   Else
      '      KeyCode = 0 '取消動作
      '   End If
      'end 2021/11/22
      
   End Select
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
 
   lblCUID.BackColor = &H8000000F
   ClearField
   'Memo by Lydia 2019/10/17 注意關聯代號只能第一碼為英文,若改變格式須修改模組PUB_Num2Id和PUB_Id2Num
   Pub_SetFTypeList Me.Combo1, 10

   m_EditMode = 0
   
   SetInputEntry
   UpdateToolbarState
   '清除查詢
   For Each oText In Txt1
      oText.Text = Empty
   Next
   
   Call SetGrd
   SetCtrlReadOnly True
   
   Me.SSTab1.Tab = 0
End Sub

Private Sub SetGrd(Optional ByRef iR As Integer = 2)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

    arrGridHeadText = Array("編號", "名稱", "FR01NA01", "國籍", "狀態", "關聯編號", "名稱", "FR02NA01", "FR02NA03", "FR02STATUS", "FR03", "關聯", _
                          "FR04", "關聯說明", "FR05", "FR06", "FR07", "FR08", "FR09", "FR10", "R004", "R005", "R006", "R010", "R011", "R012")
   arrGridHeadWidth = Array(860, 960, 0, 860, 960, 860, 960, 0, 0, 0, 0, 1000, _
                           0, 1500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   
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
       
       For idx = 1 To iR - 1
         .row = idx
         For iRow = 0 To .Cols - 1
           .col = iRow
           If iRow < 2 Then
              .CellBackColor = QBColor(15) '底色
           End If
         Next iRow
       Next idx
       
       .Visible = True
   End With
   
   If colFR02 = 0 Then
      colFR01 = PUB_MGridGetId("編號", GRD1)
      colFR02 = PUB_MGridGetId("關聯編號", GRD1)
   End If
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   If Me.Visible = True Then
      Select Case m_EditMode
         Case 1 '新增
            txtData(1).SetFocus
            Txtdata_GotFocus 1
         Case 2 '修改
            txtData(4).SetFocus
            Txtdata_GotFocus 4
         Case 4 '查詢
            txtData(1).SetFocus
            Txtdata_GotFocus 1
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
         'If m_bUpdate And txtData(1) <> "" Then
         If m_bUpdate Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         'Modifeid by Lydia 2016/01/28 改到OnAction控制
         'If m_bDelete And txtData(1) <> "" Then
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
         If m_bQuery And txtData(1) <> "" Then
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

   Set frm12040160 = Nothing
End Sub

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0, Optional ByVal bolMsg As Boolean = True) As Boolean
   
   Dim adoRst As New ADODB.Recordset
   Dim stCon As String
   
   If p_iWay = -1 Or p_iWay = 1 Or p_iWay = 0 Then
      If txtData(1) <> "" Then stCon = stCon & " and FR01=" & CNULL(txtData(1))
      If txtData(2) <> "" Then stCon = stCon & " and FR02=" & CNULL(txtData(2))
   End If
   'Added by Lydia 2021/10/21 找不到資料,按上一筆/下一筆
   If (p_iWay = -1 Or p_iWay = 1) And Trim(txtData(1) & txtData(2)) = "" Then
        If p_iWay = -1 Then
           p_iWay = -2
        ElseIf p_iWay = 1 Then
           p_iWay = 2
        End If
   End If
   'end 2021/10/21
   
   strExc(0) = "SELECT * FROM FRELATION WHERE "
   Select Case p_iWay
      '尋找
      Case 0
          strExc(0) = strExc(0) & "1=1 " & stCon & " ORDER BY 1,2"
      '首筆
      Case -2
          strExc(0) = strExc(0) & "FR01||FR02=(select min(FR01||FR02) from FRELATION) "
      '前一筆
      Case -1
          strExc(0) = strExc(0) & "FR01||FR02=(select max(FR01||FR02) from FRELATION where FR01||FR02 <'" & txtData(1) & txtData(2) & "') "
      '後一筆
      Case 1
          strExc(0) = strExc(0) & "FR01||FR02=(select min(FR01||FR02) from FRELATION where FR01||FR02 >'" & txtData(1) & txtData(2) & "') "
      '末筆
      Case 2
          strExc(0) = strExc(0) & "FR01||FR02=(select max(FR01||FR02) from FRELATION) "
   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
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
      ClearField
   End If

   If m_EditMode <> "1" And m_EditMode <> "2" Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing

End Function
' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
Dim CUID(1 To 6) As String
Dim tmpArr As Variant

   ClearField
   With p_Rst
      If .RecordCount > 0 Then
         For Each oText In txtData
            If oText.Index > 0 Then
               oText.Text = "" & .Fields("FR" & Format(oText.Index, "00"))
               oText.Tag = oText.Text
            End If
         Next
         If txtData(1).Text <> "" Then Txtdata_Validate 1, False
         If txtData(2).Text <> "" Then Txtdata_Validate 2, False
         
         If txtData(3) <> "" Then
            Call Pub_ShowSelectList(Combo1, List1, txtData(3).Text)
            'Mark by Lydia 2020/08/14 David: 先不開放複製功能
            'Remove Mark by Lydia 2021/08/18 開放複製功能
            If InStr(txtData(3), "1,") > 0 Then cmdCopy(1).Visible = True
            If InStr(txtData(3), "2,") > 0 Then cmdCopy(2).Visible = True
            If cmdCopy(1).Visible = True Or cmdCopy(2).Visible = True Then
                lblPS.Visible = True
            End If
            'end 2020/08/14
         End If
         CUID(1) = "" & .Fields("FR05")
         CUID(2) = "" & .Fields("FR06")
         CUID(3) = "" & .Fields("FR07")
         CUID(4) = "" & .Fields("FR08")
         CUID(5) = "" & .Fields("FR09")
         CUID(6) = "" & .Fields("FR10")
      End If
   End With
   
   Combo1.ListIndex = -1
   UpdateCUID CUID, lblCUID

End Sub

Private Sub ClearField()

   For Each oText In txtData
      oText.Text = Empty
      oText.Tag = Empty
   Next

   For Each oLabel In Lbl3
      oLabel.Caption = ""
   Next

   cmdCopy(1).Visible = False
   cmdCopy(2).Visible = False
   lblPS.Visible = False
   Combo1.ListIndex = -1
   List1.Clear

   m_insCopy1 = False
   m_insCopy2 = False
   m_basCopy1 = False
   m_basCopy2 = False
   m_insDup1 = ""
   m_insDup2 = ""

End Sub

Private Sub GRD1_DblClick()
   If GRD1.MouseRow > 0 And GRD1.TextMatrix(GRD1.row, colFR01) <> "" And GRD1.TextMatrix(GRD1.row, colFR02) <> "" Then
      txtData(1) = GRD1.TextMatrix(GRD1.row, colFR01)
      txtData(2) = GRD1.TextMatrix(GRD1.row, colFR02)
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
    If m_blnColOrderAsc = True Then
       Me.GRD1.Sort = 5 '字串昇冪
       m_blnColOrderAsc = False
    Else
       Me.GRD1.Sort = 6 '字串降冪
       m_blnColOrderAsc = True
    End If
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse Txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Dim iLen As Integer

   Select Case Index
      Case 0, 1
           If Txt1(Index) <> "" Then
              If InStr("X,Y", Left(Txt1(Index), 1)) = 0 Then
                 MsgBox "申請人/代理人編號必須為X/Y開頭！", vbCritical + vbOKOnly, "檢核資料"
                 GoTo JumpCancel
              Else
                 If Len(Txt1(Index)) < 6 Then
                    MsgBox "申請人/代理人編號請至少輸入六碼！", vbCritical + vbOKOnly, "檢核資料"
                    GoTo JumpCancel
                 Else
                    Txt1(Index) = Mid(Txt1(Index) & "00", 1, 8)
                 End If
              End If
           End If
   End Select
   
   If Cancel = False Then
      If Txt1(Index).MaxLength > 0 Then
         If Not CheckLengthIsOK(Txt1(Index), iLen) Then
            GoTo JumpCancel
         End If
      End If
   End If
   Exit Sub
   
JumpCancel:
   Cancel = True
End Sub


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
Private Sub OnAction(ByVal KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         m_EditMode = 1
         ClearField
         Me.SSTab1.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyF3 ' 修改
         '不在UpdateToolbarState控制
         If txtData(1) = "" Or txtData(2) = "" Or txtData(3) = "" Then
            MsgBox "無記錄可修改!!", vbCritical
            Exit Sub
         End If

         If SSTab1.Tab = 1 Then
            MsgBox "請先選擇記錄並切換到單筆資料!", vbCritical
            Exit Sub
         End If
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyF5 ' 刪除
         '不在UpdateToolbarState控制
         If txtData(1) = "" Or txtData(2) = "" Or txtData(3) = "" Then
            MsgBox "無記錄可刪除!!", vbCritical
            Exit Sub
         End If

         If DelMsg() Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         If SSTab1.Tab = 1 Then
             SSTab1.Tab = 0 '切換回單筆
         End If
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
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         
         SetInputEntry
         
      Case vbKeyF10 ' 取消
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  If m_EditMode = 1 Then
                     m_EditMode = 0
                     ClearField
                     SetInputEntry
                  Else
                     m_EditMode = 0
                     txtData(1).Text = txtData(1).Tag
                     txtData(2).Text = txtData(2).Tag
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
            bolR = True
         End If
      Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord(0, True) = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtData(1).SetFocus
               Txtdata_GotFocus 1
            End If
         End If
   End Select
   
   '多筆資料-整理
   If bolR = True Then
      If QueryData(False) = False Then
      End If
   End If
End Function

Private Function TxtValidate() As Boolean
Dim tmpArr As Variant
Dim Cancel As Boolean, ii As Integer, jj As Integer

TxtValidate = False
   
   If txtData(1) = "" Then
      MsgBox "請輸入申請人/代理人編號！", vbCritical + vbOKOnly, "檢核資料"
      txtData(1).SetFocus
      Txtdata_GotFocus 1
      Exit Function
   End If

   Txtdata_Validate 1, Cancel
   If Cancel = True Then
      txtData(1).SetFocus
      Txtdata_GotFocus 1
      Exit Function
   End If
   
   'Added by Lydia 2021/09/24 因為輸入錯誤會彈「資料庫無資料 !」，在按下確定鍵時會執行ToolBar的確定，所以要另外做檢查。
   If Trim(Lbl3(3) & Lbl3(4) & Lbl3(5)) = "" Then '判斷無名稱不可存檔
        txtData(1).SetFocus
        Txtdata_GotFocus 1
        Exit Function '之前已彈過「資料庫無資料 !」
   End If
   'end 2021/09/24
   
   '新增,修改
   If m_EditMode <> 4 Then
      If txtData(2) = "" Then
         MsgBox "請輸入關聯企業編號！", vbCritical + vbOKOnly, "檢核資料"
         txtData(2).SetFocus
         Txtdata_GotFocus 2
         Exit Function
      End If

      Txtdata_Validate 2, Cancel
      If Cancel = True Then
         txtData(2).SetFocus
         Txtdata_GotFocus 2
         Exit Function
      End If
      
      'Added by Lydia 2021/09/24 因為輸入錯誤會彈「資料庫無資料 !」，在按下確定鍵時會執行ToolBar的確定，所以要另外做檢查。
      If Trim(Lbl3(9) & Lbl3(10) & Lbl3(11)) = "" Then '判斷無名稱不可存檔
           txtData(2).SetFocus
           Txtdata_GotFocus 2
           Exit Function '之前已彈過「資料庫無資料 !」
      End If
      'end 2021/09/24
   
      If txtData(1) = txtData(2) Then
         MsgBox "關聯企業編號不可與申請人/代理人編號相同！", vbCritical + vbOKOnly, "檢核資料"
         txtData(2).SetFocus
         Txtdata_GotFocus 2
         Exit Function
      End If
      If txtData(3).Text = "" Then
         MsgBox "請新增關聯！", vbCritical + vbOKOnly, "檢核資料"
         Combo1.SetFocus
         Exit Function
      End If
   End If
   
   'Added by Lydia 2021/09/23 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If

   If m_EditMode = 1 Then
      If ChkIsExist(txtData(1), txtData(2)) = True Then
         Exit Function
      End If
   End If
   
   'Added by Lydia 2020/07/27 因為覆蓋過去的指示都預設為系統日，所以要先檢查是否有不同記錄日期的同一分類指示
   'Mark by Lydia 2022/11/25 複製對象編號ITS13：將來源X / Y編號各項指示複製至目的X / Y編號，連同"紀錄日期"一併複製(而非現行的系統日+1工作天)
   'If (m_insCopy1 = True And m_insDup1 <> "") Or (m_insCopy2 = True And m_insDup2 <> "") Then
   '    strExc(1) = IIf(m_insDup1 <> "", "案件", "財務")
   '    strExc(2) = strExc(1) & "共通指示複製：" & vbCrLf & txtData(2) & "的分類代號：" & IIf(m_insDup1 <> "", m_insDup1, m_insDup2) & vbCrLf & _
   '                     "有不同記錄日期的指示，不可進行覆蓋，請先到各項指示進行維護！"
   '    MsgBox strExc(2), vbExclamation, strExc(1) & "共通指示複製"
   '    Exit Function
   'End If
   ''end 2020/07/27
   'end 2022/11/25
TxtValidate = True
   
End Function

Private Function ChkIsExist(ByVal pK01 As String, ByVal pK02 As String) As Boolean
Dim inA As Integer
   
ChkIsExist = False

   If pK01 = "" Or pK02 = "" Then
      Exit Function
   End If
      
   strExc(0) = "select 1 ord1,count(*) cnt from frelation where fr01=" & CNULL(Mid(pK01 & String(8, "0"), 1, 8)) & " and fr02=" & CNULL(Mid(pK02 & String(8, "0"), 1, 8))
   '相反方向的檢查
   strExc(0) = strExc(0) & " Union all select 2 ord1,count(*) cnt from frelation where fr02=" & CNULL(Mid(pK01 & String(8, "0"), 1, 8)) & " and fr01=" & CNULL(Mid(pK02 & String(8, "0"), 1, 8))
   inA = 1
   Set RsTemp = ClsLawReadRstMsg(inA, strExc(0))
   If inA = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
        If Val(RsTemp.Fields("cnt")) > 0 Then
           MsgBox "關聯企業已存在 " & IIf(RsTemp.Fields("ord1") = 1, Mid(pK01 & String(8, "0"), 1, 8) & "->" & Mid(pK02 & String(8, "0"), 1, 8), Mid(pK02 & String(8, "0"), 1, 8) & "->" & Mid(pK01 & String(8, "0"), 1, 8)) & " 的關聯!", vbCritical + vbOKOnly, "關聯企業資料"
           ChkIsExist = True
           Exit Function
        End If
         RsTemp.MoveNext
      Loop
   End If
   
End Function
' 新增記錄
Private Function AddRecord() As Boolean
Dim pErrMsg As String
Dim strExp As String
Dim bolOK As Boolean

On Error GoTo ErrHand

   cnnConnection.BeginTrans
       strSql = "INSERT INTO FRELATION (FR01,FR02,FR03,FR04,FR05,FR06,FR07) " & _
                "VALUES ('" & Trim(txtData(1)) & "','" & Trim(txtData(2)) & "','" & Trim(txtData(3)) & "','" & ChgSQL(Trim(txtData(4))) & "','" & strUserNum & "'," & strSrvDate(1) & "," & Left(Format(ServerTime, "000000"), 4) & ")"
       Pub_SeekTbLog strSql
       cnnConnection.Execute strSql, intI
       
       '覆蓋指示=基本檔的欄位
       If m_basCopy1 = True Then
          bolOK = GetUpdInstruction(txtData(1), txtData(2), "A", "2", pErrMsg)
          strExp = strExp & IIf(pErrMsg <> "", vbCrLf, "") & pErrMsg
          If bolOK = False Then GoTo ErrHand
       End If
       If m_basCopy2 = True Then
          bolOK = GetUpdInstruction(txtData(1), txtData(2), "F", "2", pErrMsg)
          strExp = strExp & IIf(pErrMsg <> "", vbCrLf, "") & pErrMsg
          If bolOK = False Then GoTo ErrHand
       End If
       '覆蓋各項指示
       If m_insCopy1 = True Then
          bolOK = GetUpdInstruction(txtData(1), txtData(2), "A", "1", pErrMsg)
          strExp = strExp & IIf(pErrMsg <> "", vbCrLf, "") & pErrMsg
          If bolOK = False Then GoTo ErrHand
       End If
       If m_insCopy2 = True Then
          bolOK = GetUpdInstruction(txtData(1), txtData(2), "F", "1", pErrMsg)
          strExp = strExp & IIf(pErrMsg <> "", vbCrLf, "") & pErrMsg
          If bolOK = False Then GoTo ErrHand
       End If
   cnnConnection.CommitTrans
   AddRecord = True
   
   If strExp <> "" And InStr(strExp, "人工處理") > 0 Then
       MsgBox strExp, vbInformation, "共通指示複製"
   End If
   Call ProcSendMail 'Added by Lydia 2022/11/25
   
   Exit Function
   
ErrHand:
   If Err.Number > 0 Or strExp <> "" Then
      cnnConnection.RollbackTrans
      MsgBox strExp & vbCrLf & Err.Description
   End If

End Function

' 刪除記錄
Private Function DelRecord() As Boolean

On Error GoTo ErrHand
   cnnConnection.BeginTrans
      strSql = "DELETE FROM FRELATION WHERE FR01=" & CNULL(txtData(1)) & " AND FR02=" & CNULL(txtData(2))
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
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
Dim pErrMsg As String
Dim strExp As String
Dim bolOK As Boolean

On Error GoTo ErrHand

   strExc(1) = ""
   
   For Each oText In txtData
       If oText.Index > 0 Then
          If oText.Tag <> oText.Text Then
             strExc(1) = strExc(1) & ", FR" & Format(oText.Index, "00") & "=" & CNULL(Trim(ChgSQL(oText.Text)))
          End If
       End If
   Next
   
   If strExc(1) <> "" Or m_insCopy1 = True Or m_insCopy2 = True Or m_basCopy1 = True Or m_basCopy2 = True Then
       cnnConnection.BeginTrans
           strSql = "UPDATE FRELATION SET FR08='" & strUserNum & "', FR09=" & strSrvDate(1) & ", FR10=" & Left(Format(ServerTime, "000000"), 4) & _
               Trim(strExc(1)) & " WHERE FR01=" & CNULL(txtData(1).Tag) & " AND FR02=" & CNULL(txtData(2).Tag)
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql, intI
        '覆蓋指示=基本檔的欄位
       If m_basCopy1 = True Then
          bolOK = GetUpdInstruction(txtData(1), txtData(2), "A", "2", pErrMsg)
          strExp = strExp & IIf(pErrMsg <> "", vbCrLf, "") & pErrMsg
          If bolOK = False Then GoTo ErrHand
       End If
       If m_basCopy2 = True Then
          bolOK = GetUpdInstruction(txtData(1), txtData(2), "F", "2", pErrMsg)
          strExp = strExp & IIf(pErrMsg <> "", vbCrLf, "") & pErrMsg
          If bolOK = False Then GoTo ErrHand
       End If
       '覆蓋各項指示
       If m_insCopy1 = True Then
          bolOK = GetUpdInstruction(txtData(1), txtData(2), "A", "1", pErrMsg)
          strExp = strExp & IIf(pErrMsg <> "", vbCrLf, "") & pErrMsg
          If bolOK = False Then GoTo ErrHand
       End If
       If m_insCopy2 = True Then
          bolOK = GetUpdInstruction(txtData(1), txtData(2), "F", "1", pErrMsg)
          strExp = strExp & IIf(pErrMsg <> "", vbCrLf, "") & pErrMsg
          If bolOK = False Then GoTo ErrHand
       End If
       cnnConnection.CommitTrans
   End If
   
   ModRecord = True
   
   If strExp <> "" And InStr(strExp, "人工處理") > 0 Then
       MsgBox strExp, vbInformation, "共通指示複製"
   End If
   Call ProcSendMail 'Added by Lydia 2022/11/25
   
   Exit Function
   
ErrHand:
   If Err.Number > 0 Or strExp <> "" Then
      cnnConnection.RollbackTrans
      MsgBox strExp & vbCrLf & Err.Description
   End If
   
End Function

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtData
       oText.Locked = bLocked
   Next
   
   If m_EditMode = 1 Then
      txtData(1).Locked = False
      txtData(2).Locked = False
   ElseIf m_EditMode = 2 Then
      txtData(1).Locked = True
      txtData(2).Locked = True
   End If
   
   Me.SSTab1.TabEnabled(1) = bLocked
   
   Combo1.Locked = bLocked
   List1.Enabled = Not bLocked
   cmdAdd.Enabled = Not bLocked
   cmdRemove.Enabled = Not bLocked
   
   cmdCopy(1).Enabled = Not bLocked
   cmdCopy(2).Enabled = Not bLocked
End Sub

'覆蓋案件備註
Private Function GetUpdInstruction(ByVal mKeyTo As String, ByVal mKeyFrom As String, ByVal tTyp As String, ByVal SrcType As String, Optional pExMsg As String) As Boolean
'申請人/代理人編號mKeyTo <=關聯企業mKeyFrom
'tTyp:  F-財務通用、非F->A案件通用
'SrcType 資料來源：1-各項指示, 2-指示=基本檔的欄位
Dim intP As Integer, intQ As Integer
Dim strP1 As String, strP2 As String
Dim strCon1 As String, strCon2 As String
Dim strTmp As String
Dim pKeyFrom As String '判斷欄位的種類(mKeyFrom)
Dim pKeyTo As String '判斷欄位的種類(mKeyTo)
Dim rsPD  As New ADODB.Recordset
Dim rsAD  As New ADODB.Recordset
Dim arrTmp As Variant

GetUpdInstruction = False

On Error GoTo OutPort
    
    pExMsg = ""
    
    If SrcType = "1" Then '各項指示
        pKeyTo = Pub_GetITS01Type(mKeyTo)
        '覆蓋時則將原各項指示都改為無效指示以便留下記錄，覆蓋過來的指示都設定為系統日+1工作天
        'Mark by Lydia 2022/11/25 複製對象編號ITS13：將來源X / Y編號各項指示複製至目的X / Y編號，連同"紀錄日期"一併複製(而非現行的系統日+1工作天)
        'strTmp = "UPDATE INSTRUCTIONS SET ITS05='N', ITS10='" & strUserNum & "', ITS11=" & strSrvDate(1) & ", ITS12=" & Mid(Format(ServerTime, "000000"), 1, 4) & _
        '             " WHERE ITS01=" & CNULL(pKeyTo) & " AND ITS02=" & CNULL(mKeyTo) & " AND ITS05 IS NULL AND ITS04<>" & CompWorkDay(2, strSrvDate(1)) & _
        '             " AND SUBSTR(ITS03,1,1) " & IIf(tTyp = "F", "= 'F'", "<> 'F'")
        'Pub_SeekTbLog strTmp
        'cnnConnection.Execute strTmp, intP
        'end 2022/11/25
        
        strTmp = "SELECT ITS03,ITS04,ITS05,ITS06 FROM INSTRUCTIONS " & _
                 "WHERE NVL(ITS05,'Y')<>'N' AND ITS01=" & CNULL(Pub_GetITS01Type(mKeyFrom)) & " AND ITS02=" & CNULL(mKeyFrom) & _
                 " AND SUBSTR(ITS03,1,1) " & IIf(tTyp = "F", "= 'F'", "<> 'F'") & " ORDER BY 1,2 "
        intP = 1
        Set rsAD = ClsLawReadRstMsg(intP, strTmp)
        If intP = 1 Then
           If rsAD.RecordCount > 0 Then
               rsAD.MoveFirst
               Do While Not rsAD.EOF
                   '複製指示
                   'Modified by Lydia 2022/11/25 判斷目的編號是否存在相同複製指示
                   'strTmp = "INSERT INTO INSTRUCTIONS (ITS01,ITS02,ITS03,ITS04,ITS05,ITS06,ITS07,ITS08,ITS09) " & _
                            "VALUES ('" & pKeyTo & "','" & mKeyTo & "','" & rsAD.Fields("ITS03") & "'," & CompWorkDay(2, strSrvDate(1)) & "," & CNULL("" & rsAD.Fields("ITS05")) & _
                            ",'" & ChgSQL(rsAD.Fields("ITS06")) & "','" & strUserNum & "','" & strSrvDate(1) & "','" & Left(Format(ServerTime, "000000"), 4) & "') "
                   strP1 = "SELECT ITS03,ITS04,ITS05,ITS06 FROM INSTRUCTIONS WHERE ITS01='" & pKeyTo & "' AND ITS02=" & CNULL(mKeyTo) & _
                            " AND SUBSTR(ITS03,1,1) " & IIf(tTyp = "F", "= 'F'", "<> 'F'") & " AND ITS13='" & mKeyFrom & "' AND ITS03='" & rsAD.Fields("ITS03") & "' AND ITS04='" & rsAD.Fields("ITS04") & "' "
                   intQ = 1
                   strTmp = ""
                   Set rsPD = ClsLawReadRstMsg(intQ, strP1)
                   If intQ = 1 Then
                      If "" & rsAD.Fields("ITS06") <> "" & rsPD.Fields("ITS06") Then
                         strTmp = "Update INSTRUCTIONS Set ITS06=" & CNULL(ChgSQL(rsAD.Fields("ITS06"))) & ", ITS10='" & strUserNum & "', ITS11=TO_CHAR(SYSDATE,'YYYYMMDD'), ITS12='" & Left(Format(ServerTime, "000000"), 4) & "' " & _
                                       "Where ITS01='" & pKeyTo & "' AND ITS02=" & CNULL(mKeyTo) & " AND ITS13='" & mKeyFrom & "' AND ITS03='" & rsAD.Fields("ITS03") & "' AND ITS04='" & rsAD.Fields("ITS04") & "' "
                      End If
                   Else
                      strTmp = "INSERT INTO INSTRUCTIONS (ITS01,ITS02,ITS03,ITS04,ITS05,ITS06,ITS07,ITS08,ITS09,ITS13) " & _
                                "VALUES ('" & pKeyTo & "','" & mKeyTo & "','" & rsAD.Fields("ITS03") & "','" & rsAD.Fields("ITS04") & "'," & CNULL("" & rsAD.Fields("ITS05")) & _
                                ",'" & ChgSQL(rsAD.Fields("ITS06")) & "','" & strUserNum & "','" & strSrvDate(1) & "','" & Left(Format(ServerTime, "000000"), 4) & "','" & mKeyFrom & "') "
                   End If
                   If strTmp <> "" Then
                   'end 2022/11/25
                      Pub_SeekTbLog strTmp
                      cnnConnection.Execute strTmp, intP
                   End If 'Added by Lydia 2022/11/25
                   rsAD.MoveNext
               Loop
               GetUpdInstruction = True
           End If
        End If
        
    ElseIf SrcType = "2" Then '直接讀基本檔的欄位進行覆蓋；參考basUpdate.GetInsRpt
        
        '被覆蓋的對象=>判斷欄位的種類(mKeyTo)
        If Left(mKeyTo, 1) = "X" Then
            pKeyTo = "CU"
            mKeyTo = ChangeCustomerL(mKeyTo)
            strCon1 = " update customer set XXX where cu01='" & Mid(mKeyTo, 1, 8) & "' and cu02='" & Mid(mKeyTo, 9, 1) & "' "
        ElseIf Left(mKeyTo, 1) = "Y" Then
            pKeyTo = "FA"
            mKeyTo = ChangeCustomerL(mKeyTo)
            strCon1 = " update fagent set XXX where fa01='" & Mid(mKeyTo, 1, 8) & "' and fa02='" & Mid(mKeyTo, 9, 1) & "' "
        End If
        '來源的對象=>判斷欄位的種類(mKeyFrom)
        If Left(mKeyFrom, 1) = "X" Then
            pKeyFrom = "CU"
            mKeyFrom = ChangeCustomerL(mKeyFrom)
            strCon2 = " from customer where cu01='" & Mid(mKeyFrom, 1, 8) & "' and cu02='" & Mid(mKeyFrom, 9, 1) & "' "
        ElseIf Left(mKeyFrom, 1) = "Y" Then
            pKeyFrom = "FA"
            mKeyFrom = ChangeCustomerL(mKeyFrom)
            strCon2 = " from fagent where fa01='" & Mid(mKeyFrom, 1, 8) & "' and fa02='" & Mid(mKeyFrom, 9, 1) & "' "
        End If
        If pKeyFrom = "" Or pKeyTo = "" Then Exit Function
        
        '逐筆讀取分類檔
        strP1 = "SELECT IT01,IT02,IT03,IT10,IT11,IT12 FROM INSTTYPE WHERE IT11 IS NOT NULL AND IT01 " & IIf(tTyp = "F", "= 'F'", "<> 'F'")
        strP1 = strP1 & " ORDER BY 1,2"
        intP = 1
        Set rsPD = ClsLawReadRstMsg(intP, strP1)
        If intP = 1 Then
            rsPD.MoveFirst
            Do While Not rsPD.EOF
                arrTmp = Empty
                arrTmp = Split(UCase("" & rsPD.Fields("IT11")), ";")
                strP1 = ""
                strP2 = ""
                If InStr("C04,F02", "" & rsPD.Fields("IT01") & rsPD.Fields("IT02")) > 0 Then  '案件平台、帳單平台
                    pExMsg = pExMsg & "【" & rsPD.Fields("IT01") & rsPD.Fields("IT02") & rsPD.Fields("IT03") & "】" & vbCrLf
                Else
                    For intP = 0 To UBound(arrTmp)
                       If arrTmp(intP) <> "" Then
                          If Left(arrTmp(intP), Len(pKeyTo)) = pKeyTo Then
                             strP1 = arrTmp(intP)
                          End If
                          If Left(arrTmp(intP), Len(pKeyFrom)) = pKeyFrom Then
                             strP2 = arrTmp(intP)
                          End If
                       End If
                    Next intP
                         
                    If strP1 <> "" And strP2 <> "" Then
                        strTmp = "select " & strP2 & strCon2
                        intQ = 1
                        Set rsAD = ClsLawReadRstMsg(intQ, strTmp)
                        If intQ = 1 Then
                            If InStr("A01,A02,A03,A04", "" & rsPD.Fields("IT01") & rsPD.Fields("IT02")) > 0 Then  '申請人狀態/代理人狀態
                                 If strP2 = "CU80" Or strP2 = "FA69" Then 'A01不得代理,A03不再使用
                                    If "" & rsAD.Fields(strP2) <> "" & rsPD.Fields("IT03") Then
                                        GoTo JumpToNext
                                    End If
                                 ElseIf strP2 = "CU111" Or strP2 = "FA77" Then 'A02有呆帳
                                    If "" & rsAD.Fields(strP2) <> "Y" Then
                                        GoTo JumpToNext
                                    End If
                                 ElseIf strP2 = "CU142" Or strP2 = "FA103" Then 'A04宣告破產
                                    If "" & rsAD.Fields(strP2) <> "B" Then
                                        GoTo JumpToNext
                                    End If
                                 End If
                            Else
                                '其他-欄位有值才更新
                                If "" & rsAD.Fields(strP2) = "" Then
                                    GoTo JumpToNext
                                End If
                            End If
                            strTmp = Replace(strCon1, "XXX", strP1 & "=" & CNULL("" & rsAD.Fields(strP2)))
                            Pub_SeekTbLog strTmp
                            cnnConnection.Execute strTmp, intP
JumpToNext:
                        End If
                    End If
                End If
                rsPD.MoveNext
            Loop
            GetUpdInstruction = True
            If pExMsg <> "" Then pExMsg = pExMsg & "以上設定請人工處理！"
        End If
    End If
    
OutPort:
   Set rsAD = Nothing
   Set rsPD = Nothing
   If Err.Number <> 0 Then
       pExMsg = Err.Description
   End If
End Function

' 更新 Create 及 Update 的人
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
   'Memo by Lydia 2021/10/21 原程式搬到Form_KeyUp
   
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse txtData(Index)
End Sub

'Modified by Lydia 2021/09/23 改成Form 2.0
'Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index = 1 Or Index = 2 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

'Added by Lydia 2021/09/23 Form 2.0的TextBox增加右鍵選單功能
Private Sub txtData_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtData(Index)
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
Dim iLen As Integer

   Select Case Index
       Case 1, 2
            If txtData(Index) <> "" Then
               If InStr("X,Y", Left(txtData(Index), 1)) = 0 Then
                  MsgBox Mid(Label1(Index), 1, Len(Label1(Index)) - 1) & "必須為X/Y開頭！", vbCritical + vbOKOnly, "檢核資料"
                  GoTo JumpExit
               Else
                  If Len(txtData(Index)) < 6 Then
                     MsgBox Mid(Label1(Index), 1, Len(Label1(Index)) - 1) & "請至少輸入六碼！", vbCritical + vbOKOnly, "檢核資料"
                     GoTo JumpExit
                  Else
                     txtData(Index) = Mid(txtData(Index) & "00", 1, 8)
                     Cancel = Not ReadCuFa(Index)
                  End If
               End If
            End If

   End Select
   
   If Cancel = False Then
      If txtData(Index).MaxLength > 0 Then
         If Not CheckLengthIsOK(txtData(Index), iLen) Then
            GoTo JumpExit
         End If
      End If
   End If
   
   Exit Sub
   
JumpExit:
    Cancel = True
    
End Sub

Private Function ReadCuFa(ByVal iPost As Integer) As Boolean
Dim intM As Integer

   If Left(txtData(iPost), 1) = "X" Then
      strExc(0) = "select cu01 No,replace(cu04,'&','＆') name1,replace(cu05||' '||cu88||' '||cu89||' '||cu90,'&','＆') name2,replace(cu06,'&','＆') name3,cu10 na01,na03,cu80 status" & _
         " from customer,nation where cu01 = '" & Mid(txtData(iPost) & String(8, "0"), 1, 8) & "' and cu02='0' and cu10=na01(+) "
   Else
      strExc(0) = "select fa01 No,replace(fa04,'&','＆') name1,replace(fa05||' '||fa63||' '||fa64||' '||fa65,'&','＆') name2,replace(fa06,'&','＆') name3,fa10 na01,na03,fa69 status" & _
         " from fagent,nation where fa01 = '" & Mid(txtData(iPost) & String(8, "0"), 1, 8) & "' and fa02='0' and fa10=na01(+) "
   End If
   
   intM = 1
   Set RsTemp = ClsLawReadRstMsg(intM, strExc(0))
   If intM = 1 Then
      If iPost = 1 Then
         intM = 0
      Else
         intM = 6
      End If
      
      With RsTemp
          txtData(iPost) = Mid(.Fields("No"), 1, 8)
          Lbl3(intM).Caption = Mid("" & .Fields("na01"), 1, 3)
          Lbl3(intM + 1).Caption = Mid("" & .Fields("na03"), 1, 4)
          Lbl3(intM + 2).Caption = Trim("" & .Fields("status"))
          Lbl3(intM + 3).Caption = Trim("" & .Fields("name1"))
          Lbl3(intM + 4).Caption = Trim("" & .Fields("name2"))
          Lbl3(intM + 5).Caption = Trim("" & .Fields("name3"))
      End With
   Else
      MsgBox "資料庫無資料 !", vbInformation
      'Added by Lydia 2021/09/24
      If iPost = 1 Then
         intM = 0
      Else
         intM = 6
      End If
      txtData(iPost) = Mid(txtData(iPost) & String(8, "0"), 1, 8)
      Lbl3(intM).Caption = ""
      Lbl3(intM + 1).Caption = ""
      Lbl3(intM + 2).Caption = ""
      Lbl3(intM + 3).Caption = ""
      Lbl3(intM + 4).Caption = ""
      Lbl3(intM + 5).Caption = ""
      'end 2021/09/24
      Exit Function
   End If
   
   ReadCuFa = True
End Function

'Added by Lydia 2022/11/25 判斷有複製指示發Email通知
Private Sub ProcSendMail()
Dim strTo As String, strFileName As String
Dim intQ As Integer, strTmpQ As String
Dim rsQuery As New ADODB.Recordset
  
  If m_insCopy1 = True Or m_insCopy2 = True Or m_basCopy1 = True Or m_basCopy2 = True Then
        strFileName = App.path & "\" & strUserNum
        Pub_ChkExcelPath strFileName
        strFileName = strFileName & "\$" & txtData(1) & "各項指示清單.doc"
        If Dir(strFileName) <> "" Then
            RidFile strFileName
        End If
        strTmpQ = "select na51 from fagent ,nation where fa01=" & CNULL(Left(ChangeCustomerL(txtData(1)), 8)) & " and fa02='0' and fa10=na01(+) " & _
                         "union select na51 from customer, nation where cu01=" & CNULL(Left(ChangeCustomerL(txtData(1)), 8)) & " and cu02='0' and cu10=na01(+) "
        intQ = 1
        Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
        If intQ = 1 Then
           strTo = "" & rsQuery.Fields("na51")
        End If
        Set rsQuery = Nothing
        If strTo = "" Then
           MsgBox ChangeCustomerL(txtData(1)) & "查無FCP承辦管制人，無法寄Email通知！", vbInformation
           Exit Sub
        End If
        If PUB_GetITStoList(Me.Name, IIf(Left(txtData(1), 1) = "Y", "1", "2"), txtData(1), True, False, , strFileName) = True Then
           strTmpQ = "國外部關聯企業－複製共通指示" & vbCrLf & vbCrLf & _
                            "申請人/代理人編號：" & ChangeCustomerL(txtData(1)) & " " & IIf(Lbl3(4).Caption <> "", Lbl3(4).Caption, IIf(Lbl3(3).Caption <> "", Lbl3(3).Caption, IIf(Lbl3(5).Caption <> "", Lbl3(5).Caption, ""))) & vbCrLf & _
                            "來源編號：" & ChangeCustomerL(txtData(2)) & " " & IIf(Lbl3(10).Caption <> "", Lbl3(10).Caption, IIf(Lbl3(9).Caption <> "", Lbl3(9).Caption, IIf(Lbl3(11).Caption <> "", Lbl3(11).Caption, ""))) & vbCrLf
           PUB_SendMail strUserNum, strTo, "", ChangeCustomerL(txtData(1)) & "各項指示清單", strTmpQ, , strFileName, False
        End If
  End If
End Sub


