VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160004 
   BorderStyle     =   1  '單線固定
   Caption         =   "出差資料"
   ClientHeight    =   5040
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8170
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8170
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
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
            Picture         =   "frm160004.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160004.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160004.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160004.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160004.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160004.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160004.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160004.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160004.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160004.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160004.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8170
      _ExtentX        =   14411
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
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4350
      Left            =   30
      TabIndex        =   21
      Top             =   660
      Width           =   8115
      _ExtentX        =   14323
      _ExtentY        =   7673
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm160004.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(17)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label14"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textSB01_2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textSB09"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtNote"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label23"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "LblIsApart"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textSB06"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textSB03_2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textSB02"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textSB03_1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textSB05_1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textSB05_2"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textSB08"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textSB04"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textSB07"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textSB01"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Frame1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdABS"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160004.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdok"
      Tab(1).Control(1)=   "txt1(3)"
      Tab(1).Control(2)=   "txt1(2)"
      Tab(1).Control(3)=   "txt1(1)"
      Tab(1).Control(4)=   "txt1(0)"
      Tab(1).Control(5)=   "GRD1"
      Tab(1).Control(6)=   "Line5"
      Tab(1).Control(7)=   "Label16"
      Tab(1).Control(8)=   "Line4"
      Tab(1).Control(9)=   "Label15"
      Tab(1).ControlCount=   10
      Begin VB.CommandButton cmdABS 
         BackColor       =   &H00C0FFFF&
         Caption         =   "簽核資料"
         Height          =   315
         Left            =   5610
         Style           =   1  '圖片外觀
         TabIndex        =   47
         Top             =   1800
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   1005
         Left            =   30
         TabIndex        =   38
         Top             =   2640
         Visible         =   0   'False
         Width           =   2535
         Begin VB.ComboBox cboETime 
            Height          =   300
            ItemData        =   "frm160004.frx":212C
            Left            =   1170
            List            =   "frm160004.frx":212E
            Locked          =   -1  'True
            Style           =   2  '單純下拉式
            TabIndex        =   13
            Top             =   690
            Width           =   1005
         End
         Begin VB.ComboBox cboSTime 
            Height          =   300
            ItemData        =   "frm160004.frx":2130
            Left            =   1170
            List            =   "frm160004.frx":2132
            Locked          =   -1  'True
            Style           =   2  '單純下拉式
            TabIndex        =   12
            Top             =   360
            Width           =   1005
         End
         Begin VB.TextBox textSB10 
            Height          =   285
            Left            =   1170
            MaxLength       =   8
            TabIndex        =   11
            Top             =   60
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "迄日下班時段"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   41
            Top             =   750
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "起日上班時段"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   40
            Top             =   420
            Width           =   1080
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "表單編號"
            Height          =   180
            Left            =   420
            TabIndex        =   39
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   345
         Left            =   -68700
         TabIndex        =   19
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -70080
         MaxLength       =   7
         TabIndex        =   18
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -71070
         MaxLength       =   7
         TabIndex        =   17
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72960
         MaxLength       =   6
         TabIndex        =   16
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -74010
         MaxLength       =   6
         TabIndex        =   15
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox textSB01 
         Height          =   270
         Left            =   930
         MaxLength       =   6
         TabIndex        =   0
         Top             =   390
         Width           =   735
      End
      Begin VB.TextBox textSB07 
         Height          =   315
         Left            =   2010
         MaxLength       =   4
         TabIndex        =   10
         Top             =   2310
         Width           =   525
      End
      Begin VB.TextBox textSB04 
         Height          =   270
         Left            =   3690
         MaxLength       =   7
         TabIndex        =   6
         Top             =   1575
         Width           =   945
      End
      Begin VB.TextBox textSB08 
         Height          =   270
         Left            =   930
         MaxLength       =   1
         TabIndex        =   1
         Top             =   690
         Width           =   225
      End
      Begin VB.TextBox textSB05_2 
         Height          =   285
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1890
         Width           =   585
      End
      Begin VB.TextBox textSB05_1 
         Height          =   285
         Left            =   3210
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1890
         Width           =   585
      End
      Begin VB.TextBox textSB03_1 
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1890
         Width           =   585
      End
      Begin VB.TextBox textSB02 
         Height          =   270
         Left            =   1350
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1575
         Width           =   945
      End
      Begin VB.TextBox textSB03_2 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1890
         Width           =   585
      End
      Begin VB.TextBox textSB06 
         Height          =   315
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   9
         Top             =   2310
         Width           =   525
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160004.frx":2134
         Height          =   3615
         Left            =   -74970
         TabIndex        =   22
         Top             =   690
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6368
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
      Begin MSForms.Label LblIsApart 
         Height          =   195
         Left            =   210
         TabIndex        =   46
         Top             =   3720
         Width           =   7515
         ForeColor       =   8388736
         VariousPropertyBits=   27
         Caption         =   "拆單，此筆資料為 : "
         Size            =   "13256;344"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         FontWeight      =   700
      End
      Begin MSForms.Label Label23 
         Height          =   195
         Left            =   150
         TabIndex        =   45
         Top             =   4020
         Width           =   7785
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13732;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNote 
         Height          =   1100
         Left            =   3930
         TabIndex        =   14
         Top             =   2640
         Width           =   3350
         VariousPropertyBits=   -1466939365
         MaxLength       =   100
         ScrollBars      =   3
         Size            =   "5900;1931"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSB09 
         Height          =   285
         Left            =   930
         TabIndex        =   2
         Top             =   990
         Width           =   3345
         VariousPropertyBits=   679495707
         MaxLength       =   20
         Size            =   "5900;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label textSB01_2 
         Height          =   225
         Left            =   1710
         TabIndex        =   44
         Top             =   420
         Width           =   1395
         BackColor       =   12632256
         VariousPropertyBits=   27
         Size            =   "2461;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "修改/刪除原因："
         Height          =   180
         Left            =   2640
         TabIndex        =   43
         Top             =   2670
         Width           =   1310
      End
      Begin VB.Label Label14 
         Caption         =   "註：人事室異動假單資料跨月份時，記得要拆單輸入，因會影響算薪水！"
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3930
         TabIndex        =   42
         Top             =   420
         Width           =   4065
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   3240
         X2              =   4920
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line5 
         X1              =   -70440
         X2              =   -69690
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   -71670
         TabIndex        =   37
         Top             =   390
         Width           =   540
      End
      Begin VB.Line Line4 
         X1              =   -73290
         X2              =   -72690
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74940
         TabIndex        =   36
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號"
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   35
         Top             =   435
         Width           =   720
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   960
         X2              =   4950
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日期"
         Height          =   180
         Index           =   2
         Left            =   3270
         TabIndex        =   34
         Top             =   1620
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "地點"
         Height          =   180
         Left            =   540
         TabIndex        =   33
         Top             =   990
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "1：長程 2：短程 3：大陸 4：國外"
         Height          =   180
         Left            =   1170
         TabIndex        =   32
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "差程"
         Height          =   180
         Left            =   570
         TabIndex        =   31
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "分"
         Height          =   180
         Left            =   4710
         TabIndex        =   30
         Top             =   1950
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "時"
         Height          =   180
         Left            =   3840
         TabIndex        =   29
         Top             =   1950
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "時"
         Height          =   180
         Left            =   1590
         TabIndex        =   28
         Top             =   1950
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日期"
         Height          =   180
         Index           =   17
         Left            =   930
         TabIndex        =   27
         Top             =   1620
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "時間                起                                              迄"
         Height          =   180
         Left            =   570
         TabIndex        =   26
         Top             =   1290
         Width           =   3510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "分"
         Height          =   180
         Left            =   2430
         TabIndex        =   25
         Top             =   1950
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2790
         TabIndex        =   24
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "共              日               時"
         Height          =   180
         Left            =   960
         TabIndex        =   23
         Top             =   2370
         Width           =   1845
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   990
         X2              =   2670
         Y1              =   1500
         Y2              =   1500
      End
   End
End
Attribute VB_Name = "frm160004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/16 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/20 日期欄已修改
'Create by nickc 2006/11/28 copy from frm140401
Option Explicit

Dim RcMain As New ADODB.Recordset, RsAdo As New ADODB.Recordset
' 變數宣告區
Dim m_EditMode As Integer
Dim m_SubMode As Integer
'(執行各項功能的權限)
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
' 第一筆資料的本所案號
Dim m_FirstKEY(3) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(3) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(3) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_SB As Integer
Dim MyKind As String
'Add By Sindy 2011/9/22
Dim m_B1019 As String, m_B1004 As String, m_B1005 As String
Dim m_B1006 As String, m_B1007 As String, m_B1009 As String
Dim m_B1010 As String, m_B1014 As String, m_B1015 As String
Dim m_B1017 As String, m_B1028 As String, m_B1029 As String
Dim m_KeyCode As String 'Add By Sindy 2011/10/7
Dim BolIsApart As String '是否有拆單


'Add By Sindy 2022/10/28
Private Sub cmdABS_Click()
   Me.Hide
   Call frm180301_03.SetParent(Me)
   frm180301_03.txtB1001 = textSB10.Text
   frm180301_03.QueryData
   frm180301_03.Show
End Sub

Private Sub cmdok_Click()
If txt1(0) & txt1(1) & txt1(2) & txt1(3) <> "" Then
    If RunNick(txt1(0), txt1(1)) Then
        txt1(0).SetFocus
        Exit Sub
    End If
    If RunNick2(txt1(2), txt1(3)) Then
        txt1(2).SetFocus
        Exit Sub
    End If
    GetData
Else
    MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
End If
End Sub

Private Sub Form_Initialize()
Set rsA = New ADODB.Recordset
If rsA.State = 1 Then rsA.Close
rsA.CursorLocation = adUseClient
rsA.Open "select * from staff_busi_trip where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
tf_SB = rsA.Fields.Count
SetGrd
End Sub

' 按下按鍵
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
'edit by nickc 2006/11/13
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub
'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   ReDim m_FieldList(tf_SB) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textSB01.BackColor = &H8000000F
   textSB02.BackColor = &H8000000F
   ' 2008/12/24 Add BY SINDY
   textSB03_1.BackColor = &H8000000F
   textSB03_2.BackColor = &H8000000F
   ' 2008/12/24 END
   
   MoveFormToCenter Me
   
   InitialField
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
   
   'Add By Sindy 2021/8/11
   SetB102829Combo cboSTime, 1
   SetB102829Combo cboETime, 2
   '2021/8/11 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160004 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j
GRD1.Visible = False
tmpMouseRow = GRD1.row
GRD1.Visible = True
If tmpMouseRow <> 0 Then
    GRD1.row = tmpMouseRow
    GRD1.col = 0
    If GRD1.CellBackColor <> &HFFC0C0 Then
                  GRD1.Visible = False
         For j = 1 To GRD1.Rows - 1
             GRD1.row = j
             For i = 0 To GRD1.Cols - 1
                  GRD1.col = i
                  GRD1.CellBackColor = QBColor(15)
             Next i
        Next j
        GRD1.row = tmpMouseRow
         For i = 0 To GRD1.Cols - 1
             GRD1.col = i
             GRD1.CellBackColor = &HFFC0C0
         Next i
         '2008/12/12 ADD BY SONIA
         textSB01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         textSB02.Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 2))
         ' 2008/12/24 Add BY SINDY
         textSB03_1.Text = Mid(Trim(GRD1.TextMatrix(tmpMouseRow, 2)), Len(Trim(GRD1.TextMatrix(tmpMouseRow, 2))) - 4, 2)
         textSB03_2.Text = Right(Trim(GRD1.TextMatrix(tmpMouseRow, 2)), 2)
         ' 2008/12/24 END
         QueryRecord
         '2008/12/12 END
         GRD1.Visible = True
    End If
End If
End Sub

'Add By Sindy 2019/8/27
Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      cmdok.SetFocus
      cmdok.Default = True
   Else
      cmdok.Default = False
   End If
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("sb11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sb11")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("sb11"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sb12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sb12")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("sb12"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sb13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sb13")) = False Then
         strTemp = rsSrcTmp.Fields("sb13")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sb14")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sb14")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("sb14"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sb15")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sb15")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("sb15"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sb16")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sb16")) = False Then
         strTemp = rsSrcTmp.Fields("sb16")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

If Me.textSB01.Enabled = True Then
   Cancel = False
   textsb01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textSB01.Text = "" Then
    MsgBox "員工編號不可以空白！", vbExclamation
    textSB01.SetFocus
    Exit Function
End If

'Add By Sindy 2011/9/22
If textSB08.Text = "" Then
   MsgBox "差程不可以空白！", vbExclamation
   textSB08.SetFocus
   Exit Function
End If

'Add By Sindy 2011/9/22
If Me.Frame1.Visible = True And Me.textSB10.Enabled = True Then
   If m_EditMode = 1 And textSB10 = "" Then
      MsgBox "表單編號不可空白！", vbExclamation
      textSB10.SetFocus
      Exit Function
   End If
   
   Cancel = False
   textSB10_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSB02.Enabled = True Then
   Cancel = False
   textSb02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textSB02.Text = "" Then
    MsgBox "日期起不可以空白！", vbExclamation
    textSB02.SetFocus
    Exit Function
End If
If Me.textSB03_1.Enabled = True Then
   Cancel = False
   textSb03_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textSB03_1.Text = "" Then
    MsgBox "起始(時)不可以空白！", vbExclamation
    textSB03_1.SetFocus
    Exit Function
End If
If Me.textSB03_2.Enabled = True Then
   Cancel = False
   textSb03_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textSB03_2.Text = "" Then
    MsgBox "起始(分)不可以空白！", vbExclamation
    textSB03_2.SetFocus
    Exit Function
End If
If Me.textSB04.Enabled = True Then
   Cancel = False
   textSb04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textSB05_1.Enabled = True Then
   Cancel = False
   textSb05_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textSB05_2.Enabled = True Then
   Cancel = False
   textSb05_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2014/6/24
If Left(textSB02, Len(textSB02) - 2) <> Left(textSB04, Len(textSB04) - 2) Then
   MsgBox "假單不可跨月份！", vbExclamation
   textSB04.SetFocus
   Exit Function
End If
'2014/6/24 END

'Add By Sindy 2011/10/17 增加判斷員工代號+日期是否人員已離職
If ChkStaffST04(textSB01, True, textSB02) = True Then
   textSB01.SetFocus
   Exit Function
End If
'Add By Sindy 2011/10/17 增加判斷員工代號+日期是否人員已離職
If ChkStaffST04(textSB01, True, textSB04) = True Then
   textSB01.SetFocus
   Exit Function
End If

If Me.textSB06.Enabled = True Then
   Cancel = False
   textSB06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textSB07.Enabled = True Then
   Cancel = False
   textSB07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textSB06.Text = "" Or textSB07.Text = "" Or _
   (textSB06.Text = "0" And textSB07.Text = "0") Then
    MsgBox "無出差時數！", vbExclamation
    textSB06.SetFocus
    Exit Function
End If
If Me.textSB08.Enabled = True Then
   Cancel = False
   textSB08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textSB09.Enabled = True Then
   Cancel = False
   textSB09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2011/9/22
If Frame1.Visible = True Then
   If Me.cboSTime.Enabled = True Then
      Cancel = False
      cboSTime_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.cboETime.Enabled = True Then
      Cancel = False
      cboETime_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
End If

'Add by Sindy 2021/9/1 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me) = False Then
   Exit Function
End If
'2021/9/1 END

TxtValidate = True
End Function

'add by nickc 2006/10/24
' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To tf_SB - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To tf_SB - 1
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

' 新增記錄
Private Function AddRecord() As Boolean
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strSB01 As String
   Dim strSB02 As String
   Dim strSB03 As String
   
   AddRecord = False
   
   strSB01 = textSB01
   strSB02 = DBDATE(Trim(textSB02))
   strSB03 = Trim(textSB03_1.Text & textSB03_2.Text) ' 2008/12/24 Add BY SINDY
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strSB01, strSB02, strSB03) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO staff_busi_trip ("
   For nIndex = 0 To tf_SB - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To tf_SB - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   'Add By Sindy 2011/9/22
   If Frame1.Visible = True And textSB10 <> "" Then
      Call ProABSData
   Else
      'Add By Sindy 2019/5/24 假單完成,後續資料檢查及SendMail
      Call PUB_AutoM21Receive_SendMail(IIf(textSB10 <> "", textSB10, ""), 表單類別_出差, textSB01, DBDATE(textSB02), Trim(Format("00" & textSB03_1, "00") & Format("00" & textSB03_2, "00")), _
         DBDATE(textSB04), textSB08, Left(DBDATE(textSB02), 6), , , , m_EditMode)
   End If
   
   cnnConnection.CommitTrans
   ' 2008/12/24 Modify BY SINDY
   If ((strSB01 & strSB02 & strSB03) < (m_FirstKEY(0) & m_FirstKEY(1) & m_FirstKEY(2))) Or ((strSB01 & strSB02 & strSB03) > (m_LastKEY(0) & m_LastKEY(1) & m_LastKEY(2))) Then
   ' 2008/12/24 END
      RefreshRange
   End If
   
   ShowCurrRecord strSB01, DBDATE(strSB02), strSB03
   AddRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox " 新增失敗！" & vbCrLf & Err.Description
    
End Function

' 修改記錄
Private Function ModRecord() As Boolean
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strSB01 As String
   Dim strSB02 As String
   Dim strSB03 As String
   
   ModRecord = False
   
   strSB01 = m_CurrKEY(0)
   strSB02 = m_CurrKEY(1)
   strSB03 = m_CurrKEY(2)
   
   'Modify By Sindy 2023/11/1 mark,前面有檢查,此處應該不用
'   'Add By SINDY 2011/12/5
'   If strSB01 <> textSB01 Or _
'      strSB02 <> DBDATE(textSB02) Or _
'      Val(strSB03) <> Val(Trim(Format("00" & textSB03_1, "00") & Format("00" & textSB03_2, "00"))) Then
'      ' 檢查記錄是否已存在
'      If IsRecordExist(textSB01, DBDATE(textSB02), Trim(Format("00" & textSB03_1, "00") & Format("00" & textSB03_2, "00"))) = True Then
'         strTit = "新增資料"
'         strMsg = "該筆記錄已存在"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         'UpdateCtrlData
'         textSB02.SetFocus
'         Exit Function
'      End If
'   End If
   
   strSql = "begin user_data.user_enabled:=1; UPDATE staff_busi_trip SET "
   
   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_SB - 1
      strTmp = Empty
      'If nIndex < 10 Or nIndex > 15 Then
            If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
               If m_FieldList(nIndex).fiType = 0 Then
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
                     strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
                  End If
               Else
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
                     strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
                  End If
               End If
            End If
            If strTmp <> Empty Then
               bDifference = True
               If bFirst = True Then
                  strSql = strSql & strTmp
                  bFirst = False
               Else
                  strSql = strSql & "," & strTmp
               End If
            End If
        'End If
   Next nIndex
   
   strSql = strSql & " " & _
                  "WHERE sb01 = '" & strSB01 & "' and sb02='" & strSB02 & "' and sb03='" & strSB03 & "' ; end; "
On Error GoTo ErrHand
      cnnConnection.BeginTrans
        If bDifference = True Then
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
           'Add By Sindy 2011/9/22
           'Modify By Sindy 2019/5/24 電子紙本均要考慮發信問題
           'If Frame1.Visible = True And textSB10 <> "" Then
               Call ProABSData
           'End If
        End If
        
        cnnConnection.CommitTrans
      ShowCurrRecord strSB01, DBDATE(strSB02), strSB03
      
    ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)

End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim strSql As String
   Dim strSB01 As String
   Dim strSB02 As String
   Dim strSB03 As String
   'Add By Sindy 2013/2/1
   Dim rsTmp As New ADODB.Recordset
   Dim nResponse
   '2013/2/1 End
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   strSB01 = m_CurrKEY(0)
   strSB02 = m_CurrKEY(1)
   strSB03 = m_CurrKEY(2)
   'Add By Sindy 2013/2/1
   BolIsApart = False
   If textSB10.Text <> "" Then
      strSql = "SELECT * FROM staff_busi_trip " & _
               "WHERE SB10 = '" & textSB10.Text & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 1 Then
         BolIsApart = True
      End If
      rsTmp.Close
   End If
   If BolIsApart = True Then
      nResponse = MsgBox("此假單是跨月份拆單而產生的資料" & vbCrLf & _
                         "欲刪除資料，請在備註欄位裡註明是那一天要銷假，確定要刪除了嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問")
      If nResponse = vbNo Then
         Exit Function
      End If
   End If
   '2013/2/1 End
   
   cnnConnection.BeginTrans
   
   'Add By Sindy 2011/9/22
   'Modify By Sindy 2019/5/24 電子紙本均要考慮發信問題
'   If Frame1.Visible = True And textSB10.Text <> "" And m_B1019 <> "" Then
      PUB_FilterFormText Me 'Add by Sindy 2011/10/14 修正畫面所有含跳行符號的文字框
      'MsgBox "電子表單人事處已簽收,不可在此作業刪除！", vbExclamation
      'Exit Function
      'Call DelMark
      'Modify By Sindy 2013/2/1 檢查此表單編號必須全部刪除才可上註銷
      If BolIsApart = False Then
         Call DelMark
      Else
         Call ProABSData(True)
      End If
      '2013/2/1 End
'   End If
   
   strSql = "DELETE FROM staff_busi_trip " & _
            "WHERE sb01 = '" & strSB01 & "'  and sb02='" & strSB02 & "' and sb03='" & strSB03 & "' "
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   If (strSB01 = m_LastKEY(0) And strSB02 = m_LastKEY(1) And strSB03 = m_LastKEY(2)) Or (strSB01 = m_FirstKEY(0) And strSB02 = m_FirstKEY(1) And strSB03 = m_FirstKEY(2)) Then
      RefreshRange
   End If
   ShowCurrRecord strSB01, DBDATE(strSB02), strSB03
   DelRecord = True
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strSB01 As String
   Dim strSB02 As String
   Dim strSB03 As String
   
   QueryRecord = False
   strSB01 = textSB01
   strSB02 = DBDATE(Trim(textSB02))
   strSB03 = Trim(textSB03_1.Text & textSB03_2.Text) ' 2008/12/24 Add BY SINDY
   
   If IsRecordExist(strSB01, strSB02, strSB03) = True Then
      m_CurrKEY(0) = strSB01
      m_CurrKEY(1) = strSB02
      m_CurrKEY(2) = strSB03
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If AddRecord = True Then
                RefreshRange
            Else
                Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '修改
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         'Add By Sindy 2012/6/20
         If Trim(txtNote) = "" And Frame1.Visible = True Then
            MsgBox "刪除原因不可空白！", vbExclamation
            txtNote.SetFocus
            Exit Function
         End If
         '2012/6/20 End
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1), m_CurrKEY(2)
         Else
            Exit Function
         End If
      Case 4: '查詢
         ' 2008/12/22 Modify BY SINDY
         'If textSB01 <> "" And textSB02 <> "" Then
         If textSB01 <> "" And textSB02 <> "" _
            And textSB03_1 <> "" And textSB03_2 <> "" Then
         ' 2008/12/22 END
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            ' 2008/12/17 ADD BY SINDY
            If textSB01 = "" Or textSB02 = "" Or _
               textSB03_1 = "" Or textSB03_2 = "" Then
               MsgBox "須輸入員工代號及起始(日期)和起始(時)(分)才可進行查詢動作！", vbInformation
            End If
            ' 2008/12/17 END
            
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: If Me.Visible = True Then textSB01.SetFocus
      Case 2: If Me.Visible = True Then textSB03_1.SetFocus
      Case 4: If Me.Visible = True Then textSB01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT sb01 FROM staff_busi_trip" & _
            " WHERE sb01='" & strKEY01 & "'  and sb02='" & strKEY02 & "' and sb03='" & strKEY03 & "'"
   'Add By Sindy 2023/11/1 修改時,排除檢查此表單
   If m_EditMode = 2 Then
      strSql = strSql & " and sb10<>'" & textSB10 & "'"
   End If
   '2023/11/1 END
   'Add By Sindy 2022/7/19
   'Modify By Sindy 2023/11/29 +已核准
   strSql = strSql & " union SELECT B1003 FROM abs010" & _
            " WHERE B1003 = '" & strKEY01 & "'" & _
            " and (" & strKEY02 & Right("0000" & strKEY03, 4) & " between B1004||substr('0'||B1005,-4) and B1006||substr('0'||B1007,-4)" & _
            " or " & DBDATE(textSB04) & Format("0" & textSB05_1, "00") & Format("0" & textSB05_2, "00") & " between B1004||substr('0'||B1005,-4) and B1006||substr('0'||B1007,-4)" & _
            ") and B1018 not in('" & 註銷 & "','" & 已核准 & "')"
   '2022/7/19 END
   'Add By Sindy 2023/11/1 修改時,排除檢查此表單
   'Modify By Sindy 2025/1/6 新增時,排除檢查此表單
   'Modify By Sindy 2025/1/21 + And Frame1.Visible = True
   If (m_EditMode = 2 Or m_EditMode = 1) And Frame1.Visible = True Then
      strSql = strSql & " and B1001<>'" & textSB10 & "'"
   End If
   '2023/11/1 END
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02, strKEY03) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
      m_CurrKEY(2) = strKEY03
   Else
      strSql = "SELECT sb01,sb02,sb03 FROM staff_busi_trip " & _
               "WHERE sb01 = '" & m_CurrKEY(0) & "' and sb02='" & m_CurrKEY(1) & "' and sb03='" & m_CurrKEY(2) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("sb01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sb01")
         If IsNull(rsTmp.Fields("sb02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sb02")
         If IsNull(rsTmp.Fields("sb03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("sb03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT sb01,sb02,sb03 FROM staff_busi_trip " & _
               "WHERE sb02 = (SELECT MIN(sb02) FROM staff_busi_trip where sb01=(select min(sb01) from staff_busi_trip) ) and sb01=(select min(sb01) from staff_busi_trip) Order BY sb03 ASC "
      
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("sb01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sb01")
         If IsNull(rsTmp.Fields("sb02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sb02")
         If IsNull(rsTmp.Fields("sb03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("sb03")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   m_CurrKEY(2) = m_FirstKEY(2)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) And m_CurrKEY(2) = m_FirstKEY(2) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   ' 2008/12/24 Add BY SINDY
   strSql = "SELECT sb01,sb02,sb03 FROM staff_busi_trip " & _
            "WHERE sb01 = '" & m_CurrKEY(0) & "' AND " & _
                          "sb02 = '" & m_CurrKEY(1) & "' AND " & _
                  "sb03 = (SELECT MAX(sb03) FROM staff_busi_trip " & _
                          "WHERE sb01 = '" & m_CurrKEY(0) & "' AND " & _
                                        "sb02 = '" & m_CurrKEY(1) & "' AND " & _
                                        "sb03 < '" & m_CurrKEY(2) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sb01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sb01")
      If IsNull(rsTmp.Fields("sb02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sb02")
      If IsNull(rsTmp.Fields("sb03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("sb03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   ' 2008/12/24 END
   
   strSql = "SELECT sb01,sb02,sb03 FROM staff_busi_trip " & _
            "WHERE sb01 = '" & m_CurrKEY(0) & "' AND " & _
                  "sb02 = (SELECT MAX(sb02) FROM staff_busi_trip " & _
                          "WHERE sb01 = '" & m_CurrKEY(0) & "' AND " & _
                                "sb02 < '" & m_CurrKEY(1) & "') Order BY sb03 DESC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sb01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sb01")
      If IsNull(rsTmp.Fields("sb02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sb02")
      If IsNull(rsTmp.Fields("sb03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("sb03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT sb01,sb02,sb03 FROM staff_busi_trip " & _
            "WHERE sb01 = (SELECT MAX(sb01) FROM staff_busi_trip " & _
                           "WHERE sb01 < '" & m_CurrKEY(0) & "') AND " & _
                  "sb02 = (SELECT MAX(sb02) FROM staff_busi_trip " & _
                           "WHERE sb01 = (SELECT MAX(sb01) FROM staff_busi_trip " & _
                                          "WHERE sb01 < '" & m_CurrKEY(0) & "')) Order BY sb03 DESC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sb01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sb01")
      If IsNull(rsTmp.Fields("sb02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sb02")
      If IsNull(rsTmp.Fields("sb03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("sb03")
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) And m_CurrKEY(2) = m_LastKEY(2) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   ' 2008/12/24 Add BY SINDY
   strSql = "SELECT sb01,sb02,sb03 FROM staff_busi_trip " & _
            "WHERE sb01 = '" & m_CurrKEY(0) & "' AND " & _
                          "sb02 = '" & m_CurrKEY(1) & "' AND " & _
                  "sb03 = (SELECT MIN(sb03) FROM staff_busi_trip " & _
                          "WHERE sb01 = '" & m_CurrKEY(0) & "' AND " & _
                                        "sb02 = '" & m_CurrKEY(1) & "' AND " & _
                                        "sb03 > '" & m_CurrKEY(2) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sb01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sb01")
      If IsNull(rsTmp.Fields("sb02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sb02")
      If IsNull(rsTmp.Fields("sb03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("sb03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   ' 2008/12/24 END
   
   strSql = "SELECT sb01,sb02,sb03 FROM staff_busi_trip " & _
            "WHERE sb01 = '" & m_CurrKEY(0) & "' AND " & _
                  "sb02 = (SELECT MIN(sb02) FROM staff_busi_trip " & _
                          "WHERE sb01 = '" & m_CurrKEY(0) & "' AND " & _
                                "sb02 > '" & m_CurrKEY(1) & "') Order BY sb03 ASC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sb01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sb01")
      If IsNull(rsTmp.Fields("sb02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sb02")
      If IsNull(rsTmp.Fields("sb03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("sb03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT sb01,sb02,sb03 FROM staff_busi_trip " & _
            "WHERE sb01 = (SELECT MIN(sb01) FROM staff_busi_trip " & _
                           "WHERE sb01 > '" & m_CurrKEY(0) & "') AND " & _
                  "sb02 = (SELECT MIN(sb02) FROM staff_busi_trip " & _
                           "WHERE sb01 = (SELECT MIN(sb01) FROM staff_busi_trip " & _
                                          "WHERE sb01 > '" & m_CurrKEY(0) & "')) Order BY sb03 ASC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sb01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sb01")
      If IsNull(rsTmp.Fields("sb02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sb02")
      If IsNull(rsTmp.Fields("sb03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("sb03")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   m_CurrKEY(2) = m_LastKEY(2)
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   m_SubMode = 0
   m_KeyCode = KeyCode 'Add By Sindy 2011/10/7
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         'Add By Sindy 2013/2/1
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         '2013/2/1 End
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         If OnWork = True Then
            Me.SSTab1.TabEnabled(1) = True
            UpdateToolbarState
         Else
            Exit Sub
         End If
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  Me.SSTab1.TabEnabled(1) = True
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               Me.SSTab1.TabEnabled(1) = True
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT sb01,sb02,sb03 FROM staff_busi_trip " & _
            "WHERE sb01 = (SELECT MIN(sb01) FROM staff_busi_trip) AND " & _
                  "sb02 = (SELECT MIN(sb02) FROM staff_busi_trip " & _
                           "WHERE sb01 = (SELECT MIN(sb01) FROM staff_busi_trip)) Order BY sb03 ASC "
                           
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sb01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("sb01")
      If IsNull(rsTmp.Fields("sb02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("sb02")
      If IsNull(rsTmp.Fields("sb03")) = False Then: m_FirstKEY(2) = rsTmp.Fields("sb03")
   End If
   rsTmp.Close

   strSql = "SELECT sb01,sb02,sb03 FROM staff_busi_trip " & _
            "WHERE sb01 = (SELECT MAX(sb01) FROM staff_busi_trip) AND " & _
                  "sb02 = (SELECT MAX(sb02) FROM staff_busi_trip " & _
                           "WHERE sb01 = (SELECT MAX(sb01) FROM staff_busi_trip)) Order BY sb03 DESC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sb01")) = False Then: m_LastKEY(0) = rsTmp.Fields("sb01")
      If IsNull(rsTmp.Fields("sb02")) = False Then: m_LastKEY(1) = rsTmp.Fields("sb02")
      If IsNull(rsTmp.Fields("sb03")) = False Then: m_LastKEY(2) = rsTmp.Fields("sb03")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim i As Integer, j As Integer
   
   strSql = "SELECT * FROM staff_busi_trip " & _
            "WHERE sb01='" & m_CurrKEY(0) & "' and sb02 = '" & m_CurrKEY(1) & "' and sb03 = '" & m_CurrKEY(2) & "' "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("sb01")) = False Then: textSB01 = rsTmp.Fields("sb01")
      If IsNull(rsTmp.Fields("sb02")) = False Then: textSB02 = TAIWANDATE(rsTmp.Fields("sb02"))
      If IsNull(rsTmp.Fields("sb03")) = False Then: textSB03_1 = Mid(Format(rsTmp.Fields("sb03"), "0000"), 1, 2): textSB03_2 = Mid(Format(rsTmp.Fields("sb03"), "0000"), 3, 2)
      If IsNull(rsTmp.Fields("sb04")) = False Then: textSB04 = TAIWANDATE(rsTmp.Fields("sb04"))
      If IsNull(rsTmp.Fields("sb05")) = False Then: textSB05_1 = Mid(Format(rsTmp.Fields("sb05"), "0000"), 1, 2): textSB05_2 = Mid(Format(rsTmp.Fields("sb05"), "0000"), 3, 2)
      If IsNull(rsTmp.Fields("sb06")) = False Then: textSB06 = rsTmp.Fields("sb06")
      If IsNull(rsTmp.Fields("sb07")) = False Then: textSB07 = rsTmp.Fields("sb07")
      If IsNull(rsTmp.Fields("sb08")) = False Then: textSB08 = rsTmp.Fields("sb08")
      If IsNull(rsTmp.Fields("sb09")) = False Then: textSB09 = rsTmp.Fields("sb09")
      
      'Add By Sindy 2021/8/11
      SetB102829Combo cboSTime, 1, textSB02, textSB01
      SetB102829Combo cboETime, 2, textSB02, textSB01
      '2021/8/11 END
      
      'Add By Sindy 2021/12/27
      LblIsApart.Caption = "跨月份拆單，此筆資料為：" & ChangeTStringToTDateString(textSB02) & " " & textSB03_1 & ":" & textSB03_2 & " ~ " & ChangeTStringToTDateString(textSB04) & " " & textSB05_1 & ":" & textSB05_2
      '2021/12/27 End
      
      'Add By Sindy 2011/9/22
      If IsNull(rsTmp.Fields("SB10")) = False Then
         textSB10 = rsTmp.Fields("SB10")
         Call GetABS010(True)
         Frame1.Visible = True
         
         'Add By Sindy 2021/12/27
         BolIsApart = False
         If textSB10.Text <> "" Then
            strSql = "SELECT * FROM staff_Absence " & _
                     "WHERE SA09 = '" & textSB10.Text & "'"
            If RsTemp.State = 1 Then RsTemp.Close
            RsTemp.CursorLocation = adUseClient
            RsTemp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If RsTemp.RecordCount > 1 Then
               BolIsApart = True
            End If
            RsTemp.Close
         End If
         If BolIsApart = True Then
            LblIsApart.Visible = True
         End If
         '2021/12/27 End
      Else
         'Add By Sindy 2019/5/24 記錄原始資料
         m_B1004 = DBDATE(textSB02)
         m_B1005 = textSB03_1 & textSB03_2
         m_B1006 = DBDATE(textSB04)
         m_B1007 = textSB05_1 & textSB05_2
         m_B1009 = textSB06
         m_B1010 = textSB07
         m_B1014 = textSB08
         m_B1015 = textSB09
         If Not IsNull(rsTmp.Fields("SB17")) Then m_B1028 = IIf(Format(rsTmp.Fields("SB17"), "hhmm") = "0000", "", Format(rsTmp.Fields("SB17"), "hhmm"))
         If Not IsNull(rsTmp.Fields("SB18")) Then m_B1029 = IIf(Format(rsTmp.Fields("SB18"), "hhmm") = "0000", "", Format(rsTmp.Fields("SB18"), "hhmm"))
         '2019/5/24 END
         Frame1.Visible = False
      End If
      'Add By Sindy 2013/2/1
      '顯示起日上班時段,迄日下班時段至畫面上
      If Not IsNull(rsTmp.Fields("SB17")) Then
         For i = 0 To cboSTime.ListCount - 1
            If cboSTime.List(i) = Format(rsTmp.Fields("SB17"), "00:00") Then 'Format(Format(rsTmp.Fields("SB17"), "0000"), "##:##") Then
               cboSTime.ListIndex = i
               Exit For
            End If
         Next i
      End If
      If Not IsNull(rsTmp.Fields("SB18")) Then
         For i = 0 To cboETime.ListCount - 1
            If cboETime.List(i) = Format(rsTmp.Fields("SB18"), "00:00") Then 'Format(Format(rsTmp.Fields("SB18"), "0000"), "##:##") Then
               cboETime.ListIndex = i
               Exit For
            End If
         Next i
      End If
      '2013/2/1 End
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp

       textSB01_2 = GetStaffName(textSB01, True)
   End If

   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset
strSql = ""
If txt1(0) <> "" Then
    strSql = strSql & " and sb01>='" & txt1(0) & "' "
End If
If txt1(1) <> "" Then
    strSql = strSql & " and sb01<='" & txt1(1) & "' "
End If
'Modify By Sindy 2019/10/1
If txt1(2) <> "" Then
    strSql = strSql & " and sb02>='" & DBDATE(txt1(2)) & "' "
End If
If txt1(3) <> "" Then
    strSql = strSql & " and sb02<='" & DBDATE(txt1(3)) & "' "
End If
'If txt1(2) <> "" And txt1(3) <> "" Then
'   strSql = strSql & " AND ('" & DBDATE(txt1(2)) & "' BETWEEN sb02 AND sb04 or '" & DBDATE(txt1(3)) & "' BETWEEN sb02 AND sb04) "
'End If
'2019/10/1 END
'抓取資料
strSql = "SELECT sb01,s1.st02,sqldateT(sb02)||' '||substr(ltrim(to_char('0000'||to_char(sb03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(sb03),'0000')),3,2),sqldateT(sb04)||' '||substr(ltrim(to_char('0000'||to_char(sb05),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(sb05),'0000')),3,2),decode(sb08,'1','長程','2','短程','3','大陸','4','國外',sb08),sb09,sb06,sb07,sb10 FROM staff_busi_trip,staff s1 where sb01=s1.st01(+) " & strSql & _
        " order by sb02,sb01,sb03 "
If rsTmp.State = 1 Then rsTmp.Close
rsTmp.CursorLocation = adUseClient
rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
Set GRD1.Recordset = rsTmp
SetGrd
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
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
         ' 新增
      Case 1, 2, 3, 4:
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

Private Function CheckDataValid() As Boolean
   Dim nResponse As Boolean
   Dim strTmp  As String
   CheckDataValid = False
   
   nResponse = False
   textsb01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSb02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSb03_1_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSb03_2_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSb04_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSb05_1_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSb05_2_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSB06_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSB07_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSB08_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSB09_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSB01.Locked = bEnable
   If bEnable Then textSB01.BackColor = &H8000000F Else textSB01.BackColor = &H80000005
   If m_EditMode <> "2" Then 'Modify By Sindy 2011/12/5
      textSB02.Locked = bEnable
      If bEnable Then textSB02.BackColor = &H8000000F Else textSB02.BackColor = &H80000005
      ' 2008/12/24 Add BY SINDY
      textSB03_1.Locked = bEnable
      textSB03_2.Locked = bEnable
      If bEnable Then textSB03_1.BackColor = &H8000000F Else textSB03_1.BackColor = &H80000005
      If bEnable Then textSB03_2.BackColor = &H8000000F Else textSB03_2.BackColor = &H80000005
      ' 2008/12/24 END
   End If
   'Add By Sindy 2011/9/22
   textSB10.Locked = bEnable
   If bEnable Then textSB10.BackColor = &H8000000F Else textSB10.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   textSB01.Locked = bEnable
   textSB02.Locked = bEnable
   If bEnable Then textSB01.BackColor = &H8000000F Else textSB01.BackColor = &H80000005
   If bEnable Then textSB02.BackColor = &H8000000F Else textSB02.BackColor = &H80000005
   textSB03_1.Locked = bEnable
   textSB03_2.Locked = bEnable
   ' 2008/12/24 Add BY SINDY
   If bEnable Then textSB03_1.BackColor = &H8000000F Else textSB03_1.BackColor = &H80000005
   If bEnable Then textSB03_2.BackColor = &H8000000F Else textSB03_2.BackColor = &H80000005
   ' 2008/12/24 END
   textSB04.Locked = bEnable
   textSB05_1.Locked = bEnable
   textSB05_2.Locked = bEnable
   textSB06.Locked = bEnable
   textSB07.Locked = bEnable
   textSB08.Locked = bEnable
   textSB09.Locked = bEnable
   'Add By Sindy 2011/9/22
   textSB10.Locked = bEnable
   If bEnable Then textSB10.BackColor = &H8000000F Else textSB10.BackColor = &H80000005
   cboSTime.Locked = bEnable
   cboETime.Locked = bEnable
'   txtNote.Locked = bEnable
   
   'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'   '有表單編號的資料時,日及時欄位鎖住
'   If textSB10 <> "" Then
'      textSB06.Enabled = False
'      textSB07.Enabled = False
'   Else
'      textSB06.Enabled = True
'      textSB07.Enabled = True
'   End If
End Sub

Private Sub ClearField()
   Dim nIndex As Integer
   textSB01 = Empty
   textSB01_2 = Empty
   textSB02 = Empty
   textSB03_1 = Empty
   textSB03_2 = Empty
   textSB04 = Empty
   textSB05_1 = Empty
   textSB05_2 = Empty
   textSB06 = Empty
   textSB07 = Empty
   textSB08 = Empty
   textSB09 = Empty
   
   'Add By Sindy 2011/9/22
   textSB10 = Empty
   Frame1.Visible = False
   m_B1019 = Empty: m_B1004 = Empty: m_B1005 = Empty: m_B1006 = Empty: m_B1007 = Empty
   m_B1009 = Empty: m_B1010 = Empty: m_B1014 = Empty: m_B1015 = Empty: m_B1017 = Empty
   m_B1028 = Empty: m_B1029 = Empty
   txtNote = Empty
   
   Label23 = Empty
   SetGrd
   For nIndex = 0 To tf_SB - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
   LblIsApart.Visible = False 'Add By Sindy 2021/12/27
   cmdABS.Visible = False 'Add By Sindy 2022/10/28
End Sub

Private Sub UpdateFieldNewData()
    Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "SB01", textSB01
      'Modify By Sindy 2011/12/5 起始日期及起始時間開放修改
'      SetFieldNewData "SB02", DBDATE(textSB02)
'      SetFieldNewData "SB03", textSB03_1 & Format("00" & textSB03_2, "00")
   End If
   SetFieldNewData "SB02", DBDATE(textSB02)
   SetFieldNewData "SB03", textSB03_1 & Format("00" & textSB03_2, "00")
   SetFieldNewData "SB04", DBDATE(textSB04)
   SetFieldNewData "SB05", textSB05_1 & Format("00" & textSB05_2, "00")
   SetFieldNewData "SB06", textSB06
   SetFieldNewData "SB07", textSB07
   SetFieldNewData "SB08", textSB08
   SetFieldNewData "SB09", ChgSQL(textSB09)
   'Add By Sindy 2011/9/22
   SetFieldNewData "SB10", textSB10
   'Add By Sindy 2013/2/1
   SetFieldNewData "SB17", IIf(Frame1.Visible = False, "", Format(cboSTime, "hhmm"))
   SetFieldNewData "SB18", IIf(Frame1.Visible = False, "", Format(cboETime, "hhmm"))
   '2013/2/1 End
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To tf_SB
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "SB" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 2, 3, 4, 5, 6, 7, 17, 18:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
SetGrd
End Sub

Private Sub textsb01_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB01
End If
End Sub

Private Sub textsb01_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2011/9/22
Private Sub textSB01_LostFocus()
   '若輸入的員工代號為可寄信者,必須輸入表單編號
   If Frame1.Visible = True Then If textSB10.Enabled = True Then textSB10.SetFocus
   '新增狀態將游標停在員工代號的欄位
   If m_EditMode = 1 And textSB01 = "" Then textSB01.SetFocus
End Sub

Private Sub textsb01_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

If textSB01.Text = "" Then
   textSB01_2 = "" ' 2008/12/18 ADD BY SINDY
   'Add By Sindy 2011/9/22 預設值
   Frame1.Visible = False
End If

If m_EditMode <> 0 And textSB01 <> "" Then
    textSB01_2 = GetStaffName(textSB01, True)
    ' 2008/12/18 ADD BY SINDY
    ' 檢查員工編號規則
    If ChkStaffID(textSB01) Then
       Call textsb01_GotFocus
       Cancel = True
       Exit Sub
    End If
    ' 2008/12/18 END
    If textSB01_2 = "" Then
        MsgBox "員工編號錯誤！查無此員工！", vbInformation
        Call textsb01_GotFocus ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    If m_KeyCode = vbKeyF2 Then '按新增時
      'Add By Sindy 2011/9/22 檢查此員工是否為"不寄信"
      If ChkStaffST14(textSB01, False) = False Then
        strTit = "詢問"
        strMsg = "是否要補電子表單？" & vbCrLf & vbCrLf & _
                 "（注意：要補簽核流程，請先輸入表單編號）"
        nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
        If nResponse = vbYes Then
           '電子表單請假
           Frame1.Visible = True
        Else
           '紙本請假
           Frame1.Visible = False
        End If
      Else
        '不寄信,紙本請假
        Frame1.Visible = False
      End If
    End If
End If

If m_EditMode = 1 And textSB01 <> "" Then
    If textSB02 <> "" And Val(textSB03_1) > 0 And Val(textSB03_2) > 0 Then
      If IsRecordExist(textSB01, DBDATE(textSB02), Trim(textSB03_1.Text & textSB03_2.Text)) = True And textSB01.Enabled = True And textSB01.Locked = False Then
          MsgBox "該員工當天已有資料，請修改！", vbInformation
          Call textsb01_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
End If
End Sub

Private Sub textSb02_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB02
    CloseIme
End If
End Sub

Private Sub textSb02_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSb02_Validate(Cancel As Boolean)
If textSB03_1 = "" Then textSB03_1 = "00"
If textSB03_2 = "" Then textSB03_2 = "00"

If m_EditMode <> 0 And textSB02 <> "" Then
    If CheckIsTaiwanDate(textSB02, False) = False Then
        Call textSb02_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
        Exit Sub
    End If
'    If ChkWorkDay(DBDATE(textSb02)) = False Then
'        Call textSb02_GotFocus   ' 2008/12/18 ADD BY SINDY
'        Cancel = True
'        MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
'        Exit Sub
'    End If
    If textSB02 <> "" And textSB04 <> "" Then
      If RunNick2(textSB02, textSB04) Then
          Call textSb02_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
'    'Add By Sindy 2011/9/22
'    If Frame1.Visible = True And textSB10 <> "" And m_B1019 <> "" Then '有表單編號
'      If Val(DBDATE(textSB02)) < Val(DBDATE(m_B1004)) Or Val(DBDATE(textSB02)) > Val(DBDATE(m_B1006)) Then
'         MsgBox "請假日期只能改少不能改多！", vbInformation, "輸入日期錯誤"
'         Call textSb02_GotFocus
'         Cancel = True
'         Exit Sub
'      End If
'    End If
   
   'Add By Sindy 2021/8/13
   If m_EditMode = 1 Then '新增
      SetB102829Combo cboSTime, 1, textSB02, textSB01
      SetB102829Combo cboETime, 2, textSB02, textSB01
   End If
   '2021/8/13 END
   
   'Add By Sindy 2011/9/22
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSB10 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   End If
End If
If m_EditMode = 1 And textSB02 <> "" Then
    If textSB01 <> "" And textSB02 <> "" _
         And Val(textSB03_1) > 0 And Val(textSB03_2) > 0 Then
      If IsRecordExist(textSB01, DBDATE(textSB02), Trim(textSB03_1.Text & textSB03_2.Text)) = True And textSB01.Enabled = True And textSB01.Locked = False Then
          MsgBox "該員工當天已有資料，請修改！", vbInformation
          Call textSb02_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   arrGridHeadText = Array("員工編號", "姓名", "起始日期時間", "結束日期時間", "差程", "地點", "天數", "時數", "職務代理人")
   arrGridHeadWidth = Array(800, 900, 1200, 1200, 500, 1300, 500, 500, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub textSb03_1_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB03_1
End If
End Sub

Private Sub textSb03_1_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSb03_1_Validate(Cancel As Boolean)
If textSB03_1 = "" Then textSB03_1 = "00"

If m_EditMode = 1 And textSB03_1 <> "" Then
    If CheckLengthIsOK(textSB03_1, textSB03_1.MaxLength) = False Then
        Call textSb03_1_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    ' 2008/12/19 Add BY SINDY
    If textSB03_1.Text > 24 Then
       Call textSb03_1_GotFocus
       MsgBox "不可超過24時!", vbExclamation + vbOKOnly
       Cancel = True
       Exit Sub
    End If
    If textSB01 <> "" And textSB02 <> "" _
         And Val(textSB03_1) > 0 And Val(textSB03_2) > 0 Then
      If IsRecordExist(textSB01, DBDATE(textSB02), Trim(textSB03_1.Text & textSB03_2.Text)) = True And textSB01.Enabled = True And textSB01.Locked = False Then
          MsgBox "該員工當天已有資料，請修改！", vbInformation
          Call textSb03_1_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
    ' 2008/12/19 END
   'Add By Sindy 2011/9/22
   'If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSB10 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   'End If
End If
CloseIme
End Sub

Private Sub textSb03_2_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB03_2
End If
End Sub

Private Sub textSb03_2_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSb03_2_Validate(Cancel As Boolean)
If textSB03_2 = "" Then textSB03_2 = "00"

If m_EditMode = 1 And textSB03_2 <> "" Then
    If CheckLengthIsOK(textSB03_2, textSB03_2.MaxLength) = False Then
        Call textSb03_2_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    ' 2008/12/18 ADD BY SINDY
    If textSB03_2.Text > 59 Then
       Call textSb03_2_GotFocus
       MsgBox "不可超過59分!", vbExclamation + vbOKOnly
       Cancel = True
       Exit Sub
    End If
    If textSB01 <> "" And textSB02 <> "" _
         And Val(textSB03_1) > 0 And Val(textSB03_2) > 0 Then
      If IsRecordExist(textSB01, DBDATE(textSB02), Trim(textSB03_1.Text & textSB03_2.Text)) = True And textSB01.Enabled = True And textSB01.Locked = False Then
          MsgBox "該員工當天已有資料，請修改！", vbInformation
          Call textSb03_2_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
    ' 2008/12/18 END
   'Add By Sindy 2011/9/22
   'If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSB10 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   'End If
End If
CloseIme
End Sub

Private Sub textSb04_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB04
End If
End Sub

Private Sub textSb04_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSb04_Validate(Cancel As Boolean)
If textSB05_1 = "" Then textSB05_1 = "00"
If textSB05_2 = "" Then textSB05_2 = "00"

If m_EditMode <> 0 And textSB04 <> "" Then
    If CheckIsTaiwanDate(textSB04, False) = False Then
        Call textSb04_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
        Exit Sub
    End If
'    If ChkWorkDay(DBDATE(textSb04)) = False Then
'        Call textSb04_GotFocus   ' 2008/12/18 ADD BY SINDY
'        Cancel = True
'        MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
'        Exit Sub
'    End If
    If textSB02 <> "" And textSB04 <> "" Then
      If RunNick2(textSB02, textSB04) Then
          Call textSb04_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
'    'Add By Sindy 2011/9/22
'    If Frame1.Visible = True And textSB10 <> "" And m_B1019 <> "" Then '有表單編號
'      If Val(DBDATE(textSB04)) < Val(DBDATE(m_B1004)) Or Val(DBDATE(textSB04)) > Val(DBDATE(m_B1006)) Then
'         MsgBox "請假日期只能改少不能改多！", vbInformation, "輸入日期錯誤"
'         Call textSb04_GotFocus
'         Cancel = True
'         Exit Sub
'      End If
'    End If
   'Add By Sindy 2011/9/22
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSB10 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   End If
End If
End Sub

Private Sub textSb05_1_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB05_1
End If
End Sub

Private Sub textSb05_1_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSb05_1_Validate(Cancel As Boolean)
If textSB05_1 = "" Then textSB05_1 = "00"

If m_EditMode <> 0 And textSB05_1 <> "" Then
   If CheckLengthIsOK(textSB05_1, textSB05_1.MaxLength) = False Then
       Call textSb05_1_GotFocus   ' 2008/12/18 ADD BY SINDY
       Cancel = True
       Exit Sub
   End If
   '2008/12/19 Add BY SINDY
   If textSB05_1.Text > 24 Then
      Call textSb05_1_GotFocus
      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   '2008/12/19 END
   '2022/7/21 Add BY SINDY
   If textSB01 <> "" And textSB02 <> "" _
         And Val(textSB05_1) > 0 And Val(textSB05_2) > 0 Then
      If IsRecordExist(textSB01, DBDATE(textSB02), Trim(textSB03_1.Text & textSB03_2.Text)) = True And textSB01.Enabled = True And textSB01.Locked = False Then
          MsgBox "該員工當天已有資料，請修改！", vbInformation
          Call textSb05_1_GotFocus
          Cancel = True
          Exit Sub
      End If
   End If
   '2022/7/21 END
   'Add By Sindy 2011/9/22
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSB10 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   End If
End If
CloseIme
End Sub

Private Sub textSb05_2_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB05_2
End If
End Sub

Private Sub textSb05_2_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSb05_2_LostFocus()
If m_EditMode <> 0 And textSB05_2 <> "" Then
    If Trim(textSB02) <> "" And Trim(textSB03_1) <> "" And Trim(textSB03_2) <> "" And Trim(textSB04) <> "" And Trim(textSB05_1) <> "" And Trim(textSB05_2) <> "" Then
        If CheckIsTaiwanDate(textSB02, False) = True And CheckIsTaiwanDate(textSB04, False) = True Then
            If CompDateTime(textSB02 & Format(textSB03_1, "00") & Format(textSB03_2, "00"), textSB04 & Format(textSB05_1, "00") & Format(textSB05_2, "00")) = False Then
                Call textSb05_2_GotFocus   ' 2008/12/18 ADD BY SINDY
                MsgBox "日期時間設定錯誤！", vbInformation, "輸入錯誤！"
                'textSb04.SetFocus
                Exit Sub
            End If
        End If
    End If
End If
End Sub

Private Sub textSb05_2_Validate(Cancel As Boolean)
If textSB05_2 = "" Then textSB05_2 = "00"

If m_EditMode <> 0 And textSB05_2 <> "" Then
   If CheckLengthIsOK(textSB05_2, textSB05_2.MaxLength) = False Then
       Call textSb05_2_GotFocus   ' 2008/12/18 ADD BY SINDY
       Cancel = True
       Exit Sub
   End If
   ' 2008/12/18 ADD BY SINDY
   If textSB05_2.Text > 59 Then
      Call textSb05_2_GotFocus
      MsgBox "不可超過59分!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   ' 2008/12/18 END
   '2022/7/21 Add BY SINDY
   If textSB01 <> "" And textSB02 <> "" _
         And Val(textSB05_1) > 0 And Val(textSB05_2) > 0 Then
      If IsRecordExist(textSB01, DBDATE(textSB02), Trim(textSB03_1.Text & textSB03_2.Text)) = True And textSB01.Enabled = True And textSB01.Locked = False Then
          MsgBox "該員工當天已有資料，請修改！", vbInformation
          Call textSb05_2_GotFocus
          Cancel = True
          Exit Sub
      End If
   End If
   '2022/7/21 END
   'Modify By Sindy 2011/9/22
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSB10 <> "" Then
         '計算時數
         Call CountDayHour
      Else
         '可以人工修改
         If textSB06 = "" Or textSB07 = "" Or (textSB06 = "0" And textSB07 = "0") Then
           '計算時數
           Call CountDayHour
         End If
      End If
   End If
End If
CloseIme
End Sub

Private Sub textSB06_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB06
End If
End Sub

Private Sub textSB06_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSB06_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textSB06 <> "" Then
    If CheckLengthIsOK(textSB06, textSB06.MaxLength) = False Then
        Call textSB06_GotFocus   ' 2008/12/24 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textSB07_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB07
End If
End Sub

Private Sub textSB07_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub textSB07_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textSB07 <> "" Then
    If CheckLengthIsOK(textSB07, textSB07.MaxLength) = False Then
        Call textSB07_GotFocus   ' 2008/12/24 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    
    ' 2008/12/17 ADD BY SINDY
    'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
    'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
    'Modify By Sindy 2012/7/9 上班時數為特殊者
    Call Pub_GetSpecWorkHour(textSB01, textSB02)
'    If textSB01 = "99029" Then
'         If textSB07.Text >= 5 Then
'            Call textSB07_GotFocus
'            MsgBox "出差時數-共(時)不可超過5小時!!!", vbExclamation + vbOKOnly
'            Cancel = True
'            Exit Sub
'         End If
'    '2010/7/14 End
    'Modify By Sindy 2018/5/14 Mark:不控管
'    If Val(textSB07.Text) >= Val(PUB_intWkHour) Then
'       Call textSB07_GotFocus
'       MsgBox "出差時數-共(時)不可超過" & PUB_intWkHour & "小時!!!", vbExclamation + vbOKOnly
'       Cancel = True
'       Exit Sub
'    End If
'    ' 2008/12/17 END
End If
CloseIme
End Sub

Private Sub textSB08_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB08
End If
End Sub

Private Sub textSB08_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSB08_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textSB08 <> "" Then
    If CheckLengthIsOK(textSB08, textSB08.MaxLength) = False Then
        Call textSB08_GotFocus   ' 2008/12/24 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    ' 2008/12/24 ADD BY SINDY
    If Trim(textSB08) <> "" Then
      If textSB08 <> "1" And textSB08 <> "2" And textSB08 <> "3" And textSB08 <> "4" Then
         MsgBox "差程代碼有誤!!!", vbExclamation + vbOKOnly
         Call textSB08_GotFocus
         Cancel = True
         Exit Sub
      End If
    End If
    ' 2008/12/24 END
End If
CloseIme
End Sub

Private Sub textSB09_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB09
    OpenIme
End If
End Sub

Private Sub textSB09_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textSB09 <> "" Then
    If CheckLengthIsOK(textSB09, textSB09.MaxLength) = False Then
        Call textSB09_GotFocus   ' 2008/12/24 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = UpperCase(KeyAscii)
      Case 2, 3
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         ' 2008/12/17 ADD BY SINDY
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ' 2008/12/17 END
         ElseIf Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         
      Case 2, 3
         ' 2008/12/16 MODIFY BY SINDY
         'If CheckIsTaiwanDate(txt1(Index), False) = False Then
         If CheckIsTaiwanDate(txt1(Index), False) = False And Trim(txt1(Index)) <> "" Then
         ' 2008/12/16 END
            Call txt1_GotFocus(Index)
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
         
         ' 2008/12/17 ADD BY SINDY
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ' 2008/12/17 END
         ElseIf Index = 3 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         
      Case Else
   End Select
End Sub

'計算時數
Private Function CountDayHour()
'Dim tmpCalH As String
''Dim dblSTime As Double, dblETime As Double
Dim temp As Variant
'Dim bwk5hour As Boolean
Dim strSTime As String, strETime As String 'Add By Sindy 2011/9/22
Dim m_Day As Integer, m_Hour As Double
   
   'Add By Sindy 2010/7/14 99029伊恩一天只上4個小時
   'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
'   bwk5hour = False
'   If textSB01 = "99029" Then bwk5hour = True
   '2010/7/14 End

   'Modify By Sindy 2011/9/22 電子表單資料,日及時欄位值一律由系統計算
   If textSB06 = "" Or textSB07 = "" Or (textSB06 = "0" And textSB07 = "0") Then 'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
   'If textSB06 = "" Or textSB07 = "" Or (textSB06 = "0" And textSB07 = "0") Or (Frame1.Visible = True And textSB10 <> "") Then
      If Trim(textSB02) <> "" And Trim(textSB03_1) <> "" And Trim(textSB03_2) <> "" And Trim(textSB04) <> "" And Trim(textSB05_1) <> "" And Trim(textSB05_2) <> "" Then
          If CheckIsTaiwanDate(textSB02, False) = True And CheckIsTaiwanDate(textSB04, False) = True Then
              'Modify By Sindy 2010/7/14 增加傳入bwk4hour
              'Modify By Sindy 2011/3/8 增加傳入bwk5hour
              'Add By Sindy 2011/9/22
              strSTime = "": strETime = ""
              If cboSTime.Visible = True Then strSTime = Format(cboSTime.Text, "hhmm")
              If cboETime.Visible = True Then strETime = Format(cboETime.Text, "hhmm")

'              'Add By Sindy 2011/11/9 調整的 (劉經理：國外及大陸出差天數計算應含休假日,國內出差則以實際工作時數計算)
'              If textSB08 = "3" Or textSB08 = "4" Then
'                  tmpCalH = CalDateTime(textSB02 & Format(textSB03_1, "00") & Format(textSB03_2, "00"), textSB04 & Format(textSB05_1, "00") & Format(textSB05_2, "00"), bwk5hour, strSTime, strETime, False)
'              Else
'                  tmpCalH = CalDateTime(textSB02 & Format(textSB03_1, "00") & Format(textSB03_2, "00"), textSB04 & Format(textSB05_1, "00") & Format(textSB05_2, "00"), bwk5hour, strSTime, strETime)
'              End If
               Call PUB_CountHour_Busi_Trip(textSB02, Format(textSB03_1, "00") & Format(textSB03_2, "00"), textSB04, Format(textSB05_1, "00") & Format(textSB05_2, "00"), m_Day, m_Hour)
               If m_Day > 0 Then
                  textSB06 = m_Day
               Else
                  textSB06 = "0"
               End If
               If m_Hour > 0 Then
                  textSB07 = m_Hour
               Else
                  textSB07 = "0"
               End If

'              'Add By Sindy 98/03/13 起始時間<=12時並且迄止時間>=13時30分者，減1小時
'              dblSTime = Val(textSB03_1 & textSB03_2)
'              dblETime = Val(textSB05_1 & textSB05_2)
'              If dblSTime <= 1200 And dblETime >= 1330 Then
'                  tmpCalH = tmpCalH - 1
'              End If
'              '98/03/13 End
'
'              If tmpCalH > "" Then
'                  'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
'                  'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
'                  If textSB01 = "99029" Then
'                      If tmpCalH < 5 Then
'                          textSB06 = 0
'                      Else
'                          'textSB06 = Val(tmpCalH) \ 5
'                          temp = Split(CStr(Val(tmpCalH) / 5), ".")
'                          textSB06 = temp(0)
'                      End If
'                      textSB07 = Val(tmpCalH) - (Val(textSB06) * 5)
'                  '2010/7/14 End
'                  Else
'                      If tmpCalH < 8 Then
'                          textSB06 = 0
'                      Else
'                          'textSB06 = Val(tmpCalH) \ 8
'                          temp = Split(CStr(Val(tmpCalH) / 8), ".")
'                          textSB06 = temp(0)
'                      End If
'                      textSB07 = Val(tmpCalH) - (Val(textSB06) * 8)
'                  End If
'              Else
'                  textSB06 = ""
'                  textSB07 = ""
'                  MsgBox "日期時間設定錯誤！", vbInformation, "輸入錯誤！"
'                  Exit Function
'              End If
          Else
              textSB06 = ""
              textSB07 = ""
          End If
      Else
          textSB06 = ""
          textSB07 = ""
      End If
   End If
End Function

Private Sub textSB10_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSB10
    CloseIme
End If
End Sub

Private Sub textSB10_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSB10_LostFocus()
   '新增狀態時可以輸入表單編號做查詢
   If m_EditMode = 1 And textSB10 <> "" Then
      If GetABS010 = True Then
         textSB10.Enabled = False
      End If
   End If
End Sub

Private Sub textSB10_Validate(Cancel As Boolean)
   If Frame1.Visible = False Then Exit Sub
   
   If m_EditMode = 1 And textSB10 <> "" Then
      If CheckLengthIsOK(textSB10, textSB10.MaxLength) = False Then
         Call textSB10_GotFocus
         Cancel = True
         Exit Sub
      End If
      If ChkAbsSysB1001Exist(textSB10, "03", textSB01) = False Then
         Call textSB10_GotFocus
         Cancel = True
         Exit Sub
      End If
      If ChkPerSysB1001Exist(textSB10, textSB01, False) = True Then
         MsgBox "表單編號重覆！", vbExclamation
         Call textSB10_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub cboSTime_GotFocus()
'   InverseTextBox cboSTime
End Sub

Private Sub cboSTime_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub cboSTime_Validate(Cancel As Boolean)
If Frame1.Visible = True And cboSTime <> "" Then
   'Modify By Sindy 2015/5/13 Mark
'   If Val(Format(cboSTime.Text, "hhmm")) > Val(Right("00" & textSB03_1, 2) & Right("00" & textSB03_2, 2)) Then
'      Call cboSTime_GotFocus
'      MsgBox "起日上班時段必須小於或等於起日請假時間!", vbExclamation + vbOKOnly
'      Cancel = True
'      Exit Sub
'   End If
   '2015/5/13 END
   
'   If Val(Format(cboSTime.Text, "hhmm")) > 2400 Then
'      Call cboSTime_GotFocus
'      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
'      Cancel = True
'      Exit Sub
'   End If
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSB10 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   End If
Else
'   If textSA02 <> textSA04 Then '跨日
'      If cboSTime = "" Then
'         Call cboSTime_GotFocus
'         MsgBox "請輸入起日上班時段!", vbExclamation + vbOKOnly
'         Cancel = True
'         Exit Sub
'      End If
'   End If
End If
CloseIme
End Sub

Private Sub cboETime_GotFocus()
'   InverseTextBox cboETime
End Sub

Private Sub cboETime_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub cboETime_Validate(Cancel As Boolean)
If Frame1.Visible = True And cboETime <> "" Then
   'Modify By Sindy 2015/5/13 Mark
'   If Val(Format(cboETime.Text, "hhmm")) < Val(Right("00" & textSB05_1, 2) & Right("00" & textSB05_2, 2)) Then
'      Call cboETime_GotFocus
'      MsgBox "迄日下班時段必須大於或等於迄日請假時間!", vbExclamation + vbOKOnly
'      Cancel = True
'      Exit Sub
'   End If
   '2015/5/13 END
   
'   If Val(Format(cboETime.Text, "hhmm")) > 2400 Then
'      Call cboETime_GotFocus
'      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
'      Cancel = True
'      Exit Sub
'   End If
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSB10 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   End If
Else
'   If textSA02 <> textSA04 Then '跨日
'      If cboETime = "" Then
'         Call cboETime_GotFocus
'         MsgBox "請輸入迄日下班時段!", vbExclamation + vbOKOnly
'         Cancel = True
'         Exit Sub
'      End If
'   End If
End If
CloseIme
End Sub

'Add By Sindy 2011/9/22
Private Function GetABS010(Optional bolOnlyQrySETime As Boolean = False) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, i As Integer
   
   Screen.MousePointer = vbHourglass
   GetABS010 = False
   cmdABS.Visible = False 'Add By Sindy 2022/10/28
   
   '出缺勤電子簽核主檔
   strSql = "Select B1001,B1002,B1003,B1004,substr(ltrim(to_char('0000'||to_char(B1005),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),3,2) B1005,B1006,substr(ltrim(to_char('0000'||to_char(B1007),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),3,2) B1007,B1008||' '||AC03 B1008,B1009,B1010,B1011,B1012,B1013,B1014,B1015,B1016,B1017," & B1018CName & " B1018,B1019,B1020,B1021,B1022,B1023,B1024,B1025,B1026,B1027,substr(ltrim(to_char('0000'||to_char(B1028),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1028),'0000')),3,2) B1028,substr(ltrim(to_char('0000'||to_char(B1029),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1029),'0000')),3,2) B1029 " & _
            "From ABS010, allcode " & _
            "Where ac01(+)='04' and B1008=ac02(+) " & _
            "and B1001='" & textSB10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GetABS010 = True
      'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'      '有表單編號的資料時,日及時欄位鎖住
'      textSB06.Enabled = False
'      textSB07.Enabled = False
      
      '記錄原始資料 : 註.m_變數值必須在ClearField函數裡清值
      If Not IsNull(rsTmp.Fields("B1019")) Then m_B1019 = rsTmp.Fields("B1019")
      If Not IsNull(rsTmp.Fields("B1004")) Then m_B1004 = rsTmp.Fields("B1004")
      If Not IsNull(rsTmp.Fields("B1005")) Then m_B1005 = IIf(Format(rsTmp.Fields("B1005"), "hhmm") = "0000", "", Format(rsTmp.Fields("B1005"), "hhmm"))
      If Not IsNull(rsTmp.Fields("B1006")) Then m_B1006 = rsTmp.Fields("B1006")
      If Not IsNull(rsTmp.Fields("B1007")) Then m_B1007 = IIf(Format(rsTmp.Fields("B1007"), "hhmm") = "0000", "", Format(rsTmp.Fields("B1007"), "hhmm"))
      If Not IsNull(rsTmp.Fields("B1009")) Then m_B1009 = rsTmp.Fields("B1009")
      If Not IsNull(rsTmp.Fields("B1010")) Then m_B1010 = rsTmp.Fields("B1010")
      If Not IsNull(rsTmp.Fields("B1014")) Then m_B1014 = rsTmp.Fields("B1014")
      If Not IsNull(rsTmp.Fields("B1015")) Then m_B1015 = rsTmp.Fields("B1015")
      If Not IsNull(rsTmp.Fields("B1017")) Then m_B1017 = rsTmp.Fields("B1017")
      If Not IsNull(rsTmp.Fields("B1028")) Then m_B1028 = IIf(Format(rsTmp.Fields("B1028"), "hhmm") = "0000", "", Format(rsTmp.Fields("B1028"), "hhmm"))
      If Not IsNull(rsTmp.Fields("B1029")) Then m_B1029 = IIf(Format(rsTmp.Fields("B1029"), "hhmm") = "0000", "", Format(rsTmp.Fields("B1029"), "hhmm"))
      
      '顯示其他資料至畫面上
      'Add By Sindy 2022/10/28 + if 已簽核不要顯示於畫面上,已人事資料為主
      If Not IsNull(rsTmp.Fields("B1019")) Then
         '為防止簽核後又修改,抓人事資料
         strSql = "select * from abs012 where b1201='" & textSB10 & "' and substr(b1207,1,4)='修改資料'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            cmdABS.Visible = True
            '記錄畫面上的資料 : 註.m_變數值必須在ClearField函數裡清值
            m_B1004 = DBDATE(textSB02)
            m_B1005 = Format(textSB03_1 & textSB03_2, "0000")
            m_B1006 = DBDATE(textSB04)
            m_B1007 = Format(textSB05_1 & textSB05_2, "0000")
            m_B1009 = textSB06
            m_B1010 = textSB07
            m_B1014 = textSB08
            m_B1015 = textSB09
         End If
      Else
      '2022/10/28 END
         If Not IsNull(rsTmp.Fields("B1004")) Then textSB02 = ChangeWStringToTString(rsTmp.Fields("B1004"))
         If Not IsNull(rsTmp.Fields("B1005")) Then textSB03_1 = Left(rsTmp.Fields("B1005"), 2): textSB03_2 = Right(rsTmp.Fields("B1005"), 2)
         If Not IsNull(rsTmp.Fields("B1006")) Then textSB04 = ChangeWStringToTString(rsTmp.Fields("B1006"))
         If Not IsNull(rsTmp.Fields("B1007")) Then textSB05_1 = Left(rsTmp.Fields("B1007"), 2): textSB05_2 = Right(rsTmp.Fields("B1007"), 2)
         If Not IsNull(rsTmp.Fields("B1009")) Then textSB06 = rsTmp.Fields("B1009")
         If Not IsNull(rsTmp.Fields("B1010")) Then textSB07 = rsTmp.Fields("B1010")
         If Not IsNull(rsTmp.Fields("B1014")) Then textSB08 = rsTmp.Fields("B1014")
         If Not IsNull(rsTmp.Fields("B1015")) Then textSB09 = rsTmp.Fields("B1015")
      End If
      
      'Add By Sindy 2021/8/11
      If bolOnlyQrySETime = False Then
         SetB102829Combo cboSTime, 1, textSB02, textSB01
         SetB102829Combo cboETime, 2, textSB02, textSB01
      End If
      '2021/8/11 END
      '僅顯示起日上班時段,迄日下班時段至畫面上
      If bolOnlyQrySETime = True Then
         'If Not IsNull(rsTmp.Fields("B1028")) Then
         If m_B1028 <> "" Then
            For i = 0 To cboSTime.ListCount - 1
               If cboSTime.List(i) = Format(Format(rsTmp.Fields("B1028"), "hhmm"), "00:00") Then 'IIf(Left(m_B1028, 1) = "0", "0", "") & Format(Format(rsTmp.Fields("B1028"), "hhmm"), "##:##") Then
                  cboSTime.ListIndex = i
                  Exit For
               End If
            Next i
         End If
         'If Not IsNull(rsTmp.Fields("B1029")) Then
         If m_B1029 <> "" Then
            For i = 0 To cboETime.ListCount - 1
               If cboETime.List(i) = Format(Format(rsTmp.Fields("B1029"), "hhmm"), "00:00") Then 'IIf(Left(m_B1029, 1) = "0", "0", "") & Format(Format(rsTmp.Fields("B1029"), "hhmm"), "##:##") Then
                  cboETime.ListIndex = i
                  Exit For
               End If
            Next i
         End If
         GoTo EXITSUB
      End If
      
   Else
'      Screen.MousePointer = vbDefault
'      ShowNoData
'      rsTmp.Close
'      Set rsTmp = Nothing
'      Exit Sub
   End If
   
EXITSUB:
   rsTmp.Close
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
End Function

'Add By Sindy 2011/9/22
'Modify By Sindy 2013/2/1 +bolIsDel
Private Sub ProABSData(Optional bolIsDel As Boolean = False)
Dim strUpdDate As String, strUpdTime As String
Dim strB1004 As String, strB1005 As String, strB1006 As String
Dim strB1007 As String, strB1009 As String, strB1010 As String
Dim strB1014 As String, strB1015 As String
Dim strB1028 As String, strB1029 As String
Dim strOldData As String, strNowData As String, strNote As String
'Dim strTo As String 'Add By Sindy 2012/7/17
Dim strSubject As String, strContent As String 'Add By Sindy 2019/5/24
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   '檢查有無異動資料:
   '畫面上的欄位值
   strB1004 = DBDATE(textSB02)
   strB1005 = textSB03_1 & Format("00" & textSB03_2, "00")
   strB1006 = DBDATE(textSB04)
   strB1007 = textSB05_1 & Format("00" & textSB05_2, "00")
   strB1009 = textSB06
   strB1010 = textSB07
   strB1014 = textSB08
   strB1015 = textSB09
   'Add By Sindy 2013/4/1
   If textSB07 = 0 Or textSB07 = "" Then
      cboSTime.Clear
      cboETime.Clear
   End If
   '2013/4/1 End
   strB1028 = IIf(Frame1.Visible = False, "", Format(cboSTime, "hhmm"))
   strB1029 = IIf(Frame1.Visible = False, "", Format(cboETime, "hhmm"))
   
   '串原始資料
   If m_B1028 <> "" And m_B1029 <> "" Then
      strOldData = strOldData & "非整日," & Format(m_B1028, "##:##") & "," & Format(m_B1029, "##:##")
   End If
   strOldData = strOldData & "," & ChangeWStringToTDateString(m_B1004) & "," & Format(m_B1005, "##:##")
   strOldData = strOldData & "," & ChangeWStringToTDateString(m_B1006) & "," & Format(m_B1007, "##:##")
   strOldData = strOldData & "," & m_B1009 & "日," & m_B1010 & "時"
   strOldData = strOldData & ",差程" & m_B1014 & "," & m_B1015
   '串目前畫面上資料
   If Frame1.Visible = True And _
      cboSTime.Text <> "" And cboETime.Text <> "" Then
      strNowData = strNowData & "非整日," & Format(strB1028, "##:##") & "," & Format(strB1029, "##:##")
   End If
   strNowData = strNowData & "," & ChangeWStringToTDateString(strB1004) & "," & Format(strB1005, "##:##")
   strNowData = strNowData & "," & ChangeWStringToTDateString(strB1006) & "," & Format(strB1007, "##:##")
   strNowData = strNowData & "," & strB1009 & "日," & strB1010 & "時"
   strNowData = strNowData & ",差程" & strB1014 & "," & strB1015
   If Left(strOldData, 1) = "," Then strOldData = Right(strOldData, Len(strOldData) - 1)
   If Left(strNowData, 1) = "," Then strNowData = Right(strNowData, Len(strNowData) - 1)
   
   '流程備註檔
   If txtNote.Text <> "" And textSB10 <> "" Then
      strSql = GetInsertABS012Sql(Trim(textSB10), 人事處, strUpdDate, strUpdTime, "", txtNote)
      cnnConnection.Execute strSql
   End If
   
   If strOldData <> strNowData And textSB10 <> "" Then '電子簽核的,非紙本
      '人事處尚未簽收時,在人事系統已先建立此表單編號資料,須一併更新出缺勤電子簽核主檔資料
      If m_B1019 = "" Then
         strSql = "update ABS010 set " & _
                  "B1004= " & CNULL(DBDATE(strB1004)) & _
                  ",B1005= " & CNULL(strB1005) & _
                  ",B1006= " & CNULL(strB1006) & _
                  ",B1007= " & CNULL(strB1007) & _
                  ",B1009= " & CNULL(strB1009) & _
                  ",B1010= " & CNULL(strB1010) & _
                  ",B1014= " & CNULL(strB1014) & _
                  ",B1015= " & CNULL(strB1015) & _
                  ",B1028= " & CNULL(strB1028) & _
                  ",B1029= " & CNULL(strB1029) & _
                  " where B1001=" & CNULL(textSB10)
         cnnConnection.Execute strSql
      End If
      '檢查有異動資料時,須記錄異動資訊到表單流程備註
      strNote = "修改資料" & strOldData & "->" & strNowData
      strSql = GetInsertABS012Sql(Trim(textSB10), "M21", strUpdDate, strUpdTime, "", strNote)
      cnnConnection.Execute strSql
   End If
   
   If m_B1019 = "" And m_EditMode = 1 And textSB10 <> "" Then
      '寄E-Mail通知當事人
      PUB_SendMail strUserNum, textSB01, "", "表單人事處已先行作業，請儘速簽核。", _
      "表單內容為，" & strNowData & vbCrLf & _
      "(表單編號：" & textSB10 & ")", , , , , , , , , , True
   ElseIf m_B1019 <> "" And m_EditMode = 1 And textSB10 <> "" Then
      strSql = "update ABS010 set " & _
               "B1018='" & 已核准 & "'" & _
               " where B1001=" & CNULL(textSB10)
      cnnConnection.Execute strSql
      
      '記錄資訊到表單流程備註
      strNote = "補入資料"
      strSql = GetInsertABS012Sql(Trim(textSB10), "M21", strUpdDate, strUpdTime, "", strNote)
      cnnConnection.Execute strSql
   Else
      If strOldData <> strNowData Then
'         '寄E-Mail通知當事人有異動內容
'         'Modify By Sindy 2012/7/17 發E-Mail通知當事人之外，已簽核的職代及審核主管亦也要通知
'         strTo = GetBossB1107_All(textSB10)
'         'Add By Sindy 2012/7/17 專利處P10-P14,必須另外E-Mail通知71011王副總
'         If (GetStaffDepartment(textSB01) >= "P10" And GetStaffDepartment(textSB01) <= "P14") And _
'            InStr(strTo, "71011") = 0 Then
'            strTo = strTo + ";71011"
'         End If
         
         If textSB10 <> "" Then  '電子簽核的,非紙本
            strSubject = "[通知]人事處有修改資料(表單編號：" & textSB10 & ")"
         Else
            strSubject = "[通知]人事處有修改資料"
         End If
         'Add By Sindy 2013/2/1
         If bolIsDel = True Then
'            PUB_SendMail strUserNum, textSB01, "", "[通知]人事處有修改資料(表單編號：" & textSB10 & ")", _
'            "異動前資料：" & strOldData & vbCrLf & _
'            "註銷的資料：" & strNowData & vbCrLf & _
'            "(表單編號：" & textSB10 & ")" & vbCrLf & vbCrLf & _
'            "人事處修改原因：" & txtNote, , , , , , strTo, , , , True
            strContent = "異動前資料：" & strOldData & vbCrLf & _
                         "註銷的資料：" & strNowData & vbCrLf & _
                         IIf(textSB10 <> "", "(表單編號：" & textSB10 & ")" & vbCrLf & vbCrLf, "") & _
                         "人事處修改原因：" & txtNote
         Else
         '2013/2/1 End
'            PUB_SendMail strUserNum, textSB01, "", "[通知]人事處有修改資料(表單編號：" & textSB10 & ")", _
'            "異動前資料：" & strOldData & vbCrLf & _
'            "異動後資料：" & strNowData & vbCrLf & _
'            "(表單編號：" & textSB10 & ")" & vbCrLf & vbCrLf & _
'            "人事處修改原因：" & txtNote, , , , , , strTo, , , , True
            strContent = "異動前資料：" & strOldData & vbCrLf & _
                         "異動後資料：" & strNowData & vbCrLf & _
                         IIf(textSB10 <> "", "(表單編號：" & textSB10 & ")" & vbCrLf & vbCrLf, "") & _
                         "人事處修改原因：" & txtNote
         End If
         '2012/7/17 End
         'Add By Sindy 2019/5/24 假單完成,後續資料檢查及SendMail
         Call PUB_AutoM21Receive_SendMail(IIf(textSB10 <> "", textSB10, ""), 表單類別_出差, textSB01, DBDATE(textSB02), Trim(Format("00" & textSB03_1, "00") & Format("00" & textSB03_2, "00")), _
            DBDATE(textSB04), textSB08, Left(DBDATE(textSB02), 6), , strSubject, strContent, m_EditMode)
      End If
   End If
   
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Sub

Private Sub txtNote_GotFocus()
   InverseTextBox txtNote
   OpenIme
End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
If txtNote <> "" Then
   If CheckLengthIsOK(txtNote, txtNote.MaxLength) = False Then
      Call txtNote_GotFocus
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
End Sub

Private Sub DelMark()
'Dim strTo As String
Dim nResponse
Dim m_B1018 As String, strUpdDate As String, strUpdTime As String
Dim strSubject As String, strContent As String
   
On Error GoTo ErrHand
   
'   If txtNote.Text = "" Then
'      MsgBox "原因不可以空白！", vbExclamation
'      txtNote.SetFocus
'      Exit Sub
'   End If
'
'   nResponse = MsgBox("註銷會將人事系統的相關資料一併刪除，確定要註銷嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問")
'   If nResponse = vbNo Then Exit Sub
   
   If textSB10 <> "" Then m_B1018 = 註銷 '(06)
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   'Modify By Sindy 2019/5/24
   'strContent = GetEMailContent(textSB10, strSubject)
   strContent = GetEMailContent(IIf(textSB10 <> "", textSB10, ""), strSubject, m_B1018, , , "03", textSB01, DBDATE(textSB02), Val(Trim(Format("00" & textSB03_1, "00") & Format("00" & textSB03_2, "00"))), m_EditMode)
   strContent = strContent & vbCrLf & vbCrLf & _
                "人事處修改原因：" & txtNote
   
'   cnnConnection.BeginTrans
   
   If textSB10 <> "" Then
      '流程備註檔
      If txtNote.Text <> "" Then
         strSql = GetInsertABS012Sql(Trim(textSB10), 人事處, strUpdDate, strUpdTime, m_B1018, txtNote)
         cnnConnection.Execute strSql
      End If
      '主檔
      strSql = "update ABS010 set " & _
               "B1018='" & m_B1018 & "'" & _
               " where B1001='" & textSB10 & "' "
      cnnConnection.Execute strSql
   End If
'   '刪除人事系統該筆表單資料, 並寫Log
'   If Left(CboB1002, 2) = 表單類別_請假 Then
'      strSql = "delete from Staff_Absence where SA09='" & Trim(txtB1001) & "'"
'   ElseIf Left(CboB1002, 2) = 表單類別_加班 Then
'      strSql = "delete from Staff_Overtime where So13='" & Trim(txtB1001) & "'"
'   ElseIf Left(CboB1002, 2) = 表單類別_出差 Then
'      strSql = "delete from Staff_Busi_Trip where SB10='" & Trim(txtB1001) & "'"
'   End If
'   Pub_SeekTbLog strSql '記錄刪除Log
'   cnnConnection.Execute strSql
'
'   cnnConnection.CommitTrans
   
'   '發E-Mail通知當事人及已簽核的職代及審核主管
'   strTo = GetBossB1107_All(textSB10)
'   'Add By Sindy 2012/8/23 專利處P10-P14,必須另外E-Mail通知71011王副總
'   If (GetStaffDepartment(textSB01) >= "P10" And GetStaffDepartment(textSB01) <= "P14") And _
'      InStr(strTo, "71011") = 0 Then
'      strTo = strTo + ";71011"
'   End If
'   '2012/8/23 End
'   strContent = GetEMailContent(textSB10, strSubject)
'   'PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True
'   PUB_SendMail strUserNum, textSB01, "", strSubject, strContent & vbCrLf & vbCrLf & _
'         "人事處修改原因：" & txtNote, , , , , , strTo, , , , True
   
   'Add By Sindy 2019/5/24 假單完成,後續資料檢查及SendMail
   Call PUB_AutoM21Receive_SendMail(IIf(textSB10 <> "", textSB10, ""), 表單類別_出差, textSB01, DBDATE(textSB02), Trim(Format("00" & textSB03_1, "00") & Format("00" & textSB03_2, "00")), _
      DBDATE(textSB04), textSB08, Left(DBDATE(textSB02), 6), , strSubject, strContent, m_EditMode)
      
'   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "註銷失敗！" & vbCrLf & Err.Description
End Sub
