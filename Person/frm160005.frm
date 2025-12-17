VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160005 
   BorderStyle     =   1  '單線固定
   Caption         =   "請假資料"
   ClientHeight    =   5060
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5060
   ScaleWidth      =   8190
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
            Picture         =   "frm160005.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160005.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160005.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160005.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160005.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160005.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160005.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160005.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160005.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160005.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160005.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8190
      _ExtentX        =   14446
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
      Height          =   4380
      Left            =   30
      TabIndex        =   20
      Top             =   660
      Width           =   8120
      _ExtentX        =   14323
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm160005.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LblNote"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "textSA01_2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label23"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtNote"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LblIsApart"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textSA01"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textSA06"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "FrameSA19"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdABS"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160005.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(0)"
      Tab(1).Control(1)=   "txt1(1)"
      Tab(1).Control(2)=   "txt1(2)"
      Tab(1).Control(3)=   "txt1(3)"
      Tab(1).Control(4)=   "cmdok"
      Tab(1).Control(5)=   "GRD1"
      Tab(1).Control(6)=   "Line5"
      Tab(1).Control(7)=   "Line4"
      Tab(1).Control(8)=   "Label15"
      Tab(1).Control(9)=   "Label16"
      Tab(1).ControlCount=   10
      Begin VB.Frame Frame2 
         Height          =   1480
         Left            =   540
         TabIndex        =   39
         Top             =   960
         Width           =   4570
         Begin VB.TextBox textSA08 
            Height          =   315
            Left            =   1470
            MaxLength       =   4
            TabIndex        =   9
            Top             =   1110
            Width           =   525
         End
         Begin VB.TextBox textSA04 
            Height          =   270
            Left            =   3180
            MaxLength       =   7
            TabIndex        =   5
            Top             =   390
            Width           =   945
         End
         Begin VB.TextBox textSA05_2 
            Height          =   285
            Left            =   3540
            MaxLength       =   2
            TabIndex        =   7
            Top             =   690
            Width           =   585
         End
         Begin VB.TextBox textSA05_1 
            Height          =   285
            Left            =   2670
            MaxLength       =   2
            TabIndex        =   6
            Top             =   690
            Width           =   585
         End
         Begin VB.TextBox textSA03_1 
            Height          =   285
            Left            =   420
            MaxLength       =   2
            TabIndex        =   3
            Top             =   690
            Width           =   585
         End
         Begin VB.TextBox textSA02 
            Height          =   270
            Left            =   840
            MaxLength       =   7
            TabIndex        =   2
            Top             =   390
            Width           =   945
         End
         Begin VB.TextBox textSA03_2 
            Height          =   285
            Left            =   1260
            MaxLength       =   2
            TabIndex        =   4
            Top             =   690
            Width           =   585
         End
         Begin VB.TextBox textSA07 
            Height          =   315
            Left            =   660
            MaxLength       =   2
            TabIndex        =   8
            Top             =   1110
            Width           =   525
         End
         Begin VB.Line Line3 
            BorderWidth     =   3
            X1              =   420
            X2              =   4410
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "日期"
            Height          =   180
            Index           =   2
            Left            =   2760
            TabIndex        =   48
            Top             =   450
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "分"
            Height          =   180
            Left            =   4170
            TabIndex        =   47
            Top             =   750
            Width           =   180
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "時"
            Height          =   180
            Left            =   3300
            TabIndex        =   46
            Top             =   750
            Width           =   180
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "時"
            Height          =   180
            Left            =   1050
            TabIndex        =   45
            Top             =   750
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "日期"
            Height          =   180
            Index           =   17
            Left            =   420
            TabIndex        =   44
            Top             =   450
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "時間                起                                          迄"
            Height          =   180
            Left            =   90
            TabIndex        =   43
            Top             =   120
            Width           =   3620
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "分"
            Height          =   180
            Left            =   1890
            TabIndex        =   42
            Top             =   750
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
            Left            =   2250
            TabIndex        =   41
            Top             =   720
            Width           =   260
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "共             日               時"
            Height          =   180
            Left            =   420
            TabIndex        =   40
            Top             =   1170
            Width           =   1940
         End
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   450
            X2              =   2130
            Y1              =   330
            Y2              =   330
         End
         Begin VB.Line Line2 
            BorderWidth     =   3
            X1              =   2700
            X2              =   4380
            Y1              =   330
            Y2              =   330
         End
      End
      Begin VB.CommandButton cmdABS 
         BackColor       =   &H00C0FFFF&
         Caption         =   "簽核資料"
         Height          =   315
         Left            =   5130
         Style           =   1  '圖片外觀
         TabIndex        =   36
         Top             =   2160
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Frame FrameSA19 
         Caption         =   "請假事由"
         Height          =   1130
         Left            =   2340
         TabIndex        =   37
         Top             =   2850
         Width           =   2360
         Begin VB.TextBox textSA19_2 
            Appearance      =   0  '平面
            BackColor       =   &H8000000F&
            BorderStyle     =   0  '沒有框線
            Height          =   670
            Left            =   210
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  '垂直捲軸
            TabIndex        =   38
            Top             =   330
            Width           =   2020
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   1130
         Left            =   120
         TabIndex        =   26
         Top             =   2460
         Visible         =   0   'False
         Width           =   2360
         Begin VB.TextBox textSA09 
            Height          =   285
            Left            =   1170
            MaxLength       =   8
            TabIndex        =   10
            Top             =   60
            Width           =   1095
         End
         Begin VB.ComboBox cboSTime 
            Height          =   260
            ItemData        =   "frm160005.frx":212C
            Left            =   1170
            List            =   "frm160005.frx":212E
            Locked          =   -1  'True
            Style           =   2  '單純下拉式
            TabIndex        =   11
            Top             =   360
            Width           =   1005
         End
         Begin VB.ComboBox cboETime 
            Height          =   260
            ItemData        =   "frm160005.frx":2130
            Left            =   1170
            List            =   "frm160005.frx":2132
            Locked          =   -1  'True
            Style           =   2  '單純下拉式
            TabIndex        =   12
            Top             =   690
            Width           =   1005
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "表單編號"
            Height          =   180
            Left            =   420
            TabIndex        =   29
            Top             =   90
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "起日上班時段"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   28
            Top             =   420
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "迄日下班時段"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   27
            Top             =   750
            Width           =   1080
         End
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73980
         MaxLength       =   6
         TabIndex        =   14
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72930
         MaxLength       =   6
         TabIndex        =   15
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -71040
         MaxLength       =   7
         TabIndex        =   16
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -70050
         MaxLength       =   7
         TabIndex        =   17
         Top             =   390
         Width           =   915
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   345
         Left            =   -68670
         TabIndex        =   18
         Top             =   360
         Width           =   915
      End
      Begin VB.ComboBox textSA06 
         Height          =   260
         ItemData        =   "frm160005.frx":2134
         Left            =   1020
         List            =   "frm160005.frx":2136
         TabIndex        =   1
         Top             =   660
         Width           =   1695
      End
      Begin VB.TextBox textSA01 
         Height          =   270
         Left            =   1020
         MaxLength       =   6
         TabIndex        =   0
         Top             =   390
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160005.frx":2138
         Height          =   3615
         Left            =   -74970
         TabIndex        =   21
         Top             =   750
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
         Height          =   200
         Left            =   270
         TabIndex        =   35
         Top             =   3690
         Width           =   7520
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
      Begin MSForms.TextBox txtNote 
         Height          =   1100
         Left            =   3840
         TabIndex        =   13
         Top             =   2520
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
      Begin MSForms.Label Label23 
         Height          =   200
         Left            =   210
         TabIndex        =   34
         Top             =   4020
         Width           =   7790
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13732;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label textSA01_2 
         Height          =   230
         Left            =   1830
         TabIndex        =   33
         Top             =   390
         Width           =   1400
         BackColor       =   12632256
         VariousPropertyBits=   27
         Size            =   "2461;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LblNote 
         Alignment       =   1  '靠右對齊
         Caption         =   "修改/刪除原因："
         Height          =   180
         Left            =   2520
         TabIndex        =   32
         Top             =   2550
         Width           =   1320
      End
      Begin VB.Label Label12 
         Caption         =   $"frm160005.frx":214D
         ForeColor       =   &H000000C0&
         Height          =   1340
         Left            =   5190
         TabIndex        =   31
         Top             =   870
         Width           =   2720
      End
      Begin VB.Label Label11 
         Caption         =   "註：人事室異動假單資料跨月份時，記得要拆單輸入，因會影響算薪水！"
         ForeColor       =   &H000000C0&
         Height          =   380
         Left            =   3570
         TabIndex        =   30
         Top             =   420
         Width           =   4340
      End
      Begin VB.Line Line5 
         X1              =   -70320
         X2              =   -69720
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line4 
         X1              =   -73290
         X2              =   -72600
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74910
         TabIndex        =   25
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   -71610
         TabIndex        =   24
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "假別"
         Height          =   180
         Left            =   630
         TabIndex        =   23
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   22
         Top             =   440
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm160005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/16 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/20 日期欄已修改
'Create by nickc 2006/12/03 copy from frm140401
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
Dim tf_SA As Integer
Dim MyKind As String
Dim m_Day As Integer 'Add By Sindy 2011/9/21 記錄該筆目前DB裡的天數
Dim m_Hour As Double 'Add By Sindy 2014/12/31 記錄該筆目前DB裡的時數
'Add By Sindy 2011/9/22
Dim m_B1019 As String, m_B1004 As String, m_B1005 As String
Dim m_B1006 As String, m_B1007 As String, m_B1008 As String, m_B1009 As String
Dim m_B1010 As String, m_B1017 As String, m_B1028 As String, m_B1029 As String
Dim m_KeyCode As String 'Add By Sindy 2011/10/7
Dim strCallCont As String 'Added by Lydia 2020/03/05 記錄email內容
Dim BolIsApart As String '是否有拆單


'Add By Sindy 2022/10/28
Private Sub cmdABS_Click()
   Me.Hide
   Call frm180301_03.SetParent(Me)
   frm180301_03.txtB1001 = textSA09.Text
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
    If m_EditMode <> 1 Then 'Modify By Sindy 2011/9/22
      MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
    End If
End If
End Sub

Private Sub Form_Initialize()
Set rsA = New ADODB.Recordset
If rsA.State = 1 Then rsA.Close
rsA.CursorLocation = adUseClient
rsA.Open "select * from staff_Absence where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
tf_SA = rsA.Fields.Count
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

   ReDim m_FieldList(tf_SA) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textSA01.BackColor = &H8000000F
   textSA02.BackColor = &H8000000F
   ' 2008/12/22 Add BY SINDY
   textSA03_1.BackColor = &H8000000F
   textSA03_2.BackColor = &H8000000F
   ' 2008/12/22 END
   
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
   Set frm160005 = Nothing
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
         textSA01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         textSA02.Text = Trim(ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 2)))
         ' 2008/12/22 Add BY SINDY
         textSA03_1.Text = Mid(Trim(GRD1.TextMatrix(tmpMouseRow, 2)), Len(Trim(GRD1.TextMatrix(tmpMouseRow, 2))) - 4, 2)
         textSA03_2.Text = Right(Trim(GRD1.TextMatrix(tmpMouseRow, 2)), 2)
         ' 2008/12/22 END
         QueryRecord
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
   If IsNull(rsSrcTmp.Fields("sa10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa10")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("sa10"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sa11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa11")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("sa11"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sa12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa12")) = False Then
         strTemp = rsSrcTmp.Fields("sa12")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sa13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa13")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("sa13"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sa14")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa14")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("sa14"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sa15")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa15")) = False Then
         strTemp = rsSrcTmp.Fields("sa15")
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

If Me.textSA01.Enabled = True Then
   Cancel = False
   textSA01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textSA01.Text = "" Then
    MsgBox "員工編號不可以空白！", vbExclamation
    textSA01.SetFocus
    Exit Function
End If

'Add By Sindy 2011/9/22
If Me.Frame1.Visible = True And Me.textSA09.Enabled = True Then
   If m_EditMode = 1 And textSA09 = "" Then
      MsgBox "表單編號不可空白！", vbExclamation
      textSA09.SetFocus
      Exit Function
   End If
   
   Cancel = False
   textSA09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSA02.Enabled = True Then
   Cancel = False
   textSA02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textSA02.Text = "" Then
    MsgBox "日期起不可以空白！", vbExclamation
    textSA02.SetFocus
    Exit Function
End If
If Me.textSA03_1.Enabled = True Then
   Cancel = False
   textSA03_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
' 2008/12/18 Add BY SINDY
If textSA03_1.Text = "" Or textSA03_1.Text = "00" Then
    MsgBox "必須輸入起始(時)！", vbExclamation
    textSA03_1.SetFocus
    Exit Function
End If
' 2008/12/18 END
If Me.textSA03_2.Enabled = True Then
   Cancel = False
   textSA03_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textSA04.Enabled = True Then
   Cancel = False
   textSA04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
' 2008/12/18 Add BY SINDY
If textSA04.Text = "" Then
    MsgBox "日期迄不可以空白！", vbExclamation
    textSA04.SetFocus
    Exit Function
End If
' 2008/12/18 END
If Me.textSA05_1.Enabled = True Then
   Cancel = False
   textSA05_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
' 2008/12/18 Add BY SINDY
If textSA05_1.Text = "" Or textSA05_1.Text = "00" Then
   MsgBox "必須輸入迄止(時)！", vbExclamation
   textSA05_1.SetFocus
   Exit Function
End If
' 2008/12/18 END
If Me.textSA05_2.Enabled = True Then
   Cancel = False
   textSA05_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2014/6/24
If Left(textSA02, Len(textSA02) - 2) <> Left(textSA04, Len(textSA04) - 2) Then
   MsgBox "假單不可跨月份！", vbExclamation
   textSA04.SetFocus
   Exit Function
End If
'2014/6/24 END

'Add By Sindy 2011/10/17 增加判斷員工代號+日期是否人員已離職
If ChkStaffST04(textSA01, True, textSA02) = True Then
   textSA01.SetFocus
   Exit Function
End If
'Add By Sindy 2011/10/17 增加判斷員工代號+日期是否人員已離職
If ChkStaffST04(textSA01, True, textSA04) = True Then
   textSA01.SetFocus
   Exit Function
End If

If Me.textSA06.Enabled = True Then
   Cancel = False
   textSA06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
' 2008/12/18 Add BY SINDY
If textSA06.Text = "" Then
    MsgBox "假別不可以空白！", vbExclamation
    textSA06.SetFocus
    Exit Function
End If
' 2008/12/18 END
If Me.textSA07.Enabled = True Then
   Cancel = False
   textSA07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textSA08.Enabled = True Then
   Cancel = False
   textSA08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
' 2008/12/18 Add BY SINDY
If textSA07.Text = "" Or textSA08.Text = "" Or _
   (textSA07.Text = "0" And textSA08.Text = "0") Then
    MsgBox "無請假時數！", vbExclamation
    textSA07.SetFocus
    Exit Function
End If
If textSA07.Text <> "" Or textSA08.Text <> "" Then
   If m_EditMode <> 2 Then m_Day = 0: m_Hour = 0
   'Modify By Sindy 2017/11/3 + , m_Hour
   If ChkSA06_08(textSA07, textSA08, textSA01, textSA02, textSA03_1, textSA03_2, textSA04, textSA05_1, textSA05_2, textSA06, m_Day, m_Hour) = False Then
      textSA06.Text = ""
      textSA07.SetFocus
      Exit Function
   End If
   'Add By Sindy 2014/12/31 +健檢假
   If ChkSA06_23(textSA07, textSA08, textSA01, textSA02, textSA03_1, textSA03_2, textSA04, textSA05_1, textSA05_2, textSA06, m_Day, m_Hour) = False Then
      textSA06.Text = ""
      textSA07.SetFocus
      Exit Function
   End If
   'Add By Sindy 2024/12/10 檢查可補休
   If ChkSA06_14(textSA07, textSA08, textSA01, textSA02, textSA03_1, textSA03_2, textSA04, textSA05_1, textSA05_2, textSA06, m_Day, m_Hour) = False Then
      textSA06.Text = ""
      textSA07.SetFocus
      Exit Function
   End If
End If
' 2008/12/18 END

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
   For nIndex = 0 To tf_SA - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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
   
   For nIndex = 0 To tf_SA - 1
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
   Dim strSA01 As String
   Dim strSA02 As String
   Dim strSA03 As String
   Dim rsTmp As New ADODB.Recordset
   
   AddRecord = False
   
   strSA01 = textSA01
   strSA02 = DBDATE(textSA02)
   strSA03 = Trim(textSA03_1.Text & textSA03_2.Text) ' 2008/12/22 Add BY SINDY
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strSA01, strSA02, strSA03) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO staff_Absence ("
   For nIndex = 0 To tf_SA - 1
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
   For nIndex = 0 To tf_SA - 1
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
   
'   'Add By Sindy 2018/3/2 當月的薪資已計算過,發E-MAIL通知財務處
'   strSql = "select sm02,count(*) from SalaryMonth where sm02=" & Left(DBDATE(textSA02), 6) & " group by sm02"
'   intI = 1
'   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      strContent = GetEMailContent(textSA09, strSubject) '先取得E-Mail主旨,本文內容
'      PUB_SendMail strUserNum, "71005", "", textSA01 & GetPrjSalesNM(textSA01) & "補輸請假資料，請重新做每月(" & Left(DBDATE(textSA02), 6) & ")薪資計算！", strContent, , , , , , , , , , True
'   End If
'   '2018/3/2 END
   
   'Add By Sindy 2011/9/22
   If Frame1.Visible = True And textSA09 <> "" Then
      Call ProABSData
   Else
      'Add By Sindy 2019/5/24 假單完成,後續資料檢查及SendMail
      'Modify By Sindy 2024/10/18 +"請假事由：" & txtNote
      Call PUB_AutoM21Receive_SendMail(IIf(textSA09 <> "", textSA09, ""), 表單類別_請假, textSA01, DBDATE(textSA02), Trim(Format("00" & textSA03_1, "00") & Format("00" & textSA03_2, "00")), _
         DBDATE(textSA04), "", Left(DBDATE(textSA02), 6), , , "請假事由：" & txtNote, m_EditMode)
   End If
   
   cnnConnection.CommitTrans
   ' 2008/12/22 Modify BY SINDY
   'If ((strSA01 & strSA02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strSA01 & strSA02) > (m_LastKEY(0) & m_LastKEY(1))) Then
   If ((strSA01 & strSA02 & strSA03) < (m_FirstKEY(0) & m_FirstKEY(1) & m_FirstKEY(2))) Or ((strSA01 & strSA02 & strSA03) > (m_LastKEY(0) & m_LastKEY(1) & m_LastKEY(2))) Then
   ' 2008/12/22 END
      RefreshRange
   End If
   
   ' 2008/12/22 Modify BY SINDY
   'ShowCurrRecord strSA01, DBDATE(strSA02)
   ShowCurrRecord strSA01, DBDATE(strSA02), strSA03
   ' 2008/12/22 END
   
   Set rsTmp = Nothing
   AddRecord = True
   Exit Function
   
ErrHand:
   Set rsTmp = Nothing
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
   Dim strSA01 As String
   Dim strSA02 As String
   Dim strSA03 As String
   Dim strContent As String, strSubject As String
   Dim rsTmp As New ADODB.Recordset
   Dim strOldSa02 As String, strOldSa04 As String 'Added by Lydia 2020/03/05 修改前的日期起、迄
   
   ModRecord = False
   
   strSA01 = m_CurrKEY(0)
   strSA02 = m_CurrKEY(1)
   strSA03 = m_CurrKEY(2) ' 2008/12/22 Add BY SINDY
   
   'Modify By Sindy 2023/11/1 mark,前面有檢查,此處應該不用 ex:89037-1121101 (11206369、11206107)
'   'Add By SINDY 2011/12/5
'   If strSA01 <> textSA01 Or _
'      strSA02 <> DBDATE(textSA02) Or _
'      Val(strSA03) <> Val(Trim(Format("00" & textSA03_1, "00") & Format("00" & textSA03_2, "00"))) Then
'      ' 檢查記錄是否已存在
'      If IsRecordExist(textSA01, DBDATE(textSA02), Trim(Format("00" & textSA03_1, "00") & Format("00" & textSA03_2, "00"))) = True Then
'         strTit = "新增資料"
'         strMsg = "該筆記錄已存在"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         'UpdateCtrlData
'         textSA02.SetFocus
'         Exit Function
'      End If
'   End If
   
   strSql = "begin user_data.user_enabled:=1; UPDATE staff_Absence SET "
   
   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_SA - 1
      strTmp = Empty
      'If nIndex < 9 Or nIndex > 14 Then
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
            'Added by Lydia 2020/03/05 修改前的日期起、迄
            If m_FieldList(nIndex).fiName = "SA02" Then
                strOldSa02 = m_FieldList(nIndex).fiOldData
            ElseIf m_FieldList(nIndex).fiName = "SA04" Then
                strOldSa04 = m_FieldList(nIndex).fiOldData
            End If
            'end 2020/03/05
        'End If
   Next nIndex
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = strSQL & " " & _
'                  "WHERE SA01 = '" & strSA01 & "' and SA02='" & strSA02 & "' ; end; "
   strSql = strSql & " " & _
                  "WHERE SA01 = '" & strSA01 & "' and SA02='" & strSA02 & "' and SA03='" & strSA03 & "' ; end; "
   ' 2008/12/22 END
On Error GoTo ErrHand
      cnnConnection.BeginTrans
        If bDifference = True Then
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
           
'            'Add By Sindy 2018/3/2 當月的薪資已計算過,發E-MAIL通知財務處
'            strSql = "select sm02,count(*) from SalaryMonth where sm02=" & Left(DBDATE(textSA02), 6) & " group by sm02"
'            intI = 1
'            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               strContent = GetEMailContent(textSA09, strSubject) '先取得E-Mail主旨,本文內容
'               PUB_SendMail strUserNum, "71005", "", textSA01 & GetPrjSalesNM(textSA01) & "修改請假資料，請重新做每月(" & Left(DBDATE(textSA02), 6) & ")薪資計算！", strContent, , , , , , , , , , True
'            End If
'            '2018/3/2 END
            
           'Add By Sindy 2011/9/22
           'Modify By Sindy 2019/5/24 電子紙本均要考慮發信問題
           'If Frame1.Visible = True And textSA09 <> "" Then
               Call ProABSData
           'End If
        End If
           
        'Added by Lydia 2020/03/05 提醒至系統更改查名人狀態
        'A請假時間縮短：日期迄=改成前一工作天或當天不論時間，B請假時間延後：日期起=原本是下一工作日改成之後的日期
        If (textSA04 < TransDate(strOldSa04, 1) And (textSA04 = strSrvDate(2) Or DBDATE(textSA04) = CompWorkDay(2, strSrvDate(1), 1))) _
            Or (textSA02 > TransDate(strOldSa02, 1) And DBDATE(strOldSa02) = CompWorkDay(2, strSrvDate(1))) Then
            Call ProcTMQemail("M")
        End If
        'end 2020/03/05
           
      cnnConnection.CommitTrans
      ' 2008/12/22 Modify BY SINDY
      'ShowCurrRecord strSA01, DBDATE(strSA02)
      ShowCurrRecord strSA01, DBDATE(strSA02), strSA03
      ' 2008/12/22 END
      PUB_SendMailCache 'Added by Lydia 2020/03/05 'Added by Lydia 2020/03/05
      
   Set rsTmp = Nothing
   ModRecord = True
   Exit Function
   
ErrHand:
   Set rsTmp = Nothing
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim strSql As String
   Dim strSA01 As String
   Dim strSA02 As String
   Dim strSA03 As String
   'Add By Sindy 2013/2/1
   Dim rsTmp As New ADODB.Recordset
   Dim nResponse
   '2013/2/1 End
   Dim strContent As String, strSubject As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   strSA01 = m_CurrKEY(0)
   strSA02 = m_CurrKEY(1)
   strSA03 = m_CurrKEY(2) ' 2008/12/22 Add BY SINDY
   'Add By Sindy 2013/2/1
   BolIsApart = False
   If textSA09.Text <> "" Then
      strSql = "SELECT * FROM staff_Absence " & _
               "WHERE SA09 = '" & textSA09.Text & "'"
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
   
'   'Add By Sindy 2018/3/12 當月的薪資已計算過,發E-MAIL通知財務處
'   strSql = "select sm02,count(*) from SalaryMonth where sm02=" & Left(DBDATE(textSA02), 6) & " group by sm02"
'   intI = 1
'   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      strContent = GetEMailContent(textSA09, strSubject) '先取得E-Mail主旨,本文內容
'      PUB_SendMail strUserNum, "71005", "", textSA01 & GetPrjSalesNM(textSA01) & "刪除請假資料，請重新做每月(" & Left(DBDATE(textSA02), 6) & ")薪資計算！", strContent, , , , , , , , , , True
'   End If
'   '2018/3/12 END
   
   'Add By Sindy 2011/9/22
   'Modify By Sindy 2019/5/24 電子紙本均要考慮發信問題
'   If Frame1.Visible = True And textSA09.Text <> "" And m_B1019 <> "" Then
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

    'Added by Lydia 2020/03/05 提醒至系統更改查名人狀態
    'A: 日期迄=為前一工作天或當天不論時間，B: 日期起=下一工作日
    If (textSA04 = strSrvDate(2) Or DBDATE(textSA04) = CompWorkDay(2, strSrvDate(1), 1)) _
        Or (DBDATE(textSA02) = CompWorkDay(2, strSrvDate(1))) Then
        Call ProcTMQemail("D")
    End If
    'end 2020/03/05
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "DELETE FROM staff_Absence " & _
'            "WHERE SA01 = '" & strSA01 & "'  and SA02='" & strSA02 & "' "
   strSql = "DELETE FROM staff_Absence " & _
            "WHERE SA01 = '" & strSA01 & "'  and SA02='" & strSA02 & "' and SA03='" & strSA03 & "' "
   ' 2008/12/22 END
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   ' 2008/12/22 Modify BY SINDY
'   If (strSA01 = m_LastKEY(0) And strSA02 = m_LastKEY(1)) Or (strSA01 = m_FirstKEY(0) And strSA02 = m_FirstKEY(1)) Then
'      RefreshRange
'   End If
   If (strSA01 = m_LastKEY(0) And strSA02 = m_LastKEY(1) And strSA03 = m_LastKEY(2)) Or (strSA01 = m_FirstKEY(0) And strSA02 = m_FirstKEY(1) And strSA03 = m_FirstKEY(2)) Then
      RefreshRange
   End If
   'ShowCurrRecord strSA01, DBDATE(strSA02)
   ShowCurrRecord strSA01, DBDATE(strSA02), strSA03
   ' 2008/12/22 END
   PUB_SendMailCache 'Added by Lydia 2020/03/05
   
   DelRecord = True
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strSA01 As String
   Dim strSA02 As String
   Dim strSA03 As String
   
   QueryRecord = False
   strSA01 = textSA01
   ' 2008/12/22 Modify BY SINDY
   'strSA02 = DBDATE(textSA02)
   strSA02 = DBDATE(Trim(textSA02))
   strSA03 = Trim(textSA03_1.Text & textSA03_2.Text) ' 2008/12/22 Add BY SINDY
   ' 2008/12/22 END
   
   If IsRecordExist(strSA01, strSA02, strSA03) = True Then
      m_CurrKEY(0) = strSA01
      m_CurrKEY(1) = strSA02
      m_CurrKEY(2) = strSA03 ' 2008/12/22 Add BY SINDY
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
         Call Pub_GetSpecWorkHour(textSA01, textSA02) 'Add By Sindy 2019/10/1 上班時數為特殊者
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
            ' 2008/12/22 Modify BY SINDY
            'ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1), m_CurrKEY(2)
            ' 2008/12/22 END
         Else
            Exit Function
         End If
      Case 4: '查詢
         ' 2008/12/22 Modify BY SINDY
         'If textSA01 <> "" And textSA02 <> "" Then
         If textSA01 <> "" And textSA02 <> "" _
            And textSA03_1 <> "" And textSA03_2 <> "" Then
         ' 2008/12/22 END
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            ' 2008/12/17 ADD BY SINDY
            If textSA01 = "" Or textSA02 = "" Or _
               textSA03_1 = "" Or textSA03_2 = "" Then
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
      Case 1: If Me.Visible = True Then textSA01.SetFocus
      Case 2: If Me.Visible = True Then textSA03_1.SetFocus
      Case 4: If Me.Visible = True Then textSA01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strHHMM_S As String, strHHMM_E As String
   
   IsRecordExist = False
   ' 2008/12/19 Add BY SINDY
   '比較員編和起始日
'   strSQL = "SELECT * FROM staff_Absence " & _
'            "WHERE SA01 = '" & strKEY01 & "'  and SA02='" & strKEY02 & "'  "
   '比較員編和起(迄)日
'   strSQL = "SELECT * FROM staff_Absence " & _
'            "WHERE SA01 = '" & strKEY01 & "' and ('" & strKEY02 & "' between SA02 and SA04) "
   '比較員編和起始日及時間(起)
   strSql = "SELECT SA01 FROM staff_Absence" & _
            " WHERE SA01='" & strKEY01 & "' and SA02='" & strKEY02 & "' and SA03='" & strKEY03 & "'"
   ' 2008/12/19 END
   'Add By Sindy 2023/11/1 修改時,排除檢查此表單
   If m_EditMode = 2 Then
      strSql = strSql & " and SA09<>'" & textSA09 & "'"
   End If
   '2023/11/1 END
   'Add By Sindy 2022/7/19
   'Modify By Sindy 2023/11/29 +已核准 ex:11206863(98020)
   'Modify By Sindy 2023/11/29 有假單上下分開填,8:00~~11:00請假;11:00~24:00出差
   If textSA04 <> "" And textSA05_1 <> "" Then
      strHHMM_S = PUB_DTtoDateAdd(strKEY02, strKEY03, 1)
      strHHMM_E = PUB_DTtoDateAdd(textSA04, textSA05_1 & textSA05_2, -1)
   '2023/11/29 END
      'Right("0000" & strKEY03, 4) => strHHMM_S
      'Format("0" & textSA05_1, "00") & Format("0" & textSA05_2, "00") => strHHMM_E
      strSql = strSql & " union SELECT B1003 FROM abs010" & _
               " WHERE B1003 = '" & strKEY01 & "'" & _
               " and (" & strKEY02 & strHHMM_S & " between B1004||substr('0'||B1005,-4) and B1006||substr('0'||B1007,-4)" & _
               " or " & DBDATE(textSA04) & strHHMM_E & " between B1004||substr('0'||B1005,-4) and B1006||substr('0'||B1007,-4)" & _
               ") and B1018 not in('" & 註銷 & "','" & 已核准 & "')"
   End If
   '2022/7/19 END
   'Add By Sindy 2023/11/1 修改時,排除檢查此表單
   'Modify By Sindy 2025/1/6 新增時,排除檢查此表單
   'Modify By Sindy 2025/1/21 + And Frame1.Visible = True
   If (m_EditMode = 2 Or m_EditMode = 1) And Frame1.Visible = True Then
      strSql = strSql & " and B1001<>'" & textSA09 & "'"
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
      m_CurrKEY(2) = strKEY03 ' 2008/12/22 Add BY SINDY
   Else
      ' 2008/12/22 Modify BY SINDY
'      strSQL = "SELECT SA01,SA02 FROM staff_Absence " & _
'               "WHERE SA01 = '" & m_CurrKEY(0) & "' and SA02='" & m_CurrKEY(1) & "' "
      strSql = "SELECT SA01,SA02,SA03 FROM staff_Absence " & _
               "WHERE SA01 = '" & m_CurrKEY(0) & "' and SA02='" & m_CurrKEY(1) & "' and SA03='" & m_CurrKEY(2) & "' "
      ' 2008/12/22 END
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SA01")
         If IsNull(rsTmp.Fields("SA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SA02")
         ' 2008/12/22 Add BY SINDY
         If IsNull(rsTmp.Fields("SA03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SA03")
         ' 2008/12/22 END
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      ' 2008/12/22 Modify BY SINDY
'      strSQL = "SELECT SA01,SA02,SA03 FROM staff_Absence " & _
'               "WHERE SA02 = (SELECT MIN(SA02) FROM staff_Absence where SA01=(select min(SA01) from staff_Absence) ) and SA01=(select min(SA01) from staff_Absence) "
      strSql = "SELECT SA01,SA02,SA03 FROM staff_Absence " & _
               "WHERE SA02 = (SELECT MIN(SA02) FROM staff_Absence where SA01=(select min(SA01) from staff_Absence) ) and SA01=(select min(SA01) from staff_Absence) Order BY SA03 ASC"
      ' 2008/12/22 END
      
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SA01")
         If IsNull(rsTmp.Fields("SA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SA02")
         ' 2008/12/22 Add BY SINDY
         If IsNull(rsTmp.Fields("SA03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SA03")
         ' 2008/12/22 END
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
   m_CurrKEY(2) = m_FirstKEY(2) ' 2008/12/22 Add BY SINDY
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 2008/12/22 Modify BY SINDY
   'If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) _
      And m_CurrKEY(2) = m_FirstKEY(2) Then
   ' 2008/12/22 END
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   ' 2008/12/22 Add BY SINDY
   strSql = "SELECT SA01,SA02,SA03 FROM staff_Absence " & _
            "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
                          "SA02 = '" & m_CurrKEY(1) & "' AND " & _
                  "SA03 = (SELECT MAX(SA03) FROM staff_Absence " & _
                          "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
                                        "SA02 = '" & m_CurrKEY(1) & "' AND " & _
                                        "SA03 < '" & m_CurrKEY(2) & "' ) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SA01")
      If IsNull(rsTmp.Fields("SA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SA02")
      If IsNull(rsTmp.Fields("SA03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SA03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   ' 2008/12/22 END
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT SA01,SA02 FROM staff_Absence " & _
'            "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
'                  "SA02 = (SELECT MAX(SA02) FROM staff_Absence " & _
'                          "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
'                                "SA02 < '" & m_CurrKEY(1) & "' )"
   strSql = "SELECT SA01,SA02,SA03 FROM staff_Absence " & _
            "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
                  "SA02 = (SELECT MAX(SA02) FROM staff_Absence " & _
                          "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
                                "SA02 < '" & m_CurrKEY(1) & "') Order BY SA03 DESC"
   ' 2008/12/22 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SA01")
      If IsNull(rsTmp.Fields("SA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SA02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("SA03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SA03")
      ' 2008/12/22 END
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT SA01,SA02 FROM staff_Absence " & _
'            "WHERE SA01 = (SELECT MAX(SA01) FROM staff_Absence " & _
'                           "WHERE SA01 < '" & m_CurrKEY(0) & "') AND " & _
'                  "SA02 = (SELECT MAX(SA02) FROM staff_Absence " & _
'                           "WHERE SA01 = (SELECT MAX(SA01) FROM staff_Absence " & _
'                                          "WHERE SA01 < '" & m_CurrKEY(0) & "')) "
   strSql = "SELECT SA01,SA02,SA03 FROM staff_Absence " & _
            "WHERE SA01 = (SELECT MAX(SA01) FROM staff_Absence " & _
                           "WHERE SA01 < '" & m_CurrKEY(0) & "') AND " & _
                  "SA02 = (SELECT MAX(SA02) FROM staff_Absence " & _
                           "WHERE SA01 = (SELECT MAX(SA01) FROM staff_Absence " & _
                                          "WHERE SA01 < '" & m_CurrKEY(0) & "')) Order BY SA03 DESC"
   ' 2008/12/22 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SA01")
      If IsNull(rsTmp.Fields("SA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SA02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("SA03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SA03")
      ' 2008/12/22 END
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
   
   ' 2008/12/22 Modify BY SINDY
   'If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) And m_CurrKEY(2) = m_LastKEY(2) Then
   ' 2008/12/22 END
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   ' 2008/12/22 Add BY SINDY
   strSql = "SELECT SA01,SA02,SA03 FROM staff_Absence " & _
            "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
                          "SA02 = '" & m_CurrKEY(1) & "' AND " & _
                  "SA03 = (SELECT MIN(SA03) FROM staff_Absence " & _
                          "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
                                        "SA02 = '" & m_CurrKEY(1) & "' AND " & _
                                        "SA03 > '" & m_CurrKEY(2) & "' ) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SA01")
      If IsNull(rsTmp.Fields("SA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SA02")
      If IsNull(rsTmp.Fields("SA03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SA03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   ' 2008/12/22 END
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT SA01,SA02 FROM staff_Absence " & _
'            "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
'                  "SA02 = (SELECT MIN(SA02) FROM staff_Absence " & _
'                          "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
'                                "SA02 > '" & m_CurrKEY(1) & "' )"
   strSql = "SELECT SA01,SA02,SA03 FROM staff_Absence " & _
            "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
                  "SA02 = (SELECT MIN(SA02) FROM staff_Absence " & _
                          "WHERE SA01 = '" & m_CurrKEY(0) & "' AND " & _
                                "SA02 > '" & m_CurrKEY(1) & "' ) Order BY SA03 ASC "
   ' 2008/12/22 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SA01")
      If IsNull(rsTmp.Fields("SA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SA02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("SA03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SA03")
      ' 2008/12/22 END
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT SA01,SA02 FROM staff_Absence " & _
'            "WHERE SA01 = (SELECT MIN(SA01) FROM staff_Absence " & _
'                           "WHERE SA01 > '" & m_CurrKEY(0) & "') AND " & _
'                  "SA02 = (SELECT MIN(SA02) FROM staff_Absence " & _
'                           "WHERE SA01 = (SELECT MIN(SA01) FROM staff_Absence " & _
'                                          "WHERE SA01 > '" & m_CurrKEY(0) & "')) "
   strSql = "SELECT SA01,SA02,SA03 FROM staff_Absence " & _
            "WHERE SA01 = (SELECT MIN(SA01) FROM staff_Absence " & _
                           "WHERE SA01 > '" & m_CurrKEY(0) & "') AND " & _
                  "SA02 = (SELECT MIN(SA02) FROM staff_Absence " & _
                           "WHERE SA01 = (SELECT MIN(SA01) FROM staff_Absence " & _
                                          "WHERE SA01 > '" & m_CurrKEY(0) & "')) Order BY SA03 ASC "
   ' 2008/12/22 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SA01")
      If IsNull(rsTmp.Fields("SA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SA02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("SA03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SA03")
      ' 2008/12/22 END
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
   m_CurrKEY(2) = m_LastKEY(2) ' 2008/12/22 Add BY SINDY
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
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT SA01,SA02 FROM staff_Absence " & _
'            "WHERE SA01 = (SELECT MIN(SA01) FROM staff_Absence) AND " & _
'                  "SA02 = (SELECT MIN(SA02) FROM staff_Absence " & _
'                           "WHERE SA01 = (SELECT MIN(SA01) FROM staff_Absence)) "
   strSql = "SELECT SA01,SA02,SA03 FROM staff_Absence " & _
            "WHERE SA01 = (SELECT MIN(SA01) FROM staff_Absence) AND " & _
                  "SA02 = (SELECT MIN(SA02) FROM staff_Absence " & _
                           "WHERE SA01 = (SELECT MIN(SA01) FROM staff_Absence)) Order BY SA03 ASC"
   ' 2008/12/22 END
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SA01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("SA01")
      If IsNull(rsTmp.Fields("SA02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("SA02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("SA03")) = False Then: m_FirstKEY(2) = rsTmp.Fields("SA03")
      ' 2008/12/22 END
   End If
   rsTmp.Close
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT SA01,SA02 FROM staff_Absence " & _
'            "WHERE SA01 = (SELECT MAX(SA01) FROM staff_Absence) AND " & _
'                  "SA02 = (SELECT MAX(SA02) FROM staff_Absence " & _
'                           "WHERE SA01 = (SELECT MAX(SA01) FROM staff_Absence)) "
   strSql = "SELECT SA01,SA02,SA03 FROM staff_Absence " & _
            "WHERE SA01 = (SELECT MAX(SA01) FROM staff_Absence) AND " & _
                  "SA02 = (SELECT MAX(SA02) FROM staff_Absence " & _
                           "WHERE SA01 = (SELECT MAX(SA01) FROM staff_Absence)) Order BY SA03 DESC"
   ' 2008/12/22 END
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SA01")) = False Then: m_LastKEY(0) = rsTmp.Fields("SA01")
      If IsNull(rsTmp.Fields("SA02")) = False Then: m_LastKEY(1) = rsTmp.Fields("SA02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("SA03")) = False Then: m_LastKEY(2) = rsTmp.Fields("SA03")
      ' 2008/12/22 END
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim i As Integer, j As Integer
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT * FROM staff_Absence " & _
'            "WHERE SA01='" & m_CurrKEY(0) & "' and SA02 = '" & m_CurrKEY(1) & "'   "
   strSql = "SELECT * FROM staff_Absence " & _
            "WHERE SA01='" & m_CurrKEY(0) & "' and SA02 = '" & m_CurrKEY(1) & "' and SA03 = '" & m_CurrKEY(2) & "' "
   ' 2008/12/22 END
   
   FrameSA19.Visible = False 'Add By Sindy 2024/10/18
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("SA01")) = False Then: textSA01 = rsTmp.Fields("SA01")
      If IsNull(rsTmp.Fields("SA02")) = False Then: textSA02 = TAIWANDATE(rsTmp.Fields("SA02"))
      If IsNull(rsTmp.Fields("SA03")) = False Then: textSA03_1 = Mid(Format(rsTmp.Fields("SA03"), "0000"), 1, 2): textSA03_2 = Mid(Format(rsTmp.Fields("SA03"), "0000"), 3, 2)
      If IsNull(rsTmp.Fields("SA04")) = False Then: textSA04 = TAIWANDATE(rsTmp.Fields("SA04"))
      If IsNull(rsTmp.Fields("SA05")) = False Then: textSA05_1 = Mid(Format(rsTmp.Fields("SA05"), "0000"), 1, 2): textSA05_2 = Mid(Format(rsTmp.Fields("SA05"), "0000"), 3, 2)
      If IsNull(rsTmp.Fields("SA06")) = False Then: textSA06 = rsTmp.Fields("SA06"): textSA06_Validate False
      If IsNull(rsTmp.Fields("SA07")) = False Then: textSA07 = rsTmp.Fields("SA07")
      If IsNull(rsTmp.Fields("SA08")) = False Then: textSA08 = rsTmp.Fields("SA08")
      
      'Add By Sindy 2021/8/11
      SetB102829Combo cboSTime, 1, textSA02, textSA01
      SetB102829Combo cboETime, 2, textSA02, textSA01
      '2021/8/11 END
      
      'Add By Sindy 2021/12/27
      LblIsApart.Caption = "跨月份拆單，此筆資料為：" & ChangeTStringToTDateString(textSA02) & " " & textSA03_1 & ":" & textSA03_2 & " ~ " & ChangeTStringToTDateString(textSA04) & " " & textSA05_1 & ":" & textSA05_2
      '2021/12/27 End
      
      'Add By Sindy 2011/9/22
      If IsNull(rsTmp.Fields("SA09")) = False Then
         textSA09 = rsTmp.Fields("SA09")
         Call GetABS010(True)
         Frame1.Visible = True
         
         'Add By Sindy 2021/12/27
         BolIsApart = False
         If textSA09.Text <> "" Then
            strSql = "SELECT * FROM staff_Absence " & _
                     "WHERE SA09 = '" & textSA09.Text & "'"
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
         m_B1004 = DBDATE(textSA02)
         m_B1005 = textSA03_1 & textSA03_2
         m_B1006 = DBDATE(textSA04)
         m_B1007 = textSA05_1 & textSA05_2
         m_B1008 = Left(textSA06, 2)
         m_B1009 = textSA07
         m_B1010 = textSA08
         If Not IsNull(rsTmp.Fields("SA16")) Then m_B1028 = IIf(Format(rsTmp.Fields("SA16"), "hhmm") = "0000", "", Format(rsTmp.Fields("SA16"), "hhmm"))
         If Not IsNull(rsTmp.Fields("SA17")) Then m_B1029 = IIf(Format(rsTmp.Fields("SA17"), "hhmm") = "0000", "", Format(rsTmp.Fields("SA17"), "hhmm"))
         '2019/5/24 END
         Frame1.Visible = False
         'Add By Sindy 2024/10/18
         If IsNull(rsTmp.Fields("SA19")) = False Then
            textSA19_2 = rsTmp.Fields("SA19")
            FrameSA19.Visible = True
            FrameSA19.Top = Frame1.Top
            FrameSA19.Left = Frame1.Left
         End If
         '2024/10/18 END
      End If
      'Add By Sindy 2013/2/1
      '顯示起日上班時段,迄日下班時段至畫面上
      If Not IsNull(rsTmp.Fields("SA16")) Then
         For i = 0 To cboSTime.ListCount - 1
            If cboSTime.List(i) = Format(rsTmp.Fields("SA16"), "00:00") Then
               cboSTime.ListIndex = i
               Exit For
            End If
         Next i
      End If
      If Not IsNull(rsTmp.Fields("SA17")) Then
         For i = 0 To cboETime.ListCount - 1
            If cboETime.List(i) = Format(rsTmp.Fields("SA17"), "00:00") Then
               cboETime.ListIndex = i
               Exit For
            End If
         Next i
      End If
      '2013/2/1 End
      
      'Add By Sindy 2011/9/21 若為特別假時,記錄該筆目前DB裡的時數
'      m_Day = 0
'      If IsNull(rsTmp.Fields("SA06")) = False Then
'         If rsTmp.Fields("SA06") = "08" Then
            m_Day = Val("" & rsTmp.Fields("SA07"))
'         End If
'      End If
      m_Hour = Val("" & rsTmp.Fields("SA08")) 'Add By Sindy 2014/12/31 記錄該筆目前DB裡的時數
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp

       textSA06_Validate False
       textSA01_2 = GetStaffName(textSA01, True)
   End If

   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset
strSql = ""
If txt1(0) <> "" Then
    strSql = strSql & " and SA01>='" & txt1(0) & "' "
End If
If txt1(1) <> "" Then
    strSql = strSql & " and SA01<='" & txt1(1) & "' "
End If
'Modify By Sindy 2019/10/1
If txt1(2) <> "" Then
    strSql = strSql & " and SA02>='" & DBDATE(txt1(2)) & "' "
End If
If txt1(3) <> "" Then
    strSql = strSql & " and SA04<='" & DBDATE(txt1(3)) & "' "
End If
'If txt1(2) <> "" And txt1(3) <> "" Then
'   strSql = strSql & " AND ('" & DBDATE(txt1(2)) & "' BETWEEN SA02 AND SA04 or '" & DBDATE(txt1(3)) & "' BETWEEN SA02 AND SA04) "
'End If
'2019/10/1 END
'抓取資料
strSql = "SELECT SA01,s1.st02,sqldateT(SA02)||' '||substr(ltrim(to_char('0000'||to_char(SA03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SA03),'0000')),3,2),sqldateT(SA04)||' '||substr(ltrim(to_char('0000'||to_char(SA05),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SA05),'0000')),3,2),ac02||' '||ac03,SA07,SA08,SA09  FROM staff_Absence,staff s1,allcode where SA01=s1.st01(+) and '04'=ac01(+) and sa06=ac02(+) " & strSql & _
        " order by SA02,SA03,SA01 "
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
   
   'Add By Sindy 2014/3/21 上班時數為特殊者
   Call Pub_GetSpecWorkHour(textSA01, textSA02)
   
   nResponse = False
   textSA01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA03_1_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA03_2_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA04_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA05_1_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA05_2_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA06_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA07_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA08_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   
   CheckDataValid = True
   
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSA01.Locked = bEnable
   If bEnable Then textSA01.BackColor = &H8000000F Else textSA01.BackColor = &H80000005
   If m_EditMode <> "2" Then 'Modify By Sindy 2011/12/5
      textSA02.Locked = bEnable
      If bEnable Then textSA02.BackColor = &H8000000F Else textSA02.BackColor = &H80000005
      ' 2008/12/22 Add BY SINDY
      textSA03_1.Locked = bEnable
      textSA03_2.Locked = bEnable
      If bEnable Then textSA03_1.BackColor = &H8000000F Else textSA03_1.BackColor = &H80000005
      If bEnable Then textSA03_2.BackColor = &H8000000F Else textSA03_2.BackColor = &H80000005
      ' 2008/12/22 END
   End If
   'Add By Sindy 2011/9/22
   textSA09.Locked = bEnable
   If bEnable Then textSA09.BackColor = &H8000000F Else textSA09.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   textSA01.Locked = bEnable
   textSA02.Locked = bEnable
   If bEnable Then textSA01.BackColor = &H8000000F Else textSA01.BackColor = &H80000005
   If bEnable Then textSA02.BackColor = &H8000000F Else textSA02.BackColor = &H80000005
   textSA03_1.Locked = bEnable
   textSA03_2.Locked = bEnable
   ' 2008/12/22 Add BY SINDY
   If bEnable Then textSA03_1.BackColor = &H8000000F Else textSA03_1.BackColor = &H80000005
   If bEnable Then textSA03_2.BackColor = &H8000000F Else textSA03_2.BackColor = &H80000005
   ' 2008/12/22 END
   textSA04.Locked = bEnable
   textSA05_1.Locked = bEnable
   textSA05_2.Locked = bEnable
   textSA06.Locked = bEnable
   textSA07.Locked = bEnable
   textSA08.Locked = bEnable
   'Add By Sindy 2011/9/22
   textSA09.Locked = bEnable
   If bEnable Then textSA09.BackColor = &H8000000F Else textSA09.BackColor = &H80000005
   cboSTime.Locked = bEnable
   cboETime.Locked = bEnable
'   txtNote.Locked = bEnable
   
   'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'   '有表單編號的資料時,日及時欄位鎖住
'   If textSA09 <> "" Then
'      textSA07.Enabled = False
'      textSA08.Enabled = False
'   Else
'      textSA07.Enabled = True
'      textSA08.Enabled = True
'   End If
End Sub

Private Sub ClearField()
   Dim nIndex As Integer
   textSA01 = Empty
   If m_EditMode = 1 Then textSA01.SetFocus
   textSA01_2 = Empty
   textSA02 = Empty
   textSA03_1 = Empty
   textSA03_2 = Empty
   textSA04 = Empty
   textSA05_1 = Empty
   textSA05_2 = Empty
   textSA06 = Empty
   textSA07 = Empty
   textSA08 = Empty
   
   'Add By Sindy 2011/9/22
   textSA09 = Empty
   Frame1.Visible = False
   m_B1019 = Empty: m_B1004 = Empty: m_B1005 = Empty: m_B1006 = Empty: m_B1007 = Empty
   m_B1008 = Empty: m_B1009 = Empty: m_B1010 = Empty: m_B1017 = Empty: m_B1028 = Empty
   m_B1029 = Empty
   txtNote = Empty
   
   Label23 = Empty
   SetGrd
   For nIndex = 0 To tf_SA - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
   LblIsApart.Visible = False 'Add By Sindy 2021/12/27
   cmdABS.Visible = False 'Add By Sindy 2022/10/28
   textSA19_2 = Empty 'Add By Sindy 2024/10/18
End Sub

Private Sub UpdateFieldNewData()
    Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "SA01", textSA01
      'Modify By Sindy 2011/12/5 起始日期及起始時間開放修改
'      SetFieldNewData "SA02", DBDATE(textSA02)
'      SetFieldNewData "SA03", textSA03_1 & Format("00" & textSA03_2, "00")
   End If
   SetFieldNewData "SA02", DBDATE(textSA02)
   SetFieldNewData "SA03", textSA03_1 & Format("00" & textSA03_2, "00")
   SetFieldNewData "SA04", DBDATE(textSA04)
   SetFieldNewData "SA05", textSA05_1 & Format("00" & textSA05_2, "00")
   If textSA06.Text <> "" Then
        MyArr = Split(textSA06, " ")
        SetFieldNewData "SA06", MyArr(0)
   Else
        SetFieldNewData "SA06", Empty
   End If
   SetFieldNewData "SA07", textSA07
   SetFieldNewData "SA08", textSA08
   SetFieldNewData "SA09", textSA09 'Add By Sindy 2011/9/22
   'Add By Sindy 2024/10/18 無須走簽核流程且為新增假單時,才要儲存請假事由
   If Frame1.Visible = False And m_EditMode = 1 Then
      SetFieldNewData "SA19", txtNote
   End If
   '2024/10/18 END
   SetFieldNewData "SA16", IIf(Frame1.Visible = False, "", Format(cboSTime, "hhmm"))
   SetFieldNewData "SA17", IIf(Frame1.Visible = False, "", Format(cboETime, "hhmm"))
   '2013/2/1 End
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To tf_SA
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "SA" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 2, 3, 4, 5, 7, 8, 16, 17:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
textSA06.Clear
Dim MyRs As New ADODB.Recordset
Set MyRs = New ADODB.Recordset
If MyRs.State = 1 Then MyRs.Close
' 2008/12/18 Modify BY SINDY
' 排除不須要的代碼 : 01.忘打卡 02.遲到 03.曠職 04.出差 16.加班
'strSQL = "select ac02||' '||ac03 from allcode where ac01='04' order by ac02"
strSql = "select ac02||' '||ac03 from allcode where ac01='04' and ac02 not in ('01','02','03','04','16') order by ac02"
' 2008/12/18 END
MyRs.CursorLocation = adUseClient
MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If MyRs.RecordCount <> 0 Then
    While Not MyRs.EOF
        textSA06.AddItem "" & MyRs.Fields(0).Value
        MyRs.MoveNext
    Wend
End If
SetGrd
End Sub

Private Sub textSA01_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSA01
End If
End Sub

Private Sub textSA01_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2011/9/22
Private Sub textSA01_LostFocus()
   '若輸入的員工代號為可寄信者,必須輸入表單編號
   If Frame1.Visible = True Then If textSA09.Enabled = True Then textSA09.SetFocus
   '新增狀態將游標停在員工代號的欄位
   If m_EditMode = 1 And textSA01 = "" Then textSA01.SetFocus
End Sub

Private Sub textSA01_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

If textSA01.Text = "" Then
   textSA01_2 = "" ' 2008/12/18 ADD BY SINDY
   'Add By Sindy 2011/9/22 預設值
   Frame1.Visible = False
End If

If m_EditMode <> 0 And textSA01 <> "" Then
    textSA01_2 = GetStaffName(textSA01, True)
    ' 2008/12/18 ADD BY SINDY
    ' 檢查員工編號規則
    If ChkStaffID(textSA01) Then
       Call textSA01_GotFocus
       Cancel = True
       Exit Sub
    End If
    ' 2008/12/18 END
    If textSA01_2 = "" Then
        MsgBox "員工編號錯誤！查無此員工！", vbInformation
        Call textSA01_GotFocus ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    LblNote.Caption = "修改/刪除原因：" 'Add By Sindy 2024/10/18
    If m_KeyCode = vbKeyF2 Then '按新增時
      LblNote.Caption = "請假事由：" 'Add By Sindy 2024/10/18
      'Add By Sindy 2011/9/22 檢查此員工是否為"不寄信"
      If ChkStaffST14(textSA01, False) = False Then
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
If m_EditMode = 1 And textSA01 <> "" Then
    If textSA02 <> "" And Val(textSA03_1) > 0 And Val(textSA03_2) > 0 Then
      If IsRecordExist(textSA01, DBDATE(textSA02), Trim(textSA03_1.Text & textSA03_2.Text)) = True And textSA01.Enabled = True And textSA01.Locked = False Then
          MsgBox "該員工當天已有資料，請修改！", vbInformation
          Call textSA01_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
End If
End Sub

'Private Sub textSA02_Change()
'Dim tmpCalH As String
'If m_EditMode <> 0 Then
'    If Trim(textSA02) <> "" And Trim(textSA03_1) <> "" And Trim(textSA03_2) <> "" And Trim(textSA04) <> "" And Trim(textSA05_1) <> "" And Trim(textSA05_2) <> "" Then
'        If CheckIsTaiwanDate(textSA02, False) = True And CheckIsTaiwanDate(textSA04, False) = True Then
'            tmpCalH = CalDateTime(textSA02 & Format(textSA03_1, "00") & Format(textSA03_2, "00"), textSA04 & Format(textSA05_1, "00") & Format(textSA05_2, "00"))
'            textSA07 = Val(tmpCalH) \ 8
'            textSA08 = Val(tmpCalH) - (Val(textSA07) * 8)
'        Else
'            textSA07 = 0
'            textSA08 = 0
'        End If
'    Else
'        textSA07 = 0
'        textSA08 = 0
'    End If
'End If
'End Sub

Private Sub textSA02_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSA02
    CloseIme
End If
End Sub

Private Sub textSA02_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSA02_Validate(Cancel As Boolean)
If textSA03_1 = "" Then textSA03_1 = "00"
If textSA03_2 = "" Then textSA03_2 = "00"

If m_EditMode <> 0 And textSA02 <> "" Then
    If CheckIsTaiwanDate(textSA02, False) = False Then
        Call textSA02_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
        Exit Sub
    End If
'    If ChkWorkDay(DBDATE(textSA02)) = False Then
'        Call textSA02_GotFocus   ' 2008/12/18 ADD BY SINDY
'        Cancel = True
'        MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
'        Exit Sub
'    End If
    If textSA02 <> "" And textSA04 <> "" Then
      If RunNick2(textSA02, textSA04) Then
          Call textSA02_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
'    'Add By Sindy 2011/9/22
'    If Frame1.Visible = True And textSA09 <> "" And m_B1019 <> "" Then '有表單編號
'      If Val(DBDATE(textSA02)) < Val(DBDATE(m_B1004)) Or Val(DBDATE(textSA02)) > Val(DBDATE(m_B1006)) Then
'         MsgBox "請假日期只能改少不能改多！", vbInformation, "輸入日期錯誤"
'         Call textSA02_GotFocus
'         Cancel = True
'         Exit Sub
'      End If
'    End If
   
   'Add By Sindy 2021/8/13
   If m_EditMode = 1 Then '新增
      SetB102829Combo cboSTime, 1, textSA02, textSA01
      SetB102829Combo cboETime, 2, textSA02, textSA01
   End If
   '2021/8/13 END
   
   'Add By Sindy 2011/9/22
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSA09 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   End If
End If
If m_EditMode = 1 And textSA02 <> "" Then
    If textSA01 <> "" And textSA02 <> "" _
         And Val(textSA03_1) > 0 And Val(textSA03_2) > 0 Then
      If IsRecordExist(textSA01, DBDATE(textSA02), Trim(textSA03_1.Text & textSA03_2.Text)) = True And textSA01.Enabled = True And textSA01.Locked = False Then
          MsgBox "該員工當天已有資料，請修改！", vbInformation
          Call textSA02_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   ' 2008/12/18 Modify SINDY
'   arrGridHeadText = Array("員工編號", "姓名", "起始日期時間", "結束日期時間", "假別", "天數", "時數", "職務代理人編號", "職務代理人")
'   arrGridHeadWidth = Array(800, 1200, 1200, 1200, 1200, 800, 800, 800, 1200)
   arrGridHeadText = Array("員工編號", "姓名", "起始日期時間", "結束日期時間", "假別", "天數", "時數", "職務代理人")
   arrGridHeadWidth = Array(800, 1200, 1200, 1200, 1200, 800, 800, 0)
   ' 2008/12/18 END
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

'Private Sub textSA03_1_Change()
'Dim tmpCalH As String
'If m_EditMode <> 0 Then
'    If Trim(textSA02) <> "" And Trim(textSA03_1) <> "" And Trim(textSA03_2) <> "" And Trim(textSA04) <> "" And Trim(textSA05_1) <> "" And Trim(textSA05_2) <> "" Then
'        If CheckIsTaiwanDate(textSA02, False) = True And CheckIsTaiwanDate(textSA04, False) = True Then
'            tmpCalH = CalDateTime(textSA02 & Format(textSA03_1, "00") & Format(textSA03_2, "00"), textSA04 & Format(textSA05_1, "00") & Format(textSA05_2, "00"))
'            textSA07 = Val(tmpCalH) \ 8
'            textSA08 = Val(tmpCalH) - (Val(textSA07) * 8)
'        Else
'            textSA07 = 0
'            textSA08 = 0
'        End If
'    Else
'        textSA07 = 0
'        textSA08 = 0
'    End If
'End If
'End Sub

Private Sub textSA03_1_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSA03_1
End If
End Sub

Private Sub textSA03_1_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSA03_1_Validate(Cancel As Boolean)
If textSA03_1 = "" Then textSA03_1 = "00"

If m_EditMode = 1 And textSA03_1 <> "" Then
    If CheckLengthIsOK(textSA03_1, textSA03_1.MaxLength) = False Then
        Call textSA03_1_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    ' 2008/12/19 Add BY SINDY
    If textSA03_1.Text > 24 Then
       Call textSA03_1_GotFocus
       MsgBox "不可超過24時!", vbExclamation + vbOKOnly
       Cancel = True
       Exit Sub
    End If
    If textSA01 <> "" And textSA02 <> "" _
         And Val(textSA03_1) > 0 And Val(textSA03_2) > 0 Then
      If IsRecordExist(textSA01, DBDATE(textSA02), Trim(textSA03_1.Text & textSA03_2.Text)) = True And textSA01.Enabled = True And textSA01.Locked = False Then
          MsgBox "該員工當天已有資料，請修改！", vbInformation
          Call textSA03_1_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
    ' 2008/12/19 END
   'Add By Sindy 2011/9/22
   'If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSA09 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   'End If
End If
CloseIme
End Sub

'Private Sub textSA03_2_Change()
'Dim tmpCalH As String
'If m_EditMode <> 0 Then
'    If Trim(textSA02) <> "" And Trim(textSA03_1) <> "" And Trim(textSA03_2) <> "" And Trim(textSA04) <> "" And Trim(textSA05_1) <> "" And Trim(textSA05_2) <> "" Then
'        If CheckIsTaiwanDate(textSA02, False) = True And CheckIsTaiwanDate(textSA04, False) = True Then
'            tmpCalH = CalDateTime(textSA02 & Format(textSA03_1, "00") & Format(textSA03_2, "00"), textSA04 & Format(textSA05_1, "00") & Format(textSA05_2, "00"))
'            textSA07 = Val(tmpCalH) \ 8
'            textSA08 = Val(tmpCalH) - (Val(textSA07) * 8)
'        Else
'            textSA07 = 0
'            textSA08 = 0
'        End If
'    Else
'        textSA07 = 0
'        textSA08 = 0
'    End If
'End If
'End Sub

Private Sub textSA03_2_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSA03_2
End If
End Sub

Private Sub textSA03_2_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSA03_2_Validate(Cancel As Boolean)
If textSA03_2 = "" Then textSA03_2 = "00"

If m_EditMode = 1 And textSA03_2 <> "" Then
    If CheckLengthIsOK(textSA03_2, textSA03_2.MaxLength) = False Then
        Call textSA03_2_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    ' 2008/12/18 ADD BY SINDY
    If textSA03_2.Text > 59 Then
       Call textSA03_2_GotFocus
       MsgBox "不可超過59分!", vbExclamation + vbOKOnly
       Cancel = True
       Exit Sub
    End If
    If textSA01 <> "" And textSA02 <> "" _
         And Val(textSA03_1) > 0 And Val(textSA03_2) > 0 Then
      If IsRecordExist(textSA01, DBDATE(textSA02), Trim(textSA03_1.Text & textSA03_2.Text)) = True And textSA01.Enabled = True And textSA01.Locked = False Then
          MsgBox "該員工當天已有資料，請修改！", vbInformation
          Call textSA03_2_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
    ' 2008/12/18 END
   'Add By Sindy 2011/9/22
   'If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSA09 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   'End If
End If
CloseIme
End Sub

'Private Sub textSA04_Change()
'Dim tmpCalH As String
'If m_EditMode <> 0 Then
'    If Trim(textSA02) <> "" And Trim(textSA03_1) <> "" And Trim(textSA03_2) <> "" And Trim(textSA04) <> "" And Trim(textSA05_1) <> "" And Trim(textSA05_2) <> "" Then
'        If CheckIsTaiwanDate(textSA02, False) = True And CheckIsTaiwanDate(textSA04, False) = True Then
'            tmpCalH = CalDateTime(textSA02 & Format(textSA03_1, "00") & Format(textSA03_2, "00"), textSA04 & Format(textSA05_1, "00") & Format(textSA05_2, "00"))
'            textSA07 = Val(tmpCalH) \ 8
'            textSA08 = Val(tmpCalH) - (Val(textSA07) * 8)
'        Else
'            textSA07 = 0
'            textSA08 = 0
'        End If
'    Else
'        textSA07 = 0
'        textSA08 = 0
'    End If
'End If
'End Sub

Private Sub textSA04_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSA04
End If
End Sub

Private Sub textSA04_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSA04_Validate(Cancel As Boolean)
If textSA05_1 = "" Then textSA05_1 = "00"
If textSA05_2 = "" Then textSA05_2 = "00"

If m_EditMode <> 0 And textSA04 <> "" Then
    If CheckIsTaiwanDate(textSA04, False) = False Then
        Call textSA04_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
        Exit Sub
    End If
'    If ChkWorkDay(DBDATE(textSA04)) = False Then
'        Call textSA04_GotFocus   ' 2008/12/18 ADD BY SINDY
'        Cancel = True
'        MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
'        Exit Sub
'    End If
    If textSA02 <> "" And textSA04 <> "" Then
      If RunNick2(textSA02, textSA04) Then
          Call textSA04_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
'    'Add By Sindy 2011/9/22
'    If Frame1.Visible = True And textSA09 <> "" And m_B1019 <> "" Then '有表單編號
'      If Val(DBDATE(textSA04)) < Val(DBDATE(m_B1004)) Or Val(DBDATE(textSA04)) > Val(DBDATE(m_B1006)) Then
'         MsgBox "請假日期只能改少不能改多！", vbInformation, "輸入日期錯誤"
'         Call textSA04_GotFocus
'         Cancel = True
'         Exit Sub
'      End If
'    End If
   'Add By Sindy 2011/9/22
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSA09 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   End If
End If
End Sub

'Private Sub textSA05_1_Change()
'Dim tmpCalH As String
'If m_EditMode <> 0 Then
'    If Trim(textSA02) <> "" And Trim(textSA03_1) <> "" And Trim(textSA03_2) <> "" And Trim(textSA04) <> "" And Trim(textSA05_1) <> "" And Trim(textSA05_2) <> "" Then
'        If CheckIsTaiwanDate(textSA02, False) = True And CheckIsTaiwanDate(textSA04, False) = True Then
'            tmpCalH = CalDateTime(textSA02 & Format(textSA03_1, "00") & Format(textSA03_2, "00"), textSA04 & Format(textSA05_1, "00") & Format(textSA05_2, "00"))
'            textSA07 = Val(tmpCalH) \ 8
'            textSA08 = Val(tmpCalH) - (Val(textSA07) * 8)
'        Else
'            textSA07 = 0
'            textSA08 = 0
'        End If
'    Else
'        textSA07 = 0
'        textSA08 = 0
'    End If
'End If
'End Sub

Private Sub textSA05_1_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSA05_1
End If
End Sub

Private Sub textSA05_1_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSA05_1_Validate(Cancel As Boolean)
If textSA05_1 = "" Then textSA05_1 = "00"

If m_EditMode <> 0 And textSA05_1 <> "" Then
   If CheckLengthIsOK(textSA05_1, textSA05_1.MaxLength) = False Then
       Call textSA05_1_GotFocus   ' 2008/12/18 ADD BY SINDY
       Cancel = True
       Exit Sub
   End If
   ' 2008/12/19 Add BY SINDY
   If textSA05_1.Text > 24 Then
      Call textSA05_1_GotFocus
      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   ' 2008/12/19 END
   'Add By Sindy 2022/7/19
   If textSA01 <> "" And textSA02 <> "" _
      And Val(textSA05_1) > 0 And Val(textSA05_2) > 0 Then
      If IsRecordExist(textSA01, DBDATE(textSA02), Trim(textSA03_1.Text & textSA03_2.Text)) = True And textSA01.Enabled = True And textSA01.Locked = False Then
         MsgBox "該員工當天已有資料，請修改！", vbInformation
         Call textSA05_1_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   '2022/7/19 END
   'Add By Sindy 2011/9/22
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSA09 <> "" Then
         '計算時數
         Call CountDayHour
      End If
   End If
End If
CloseIme
End Sub

'Private Sub textSA05_2_Change()
'Dim tmpCalH As String
'If m_EditMode <> 0 Then
'    If Trim(textSA02) <> "" And Trim(textSA03_1) <> "" And Trim(textSA03_2) <> "" And Trim(textSA04) <> "" And Trim(textSA05_1) <> "" And Trim(textSA05_2) <> "" Then
'        If CheckIsTaiwanDate(textSA02, False) = True And CheckIsTaiwanDate(textSA04, False) = True Then
'            tmpCalH = CalDateTime(textSA02 & Format(textSA03_1, "00") & Format(textSA03_2, "00"), textSA04 & Format(textSA05_1, "00") & Format(textSA05_2, "00"))
'            textSA07 = Val(tmpCalH) \ 8
'            textSA08 = Val(tmpCalH) - (Val(textSA07) * 8)
'        Else
'            textSA07 = 0
'            textSA08 = 0
'        End If
'    Else
'        textSA07 = 0
'        textSA08 = 0
'    End If
'End If
'End Sub

Private Sub textSA05_2_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSA05_2
End If
End Sub

Private Sub textSA05_2_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSA05_2_LostFocus()
If m_EditMode <> 0 And textSA05_2 <> "" Then
    If Trim(textSA02) <> "" And Trim(textSA03_1) <> "" And Trim(textSA03_2) <> "" And Trim(textSA04) <> "" And Trim(textSA05_1) <> "" And Trim(textSA05_2) <> "" Then
        If CheckIsTaiwanDate(textSA02, False) = True And CheckIsTaiwanDate(textSA04, False) = True Then
            If CompDateTime(textSA02 & Format(textSA03_1, "00") & Format(textSA03_2, "00"), textSA04 & Format(textSA05_1, "00") & Format(textSA05_2, "00")) = False Then
                Call textSA05_2_GotFocus   ' 2008/12/18 ADD BY SINDY
                MsgBox "日期時間設定錯誤！", vbInformation, "輸入錯誤！"
                'textSA04.SetFocus
                Exit Sub
            End If
        End If
    End If
End If
End Sub

Private Sub textSA05_2_Validate(Cancel As Boolean)
If textSA05_2 = "" Then textSA05_2 = "00"

If m_EditMode <> 0 And textSA05_2 <> "" Then
   If CheckLengthIsOK(textSA05_2, textSA05_2.MaxLength) = False Then
       Call textSA05_2_GotFocus   ' 2008/12/18 ADD BY SINDY
       Cancel = True
       Exit Sub
   End If
   ' 2008/12/18 ADD BY SINDY
   If textSA05_2.Text > 59 Then
      Call textSA05_2_GotFocus
      MsgBox "不可超過59分!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   ' 2008/12/18 END
   'Add By Sindy 2022/7/19
   If textSA01 <> "" And textSA02 <> "" _
      And Val(textSA05_1) > 0 And Val(textSA05_2) > 0 Then
      If IsRecordExist(textSA01, DBDATE(textSA02), Trim(textSA03_1.Text & textSA03_2.Text)) = True And textSA01.Enabled = True And textSA01.Locked = False Then
         MsgBox "該員工當天已有資料，請修改！", vbInformation
         Call textSA05_2_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   '2022/7/19 END
   'Modify By Sindy 2011/9/22
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSA09 <> "" Then
         '計算時數
         Call CountDayHour
      Else
         '可以人工修改
         If textSA07 = "" Or textSA08 = "" Or (textSA07 = "0" And textSA08 = "0") Then
           '計算時數
           Call CountDayHour
         End If
      End If
   End If
End If
CloseIme
End Sub

Private Sub textSA06_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSA06
End If
End Sub

Private Sub textSA06_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSA06_Validate(Cancel As Boolean)
Dim MyRs As New ADODB.Recordset
Dim MyArr As Variant

If textSA06.Text <> "" Then
    MyArr = Split(textSA06, " ")
    Set MyRs = New ADODB.Recordset
    If MyRs.State = 1 Then MyRs.Close
    ' 2008/12/18 Modify BY SINDY
    ' 排除不須要的代碼 : 01.忘打卡 02.遲到 03.曠職 04.出差 16.加班
    'strSQL = "select ac02||' '||ac03 from allcode where ac01='04' and ac02='" & MyArr(0) & "' order by ac02"
    strSql = "select ac02||' '||ac03 from allcode where ac01='04' and ac02='" & MyArr(0) & "' and ac02 not in ('01','02','03','04','16') order by ac02"
    ' 2008/12/18 END
    MyRs.CursorLocation = adUseClient
    MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If MyRs.RecordCount <> 0 Then
         textSA06.Text = "" & MyRs.Fields(0).Value
    Else
        If m_EditMode <> 0 Then
            Call textSA06_GotFocus   ' 2008/12/18 ADD BY SINDY
            MsgBox "假別代號輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub textSA07_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSA07
End If
End Sub

Private Sub textSA07_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSA07_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textSA07 <> "" Then
    If CheckLengthIsOK(textSA07, textSA07.MaxLength) = False Then
        Call textSA07_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textSA08_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSA08
End If
End Sub

Private Sub textSA08_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub textSA08_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textSA08 <> "" Then
    If CheckLengthIsOK(textSA08, textSA08.MaxLength) = False Then
        Call textSA08_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    
    ' 2008/12/17 ADD BY SINDY
    'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
    'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
    'Modify By Sindy 2012/7/9 上班時數為特殊者
'    If textSA01 = "99029" Then
'      If textSA08.Text >= 5 Then
'         Call textSA08_GotFocus
'         MsgBox "請假時數-共(時)不可超過5小時!!!", vbExclamation + vbOKOnly
'         Cancel = True
'         Exit Sub
'      End If
    '2010/7/14 End
'    Call Pub_GetSpecWorkHour(textSA01, textSA02)
'    If Val(textSA08.Text) >= Val(PUB_intWkHour) Then
'       Call textSA08_GotFocus
'       MsgBox "請假時數-共(時)不可超過" & PUB_intWkHour & "小時!!!", vbExclamation + vbOKOnly
'       Cancel = True
'       Exit Sub
'    End If
    ' 2008/12/17 END
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
   'Modify By Sindy 2012/7/9
   Dim strSTime As String, strETime As String
   Dim strSA07 As String, strSA08 As String
   strSTime = "": strETime = ""
   If Frame1.Visible = True Then strSTime = Format(cboSTime.Text, "hhmm")
   If Frame1.Visible = True Then strETime = Format(cboETime.Text, "hhmm")
   strSA07 = textSA07
   strSA08 = textSA08
   'Modify by Sindy 2012/10/15
   'Call PUB_CountDayHour(textSA01, DBDATE(textSA02), Format(textSA03_1, "00") & Format(textSA03_2, "00"), DBDATE(textSA04), Format(textSA05_1, "00") & Format(textSA05_2, "00"), strSTime, strETime, strSA07, strSA08, True, False)
   Call PUB_CountDayHour(textSA01, DBDATE(textSA02), Format(textSA03_1, "00") & Format(textSA03_2, "00"), DBDATE(textSA04), Format(textSA05_1, "00") & Format(textSA05_2, "00"), strSTime, strETime, strSA07, strSA08, textSA06, True, False)
   '2012/10/15 End
   textSA07 = strSA07
   textSA08 = strSA08
   '2012/7/9 End
   
'Dim tmpCalH As String
''Dim dblSTime As Double, dblETime As Double
'Dim temp As Variant ', bwk5hour As Boolean
'Dim strSTime As String, strETime As String 'Add By Sindy 2011/9/22
'
'   'Add By Sindy 2010/7/14 99029伊恩一天只上4個小時
'   'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
''   bwk5hour = False
''   If textSA01 = "99029" Then bwk5hour = True
'   'Modify By Sindy 2012/7/9 上班時數為特殊者
'   Call Pub_GetSpecWorkHour(textSA01)
'   '2010/7/14 End
'
'   'Modify By Sindy 2011/9/22 電子表單資料,日及時欄位值一律由系統計算
'   If textSA07 = "" Or textSA08 = "" Or (textSA07 = "0" And textSA08 = "0") Then 'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'   'If textSA07 = "" Or textSA08 = "" Or (textSA07 = "0" And textSA08 = "0") Or (Frame1.Visible = True And textSA09 <> "") Then
'      If Trim(textSA02) <> "" And Trim(textSA03_1) <> "" And Trim(textSA03_2) <> "" And Trim(textSA04) <> "" And Trim(textSA05_1) <> "" And Trim(textSA05_2) <> "" Then
'          If CheckIsTaiwanDate(textSA02, False) = True And CheckIsTaiwanDate(textSA04, False) = True Then
'              'Modify By Sindy 2010/7/14 增加傳入bwk4hour
'              'Modify By Sindy 2011/3/8 增加傳入bwk5hour
'              'Add By Sindy 2011/9/22
'              strSTime = "": strETime = ""
'              If cboSTime.Visible = True Then strSTime = Format(cboSTime.Text, "hhmm")
'              If cboETime.Visible = True Then strETime = Format(cboETime.Text, "hhmm")
'              tmpCalH = CalDateTime(textSA02 & Format(textSA03_1, "00") & Format(textSA03_2, "00"), textSA04 & Format(textSA05_1, "00") & Format(textSA05_2, "00"), PUB_bWkSpec, strSTime, strETime)
'
''              'Add By Sindy 98/03/13 起始時間<=12時並且迄止時間>=13時30分者，減1小時
''              dblSTime = Val(textSA03_1 & textSA03_2)
''              dblETime = Val(textSA05_1 & textSA05_2)
''              If dblSTime <= 1200 And dblETime >= 1330 Then
''                  tmpCalH = tmpCalH - 1
''              End If
''              '98/03/13 End
'
'              If tmpCalH > "" Then
'                  'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
'                  'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
'                  'Modify By Sindy 2012/7/9 上班時數為特殊者
''                  If textSA01 = "99029" Then
''                      If tmpCalH < 5 Then
''                          textSA07 = 0
''                      Else
''                          temp = Split(CStr(Val(tmpCalH) / 5), ".")
''                          textSA07 = temp(0)
''                      End If
''                      textSA08 = Val(tmpCalH) - (Val(textSA07) * 5)
'                  '2010/7/14 End
'                  If PUB_bWkSpec = True Then
'                      If Val(tmpCalH) < Val(PUB_intWkHour) Then
'                          textSA07 = 0
'                      Else
'                          temp = Split(CStr(Val(tmpCalH) / PUB_intWkHour), ".")
'                          textSA07 = temp(0)
'                      End If
'                      textSA08 = Val(tmpCalH) - (Val(textSA07) * PUB_intWkHour)
'                  Else
'                  '2012/7/9 End
'                      If tmpCalH < 8 Then
'                          textSA07 = 0
'                      Else
'                          temp = Split(CStr(Val(tmpCalH) / 8), ".")
'                          textSA07 = temp(0)
'                      End If
'                      textSA08 = Val(tmpCalH) - (Val(textSA07) * 8)
'                  End If
'              Else
'                  textSA07 = ""
'                  textSA08 = ""
'                  MsgBox "日期時間設定錯誤！", vbInformation, "輸入錯誤！"
'                  Exit Function
'              End If
'          Else
'              textSA07 = ""
'              textSA08 = ""
'          End If
'      Else
'          textSA07 = ""
'          textSA08 = ""
'      End If
'   End If
End Function

Private Sub textSA09_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSA09
    CloseIme
End If
End Sub

Private Sub textSA09_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSA09_LostFocus()
   '新增狀態時可以輸入表單編號做查詢
   If m_EditMode = 1 And textSA09 <> "" Then
      If GetABS010 = True Then
         textSA09.Enabled = False
      End If
   End If
End Sub

Private Sub textSA09_Validate(Cancel As Boolean)
   If Frame1.Visible = False Then Exit Sub
   
   If m_EditMode = 1 And textSA09 <> "" Then
      If CheckLengthIsOK(textSA09, textSA09.MaxLength) = False Then
         Call textSA09_GotFocus
         Cancel = True
         Exit Sub
      End If
      If ChkAbsSysB1001Exist(textSA09, "01", textSA01) = False Then
         Call textSA09_GotFocus
         Cancel = True
         Exit Sub
      End If
      If ChkPerSysB1001Exist(textSA09, textSA01, False) = True Then
         MsgBox "表單編號重覆！", vbExclamation
         Call textSA09_GotFocus
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
   If Val(Format(cboSTime.Text, "hhmm")) > Val(Right("00" & textSA03_1, 2) & Right("00" & textSA03_2, 2)) Then
      Call cboSTime_GotFocus
      MsgBox "起日上班時段必須小於或等於起日請假時間!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
'   If Val(Format(cboSTime.Text, "hhmm")) > 2400 Then
'      Call cboSTime_GotFocus
'      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
'      Cancel = True
'      Exit Sub
'   End If
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSA09 <> "" Then
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
   If Val(Format(cboETime.Text, "hhmm")) < Val(Right("00" & textSA05_1, 2) & Right("00" & textSA05_2, 2)) Then
      Call cboETime_GotFocus
      MsgBox "迄日下班時段必須大於或等於迄日請假時間!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
'   If Val(Format(cboETime.Text, "hhmm")) > 2400 Then
'      Call cboETime_GotFocus
'      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
'      Cancel = True
'      Exit Sub
'   End If
   If m_EditMode = 1 Or m_EditMode = 2 Then
      '電子表單資料,日及時欄位值一律由系統計算
      If Frame1.Visible = True And textSA09 <> "" Then
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
            "and B1001='" & textSA09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GetABS010 = True
      'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'      '有表單編號的資料時,日及時欄位鎖住
'      textSA07.Enabled = False
'      textSA08.Enabled = False

      '記錄原始資料 : 註.m_變數值必須在ClearField函數裡清值
      If Not IsNull(rsTmp.Fields("B1019")) Then m_B1019 = rsTmp.Fields("B1019")
      If Not IsNull(rsTmp.Fields("B1004")) Then m_B1004 = rsTmp.Fields("B1004")
      If Not IsNull(rsTmp.Fields("B1005")) Then m_B1005 = IIf(Format(rsTmp.Fields("B1005"), "hhmm") = "0000", "", Format(rsTmp.Fields("B1005"), "hhmm"))
      If Not IsNull(rsTmp.Fields("B1006")) Then m_B1006 = rsTmp.Fields("B1006")
      If Not IsNull(rsTmp.Fields("B1007")) Then m_B1007 = IIf(Format(rsTmp.Fields("B1007"), "hhmm") = "0000", "", Format(rsTmp.Fields("B1007"), "hhmm"))
      If Not IsNull(rsTmp.Fields("B1008")) Then m_B1008 = Trim(rsTmp.Fields("B1008"))
      If Not IsNull(rsTmp.Fields("B1009")) Then m_B1009 = rsTmp.Fields("B1009")
      If Not IsNull(rsTmp.Fields("B1010")) Then m_B1010 = rsTmp.Fields("B1010")
      If Not IsNull(rsTmp.Fields("B1017")) Then m_B1017 = rsTmp.Fields("B1017")
      If Not IsNull(rsTmp.Fields("B1028")) Then m_B1028 = IIf(Format(rsTmp.Fields("B1028"), "hhmm") = "0000", "", Format(rsTmp.Fields("B1028"), "hhmm"))
      If Not IsNull(rsTmp.Fields("B1029")) Then m_B1029 = IIf(Format(rsTmp.Fields("B1029"), "hhmm") = "0000", "", Format(rsTmp.Fields("B1029"), "hhmm"))
      
      '顯示其他資料至畫面上
      'Add By Sindy 2022/10/28 + if 已簽核不要顯示於畫面上,已人事資料為主
      If Not IsNull(rsTmp.Fields("B1019")) Then
         '為防止簽核後又修改,抓人事資料
         strSql = "select * from abs012 where b1201='" & textSA09 & "' and substr(b1207,1,4)='修改資料'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            cmdABS.Visible = True
            '記錄畫面上的資料 : 註.m_變數值必須在ClearField函數裡清值
            m_B1004 = DBDATE(textSA02)
            m_B1005 = Format(textSA03_1 & textSA03_2, "0000")
            m_B1006 = DBDATE(textSA04)
            m_B1007 = Format(textSA05_1 & textSA05_2, "0000")
            m_B1008 = textSA06
            m_B1009 = textSA07
            m_B1010 = textSA08
         End If
      Else
      '2022/10/28 END
         If Not IsNull(rsTmp.Fields("B1004")) Then textSA02 = ChangeWStringToTString(rsTmp.Fields("B1004"))
         If Not IsNull(rsTmp.Fields("B1005")) Then textSA03_1 = Left(rsTmp.Fields("B1005"), 2): textSA03_2 = Right(rsTmp.Fields("B1005"), 2)
         If Not IsNull(rsTmp.Fields("B1006")) Then textSA04 = ChangeWStringToTString(rsTmp.Fields("B1006"))
         If Not IsNull(rsTmp.Fields("B1007")) Then textSA05_1 = Left(rsTmp.Fields("B1007"), 2): textSA05_2 = Right(rsTmp.Fields("B1007"), 2)
         If Not IsNull(rsTmp.Fields("B1008")) Then textSA06 = Trim(rsTmp.Fields("B1008"))
         If Not IsNull(rsTmp.Fields("B1009")) Then textSA07 = rsTmp.Fields("B1009")
         If Not IsNull(rsTmp.Fields("B1010")) Then textSA08 = rsTmp.Fields("B1010")
      End If
      
      'Add By Sindy 2021/8/11
      If bolOnlyQrySETime = False Then
         SetB102829Combo cboSTime, 1, textSA02, textSA01
         SetB102829Combo cboETime, 2, textSA02, textSA01
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
Dim strB1007 As String, strB1008 As String, strB1009 As String, strB1010 As String
Dim strB1028 As String, strB1029 As String
Dim strOldData As String, strNowData As String, strNote As String
'Dim strTo As String 'Add By Sindy 2012/7/17
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2016/10/26
Dim strB1009_2 As String 'Add By Sindy 2016/10/27
Dim strSubject As String, strContent As String 'Add By Sindy 2019/5/24
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   '檢查有無異動資料:
   '畫面上的欄位值
   strB1004 = DBDATE(textSA02)
   strB1005 = textSA03_1 & Format("00" & textSA03_2, "00")
   strB1006 = DBDATE(textSA04)
   strB1007 = textSA05_1 & Format("00" & textSA05_2, "00")
   strB1008 = Left(Trim(textSA06), 2)
   strB1009 = textSA07
   strB1010 = textSA08
   'Add By Sindy 2013/4/1
   If textSA08 = 0 Or textSA08 = "" Then
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
   strOldData = strOldData & "," & GetAllCode04(Left(m_B1008, 2))
   strOldData = strOldData & "," & ChangeWStringToTDateString(m_B1004) & "," & Format(m_B1005, "##:##")
   strOldData = strOldData & "," & ChangeWStringToTDateString(m_B1006) & "," & Format(m_B1007, "##:##")
   strOldData = strOldData & "," & m_B1009 & "日," & m_B1010 & "時"
   '串目前畫面上資料
   If Frame1.Visible = True And _
      cboSTime.Text <> "" And cboETime.Text <> "" Then
      strNowData = strNowData & "非整日," & Format(strB1028, "##:##") & "," & Format(strB1029, "##:##")
   End If
   strNowData = strNowData & "," & GetAllCode04(Left(strB1008, 2))
   strNowData = strNowData & "," & ChangeWStringToTDateString(strB1004) & "," & Format(strB1005, "##:##")
   strNowData = strNowData & "," & ChangeWStringToTDateString(strB1006) & "," & Format(strB1007, "##:##")
   strNowData = strNowData & "," & strB1009 & "日," & strB1010 & "時"
   If Left(strOldData, 1) = "," Then strOldData = Right(strOldData, Len(strOldData) - 1)
   If Left(strNowData, 1) = "," Then strNowData = Right(strNowData, Len(strNowData) - 1)
   
   'Added by Lydia 2020/03/05 記錄email內容
   strCallCont = ""
   If strOldData <> strNowData Then
       If bolIsDel = True Then
           strCallCont = "已註銷：" & strNowData
       Else
           strCallCont = "異動前資料：" & strOldData & vbCrLf & _
                               "異動後資料：" & strNowData
       End If
   End If
   'end 2020/03/05
   
   '流程備註檔
   If txtNote.Text <> "" And textSA09 <> "" Then
      strSql = GetInsertABS012Sql(Trim(textSA09), 人事處, strUpdDate, strUpdTime, "", txtNote)
      cnnConnection.Execute strSql
   End If
   
   If strOldData <> strNowData And textSA09 <> "" Then '電子簽核的,非紙本
      '人事處尚未簽收時,在人事系統已先建立此表單編號資料,須一併更新出缺勤電子簽核主檔資料
      If m_B1019 = "" Then
         strSql = "update ABS010 set " & _
                  "B1004= " & CNULL(DBDATE(strB1004)) & _
                  ",B1005= " & CNULL(strB1005) & _
                  ",B1006= " & CNULL(strB1006) & _
                  ",B1007= " & CNULL(strB1007) & _
                  ",B1008= " & CNULL(strB1008) & _
                  ",B1009= " & CNULL(strB1009) & _
                  ",B1010= " & CNULL(strB1010) & _
                  ",B1028= " & CNULL(strB1028) & _
                  ",B1029= " & CNULL(strB1029) & _
                  " where B1001=" & CNULL(textSA09)
         cnnConnection.Execute strSql
      End If
      '檢查有異動資料時,須記錄異動資訊到表單流程備註
      strNote = "修改資料" & strOldData & "->" & strNowData
      strSql = GetInsertABS012Sql(Trim(textSA09), "M21", strUpdDate, strUpdTime, "", strNote)
      cnnConnection.Execute strSql
   End If
   
   If m_B1019 = "" And m_EditMode = 1 And textSA09 <> "" Then
      '寄E-Mail通知當事人
      PUB_SendMail strUserNum, textSA01, "", "表單人事處已先行作業，請儘速簽核。", _
      "表單內容為，" & strNowData & vbCrLf & _
      "(表單編號：" & textSA09 & ")", , , , , , , , , , True
   ElseIf m_B1019 <> "" And m_EditMode = 1 And textSA09 <> "" Then
      strSql = "update ABS010 set " & _
               "B1018='" & 已核准 & "'" & _
               " where B1001=" & CNULL(textSA09)
      cnnConnection.Execute strSql
      
      '記錄資訊到表單流程備註
      strNote = "補入資料"
      strSql = GetInsertABS012Sql(Trim(textSA09), "M21", strUpdDate, strUpdTime, "", strNote)
      cnnConnection.Execute strSql
   Else
      If strOldData <> strNowData Then
'         '寄E-Mail通知當事人有異動內容
'         'Modify By Sindy 2012/7/17 發E-Mail通知當事人之外，已簽核的職代及審核主管亦也要通知
'         strTo = GetBossB1107_All(textSA09)
'         'Add By Sindy 2012/7/17 專利處P10-P14,必須另外E-Mail通知71011王副總
'         If (GetStaffDepartment(textSA01) >= "P10" And GetStaffDepartment(textSA01) <= "P14") And _
'            InStr(strTo, "71011") = 0 Then
'            strTo = strTo + ";71011"
'         End If
         
         'Add By Sindy 2016/10/26 若有跨月的資料,重組異動後資料
         If textSA09 <> "" Then '電子簽核的,非紙本
            strSql = "select * from staff_absence where sa01='" & textSA01 & "' and SA09='" & textSA09 & "' order by sa02 asc"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 1 Then '有一筆以上的資料(跨月) ex:97031林若璇/表單編號：10506085
               strNowData = ""
               rsTmp.MoveFirst
               strB1004 = rsTmp.Fields("SA02") '起始日期
               strB1005 = rsTmp.Fields("SA03") '起始時間
               strB1008 = rsTmp.Fields("SA06") '假別
               strB1009 = "0": strB1010 = "0": strB1009_2 = "0"
               Do While Not rsTmp.EOF
                  strB1009 = Int(strB1009) + Int(rsTmp.Fields("SA07"))
                  strB1010 = Val(strB1010) + Val(rsTmp.Fields("SA08"))
                  strB1006 = rsTmp.Fields("SA04") '截止日期
                  strB1007 = rsTmp.Fields("SA05") '截止時間
                  rsTmp.MoveNext
               Loop
               Dim strTemp As Variant
               strTemp = Split(CStr(Val(strB1010) / PUB_intWkHour), ".")
               strB1009_2 = CStr(strTemp(0))
               strB1009 = Int(strB1009) + Int(strB1009_2)
               strB1010 = Val(strB1010) - (Val(strB1009_2) * PUB_intWkHour)
               If Frame1.Visible = True And _
                  cboSTime.Text <> "" And cboETime.Text <> "" Then
                  strNowData = strNowData & "非整日," & Format(strB1028, "##:##") & "," & Format(strB1029, "##:##")
               End If
               strNowData = strNowData & "," & GetAllCode04(strB1008)
               strNowData = strNowData & "," & ChangeWStringToTDateString(strB1004) & "," & Format(strB1005, "##:##")
               strNowData = strNowData & "," & ChangeWStringToTDateString(strB1006) & "," & Format(strB1007, "##:##")
               strNowData = strNowData & "," & strB1009 & "日," & strB1010 & "時"
               If Left(strNowData, 1) = "," Then strNowData = Right(strNowData, Len(strNowData) - 1)
            End If
            rsTmp.Close
         End If
         '2016/10/26 END
         
         If textSA09 <> "" Then '電子簽核的,非紙本
            strSubject = "[通知]人事處有修改資料(表單編號：" & textSA09 & ")"
         Else
            strSubject = "[通知]人事處有修改資料"
         End If
         'Add By Sindy 2013/2/1
         If bolIsDel = True Then
'            PUB_SendMail strUserNum, textSA01, "", "[通知]人事處有修改資料(表單編號：" & textSA09 & ")", _
'            "異動前資料：" & strOldData & vbCrLf & _
'            "註銷的資料：" & strNowData & vbCrLf & _
'            "(表單編號：" & textSA09 & ")" & vbCrLf & vbCrLf & _
'            "人事處修改原因：" & txtNote, , , , , , strTo, , , , True
            strContent = "異動前資料：" & strOldData & vbCrLf & _
                         "註銷的資料：" & strNowData & vbCrLf & _
                         IIf(textSA09 <> "", "(表單編號：" & textSA09 & ")" & vbCrLf & vbCrLf, "") & _
                         "人事處修改原因：" & txtNote
         Else
         '2013/2/1 End
'            PUB_SendMail strUserNum, textSA01, "", "[通知]人事處有修改資料(表單編號：" & textSA09 & ")", _
'            "異動前資料：" & strOldData & vbCrLf & _
'            "異動後資料：" & strNowData & vbCrLf & _
'            "(表單編號：" & textSA09 & ")" & vbCrLf & vbCrLf & _
'            "人事處修改原因：" & txtNote, , , , , , strTo, , , , True
            strContent = "異動前資料：" & strOldData & vbCrLf & _
                         "異動後資料：" & strNowData & vbCrLf & _
                         IIf(textSA09 <> "", "(表單編號：" & textSA09 & ")" & vbCrLf & vbCrLf, "") & _
                         "人事處修改原因：" & txtNote
         End If
         '2012/7/17 End
         'Add By Sindy 2019/5/23 假單完成,後續資料檢查及SendMail
         Call PUB_AutoM21Receive_SendMail(IIf(textSA09 <> "", textSA09, ""), 表單類別_請假, textSA01, DBDATE(textSA02), Trim(Format("00" & textSA03_1, "00") & Format("00" & textSA03_2, "00")), _
            DBDATE(textSA04), "", Left(DBDATE(textSA02), 6), , strSubject, strContent, m_EditMode)
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
   
   If textSA09 <> "" Then m_B1018 = 註銷 '(06)
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   'Modify By Sindy 2019/5/24
   'strContent = GetEMailContent(textSA09, strSubject)
   strContent = GetEMailContent(IIf(textSA09 <> "", textSA09, ""), strSubject, m_B1018, , , "01", textSA01, DBDATE(textSA02), Val(Trim(Format("00" & textSA03_1, "00") & Format("00" & textSA03_2, "00"))), m_EditMode)
   strContent = strContent & vbCrLf & vbCrLf & _
                "人事處修改原因：" & txtNote
   
'   cnnConnection.BeginTrans
   
   If textSA09 <> "" Then
      '流程備註檔
      If txtNote.Text <> "" Then
         strSql = GetInsertABS012Sql(Trim(textSA09), 人事處, strUpdDate, strUpdTime, m_B1018, txtNote)
         cnnConnection.Execute strSql
      End If
      '主檔
      strSql = "update ABS010 set " & _
               "B1018='" & m_B1018 & "'" & _
               " where B1001='" & textSA09 & "' "
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
'   strTo = GetBossB1107_All(textSA09)
'   'Add By Sindy 2012/8/23 專利處P10-P14,必須另外E-Mail通知71011王副總
'   If (GetStaffDepartment(textSA01) >= "P10" And GetStaffDepartment(textSA01) <= "P14") And _
'      InStr(strTo, "71011") = 0 Then
'      strTo = strTo + ";71011"
'   End If
'   '2012/8/23 End
   
   'PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True
'   PUB_SendMail strUserNum, textSA01, "", strSubject, strContent & vbCrLf & vbCrLf & _
'         "人事處修改原因：" & txtNote, , , , , , strTo, , , , True
   
   'Add By Sindy 2019/5/23 假單完成,後續資料檢查及SendMail
   Call PUB_AutoM21Receive_SendMail(IIf(textSA09 <> "", textSA09, ""), 表單類別_請假, textSA01, DBDATE(textSA02), Trim(Format("00" & textSA03_1, "00") & Format("00" & textSA03_2, "00")), _
      DBDATE(textSA04), "", Left(DBDATE(textSA02), 6), , strSubject, strContent, m_EditMode)
      
'   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "註銷失敗！" & vbCrLf & Err.Description
End Sub

'Added by Lydia 2020/03/05 若幫查名人員銷假上班(請假資料整筆刪除或修改日期)，系統將發Email通知「內商查名銷假通知」，提醒至系統更改查名人狀態。
'bolAll: True=無視查名人狀態(tmqsr17), False:查名人狀態=不分單(tmqsr17=N)
Private Sub ProcTMQemail(ByVal pType As String, Optional ByVal bolAll = False)
'參考frmAutoBatchDay.StrMenu62
'2019/01/28 查名單分發規則調整如下：
    '1.特休(請)假整日者，前一天不發查名單，
    '2.特休(請)假非整日者，如僅半日或不及半日（小時）之非整日特休（請）假，前一天仍正常發給查名單
    '3.特休(請)假整日者，特休(請)假當日不發查名單
    '4.特休(請)假非整日者，以查名單分發當時段是否上班為原則發給查名單，例：上午請半日特休假，下午上班者，則下午正常發給查名單
    '程式修改第2點和第4點,非整日改由當日判斷是否出缺勤
Dim intQ As Integer
Dim strQ1 As String, strA1 As String
Dim rsQuery As New ADODB.Recordset
Dim strContent As String

On Error GoTo ErrHandle

    strQ1 = "Select Tmqm01,Tmqm02,decode(tmqm01,tmqm01,'N',tmqm03) tmqm03,tmqsr17 From Tmqmember,Staff,tmqsumr " & _
                "Where Tmqm01='" & textSA01 & "' And Tmqm01=St01(+) And St04='1' and tmqm01=tmqsr01(+) "
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
    If intQ = 1 Then
         If (bolAll = False And "" & rsQuery.Fields("tmqsr17") = "N") Or bolAll = True Then 'bolAll=False只提醒不分單改為要分單的情況
            strA1 = "" & rsQuery.Fields("tmqm03")
            If strA1 = "Y" Then
               'TMQM03=Y，表示統計人員(ex.P2002, P2003)底下的查名人有一位請假(排除假單的人員)，這個代號今天請假
               strQ1 = "select count(*) cnt from tmqmember,tmqsumr where tmqm01<>'" & rsQuery.Fields("tmqm01") & "' " & _
                           " and tmqm02='" & rsQuery.Fields("tmqm02") & "' and tmqm01<>tmqm02 and tmqm01=tmqsr01 and tmqsr17='N' "
               intQ = 1
               Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
               If intQ = 0 Then
                  strA1 = "N"
               Else
                  If Val("" & rsQuery.Fields("cnt")) = 0 Then
                      strA1 = "N"
                  End If
               End If
            End If
            If strA1 = "N" Then
                strA1 = Pub_GetSpecMan("內商查名銷假通知")
                strContent = "查名人員(" & textSA01 & textSA01_2 & ")請假" & IIf(pType = "M", "時間異動", "已註銷") & "，" & vbCrLf
                If strCallCont <> "" Then
                    strContent = strContent & strCallCont & vbCrLf & vbCrLf
                Else
                    strContent = strContent & IIf(pType = "M", "異動後資料：", "已註銷：") & textSA06.Text & "," & ChangeTStringToTDateString(textSA02) & "," & Format(textSA03_1 & textSA03_2, "##:##") & _
                                                                        "," & ChangeTStringToTDateString(textSA04) & "," & Format(textSA05_1 & textSA05_2, "##:##") & _
                                                                        "," & textSA07 & "日," & textSA08 & "時" & vbCrLf & vbCrLf
                End If
                strContent = strContent & "請判斷是否要更改查名人狀態，" & vbCrLf & _
                                "程式路徑：承辦人系統-商標處-查名人狀態。"
                strQ1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                   " values( '" & strUserNum & "','" & strA1 & "',to_char(sysdate,'yyyymmdd')" & _
                   ",to_char(sysdate,'hh24miss'),'查名人員(" & textSA01 & textSA01_2 & ")請假" & IIf(pType = "M", "時間異動", "已註銷") & "，請至承辦人系統-商標處-查名人狀態進行更改。' " & _
                   ",'" & strContent & "')"
                cnnConnection.Execute strQ1, intQ
            End If
         End If
    End If
    
    Set rsQuery = Nothing
    Exit Sub
    
ErrHandle:
    If Err.Number <> 0 Then
        cnnConnection.RollbackTrans
    End If
End Sub
