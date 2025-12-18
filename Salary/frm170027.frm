VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170027 
   BorderStyle     =   1  '單線固定
   Caption         =   "其他各類所得資料"
   ClientHeight    =   5328
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8220
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5328
   ScaleWidth      =   8220
   Begin TabDlg.SSTab SSTab1 
      Height          =   4620
      Left            =   50
      TabIndex        =   10
      Top             =   636
      Width           =   8112
      _ExtentX        =   14330
      _ExtentY        =   8149
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm170027.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDsp(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDsp(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textCUID"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(9)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(11)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtOID(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtOID(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtOID(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtOID(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtOID(8)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtOID(9)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtOID(7)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtOID(5)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtOID(6)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm170027.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(4)"
      Tab(1).Control(1)=   "txtSum(0)"
      Tab(1).Control(2)=   "txtSum(1)"
      Tab(1).Control(3)=   "GRD1"
      Tab(1).Control(4)=   "txt1(3)"
      Tab(1).Control(5)=   "txt1(2)"
      Tab(1).Control(6)=   "txt1(0)"
      Tab(1).Control(7)=   "txt1(1)"
      Tab(1).Control(8)=   "cmdok"
      Tab(1).Control(9)=   "Label1(12)"
      Tab(1).Control(10)=   "Label5"
      Tab(1).Control(11)=   "Label6"
      Tab(1).Control(12)=   "Label8"
      Tab(1).Control(13)=   "Label16"
      Tab(1).Control(14)=   "Label12"
      Tab(1).ControlCount=   15
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   4
         Left            =   -73920
         MaxLength       =   1
         TabIndex        =   37
         Top             =   720
         Width           =   405
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   0
         Left            =   -72795
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   4296
         Width           =   1005
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   1
         Left            =   -70500
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   4296
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm170027.frx":0038
         Height          =   3072
         Left            =   -75000
         TabIndex        =   30
         Top             =   1080
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   5440
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "年度|月份|所得人代號|名　稱|公司別|格式|共用欄位|所得總額|扣繳稅額"
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
      Begin VB.TextBox txtOID 
         Height          =   270
         Index           =   6
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "12"
         Top             =   2120
         Width           =   405
      End
      Begin VB.TextBox txtOID 
         Height          =   270
         Index           =   5
         Left            =   1530
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "01"
         Top             =   2120
         Width           =   405
      End
      Begin VB.TextBox txtOID 
         Height          =   270
         Index           =   7
         Left            =   1530
         MaxLength       =   12
         TabIndex        =   6
         Text            =   "999999999999"
         Top             =   2440
         Width           =   1425
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -69810
         MaxLength       =   12
         TabIndex        =   16
         Top             =   434
         Width           =   1425
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -71400
         MaxLength       =   12
         TabIndex        =   15
         Top             =   434
         Width           =   1425
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73920
         MaxLength       =   3
         TabIndex        =   13
         Top             =   434
         Width           =   500
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -73140
         MaxLength       =   3
         TabIndex        =   14
         Top             =   434
         Width           =   500
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢(&S)"
         Height          =   345
         Left            =   -68190
         TabIndex        =   17
         Top             =   380
         Width           =   1095
      End
      Begin VB.TextBox txtOID 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   9
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "9999999"
         Top             =   3080
         Width           =   1200
      End
      Begin VB.TextBox txtOID 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   8
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "9999999"
         Top             =   2760
         Width           =   1200
      End
      Begin VB.TextBox txtOID 
         Height          =   270
         Index           =   3
         Left            =   1530
         MaxLength       =   1
         TabIndex        =   2
         Text            =   "1"
         Top             =   1480
         Width           =   405
      End
      Begin VB.TextBox txtOID 
         Height          =   270
         Index           =   1
         Left            =   1530
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "96"
         Top             =   520
         Width           =   500
      End
      Begin VB.TextBox txtOID 
         Height          =   270
         Index           =   2
         Left            =   1530
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "999999999999"
         Top             =   840
         Width           =   1425
      End
      Begin VB.TextBox txtOID 
         Height          =   270
         Index           =   4
         Left            =   1530
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "50"
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "格式 9B 講演所得：費用別代號"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   11
         Left            =   4392
         TabIndex        =   41
         Top             =   3000
         Width           =   2448
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "格式 9A 執行業務所得：業別代號"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   10
         Left            =   4392
         TabIndex        =   40
         Top             =   2750
         Width           =   2640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "格式92其他所得：給付項目代號"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   9
         Left            =   4392
         TabIndex        =   39
         Top             =   3250
         Width           =   2508
      End
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3960
         Width           =   5700
         VariousPropertyBits=   671105055
         Size            =   "7223;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別： "
         Height          =   180
         Index           =   12
         Left            =   -74880
         TabIndex        =   38
         Top             =   768
         Width           =   948
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "所得總額："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   -73788
         TabIndex        =   36
         Top             =   4308
         Width           =   972
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "扣繳稅額："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   -71484
         TabIndex        =   35
         Top             =   4308
         Width           =   972
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "合計："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   -74640
         TabIndex        =   34
         Top             =   4308
         Width           =   588
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PS：共用欄位：格式 51 租賃所得：房屋稅籍編號"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   8
         Left            =   3120
         TabIndex        =   31
         Top             =   2500
         Width           =   3876
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "共用欄位："
         Height          =   180
         Index           =   7
         Left            =   600
         TabIndex        =   29
         Top             =   2480
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "起迄月份：          －"
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   28
         Top             =   2160
         Width           =   1530
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "所得人代號：                              －"
         Height          =   180
         Left            =   -72480
         TabIndex        =   27
         Top             =   468
         Width           =   2700
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "所得年度：             －"
         Height          =   180
         Left            =   -74880
         TabIndex        =   26
         Top             =   468
         Width           =   1704
      End
      Begin MSForms.Label lblDsp 
         Height          =   300
         Index           =   2
         Left            =   2100
         TabIndex        =   24
         Top             =   1530
         Width           =   3090
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "扣繳稅額： "
         Height          =   180
         Index           =   6
         Left            =   600
         TabIndex        =   23
         Top             =   3120
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所得總額： "
         Height          =   180
         Index           =   5
         Left            =   600
         TabIndex        =   22
         Top             =   2800
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別： "
         Height          =   180
         Index           =   3
         Left            =   600
         TabIndex        =   21
         Top             =   1520
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所得年度： "
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   20
         Top             =   555
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "名　　稱："
         Height          =   180
         Left            =   600
         TabIndex        =   19
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所得人代號："
         Height          =   180
         Index           =   4
         Left            =   420
         TabIndex        =   18
         Top             =   880
         Width           =   1080
      End
      Begin MSForms.Label lblDsp 
         Height          =   300
         Index           =   1
         Left            =   1530
         TabIndex        =   12
         Top             =   1200
         Width           =   1440
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2540;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "格式代號：            (50-53；5A-5E；91-92；9A-9B)"
         Height          =   180
         Index           =   2
         Left            =   600
         TabIndex        =   11
         Top             =   1845
         Width           =   3885
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
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
            Picture         =   "frm170027.frx":004D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170027.frx":0369
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170027.frx":0685
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170027.frx":0861
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170027.frx":0B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170027.frx":0E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170027.frx":11B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170027.frx":14D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170027.frx":17ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170027.frx":1B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170027.frx":1E25
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8220
      _ExtentX        =   14499
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
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frm170027"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/22 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2008/12/26 add by sonia
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_OID As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim m_bConfirmCheck As Boolean
Dim m_bActived As Boolean


Private Sub cmdok_Click()
   If txt1(0) & txt1(1) & txt1(2) & txt1(3) <> "" Then
      If RunNick(txt1(0), txt1(1)) Then
         txt1(0).SetFocus
         Exit Sub
      End If
      If RunNick(txt1(2), txt1(3)) Then
         txt1(2).SetFocus
         Exit Sub
      End If
      GetData
   Else
      MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
      txt1(0).SetFocus
   End If
End Sub

Sub GetData()
Dim stCon As String
   txtSum(0) = ""
   txtSum(1) = ""
   
   stCon = ""
   If txt1(0) <> "" Then
      stCon = stCon & " and oid01>=" & Val(txt1(0)) + 1911
   End If
   If txt1(1) <> "" Then
      stCon = stCon & " and oid01<=" & Val(txt1(1)) + 1911
   End If
   If txt1(2) <> "" Then
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'stCon = stCon & " and replace(oid02,'A','0')>='" & txt1(2) & "' "
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      stCon = stCon & " and substr(oid02,1,2)||replace(substr(oid02,3,1),'A','0')||substr(oid02,4)>='" & txt1(2) & "' "
   End If
   If txt1(3) <> "" Then
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'stCon = stCon & " and replace(oid02,'A','0')<='" & txt1(3) & "' "
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      stCon = stCon & " and substr(oid02,1,2)||replace(substr(oid02,3,1),'A','0')||substr(oid02,4)<='" & txt1(3) & "' "
   End If
   'add by sonia 2018/12/25 加公司別條件
   If txt1(4) <> "" Then
      stCon = stCon & " and oid03='" & txt1(4) & "' "
   End If
   'end 2018/12/25
   
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'strExc(0) = "SELECT oid01-1911,oid05||'~'||oid06,oid02,nvl(oi04,st02),oid03,oid04,oid07,oid08,oid09 FROM OtherIncomeData,Otherincomer,staff " & _
               " where oid02=oi01(+) and replace(oid02,'A','0')=st01(+) " & stCon & " order by oid01,oid02,oid03,oid04"
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   strExc(0) = "SELECT oid01-1911,oid05||'~'||oid06,oid02,nvl(oi04,st02),oid03,oid04,oid07,oid08,oid09 FROM OtherIncomeData,Otherincomer,staff " & _
               " where oid02=oi01(+) and substr(oid02,1,2)||replace(substr(oid02,3,1),'A','0')||substr(oid02,4)=st01(+) " & stCon & " order by oid01,oid02,oid03,oid04"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set GRD1.Recordset = RsTemp.Clone
      GRD1.FormatString = GRD1.FormatString
      SetGrd
      
      'Added by Morgan 2013/10/17
      If RsTemp.RecordCount > 0 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            txtSum(0) = Val(txtSum(0)) + Val("" & RsTemp("oid08"))
            txtSum(1) = Val(txtSum(1)) + Val("" & RsTemp("oid09"))
            RsTemp.MoveNext
         Loop
         txtSum(0) = Format(txtSum(0), "#,##0")
         txtSum(1) = Format(txtSum(1), "#,##0")
      End If
      'end 2013/10/17
   End If
End Sub

Private Sub Form_Activate()
   If m_bActived = False Then
      SetInputEntry
      m_bActived = True
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   
   InitialField
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      Form_KeyDown vbKeyF2, 0
   End If
   UpdateToolbarState
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170027 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from OtherIncomeData where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_OID = .Fields.Count
      ReDim m_FieldList(TF_OID) As FIELDITEM
      For Each oText In txtOID
         idx = oText.Index
         m_FieldList(idx).fiName = "OID" & Format(idx, "00")
         'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
         'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
            m_FieldList(idx).fiType = 0
         'Else
         '   m_FieldList(idx).fiType = 1
         'End If
         'end 2017/06/29
      Next
      End With
   End If
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
Dim stKey01 As String
Dim stKey02 As String
Dim stKey03 As String
Dim stKey04 As String
Dim adoRst As New ADODB.Recordset
   
   stKey01 = Val(txtOID(1)) + 1911
   stKey02 = txtOID(2)
   stKey03 = txtOID(3)
   stKey04 = txtOID(4)
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM OtherIncomeData" & _
            " WHERE oid01 = '" & stKey01 & "' and oid02= '" & stKey02 & "' and oid03= '" & stKey03 & "' and oid04= '" & stKey04 & "'"
      Case -2
         strExc(0) = "SELECT * FROM OtherIncomeData order by 1 ASC,2 ASC,3 ASC,4 ASC"
      Case -1
         strExc(0) = "SELECT * FROM OtherIncomeData" & _
            " WHERE oid01||oid02||oid03||oid04 <'" & stKey01 & stKey02 & stKey03 & stKey04 & "' order by 1 DESC,2 DESC,3 DESC,4 DESC"
      Case 1
         strExc(0) = "SELECT * FROM OtherIncomeData" & _
            " WHERE oid01||oid02||oid03||oid04 >'" & stKey01 & stKey02 & stKey03 & stKey04 & "' order by 1 ASC,2 ASC,3 ASC,4 ASC"
      Case 2
         strExc(0) = "SELECT * FROM OtherIncomeData order by 1 DESC,2 DESC,3 DESC,4 DESC"
   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      UpdateCtrlData adoRst
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         MsgBox "查無資料！", vbInformation
         ClearField
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtOID(1).SetFocus
      txtOID_GotFocus 1
   End If
End Function

Private Sub GRD1_Click()
   Dim lCurRow As Long, i As Integer, j As Integer
   lCurRow = GRD1.row
   If lCurRow > 0 Then
      If GRD1.TextMatrix(lCurRow, 0) <> "" Then
         If GRD1.CellBackColor <> &HFFC0C0 Then
            GRD1.Visible = False
            For j = 1 To GRD1.Rows - 1
               GRD1.row = j
               If GRD1.CellBackColor <> QBColor(15) Then
                  For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = QBColor(15)
                  Next i
               End If
            Next j
            GRD1.row = lCurRow
            For i = 0 To GRD1.Cols - 1
                GRD1.col = i
                GRD1.CellBackColor = &HFFC0C0
            Next i
            GRD1.Visible = True
         End If
      End If
   End If
End Sub

Private Sub GRD1_DblClick()
Dim lCurRow As Long
   
   lCurRow = GRD1.row
   '呼叫查詢
   If lCurRow > 0 Then
      If GRD1.TextMatrix(lCurRow, 0) <> "" Then
         If TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
            If txtOID(1).Locked = False Then
               txtOID(1).Text = GRD1.TextMatrix(lCurRow, 0)
               txtOID(2).Text = GRD1.TextMatrix(lCurRow, 2)
               txtOID(3).Text = GRD1.TextMatrix(lCurRow, 4)
               txtOID(4).Text = GRD1.TextMatrix(lCurRow, 5)
               If TBar1.Buttons(11).Enabled = True Then
                  Call Tbar1_ButtonClick(TBar1.Buttons(11))
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 2 Then
      txt1(0).SetFocus
      TextInverse txt1(0)
   ElseIf SSTab1.Tab = 0 And PreviousTab = 2 Then
      GRD1_DblClick
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   CloseIme
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtOID_GotFocus(Index As Integer)
   TextInverse txtOID(Index)
   CloseIme
End Sub

Private Sub ClearField()
   For Each oText In txtOID
      oText.Text = Empty
   Next
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_OID
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   m_bConfirmCheck = False
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   With p_Rst
   If .RecordCount > 0 Then
      For Each oText In txtOID
         idx = oText.Index
         '所得年度轉民國年
         If idx = 1 Then
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName) - 1911
         Else
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         End If
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         oText.Text = m_FieldList(idx).fiOldData
         'oText.Tag = m_FieldList(idx).fiOldData
      Next
      
      If ClsPDGetOtherIncomer(txtOID(2), strExc(1)) = True Then
         lblDsp(1) = strExc(1)
      ElseIf ClsPDGetStaffN(txtOID(2), strExc(1), , True) Then
         lblDsp(1) = strExc(1)
      End If
      lblDsp(2) = CompNameQuery(txtOID(3))
      
      CUID(1) = "" & .Fields("oid10")
      CUID(2) = "" & .Fields("oid11")
      CUID(3) = "" & .Fields("oid12")
      CUID(4) = "" & .Fields("oid13")
      CUID(5) = "" & .Fields("oid14")
      CUID(6) = "" & .Fields("oid15")
   End If
   End With
   UpdateCUID CUID, textCUID
   txtOID(1).Tag = txtOID(1)
   txtOID(2).Tag = txtOID(2)
   txtOID(3).Tag = txtOID(3)
   txtOID(4).Tag = txtOID(4)
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtOID
      oText.Locked = bLocked
   Next
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
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

' 執行指令
Public Sub OnAction(ByVal KeyCode As Integer)
   Dim bCancel As Boolean
   
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         SSTab1.Tab = 0
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF3 ' 修改
         SSTab1.Tab = 0
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF5 ' 刪除
         SSTab1.Tab = 0
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         SSTab1.Tab = 0
         m_EditMode = 4
         SetCtrlReadOnly True
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
         bCancel = False
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  bCancel = True
               End If
            Case Else
               bCancel = True
         End Select
         If bCancel = True Then
            txtOID(1) = txtOID(1).Tag
            txtOID(2) = txtOID(2).Tag
            txtOID(3) = txtOID(3).Tag
            txtOID(4) = txtOID(4).Tag
            m_EditMode = 0
            SetInputEntry
            ShowRecord
            UpdateToolbarState
         End If
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
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
         If m_bUpdate And txtOID(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtOID(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtOID(1) <> "" Then
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

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1
         txtOID(1).Locked = False
         If Me.Visible = True Then
            txtOID(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 2
         txtOID(1).Locked = True
         If Me.Visible = True Then
            txtOID(3).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 4
         txtOID(1).Locked = False
         txtOID(2).Locked = False
         txtOID(3).Locked = False
         txtOID(4).Locked = False
         If Me.Visible = True Then
            txtOID(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case Else
         txtOID(1).Locked = True
         If Me.Visible = True Then
            txtOID(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = True
   End Select
End Sub

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If AddRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 3: '刪除
         If DelRecord = True Then
            OnWork = True
            m_EditMode = 0
            ShowRecord 2
         End If
      
      Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtOID(1).SetFocus
               txtOID_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   For Each oText In txtOID
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtOID_Validate idx, bCancel
         If bCancel = True Then
            txtOID(idx).SetFocus
            txtOID_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   '查詢
   If m_EditMode = 4 Then
      If txtOID(1) = "" Then
         ShowMsg "請輸入所得年度 !"
         txtOID(1).SetFocus
         txtOID_GotFocus 1
         GoTo EscPoint
      End If
      If txtOID(2) = "" Then
         ShowMsg "請輸入所得人代號 !"
         txtOID(2).SetFocus
         txtOID_GotFocus 2
         GoTo EscPoint
      End If
      If txtOID(3) = "" Then
         ShowMsg "請輸入公司別 !"
         txtOID(3).SetFocus
         txtOID_GotFocus 3
         GoTo EscPoint
      End If
      If txtOID(4) = "" Then
         ShowMsg "請輸入格式 !"
         txtOID(4).SetFocus
         txtOID_GotFocus 4
         GoTo EscPoint
      End If
      
   '維護
   Else
      If txtOID(1) = "" And txtOID(1).Locked = False Then
         ShowMsg "請輸入所得年度 !"
         txtOID(1).SetFocus
         txtOID_GotFocus 1
         GoTo EscPoint
      End If
      If txtOID(2) = "" And txtOID(2).Locked = False Then
         ShowMsg "請輸入所得人代號 !"
         txtOID(2).SetFocus
         txtOID_GotFocus 2
         GoTo EscPoint
      End If
      If txtOID(3) = "" And txtOID(3).Locked = False Then
         ShowMsg "請輸入公司別 !"
         txtOID(3).SetFocus
         txtOID_GotFocus 3
         GoTo EscPoint
      End If
      If txtOID(4) = "" And txtOID(4).Locked = False Then
         ShowMsg "請輸入格式代號 !"
         txtOID(4).SetFocus
         txtOID_GotFocus 4
         GoTo EscPoint
      End If
                  
      If txtOID(5) = "" And txtOID(5).Locked = False Then
         ShowMsg "請輸入起始月份 !"
         txtOID(5).SetFocus
         txtOID_GotFocus 5
         GoTo EscPoint
      End If
      If txtOID(6) = "" And txtOID(6).Locked = False Then
         ShowMsg "請輸入截止月份 !"
         txtOID(6).SetFocus
         txtOID_GotFocus 6
         GoTo EscPoint
      End If
      If txtOID(4) = "51" And txtOID(7) = "" And txtOID(7).Locked = False Then
         ShowMsg "格式代號為51時, 請於共用欄位輸入房屋稅籍編號 !"
         txtOID(7).SetFocus
         txtOID_GotFocus 7
         GoTo EscPoint
      End If
      If txtOID(4) = "9A" And txtOID(7) = "" And txtOID(7).Locked = False Then
         ShowMsg "格式代號為9A時, 請於共用欄位輸入執行業務業別代號 !"
         txtOID(7).SetFocus
         txtOID_GotFocus 7
         GoTo EscPoint
      End If
      If txtOID(4) = "9B" And txtOID(7) = "" And txtOID(7).Locked = False Then
         ShowMsg "格式代號為9B時, 請於共用欄位輸入費用代號 !"
         txtOID(7).SetFocus
         txtOID_GotFocus 7
         GoTo EscPoint
      End If
      If txtOID(4) = "92" And txtOID(7) = "" And txtOID(7).Locked = False Then
         ShowMsg "格式代號為92時, 請於共用欄位輸入給付項目代號 !"
         txtOID(7).SetFocus
         txtOID_GotFocus 7
         GoTo EscPoint
      End If
      If txtOID(8) = "" And txtOID(8).Locked = False Then
         ShowMsg "請輸入所得總額 !"
         txtOID(8).SetFocus
         txtOID_GotFocus 8
         GoTo EscPoint
      End If
   End If
   TxtValidate = True
   
EscPoint:
   m_bConfirmCheck = False
    
End Function

Private Function AddRecord() As Boolean
Dim stCols As String, stValues As String, stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtOID
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            '所得年度轉西元年, 取消因為UpdateFieldNewData已+1911
            'If idx = 1 Then
            '   stValues = stValues & "," & CNULL(Val((m_FieldList(idx).fiNewData) + 1911), True)
            'Else
               stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
            'End If
         End If
      End If
   Next
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO OtherIncomeData (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
'   stSQL = "select max(oid02) from OtherIncomeData where oid01='" & txtOID(1) & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
'   If intI = 1 Then
'      txtOID(2) = RsTemp.Fields(0)
'   End If
   
   cnnConnection.CommitTrans
   
   AddRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
    
End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE OtherIncomeData SET "
   stSet = ""
   For Each oText In txtOID
      idx = oText.Index
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
      stSQL = stSQL & stSet & " where oid01=" & Val(txtOID(1)) + 1911 & " and oid02='" & txtOID(2) & "' and oid03='" & txtOID(3) & "' and oid04='" & txtOID(4) & "'; end; "
      
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

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

Private Sub UpdateFieldNewData()
   For Each oText In txtOID
      idx = oText.Index
      Select Case idx
         Case 1
            m_FieldList(idx).fiNewData = Val(oText.Text) + 1911
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub

Private Sub txtOID_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 2, 3, 4, 7
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtOID_Validate(Index As Integer, Cancel As Boolean)
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Select Case Index
         'Added by Morgan 2013/2/5
         Case 1
            If txtOID(1) <> "" Then
               If Val(txtOID(1)) > strSrvDate(2) \ 10000 Then
                  MsgBox "所得年度不可晚於當年度！"
                  Cancel = True
               ElseIf Val(txtOID(1)) < strSrvDate(2) \ 10000 - 1 Then
                  MsgBox "所得年度不可早於去年度！"
                  Cancel = True
               End If
            End If
         'end 2013/2/5
         Case 2
            If txtOID(Index) <> "" Then
               If ClsPDGetOtherIncomer(txtOID(Index), strExc(1)) = True Then
                  lblDsp(1) = strExc(1)
               Else
                  'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
                  'If ChkStaffID(Replace(txtOID(Index), "A", "0")) = True Then
                  If ChkStaffID(Left(txtOID(Index), 1) & Replace(Mid(txtOID(Index), 2), "A", "0")) = True Then
                     Cancel = True
                  End If
                  If Cancel = False Then
                     If ClsPDGetStaffN(txtOID(Index), strExc(1), , True) = False Then
                        Cancel = True
                        lblDsp(1) = ""
                     Else
                        lblDsp(1) = strExc(1)
                     End If
                  End If
               End If
            End If
         Case 3
            If txtOID(Index) <> "" Then
               lblDsp(2) = CompNameQuery(txtOID(Index))
               If lblDsp(2) = "" Then
                  ShowMsg "公司別錯誤 !"
                  Cancel = True
               End If
            End If
         Case 4
            If txtOID(Index) <> "" Then
               'modify by sonia 2023/2/7 補滿2碼再比較,否則只輸2也會過
               'If InStr("50,51,52,53,5A,5B,5C,5D,5E,91,92,9A,9B", txtOID(Index)) > 0 Then
               If InStr("50,51,52,53,5A,5B,5C,5D,5E,91,92,9A,9B", Left(txtOID(Index) & Space(2), 2)) > 0 Then
                  'Added by Morgan 2013/1/23
                  If InStr("50,51,52,5A,5B,5C,5D,9A,9B", txtOID(Index)) > 0 Then
                     MsgBox "102年起，50,51,52,5A,5B,5C,5D,9A,9B 格式有補充保費問題，請改至【 其他各類所得資料(平日) 】輸入！", vbInformation
                     Cancel = True
                  End If
               Else
                  ShowMsg "格式代號錯誤 !"
                  Cancel = True
               End If
            End If
         Case 5, 6
            If txtOID(Index) <> "" Then
               If Val(txtOID(Index)) < 1 Or Val(txtOID(Index)) > 12 Then
                  ShowMsg "月份不可小於 1 或超過 12 !"
                  Cancel = True
               End If
            End If
         Case 7  '格式為9A時,共用欄位檢查AllCode的執行業務業別代號
            If txtOID(Index) <> "" Then
               If txtOID(4) = "9A" Then
                  strExc(0) = "SELECT ac03 FROM AllCode where ac01='07' and ac02='" & txtOID(Index) & "' "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI <> 1 Then
                     ShowMsg "共用欄位輸入的執行業務業別代號錯誤 !"
                     Cancel = True
                  End If
               End If
            End If
      End Select
      
      If Cancel = True Then TextInverse txtOID(Index)
      
      '若是按確定的檢查時略過, 檢查代號檔
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
         End Select
      End If
   End If
End Sub

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '刪除
   stSQL = "delete from OtherIncomeData where oid01=" & Val(txtOID(1)) + 1911 & " and oid02='" & txtOID(2) & "' and oid03='" & txtOID(3) & "' and oid04='" & txtOID(4) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtOID(1).Tag = ""
   txtOID(2).Tag = ""
   txtOID(3).Tag = ""
   txtOID(4).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, iCol As Integer
   
   arrGridHeadText = Array("年度", "月份", "所得人代號", "名　稱", "公司別", "格式", "共用欄位", "所得總額", "扣繳稅額")
   arrGridHeadWidth = Array(450, 600, 1100, 1200, 600, 500, 1250, 1000, 950)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
      'Added by Morgan 2013/10/17
      If iRow < 7 Then
         GRD1.ColAlignment(iRow) = flexAlignLeftCenter
      Else
         GRD1.ColAlignment(iRow) = flexAlignRightCenter
      End If
   Next
   
   
   GRD1.Visible = True
End Sub
