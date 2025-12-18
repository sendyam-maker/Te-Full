VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060508 
   BorderStyle     =   1  '單線固定
   Caption         =   "請款函預設備註維護"
   ClientHeight    =   5628
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8268
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5628
   ScaleWidth      =   8268
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7695
      Top             =   450
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
            Picture         =   "frm060508.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060508.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060508.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060508.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060508.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060508.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060508.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060508.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060508.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060508.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060508.frx":1DD8
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
      Width           =   8268
      _ExtentX        =   14584
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
      Height          =   4860
      Left            =   96
      TabIndex        =   7
      Top             =   720
      Width           =   8112
      _ExtentX        =   14309
      _ExtentY        =   8573
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm060508.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label2(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtDB(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtDB(5)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtDB(4)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtDB(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDB(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textCUID"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtDB(12)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label2(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm060508.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblPS"
      Tab(1).Control(1)=   "Label1(12)"
      Tab(1).Control(2)=   "Label1(11)"
      Tab(1).Control(3)=   "Label1(10)"
      Tab(1).Control(4)=   "txtFM2(0)"
      Tab(1).Control(5)=   "txtFM2(1)"
      Tab(1).Control(6)=   "txtFM2(2)"
      Tab(1).Control(7)=   "lblFM2(1)"
      Tab(1).Control(8)=   "lblFM2(2)"
      Tab(1).Control(9)=   "Label1(9)"
      Tab(1).Control(10)=   "Label1(13)"
      Tab(1).Control(11)=   "txtFM2(3)"
      Tab(1).Control(12)=   "Label1(7)"
      Tab(1).Control(13)=   "txtFM2(4)"
      Tab(1).Control(14)=   "lblFM2(0)"
      Tab(1).Control(15)=   "GRD1"
      Tab(1).Control(16)=   "cmdQuery"
      Tab(1).ControlCount=   17
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   300
         Left            =   -72270
         TabIndex        =   13
         Top             =   420
         Width           =   885
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm060508.frx":212C
         Height          =   2535
         Left            =   -74910
         TabIndex        =   8
         Top             =   2220
         Width           =   7905
         _ExtentX        =   13949
         _ExtentY        =   4466
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "流水號|備註內容|本所案號|代理人|申請人"
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
         _Band(0).Cols   =   5
      End
      Begin MSForms.Label lblFM2 
         Height          =   252
         Index           =   0
         Left            =   -69432
         TabIndex        =   40
         Top             =   436
         Width           =   2256
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "3979;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   288
         Index           =   4
         Left            =   -70080
         TabIndex        =   14
         Top             =   418
         Width           =   600
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1058;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   180
         Index           =   7
         Left            =   -71064
         TabIndex        =   39
         Top             =   472
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   180
         Index           =   4
         Left            =   132
         TabIndex        =   38
         Top             =   3108
         Width           =   900
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   0
         Left            =   1824
         TabIndex        =   37
         Top             =   3108
         Width           =   3012
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "5313;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   12
         Left            =   1080
         TabIndex        =   5
         Top             =   3048
         Width           =   600
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1058;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   520
         Index           =   3
         Left            =   -73860
         TabIndex        =   17
         Top             =   1350
         Width           =   6200
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "10936;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註內容："
         Height          =   180
         Index           =   13
         Left            =   -74880
         TabIndex        =   36
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "模糊比對"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   9
         Left            =   -74880
         TabIndex        =   35
         Top             =   1620
         Visible         =   0   'False
         Width           =   720
      End
      Begin MSForms.TextBox textCUID 
         Height          =   288
         Left            =   120
         TabIndex        =   34
         Top             =   4272
         Width           =   7860
         VariousPropertyBits=   671105055
         Size            =   "13864;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   1
         Left            =   1080
         TabIndex        =   0
         Top             =   540
         Width           =   630
         VariousPropertyBits=   671105049
         MaxLength       =   4
         Size            =   "1111;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   1065
         Index           =   2
         Left            =   1080
         TabIndex        =   1
         Top             =   870
         Width           =   5580
         VariousPropertyBits=   -1466941413
         MaxLength       =   600
         ScrollBars      =   2
         Size            =   "9842;1879"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   4
         Left            =   1080
         TabIndex        =   3
         Top             =   2370
         Width           =   1170
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2064;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   5
         Left            =   1080
         TabIndex        =   4
         Top             =   2730
         Width           =   1170
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2064;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   3
         Left            =   1080
         TabIndex        =   2
         Top             =   2040
         Width           =   1575
         VariousPropertyBits=   671105051
         MaxLength       =   12
         Size            =   "2778;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   2
         Left            =   2325
         TabIndex        =   33
         Top             =   2790
         Width           =   5505
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "9710;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   1
         Left            =   2325
         TabIndex        =   32
         Top             =   2430
         Width           =   5505
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "9710;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   255
         Index           =   2
         Left            =   -72720
         TabIndex        =   31
         Top             =   1065
         Width           =   5595
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "9869;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   255
         Index           =   1
         Left            =   -72720
         TabIndex        =   30
         Top             =   750
         Width           =   5595
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "9869;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   285
         Index           =   2
         Left            =   -73860
         TabIndex        =   16
         Top             =   1050
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1940;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   285
         Index           =   1
         Left            =   -73860
         TabIndex        =   15
         Top             =   735
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1940;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   285
         Index           =   0
         Left            =   -73860
         TabIndex        =   12
         Top             =   420
         Width           =   1515
         VariousPropertyBits=   671105051
         MaxLength       =   12
         Size            =   "2672;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   10
         Left            =   -74880
         TabIndex        =   29
         Top             =   465
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   11
         Left            =   -74880
         TabIndex        =   28
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         Height          =   180
         Index           =   12
         Left            =   -74880
         TabIndex        =   27
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblPS 
         Caption         =   "P.S. 輸入本所案號會另外帶該案代理人和申請人的其他設定"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74880
         TabIndex        =   26
         Top             =   2010
         Width           =   4845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "　　　2.備註內容不輸入表示請款函不要預設備註"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   192
         Index           =   4
         Left            =   1080
         TabIndex        =   25
         Top             =   3924
         Width           =   4236
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "　　　   FCP03914200,FCP04537300,FCP03577100,FCP05102400皆不帶預設備註"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   192
         Index           =   3
         Left            =   1080
         TabIndex        =   24
         Top             =   3720
         Width           =   6420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "注意：1.申請人X5403200,X5988900,如遇到605年費或FCP03951700,"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   192
         Index           =   2
         Left            =   1080
         TabIndex        =   23
         Top             =   3480
         Width           =   5460
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "模糊比對"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   6
         Left            =   6750
         TabIndex        =   22
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "流水號："
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   21
         Top             =   585
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註內容："
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   20
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         Height          =   180
         Index           =   2
         Left            =   312
         TabIndex        =   19
         Top             =   2430
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   3
         Left            =   312
         TabIndex        =   18
         Top             =   2790
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   5
         Left            =   132
         TabIndex        =   11
         Top             =   2064
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "注意：1.代理人/申請人可輸入6碼或8碼，6碼代表含關係企業。 112/1/9 取消"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   192
         Index           =   1
         Left            =   1872
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   6348
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "　　　2.代理人/申請人無論6碼或8碼均包含更名前編號。112/1/9 取消"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   192
         Index           =   0
         Left            =   1872
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   5820
      End
   End
End
Attribute VB_Name = "frm060508"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/02/18 整合特殊備註維護：1.將現有資料6碼Y/X編號補足為8碼；2.在輸入Y/X編號若為6碼，統一補足為8碼。
'Memo by Lydia 2021/11/01 改成Form2.0 ; GRD1改字型=新細明體-ExtB、txtDB(index)、Label2(index)、textCUID、txtFM2(index)、lblFM2(index)
'Memo by Lydia 2021/11/01 畫面頁籤改成「單筆資料」和「多筆查詢」：上方工具列的「查詢」帶出第一筆符合的資料，在多筆查詢的頁籤可以輸入條件進行查詢，並且在下方的Grid呈現多筆資料。
'Created by Lydia 2015/02/06 新增-請款函預設備註維護(DebitNotePS)
Option Explicit

Dim m_EditMode As Integer '0:瀏覽 1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_blnColOrderAsc As Boolean 'Added by Lydia 2021/11/11 欄位資料由小到大排序
Dim oText As Control, oLabel As Control
Dim stCon As String, stSQL As String, intR As Integer
Dim rsRead As New ADODB.Recordset
Private Const mPtyList As String = "101,102,103,104,105,125,301,302,303,304,305,306,307,601,602,603,605,701,702,703,704,401,413,411,1001,1008,907,913,1604,416,608,417,425" 'Added by Lydia 2023/08/01 可輸入案件性質參考frm060306的範圍。

'Added by Lydia 2021/11/01
Private Sub cmdQuery_Click()
   
   stCon = ""
   If txtFM2(0) <> "" Then
      If Trim(txtFM2(1).Tag & txtFM2(2).Tag) = "" Then
          stCon = stCon & " and dnps03='" & txtFM2(0) & "'"
      Else
          '另外抓本所案號的相關Y編號、X編號條件
          stCon = stCon & " and (dnps03='" & txtFM2(0) & "'"
          If txtFM2(1).Tag <> "" Then stCon = stCon & " or instr(" & CNULL(txtFM2(1).Tag) & ", dnps04) > 0 "
          If txtFM2(2).Tag <> "" Then stCon = stCon & " or instr(" & CNULL(txtFM2(2).Tag) & ", dnps05) > 0 "
          stCon = stCon & ") "
      End If
   Else
      txtFM2(1).Tag = "": txtFM2(2).Tag = ""   '清空本所案號的相關Y編號、X編號條件
   End If
   If txtFM2(1) <> "" Then
      stCon = stCon & " and dnps04 like '" & txtFM2(1) & "%'"
   End If
   If txtFM2(2) <> "" Then
      stCon = stCon & " and dnps05 like '" & txtFM2(2) & "%'"
   End If
   'Added by Lydia 2022/10/03 增加"備註"查詢
   If txtFM2(3) <> "" Then
       stCon = stCon & " and upper(dnps02) like '%" & ChgSQL(UCase(txtFM2(3))) & "%' "
   End If
   'end 2022/10/03
   'Added by Lydia 2023/08/01 案件性質
   If txtFM2(4) <> "" Then
      stCon = stCon & " and dnps12 = '" & txtFM2(4) & "'"
   End If
   
   'Modified by Lydia 2023/08/01
   'stSQL = "SELECT DNPS01,DNPS02,DNPS03,DNPS04,DNPS05 FROM DEBITNOTEPS WHERE 1=1 " & stCon
   stSQL = "SELECT DNPS01,DNPS02,DNPS03,DNPS04,DNPS05,DNPS12,CPM03 AS DNPS12T " & _
           "FROM DEBITNOTEPS,CASEPROPERTYMAP WHERE 'FCP'=CPM01(+) AND DNPS12=CPM02(+) " & stCon
   stSQL = stSQL & " ORDER BY DNPS01"
   intR = 0
   Set rsRead = ClsLawReadRstMsg(intR, stSQL)
   
   Call SetGrd(True)
   If intR = 1 Then
        grd1.FixedCols = 0
        Set grd1.Recordset = rsRead
        Call SetGrd
        grd1.FixedCols = 3
   End If
End Sub

'Added by Lydia 2021/11/01
Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modified by Lydia 2023/08/01
   'arrGridHeadText = Array("流水號", "備註內容", "本所案號", "代理人", "申請人")
   'arrGridHeadWidth = Array(800, 1200, 1200, 1000, 1000)
   arrGridHeadText = Array("流水號", "備註內容", "本所案號", "代理人", "申請人", "DNPS12", "案件性質")
   arrGridHeadWidth = Array(800, 1200, 1200, 1000, 1000, 0, 1000)
   
   grd1.Visible = False
   grd1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
        grd1.Clear
        grd1.Rows = 2
   End If
   For iRow = 0 To grd1.Cols - 1
      grd1.row = 0
      grd1.col = iRow
      grd1.Text = arrGridHeadText(iRow)
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1.CellAlignment = flexAlignCenterCenter
   Next

   grd1.Visible = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Memo by Lydia 2021/11/01 原程式搬到Form_KeyUp

End Sub

'Added by Lydia 2021/11/01
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Memo by Lydia 2021/11/01 從Form_KeyDown搬來
   Screen.MousePointer = vbHourglass
   Select Case KeyCode
      Case vbKeyF2 '新增
         KeyCode = 0: Action 1
      Case vbKeyF3 '修改
         KeyCode = 0: Action 2
      Case vbKeyF4: '查詢
         KeyCode = 0: Action 4
      Case vbKeyF5 '刪除
         KeyCode = 0: Action 3
      Case vbKeyHome '第一筆
         KeyCode = 0: Action 6
      Case vbKeyPageUp '上一筆
         KeyCode = 0: Action 7
      Case vbKeyPageDown '下一筆
         KeyCode = 0: Action 8
      Case vbKeyEnd: '最後筆
         KeyCode = 0: Action 9
      'Modified by Lydia 2021/11/22 Lydia 2021/11/22 取消以ENTER控制為換行的功能 (Form2.0修改之維護資料功能Toolbar之修改統一)
'      Case vbKeyF9, vbKeyReturn '確定
      Case vbKeyF9 '確定
         KeyCode = 0: Action 11
      
      Case vbKeyF10 '取消
         KeyCode = 0: Action 12
      Case vbKeyEscape '結束
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            KeyCode = 0: Action 14
         End If
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction("frm060508", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm060508", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm060508", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm060508", strFind, False)
  
   MoveFormToCenter Me
   
   'Added by Lydia 2021/11/01
   For Each oLabel In lblFM2
       oLabel.BackColor = &H8000000F
   Next
   Call SetGrd(True)
   'end 2021/11/01
   
   textCUID.BackColor = &H8000000F
   Action 6 '預設第一筆
   UpdateToolbarState
   
   Me.SSTab1.Tab = 1 'Added by Lydia 2021/11/01 改從多筆查詢頁籤開始
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060508 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
   Action Button.Index
   Screen.MousePointer = vbDefault
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
         If m_bUpdate And txtDB(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtDB(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtDB(1) <> "" Then
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

Private Sub TxtLock()
   Select Case m_EditMode
   Case 0 '瀏覽
      For Each oText In txtDB
         oText.Locked = True
      Next
      SSTab1.TabEnabled(1) = True
   Case Else
      For Each oText In txtDB
         oText.Locked = False
      Next
      If m_EditMode <> 4 Then
         txtDB(1).Locked = True
         txtDB(2).SetFocus
         txtDB_GotFocus 2
      End If
      SSTab1.TabEnabled(1) = False
   End Select
End Sub
Private Sub Action(Index As Integer)
Dim bCancel As Boolean 'Added by Lydia 2019/05/20
Dim strKind As String 'Added by Lydia 2021/11/01

   If TBar1.Buttons(Index).Enabled = False Then Exit Sub

On Error GoTo ErrHand

   SSTab1.Tab = 0
   Select Case Index
      Case 1 '按下新增
        m_EditMode = 1
        FormReset
        
      Case 2 '按下修改
         m_EditMode = 2

      Case 3 '按下刪除
         If txtDB(1).Text = "" Then
             MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
             Exit Sub
         End If

         If DelMsg() = True Then
            If FormDelete() = False Then
               MsgBox "刪除失敗!", vbCritical
               Exit Sub
            '刪除後移到最末筆
            Else
               ShowRecord 3
            End If
         End If

      Case 4 '按下查詢
         FormReset
         m_EditMode = 4
         txtDB(1).Enabled = True
         txtDB(1).SetFocus
         Label1(6).Visible = True
         
      Case 6 '第一筆
         ShowRecord 0
      Case 7 '前一筆
         ShowRecord 1
      Case 8 '後一筆
         ShowRecord 2
      Case 9 '最後筆
         ShowRecord 3
      Case 11 '按下確定
         'Added by Lydia 2019/05/20 使用者輸入案號後，直接按Enter無法觸發檢查案號之功能 (by Winfrey)
         If Val(m_EditMode) > 0 And Trim(txtDB(3)) <> "" And ((Left(Trim(txtDB(3)), 1) = "P" And Len(Trim(txtDB(3))) < 10) Or (Left(Trim(txtDB(3)), 3) = "FCP" And Len(Trim(txtDB(3))) < 12)) Then
             Call txtDB_Validate(3, bCancel)
             If bCancel = True Then
                 Exit Sub
             End If
         End If
         
         Select Case m_EditMode
            '新增,修改
            Case 1, 2
               'Modified by Lydia 2021/11/01 新增,修改都要判斷
               'If m_EditMode = 1 Then
               '   If RecIsExist = True Then Exit Sub
               'End If
               If RecIsExist = True Then Exit Sub
               
               If TxtValidate = False Then
                  Exit Sub
               Else
                  If FormSave() = False Then
                     MsgBox "存檔失敗!", vbCritical
                     Exit Sub
                  Else
                     strKind = m_EditMode 'Added by Lydia 2021/11/01 記錄新增模式
                     m_EditMode = 0
                     'Modified by Morgan 2017/9/13
                     'If m_EditMode = 1 Then
                     If txtDB(1) = "" Then
                     'end 2017/9/13
                        ShowRecord 3
                     Else
                        ReadData txtDB(1)
                     End If
                     'Added by Lydia 2021/11/01 在新增存檔後自動帶入多筆查詢顯示本次新增記錄
                     If strKind = "1" Then
                        For Each oText In txtFM2
                            oText.Text = ""
                            oText.Tag = ""
                        Next
                        For Each oLabel In lblFM2
                            oLabel.Caption = ""
                        Next
                        If txtDB(3) <> "" Then
                            txtFM2(0) = txtDB(3)
                            Call txtFM2_Validate(0, False)
                        Else
                            If txtDB(4) <> "" Then
                               txtFM2(1) = ChangeCustomerS(txtDB(4))
                               Call txtFM2_Validate(1, False)
                            End If
                            If txtDB(5) <> "" Then
                               txtFM2(2) = ChangeCustomerS(txtDB(5))
                               Call txtFM2_Validate(2, False)
                            End If
                        End If
                        SSTab1.Tab = 1
                        Call cmdQuery_Click
                     End If
                     'end 2021/11/01
                  End If
               End If
            '查詢
            Case 4
               If ReadData(txtDB(1)) = False Then
                  MsgBox "無資料!", vbExclamation
                  Exit Sub
               Else
                  m_EditMode = 0
               End If
         End Select
      Case 12 '按下取消
         m_EditMode = 0
         txtDB(1) = txtDB(1).Tag
         If txtDB(1) <> "" Then
            If ReadData(txtDB(1)) = False Then
               ShowRecord 3
            End If
         End If
      Case 14 '結束
         Unload Me
         Exit Sub
   End Select
   UpdateToolbarState
   TxtLock
   Exit Sub
   
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

' 顯示資料
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
 Dim stKey As String
    
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case p_iWay
      Case 0 '第一筆
         strExc(0) = "SELECT nvl(min(DNPS01),0) FROM DebitNotePS"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKey = RsTemp.Fields(0)
            End If
         End If
         
      Case 1 '前一筆
         strExc(0) = "SELECT nvl(max(DNPS01),0) FROM DebitNotePS where DNPS01<" & txtDB(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 6
            Else
               stKey = RsTemp.Fields(0)
            End If
         End If
         
      Case 2 '後一筆
         strExc(0) = "SELECT nvl(min(DNPS01),0) FROM DebitNotePS where DNPS01>" & txtDB(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 7
            Else
               stKey = RsTemp.Fields(0)
            End If
         End If
         
      Case 3 '最後筆
         strExc(0) = "SELECT nvl(max(DNPS01),0) FROM DebitNotePS"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKey = RsTemp.Fields(0)
            End If
         End If
   End Select
   
   
   If stKey <> "" Then
      ReadData stKey
      ShowRecord = True
   End If
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Private Function ReadData(Optional ByVal pKey As String) As Boolean

   '單筆
   If pKey <> "" Then
      stCon = " and DNPS01=" & pKey
   '多筆
   Else
      If txtDB(2) <> "" Then
         'Modified by Morgan 2017/9/13
         'stCon = stCon & " and DNPS02 like '%" & txtDB(2) & "%'"
         stCon = stCon & " and DNPS02 like '%" & ChgSQL(txtDB(2)) & "%'"
      End If
      If txtDB(3) <> "" Then
         stCon = stCon & " and DNPS03='" & txtDB(3) & "'"
      End If
      If txtDB(4) <> "" Then
         stCon = stCon & " and DNPS04 like '" & txtDB(4) & "%'"
      End If
      If txtDB(5) <> "" Then
         stCon = stCon & " and DNPS05 like '" & txtDB(5) & "%'"
      End If
      'Added by Lydia 2023/08/01
      If txtDB(12) <> "" Then
         stCon = stCon & " and DNPS12='" & txtDB(12) & "'"
      End If
      'end 2023/08/01
   End If
   
   FormReset
   
   strExc(0) = "select * from DebitNotePS where 1=1 " & stCon & " order by DNPS01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If m_EditMode = 4 Then
         'Modified by Lydia 2021/11/01 改成單筆查詢
         'Set GRD1.Recordset = RsTemp.Clone
         'GRD1.FormatString = GRD1.FormatString
         'GRD1.ColWidth(1) = 2775
         'GRD1.ColWidth(2) = 1290
         'GRD1.ColWidth(3) = 1500
         'GRD1.ColWidth(4) = 1500
         'For intI = 5 To GRD1.Cols - 1
         '   GRD1.ColWidth(intI) = 0
         'Next
         'If RsTemp.RecordCount > 1 Then
         '   GRD1.Recordset.MoveFirst
         '   SSTab1.Tab = 1
         'Else
         '   SSTab1.Tab = 0
         'End If
         RsTemp.MoveFirst
         'end 2021/11/01
      Else
         SSTab1.Tab = 0
      End If
      SetData RsTemp
      ReadData = True
   End If
End Function

Private Sub SetData(ByRef rsQuery As ADODB.Recordset, Optional ByVal iRow As Integer)
   If iRow > 0 Then
      rsQuery.MoveFirst
      If iRow > 1 Then
         rsQuery.Move iRow - 1
      End If
      SSTab1.Tab = 0
   End If
   
   With rsQuery
   For Each oText In txtDB
      oText = "" & .Fields("DNPS" & Format(oText.Index, "00"))
   Next
   End With
   UpdateCUID rsQuery
   
   txtDB(1).Tag = txtDB(1)
   If txtDB(4) <> "" Then txtDB_Validate 4, False
   If txtDB(5) <> "" Then txtDB_Validate 5, False
   If txtDB(12).Text <> "" Then txtDB_Validate 12, False 'Added by Lydia 2023/08/01
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
   If IsNull(rsSrcTmp.Fields("DNPS06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("DNPS06")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("DNPS06"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("DNPS07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("DNPS07")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("DNPS07"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("DNPS08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("DNPS08")) = False Then
         strTemp = rsSrcTmp.Fields("DNPS08")
         strCTime = Format(strTemp, "00:00:00")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("DNPS09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("DNPS09")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("DNPS09"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("DNPS10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("DNPS10")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("DNPS10"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("DNPS11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("DNPS11")) = False Then
         strTemp = rsSrcTmp.Fields("DNPS11")
         strUTime = Format(strTemp, "00:00:00")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub FormReset()
   
   For Each oText In txtDB
      oText.Text = ""
   Next
   
   For Each oLabel In Label2
      oLabel.Caption = ""
   Next
   
   textCUID = ""
   Label1(6).Visible = False
End Sub

Private Sub txtDB_GotFocus(Index As Integer)
   TextInverse txtDB(Index)
   If Index = 2 Then
      OpenIme
   Else
      CloseIme
   End If
End Sub

'Modified by Lydia 2021/11/01 改成Form 2.0
'Private Sub txtDB_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtDB_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index <> 2 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtDB_Validate(Index As Integer, Cancel As Boolean)
   Dim strCusTemp As String, strTemp As String
   Select Case Index
      Case 3 '本所案號
         If txtDB(Index) <> "" Then
            strExc(0) = "select PA01||PA02||PA03||PA04 from patent where " & ChgPatent(txtDB(Index))
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               If m_EditMode <> 0 Then 'Added by Lydia 2021/11/01 排除非編輯模式: 因為案號有可能刪除
                    MsgBox "本所案號輸入錯誤!", vbExclamation
                    Cancel = True
               End If 'Added by Lydia 2021/11/01
               'If m_EditMode <> 0 Then Cancel = True 'Remove by Lydia 2021/11/01
            Else
               txtDB(Index) = RsTemp(0)
            End If
         End If
      Case 4 '代理人
         Label2(1).Caption = ""
         If txtDB(Index) <> "" Then
            'Modified by Morgan 2019/7/25 加碼數檢查
            If Len(txtDB(Index)) = 6 Or Len(txtDB(Index)) = 8 Then
               strCusTemp = ChangeCustomerL(txtDB(Index))
               If ClsPDGetAgent(strCusTemp, strTemp) Then
                  Label2(1).Caption = strTemp
                  'Added by Lydia 2023/02/18 整合特殊備註維護：在輸入Y/X編號若為6碼，統一補足為8碼。
                  If m_EditMode <> 0 Then
                     txtDB(Index) = Left(ChangeCustomerL(txtDB(Index)), 8)
                  End If
                  'end 2023/02/18
               Else
                  'MsgBox "代理人編號輸入錯誤！", vbCritical 'Remove by Lydia 2021/11/01 模組已彈訊息
                  If m_EditMode <> 0 Then Cancel = True
               End If
            Else
               MsgBox "代理人編號只可輸入6碼或8碼！", vbCritical
               If m_EditMode <> 0 Then Cancel = True
            End If
         End If
      Case 5 '申請人
         Label2(2).Caption = ""
         If txtDB(Index) <> "" Then
            'Modified by Morgan 2019/7/25 加碼數檢查
            If Len(txtDB(Index)) = 6 Or Len(txtDB(Index)) = 8 Then
               strCusTemp = ChangeCustomerL(txtDB(Index))
               If ClsPDGetCustomer(strCusTemp, strTemp) Then
                  Label2(2).Caption = strTemp
                  'Added by Lydia 2023/02/18 整合特殊備註維護：在輸入Y/X編號若為6碼，統一補足為8碼。
                  If m_EditMode <> 0 Then
                     txtDB(Index) = Left(ChangeCustomerL(txtDB(Index)), 8)
                  End If
                  'end 2023/02/18
               Else
                  'MsgBox "客戶編號輸入錯誤！", vbCritical  'Remove by Lydia 2021/11/01 模組已彈訊息
                  If m_EditMode <> 0 Then Cancel = True
               End If
            Else
               MsgBox "客戶編號只可輸入6碼或8碼！", vbCritical
               If m_EditMode <> 0 Then Cancel = True
            End If
         End If
      'Added by Lydia 2023/08/01 案件性質
      Case 12 '案件性質
         Label2(0).Caption = ""
         If txtDB(Index) <> "" Then
            If InStr(mPtyList & ",", txtDB(Index) & ",") = 0 Then
               MsgBox "目前僅開放設定可產生請款通知函的案件性質！", vbExclamation
               Cancel = True
            Else
               If ClsPDGetCaseProperty("FCP", txtDB(Index), strTemp) Then
                  Label2(0).Caption = strTemp
               Else
                  If m_EditMode <> 0 Then Cancel = True
               End If
            End If
         End If
      'end 2023/08/01
   End Select
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean, idx As Integer
   
   If txtDB(2) = "" Then
      'Modified by Morgan 2015/12/16 不必限制,因有個案需要設定不待備註
      'MsgBox "備註內容不可空白！", vbExclamation
      If MsgBox("備註內容未輸入是否確定為請款函不要預設備註？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         txtDB(2).SetFocus
         Exit Function
      End If
   End If
   
   If txtDB(3) & txtDB(4) & txtDB(5) = "" Then
      MsgBox "請輸入本所案號、代理人或申請人！", vbExclamation
      txtDB(3).SetFocus
      Exit Function
   End If
   
   For idx = 3 To 5
      txtDB_Validate idx, bCancel
      If bCancel = True Then
         txtDB(idx).SetFocus
         Exit Function
      End If
   Next
   
   'Added by Lydia 2021/11/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
   If m_EditMode = 1 Then
      'Modified by Lydia 2017/01/11 避免LOG語法分析錯誤
      'strSql = "insert into DebitNotePS(DNPS01,DNPS02,DNPS03,DNPS04,DNPS05)" & _
         " select nvl(max(DNPS01),0)+1 DNPS01,'" & ChgSQL(txtDB(2)) & "' DNPS02" & _
         ",'" & txtDB(3) & "' DNPS03,'" & txtDB(4) & "' DNPS04,'" & txtDB(5) & "' DNPS05 from DebitNotePS "
      'Modified by Lydia 2023/08/01 +案件性質DNPS12
      strSql = "insert into DebitNotePS(DNPS01,DNPS02,DNPS03,DNPS04,DNPS05,DNPS12) VALUES " & _
               "('" & Pub_GetDefColMaxNo("DebitNotePS", "DNPS01") & "','" & ChgSQL(txtDB(2)) & "','" & txtDB(3) & "' ,'" & txtDB(4) & "' ,'" & txtDB(5) & "' ,'" & ChgSQL(txtDB(12)) & "') "
   Else
      'Modified by Lydia 2023/08/01 +案件性質DNPS12
      strSql = "update DebitNotePS set DNPS02='" & ChgSQL(txtDB(2)) & "', DNPS03='" & txtDB(3) & "'" & _
         ", DNPS04='" & txtDB(4) & "', DNPS05='" & txtDB(5) & "', DNPS12='" & ChgSQL(txtDB(12)) & "'" & _
         " where DNPS01=" & txtDB(1)
   End If
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

Private Sub GRD1_DblClick()
   'Modified by Lydia 2021/11/11 因為加上Grid排序,所以改寫法
   'If GRD1.row > 0 And GRD1.TextMatrix(GRD1.row, 0) <> "" Then
   '   ReadData GRD1.TextMatrix(GRD1.row, 0)
   'End If
Dim intRow As Integer
   With grd1
       If .MouseRow > 0 Then
          intRow = .MouseRow
          .row = intRow
          If .row > 0 And .TextMatrix(intRow, 0) <> "" Then
              ReadData .TextMatrix(intRow, 0)
          End If
       End If
   End With
'end 2021/11/11
End Sub

'Added by Lydia 2021/11/11
Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow grd1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   grd1.col = nCol
   grd1.row = nRow
   If Me.grd1.row < 1 And Me.grd1.Text <> "V" Then
      If InStr("流水號,", Me.grd1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.grd1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.grd1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.grd1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.grd1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub
Private Function FormDelete() As Boolean
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   strSql = "delete from DebitNotePS where DNPS01=" & txtDB(1)
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   FormDelete = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

Private Function RecIsExist() As Boolean

stCon = ""
If Trim(txtDB(3)) <> "" Then
   stCon = stCon & "and dnps03='" & Trim(txtDB(3)) & "' "
End If
If Trim(txtDB(4)) <> "" Then
   'Modified by Lydia 2019/07/31 改成=判斷; 因為無法先輸入8碼後再輸入6碼
   'stcon = stcon & "and instr(dnps04,'" & Trim(txtDB(4)) & "') > 0 "
   stCon = stCon & "and dnps04='" & Trim(txtDB(4)) & "' "
   'Added by Lydia 2019/03/04 區別只有代理人或客戶的條件
   If Trim(txtDB(5)) = "" Then stCon = stCon & "and dnps05 is null "
End If
If Trim(txtDB(5)) <> "" Then
   'Modified by Lydia 2019/07/31 改成=判斷; 因為無法先輸入8碼後再輸入6碼
   'stcon = stcon & "and instr(dnps05,'" & Trim(txtDB(5)) & "') > 0 "
    stCon = stCon & "and dnps05='" & Trim(txtDB(5)) & "' "
   'Added by Lydia 2019/03/04 區別只有代理人或客戶的條件
   If Trim(txtDB(4)) = "" Then stCon = stCon & "and dnps04 is null "
End If
'Added by Lydia 2023/08/01 案件性質
If Trim(txtDB(12)) <> "" Then
   stCon = stCon & "and dnps12='" & Trim(txtDB(12)) & "' "
End If


'Modified by Lydia 2019/03/11 若新增時在備註欄按Enter會跳到檢查
'If Left(stcon, 3) = "and" Then stcon = Mid(stcon, 4, Len(stcon) - 4)
If Left(stCon, 3) = "and" Then
    stCon = Mid(stCon, 4, Len(stCon) - 4)
ElseIf stCon = "" Then
    Exit Function
End If
'end 2019/03/11

   strExc(1) = " select * from DebitNotePS where " & stCon
   intR = 1
   Set rsRead = ClsLawReadRstMsg(intR, strExc(1))
   If intR = 1 Then
      'Added by Lydia 2021/11/01 排除現在修改的記錄
      If rsRead.RecordCount = 1 And Trim(rsRead.Fields("DNPS01")) = Trim(txtDB(1)) Then
          RecIsExist = False
      Else
      'end 2021/11/01
          RecIsExist = True
          MsgBox "已存在同樣條件的記錄(流水號 " & rsRead(0) & " )，請先查詢!!", vbCritical
      End If 'Added by Lydia 2021/11/01
   Else
      RecIsExist = False
   End If
   Set rsRead = Nothing
   
End Function

'Added by Lydia 2021/11/01
Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    If Index <> 3 Then 'Added by Lydia 2022/10/03
        KeyAscii = UpperCase(KeyAscii)
    End If
End Sub

Private Sub txtFM2_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp As String

   Select Case Index
   Case 0 '本所案號
      txtFM2(1).Tag = "": txtFM2(2).Tag = ""   '清空本所案號的相關Y編號、X編號條件
      If txtFM2(Index) <> "" Then
         strExc(0) = "select PA01||PA02||PA03||PA04,PA75, PA26||','||PA27||','||PA28||','||PA29||','||PA30 AS appno from patent where " & ChgPatent(txtFM2(Index))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            MsgBox "本所案號輸入錯誤!", vbExclamation
         Else
            txtFM2(Index) = RsTemp(0)
            txtFM2(1).Tag = "" & RsTemp.Fields("pa75")
            txtFM2(2).Tag = "" & RsTemp.Fields("appno")
         End If
      End If
   Case 1 '代理人
      lblFM2(Index).Caption = ""
      If txtFM2(Index) <> "" Then
         If Len(txtFM2(Index)) = 6 Or Len(txtFM2(Index)) = 8 Then
            stCon = Left(txtFM2(Index) & "000", 9)
            If ClsPDGetAgent(stCon, strTemp) Then
               lblFM2(1).Caption = strTemp
            Else
               '模組已彈訊息
            End If
         End If
      End If
   Case 2 '申請人
      lblFM2(Index).Caption = ""
      If txtFM2(Index) <> "" Then
         If Len(txtFM2(Index)) = 6 Or Len(txtFM2(Index)) = 8 Then
            stCon = Left(txtFM2(Index) & "000", 9)
            If ClsPDGetCustomer(stCon, strTemp) Then
               lblFM2(2).Caption = strTemp
            Else
               '模組已彈訊息
            End If
         End If
      End If
   'Added by Lydia 2023/08/01
   Case 4 '案件性質
      lblFM2(0).Caption = ""
      If txtFM2(Index) <> "" Then
         If InStr(mPtyList & ",", txtFM2(Index) & ",") = 0 Then
            MsgBox "目前僅開放設定可產生請款通知函的案件性質！", vbExclamation
            Cancel = True
         Else
            If ClsPDGetCaseProperty("FCP", txtFM2(Index), strTemp) Then
               lblFM2(0).Caption = strTemp
            Else
               '模組已彈訊息
            End If
         End If
      End If
   End Select
End Sub
