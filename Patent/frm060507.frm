VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060507 
   BorderStyle     =   1  '單線固定
   Caption         =   "核駁及審查意見通知函備註"
   ClientHeight    =   6660
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9732
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9732
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
            Picture         =   "frm060507.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060507.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060507.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060507.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060507.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060507.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060507.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060507.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060507.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060507.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060507.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9732
      _ExtentX        =   17166
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
      Height          =   5916
      Left            =   24
      TabIndex        =   16
      Top             =   720
      Width           =   9696
      _ExtentX        =   17103
      _ExtentY        =   10435
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm060507.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
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
      Tab(0).Control(8)=   "Label1(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label3(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textCUID"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label2(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label2(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtIM(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtIM(6)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtIM(5)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtIM(4)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtIM(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtIM(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtIM(7)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtIM(8)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtIM(9)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtIM(11)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label1(16)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label1(8)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtIM(10)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtIM(18)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label3(3)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label5(0)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label5(1)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label5(2)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label5(3)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label1(17)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label1(18)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtIM(19)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label1(19)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label2(4)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtIM(20)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Label1(23)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Label1(24)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Label1(25)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtIM(21)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Label1(26)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Label1(27)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "CmdMsg"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "cmdPS"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).ControlCount=   52
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm060507.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblFM2(2)"
      Tab(1).Control(1)=   "lblFM2(1)"
      Tab(1).Control(2)=   "txtFM2(3)"
      Tab(1).Control(3)=   "txtFM2(2)"
      Tab(1).Control(4)=   "txtFM2(1)"
      Tab(1).Control(5)=   "txtFM2(0)"
      Tab(1).Control(6)=   "Label1(12)"
      Tab(1).Control(7)=   "Label1(13)"
      Tab(1).Control(8)=   "Label1(14)"
      Tab(1).Control(9)=   "Label1(15)"
      Tab(1).Control(10)=   "lblPS"
      Tab(1).Control(11)=   "Label1(20)"
      Tab(1).Control(12)=   "txtFM2(4)"
      Tab(1).Control(13)=   "lblFM2(4)"
      Tab(1).Control(14)=   "Label1(21)"
      Tab(1).Control(15)=   "Label1(22)"
      Tab(1).Control(16)=   "txtFM2(5)"
      Tab(1).Control(17)=   "GRD1"
      Tab(1).Control(18)=   "cmdQuery"
      Tab(1).ControlCount=   19
      Begin VB.CommandButton cmdPS 
         Caption         =   "其他例外情況說明"
         Height          =   495
         Left            =   396
         TabIndex        =   54
         Top             =   4788
         Width           =   1005
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   300
         Left            =   -72270
         TabIndex        =   19
         Top             =   390
         Width           =   885
      End
      Begin VB.CommandButton CmdMsg 
         BackColor       =   &H00C0FFC0&
         Caption         =   "規則說明"
         Height          =   360
         Left            =   6516
         Style           =   1  '圖片外觀
         TabIndex        =   39
         Top             =   5268
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm060507.frx":212C
         Height          =   3675
         Left            =   -74880
         TabIndex        =   17
         Top             =   2130
         Width           =   9375
         _ExtentX        =   16531
         _ExtentY        =   6477
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "流水號|備註內容|本所案號|代理人|申請人|C類性質|C類承辦|系統別|IM06| B類收文|B類承辦"
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
         _Band(0).Cols   =   11
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "空白: 依設定規則讀取)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   27
         Left            =   7776
         TabIndex        =   72
         Top             =   2904
         Width           =   1776
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y: 新增,  N: 確定不新增"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   26
         Left            =   7728
         TabIndex        =   71
         Top             =   2664
         Width           =   1896
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   21
         Left            =   7296
         TabIndex        =   14
         Top             =   4116
         Width           =   372
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "起算：          (1.系統日 2.官方發文日)"
         Height          =   180
         Index           =   25
         Left            =   6720
         TabIndex        =   70
         Top             =   4128
         Width           =   2868
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " (1.系統日 2.官方發文日)"
         Height          =   180
         Index           =   24
         Left            =   3720
         TabIndex        =   69
         Top             =   4050
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "期限起算類型："
         Height          =   180
         Index           =   23
         Left            =   2070
         TabIndex        =   68
         Top             =   4050
         Width           =   1260
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   20
         Left            =   3330
         TabIndex        =   8
         Top             =   3990
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   520
         Index           =   5
         Left            =   -73860
         TabIndex        =   24
         Top             =   1320
         Width           =   6690
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "11800;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "模糊比對"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   22
         Left            =   -74850
         TabIndex        =   67
         Top             =   1590
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註內容："
         Height          =   180
         Index           =   21
         Left            =   -74880
         TabIndex        =   66
         Top             =   1380
         Width           =   900
      End
      Begin MSForms.Label lblFM2 
         Height          =   255
         Index           =   4
         Left            =   -67200
         TabIndex        =   65
         Top             =   405
         Width           =   1305
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "2302;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   285
         Index           =   4
         Left            =   -67860
         TabIndex        =   21
         Top             =   390
         Width           =   615
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1085;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶C類來函性質："
         Height          =   180
         Index           =   20
         Left            =   -69420
         TabIndex        =   64
         Top             =   442
         Width           =   1560
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   2370
         TabIndex        =   63
         Top             =   3360
         Width           =   1785
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "3149;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶C類來函性質："
         Height          =   180
         Index           =   19
         Left            =   60
         TabIndex        =   62
         Top             =   3390
         Width           =   1560
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   19
         Left            =   1710
         TabIndex        =   5
         Top             =   3330
         Width           =   615
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(計算客戶指定送件日期)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   18
         Left            =   60
         TabIndex        =   61
         Top             =   3990
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(計算本所期限)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   17
         Left            =   7776
         TabIndex        =   60
         Top             =   3516
         Width           =   1200
      End
      Begin MSForms.Label Label5 
         Height          =   192
         Index           =   3
         Left            =   1968
         TabIndex        =   59
         Top             =   4512
         Width           =   1092
         ForeColor       =   16711680
         BackColor       =   12648447
         Caption         =   "指定送件日期"
         Size            =   "1931;344"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label5 
         Height          =   192
         Index           =   2
         Left            =   3888
         TabIndex        =   58
         Top             =   4800
         Width           =   732
         ForeColor       =   16711680
         BackColor       =   12648447
         Caption         =   "本所期限"
         Size            =   "1291;339"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label5 
         Height          =   192
         Index           =   1
         Left            =   2448
         TabIndex        =   57
         Top             =   4800
         Width           =   1092
         ForeColor       =   16711680
         BackColor       =   12648447
         Caption         =   "指定送件日期"
         Size            =   "1931;344"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label5 
         Height          =   192
         Index           =   0
         Left            =   4680
         TabIndex        =   56
         Top             =   4512
         Width           =   732
         ForeColor       =   16711680
         BackColor       =   12648447
         Caption         =   "本所期限"
         Size            =   "1291;339"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "承辦期限=　　　　　　  或　　　　　-2個工作天。"
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
         Index           =   3
         Left            =   1476
         TabIndex        =   55
         Top             =   4788
         Width           =   4440
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   18
         Left            =   2955
         TabIndex        =   7
         Top             =   3660
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   10
         Left            =   7296
         TabIndex        =   13
         Top             =   3780
         Width           =   372
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "類型：          (1.日曆天 2.工作天)"
         Height          =   180
         Index           =   8
         Left            =   6720
         TabIndex        =   53
         Top             =   3840
         Width           =   2508
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "類型：         (1.日曆天 2.工作天)"
         Height          =   180
         Index           =   16
         Left            =   2445
         TabIndex        =   52
         Top             =   3720
         Width           =   2460
      End
      Begin VB.Label lblPS 
         Caption         =   "P.S. 輸入本所案號會另外帶該案代理人和申請人的其他設定"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74910
         TabIndex        =   51
         Top             =   1920
         Width           =   4845
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   11
         Left            =   2040
         TabIndex        =   6
         Top             =   3660
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   9
         Left            =   7290
         TabIndex        =   10
         Top             =   2670
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   8
         Left            =   7290
         TabIndex        =   9
         Top             =   2340
         Width           =   615
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   7
         Left            =   7296
         TabIndex        =   12
         Top             =   3456
         Width           =   372
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   1
         Left            =   1050
         TabIndex        =   0
         Top             =   390
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
      Begin MSForms.TextBox txtIM 
         Height          =   1575
         Index           =   2
         Left            =   1050
         TabIndex        =   1
         Top             =   720
         Width           =   8295
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "14631;2778"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   4
         Left            =   1050
         TabIndex        =   3
         Top             =   2670
         Width           =   1170
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2064;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   5
         Left            =   1050
         TabIndex        =   4
         Top             =   3000
         Width           =   1170
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2064;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   6
         Left            =   7296
         TabIndex        =   11
         Top             =   3120
         Width           =   612
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtIM 
         Height          =   300
         Index           =   3
         Left            =   1050
         TabIndex        =   2
         Top             =   2340
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
         Height          =   252
         Index           =   3
         Left            =   7968
         TabIndex        =   50
         Top             =   3168
         Width           =   1152
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "2037;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   49
         Top             =   2693
         Width           =   3195
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "5636;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   48
         Top             =   3030
         Width           =   3195
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "5644;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID 
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   5550
         Width           =   7860
         VariousPropertyBits=   671105055
         Size            =   "13864;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         Height          =   180
         Index           =   15
         Left            =   -74880
         TabIndex        =   46
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   14
         Left            =   -74880
         TabIndex        =   45
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "系統別："
         Height          =   180
         Index           =   13
         Left            =   -71160
         TabIndex        =   44
         Top             =   435
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   12
         Left            =   -74880
         TabIndex        =   43
         Top             =   435
         Width           =   900
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   285
         Index           =   0
         Left            =   -73860
         TabIndex        =   18
         Top             =   390
         Width           =   1515
         VariousPropertyBits=   671105051
         MaxLength       =   12
         Size            =   "2672;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   285
         Index           =   1
         Left            =   -73860
         TabIndex        =   22
         Top             =   705
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
         Index           =   2
         Left            =   -73860
         TabIndex        =   23
         Top             =   1020
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
         Index           =   3
         Left            =   -70410
         TabIndex        =   20
         Top             =   390
         Width           =   615
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1085;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   255
         Index           =   1
         Left            =   -72720
         TabIndex        =   42
         Top             =   720
         Width           =   5595
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "9878;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   255
         Index           =   2
         Left            =   -72720
         TabIndex        =   41
         Top             =   1035
         Width           =   5595
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "9878;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "　　　2.因為核駁及審查意見通知函備註可顯示多筆備註，所以有"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   4
         Left            =   876
         TabIndex        =   40
         Top             =   5292
         Width           =   5592
      End
      Begin VB.Label Label4 
         Caption         =   "期限皆不包含當天，若使用日曆天的期限遇到非工作日往前移到工作天。"
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
         Height          =   252
         Left            =   1476
         TabIndex        =   38
         Top             =   5028
         Width           =   6768
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "注意：1.C類來函的　　　　　　  和B類內部收文的　　　　  =系統日+X個(1.日曆天，2.工作天)，"
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
         Index           =   2
         Left            =   312
         TabIndex        =   37
         Top             =   4512
         Width           =   8316
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否新增內部收文："
         Height          =   180
         Index           =   11
         Left            =   5610
         TabIndex        =   36
         Top             =   2715
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "系統別："
         Height          =   180
         Index           =   10
         Left            =   6450
         TabIndex        =   35
         Top             =   2385
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶C類來函承辦天數："
         Height          =   180
         Index           =   9
         Left            =   60
         TabIndex        =   34
         Top             =   3720
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦天數："
         Height          =   180
         Index           =   7
         Left            =   6348
         TabIndex        =   33
         Top             =   3516
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "模糊比對"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   6
         Left            =   6750
         TabIndex        =   32
         Top             =   450
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "流水號："
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   31
         Top             =   435
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註內容："
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   30
         Top             =   810
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   29
         Top             =   2715
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   3
         Left            =   285
         TabIndex        =   28
         Top             =   3045
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   180
         Index           =   4
         Left            =   6348
         TabIndex        =   27
         Top             =   3168
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   5
         Left            =   105
         TabIndex        =   26
         Top             =   2385
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "注意：1.代理人/申請人可輸入6碼或8碼，6碼代表含關係企業，無論6碼或8碼均包含更名前編號。112/1/9  取消"
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   9255
      End
   End
End
Attribute VB_Name = "frm060507"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/02/18 整合特殊備註維護：1.將現有資料6碼Y/X編號補足為8碼；2.在輸入Y/X編號若為6碼，統一補足為8碼。
'Memo by Lydia 2021/11/01 改成Form2.0 ; GRD1改字型=新細明體-ExtB、txtIM(index)、Label2(index)、textCUID、txtFM2(index)、lblFM2(index)
'Memo by Lydia 2021/11/01 原本IM11(說明)暫存原程式PUB_ChkAutoRec的iType分類，現在改為「客戶C類來函承辦天數」
'Memo by Lydia 2021/11/01 畫面頁籤改成「單筆資料」和「多筆查詢」：上方工具列的「查詢」帶出第一筆符合的資料，在多筆查詢的頁籤可以輸入條件進行查詢，並且在下方的Grid呈現多筆資料。
'Created by Lydia 2014/12/03 核駁及審查意見通知函備註
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
'Remove by Lydia 2018/02/14  開放輸入案件性質
'Dim mInPty As String, mInPtyStr As String
'Remove by Lydia 2018/02/14 開放輸入案件性質
'Private Const mInPty = "901,924" '案件性質
'Private Const mInPtyStr = "告知代理人(901)、會稿(924)"    '案件性質message
'end remove

'Added by Lydia 2021/11/01 減少畫面的說明
Private Sub cmdPS_Click()
   'Modified by Lydia 2022/10/13 改成欄位設定IM20,IM21
   'MsgBox "1.當系統別為P案且代理人為Y20065時，預設承辦期限＝官方發文日＋承辦天數。" & vbCrLf & _
                "    其他案件或代理人，預設承辦期限＝系統日＋承辦天數。" & vbCrLf & vbCrLf & _
                "2.先正達來函下一程序案件性質'掛804或501-509時C類接洽單另有特殊備註。", vbInformation + vbOKOnly, "其他例外情況說明"
   MsgBox "先正達來函下一程序案件性質掛804或501-509時C類接洽單另有特殊備註。", vbInformation + vbOKOnly, "其他例外情況說明"
End Sub

'Added by Lydia 2021/11/01
Private Sub cmdQuery_Click()
   
   stCon = ""
   If txtFM2(0) <> "" Then
      If Trim(txtFM2(1).Tag & txtFM2(2).Tag) = "" Then
          stCon = stCon & " and im03='" & txtFM2(0) & "'"
      Else
          '另外抓本所案號的相關Y編號、X編號條件
          stCon = stCon & " and (im03='" & txtFM2(0) & "'"
          If txtFM2(1).Tag <> "" Then stCon = stCon & " or instr(" & CNULL(txtFM2(1).Tag) & ", im04) > 0 "
          If txtFM2(2).Tag <> "" Then stCon = stCon & " or instr(" & CNULL(txtFM2(2).Tag) & ", im05) > 0 "
          stCon = stCon & ") "
      End If
   Else
      txtFM2(1).Tag = "": txtFM2(2).Tag = ""   '清空本所案號的相關Y編號、X編號條件
   End If
   If txtFM2(1) <> "" Then
      stCon = stCon & " and im04 like '" & txtFM2(1) & "%'"
   End If
   If txtFM2(2) <> "" Then
      stCon = stCon & " and im05 like '" & txtFM2(2) & "%'"
   End If
   If txtFM2(3) <> "" Then
      stCon = stCon & " and im08 = '" & txtFM2(3) & "'"
   End If
   'Added by Lydia 2021/11/11 客戶C類來函性質
   If txtFM2(4) <> "" Then
      stCon = stCon & " and im19 = '" & txtFM2(4) & "'"
   End If
   'end 2021/11/11
   'Added by Lydia 2022/10/03 增加"備註"查詢
   If txtFM2(5) <> "" Then
       stCon = stCon & " and upper(im02) like '%" & ChgSQL(UCase(txtFM2(5))) & "%' "
   End If
   'end 2022/10/03
   
   'Modified by Lydia 2021/11/11
   'stSQL = "SELECT IM01,substr(IM02,1,500) as IM02,IM03,IM04,IM05,IM11,IM08,IM06,CPM03,IM07 " & _
                "FROM INCOMMEMO,CASEPROPERTYMAP WHERE 'FCP'=CPM01(+) AND IM06=CPM02(+) " & stCon
   stSQL = "SELECT IM01,substr(IM02,1,500) as IM02,IM03,IM04,IM05,c1.CPM03 as ccpm03,IM11,IM08,IM06,c2.CPM03 as bcpm03,IM07 " & _
                "FROM INCOMMEMO,CASEPROPERTYMAP c1,CASEPROPERTYMAP c2 " & _
                "WHERE 'FCP'=c1.CPM01(+) AND IM19=c1.CPM02(+) AND 'FCP'=c2.CPM01(+) AND IM06=c2.CPM02(+) " & stCon
   stSQL = stSQL & " ORDER BY IM01"
   intR = 0
   Set rsRead = ClsLawReadRstMsg(intR, stSQL)
   
   Call SetGrd(True)
   If intR = 1 Then
        GRD1.FixedCols = 0
        Set GRD1.Recordset = rsRead
        Call SetGrd
        GRD1.FixedCols = 5
   End If
End Sub

'Added by Lydia 2021/11/01
Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer, iR As Integer
   
   'Modified by Lydia 2021/11/11 +C類性質IM19=>CCPM03
   arrGridHeadText = Array("流水號", "備註內容", "本所案號", "代理人", "申請人", "C類性質", "C類承辦", "系統別", "IM06", " B類收文", "B類承辦")
   arrGridHeadWidth = Array(800, 1200, 1200, 1000, 1000, 1000, 800, 720, 0, 1000, 800)
   
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
        GRD1.Clear
        GRD1.Rows = 2
   End If
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   
   'Added by Lydia 2021/11/11
   For iR = 1 To GRD1.Rows - 1
      For iRow = 0 To GRD1.Cols - 1
         GRD1.row = iR
         GRD1.col = iRow
         If InStr("06,10", Format(iRow, "00")) > 0 Then '置中
            GRD1.CellAlignment = flexAlignCenterCenter
         End If
      Next iRow
   Next iR
   'end 2021/11/11
   
   GRD1.Visible = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Memo by Lydia 2021/11/01 原程式搬到Form_KeyUp

End Sub

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
         If m_EditMode = 4 Then
            KeyCode = 0: Action 6
         End If
      Case vbKeyPageUp '上一筆
         If m_EditMode = 4 Then
            KeyCode = 0: Action 7
         End If
      Case vbKeyPageDown '下一筆
         If m_EditMode = 4 Then
            KeyCode = 0: Action 8
         End If
      Case vbKeyEnd: '最後筆
         If m_EditMode = 4 Then
            KeyCode = 0: Action 9
         End If
      'Modified by Lydia 2021/11/22 Lydia 2021/11/22 取消以ENTER控制為換行的功能 (Form2.0修改之維護資料功能Toolbar之修改統一)
      'Case vbKeyF9, vbKeyReturn '確定
      Case vbKeyF9
         'Modified by Lydia 2021/11/22 取消
         'If Me.ActiveControl <> txtIM(2) Then KeyCode = 0: Action 11
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
   m_bInsert = IsUserHasRightOfFunction("frm060507", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm060507", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm060507", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm060507", strFind, False)
  
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
   Set frm060507 = Nothing
End Sub

'Added by Lydia 2021/11/11
Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If InStr("流水號,C類承辦,B類承辦", Me.GRD1.Text) > 0 Then
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
         If m_bUpdate And txtIM(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtIM(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtIM(1) <> "" Then
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
      For Each oText In txtIM
         oText.Locked = True
      Next
      SSTab1.TabEnabled(1) = True
   Case Else
      For Each oText In txtIM
         oText.Locked = False
      Next
      If m_EditMode <> 4 Then
         txtIM(1).Locked = True
         txtIM(2).SetFocus
         txtIM_GotFocus 2
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
         If txtIM(1).Text = "" Then
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
         txtIM(1).Enabled = True
         txtIM(1).SetFocus
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
         If Val(m_EditMode) > 0 And Trim(txtIM(3)) <> "" And ((Left(Trim(txtIM(3)), 1) = "P" And Len(Trim(txtIM(3))) < 10) Or (Left(Trim(txtIM(3)), 3) = "FCP" And Len(Trim(txtIM(3))) < 12)) Then
             Call txtIM_Validate(3, bCancel)
             If bCancel = True Then
                 Exit Sub
             End If
         End If
         
         Select Case m_EditMode
            '新增,修改
            Case 1, 2
               If TxtValidate = False Then
                  Exit Sub
               Else
                  'Added by Lydia 2015/01/07 判斷是否存在相同條件的記錄
                  'Modified by Lydia 2021/11/01 新增,修改都要判斷
                  'If m_EditMode = 1 Then
                  '   If RecIsExist = True Then Exit Sub
                  'End If
                  ''end 2015/01/07
                  If RecIsExist = True Then Exit Sub
                  
                    If FormSave() = False Then
                       MsgBox "存檔失敗!", vbCritical
                       Exit Sub
                    Else
                       strKind = m_EditMode 'Added by Lydia 2021/11/01 記錄新增模式
                       m_EditMode = 0
                       'Modified by Morgan 2017/9/13
                       'If m_EditMode = 1 Then
                       If txtIM(1) = "" Then
                       'end 2017/9/13
                          ShowRecord 3
                       Else
                          ReadData txtIM(1)
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
                           If txtIM(3) <> "" Then
                               txtFM2(0) = txtIM(3)
                               Call txtFM2_Validate(0, False)
                           Else
                               If txtIM(4) <> "" Then
                                  txtFM2(1) = ChangeCustomerS(txtIM(4))
                                  Call txtFM2_Validate(1, False)
                               End If
                               If txtIM(5) <> "" Then
                                  txtFM2(2) = ChangeCustomerS(txtIM(5))
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
               If ReadData(txtIM(1)) = False Then
                  MsgBox "無資料!", vbExclamation
                  Exit Sub
               Else
                  m_EditMode = 0
               End If
         End Select
      Case 12 '按下取消
         m_EditMode = 0
         txtIM(1) = txtIM(1).Tag
         If txtIM(1) <> "" Then
            If ReadData(txtIM(1)) = False Then
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
         strExc(0) = "SELECT nvl(min(im01),0) FROM IncomMemo"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKey = RsTemp.Fields(0)
            End If
         End If
         
      Case 1 '前一筆
         strExc(0) = "SELECT nvl(max(im01),0) FROM IncomMemo where im01<" & txtIM(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 6
            Else
               stKey = RsTemp.Fields(0)
            End If
         End If
         
      Case 2 '後一筆
         strExc(0) = "SELECT nvl(min(im01),0) FROM IncomMemo where im01>" & txtIM(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 7
            Else
               stKey = RsTemp.Fields(0)
            End If
         End If
         
      Case 3 '最後筆
         strExc(0) = "SELECT nvl(max(im01),0) FROM IncomMemo"
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
   
   stCon = ""
   '單筆
   If pKey <> "" Then
      stCon = " and im01=" & pKey
   '多筆
   Else
      If txtIM(2) <> "" Then
         'Modified by Morgan 2017/9/13
         'stCon = stCon & " and im02 like '%" & txtIM(2) & "%'"
         stCon = stCon & " and im02 like '%" & ChgSQL(txtIM(2)) & "%'"
      End If
      If txtIM(3) <> "" Then
         stCon = stCon & " and im03='" & txtIM(3) & "'"
      End If
      If txtIM(4) <> "" Then
         stCon = stCon & " and im04 like '" & txtIM(4) & "%'"
      End If
      If txtIM(5) <> "" Then
         stCon = stCon & " and im05 like '" & txtIM(5) & "%'"
      End If
      For intI = 6 To 11
          If txtIM(intI) <> "" Then
             stCon = stCon & " and im" & Format(intI, "00") & "='" & txtIM(intI) & "'"
          End If
      Next intI
   End If
   
   FormReset
   
   strExc(0) = "select * from IncomMemo where 1=1 " & stCon & " order by im01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If m_EditMode = 4 Then
         'Remove by Lydia 2021/11/01 改成單筆查詢
         'Set GRD1.Recordset = RsTemp.Clone
         'GRD1.FormatString = GRD1.FormatString
         'GRD1.ColWidth(1) = 2775
         'GRD1.ColWidth(2) = 1290
         'GRD1.ColWidth(3) = 1020
         'GRD1.ColWidth(4) = 1020
         'For intI = 5 To 11
         '   GRD1.ColWidth(intI) = 810
         'Next
         'For intI = 11 To GRD1.Cols - 1
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
   For Each oText In txtIM
      oText = "" & .Fields("im" & Format(oText.Index, "00"))
   Next
   End With
   UpdateCUID rsQuery
   
   txtIM(1).Tag = txtIM(1)
   If txtIM(4) <> "" Then txtIM_Validate 4, False
   If txtIM(5) <> "" Then txtIM_Validate 5, False
   If txtIM(6) <> "" Then txtIM_Validate 6, False
   If txtIM(7) <> "" Then txtIM_Validate 7, False
   If txtIM(8) <> "" Then txtIM_Validate 8, False
   If txtIM(9) <> "" Then txtIM_Validate 9, False
   If txtIM(19) <> "" Then txtIM_Validate 19, False  'Added by Lydia 2021/11/19
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
   If IsNull(rsSrcTmp.Fields("im12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("im12")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("im12"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("im13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("im13")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("im13"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("im14")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("im14")) = False Then
         strTemp = rsSrcTmp.Fields("im14")
         strCTime = Format(strTemp, "00:00:00")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("im15")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("im15")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("im15"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("im16")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("im16")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("im16"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("im17")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("im17")) = False Then
         strTemp = rsSrcTmp.Fields("im17")
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

   For Each oText In txtIM
      oText.Text = ""
   Next
   
   For Each oLabel In Label2
      oLabel.Caption = ""
   Next
   
   textCUID = ""
   Label1(6).Visible = False
End Sub

Private Sub txtIM_GotFocus(Index As Integer)
   TextInverse txtIM(Index)
   If Index = 2 Then
      OpenIme
   Else
      CloseIme
   End If
End Sub

'Modified by Lydia 2021/11/01 改成Form 2.0
'Private Sub txtIM_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtIM_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   'Added by Lydia 2021/11/01  承辦天數
   'Modified by Lydia 2021/11/01 +承辦期限類型IM10,IM18
   'Modified by Lydia 2021/11/11 +IM19 客戶C類來函性質
   'Modified by Lydia 2022/10/13 +IM20, IM21
   If Index = 7 Or Index = 11 Or Index = 10 Or Index = 18 Or Index = 19 Or Index = 20 Or Index = 21 Then
      KeyAscii = Pub_NumAscii(KeyAscii)
   Else
   'end 2021/11/01
      If Index <> 2 Then
         KeyAscii = UpperCase(KeyAscii)
      End If
   End If 'Added by Lydia 2021/11/01
End Sub

Private Sub txtIM_Validate(Index As Integer, Cancel As Boolean)
   Dim strCusTemp As String, strTemp As String
   Select Case Index
   Case 3 '本所案號
      If txtIM(Index) <> "" Then
         strExc(0) = "select PA01||PA02||PA03||PA04 from patent where " & ChgPatent(txtIM(Index))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            If m_EditMode <> 0 Then 'Added by Lydia 2021/11/01 排除非編輯模式: 因為案號有可能刪除
                 MsgBox "本所案號輸入錯誤!", vbExclamation
                 Cancel = True
            End If 'Added by Lydia 2021/11/01
            'If m_EditMode <> 0 Then Cancel = True 'Remove by Lydia 2021/11/01
         Else
            txtIM(Index) = RsTemp(0)
         End If
      End If
   Case 4 '代理人
      Label2(1).Caption = ""
      If txtIM(Index) <> "" Then
         'Modified by Morgan 2019/7/25 加碼數檢查
         If Len(txtIM(Index)) = 6 Or Len(txtIM(Index)) = 8 Then
            strCusTemp = ChangeCustomerL(txtIM(Index))
            If ClsPDGetAgent(strCusTemp, strTemp) Then
               Label2(1).Caption = LeftB(strTemp, 50)
               'Added by Lydia 2023/02/18 整合特殊備註維護：在輸入Y/X編號若為6碼，統一補足為8碼。
               If m_EditMode <> 0 Then
                  txtIM(Index) = Left(ChangeCustomerL(txtIM(Index)), 8)
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
      If txtIM(Index) <> "" Then
         'Modified by Morgan 2019/7/25 加碼數檢查
         If Len(txtIM(Index)) = 6 Or Len(txtIM(Index)) = 8 Then
            strCusTemp = ChangeCustomerL(txtIM(Index))
            If ClsPDGetCustomer(strCusTemp, strTemp) Then
               Label2(2).Caption = LeftB(strTemp, 50)
               'Added by Lydia 2023/02/18 整合特殊備註維護：在輸入Y/X編號若為6碼，統一補足為8碼。
               If m_EditMode <> 0 Then
                  txtIM(Index) = Left(ChangeCustomerL(txtIM(Index)), 8)
               End If
               'end 2023/02/18
            Else
               'MsgBox "客戶編號輸入錯誤！", vbCritical 'Remove by Lydia 2021/11/01 模組已彈訊息
               If m_EditMode <> 0 Then Cancel = True
            End If
         Else
            MsgBox "客戶編號只可輸入6碼或8碼！", vbCritical
            If m_EditMode <> 0 Then Cancel = True
         End If
      End If
   Case 6 '案件性質
      Label2(3).Caption = ""
      If txtIM(Index) <> "" Then
         'Remove by Lydia 2018/02/14  開放輸入案件性質
         'If InStr(mInPty, txtIM(Index)) = 0 Then
         '   MsgBox "目前僅開放設定" & mInPtyStr & "！", vbExclamation
         If Len(txtIM(Index)) <> 3 Then
            MsgBox "案件性質請輸入3碼！", vbExclamation
         'end 2018/02/14
            Cancel = True
         Else
            'Added by Lydia 2018/02/14 排除新申請案
            If InStr(NewCasePtyList, txtIM(Index)) > 0 Then
                MsgBox "案件性質不可輸入新申請案的案件性質！", vbExclamation
                Cancel = True
            End If
            'end 2018/02/14
            'Added by Lydia 2025/08/26
            If Trim(txtIM(9)) = "N" Then
                MsgBox "確定不新增內部收文！", vbExclamation
                Cancel = True
                Exit Sub
            End If
            'end 2025/08/26
            If ClsPDGetCaseProperty("FCP", txtIM(Index), strTemp) Then
               Label2(3).Caption = strTemp
               If m_EditMode <> 0 Then
                 txtIM(9) = "Y" '預設901+924=>產生B類單
                 'txtIM(10) = "Y" 'Remove by Lydia 2021/11/01 拿掉IM10='Y'
                 If txtIM(7) = "" Then txtIM(7) = "7"
               End If
            Else
               If m_EditMode <> 0 Then Cancel = True
            End If
         End If
      Else
         If m_EditMode <> 0 Then
            txtIM(7) = ""
            'Modified by Lydia 2025/08/26
            'txtIM(9) = ""
            If Trim(txtIM(9)) <> "N" Then txtIM(9) = ""
            txtIM(10) = ""
            txtIM(21) = ""
            'end 2025/08/26
            'txtIM(10) = "Y"  'Remove by Lydia 2021/11/01 拿掉IM10='Y'
         End If
      End If
   Case 8
      If txtIM(Index) <> "" Then
      '以FCP代表國外案,如果除了系統別其他條件相同,建議系統別皆要設定
         If Not (txtIM(Index) = "P" Or txtIM(Index) = "FCP") Then
            MsgBox "系統別只能輸入P和FCP ！", vbExclamation
            Cancel = True
         End If
      End If
   'Modified by Lydia 2021/11/01 拿掉IM10=Y
   Case 9
      If m_EditMode <> 0 And txtIM(Index) <> "" Then
         'Modified by Lydia 2025/08/26
         'If txtIM(Index) <> "Y" Then
         '   MsgBox "只能輸入Y ！", vbExclamation
         If txtIM(Index) <> "Y" And txtIM(Index) <> "N" Then
            MsgBox "只能輸入Y 或 N！", vbExclamation
         'end 2025/08/26
            Cancel = True
         Else
            If Index = 9 And txtIM(Index) = "Y" Then
                If txtIM(7) = "" Then txtIM(7) = "7"
                If txtIM(6) = "" Then txtIM(6) = "901": txtIM_Validate 6, False
                'txtIM(10) = "Y" 'Remove by Lydia 2021/11/01
            End If
         End If
      ElseIf m_EditMode <> 0 Then
            txtIM(7) = ""
            txtIM(9) = ""
            'txtIM(10) = "Y" 'Remove by Lydia 2021/11/01
            txtIM(6) = ""
            Label2(3).Caption = ""
            'Added by Lydia 2025/08/26
            txtIM(10) = ""
            txtIM(21) = ""
            'end 2025/08/26
      End If
   'Added by Lydia 2021/11/01 承辦期限類型
   'Modified by Lydia 2022/10/13 +IM20, IM21
   Case 10, 18, 20, 21
      If m_EditMode <> 0 And txtIM(Index) <> "" Then
         If txtIM(Index) <> "1" And txtIM(Index) <> "2" Then
            MsgBox "只能輸入1 或 2 ！", vbExclamation
            Cancel = True
         End If
      End If
   'end 2021/11/01
   'Added by Lydia 2021/11/11
   Case 19 '客戶C類來函性質
      Label2(4).Caption = ""
      If txtIM(Index) <> "" Then
         If Len(txtIM(Index)) <> 4 Then
            MsgBox "案件性質請輸入4碼！", vbExclamation
            Cancel = True
         Else
            '參考ClsPrtForm001.PrintCFormNew
            If InStr("1001,1008,1204,1217,1913,1603,1604", txtIM(Index)) > 0 Then
                MsgBox "案件性質不可輸入1001核准,1008核發,1204通知實審日,1217通知形式審查,1603專利證書,1604專利權消滅,1913通知期限！", vbExclamation
                Cancel = True
            End If
            If ClsPDGetCaseProperty("FCP", txtIM(Index), strTemp) Then
               Label2(4).Caption = strTemp
            Else
               If m_EditMode <> 0 Then Cancel = True
            End If
         End If
      End If
   'end 2021/11/11
   End Select
   
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean, idx As Integer
   If txtIM(2) = "" Then
      MsgBox "備註內容不可空白！", vbExclamation
      txtIM(2).SetFocus
      Exit Function
   End If
   If txtIM(3) & txtIM(4) & txtIM(5) = "" Then
      MsgBox "請輸入本所案號、代理人或申請人！", vbExclamation
      txtIM(3).SetFocus
      Exit Function
   End If
   
   For idx = 3 To 10
      txtIM_Validate idx, bCancel
      If bCancel = True Then
         txtIM(idx).SetFocus
         Exit Function
      End If
   Next
   
   'Added by Lydia 2025/08/26
   If Trim(txtIM(6)) = "" Or txtIM(9) <> "Y" Then
      If Trim(txtIM(8) & txtIM(6) & txtIM(7) & txtIM(10) & txtIM(21)) <> "" Then
         MsgBox "不新增內部收文，請清空其他有關內部收文的設定！！", vbExclamation
         Exit Function
      End If
   End If
   'end 2025/08/26
   
   'Added by Lydia 2021/11/01 承辦期限的檢查
   If txtIM(6) <> "" And txtIM(10) <> "1" And txtIM(10) <> "2" Then
      MsgBox "請輸入內部收文承辦天數的類型！", vbExclamation
      txtIM(10).SetFocus 'Added by Lydia 2021/11/22
      Exit Function
   End If
   'Added by Lydia 2022/10/13 內部收文期限起算類型
   If txtIM(6) <> "" And txtIM(21) <> "1" And txtIM(21) <> "2" Then
      MsgBox "請輸入內部收文期限起算類型！", vbExclamation
      txtIM(21).SetFocus
      Exit Function
   End If

   'Added by Lydia 2021/11/11 客戶C類來函性質
   If txtIM(19) <> "" Then
      txtIM_Validate 19, bCancel
      If bCancel = True Then
         Exit Function
      End If
      If txtIM(11) = "" Then
         MsgBox "請輸入客戶C類來函承辦天數！", vbExclamation
         txtIM(11).SetFocus 'Added by Lydia 2021/11/22
         Exit Function
      End If
      If txtIM(18) = "" Then
         MsgBox "請輸入客戶C類來函承辦天數的類型！", vbExclamation
         txtIM(18).SetFocus 'Added by Lydia 2021/11/22
         Exit Function
      End If
   End If
   'end 2021/11/11
   'Added by Lydia 2021/11/01
   If txtIM(11) <> "" And txtIM(18) <> "1" And txtIM(18) <> "2" Then
      MsgBox "請輸入客戶C類來函承辦天數的類型！", vbExclamation
      txtIM(18).SetFocus 'Added by Lydia 2021/11/22
      Exit Function
   End If
   'end 2021/11/01
   
   'Added by Lydia 2022/10/13 客戶C類來函期限起算類型
   If txtIM(11) <> "" And txtIM(20) <> "1" And txtIM(20) <> "2" Then
      MsgBox "請輸入客戶C類來函期限起算類型！", vbExclamation
      txtIM(20).SetFocus
      Exit Function
   End If
   
   'Added by Lydia 2021/11/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
On Error GoTo ErrHnd
   strExc(0) = strUserNum
   strExc(1) = strSrvDate(1)
   strExc(2) = Right("000000" & ServerTime, 6)
   
   cnnConnection.BeginTrans
   If m_EditMode = 1 Then
      'Modified by Lydia 2017/01/11 避免LOG語法分析錯誤
      'strSql = "insert into IncomMemo(im01,im02,im03,im04,im05,im06,im07,im08,im09,im10,im11,im12,im13,im14)" & _
         " select nvl(max(im01),0)+1 im01,'" & ChgSQL(txtIM(2)) & "' im02" & _
         ",'" & txtIM(3) & "' im03,'" & txtIM(4) & "' im04,'" & txtIM(5) & "' im05,'" & txtIM(6) & "' im06,'" & txtIM(7) & "' im07" & _
         ",'" & txtIM(8) & "' im08,'" & txtIM(9) & "' im09,'" & txtIM(10) & "' im10,'" & txtIM(11) & "' im11, '" & strExc(0) & "' am12, '" & strExc(1) & "' am13, '" & strExc(2) & "' am14 from IncomMemo "
       'Modified by Lydia 2021/11/01 +IM18 客戶C類來函承辦期限類型(IM11承辦期限類型)
       'Modified by Lydia 2021/11/11 +IM19 客戶C類來函性質(IM11客戶C類來函性質)
       'Modified by Lydia 2022/10/13 +IM20客戶C類來函期限起算類型, IM21內部收文期限起算類型
       strSql = "insert into IncomMemo(im01,im02,im03,im04,im05,im06,im07,im08,im09,im10,im11,im12,im13,im14,im18,im19,im20,im21) " & _
         " VALUES ('" & Pub_GetDefColMaxNo("IncomMemo", "IM01") & "','" & ChgSQL(txtIM(2)) & "'" & _
         ",'" & txtIM(3) & "','" & txtIM(4) & "','" & txtIM(5) & "','" & txtIM(6) & "','" & txtIM(7) & "'" & _
         ",'" & txtIM(8) & "','" & txtIM(9) & "','" & txtIM(10) & "','" & txtIM(11) & "' " & _
         ", '" & strExc(0) & "', '" & strExc(1) & "' , '" & strExc(2) & "','" & txtIM(18) & "','" & txtIM(19) & "','" & txtIM(20) & "','" & txtIM(21) & "') "
         
   Else
      'Modified by Lydia 2021/11/01 +IM18 客戶C類來函承辦期限類型(IM11承辦期限類型)
      'Modified by Lydia 2021/11/11 +IM19 客戶C類來函性質(IM11客戶C類來函性質)
      'Modified by Lydia 2022/10/13 +IM20客戶C類來函期限起算類型, IM21內部收文期限起算類型
      strSql = "update IncomMemo set im02='" & ChgSQL(txtIM(2)) & "',im03='" & txtIM(3) & "'" & _
         ",im04='" & txtIM(4) & "',im05='" & txtIM(5) & "',im06='" & txtIM(6) & "'" & _
         ",im07='" & txtIM(7) & "',im08='" & txtIM(8) & "',im09='" & txtIM(9) & "'" & _
         ",im10='" & txtIM(10) & "',im11='" & txtIM(11) & "'" & _
         ",im15='" & strExc(0) & "',im16='" & strExc(1) & "',im17='" & strExc(2) & "' " & _
         ",im18='" & txtIM(18) & "',im19='" & txtIM(19) & "',im20='" & txtIM(20) & "',im21='" & txtIM(21) & "'  " & _
         "where im01=" & txtIM(1)
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
   With GRD1
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

Private Function FormDelete() As Boolean
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   strSql = "delete from IncomMemo where im01=" & txtIM(1)
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   FormDelete = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

'Modified by Lydia 2015/01/07 判斷是否存在相同條件的記錄
Private Function RecIsExist() As Boolean
   
strExc(0) = ""
If Trim(txtIM(3)) <> "" Then
   strExc(0) = strExc(0) & "and im03='" & Trim(txtIM(3)) & "' "
End If
If Trim(txtIM(4)) <> "" Then
   'Modified by Lydia 2019/07/31 改成=判斷; 因為無法先輸入8碼後再輸入6碼
   'strExc(0) = strExc(0) & "and instr(im04,'" & Trim(txtIM(4)) & "') > 0 "
   strExc(0) = strExc(0) & "and im04='" & Trim(txtIM(4)) & "' "
   'Added by Lydia 2016/12/28 區別只有代理人或客戶的條件
   If Trim(txtIM(5)) = "" Then strExc(0) = strExc(0) & "and im05 is null "
End If
If Trim(txtIM(5)) <> "" Then
   'Modified by Lydia 2019/07/31 改成=判斷; 因為無法先輸入8碼後再輸入6碼
   'strExc(0) = strExc(0) & "and instr(im05,'" & Trim(txtIM(5)) & "') > 0 "
   strExc(0) = strExc(0) & "and im05='" & Trim(txtIM(5)) & "' "
   'Added by Lydia 2016/12/28 區別只有代理人或客戶的條件
   If Trim(txtIM(4)) = "" Then strExc(0) = strExc(0) & "and im04 is null "
End If
'案件性質
'Remove by Lydia 2021/10/22 取消; Y54116+X48637能存在兩筆記錄，是因為B類收文設定也被視做判斷條件
'If Trim(txtIM(6)) <> "" Then
'   strExc(0) = strExc(0) & "and im06='" & Trim(txtIM(6)) & "' "
'End If
'end 2021/10/22
'系統別
If Trim(txtIM(8)) <> "" Then
   strExc(0) = strExc(0) & "and im08='" & Trim(txtIM(8)) & "' "
End If
'是否新增內部收文
'Remove by Lydia 2021/10/22 取消; Y54116+X48637能存在兩筆記錄，是因為B類收文設定也被視做判斷條件
'If Trim(txtIM(9)) <> "" Then
'   strExc(0) = strExc(0) & "and im09='" & Trim(txtIM(9)) & "' "
'End If
'end 2021/10/22

'Added by Lydia 2021/11/11 客戶C類來函性質
If Trim(txtIM(19)) <> "" Then
   strExc(0) = strExc(0) & "and im19='" & Trim(txtIM(19)) & "' "
'Added by Lydia 2024/05/30
Else
   strExc(0) = strExc(0) & "and im19 is null "
'end 2024/05/30
End If
'end 2021/11/11

strExc(0) = Mid(strExc(0), 4, Len(strExc(0)) - 4)
   strExc(1) = " select * from IncomMemo where " & strExc(0)
   intR = 1
   Set rsRead = ClsLawReadRstMsg(intR, strExc(1))
   If intR = 1 Then
      'Added by Lydia 2021/11/01 排除現在修改的記錄
      If rsRead.RecordCount = 1 And Trim(rsRead.Fields("IM01")) = Trim(txtIM(1)) Then
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

'Added by Lydia 2018/08/29 顯示備註的優先順序
Private Sub CmdMsg_Click()
      frm880004.iStiu = 7
      frm880004.Show
End Sub

'Added by Lydia 2021/11/01
Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    If Index <> 5 Then 'Added by Lydia 2022/10/03
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
   Case 3 '系統別
      If txtFM2(Index) <> "" Then
      '以FCP代表國外案,如果除了系統別其他條件相同,建議系統別皆要設定
         If Not (txtFM2(Index) = "P" Or txtFM2(Index) = "FCP") Then
            MsgBox "系統別只能輸入P和FCP ！", vbExclamation
            Cancel = True
         End If
      End If
   'Added by Lydia 2021/11/11
   Case 4 '客戶C類來函性質
      lblFM2(Index).Caption = ""
      If txtFM2(Index) <> "" Then
         If Len(txtFM2(Index)) <> 4 Then
            MsgBox "案件性質請輸入4碼！", vbExclamation
            Cancel = True
         Else
            '參考ClsPrtForm001.PrintCFormNew
            If InStr("1001,1008,1204,1217,1913,1603,1604", txtFM2(Index)) > 0 Then
                MsgBox "案件性質不可輸入1001核准,1008核發,1204通知實審日,1217通知形式審查,1603專利證書,1604專利權消滅,1913通知期限！", vbExclamation
                Cancel = True
            End If
            If ClsPDGetCaseProperty("FCP", txtFM2(Index), strTemp) Then
               lblFM2(4).Caption = strTemp
            Else
               Cancel = True
            End If
         End If
      End If
   'end 2021/11/11
   End Select
End Sub
