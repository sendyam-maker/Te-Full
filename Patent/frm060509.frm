VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060509 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦單設定維護"
   ClientHeight    =   6072
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6072
   ScaleWidth      =   8952
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
            Picture         =   "frm060509.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060509.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060509.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060509.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060509.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060509.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060509.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060509.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060509.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060509.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060509.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   8952
      _ExtentX        =   15790
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
      Height          =   5370
      Left            =   30
      TabIndex        =   25
      Top             =   660
      Width           =   8895
      _ExtentX        =   15685
      _ExtentY        =   9462
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm060509.frx":20F4
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
      Tab(0).Control(8)=   "Label3(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Lbl07"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(9)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(10)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(11)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textCUID"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label2(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtDB(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtDB(5)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtDB(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtDB(2)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtDB(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtDB(6)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtDB(7)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtDB(9)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtDB(10)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtDB(11)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Combo1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Frame1"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Frame2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Frame3"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Cob08"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm060509.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo2"
      Tab(1).Control(1)=   "cmdQuery"
      Tab(1).Control(2)=   "GRD1"
      Tab(1).Control(3)=   "txtFM2(3)"
      Tab(1).Control(4)=   "Label1(18)"
      Tab(1).Control(5)=   "Label1(17)"
      Tab(1).Control(6)=   "lblPS"
      Tab(1).Control(7)=   "Label1(14)"
      Tab(1).Control(8)=   "Label1(16)"
      Tab(1).Control(9)=   "Label1(15)"
      Tab(1).Control(10)=   "Label1(13)"
      Tab(1).Control(11)=   "txtFM2(0)"
      Tab(1).Control(12)=   "txtFM2(1)"
      Tab(1).Control(13)=   "txtFM2(2)"
      Tab(1).Control(14)=   "lblFM2(1)"
      Tab(1).Control(15)=   "lblFM2(2)"
      Tab(1).ControlCount=   16
      Begin VB.ComboBox Combo2 
         Height          =   276
         Left            =   -69960
         TabIndex        =   52
         Text            =   "Combo2"
         Top             =   382
         Width           =   2175
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   300
         Left            =   -72300
         TabIndex        =   51
         Top             =   390
         Width           =   885
      End
      Begin VB.ComboBox Cob08 
         Height          =   276
         Left            =   1080
         TabIndex        =   63
         Text            =   "Combo1"
         Top             =   2394
         Width           =   2985
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   735
         Left            =   150
         TabIndex        =   57
         Top             =   3360
         Width           =   8625
         Begin MSForms.TextBox txtDB 
            Height          =   480
            Index           =   25
            Left            =   6600
            TabIndex        =   17
            Top             =   210
            Width           =   1350
            VariousPropertyBits=   -1466941413
            MaxLength       =   20
            ScrollBars      =   2
            Size            =   "2381;847"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtDB 
            Height          =   675
            Index           =   24
            Left            =   930
            TabIndex        =   16
            Top             =   0
            Width           =   5580
            VariousPropertyBits=   -1466941413
            MaxLength       =   500
            ScrollBars      =   2
            Size            =   "9842;1191"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label9 
            Caption         =   "排除的X,Y編號："
            Height          =   225
            Left            =   6630
            TabIndex        =   62
            Top             =   30
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "加註內容："
            Height          =   180
            Index           =   12
            Left            =   0
            TabIndex        =   59
            Top             =   270
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "OutLook"
            Height          =   180
            Index           =   7
            Left            =   0
            TabIndex        =   58
            Top             =   60
            Width           =   630
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  '沒有框線
         Height          =   1185
         Left            =   4140
         TabIndex        =   49
         Top             =   4110
         Width           =   4635
         Begin VB.CheckBox Chk26 
            Caption         =   "不印地址條"
            Height          =   180
            Left            =   2550
            TabIndex        =   15
            Top             =   990
            Width           =   1365
         End
         Begin VB.CheckBox Chk23 
            Caption         =   "同時寄專利證書+公告公報"
            Height          =   180
            Left            =   60
            TabIndex        =   14
            Top             =   990
            Width           =   2475
         End
         Begin MSForms.TextBox txtDB 
            Height          =   300
            Index           =   22
            Left            =   1260
            TabIndex        =   13
            Top             =   670
            Width           =   3300
            VariousPropertyBits=   671105055
            MaxLength       =   60
            Size            =   "5821;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtDB 
            Height          =   300
            Index           =   21
            Left            =   1260
            TabIndex        =   12
            Top             =   350
            Width           =   1200
            VariousPropertyBits=   671105055
            MaxLength       =   9
            Size            =   "2117;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtDB 
            Height          =   300
            Index           =   20
            Left            =   1020
            TabIndex        =   11
            Top             =   30
            Width           =   450
            VariousPropertyBits=   671105055
            MaxLength       =   1
            Size            =   "794;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label Label2 
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   72
            Top             =   373
            Width           =   2000
            VariousPropertyBits=   27
            Caption         =   "1111"
            Size            =   "3528;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "聯絡人："
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   600
            TabIndex        =   61
            Top             =   730
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "地址條收件人："
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   60
            TabIndex        =   60
            Top             =   410
            Width           =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "證書正本：　　　1.不寄 2.另寄"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   60
            TabIndex        =   56
            Top             =   90
            Width           =   2475
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Height          =   950
         Left            =   4110
         TabIndex        =   42
         Top             =   2370
         Width           =   3645
         Begin VB.CheckBox Chk18 
            Caption         =   "工程師"
            Height          =   180
            Index           =   21
            Left            =   3660
            TabIndex        =   48
            Top             =   60
            Width           =   1100
         End
         Begin VB.CheckBox Chk19 
            Caption         =   "工程師"
            Height          =   180
            Index           =   21
            Left            =   3660
            TabIndex        =   47
            Top             =   480
            Width           =   1100
         End
         Begin VB.CheckBox Chk19 
            Caption         =   "工程師主管"
            Height          =   180
            Index           =   22
            Left            =   3660
            TabIndex        =   46
            Top             =   720
            Width           =   1305
         End
         Begin VB.CheckBox Chk19 
            Caption         =   "不要副本"
            ForeColor       =   &H00FF00FF&
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   45
            Top             =   720
            Width           =   1100
         End
         Begin VB.CheckBox Chk19 
            Caption         =   "程序主管"
            Height          =   180
            Index           =   12
            Left            =   2535
            TabIndex        =   23
            Top             =   720
            Width           =   1100
         End
         Begin VB.CheckBox Chk19 
            Caption         =   "承辦主管"
            Height          =   180
            Index           =   2
            Left            =   1380
            TabIndex        =   22
            Top             =   720
            Width           =   1100
         End
         Begin VB.CheckBox Chk19 
            Caption         =   "程序"
            Height          =   180
            Index           =   11
            Left            =   2535
            TabIndex        =   21
            Top             =   480
            Width           =   1100
         End
         Begin VB.CheckBox Chk19 
            Caption         =   "承辦"
            Height          =   180
            Index           =   1
            Left            =   1380
            TabIndex        =   20
            Top             =   480
            Width           =   1100
         End
         Begin VB.CheckBox Chk18 
            Caption         =   "程序"
            Height          =   180
            Index           =   11
            Left            =   2535
            TabIndex        =   19
            Top             =   60
            Width           =   1100
         End
         Begin VB.CheckBox Chk18 
            Caption         =   "承辦"
            Height          =   180
            Index           =   1
            Left            =   1380
            TabIndex        =   18
            Top             =   60
            Width           =   1100
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            X1              =   60
            X2              =   4770
            Y1              =   330
            Y2              =   330
         End
         Begin VB.Label Label5 
            Caption         =   "EMAIL副本: "
            Height          =   195
            Left            =   60
            TabIndex        =   44
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label Label4 
            Caption         =   "EMAIL收件人: "
            Height          =   195
            Left            =   60
            TabIndex        =   43
            Top             =   60
            Width           =   1185
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   3960
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   420
         Width           =   2175
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm060509.frx":212C
         Height          =   3075
         Left            =   -74940
         TabIndex        =   26
         Top             =   2190
         Width           =   8745
         _ExtentX        =   15431
         _ExtentY        =   5419
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "流水號|備註內容|本所案號|代理人|申請人|承辦單類別|訊息類別|速別|主旨|附件|其他"
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
      Begin MSForms.TextBox txtFM2 
         Height          =   520
         Index           =   3
         Left            =   -73890
         TabIndex        =   55
         Top             =   1320
         Width           =   6765
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "11933;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註內容："
         Height          =   180
         Index           =   18
         Left            =   -74910
         TabIndex        =   76
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "模糊比對"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   17
         Left            =   -74910
         TabIndex        =   75
         Top             =   1590
         Visible         =   0   'False
         Width           =   720
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   11
         Left            =   5070
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   3000
         VariousPropertyBits=   671105055
         MaxLength       =   50
         Size            =   "5292;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   10
         Left            =   1080
         TabIndex        =   9
         Top             =   3030
         Width           =   3000
         VariousPropertyBits=   671105055
         MaxLength       =   50
         Size            =   "5292;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   9
         Left            =   1080
         TabIndex        =   8
         Top             =   2710
         Width           =   3000
         VariousPropertyBits=   671105055
         MaxLength       =   50
         Size            =   "5292;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   7
         Left            =   3840
         TabIndex        =   5
         Top             =   1446
         Width           =   450
         VariousPropertyBits=   671105055
         MaxLength       =   1
         Size            =   "794;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   6
         Left            =   6510
         TabIndex        =   1
         Top             =   420
         Visible         =   0   'False
         Width           =   630
         VariousPropertyBits=   671105053
         MaxLength       =   2
         Size            =   "1111;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   1
         Left            =   1080
         TabIndex        =   0
         Top             =   420
         Width           =   630
         VariousPropertyBits=   671105053
         MaxLength       =   4
         Size            =   "1111;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   680
         Index           =   2
         Left            =   1080
         TabIndex        =   3
         Top             =   750
         Width           =   5580
         VariousPropertyBits=   -1466941413
         MaxLength       =   800
         ScrollBars      =   2
         Size            =   "9842;1199"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   4
         Left            =   1080
         TabIndex        =   6
         Top             =   1762
         Width           =   1170
         VariousPropertyBits=   671105055
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
         TabIndex        =   7
         Top             =   2078
         Width           =   1170
         VariousPropertyBits=   671105055
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
         TabIndex        =   4
         Top             =   1446
         Width           =   1575
         VariousPropertyBits=   671105055
         MaxLength       =   12
         Size            =   "2778;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   2340
         TabIndex        =   74
         Top             =   1785
         Width           =   6165
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "10874;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   2340
         TabIndex        =   73
         Top             =   2101
         Width           =   6165
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "10874;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID 
         Height          =   285
         Left            =   90
         TabIndex        =   71
         Top             =   5040
         Width           =   7860
         VariousPropertyBits=   671105055
         Size            =   "13864;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblPS 
         Caption         =   "P.S. 輸入本所案號會另外帶該案代理人和申請人的其他設定"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74910
         TabIndex        =   70
         Top             =   1950
         Width           =   4845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦單類別："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   14
         Left            =   -71100
         TabIndex        =   69
         Top             =   442
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         Height          =   180
         Index           =   16
         Left            =   -74910
         TabIndex        =   68
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   15
         Left            =   -74910
         TabIndex        =   67
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   13
         Left            =   -74910
         TabIndex        =   66
         Top             =   435
         Width           =   900
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   285
         Index           =   0
         Left            =   -73890
         TabIndex        =   50
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
         Left            =   -73890
         TabIndex        =   53
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
         Left            =   -73890
         TabIndex        =   54
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
      Begin MSForms.Label lblFM2 
         Height          =   255
         Index           =   1
         Left            =   -72750
         TabIndex        =   65
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
         Left            =   -72750
         TabIndex        =   64
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "其他："
         Height          =   180
         Index           =   11
         Left            =   3720
         TabIndex        =   41
         Top             =   1845
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "附件："
         Height          =   180
         Index           =   10
         Left            =   495
         TabIndex        =   40
         Top             =   3090
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "主旨："
         Height          =   180
         Index           =   9
         Left            =   495
         TabIndex        =   39
         Top             =   2770
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "速別："
         Height          =   180
         Index           =   8
         Left            =   495
         TabIndex        =   38
         Top             =   2454
         Width           =   540
      End
      Begin VB.Label Lbl07 
         AutoSize        =   -1  'True
         Caption         =   "訊息類別：　　　1.彈訊息+列印在承辦單 2.僅列印在承辦單"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2880
         TabIndex        =   37
         Top             =   1506
         Width           =   4725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦單類別："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   4
         Left            =   2880
         TabIndex        =   36
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "注意：1.若欄位內容不輸入，承辦單各欄位使用預設內容。"
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
         Index           =   4
         Left            =   330
         TabIndex        =   35
         Top             =   4815
         Width           =   5010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "模糊比對"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   6
         Left            =   6750
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "流水號："
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   33
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註內容："
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   32
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         Height          =   180
         Index           =   2
         Left            =   315
         TabIndex        =   31
         Top             =   1822
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   3
         Left            =   315
         TabIndex        =   30
         Top             =   2138
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   5
         Left            =   135
         TabIndex        =   29
         Top             =   1506
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "注意：1.代理人/申請人可輸入6碼或8碼，6碼代表含關係企業。112/1/9取消"
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
         Left            =   330
         TabIndex        =   28
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(112/1/9取消) 2.代理人/申請人無論6碼或8碼均包含更名前編號。"
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
         Index           =   0
         Left            =   330
         TabIndex        =   27
         Top             =   4560
         Visible         =   0   'False
         Width           =   5355
      End
   End
End
Attribute VB_Name = "frm060509"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/02/18 整合特殊備註維護：1.將現有資料6碼Y/X編號補足為8碼；2.在新增Y/X編號若為6碼，統一補足為8碼。
'Memo by Lydia 2021/11/01 改成Form2.0 ; GRD1改字型=新細明體-ExtB、txtDB(index)、Label2(index)、textCUID、txtFM2(index)、lblFM2(index)
'Memo by Lydia 2021/11/01 畫面頁籤改成「單筆資料」和「多筆查詢」：上方工具列的「查詢」帶出第一筆符合的資料，在多筆查詢的頁籤可以輸入條件進行查詢，並且在下方的Grid呈現多筆資料。
'Created by Lydia 2019/02/27 通知函承辦單設定維護(FcpEMPbill)
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
'Added by Lydia 2020/08/17
Dim m_FEB18 As String, m_FEB19 As String 'EMAIL收件人、EMAIL副本收件人
Dim oChkBox As CheckBox

'Added by Lydia 2021/11/01
Private Sub cmdQuery_Click()
   
   stCon = ""
   If txtFM2(0) <> "" Then
      If Trim(txtFM2(1).Tag & txtFM2(2).Tag) = "" Then
          stCon = stCon & " and feb03='" & txtFM2(0) & "'"
      Else
          '另外抓本所案號的相關Y編號、X編號條件
          stCon = stCon & " and (feb03='" & txtFM2(0) & "'"
          If txtFM2(1).Tag <> "" Then stCon = stCon & " or instr(" & CNULL(txtFM2(1).Tag) & ", feb04) > 0 "
          If txtFM2(2).Tag <> "" Then stCon = stCon & " or instr(" & CNULL(txtFM2(2).Tag) & ", feb05) > 0 "
          stCon = stCon & ") "
      End If
   Else
      txtFM2(1).Tag = "": txtFM2(2).Tag = ""   '清空本所案號的相關Y編號、X編號條件
   End If
   If txtFM2(1) <> "" Then
      stCon = stCon & " and feb04 like '" & txtFM2(1) & "%'"
   End If
   If txtFM2(2) <> "" Then
      stCon = stCon & " and feb05 like '" & txtFM2(2) & "%'"
   End If
   If Combo2.Text <> "" Then
      stCon = stCon & " and feb06 like '" & Left(Combo2.Text, 2) & "%'"
   End If
   'Added by Lydia 2022/10/03 增加"備註"查詢
   If txtFM2(3) <> "" Then
       stCon = stCon & " and upper(feb02) like '%" & ChgSQL(UCase(txtFM2(3))) & "%' "
   End If
   'end 2022/10/03
   
   strExc(1) = ""
   For intR = 1 To 6
       strExc(1) = strExc(1) & "," & CNULL(Format(intR, "00")) & "," & CNULL(GetFEB06Name(Format(intR, "00")))
   Next intR

   'Modified by Lydia 2021/11/01 拿掉FEB12~FEB17
   'stSQL = "SELECT FEB01,FEB02,FEB03,FEB04,FEB05," & _
                     "DECODE(FEB06 " & strExc(1) & ",FEB06) AS FEB06," & _
                     "FEB07,FEB08,FEB09,FEB10,FEB11,FEB12,FEB13,FEB14,FEB15,FEB16,FEB17,FEB18,FEB19," & _
                     "FEB20,FEB21,FEB22,FEB23,FEB24,FEB25,FEB26 From FCPEMPBILL "
   stSQL = "SELECT FEB01,FEB02,FEB03,FEB04,FEB05," & _
                     "DECODE(FEB06 " & strExc(1) & ",FEB06) AS FEB06," & _
                     "FEB07,FEB08,FEB09,FEB10,FEB11,FEB18,FEB19," & _
                     "FEB20,FEB21,FEB22,FEB23,FEB24,FEB25,FEB26 From FCPEMPBILL "
   stSQL = stSQL & "WHERE 1=1 " & stCon & " order by FEB01"
   intR = 0
   Set rsRead = ClsLawReadRstMsg(intR, stSQL)
   
   Call SetGrd(True)
   If intR = 1 Then
        grd1.FixedCols = 0
        Set grd1.Recordset = rsRead
        Call SetGrd
        grd1.FixedCols = 6
   End If
End Sub


'Add By Sindy 2021/8/6
Private Sub Chk23_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Chk23.ToolTipText = "打勾:二次核對不附公報"
End Sub

'Add By Sindy 2021/3/31
Private Sub Cob08_Validate(Cancel As Boolean)
   '亭妙:速別有快遞字眼，不印地紙條選項請自動打勾。
   If Left(Combo1.Text, 2) = "03" Then '寄證書
      If m_EditMode = "1" Or m_EditMode = "2" Then
         If InStr(Cob08.Text, "快遞") > 0 Then
            Chk26.Value = 1
         End If
      End If
   End If
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   If m_EditMode = "1" Or m_EditMode = "2" Then
       If Combo1.Text = "" Then
            MsgBox "請輸入承辦單類別!", vbCritical
            Combo1.SetFocus
            Cancel = True
       Else
            If InStr(Combo1.Tag, Left(Combo1.Text, 2)) = 0 Then
                MsgBox "請輸入承辦單類別!", vbCritical
                Combo1.SetFocus
                Cancel = True
            Else
                'Added by Lydia 2019/10/24 Sharon: 避免人員輸入資料，鎖起來。
                'Mark by Lydia 2020/02/11 開放輸入; 109021101 外專之核准函列印承辦單，備註內容加上告准承辦單之特殊指示。
                'If Left(Combo1.Text, 2) = "01" Then
                '    txtDB(2).Text = ""
                '    txtDB(2).Enabled = False
                'Else
                    txtDB(2).Enabled = True
                'End If
                'end 2019/10/24
                
                Combo1.Text = GetFEB06Name(Left(Combo1.Text, 2))
                txtDB(6).Text = Left(Combo1.Text, 2)
                txtDB_Validate 6, False
            End If
       End If
   'Add By Sindy 2021/3/31
   ElseIf m_EditMode = "4" Then
      Call ChgFrame
   '2021/3/31 END
   End If
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
         'Modified by Lydia 2021/11/22 取消
         'If Me.ActiveControl <> txtDB(2) Then KeyCode = 0: Action 11 '備註可用Enter換行
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
   m_bInsert = IsUserHasRightOfFunction("frm060509", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm060509", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm060509", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm060509", strFind, False)
  
   MoveFormToCenter Me
   
   Combo1.Clear
   strExc(1) = ""
   For intR = 1 To 6
        Combo1.AddItem GetFEB06Name(Format(intR, "00"))
        strExc(1) = strExc(1) & Format(intR, "00") & ","
   Next intR
   Combo1.Tag = strExc(1)
   
   'Added by Lydia 2021/11/01
   Frame1.BackColor = &H8000000F
   Frame2.BackColor = &H8000000F
   Frame3.BackColor = &H8000000F
   For Each oLabel In lblFM2
       oLabel.BackColor = &H8000000F
   Next
   Combo2.Clear
   For intR = 1 To 6
        Combo2.AddItem GetFEB06Name(Format(intR, "00"))
   Next intR
   Call SetGrd(True)
   'end 2021/11/01
   
   textCUID.BackColor = &H8000000F
   Action 6 '預設第一筆
   UpdateToolbarState
   
   'Add By Sindy 2021/3/22
   Frame1.Left = 4110
   Frame1.Top = 2370
   Frame2.Left = 4170
   Frame2.Top = 2100
   Call SetCob08
   '2021/3/22 END
   
   Me.SSTab1.Tab = 1 'Added by Lydia 2021/11/01 改從多筆查詢頁籤開始
End Sub

'Add By Sindy 2021/3/31
Private Sub SetCob08()
   '速別
   Cob08.Clear
   Cob08.AddItem ""
   Cob08.AddItem "E-MAIL + 掛號"
   Cob08.AddItem "E-MAIL + 快遞"
   Cob08.AddItem "E-MAIL（不寄紙本）"
   Cob08.AddItem "E-MAIL + FAX + 快遞"
   Cob08.AddItem "FAX + 快遞"
   Cob08.AddItem "快遞（不E-MAIL）"
   Cob08.AddItem "上傳平台 + 掛號"
   Cob08.AddItem "上傳平台 + 快遞"
   Cob08.AddItem "上傳平台（不寄紙本）"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060509 = Nothing
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
      Combo1.Locked = True
      Cob08.Locked = True 'Add By Sindy 2021/3/31
      txtDB(2).Enabled = True 'Added by Lydia 2019/10/24
      Frame1.Enabled = False 'Added by Lydia 2020/08/17
      Frame2.Enabled = False 'Added by Sindy 2021/3/22
      Frame3.Enabled = False 'Added by Sindy 2021/3/22
   Case Else
      For Each oText In txtDB
         oText.Locked = False
      Next
      Combo1.Locked = False
      Cob08.Locked = False 'Add By Sindy 2021/3/31
      Frame1.Enabled = True 'Added by Lydia 2020/08/17
      Frame2.Enabled = True 'Added by Sindy 2021/3/22
      Frame3.Enabled = True 'Added by Sindy 2021/3/22
      If m_EditMode <> 4 Then
         txtDB(1).Locked = True
         '先設定承辦單類別
         'Modified by Lydia 2019/03/05 不清空承辦單類別,改為備註優先
         'Combo1.SetFocus
         txtDB(2).SetFocus
         Call txtDB_GotFocus(2)
         'end 2019/03/05
         'Added by Lydia 2019/10/24 Sharon: 避免人員輸入資料，鎖起來。
         'Mark by Lydia 2020/02/11 開放輸入; 109021101 外專之核准函列印承辦單，備註內容加上告准承辦單之特殊指示。
         'If Left(Combo1.Text, 2) = "01" Then
         '    txtDB(2).Enabled = False
         'Else
             txtDB(2).Enabled = True
         'End If
         'end 2019/10/24
         Call txtDB_Validate(20, False) 'Add By Sindy 2021/3/23
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
               If TxtValidate = False Then
                  Exit Sub
               Else
                  'Modified by Lydia 2021/11/01 新增,修改都要判斷
                  'If m_EditMode = 1 Then
                  '   If RecIsExist = True Then Exit Sub
                  'End If
                  If RecIsExist = True Then Exit Sub
                  
                  If FormSave() = False Then
                     MsgBox "存檔失敗!", vbCritical
                     Exit Sub
                  Else
                     strKind = m_EditMode 'Added by Lydia 2021/11/01 記錄新增模式
                     m_EditMode = 0
                     If txtDB(1) = "" Then
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
                        Combo2.Text = Combo1.Text
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
 Dim stKEY As String
    
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case p_iWay
      Case 0 '第一筆
         strExc(0) = "SELECT nvl(min(FEB01),0) FROM FcpEMPbill"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 1 '前一筆
         strExc(0) = "SELECT nvl(max(FEB01),0) FROM FcpEMPbill where FEB01<" & txtDB(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 6
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 2 '後一筆
         strExc(0) = "SELECT nvl(min(FEB01),0) FROM FcpEMPbill where FEB01>" & txtDB(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 7
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 3 '最後筆
         strExc(0) = "SELECT nvl(max(FEB01),0) FROM FcpEMPbill"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
   End Select
   
   
   If stKEY <> "" Then
      ReadData stKEY
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
      stCon = " and FEB01=" & pKey
   '多筆
   Else
      Call SetGrd(True) '清空
      If txtDB(2) <> "" Then
         stCon = stCon & " and FEB02 like '%" & ChgSQL(txtDB(2)) & "%'"
      End If
      If txtDB(3) <> "" Then
         stCon = stCon & " and FEB03='" & txtDB(3) & "'"
      End If
      If txtDB(4) <> "" Then
         stCon = stCon & " and FEB04 like '" & txtDB(4) & "%'"
      End If
      If txtDB(5) <> "" Then
         stCon = stCon & " and FEB05 like '" & txtDB(5) & "%'"
      End If
      If Combo1.Text <> "" Then
         stCon = stCon & " and FEB06 like '" & Left(Combo1.Text, 2) & "%'"
      End If
      'Add By Sindy 2021/3/31
      If Cob08.Text <> "" Then
         stCon = stCon & " and FEB08 like '" & Cob08.Text & "%'"
      End If
      If txtDB(20) <> "" Then
         stCon = stCon & " and FEB20='" & txtDB(20) & "'"
      End If
      If Chk23.Value = 1 Then
         stCon = stCon & " and FEB23='Y'"
      End If
      If txtDB(24) <> "" Then
         stCon = stCon & " and FEB24 like '%" & ChgSQL(txtDB(24)) & "%'"
      End If
      If txtDB(25) <> "" Then
         stCon = stCon & " and FEB25 like '%" & ChgSQL(txtDB(25)) & "%'"
      End If
      If Chk26.Value = 1 Then
         stCon = stCon & " and FEB26='Y'"
      End If
      '2021/3/31 END
   End If
   
   FormReset
   
   strExc(1) = ""
   For intI = 1 To 6
       strExc(1) = strExc(1) & "," & CNULL(Format(intI, "00")) & "," & CNULL(GetFEB06Name(Format(intI, "00")))
   Next intI
   'Modified by Lydia 2020/08/17 +FEB18,FEB19
   'Modify By Sindy 2021/3/22 + ,FEB20,FEB21,FEB22,FEB23,FEB24,FEB25,FEB26
   strExc(0) = "SELECT FEB01,FEB02,FEB03,FEB04,FEB05," & _
                     "DECODE(FEB06 " & strExc(1) & ",FEB06) AS FEB06," & _
                     "FEB07,FEB08,FEB09,FEB10,FEB11,FEB12,FEB13,FEB14,FEB15,FEB16,FEB17,FEB18,FEB19," & _
                     "FEB20,FEB21,FEB22,FEB23,FEB24,FEB25,FEB26 From FCPEMPBILL "
   strExc(0) = strExc(0) & "WHERE 1=1 " & stCon & " order by FEB01"
   intI = 1
   Set rsRead = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If m_EditMode = 4 Then
         'Modified by Lydia 2021/11/01 改成單筆查詢
         'Set GRD1.Recordset = rsRead
         'Call SetGrd
         'If rsRead.RecordCount > 1 Then
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
      SetData rsRead
      ReadData = True
   End If
   Call ChgFrame 'Add By Sindy 2021/3/23
End Function

Private Sub ChgFrame()
   Frame1.Visible = False
   Frame2.Visible = False
   Frame3.Visible = False
   'Added by Lydia 2020/08/17 年證費請款函才顯示
   'Modified by Lydia 2021/01/21 增加:實審05
   If Left(Combo1.Text, 2) = "06" Or Left(Combo1.Text, 2) = "05" Then
       Frame1.Visible = True
       'Added by Lydia 2021/01/21 工程師選項
       If Left(Combo1.Text, 2) <> "05" Then
           Chk18(21).Visible = False
           Chk19(21).Visible = False
           Chk19(22).Visible = False
       Else
           Chk18(21).Visible = True
           Chk19(21).Visible = True
           Chk19(22).Visible = True
       End If
       'end 2021/01/21
   'Add By Sindy 2021/3/22
   ElseIf Left(Combo1.Text, 2) = "03" Then '寄證書才顯示
      Frame2.Visible = True
      Frame3.Visible = True
   '2021/3/22 END
   End If
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer, iR As Integer
   
   'Modify By Sindy 2021/3/23 + "FEB12", "FEB13", "FEB14", "FEB15", "FEB16", "FEB17", "FEB18", "FEB19", _
                           "證書正本", "收件人", "聯絡人", "公報", "Outlook", "排除編號", "不印地址條")
   'Modified by Lydia 2021/11/01 不顯示"其他";拿掉FEB12~FEB17
   'arrGridHeadText = Array("流水號", "備註內容", "本所案號", "代理人", "申請人", "承辦單類別", "訊息類別", "速別", "主旨", "附件", "其他", _
                           "FEB12", "FEB13", "FEB14", "FEB15", "FEB16", "FEB17", "FEB18", "FEB19", _
                           "證書正本", "收件人", "聯絡人", "公報", "Outlook", "排除編號", "不印地址條")
   'arrGridHeadWidth = Array(800, 1200, 1200, 1000, 1000, 1000, 0, 1000, 1000, 1000, 1000, _
                           0, 0, 0, 0, 0, 0, 0, 0, _
                           1000, 1000, 1000, 1000, 1000, 1000, 1000)
   arrGridHeadText = Array("流水號", "備註內容", "本所案號", "代理人", "申請人", "承辦單類別", "訊息類別", "速別", "主旨", "附件", "其他", _
                            "FEB18", "FEB19", "證書正本", "收件人", "聯絡人", "公報", "Outlook", "排除編號", "不印地址條")
   arrGridHeadWidth = Array(800, 1200, 1200, 1000, 1000, 1000, 0, 1000, 1000, 1000, 0, _
                           0, 0, 1000, 1000, 1000, 1000, 1000, 1000, 1000)
                           
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
      grd1.CellAlignment = flexAlignCenterCenter
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow) 'Added by Lydia 2021/02/02
   Next
   
   For intI = 1 To grd1.Rows - 1
        grd1.row = intI
        For iRow = 0 To grd1.Cols - 1
           grd1.col = iRow
           'Added by Lydia 2021/11/11
           If InStr("13,19", Format(iRow, "00")) > 0 Then '置中
                grd1.CellAlignment = flexAlignCenterCenter
           Else
           'end 2021/11/11
                grd1.CellAlignment = flexAlignLeftCenter '內容靠左
           End If 'Added by Lydia 2021/11/11
        Next iRow
   Next intI
   
   grd1.Visible = True
End Sub

Private Sub SetData(ByRef rsQuery As ADODB.Recordset, Optional ByVal iRow As Integer)
Dim tmpArr1 As Variant, intP As Integer 'Added by Lydia 2020/08/17

   If iRow > 0 Then
      rsQuery.MoveFirst
      If iRow > 1 Then
         rsQuery.Move iRow - 1
      End If
      SSTab1.Tab = 0
   End If
   
   With rsQuery
        For Each oText In txtDB
           If oText.Index <> 6 Then
               oText = "" & .Fields("FEB" & Format(oText.Index, "00"))
           Else   '承辦單類別(前2碼)
               oText = Left("" & .Fields("FEB" & Format(oText.Index, "00")), 2)
               Combo1.Text = "" & .Fields("FEB" & Format(oText.Index, "00"))
           End If
        Next
        'Added by Lydia 2020/08/17 (設定) email收件人/副本
        m_FEB18 = "" & .Fields("feb18")
        m_FEB19 = "" & .Fields("feb19")
        'end 2020/08/17
        'Add By Sindy 2021/3/23
        Cob08.Text = "" & .Fields("feb08")
        If Left(Combo1.Text, 2) <> "03" And Cob08.Text <> "" Then Cob08.AddItem Cob08.Text
        If "" & .Fields("FEB23") = "Y" Then '優先寄專利證書+公告公報
           Chk23.Value = 1
        Else
           Chk23.Value = 0
        End If
        If "" & .Fields("FEB26") = "Y" Then '不印地址條
           Chk26.Value = 1
        Else
           Chk26.Value = 0
        End If
        '2021/3/23 END
   End With
   UpdateCUID rsQuery
   
   txtDB(1).Tag = txtDB(1)
   If txtDB(4) <> "" Then txtDB_Validate 4, False '代理人
   If txtDB(5) <> "" Then txtDB_Validate 5, False '客戶
   If txtDB(6) <> "" Then txtDB_Validate 6, False '承辦單類別
   If txtDB(21) <> "" Then txtDB_Validate 21, False '地址條收件人 Add By Sindy 2021/3/23
   'Added by Lydia 2020/08/17 (設定) email收件人/副本
   tmpArr1 = Empty
   If m_FEB18 <> "" Then
       tmpArr1 = Split(m_FEB18, ",")
       For intP = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(intP)) <> "" Then
               Chk18(Val(tmpArr1(intP))).Value = 1
           End If
       Next intP
   End If
   tmpArr1 = Empty
   If m_FEB19 <> "" Then
       tmpArr1 = Split(m_FEB19, ",")
       For intP = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(intP)) <> "" Then
               Chk19(Val(tmpArr1(intP))).Value = 1
           End If
       Next intP
   'Added by Lydia 2020/09/17
   Else   '因為有勾選程序會自動勾選「不要副本」，還原預設
       Chk19(0).Value = 0
   End If
   'end 2020/08/17
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
   If IsNull(rsSrcTmp.Fields("FEB12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FEB12")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("FEB12"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("FEB13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FEB13")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("FEB13"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("FEB14")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FEB14")) = False Then
         strTemp = rsSrcTmp.Fields("FEB14")
         strCTime = Format(strTemp, "00:00:00")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("FEB15")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FEB15")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("FEB15"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("FEB16")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FEB16")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("FEB16"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("FEB17")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FEB17")) = False Then
         strTemp = rsSrcTmp.Fields("FEB17")
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
   'Modified by Lydia 2019/03/06 按下"新增"按鈕不清空承辦單類別，方便延續設定;
   'Combo1.Text = ""
   txtDB(6).Text = Left(Combo1.Text, 2)
   'Memo by Lydia 2019/03/05 隱藏"其他欄位"暫不開放。
   
   '非證書函不輸入訊息類別
   txtDB(7).Visible = False
   txtDB(7).Locked = True
   Lbl07.Visible = False
   
   'Added by Lydia 2020/08/17
   Frame1.Visible = False '年證費請款函才顯示
   m_FEB18 = ""
   m_FEB19 = ""
   For Each oChkBox In Chk18
      oChkBox.Value = 0
   Next
   For Each oChkBox In Chk19
      oChkBox.Value = 0
   Next
   'Add By Sindy 2021/3/22
   Cob08.Text = "": Call SetCob08
   '寄證書才顯示
   Frame2.Visible = False
   Frame3.Visible = False
   Chk23.Value = 0
   Chk26.Value = 0
   '2021/3/22 END
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
   'Add By Sindy 2021/3/23
   If Index = 25 Then '排除的X,Y編號
      KeyAscii = UpperCase(KeyAscii)
      ',=44
      'X=88
      'Y=89
      If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 88 And KeyAscii <> 89 Then
         KeyAscii = 0
         Beep
      End If
   '2021/3/23 END
   'Modify By Sindy 2021/3/23 + And Index <> 22
   ElseIf Index <> 2 And Index <> 22 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtDB_Validate(Index As Integer, Cancel As Boolean)
Dim strCusTemp As String, strTemp As String
Dim tmpArr As Variant, intP As Integer 'Add By Sindy 2021/3/23
   
   Select Case Index
         'Added by Lydia 2019/10/07 改用模組控制長度
         Case 2 '備註內容
             If txtDB(Index).Text = "" Then Exit Sub
             'Modififeid by Lydia 2021/03/05 欄位放大到800, 因為存Char所以中文也算一個
             'If Not CheckLengthIsOK(txtDB(Index), 400) Then
             If Len(txtDB(Index)) >= 800 Then
                MsgBox MsgText(9205) & "800個字!", vbCritical + vbOKOnly, MsgText(9001)
                txtDB(Index).SetFocus
                Cancel = True
             End If
        'end 2019/10/07
        Case 3 '本所案號
           If txtDB(Index) <> "" Then
              strExc(0) = "select PA01||PA02||PA03||PA04 from patent where " & ChgPatent(txtDB(Index))
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
              If intI = 0 Then
                 If m_EditMode <> 0 Then 'Added by Lydia 2021/11/01 排除非編輯模式: 因為案號有可能刪除
                      MsgBox "本所案號輸入錯誤!", vbExclamation
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
                  'Modified by Lydia 2020/08/26 debug
                  'If ClsPDGetAgent(strCusTemp, strTemp) Then
                  If ClsPDGetCustomer(strCusTemp, strTemp) Then
                     Label2(2).Caption = strTemp
                     'Added by Lydia 2023/02/18 整合特殊備註維護：在輸入Y/X編號若為6碼，統一補足為8碼。
                     If m_EditMode <> 0 Then
                        txtDB(Index) = Left(ChangeCustomerL(txtDB(Index)), 8)
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
        Case 6 '承辦單類別
           If txtDB(Index) <> "" Then
              If InStr(Combo1.Tag, txtDB(Index)) = 0 Then
                   MsgBox "承辦單類別輸入錯誤！", vbCritical
                   If m_EditMode <> 0 Then Cancel = True
                   Exit Sub
              Else
                  If txtDB(Index) <> "03" Then  '非證書函不輸入訊息類別
                      txtDB(7).Visible = False
                      txtDB(7).Locked = True
                      Lbl07.Visible = False
                      'Add By 2021/3/30
                      txtDB(9).Enabled = True: txtDB(9).BackColor = &H80000005  '主旨
                      txtDB(10).Enabled = True: txtDB(10).BackColor = &H80000005  '附件
                      '2021/3/30 END
                  Else '證書函
                      txtDB(7).Visible = True
                      txtDB(7).Locked = False
                      Lbl07.Visible = True
                      'Add By Sindy 2021/3/22
                      Frame2.Visible = True
                      Frame3.Visible = True
                      'Add By 2021/3/30 亭妙:主旨,附件欄位反灰,不需輸入
                      txtDB(9).Enabled = False: txtDB(9).BackColor = &H8000000F '主旨
                      txtDB(10).Enabled = False: txtDB(10).BackColor = &H8000000F '附件
                      '2021/3/30 END
                      '2021/3/22 END
                  End If
                  Call ChgFrame 'Add By Sindy 2021/3/23
              End If
           End If
        Case 7 '訊息類別
           If txtDB(Index) <> "" Then
              If txtDB(Index) <> "1" And txtDB(Index) <> "2" Then
                   MsgBox "訊息類別輸入錯誤！", vbCritical
                   If m_EditMode <> 0 Then Cancel = True
                   Exit Sub
              End If
           End If
        'Add By Sindy 2021/3/22
        Case 20 '證書正本
            txtDB(21).Enabled = False: txtDB(21).BackColor = &H8000000F
            txtDB(22).Enabled = False: txtDB(22).BackColor = &H8000000F
            Chk26.Enabled = True
            If txtDB(Index) <> "" Then
               If txtDB(Index) <> "1" And txtDB(Index) <> "2" Then
                   MsgBox "證書正本輸入錯誤，只可輸入 1 或 2！", vbCritical
                   If m_EditMode <> 0 Then Cancel = True
                   Exit Sub
               End If
               If m_EditMode = "1" Or m_EditMode = "2" Then
                  If txtDB(Index) = "1" Then '1.不寄
                      Chk26.Value = 1: Chk26.Enabled = False '份數為1
                      '速別(預設值)
                      If Cob08.Text = "" Then Cob08.Text = "E-MAIL（不寄紙本）"
                      '備註內容(預設值)
                      If txtDB(2) = "" Then txtDB(2) = "寄證書時僅E-mail（不須寄專利證書正本，存卷即可）"
                  ElseIf txtDB(Index) = "2" Then '2.另寄
                      txtDB(21).Enabled = True: txtDB(21).BackColor = &H80000005
                      txtDB(22).Enabled = True: txtDB(22).BackColor = &H80000005
                      Chk26.Value = 0: Chk26.Enabled = False '份數為2
                  End If
               End If
            End If
        Case 21 '地址條收件人
            Label2(0).Caption = ""
            If txtDB(Index) <> "" Then
               If Left(txtDB(Index), 1) = "Y" Then '代理人
                  '加碼數檢查
                  If Len(txtDB(Index)) = 6 Or Len(txtDB(Index)) = 8 Then
                     strCusTemp = Left(txtDB(Index) & "000", 9)
                     If ClsPDGetAgent(strCusTemp, strTemp) Then
                        Label2(0).Caption = strTemp
                     Else
                        'MsgBox "代理人編號輸入錯誤！", vbCritical  'Remove by Lydia 2021/11/01 模組已彈訊息
                        If m_EditMode <> 0 Then Cancel = True
                     End If
                  Else
                     MsgBox "代理人編號只可輸入6碼或8碼！", vbCritical
                     If m_EditMode <> 0 Then Cancel = True
                  End If
               Else 'X.申請人
                  '加碼數檢查
                  If Len(txtDB(Index)) = 6 Or Len(txtDB(Index)) = 8 Then
                     strCusTemp = Left(txtDB(Index) & "000", 9)
                     If ClsPDGetCustomer(strCusTemp, strTemp) Then
                        Label2(0).Caption = strTemp
                     Else
                        'MsgBox "客戶編號輸入錯誤！", vbCritical  'Remove by Lydia 2021/11/01 模組已彈訊息
                        If m_EditMode <> 0 Then Cancel = True
                     End If
                  Else
                     MsgBox "客戶編號只可輸入6碼或8碼！", vbCritical
                     If m_EditMode <> 0 Then Cancel = True
                  End If
               End If
            End If
        Case 25 '排除的X,Y編號
            If txtDB(Index) <> "" Then
               '固定的Y及X編號設定時,無須輸入排除編號
               If (txtDB(4) <> "" And txtDB(5) <> "") Or txtDB(3) <> "" Then
                  MsgBox "無須輸入排除的X,Y編號！", vbCritical
                  If m_EditMode <> 0 Then Cancel = True
                  Exit Sub
               '固定Y編號設定時,輸入排除的X編號
               ElseIf txtDB(4) <> "" Then
                  If InStr(txtDB(Index), "Y") > 0 Then
                     MsgBox "請輸入欲排除的X編號！", vbCritical
                     If m_EditMode <> 0 Then Cancel = True
                     Exit Sub
                  End If
               '固定X編號設定時,輸入排除的Y編號
               ElseIf txtDB(5) <> "" Then
                  If InStr(txtDB(Index), "X") > 0 Then
                     MsgBox "請輸入欲排除的Y編號！", vbCritical
                     If m_EditMode <> 0 Then Cancel = True
                     Exit Sub
                  End If
               End If
               tmpArr = Split(txtDB(Index).Text, ",")
               For intP = 0 To UBound(tmpArr)
                  If Left(tmpArr(intP), 1) = "Y" Then '代理人
                     '加碼數檢查
                     If Len(tmpArr(intP)) = 6 Or Len(tmpArr(intP)) = 8 Then
                        strCusTemp = Left(tmpArr(intP) & "000", 9)
                        If ClsPDGetAgent(strCusTemp, strTemp) = False Then
                           'MsgBox "代理人編號輸入錯誤！", vbCritical 'Remove by Lydia 2021/11/01 模組已彈訊息
                           If m_EditMode <> 0 Then Cancel = True
                        End If
                     Else
                        MsgBox "代理人編號只可輸入6碼或8碼！", vbCritical
                        If m_EditMode <> 0 Then Cancel = True
                     End If
                  Else 'X.申請人
                     '加碼數檢查
                     If Len(tmpArr(intP)) = 6 Or Len(tmpArr(intP)) = 8 Then
                        strCusTemp = Left(tmpArr(intP) & "000", 9)
                        If ClsPDGetCustomer(strCusTemp, strTemp) = False Then
                           'MsgBox "客戶編號輸入錯誤！", vbCritical  'Remove by Lydia 2021/11/01 模組已彈訊息
                           If m_EditMode <> 0 Then Cancel = True
                        End If
                     Else
                        MsgBox "客戶編號只可輸入6碼或8碼！", vbCritical
                        If m_EditMode <> 0 Then Cancel = True
                     End If
                  End If
               Next intP
            End If
        '2021/3/22 END
   End Select
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean, idx As Integer
   
   For idx = 3 To 25 '11
      txtDB_Validate idx, bCancel
      If bCancel = True Then
         txtDB(idx).SetFocus
         Exit Function
      End If
   Next
   
   If txtDB(6) = "" Then
      MsgBox "請輸入承辦單類別！", vbExclamation
      Combo1.SetFocus
      Exit Function
   End If
   
   If txtDB(7).Visible = True And txtDB(7).Locked = False Then
      If txtDB(7) = "" Then
        MsgBox "請輸入訊息類別！", vbExclamation
        txtDB(7).SetFocus
        Exit Function
      ElseIf txtDB(7) = "1" And Trim(txtDB(2)) = "" Then
        MsgBox "若無備註內容,請輸入訊息類別=2 ", vbExclamation
        txtDB(7) = "2"
        txtDB(7).SetFocus
        Exit Function
      End If
   End If
   
   If txtDB(3) & txtDB(4) & txtDB(5) = "" Then
      MsgBox "請輸入本所案號、代理人或申請人！", vbExclamation
      txtDB(3).SetFocus
      Exit Function
   End If
   
   'Modify By Sindy 2021/3/31
   'If Trim(txtDB(2) & txtDB(8) & txtDB(9) & txtDB(10) & txtDB(11)) = "" Then
   If Trim(txtDB(2) & Cob08.Text & txtDB(9) & txtDB(10) & txtDB(11)) = "" Then
   '2021/3/31 END
      'Added by Lydia 2019/10/24
      'Mark by Lydia 2020/02/11 開放輸入; 109021101 外專之核准函列印承辦單，備註內容加上告准承辦單之特殊指示。
      'If Left(Combo1.Text, 2) = "01" Then
      '   MsgBox "請輸入速別、主旨、附件或其他！", vbExclamation
      'Else
      ''end 2019/10/24
          MsgBox "請輸入備註內容、速別、主旨、附件或其他！", vbExclamation
      'End If 'end 2019/10/24
      Exit Function
   End If
   
   'Added by Lydia 2019/03/06
   'Mark by Lydia 2020/02/11 開放輸入; 109021101 外專之核准函列印承辦單，備註內容加上告准承辦單之特殊指示。
   'If Left(Combo1.Text, 2) = "01" And Trim(txtDB(2)) <> "" Then
   '    MsgBox "告准的備註請改到核准函輸入備註維護進行設定！", vbInformation, "例外"
   '    txtDB(2).Text = "" 'Added by Lydia 2019/10/24 Sharon: 避免人員輸入資料，鎖起來。
   'End If
   ''end 2019/03/06
   
   'Added by Lydia 2020/08/17 年證費請款函若有指定收件者,一定要含程序
   If Left(Combo1.Text, 2) = "06" And (Chk18(1).Value + Chk18(11).Value + Chk19(1).Value + Chk19(2).Value + Chk19(11).Value + Chk19(12).Value) > 0 Then
       If Chk18(11).Value + Chk19(11).Value + Chk19(12).Value = 0 Then
           MsgBox "請在Email收件人或副本中，增加程序人員！", vbCritical, "檢核資料"
           Exit Function
       End If
   End If
   
   'Add By Sindy 2021/3/22
   If Left(Combo1.Text, 2) = "03" Then '寄證書
      If txtDB(20) = "1" Then '1.不寄
         txtDB(21).Text = ""
         txtDB(22).Text = ""
      ElseIf txtDB(20) = "2" Then '2.另寄
         If txtDB(21).Text = "" Then
            MsgBox "請輸入另寄的收件人編號！", vbCritical, "檢核資料"
            txtDB(21).SetFocus
            Exit Function
         End If
      End If
      If InStr(Cob08.Text, "快遞") = 0 And txtDB(20) <> "1" And Chk26.Value = 1 Then
         If MsgBox("速別中無（快遞）字樣，確定不印地址條嗎？", vbYesNo + vbDefaultButton1 + vbInformation) = vbNo Then
            Cob08.SetFocus
            Exit Function
         End If
      ElseIf InStr(Cob08.Text, "快遞") > 0 Or txtDB(20) = "1" Then
         Chk26.Value = 1
      End If
   End If
   '2021/3/22 END
   
   'Added by Lydia 2021/11/01 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean

On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
   'Added by Lydia 2020/08/17 組合CheckBox
   m_FEB18 = "": m_FEB19 = ""
   If Frame1.Visible = True Then
       For Each oChkBox In Chk18
           If oChkBox.Value = 1 Then
              m_FEB18 = m_FEB18 & "," & Format(oChkBox.Index, "00")
           End If
       Next
       If m_FEB18 <> "" Then m_FEB18 = Mid(m_FEB18, 2)
       
       For Each oChkBox In Chk19
           If oChkBox.Value = 1 Then
              m_FEB19 = m_FEB19 & "," & Format(oChkBox.Index, "00")
           End If
       Next
       If m_FEB19 <> "" Then m_FEB19 = Mid(m_FEB19, 2)
   End If
   'end 2020/08/17
   
   'Create和Update由Trigger設定
   If m_EditMode = 1 Then
      'Modified by Lydia 2020/08/17 拿掉Trigger
      'strSql = "insert into FcpEMPbill(FEB01,FEB02,FEB03,FEB04,FEB05,FEB06,FEB07,FEB08,FEB09,FEB10,FEB11) " & _
                   "VALUES ('" & Pub_GetDefColMaxNo("FcpEMPbill", "FEB01") & "'," & CNULL(ChgSQL(txtDB(2))) & "," & CNULL(txtDB(3)) & " ," & CNULL(txtDB(4)) & " ," & CNULL(txtDB(5)) & " , " & _
                    CNULL(txtDB(6)) & ", " & CNULL(txtDB(7)) & ", " & CNULL(ChgSQL(txtDB(8))) & ", " & CNULL(ChgSQL(txtDB(9))) & ", " & CNULL(ChgSQL(txtDB(10))) & ", " & CNULL(ChgSQL(txtDB(11))) & " ) "
      'Add By Sindy 2021/3/22 + ,FEB20,FEB21,FEB22,FEB23,FEB24,FEB25,FEB26
      strSql = "insert into FcpEMPbill(FEB01,FEB02,FEB03,FEB04,FEB05,FEB06,FEB07,FEB08,FEB09,FEB10,FEB11,FEB12,FEB13,FEB14,FEB18,FEB19" & _
               ",FEB20,FEB21,FEB22,FEB23,FEB24,FEB25,FEB26) " & _
               "VALUES ('" & Pub_GetDefColMaxNo("FcpEMPbill", "FEB01") & "'," & CNULL(ChgSQL(txtDB(2))) & "," & CNULL(txtDB(3)) & " ," & CNULL(txtDB(4)) & " ," & CNULL(txtDB(5)) & " , " & _
               CNULL(txtDB(6)) & ", " & CNULL(txtDB(7)) & ", " & CNULL(ChgSQL(Cob08.Text)) & ", " & CNULL(ChgSQL(txtDB(9))) & ", " & CNULL(ChgSQL(txtDB(10))) & ", " & CNULL(ChgSQL(txtDB(11))) & " , " & _
               CNULL(strUserNum) & ",to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'), " & CNULL(m_FEB18) & "," & CNULL(m_FEB19) & _
               "," & CNULL(txtDB(20)) & "," & CNULL(txtDB(21)) & "," & CNULL(txtDB(22)) & "," & CNULL(IIf(Chk23.Value = 1, "Y", "")) & "," & CNULL(ChgSQL(txtDB(24))) & "," & CNULL(txtDB(25)) & "," & CNULL(IIf(Chk26.Value = 1, "Y", "")) & ") "
   Else
      'Modified by Lydia 2020/08/17 拿掉Trigger
      'strSql = "update FcpEMPbill set FEB02=" & CNULL(ChgSQL(txtDB(2))) & " ,FEB03=" & CNULL(txtDB(3)) & "" & _
         ",FEB04=" & CNULL(txtDB(4)) & " ,FEB05=" & CNULL(txtDB(5)) & "" & _
         ",FEB06=" & CNULL(txtDB(6)) & " ,FEB07=" & CNULL(txtDB(7)) & "" & _
         ",FEB08=" & CNULL(ChgSQL(txtDB(8))) & " ,FEB09=" & CNULL(ChgSQL(txtDB(9))) & "" & _
         ",FEB10=" & CNULL(ChgSQL(txtDB(10))) & " ,FEB11=" & CNULL(ChgSQL(txtDB(11))) & "" & _
         " where FEB01=" & txtDB(1)
      'Add By Sindy 2021/3/22 + ,FEB20,FEB21,FEB22,FEB23,FEB24,FEB25,FEB26
      strSql = "update FcpEMPbill set FEB02=" & CNULL(ChgSQL(txtDB(2))) & ",FEB03=" & CNULL(txtDB(3)) & _
               ",FEB04=" & CNULL(txtDB(4)) & ",FEB05=" & CNULL(txtDB(5)) & _
               ",FEB06=" & CNULL(txtDB(6)) & ",FEB07=" & CNULL(txtDB(7)) & _
               ",FEB08=" & CNULL(ChgSQL(Cob08.Text)) & ",FEB09=" & CNULL(ChgSQL(txtDB(9))) & _
               ",FEB10=" & CNULL(ChgSQL(txtDB(10))) & ",FEB11=" & CNULL(ChgSQL(txtDB(11))) & _
               ",FEB15=" & CNULL(strUserNum) & ",FEB16=to_char(sysdate,'yyyymmdd'),FEB17=to_char(sysdate,'hh24miss')" & _
               ",FEB18=" & CNULL(ChgSQL(m_FEB18)) & ",FEB19=" & CNULL(ChgSQL(m_FEB19)) & _
               ",FEB20=" & CNULL(txtDB(20)) & ",FEB21=" & CNULL(txtDB(21)) & _
               ",FEB22=" & CNULL(txtDB(22)) & ",FEB23=" & CNULL(IIf(Chk23.Value = 1, "Y", "")) & _
               ",FEB24=" & CNULL(ChgSQL(txtDB(24))) & ",FEB25=" & CNULL(txtDB(25)) & _
               ",FEB26=" & CNULL(IIf(Chk26.Value = 1, "Y", "")) & _
               " where FEB01=" & txtDB(1)
   End If
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI '記錄log
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
      If InStr("流水號,證書正本", Me.grd1.Text) > 0 Then
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
   strSql = "delete from FcpEMPbill where FEB01=" & txtDB(1)
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   FormDelete = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

Private Function GetFEB06Name(ByVal pB06 As String) As String
    Select Case pB06
'Modified by Lydia 2019/03/05 統一更名
        Case "01"
             'GetFEB06Name = pB06 & "-核准函"
             GetFEB06Name = pB06 & "-告准承辦單"
        Case "02"
             'GetFEB06Name = pB06 & "-公告通知函"
             GetFEB06Name = pB06 & "-公告公報承辦單"
        Case "03"
             'GetFEB06Name = pB06 & "-證書函"
             GetFEB06Name = pB06 & "-寄證書承辦單"
        Case "04"
             'GetFEB06Name = pB06 & "-繳年費通知函"
             'Modified by Lydia 2019/10/18 更名
             'GetFEB06Name = pB06 & "-年費請款承辦單"
             GetFEB06Name = pB06 & "-繳年費通知承辦單"
        Case "05"
             'GetFEB06Name = pB06 & "-發文實體審查請款函"
             GetFEB06Name = pB06 & "-實審請款承辦單"
        Case "06"
             'GetFEB06Name = pB06 & "-年證費請款函"
             GetFEB06Name = pB06 & "-年證費請款承辦單"
'end 2019/03/05
    End Select
End Function

Private Function RecIsExist() As Boolean

stCon = ""
If Trim(txtDB(3)) <> "" Then
   stCon = stCon & "and FEB03='" & Trim(txtDB(3)) & "' "
End If
If Trim(txtDB(4)) <> "" Then
   'Modified by Lydia 2019/07/31 改成9碼判斷; 因為無法先輸入8碼後再輸入6碼
   'stcon = stcon & "and instr(FEB04,'" & Trim(txtDB(4)) & "') > 0 "
   stCon = stCon & "and feb04='" & Trim(txtDB(4)) & "' "
   '區別只有代理人或客戶的條件
   If Trim(txtDB(5)) = "" Then stCon = stCon & "and FEB05 is null "
End If
If Trim(txtDB(5)) <> "" Then
   'Modified by Lydia 2019/07/31 改成9碼判斷; 因為無法先輸入8碼後再輸入6碼
   'stcon = stcon & "and instr(FEB05,'" & Trim(txtDB(5)) & "') > 0 "
   stCon = stCon & "and feb05='" & Trim(txtDB(5)) & "' "
   '區別只有代理人或客戶的條件
   If Trim(txtDB(4)) = "" Then stCon = stCon & "and FEB04 is null "
End If
If Trim(txtDB(6)) <> "" Then
   stCon = stCon & "and instr(FEB06,'" & Trim(txtDB(6)) & "') > 0 "
End If

If Left(stCon, 3) = "and" Then stCon = Mid(stCon, 4, Len(stCon) - 4)

   stSQL = " select * from FcpEMPbill where " & stCon
   intR = 1
   Set rsRead = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      'Added by Lydia 2021/11/01 排除現在修改的記錄
      If rsRead.RecordCount = 1 And Trim(rsRead.Fields("FEB01")) = Trim(txtDB(1)) Then
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

'Added by Lydia 2020/09/18
Private Sub Chk19_Click(Index As Integer)
   If Chk19(Index).Value = 1 Then
       If Index = 0 Then
           Chk19(1).Value = 0: Chk19(2).Value = 0
           Chk19(11).Value = 0: Chk19(12).Value = 0
           Chk19(21).Value = 0: Chk19(22).Value = 0 'Added by Lydia 2021/01/21
       Else
           Chk19(0).Value = 0
       End If
   End If
End Sub

'Added by Lydia 2020/09/18
Private Sub Chk18_Click(Index As Integer)
'在發承辦單Email是逐一讀取特殊設定 , 若無則預設為一般設定:
'1. EMAIL副本增加勾選項「不要副本」。
'2.當EMAIL收件人勾選程序則自動勾選「不要副本」。
'3.當EMAIL收件人勾選承辦則EMAIL副本自動勾選「承辦主管」。
'4.當EMAIL收件人勾選承辦則EMAIL副本自動勾選「承辦主管」。--Added by Lydia 2021/01/21
   
   'Modified by Lydia 2021/01/21 編輯狀態才預設
   'If Chk18(Index).Value = 1 Then
   If Chk18(Index).Value = 1 And (m_EditMode = 1 Or m_EditMode = 2) Then
      If Index = 11 Then
          Chk19(0).Value = 1
      ElseIf Index = 1 Then
          Chk19(0).Value = 0
          Chk19(2).Value = 1
      'Added by Lydia 2021/01/21
      ElseIf Index = 21 Then
           Chk19(22).Value = 1
      End If
   End If
End Sub

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
   End Select
End Sub

