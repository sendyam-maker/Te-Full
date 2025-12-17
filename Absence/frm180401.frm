VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180401 
   BorderStyle     =   1  '單線固定
   Caption         =   "人事職代及審核主管設定"
   ClientHeight    =   6950
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9240
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6950
   ScaleWidth      =   9240
   Begin TabDlg.SSTab SSTab1 
      Height          =   1330
      Left            =   30
      TabIndex        =   52
      Top             =   2970
      Width           =   9190
      _ExtentX        =   16210
      _ExtentY        =   2346
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   476
      TabCaption(0)   =   "其他"
      TabPicture(0)   =   "frm180401.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Combo2(15)"
      Tab(0).Control(1)=   "Combo2(16)"
      Tab(0).Control(2)=   "Label1(25)"
      Tab(0).Control(3)=   "Label1(24)"
      Tab(0).Control(4)=   "Label1(23)"
      Tab(0).Control(5)=   "Label1(22)"
      Tab(0).Control(6)=   "txtB0124"
      Tab(0).Control(7)=   "ChkB0127"
      Tab(0).Control(8)=   "txtB0125"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "工作評價"
      TabPicture(1)   =   "frm180401.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Combo2(18)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Combo2(19)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Combo2(17)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(26)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(27)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(28)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtB0125 
         Height          =   300
         Left            =   -72780
         MaxLength       =   1
         TabIndex        =   54
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox ChkB0127 
         Caption         =   "居家無法連線"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   180
         Left            =   -74940
         TabIndex        =   53
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "主管3："
         Height          =   180
         Index           =   28
         Left            =   750
         TabIndex        =   67
         Top             =   990
         Width           =   590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "主管2："
         Height          =   180
         Index           =   27
         Left            =   750
         TabIndex        =   66
         Top             =   690
         Width           =   590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "主管1："
         Height          =   180
         Index           =   26
         Left            =   750
         TabIndex        =   65
         Top             =   390
         Width           =   590
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   17
         Left            =   1410
         TabIndex        =   64
         Top             =   330
         Width           =   1520
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2672;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   19
         Left            =   1410
         TabIndex        =   63
         Top             =   930
         Width           =   1520
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2672;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   18
         Left            =   1410
         TabIndex        =   62
         Top             =   630
         Width           =   1520
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2672;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtB0124 
         Height          =   540
         Left            =   -72210
         TabIndex        =   57
         Top             =   750
         Width           =   4520
         VariousPropertyBits=   -1466941413
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "7964;952"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "核准通知是否含簽核主管：　   （N.不含）"
         Height          =   180
         Index           =   22
         Left            =   -74940
         TabIndex        =   61
         Top             =   540
         Width           =   3360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   $"frm180401.frx":0038
         Height          =   180
         Index           =   23
         Left            =   -74940
         TabIndex        =   60
         Top             =   770
         Width           =   9050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(2)"
         Height          =   180
         Index           =   24
         Left            =   -69660
         TabIndex        =   59
         Top             =   300
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "居家職務代理人(1)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   25
         Left            =   -72720
         TabIndex        =   58
         Top             =   300
         Width           =   1470
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   16
         Left            =   -69420
         TabIndex        =   56
         Top             =   240
         Width           =   1520
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2672;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   15
         Left            =   -71220
         TabIndex        =   55
         Top             =   240
         Width           =   1520
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2672;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CheckBox ChkEMailAll 
      Caption         =   "案件EMail須發職代時，案件或人事職代全發"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   4680
      TabIndex        =   51
      Top             =   2130
      Width           =   4035
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   585
      Left            =   7440
      MultiLine       =   -1  'True
      TabIndex        =   49
      Text            =   "frm180401.frx":00C9
      Top             =   2340
      Width           =   1755
   End
   Begin VB.TextBox txtType 
      Height          =   300
      Index           =   1
      Left            =   5490
      MaxLength       =   1
      TabIndex        =   19
      Top             =   2340
      Width           =   375
   End
   Begin VB.TextBox txtType 
      Height          =   300
      Index           =   3
      Left            =   5490
      MaxLength       =   1
      TabIndex        =   21
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtType 
      Height          =   300
      Index           =   0
      Left            =   2550
      MaxLength       =   1
      TabIndex        =   15
      Top             =   2340
      Width           =   375
   End
   Begin VB.TextBox txtType 
      Height          =   300
      Index           =   2
      Left            =   2550
      MaxLength       =   1
      TabIndex        =   17
      Top             =   2640
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   260
      Left            =   6630
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   540
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.TextBox txtB0101 
      Height          =   300
      Left            =   1470
      MaxLength       =   6
      TabIndex        =   0
      Top             =   540
      Width           =   735
   End
   Begin VB.TextBox txtDay 
      Height          =   300
      Index           =   3
      Left            =   7470
      MaxLength       =   2
      TabIndex        =   14
      Top             =   1740
      Width           =   375
   End
   Begin VB.TextBox txtDay 
      Height          =   300
      Index           =   1
      Left            =   7470
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1140
      Width           =   375
   End
   Begin VB.TextBox txtDay 
      Height          =   300
      Index           =   2
      Left            =   7470
      MaxLength       =   2
      TabIndex        =   12
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtDay 
      Height          =   300
      Index           =   0
      Left            =   7470
      MaxLength       =   2
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
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
            Picture         =   "frm180401.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180401.frx":041E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180401.frx":073A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180401.frx":0916
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180401.frx":0C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180401.frx":0F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180401.frx":126A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180401.frx":1586
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180401.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180401.frx":1BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180401.frx":1EDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9240
      _ExtentX        =   16298
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Bindings        =   "frm180401.frx":21F6
      Height          =   2390
      Left            =   30
      TabIndex        =   24
      Top             =   4320
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   4216
      _Version        =   393216
      Cols            =   30
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
      _Band(0).Cols   =   30
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   11
      Left            =   2970
      TabIndex        =   16
      Top             =   2340
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   13
      Left            =   2970
      TabIndex        =   18
      Top             =   2640
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   12
      Left            =   5910
      TabIndex        =   20
      Top             =   2340
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   14
      Left            =   5910
      TabIndex        =   22
      Top             =   2640
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   10
      Left            =   5940
      TabIndex        =   13
      Top             =   1740
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   8
      Left            =   5940
      TabIndex        =   9
      Top             =   1140
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   9
      Left            =   5940
      TabIndex        =   11
      Top             =   1440
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   7
      Left            =   5940
      TabIndex        =   7
      Top             =   840
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   6
      Left            =   3270
      TabIndex        =   6
      Top             =   1440
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   4
      Left            =   3270
      TabIndex        =   4
      Top             =   1140
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   2
      Left            =   3270
      TabIndex        =   2
      Top             =   840
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   5
      Left            =   1470
      TabIndex        =   5
      Top             =   1440
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   3
      Left            =   1470
      TabIndex        =   3
      Top             =   1140
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   1
      Left            =   1470
      TabIndex        =   1
      Top             =   840
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註：若有異動資料需通知人事室！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   21
      Left            =   120
      TabIndex        =   50
      Top             =   6750
      Width           =   3210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件職務代理人2：(1)代理類型"
      Height          =   180
      Index           =   20
      Left            =   60
      TabIndex        =   48
      Top             =   2700
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件職務代理人1：(1)代理類型"
      Height          =   180
      Index           =   19
      Left            =   60
      TabIndex        =   47
      Top             =   2400
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(2)代理類型"
      Height          =   180
      Index           =   18
      Left            =   4530
      TabIndex        =   46
      Top             =   2400
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(2)代理類型"
      Height          =   180
      Index           =   17
      Left            =   4530
      TabIndex        =   45
      Top             =   2700
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件代理類型：空白->所有案件，1->台灣案，2->非台灣案"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   16
      Left            =   60
      TabIndex        =   44
      Top             =   2130
      Width           =   4590
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   9180
      Y1              =   2060
      Y2              =   2060
   End
   Begin MSForms.Label LabDept 
      Height          =   290
      Left            =   4560
      TabIndex        =   43
      Top             =   600
      Width           =   2030
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門別："
      Height          =   180
      Index           =   15
      Left            =   3810
      TabIndex        =   41
      Top             =   600
      Width           =   720
   End
   Begin MSForms.Label txtB0101_2 
      Height          =   290
      Left            =   2250
      TabIndex        =   40
      Top             =   600
      Width           =   1490
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "天以上(不含)"
      Height          =   180
      Index           =   14
      Left            =   7890
      TabIndex        =   39
      Top             =   1830
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "天以上(不含)"
      Height          =   180
      Index           =   13
      Left            =   7890
      TabIndex        =   38
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "天以上(不含)"
      Height          =   180
      Index           =   12
      Left            =   7890
      TabIndex        =   37
      Top             =   1530
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "天以上(不含)"
      Height          =   180
      Index           =   0
      Left            =   7890
      TabIndex        =   36
      Top             =   900
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(2)"
      Height          =   180
      Index           =   11
      Left            =   3030
      TabIndex        =   35
      Top             =   1500
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(2)"
      Height          =   180
      Index           =   3
      Left            =   3030
      TabIndex        =   34
      Top             =   1200
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(2)"
      Height          =   180
      Index           =   2
      Left            =   3030
      TabIndex        =   33
      Top             =   900
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "審核主管1："
      Height          =   180
      Index           =   10
      Left            =   4920
      TabIndex        =   32
      Top             =   900
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "審核主管2："
      Height          =   180
      Index           =   9
      Left            =   4920
      TabIndex        =   31
      Top             =   1200
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "審核主管3："
      Height          =   180
      Index           =   8
      Left            =   4920
      TabIndex        =   30
      Top             =   1500
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "審核主管4："
      Height          =   180
      Index           =   7
      Left            =   4920
      TabIndex        =   29
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人1：(1)"
      Height          =   180
      Index           =   6
      Left            =   60
      TabIndex        =   28
      Top             =   900
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人2：(1)"
      Height          =   180
      Index           =   5
      Left            =   60
      TabIndex        =   27
      Top             =   1200
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人3：(1)"
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   26
      Top             =   1500
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   540
      TabIndex        =   25
      Top             =   600
      Width           =   900
   End
End
Attribute VB_Name = "frm180401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2023/12/20 修改抓新部門程式
'Memo By Sindy 2021/11/17 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Create by Sindy 2011/8/8
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 第一筆資料的Key
Dim m_FirstKEY(2) As String
' 最後一筆資料的Key
Dim m_LastKEY(2) As String
' 目前正在顯示的Key
Dim m_CurrKEY(2) As String
Dim i As Integer, j As Integer
Dim dblPrevRow As Double


Private Sub Combo2_GotFocus(Index As Integer)
   InverseTextBox Combo2(Index)
End Sub

Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
   If m_EditMode <> 0 And Combo2(Index).Text > "" And Len(Trim(Combo2(Index).Text)) = 5 Then
      '抓取員工姓名
      Combo2(Index).Text = SetCboStaffName(Combo2(Index).Text)
   End If
End Sub

Private Sub Combo2_Validate(Index As Integer, Cancel As Boolean)
Dim strText As String

   If m_EditMode <> 0 And Combo2(Index) <> "" Then
      If Index <> 0 Then
         If Left(Combo2(Index), 5) = txtB0101 Then
            MsgBox "不可為本人！", vbExclamation
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(Combo2(Index), 5)) = True Then
         Call Combo2_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查 員工不可為”不寄信”
      If ChkStaffST14(Left(Combo2(Index), 5)) = True Then
         Call Combo2_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查職代輸入順序
      If Index >= 1 And Index <= 6 Then
         If (Trim(Combo2(2)) <> "" And Trim(Combo2(1)) = "") Or _
            (Trim(Combo2(4)) <> "" And Trim(Combo2(3)) = "") Or _
            (Trim(Combo2(6)) <> "" And Trim(Combo2(5)) = "") Then
            MsgBox "請依序輸入職務代理人！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         If (Trim(Combo2(3)) <> "" And Trim(Combo2(1)) = "") Or _
            (Trim(Combo2(5)) <> "" And Trim(Combo2(3)) = "") Then
            MsgBox "請依序輸入職務代理人！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         If (Combo2(2) <> "" And Left(Combo2(2), 5) = Left(Combo2(1), 5)) Or _
            (Combo2(4) <> "" And Left(Combo2(4), 5) = Left(Combo2(3), 5)) Or _
            (Combo2(6) <> "" And Left(Combo2(6), 5) = Left(Combo2(5), 5)) Then
            MsgBox "資料重覆！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
      '檢查案件職代輸入順序
      If Index >= 11 And Index <= 14 Then
         If (Trim(Combo2(12)) <> "" And Trim(Combo2(11)) = "") Or _
            (Trim(Combo2(14)) <> "" And Trim(Combo2(13)) = "") Then
            MsgBox "請依序輸入案件職務代理人！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         If (Trim(Combo2(13)) <> "" And Trim(Combo2(11)) = "") Then
            MsgBox "請依序輸入案件職務代理人！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         If (Combo2(12) <> "" And Left(Combo2(12), 5) = Left(Combo2(11), 5)) Or _
            (Combo2(14) <> "" And Left(Combo2(14), 5) = Left(Combo2(13), 5)) Then
            MsgBox "資料重覆！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
      
      'Add By Sindy 2021/5/24
      '檢查居家職務輸入順序
      If Index >= 15 And Index <= 16 Then
         If (Trim(Combo2(16)) <> "" And Trim(Combo2(15)) = "") Then
            MsgBox "請依序輸入居家職務代理人！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         If (Combo2(16) <> "" And Left(Combo2(16), 5) = Left(Combo2(15), 5)) Then
            MsgBox "資料重覆！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
      '2021/5/24 END
      
      'Add By Sindy 2023/10/19
      '檢查工作評價主管輸入順序
      If Index >= 17 And Index <= 19 Then
         For i = 17 To 19
            If Trim(Combo2(i)) <> "" Then
               strText = Left(Combo2(i), 5)
            Else
               strText = ""
            End If
            For j = i + 1 To 19
               If strText = "" Then
                  '檢查是否依序輸入
                  If Trim(Combo2(j)) <> "" Then
                     MsgBox "請依序輸入工作評價主管！", vbExclamation
                     Combo2(j).SetFocus
                     Call Combo2_GotFocus(j)
                     Cancel = True
                     Exit Sub
                  End If
               Else
                  '資料重覆
                  If strText = Left(Combo2(j), 5) Then
                     MsgBox "資料重覆！", vbExclamation
                     Combo2(Index).SetFocus
                     Call Combo2_GotFocus(Index)
                     Cancel = True
                     Exit Sub
                  End If
               End If
            Next j
         Next i
      End If
      '2023/10/19 END
      
      '檢查審核主管輸入順序
      If Index >= 7 And Index <= 10 Then
         If (Trim(Combo2(8)) <> "" And Trim(Combo2(7)) = "") Or _
            (Trim(Combo2(9)) <> "" And Trim(Combo2(8)) = "") Or _
            (Trim(Combo2(10)) <> "" And Trim(Combo2(9)) = "") Then
            MsgBox "請依序輸入審核主管！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         If (Combo2(8) <> "" And Left(Combo2(8), 5) = Left(Combo2(7), 5)) Or _
            (Combo2(9) <> "" And Left(Combo2(9), 5) = Left(Combo2(8), 5)) Or _
            (Combo2(10) <> "" And Left(Combo2(10), 5) = Left(Combo2(9), 5)) Then
            MsgBox "資料重覆！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
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
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
      
   MoveFormToCenter Me
   
   SetDataListWidth
   
   ClearField
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   ReadAllData
   'OnAction vbKeyF4
   OnAction vbKeyF10
   SSTab1.Tab = 0 'Add By Sindy 2023/10/19
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180401 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
'   '上一筆資料列清除反白
'   If dblPrevRow > 0 Then
'      grd1.col = 2
'      grd1.row = dblPrevRow
'      For i = 0 To 1
'         grd1.col = i
'         grd1.CellBackColor = &H8000000F
'      Next i
'      For i = 2 To grd1.Cols - 1
'         grd1.col = i
'         grd1.CellBackColor = QBColor(15)
'      Next i
'   End If
'   '目前資料列反白
'   grd1.col = 0
'   grd1.row = grd1.MouseRow
'   dblPrevRow = grd1.row
'   For i = 0 To grd1.Cols - 1
'      grd1.col = i
'      grd1.CellBackColor = &HFFC0C0
'   Next i
   '查詢目前資料列
   'ShowCurrRecord grd1.TextMatrix(grd1.row, 16), grd1.TextMatrix(grd1.row, 17)
   'Modify By Sindy 2015/10/16
   If m_EditMode <> 2 Then
   '2015/10/16 END
      m_CurrKEY(0) = GRD1.TextMatrix(GRD1.row, 16)
      m_CurrKEY(1) = GRD1.TextMatrix(GRD1.row, 17)
      UpdateCtrlData
   End If
End If
GRD1.Visible = True
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

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim strST20 As String

TxtValidate = False

If txtB0101.Text = "" Then
    MsgBox "員工代號不可以空白！", vbExclamation
    txtB0101.SetFocus
    Exit Function
End If

If m_EditMode = 1 Then
   ' 檢查記錄是否已存在
   If IsRecordExist(txtB0101) = True Then
      MsgBox "該筆記錄已存在", vbOKOnly, "更新資料"
      'If txtB0101_2 = "" Then Call txtB0101_LostFocus
      txtB0101.SetFocus
      Exit Function
   End If
End If

'Modify By Sindy 2019/5/23
'員工為”不寄信”時,不檢查
If ChkStaffST14(txtB0101, False) = False Then
'2019/5/23 END
   If Combo2(1).Text = "" _
      And Combo2(2).Text = "" _
      And Combo2(3).Text = "" _
      And Combo2(4).Text = "" _
      And Combo2(5).Text = "" _
      And Combo2(6).Text = "" Then
       MsgBox "職務代理人不可以空白！", vbExclamation
       Combo2(1).SetFocus
       Exit Function
   End If
   
   'Modify By Sindy 2022/5/20
   '所長可以無審核主管
   strSql = "SELECT st20 FROM staff WHERE ST01='" & txtB0101 & "' "
   intI = 1: strST20 = ""
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Not IsNull(RsTemp("ST20")) Then strST20 = RsTemp("ST20")
   End If
   'If txtB0101 <> "76012" Then
   If strST20 <> "11" Then
   '2022/5/20 END
      If Combo2(7).Text = "" _
         And Combo2(8).Text = "" _
         And Combo2(9).Text = "" _
         And Combo2(10).Text = "" Then
          MsgBox "審核主管不可以空白！", vbExclamation
          Combo2(7).SetFocus
          Exit Function
      End If
   End If
End If

For i = 1 To Combo2.UBound
   Cancel = False
   Combo2_Validate i, Cancel
   If Cancel = True Then
      Exit Function
   End If
Next i

For i = 0 To txtDay.UBound
   Cancel = False
   txtDay_Validate i, Cancel
   If Cancel = True Then
      Exit Function
   End If
Next i

'Add By Sindy 2015/10/16
Cancel = False
txtB0124_Validate Cancel
If Cancel = True Then
   Exit Function
End If
Cancel = False
txtB0125_Validate Cancel
If Cancel = True Then
   Exit Function
End If
'2015/10/16 END

'Add By Sindy 2021/5/24
If ChkB0127.Value = 0 And Trim(Combo2(15).Text) <> "" Then
   If MsgBox("有輸入居家職務代理人，是否屬「居家無法連線」人員？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
      ChkB0127.SetFocus
      Exit Function
   End If
End If
'2021/5/24 END

TxtValidate = True
End Function

' 更新資料
Private Function SaveData(strEditMode As Integer) As Boolean
Dim strKEY01 As String, strKEY02 As String
   
On Error GoTo ErrHand
   
   SaveData = False
   
   strKEY01 = txtB0101
   'Modify By Sindy 2024/1/8
   'strKEY02 = GetStaffDepartment(txtB0101)
   strKEY02 = PUB_GetST93(txtB0101)
   '2024/1/8 END
   
   cnnConnection.BeginTrans
   '新增
   If strEditMode = 1 Then
      'Modify By Sindy 2015/10/16 +b0124,b0125,b0126
      'Modify By Sindy 2015/10/16 +b0127,b0128,b0129
      'Modfiy By Sindy 2023/10/19 +b0130,b0131,b0132
      strSql = "INSERT INTO ABS001(b0101,b0102,b0103,b0104,b0105,b0106,b0107,b0108,b0109,b0110,b0111,b0112,b0113,b0114,b0115,b0116,b0117,b0118,b0119,b0120,b0121,b0122,b0123,b0124,b0125,b0126,b0127,b0128,b0129,b0130,b0131,b0132)" & _
               " VALUES(" & CNULL(strKEY01) & _
                  "," & CNULL(Left(Trim(Combo2(1)), 5)) & "," & CNULL(Left(Trim(Combo2(2)), 5)) & _
                  "," & CNULL(Left(Trim(Combo2(3)), 5)) & "," & CNULL(Left(Trim(Combo2(4)), 5)) & _
                  "," & CNULL(Left(Trim(Combo2(5)), 5)) & "," & CNULL(Left(Trim(Combo2(6)), 5)) & _
                  "," & CNULL(Left(Trim(Combo2(7)), 5)) & "," & CNULL(Left(Trim(Combo2(8)), 5)) & _
                  "," & CNULL(Left(Trim(Combo2(9)), 5)) & "," & CNULL(Left(Trim(Combo2(10)), 5)) & _
                  "," & CNULL(txtDay(0)) & "," & CNULL(txtDay(1)) & _
                  "," & CNULL(txtDay(2)) & "," & CNULL(txtDay(3)) & _
                  "," & CNULL(txtType(0)) & "," & CNULL(Left(Trim(Combo2(11)), 5)) & _
                  "," & CNULL(txtType(1)) & "," & CNULL(Left(Trim(Combo2(12)), 5)) & _
                  "," & CNULL(txtType(2)) & "," & CNULL(Left(Trim(Combo2(13)), 5)) & _
                  "," & CNULL(txtType(3)) & "," & CNULL(Left(Trim(Combo2(14)), 5)) & _
                  "," & CNULL(txtB0124) & "," & CNULL(txtB0125) & "," & CNULL(IIf(ChkEMailAll.Value = 1, "Y", "")) & _
                  "," & CNULL(IIf(ChkB0127.Value = 1, "Y", "")) & _
                  "," & CNULL(Left(Trim(Combo2(15)), 5)) & "," & CNULL(Left(Trim(Combo2(16)), 5)) & _
                  "," & CNULL(Left(Trim(Combo2(17)), 5)) & "," & CNULL(Left(Trim(Combo2(18)), 5)) & "," & CNULL(Left(Trim(Combo2(19)), 5)) & ")"
   '修改
   ElseIf strEditMode = 2 Then
      'Modify By Sindy 2015/10/16 +b0124,b0125,b0126
      'Modify By Sindy 2015/10/16 +b0127,b0128,b0129
      'Modfiy By Sindy 2023/10/19 +b0130,b0131,b0132
      strSql = "UPDATE ABS001 SET " & _
                  "B0102=" & CNULL(Left(Trim(Combo2(1)), 5)) & ",B0103=" & CNULL(Left(Trim(Combo2(2)), 5)) & _
                  ",B0104=" & CNULL(Left(Trim(Combo2(3)), 5)) & ",B0105=" & CNULL(Left(Trim(Combo2(4)), 5)) & _
                  ",B0106=" & CNULL(Left(Trim(Combo2(5)), 5)) & ",B0107=" & CNULL(Left(Trim(Combo2(6)), 5)) & _
                  ",B0108=" & CNULL(Left(Trim(Combo2(7)), 5)) & ",B0109=" & CNULL(Left(Trim(Combo2(8)), 5)) & _
                  ",B0110=" & CNULL(Left(Trim(Combo2(9)), 5)) & ",B0111=" & CNULL(Left(Trim(Combo2(10)), 5)) & _
                  ",B0112=" & CNULL(txtDay(0)) & ",B0113=" & CNULL(txtDay(1)) & _
                  ",B0114=" & CNULL(txtDay(2)) & ",B0115=" & CNULL(txtDay(3)) & _
                  ",B0116=" & CNULL(txtType(0)) & ",B0117=" & CNULL(Left(Trim(Combo2(11)), 5)) & _
                  ",B0118=" & CNULL(txtType(1)) & ",B0119=" & CNULL(Left(Trim(Combo2(12)), 5)) & _
                  ",B0120=" & CNULL(txtType(2)) & ",B0121=" & CNULL(Left(Trim(Combo2(13)), 5)) & _
                  ",B0122=" & CNULL(txtType(3)) & ",B0123=" & CNULL(Left(Trim(Combo2(14)), 5)) & _
                  ",B0124=" & CNULL(txtB0124) & ",B0125=" & CNULL(txtB0125) & _
                  ",B0126=" & CNULL(IIf(ChkEMailAll.Value = 1, "Y", "")) & _
                  ",B0127=" & CNULL(IIf(ChkB0127.Value = 1, "Y", "")) & _
                  ",B0128=" & CNULL(Left(Trim(Combo2(15)), 5)) & ",B0129=" & CNULL(Left(Trim(Combo2(16)), 5)) & _
                  ",B0130=" & CNULL(Left(Trim(Combo2(17)), 5)) & ",B0131=" & CNULL(Left(Trim(Combo2(18)), 5)) & _
                  ",B0132=" & CNULL(Left(Trim(Combo2(19)), 5)) & _
               " WHERE B0101=" & CNULL(strKEY01)
   End If
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   
   If (strKEY01 < m_FirstKEY(0)) Or (strKEY01 > m_LastKEY(0)) Then
      RefreshRange
   End If
   ShowCurrRecord strKEY01, strKEY02
   
   SaveData = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 更新失敗！" & vbCrLf & Err.Description
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strKEY01 As String
   
On Error GoTo ErrHand
   
   DelRecord = False
   
   strKEY01 = txtB0101
   
   cnnConnection.BeginTrans
   
   strSql = "DELETE FROM ABS001 WHERE b0101 = " & CNULL(strKEY01)
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strKEY01 As String
Dim strKEY02 As String
   
   QueryRecord = False
   strKEY01 = txtB0101
   'Modify By Sindy 2024/1/8
   'strKEY02 = GetStaffDepartment(txtB0101)
   strKEY02 = PUB_GetST93(txtB0101)
   '2024/1/8 END
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
      QueryRecord = True
      UpdateCtrlData
'      ReadAllData
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
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         If SaveData(m_EditMode) = True Then
             RefreshRange
             ReadAllData
             SetKeyReadOnly True
         Else
             Exit Function
         End If
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         If SaveData(m_EditMode) = False Then Exit Function
         ReadAllData
         SetKeyReadOnly True
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
            ReadAllData
            SetKeyReadOnly True
         Else
            Exit Function
         End If
      Case 4: '查詢
         If txtB0101 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
            SetKeyReadOnly True
         Else
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
      Case 0, 1, 4: If Me.txtB0101.Visible = True Then txtB0101.SetFocus
      Case 2: If Me.Combo2(1).Visible = True Then Combo2(1).SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
   'strSql = "SELECT * FROM ABS001 WHERE b0101=" & CNULL(strKEY01)
   strSql = "SELECT * FROM ABS001,staff WHERE b0101=st01(+) and b0101=" & CNULL(strKEY01) & " and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101)))"
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
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
      'strSql = "SELECT B0101,A0901 FROM ABS001,STAFF,ACC090 WHERE B0101=ST01(+) and ST03=A0901(+) and B0101='" & m_CurrKEY(0) & "' "
      strSql = "SELECT B0101,A0921 FROM ABS001,STAFF,ACC090NEW WHERE B0101=ST01(+) and ST93=A0921(+) and B0101='" & m_CurrKEY(0) & "' and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
      'strSql = "SELECT B0101,A0901 FROM ABS001,STAFF,ACC090 WHERE B0101=ST01(+) and ST03=A0901(+) order by A0901 asc,B0101 asc "
      strSql = "SELECT B0101,A0921 FROM ABS001,STAFF,ACC090NEW WHERE B0101=ST01(+) and ST93=A0921(+) and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) order by A0921 asc,B0101 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
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
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
   'strSql = "SELECT B0101,A0901 FROM ABS001,STAFF,ACC090 WHERE B0101=ST01(+) and ST03=A0901(+) and A0901||B0101<'" & m_CurrKEY(1) & m_CurrKEY(0) & "' order by A0901 desc,B0101 desc "
   strSql = "SELECT B0101,A0921 FROM ABS001,STAFF,ACC090NEW WHERE B0101=ST01(+) and ST93=A0921(+) and A0921||B0101<'" & m_CurrKEY(1) & m_CurrKEY(0) & "' and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) order by A0921 desc,B0101 desc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
   'strSql = "SELECT B0101,A0901 FROM ABS001,STAFF,ACC090 WHERE B0101=ST01(+) and ST03=A0901(+) order by A0901 asc,B0101 asc "
   strSql = "SELECT B0101,A0921 FROM ABS001,STAFF,ACC090NEW WHERE B0101=ST01(+) and ST93=A0921(+) and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) order by A0921 asc,B0101 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
   'strSql = "SELECT B0101,A0901 FROM ABS001,STAFF,ACC090 WHERE B0101=ST01(+) and ST03=A0901(+) and A0901||B0101>'" & m_CurrKEY(1) & m_CurrKEY(0) & "' order by A0901 asc,B0101 asc "
   strSql = "SELECT B0101,A0921 FROM ABS001,STAFF,ACC090NEW WHERE B0101=ST01(+) and ST93=A0921(+) and A0921||B0101>'" & m_CurrKEY(1) & m_CurrKEY(0) & "' and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) order by A0921 asc,B0101 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
   'strSql = "SELECT B0101,A0901 FROM ABS001,STAFF,ACC090 order by A0901 asc,B0101 asc "
   strSql = "SELECT B0101,A0921 FROM ABS001,STAFF,ACC090NEW where (st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) order by A0921 asc,B0101 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
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
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         SetKeyReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         Call txtB0101_LostFocus
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
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
         If OnWork = True Then
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
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  SetKeyReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               SetKeyReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   txtB0101.Locked = bEnable
   If bEnable Then txtB0101.BackColor = &H8000000F Else txtB0101.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   For i = 1 To Combo2.UBound
      Combo2(i).Locked = bEnable
      If bEnable Then Combo2(i).BackColor = &H8000000F Else Combo2(i).BackColor = &H80000005
   Next i
   For i = 0 To txtDay.UBound
      txtDay(i).Locked = bEnable
      If bEnable Then txtDay(i).BackColor = &H8000000F Else txtDay(i).BackColor = &H80000005
   Next i
   For i = 0 To txtType.UBound
      txtType(i).Locked = bEnable
      If bEnable Then txtType(i).BackColor = &H8000000F Else txtType(i).BackColor = &H80000005
   Next i
   'Add By Sindy 2015/10/16
   ChkEMailAll.Enabled = Not bEnable
   'If bEnable Then ChkEMailAll.BackColor = &H8000000F Else ChkEMailAll.BackColor = &H80000005
   txtB0124.Locked = bEnable
   If bEnable Then txtB0124.BackColor = &H8000000F Else txtB0124.BackColor = &H80000005
   txtB0125.Locked = bEnable
   If bEnable Then txtB0125.BackColor = &H8000000F Else txtB0125.BackColor = &H80000005
   '2015/10/16 END
   ChkB0127.Enabled = Not bEnable 'Add By Sindy 2021/5/24
End Sub

Private Sub ClearField()
   LabDept.Caption = Empty
   txtB0101 = Empty
   txtB0101_2 = Empty
   For i = 1 To Combo2.UBound
      'Combo2(i).Clear
      Combo2(i).Text = Empty
   Next i
   For i = 0 To txtDay.UBound
      txtDay(i).Text = Empty
   Next i
   For i = 0 To txtType.UBound
      txtType(i).Text = Empty
   Next i
   'Add By Sindy 2015/10/16
   ChkEMailAll.Value = 0
   txtB0124 = Empty
   txtB0125 = Empty
   '2015/10/16 END
   ChkB0127.Value = 0 'Add By Sindy 2021/5/24
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub ReadAllData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   GRD1.Rows = 2
   GRD1.Clear
   GRD1.FixedCols = 0
   'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
   strSql = "select A0922,s0.ST02 " & _
            ",s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02,s7.ST02,B0112,s8.ST02,B0113,s9.ST02,B0114,s10.ST02,B0115,B0101,A0921,s0.ST04 " & _
            ",B0116,s11.ST02,B0118,s12.ST02,B0120,s13.ST02,B0122,s14.ST02,B0124,B0125,B0126 " & _
            "from ABS001,ACC090NEW,STAFF s0 " & _
            ",STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF s7,STAFF s8,STAFF s9,STAFF s10,STAFF s11,STAFF s12,STAFF s13,STAFF s14 " & _
            "where B0101=s0.ST01(+) " & _
            "and s0.ST93=A0921(+) " & _
            "and B0102=s1.ST01(+) " & _
            "and B0103=s2.ST01(+) " & _
            "and B0104=s3.ST01(+) " & _
            "and B0105=s4.ST01(+) " & _
            "and B0106=s5.ST01(+) " & _
            "and B0107=s6.ST01(+) " & _
            "and B0108=s7.ST01(+) " & _
            "and B0109=s8.ST01(+) " & _
            "and B0110=s9.ST01(+) " & _
            "and B0111=s10.ST01(+) " & _
            "and B0117=s11.ST01(+) " & _
            "and B0119=s12.ST01(+) " & _
            "and B0121=s13.ST01(+) " & _
            "and B0123=s14.ST01(+) " & _
            "and (s0.st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) " & _
            "order by A0921,B0101"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      GRD1.FixedCols = 2
   End If
   rsTmp.Close
   SetDataListWidth
   GetSelChage
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub GetSelChage()
GRD1.Visible = False
If GRD1.Rows - 1 > 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 2
      GRD1.row = dblPrevRow
      For i = 0 To 1
         GRD1.col = i
         GRD1.CellBackColor = &H8000000F
      Next i
      For i = 2 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   End If
   '尋找目前資料列
   For j = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(j, 16) = m_CurrKEY(0) Then
         GRD1.col = 0
         GRD1.row = j
         dblPrevRow = GRD1.row
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
'            grd1.TopRow = j
         Exit For
      End If
   Next j
End If
GRD1.Visible = True
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   ClearField
   
   'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
   strSql = "SELECT ABS001.*,A0921,A0922,A0925,s1.ST02 s1_ST02 " & _
            ",s2.ST02 s2_ST02,s3.ST02 s3_ST02,s4.ST02 s4_ST02,s5.ST02 s5_ST02 " & _
            ",s6.ST02 s6_ST02,s7.ST02 s7_ST02,s8.ST02 s8_ST02,s9.ST02 s9_ST02 " & _
            ",s10.ST02 s10_ST02,s11.ST02 s11_ST02 " & _
            ",s17.ST02 s17_ST02,s19.ST02 s19_ST02,s21.ST02 s21_ST02,s23.ST02 s23_ST02,s28.ST02 s28_ST02,s29.ST02 s29_ST02 " & _
            "FROM ABS001,STAFF s1,ACC090NEW " & _
            ",STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF s7,STAFF s8,STAFF s9,STAFF s10,STAFF s11 " & _
            ",STAFF s17,STAFF s19,STAFF s21,STAFF s23,STAFF s28,STAFF s29 " & _
            "WHERE B0101=s1.ST01(+) and s1.ST93=A0921(+) and B0101='" & m_CurrKEY(0) & "' " & _
            "and B0102=s2.ST01(+) and B0103=s3.ST01(+) and B0104=s4.ST01(+) " & _
            "and B0105=s5.ST01(+) and B0106=s6.ST01(+) and B0107=s7.ST01(+) " & _
            "and B0108=s8.ST01(+) and B0109=s9.ST01(+) and B0110=s10.ST01(+) " & _
            "and B0111=s11.ST01(+) and B0117=s17.ST01(+) and B0119=s19.ST01(+) " & _
            "and B0121=s21.ST01(+) and B0123=s23.ST01(+) and B0128=s28.ST01(+) and B0129=s29.ST01(+) " & _
            "and (s1.st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) "
   'strSql = "SELECT ABS001.*,A0901,A0902,A0911 FROM ABS001,STAFF,ACC090 WHERE B0101=ST01(+) and ST03=A0901(+) and B0101='" & m_CurrKEY(0) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If m_EditMode = 1 And txtB0101.Enabled = True Then
         '員工代號欄位為空白,使用者自行輸入欲新增的員工代號
      Else
         If IsNull(rsTmp.Fields("A0922")) = False Then LabDept.Caption = rsTmp.Fields("A0921") & "  " & rsTmp.Fields("A0922")
         If IsNull(rsTmp.Fields("B0101")) = False Then txtB0101 = rsTmp.Fields("B0101"): txtB0101_2 = rsTmp.Fields("s1_ST02")
      End If
      If IsNull(rsTmp.Fields("B0102")) = False Then Combo2(1).Text = Left(Trim(rsTmp.Fields("B0102")) & Space(5), 7) & rsTmp.Fields("s2_ST02")
      If IsNull(rsTmp.Fields("B0103")) = False Then Combo2(2).Text = Left(Trim(rsTmp.Fields("B0103")) & Space(5), 7) & rsTmp.Fields("s3_ST02")
      If IsNull(rsTmp.Fields("B0104")) = False Then Combo2(3).Text = Left(Trim(rsTmp.Fields("B0104")) & Space(5), 7) & rsTmp.Fields("s4_ST02")
      If IsNull(rsTmp.Fields("B0105")) = False Then Combo2(4).Text = Left(Trim(rsTmp.Fields("B0105")) & Space(5), 7) & rsTmp.Fields("s5_ST02")
      If IsNull(rsTmp.Fields("B0106")) = False Then Combo2(5).Text = Left(Trim(rsTmp.Fields("B0106")) & Space(5), 7) & rsTmp.Fields("s6_ST02")
      If IsNull(rsTmp.Fields("B0107")) = False Then Combo2(6).Text = Left(Trim(rsTmp.Fields("B0107")) & Space(5), 7) & rsTmp.Fields("s7_ST02")
      If IsNull(rsTmp.Fields("B0108")) = False Then Combo2(7).Text = Left(Trim(rsTmp.Fields("B0108")) & Space(5), 7) & rsTmp.Fields("s8_ST02")
      If IsNull(rsTmp.Fields("B0109")) = False Then Combo2(8).Text = Left(Trim(rsTmp.Fields("B0109")) & Space(5), 7) & rsTmp.Fields("s9_ST02")
      If IsNull(rsTmp.Fields("B0110")) = False Then Combo2(9).Text = Left(Trim(rsTmp.Fields("B0110")) & Space(5), 7) & rsTmp.Fields("s10_ST02")
      If IsNull(rsTmp.Fields("B0111")) = False Then Combo2(10).Text = Left(Trim(rsTmp.Fields("B0111")) & Space(5), 7) & rsTmp.Fields("s11_ST02")
      If IsNull(rsTmp.Fields("B0112")) = False Then txtDay(0) = rsTmp.Fields("B0112")
      If IsNull(rsTmp.Fields("B0113")) = False Then txtDay(1) = rsTmp.Fields("B0113")
      If IsNull(rsTmp.Fields("B0114")) = False Then txtDay(2) = rsTmp.Fields("B0114")
      If IsNull(rsTmp.Fields("B0115")) = False Then txtDay(3) = rsTmp.Fields("B0115")
      If IsNull(rsTmp.Fields("B0116")) = False Then txtType(0) = rsTmp.Fields("B0116")
      If IsNull(rsTmp.Fields("B0117")) = False Then Combo2(11).Text = Left(Trim(rsTmp.Fields("B0117")) & Space(5), 7) & rsTmp.Fields("s17_ST02")
      If IsNull(rsTmp.Fields("B0118")) = False Then txtType(1) = rsTmp.Fields("B0118")
      If IsNull(rsTmp.Fields("B0119")) = False Then Combo2(12).Text = Left(Trim(rsTmp.Fields("B0119")) & Space(5), 7) & rsTmp.Fields("s19_ST02")
      If IsNull(rsTmp.Fields("B0120")) = False Then txtType(2) = rsTmp.Fields("B0120")
      If IsNull(rsTmp.Fields("B0121")) = False Then Combo2(13).Text = Left(Trim(rsTmp.Fields("B0121")) & Space(5), 7) & rsTmp.Fields("s21_ST02")
      If IsNull(rsTmp.Fields("B0122")) = False Then txtType(3) = rsTmp.Fields("B0122")
      If IsNull(rsTmp.Fields("B0123")) = False Then Combo2(14).Text = Left(Trim(rsTmp.Fields("B0123")) & Space(5), 7) & rsTmp.Fields("s23_ST02")
      'Add By Sindy 2015/10/16
      If IsNull(rsTmp.Fields("B0124")) = False Then txtB0124 = rsTmp.Fields("B0124")
      If IsNull(rsTmp.Fields("B0125")) = False Then txtB0125 = rsTmp.Fields("B0125")
      ChkEMailAll.Value = 0
      If IsNull(rsTmp.Fields("B0126")) = False Then
         If rsTmp.Fields("B0126") = "Y" Then
            ChkEMailAll.Value = 1
         End If
      End If
      '2015/10/16 END
      'Add By Sindy 2021/5/24
      ChkB0127.Value = 0
      If IsNull(rsTmp.Fields("B0127")) = False Then
         If rsTmp.Fields("B0127") = "Y" Then
            ChkB0127.Value = 1
         End If
      End If
      If IsNull(rsTmp.Fields("B0128")) = False Then Combo2(15).Text = Left(Trim(rsTmp.Fields("B0128")) & Space(5), 7) & rsTmp.Fields("s28_ST02")
      If IsNull(rsTmp.Fields("B0129")) = False Then Combo2(16).Text = Left(Trim(rsTmp.Fields("B0129")) & Space(5), 7) & rsTmp.Fields("s29_ST02")
      '2021/5/24 END
      'Add By Sindy 2023/10/19
      If IsNull(rsTmp.Fields("B0130")) = False Then Combo2(17).Text = Left(Trim(rsTmp.Fields("B0130")) & Space(5), 7) & GetPrjSalesNM(Trim(rsTmp.Fields("B0130")))
      If IsNull(rsTmp.Fields("B0131")) = False Then Combo2(18).Text = Left(Trim(rsTmp.Fields("B0131")) & Space(5), 7) & GetPrjSalesNM(Trim(rsTmp.Fields("B0131")))
      If IsNull(rsTmp.Fields("B0132")) = False Then Combo2(19).Text = Left(Trim(rsTmp.Fields("B0132")) & Space(5), 7) & GetPrjSalesNM(Trim(rsTmp.Fields("B0132")))
      '2023/10/19 END
   End If
   rsTmp.Close
   GetSelChage
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
   'strSql = "select B0101,A0901 from ABS001,STAFF,ACC090 where B0101=ST01(+) and ST03=A0901(+) order by A0901 asc,B0101 asc "
   strSql = "select B0101,A0921 from ABS001,STAFF,ACC090NEW where B0101=ST01(+) and ST93=A0921(+) and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) order by A0921 asc,B0101 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then m_FirstKEY(0) = Trim(rsTmp.Fields(0))
      If IsNull(rsTmp.Fields(1)) = False Then m_FirstKEY(1) = Trim(rsTmp.Fields(1))
   End If
   rsTmp.Close
   
   'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
   'strSql = "select B0101,A0901 from ABS001,STAFF,ACC090 where B0101=ST01(+) and ST03=A0901(+) order by A0901 desc,B0101 desc "
   strSql = "select B0101,A0921 from ABS001,STAFF,ACC090NEW where B0101=ST01(+) and ST93=A0921(+) and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=B0101 and sc02=(select max(sc02) from Staff_Change where sc01=B0101))) order by A0921 desc,B0101 desc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then m_LastKEY(0) = Trim(rsTmp.Fields(0))
      If IsNull(rsTmp.Fields(1)) = False Then m_LastKEY(1) = Trim(rsTmp.Fields(1))
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
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

Private Sub SetDataListWidth()
GRD1.row = 0
GRD1.col = 0: GRD1.Text = "部門"
GRD1.ColWidth(0) = 1000
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 1: GRD1.Text = "員工姓名"
GRD1.ColWidth(1) = 850
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 2: GRD1.Text = "職代一(1)"
GRD1.ColWidth(2) = 850
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 3: GRD1.Text = "職代一(2)"
GRD1.ColWidth(3) = 850
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 4: GRD1.Text = "職代二(1)"
GRD1.ColWidth(4) = 850
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 5: GRD1.Text = "職代二(2)"
GRD1.ColWidth(5) = 850
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 6: GRD1.Text = "職代三(1)"
GRD1.ColWidth(6) = 850
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 7: GRD1.Text = "職代三(2)"
GRD1.ColWidth(7) = 850
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 8: GRD1.Text = "審核主管1"
GRD1.ColWidth(8) = 900
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 9: GRD1.Text = "天數"
GRD1.ColWidth(9) = 500
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 10: GRD1.Text = "審核主管2"
GRD1.ColWidth(10) = 900
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 11: GRD1.Text = "天數"
GRD1.ColWidth(11) = 500
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 12: GRD1.Text = "審核主管3"
GRD1.ColWidth(12) = 900
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 13: GRD1.Text = "天數"
GRD1.ColWidth(13) = 500
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 14: GRD1.Text = "審核主管4"
GRD1.ColWidth(14) = 900
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 15: GRD1.Text = "天數"
GRD1.ColWidth(15) = 500
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 16: GRD1.Text = "B0101"
GRD1.ColWidth(16) = 0
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 17: GRD1.Text = "A0921"
GRD1.ColWidth(17) = 0
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 18: GRD1.Text = "ST04"
GRD1.ColWidth(18) = 0
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 19: GRD1.Text = "(1-1)類型"
GRD1.ColWidth(19) = 600
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 20: GRD1.Text = "案件職代一(1)"
GRD1.ColWidth(20) = 1000
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 21: GRD1.Text = "(1-2)類型"
GRD1.ColWidth(21) = 600
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 22: GRD1.Text = "案件職代一(2)"
GRD1.ColWidth(22) = 1000
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 23: GRD1.Text = "(2-1)類型"
GRD1.ColWidth(23) = 600
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 24: GRD1.Text = "案件職代二(1)"
GRD1.ColWidth(24) = 1000
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 25: GRD1.Text = "(2-2)類型"
GRD1.ColWidth(25) = 600
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 26: GRD1.Text = "案件職代二(2)"
GRD1.ColWidth(26) = 1000
GRD1.CellAlignment = flexAlignLeftCenter
'Add By Sindy 2015/10/16
GRD1.col = 27: GRD1.Text = "核准後通知人員"
GRD1.ColWidth(27) = 1200
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 28: GRD1.Text = "核准通知是否含簽核主管"
GRD1.ColWidth(28) = 1200
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 29: GRD1.Text = "案件EMail職代全發"
GRD1.ColWidth(29) = 1200
GRD1.CellAlignment = flexAlignLeftCenter
'2015/10/16 END
End Sub

Private Sub txtB0101_GotFocus()
   InverseTextBox txtB0101
End Sub

Private Sub txtB0101_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtB0101_LostFocus()
Dim Rs As New ADODB.Recordset
Dim strText(20) As String
Dim strA0911 As String, strA0925 As String
   
   If m_EditMode <> 0 And txtB0101 <> "" Then
      txtB0101_2 = GetStaffName(txtB0101, True)
      LabDept.Caption = PUB_GetST93(txtB0101) & "  " & GetDeptNameA0922(txtB0101)
      strA0911 = GetStaffA0911(txtB0101, strA0925)
      
      For i = 1 To Combo2.UBound
         strText(i) = Combo2(i).Text
         'Modify By Sindy 2023/12/20
         'Combo2(i).Clear
         'Combo2(i).AddItem ""
         Call SetB1003Combo(Combo2(i), strA0911, strA0925)
         Combo2(i).Text = strText(i)
         '2023/12/20 END
      Next i
   End If
End Sub

Private Sub txtB0101_Validate(Cancel As Boolean)
Dim Rs As New ADODB.Recordset
   
   If txtB0101.Text = "" Then txtB0101_2 = ""
   
   If m_EditMode <> 0 And txtB0101 <> "" Then
      ' 檢查員工編號規則
      If ChkStaffID(txtB0101) Then
         Call txtB0101_GotFocus
         Cancel = True
         Exit Sub
      End If
      txtB0101_2 = GetStaffName(txtB0101, True)
      If txtB0101_2 = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         Call txtB0101_GotFocus
         Cancel = True
         Exit Sub
      End If
      
      'Add By Sindy 2012/8/28 只查詢在職或留職停薪人員的資料
      Rs.CursorLocation = adUseClient
      strSql = "select ST01,ST02 " & _
               "From staff " & _
               "where st01='" & txtB0101 & "' " & _
               "and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=ST01 and sc02=(select max(sc02) from Staff_Change where sc01=ST01)))"
      Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If Rs.RecordCount = 0 Then
         MsgBox "此員工已離職！", vbInformation
         Call txtB0101_GotFocus
         Cancel = True
         If Rs.State <> adStateClosed Then Rs.Close
         Exit Sub
      End If
      If Rs.State <> adStateClosed Then Rs.Close
      '2012/8/28 End
      
      If m_EditMode = 1 Then
         ' 檢查記錄是否已存在
         If IsRecordExist(txtB0101) = True Then
            MsgBox "該筆記錄已存在", vbInformation
            Call txtB0101_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtDay_GotFocus(Index As Integer)
   InverseTextBox txtDay(Index)
End Sub

Private Sub txtDay_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtDay_Validate(Index As Integer, Cancel As Boolean)
Dim intDay As Integer
Dim i As Integer

If m_EditMode <> 0 And txtDay(Index) <> "" Then
   If CheckLengthIsOK(txtDay(Index), txtDay(Index).MaxLength) = False Then
      Call txtDay_GotFocus(Index)
      Cancel = True
      Exit Sub
   End If
   'Add By Sindy 2024/5/8 審核主管=Combo2(7 To 10)
   If Combo2(Index + 7) = "" Then
      MsgBox "有輸入天數，審核主管不可空白！", vbExclamation
      Combo2(Index + 7).SetFocus
      Cancel = True
      Exit Sub
   End If
   intDay = txtDay(Index)
   For i = Index + 1 To 3
      If Combo2(i + 7).Text <> "" Then
         If txtDay(i) = "" Then
            MsgBox Replace(Combo2(i + 7).Text, " ", "") & " 的天數不可空白！", vbExclamation
            txtDay(i).SetFocus
            Cancel = True
            Exit Sub
         ElseIf txtDay(i) < intDay Then
            MsgBox Replace(Combo2(i + 7).Text, " ", "") & " 的天數不可小於" & intDay & "！", vbExclamation
            txtDay(i).SetFocus
            Cancel = True
            Exit Sub
         End If
      End If
   Next i
   '2024/5/8 END
End If
End Sub

Private Sub txtType_GotFocus(Index As Integer)
   InverseTextBox txtType(Index)
End Sub

Private Sub txtType_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtType_Validate(Index As Integer, Cancel As Boolean)
If m_EditMode <> 0 And txtType(Index) <> "" Then
   If CheckLengthIsOK(txtType(Index), txtType(Index).MaxLength) = False Then
      Call txtType_GotFocus(Index)
      Cancel = True
      Exit Sub
   End If
End If
End Sub

'Add By Sindy 2015/10/16
Private Sub txtB0124_GotFocus()
   InverseTextBox txtB0124
End Sub
Private Sub txtB0124_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtB0124_Validate(Cancel As Boolean)
Dim strTemp As Variant
Dim strData As String
   If txtB0124.Text = "" Then Exit Sub
   If m_EditMode <> 0 And txtB0124 <> "" Then
      Cancel = False
      strTemp = Split(txtB0124, ";")
      For i = 0 To UBound(strTemp)
         strData = strTemp(i)
         If strData = txtB0101 Then
            MsgBox "不可為本人！", vbExclamation
            txtB0124.SetFocus
            txtB0124_GotFocus
            Cancel = True
            Exit Sub
         End If
         '檢查人員是否存在或離職
         If ChkStaffST04(strData) = True Then
            txtB0124.SetFocus
            txtB0124_GotFocus
            Cancel = True
            Exit Sub
         End If
         For j = i + 1 To UBound(strTemp)
            If strData = strTemp(j) Then
               MsgBox "人員編號不可重覆！", vbExclamation
               txtB0124.SetFocus
               txtB0124_GotFocus
               Cancel = True
               Exit Sub
            End If
         Next j
      Next i
   End If
End Sub
Private Sub txtB0125_GotFocus()
   InverseTextBox txtB0125
End Sub
Private Sub txtB0125_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtB0125_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   Cancel = False
   If IsEmptyText(txtB0125) = False Then
      Select Case txtB0125
         Case "N":
         Case Else:
            strTit = "檢核資料"
            strMsg = "核准通知是否含簽核主管，只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            txtB0125.SetFocus
            txtB0125_GotFocus
            Cancel = True
            Exit Sub
      End Select
   End If
End Sub
'2015/10/16 END
