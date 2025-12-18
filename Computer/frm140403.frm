VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140403 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶/代理人聯絡人資料維護"
   ClientHeight    =   6600
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   9156
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9156
   Begin VB.CommandButton cmdContact 
      Caption         =   "新增"
      Height          =   285
      Index           =   1
      Left            =   6570
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdContact 
      Caption         =   "刪除"
      Height          =   285
      Index           =   3
      Left            =   8100
      TabIndex        =   29
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdContact 
      Caption         =   "加入"
      Height          =   285
      Index           =   2
      Left            =   7335
      TabIndex        =   28
      Top             =   2400
      Width           =   735
   End
   Begin VB.Frame fraContact 
      Height          =   3945
      Left            =   270
      TabIndex        =   34
      Top             =   2610
      Width           =   8610
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "上傳相片"
         Height          =   276
         Left            =   1920
         Style           =   1  '圖片外觀
         TabIndex        =   58
         Top             =   144
         Width           =   948
      End
      Begin VB.TextBox txtPCC20 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   790
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         Height          =   675
         Left            =   4275
         TabIndex        =   36
         Top             =   1020
         Width           =   4155
         Begin VB.CommandButton cmdRemoveDept 
            Caption         =   "移除 ->"
            Height          =   255
            Left            =   45
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   390
            Width           =   735
         End
         Begin VB.CommandButton cmdAddDept 
            Caption         =   "<- 新增"
            Height          =   255
            Left            =   45
            TabIndex        =   10
            Top             =   120
            Width           =   735
         End
         Begin VB.ComboBox cboDept 
            Height          =   300
            ItemData        =   "frm140403.frx":0000
            Left            =   810
            List            =   "frm140403.frx":0002
            TabIndex        =   8
            Text            =   "cboDept"
            Top             =   120
            Width           =   3285
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   6
            Left            =   810
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   420
            Visible         =   0   'False
            Width           =   3285
            VariousPropertyBits=   671105051
            MaxLength       =   70
            Size            =   "5794;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Frame Frame2 
         Height          =   675
         Left            =   4275
         TabIndex        =   37
         Top             =   1620
         Width           =   4155
         Begin VB.CommandButton cmdAddTit 
            Caption         =   "<- 新增"
            Height          =   255
            Left            =   45
            TabIndex        =   15
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton cmdRemoveTit 
            Caption         =   "移除 ->"
            Height          =   255
            Left            =   45
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   390
            Width           =   735
         End
         Begin VB.ComboBox cboTitle 
            Height          =   300
            ItemData        =   "frm140403.frx":0004
            Left            =   810
            List            =   "frm140403.frx":0006
            TabIndex        =   13
            Text            =   "cboTitle"
            Top             =   120
            Width           =   3300
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   7
            Left            =   810
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   420
            Visible         =   0   'False
            Width           =   3285
            VariousPropertyBits=   671105051
            MaxLength       =   70
            Size            =   "5794;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Frame Frame3 
         Height          =   705
         Left            =   2330
         TabIndex        =   38
         Top             =   3210
         Width           =   2610
         Begin VB.CommandButton cmdRemove 
            Caption         =   "移除 ->"
            Height          =   255
            Index           =   1
            Left            =   45
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   390
            Width           =   735
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "<- 新增"
            Height          =   255
            Index           =   1
            Left            =   45
            TabIndex        =   25
            Top             =   120
            Width           =   735
         End
         Begin VB.TextBox txtUserNo 
            Height          =   270
            Index           =   1
            Left            =   810
            MaxLength       =   6
            TabIndex        =   23
            Top             =   120
            Width           =   680
         End
         Begin MSForms.TextBox txtPCC 
            Height          =   300
            Index           =   12
            Left            =   810
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   390
            Visible         =   0   'False
            Width           =   1770
            VariousPropertyBits=   671105051
            MaxLength       =   70
            Size            =   "3122;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblName 
            Height          =   300
            Index           =   1
            Left            =   1530
            TabIndex        =   56
            Top             =   150
            Width           =   870
            VariousPropertyBits=   27
            Caption         =   "lblFM2"
            Size            =   "1535;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.ListBox lstTitle 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         ItemData        =   "frm140403.frx":0008
         Left            =   1080
         List            =   "frm140403.frx":000F
         MultiSelect     =   1  '簡易多重選取
         Sorted          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1710
         Width           =   3180
      End
      Begin VB.ListBox lstDept 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         ItemData        =   "frm140403.frx":001D
         Left            =   1080
         List            =   "frm140403.frx":0024
         MultiSelect     =   1  '簡易多重選取
         Sorted          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1110
         Width           =   3180
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   24
         Left            =   4530
         TabIndex        =   22
         Top             =   2940
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   25
         Left            =   5610
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2340
         Width           =   2055
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "3625;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   5
         Left            =   1080
         TabIndex        =   6
         Top             =   790
         Width           =   3180
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5609;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   3
         Left            =   1080
         TabIndex        =   4
         Top             =   470
         Width           =   3180
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "5609;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   4
         Left            =   5310
         TabIndex        =   5
         Top             =   470
         Width           =   3180
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "5609;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   10
         Left            =   6912
         TabIndex        =   21
         Top             =   2640
         Width           =   288
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   8
         Left            =   1080
         TabIndex        =   17
         Top             =   2340
         Width           =   3180
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5609;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   615
         Index           =   13
         Left            =   5490
         TabIndex        =   27
         Top             =   3270
         Width           =   3060
         VariousPropertyBits=   -1466941413
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "5397;1085"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   11
         Left            =   1080
         TabIndex        =   19
         Top             =   2640
         Width           =   1035
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1826;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   9
         Left            =   4530
         TabIndex        =   20
         Top             =   2640
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   2
         Left            =   1245
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   150
         Width           =   600
         VariousPropertyBits=   671105055
         Size            =   "1058;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   560
         Index           =   1
         Left            =   1050
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   3300
         Width           =   1290
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2275;988"
         MatchEntry      =   0
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID1 
         Height          =   300
         Left            =   2976
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   156
         Width           =   5508
         VariousPropertyBits=   -2147467233
         BackColor       =   16777215
         Size            =   "9716;529"
         Caption         =   "LblFM2"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄專利雙週報：      （N:不寄)"
         Height          =   180
         Index           =   1
         Left            =   2895
         TabIndex        =   53
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "名片臨時編號："
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   14
         Left            =   4320
         TabIndex        =   52
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "名稱( 中 )："
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   51
         Top             =   850
         Width           =   930
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "名稱( 英 )："
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   50
         Top             =   530
         Width           =   930
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "名稱( 日 )："
         Height          =   180
         Index           =   2
         Left            =   4365
         TabIndex        =   49
         Top             =   530
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄電子報：      （Y:寄／ N:不寄)"
         Height          =   180
         Index           =   11
         Left            =   5664
         TabIndex        =   48
         Top             =   2700
         Width           =   2892
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "職稱："
         Height          =   180
         Index           =   3
         Left            =   135
         TabIndex        =   47
         Top             =   1710
         Width           =   540
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "部門："
         Height          =   180
         Index           =   4
         Left            =   135
         TabIndex        =   46
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發人員："
         Height          =   180
         Index           =   12
         Left            =   135
         TabIndex        =   45
         Top             =   3300
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   14
         Left            =   4950
         TabIndex        =   44
         Top             =   3300
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄台一雜誌：          （N:不寄)"
         Height          =   180
         Index           =   7
         Left            =   2892
         TabIndex        =   43
         Top             =   2700
         Width           =   2700
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人編號："
         Height          =   180
         Index           =   7
         Left            =   135
         TabIndex        =   42
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "E-MAIL："
         Height          =   180
         Index           =   5
         Left            =   135
         TabIndex        =   41
         Top             =   2400
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發日期：                         ( 西元 )"
         Height          =   180
         Index           =   9
         Left            =   135
         TabIndex        =   40
         Top             =   2700
         Width           =   2595
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "相關聯絡人編號："
         Height          =   180
         Index           =   8
         Left            =   4365
         TabIndex        =   39
         Top             =   850
         Width           =   1440
      End
   End
   Begin VB.TextBox txtPCU 
      Enabled         =   0   'False
      Height          =   276
      Index           =   2
      Left            =   2070
      MaxLength       =   1
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   675
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8415
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
            Picture         =   "frm140403.frx":0031
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140403.frx":034D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140403.frx":0669
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140403.frx":0845
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140403.frx":0B61
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140403.frx":0E7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140403.frx":1199
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140403.frx":14B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140403.frx":17D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140403.frx":1AED
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140403.frx":1E09
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPCU 
      Height          =   276
      Index           =   1
      Left            =   1005
      MaxLength       =   8
      TabIndex        =   0
      Top             =   675
      Width           =   1092
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   9156
      _ExtentX        =   16150
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
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7425
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   572
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm140403.frx":2125
      Height          =   1065
      Left            =   270
      TabIndex        =   32
      Top             =   1320
      Width           =   8625
      _ExtentX        =   15219
      _ExtentY        =   1884
      _Version        =   393216
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "X1"
         Caption         =   "編號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "PCC03"
         Caption         =   "英文名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "PCC04"
         Caption         =   "日文名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "PCC05"
         Caption         =   "中文名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "PCC06"
         Caption         =   "部門"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PCC07"
         Caption         =   "職稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "PCC08"
         Caption         =   "EMail"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "PCC09"
         Caption         =   "寄台一雜誌"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "PCC10"
         Caption         =   "寄電子報"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "PCC11"
         Caption         =   "開發日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "PCC12"
         Caption         =   "開發人員"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "PCC25"
         Caption         =   "名片臨時編號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   315
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   2580.095
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1607.811
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1272.189
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1595.906
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   887.811
         EndProperty
         BeginProperty Column09 
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column11 
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox textName 
      Height          =   600
      Left            =   2370
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   660
      Width           =   6525
      VariousPropertyBits=   -2147467233
      BackColor       =   16777215
      Size            =   "11509;1058"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "表格內的編號欄若有＊號則表示該連絡人已離職"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   10
      Left            =   270
      TabIndex        =   33
      Top             =   2400
      Width           =   3780
   End
   Begin VB.Label Label1 
      Caption         =   "編號："
      Height          =   210
      Index           =   0
      Left            =   315
      TabIndex        =   31
      Top             =   675
      Width           =   585
   End
End
Attribute VB_Name = "frm140403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/10 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、textCUID1、textName、lblName(1)、lblUsers(1)、txtPCC(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2007/11/28
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim TF_PCC As Integer
Dim strTmp As String
Dim oText As Control
Dim idx As Integer
Dim m_arrConRefList() As String '相關聯絡人資料
Dim rsContact As ADODB.Recordset
Dim rsContactOld As ADODB.Recordset
Dim rsContactSim As ADODB.Recordset
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_SHOWDROPDOWN = &H14F
Dim m_iConEditMode As Integer '聯絡人狀態 1:新增 2:修改
Dim m_bReadGrid As Boolean '是否要讀取被點選聯絡人資料
Dim stFormName As String

Private Sub cboDept_GotFocus()
   If cboDept.Locked = False Then
      CloseIme
      SendMessage cboDept.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

Private Sub cboTitle_GotFocus()
   If cboTitle.Locked = False Then
      CloseIme
      SendMessage cboTitle.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

'新增開發人員
Private Sub cmdAdd_Click(Index As Integer)
   AddlstUsers Index
   txtPCC(12) = ComposeListX(Index)
   txtUserNo(Index).SetFocus
End Sub

'新增部門
Private Sub cmdAddDept_Click()
   If InStr(cboDept, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
      cboDept.SetFocus
      Exit Sub
   End If
   AddLstFrmCbo cboDept, lstDept
   txtPCC(6) = ComposeList(lstDept)
   cboDept.SetFocus
End Sub

'新增職稱
Private Sub cmdAddTit_Click()
   If InStr(cboTitle, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
      cboTitle.SetFocus
      Exit Sub
   End If
   AddLstFrmCbo cboTitle, lstTitle
   txtPCC(7) = ComposeList(lstTitle)
   cboTitle.SetFocus
End Sub

'移除開發人員
Private Sub cmdRemove_Click(Index As Integer)
   RemovelstUsers Index
   txtPCC(12) = ComposeListX(Index)
   txtUserNo(Index).SetFocus
End Sub
'聯絡人
Private Sub cmdContact_Click(Index As Integer)
   Dim bDupeCheck As Boolean, sPCC(1 To 5) As String
   Select Case Index
      Case 1 '新增
         ClearField1
         txtPCC(2).Text = getNewNo
         txtPCC(3).SetFocus
         txtPCC(9) = "N" 'Add by Morgan 2007/12/19 預設不寄台一雜誌
         
         'Modify by Morgan 2007/12/19
         '打字室輸入時預設:1.開發日期=20070101,2.開發人員=81040
         If Pub_StrUserSt03 = "M13" Then
            txtPCC(11) = 20070101
            txtPCC(12) = "81040"
            SetlstUsers 1, txtPCC(12)
            
         '新增時開發日期預設當天
         'Modify by Morgan 2008/3/13 因為X,Y多為舊資料故不預設
         'Else
            'txtPCC(11) = strSrvDate(1)
         End If
         m_iConEditMode = 1
         
      Case 2 '加入
         If TxtValidate1 = True Then
            UpdateContact
            '檢查是否有名稱相同的聯絡人
            bDupeCheck = False
            If m_iConEditMode = 1 Then
               bDupeCheck = True
            ElseIf txtPCC20 = "" Then
               bDupeCheck = ContNameChanged
            End If
            If bDupeCheck = True Then
               sPCC(1) = txtPCU(1)
               sPCC(2) = txtPCC(2)
               sPCC(3) = txtPCC(3)
               sPCC(4) = txtPCC(4)
               sPCC(5) = txtPCC(5)
               If PUB_DupeContactCheck(sPCC, rsContactSim) = True Then
                  Me.Tag = ""
                  Set frm140402_1.grdDataList.Recordset = rsContactSim
                  Set frm140402_1.fmParent = Me
                  frm140402_1.Show vbModal
                  UpdateConRefList txtPCC(2), Me.Tag
               End If
            End If
            
            ClearField1
         End If
            
      Case 3 '刪除
         If txtPCC(2) <> "" Then
            If Not (rsContact.EOF Or rsContact.BOF) Then
               If PUB_PCCDelCheck(txtPCU(1), txtPCU(2), txtPCC(2)) = True Then
                  rsContact.Delete
                  rsContact.UpdateBatch
                  RemoveConRefList
                  ClearField1
               End If
            End If
         End If
   End Select
End Sub

Private Sub cmdRemoveDept_Click()
   RemoveList lstDept
   txtPCC(6) = ComposeList(lstDept)
End Sub

Private Sub cmdRemoveTit_Click()
   RemoveList lstTitle
   txtPCC(7) = ComposeList(lstTitle)
End Sub

Private Sub DataGrid1_Click()
   '點選同一列可能不會觸發RowColChange
   If DataGrid1.col = -1 Then
      ReadContact
   End If
   m_bReadGrid = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If m_bReadGrid = True Then
      ReadContact
   End If
End Sub

Private Sub DataGrid1_Validate(Cancel As Boolean)
   m_bReadGrid = False
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   textName.BackColor = &H8000000F
   textCUID1.BackColor = &H8000000F
   
   AddCombo 1
   AddCombo 2

   'Modify by Morgan 2008/6/16 預設查詢
   'm_EditMode = 0
   'ShowRecord -2
   'SetInputEntry
   'UpdateToolbarState
   stFormName = Me.Caption
   OnAction vbKeyF4
End Sub

Private Sub Form_Initialize()
   strExc(0) = "select * from PotCustCont where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   TF_PCC = RsTemp.Fields.Count
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
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
         
      Case vbKeyInsert
         If cmdContact(2).Enabled = True Then
            cmdContact_Click 2
         End If
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140403 = Nothing
End Sub


Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      ' 修改
      Case 2: OnAction vbKeyF3
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

Private Sub txtPCC_GotFocus(Index As Integer)
   Select Case Index
      Case 4, 5, 13
         OpenIme
         
      Case Else
         CloseIme
         
   End Select
   TextInverse txtPCC(Index)
End Sub

'Modified by Lydia 2022/01/10 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txtPCC_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 8
         'Modified by Lydia 2022/01/10 +val()
         PUB_EMailFilter Val(KeyAscii) 'Added by Morgan 2011/11/30 Email輸入字元檢查
      'Modified by Lydia 2020/11/25 拿掉”是否寄電子報”
      'Case 9, 10
      'Modify By Sindy 2021/11/30 + 24
      Case 9, 24
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
      'Modify By Sindy 2018/5/31 +25
      Case 25
         KeyAscii = UpperCase(KeyAscii)
      'Added by Lydia 2020/11/25 是否寄電子報欄位，請新增Y:寄
      Case 10
         KeyAscii = UpperCase(KeyAscii)
         'Added by Lydia 2025/05/15 只有代理人才可輸入Y
         If Left(txtPCU(1), 1) <> "Y" Then
            If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
               KeyAscii = 0
               Beep
            End If
         Else
         'end 2025/05/15
            If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
               KeyAscii = 0
               Beep
            End If
         End If
   End Select
End Sub

Private Sub txtPCC_Validate(Index As Integer, Cancel As Boolean)
   Dim iLen As Integer
   Select Case Index
      Case 8
         If txtPCC(Index) <> "" Then
            If InStr(1, txtPCC(Index), "@") = 0 Then
                MsgBox "Mail 必需要有 @ 符號！"
                Cancel = True
            ElseIf InStr(1, txtPCC(Index), ",") > 0 Or InStr(1, txtPCC(Index), "[") > 0 Or InStr(1, txtPCC(Index), "]") > 0 Or InStr(1, txtPCC(Index), "!") > 0 Or InStr(1, txtPCC(Index), "(") > 1 Or InStr(1, txtPCC(Index), ")") > 0 Or InStr(1, txtPCC(Index), "=") > 0 Or InStr(1, txtPCC(Index), "\") > 0 Or InStr(1, txtPCC(Index), "/") > 0 Or InStr(1, txtPCC(Index), "<") > 0 Or InStr(1, txtPCC(Index), ">") > 0 Or InStr(1, txtPCC(Index), "~") > 0 Or InStr(1, txtPCC(Index), "$") > 0 Or InStr(1, txtPCC(Index), "%") > 0 Or InStr(1, txtPCC(Index), "^") > 0 Or InStr(1, txtPCC(Index), "&") > 0 Or InStr(1, txtPCC(Index), "*") > 0 Then
                MsgBox "Mail 不允許有下列符號！" & vbCrLf & ",、[、]、!、(、)、=、\、/、<、>、~、$、%、^、&、* "
                Cancel = True
            End If
         End If
         
      Case 11
         If txtPCC(Index) <> "" Then
            If CheckIsDate(txtPCC(Index)) = False Then
               txtPCC_GotFocus Index
               Cancel = True
            End If
         End If
   End Select
   If Cancel = False Then
      '欄位長度檢查
      Select Case Index
         '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
         'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
         'Case 4, 5, 6, 7, 13
         'Modified by Lydia 2018/07/04 日文名稱、中文名稱、部門、職稱、備註可輸入造字
         'Case 5: iLen = 30
         'Case 4, 6, 7, 13
         'end 2017/06/14
         Case 4, 5, 6, 7, 13
            iLen = txtPCC(Index).MaxLength - 1
         Case Else
            iLen = txtPCC(Index).MaxLength
      End Select
      
      If Not CheckLengthIsOK(txtPCC(Index), iLen) Then
         Cancel = True
      End If
   End If
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF3 ' 修改
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF4 ' 查詢
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
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  txtPCU(1) = txtPCU(1).Tag
                  txtPCU(2) = txtPCU(2).Tag
                  m_EditMode = 0
                  SetInputEntry
                  ShowRecord
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               txtPCU(1) = txtPCU(1).Tag
               txtPCU(2) = txtPCU(2).Tag
               If txtPCU(1) <> "" Then
                  SetInputEntry
                  ShowRecord
               Else
                  ClearField
               End If
               UpdateToolbarState
         End Select
         
      Case vbKeyEscape ' 離開
         Unload Me
         Exit Sub
   End Select
   
   Select Case m_EditMode
      Case 1
         Me.Caption = stFormName & "(新增)"
      Case 2
         Me.Caption = stFormName & "(修改)"
      Case 4
         Me.Caption = stFormName & "(查詢)"
      Case Else
         Me.Caption = stFormName
   End Select
End Sub

Private Sub ClearField()
   txtPCU(1) = ""
   textName = ""
   OpenContactTable
   ClearField1
End Sub

Private Sub ClearField1()
   For Each oText In txtPCC
      oText.Text = Empty
   Next
   lstDept.Clear
   lstTitle.Clear
   txtUserNo(1) = ""
   lblName(1) = ""
   lstUsers(1).Clear
   lstUsers(1).Tag = "" 'Added by Lydia 2022/01/10
   cboDept = ""
   cboTitle = ""
   textCUID1 = ""
   txtPCC20 = ""
   'Added by Lydia 2024/05/10
   Command1.Visible = False
   Command1.Caption = "上傳相片"
   Command1.BackColor = &H8080FF     '紅色
   'end 2024/05/10
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   cmdContact(1).Enabled = Not bLocked
   cmdContact(2).Enabled = Not bLocked
   cmdContact(3).Enabled = Not bLocked
   For Each oText In txtPCC
      oText.Locked = bLocked
   Next
   txtPCC(2).Locked = True
   cboDept.Locked = bLocked
   
   Frame2.Visible = Not bLocked
   Frame3.Visible = Not bLocked
   Frame4.Visible = Not bLocked
   
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
         If m_bUpdate And txtPCU(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtPCU(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtPCU(1) <> "" Then
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
   
   'Add by Morgan 2008/7/4 不允許上下筆
   TBar1.Buttons(6).Enabled = False
   TBar1.Buttons(7).Enabled = False
   TBar1.Buttons(8).Enabled = False
   TBar1.Buttons(9).Enabled = False
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   If Me.Visible = True Then
      Select Case m_EditMode
         Case 2
            txtPCU(1).Locked = True
         Case 4
            txtPCU(1).Locked = False
            txtPCU(1).SetFocus
         Case Else
            txtPCU(1).Locked = True
            txtPCU(1).SetFocus
      End Select
   End If
End Sub

Private Function TxtValidate() As Boolean
   
   Dim Cancel As Boolean, ii As Integer, jj As Integer

   For Each oText In txtPCU
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         Cancel = False
         txtPCU_Validate idx, Cancel
         If Cancel = True Then
            txtPCU(idx).SetFocus
            txtPCU_GotFocus idx
            Exit Function
         End If
      End If
   Next
   '查詢
   If m_EditMode = 4 Then
      If txtPCU(1) = "" Then
         ShowMsg "請輸入欲查詢之客戶編號 !"
         txtPCU(1).SetFocus
         txtPCU_GotFocus 1
         Exit Function
      'Add by Morgan 2008/7/4 控制只可查詢智權人員為國外部之客戶
      ElseIf Left(txtPCU(1), 1) = "X" Then
         strExc(0) = "select st03,cu13 from customer,staff where cu01='" & Left(txtPCU(1) & "000", 8) & "' and cu02='0' and st01(+)=cu13"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modify by Morgan 2009/1/10 沒有智權人員的不控制
            'If Left("" & RsTemp(0), 1) <> "F" Then
            If Not IsNull(RsTemp(1)) And Left("" & RsTemp(0), 1) <> "F" Then
               ShowMsg "請輸入智權人員為國外部之客戶編號 !"
               txtPCU(1).SetFocus
               txtPCU_GotFocus 1
               Exit Function
            End If
         Else
            ShowMsg "客戶編號不存在!"
            txtPCU(1).SetFocus
            txtPCU_GotFocus 1
            Exit Function
         End If
      End If
   End If
   
   TxtValidate = True
   
End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   '更新聯絡人資料
   If rsContact.RecordCount = 0 Then
      If rsContactOld.RecordCount > 0 Then
         '清除相關聯絡人資料
         stSQL = "update potcustcont set pcc20=null where substr(pcc20,1,8)='" & txtPCU(1) & "'"
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intI
         'Added by Lydia 2024/05/10 刪除聯絡人相片
         strExc(0) = "select imgbytefile.* from potcustcont,imgbytefile where pcc01='" & txtPCU(1) & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               PUB_DelFtpFile2 RsTemp.Fields("IBF01") & "-" & RsTemp.Fields("IBF02") & "-" & RsTemp.Fields("IBF03") & "-" & RsTemp.Fields("IBF04") & "-" & RsTemp.Fields("IBF05"), , UCase("ImgByteFile")
               stSQL = "DELETE FROM IMGBYTEFILE WHERE IBF01='" & RsTemp.Fields("IBF01") & "' AND IBF02='" & RsTemp.Fields("IBF02") & "' AND IBF03='" & RsTemp.Fields("IBF03") & "' AND IBF04='" & RsTemp.Fields("IBF04") & "' AND IBF05='" & RsTemp.Fields("IBF05") & "' "
               cnnConnection.Execute stSQL
               RsTemp.MoveNext
            Loop
         End If
         'end 2024/05/10
         '刪除聯絡人資料
         stSQL = "delete from potcustcont where pcc01='" & txtPCU(1) & "'"
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intI
      End If
   Else
      '刪除聯絡人(原來的編號在新的聯絡人資料中找不到的)
      With rsContactOld
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            rsContact.MoveFirst
            rsContact.Find "PCC02='" & .Fields("PCC02") & "'"
            If rsContact.EOF Then
               '清除相關聯絡人資料
               stSQL = "update potcustcont set pcc20=null where pcc20='" & txtPCU(1) & .Fields("PCC02") & "'"
               Pub_SeekTbLog stSQL
               cnnConnection.Execute stSQL, intI
               'Added by Lydia 2024/05/10 刪除聯絡人相片
               strExc(0) = "select imgbytefile.* from potcustcont,imgbytefile where pcc01='" & txtPCU(1) & "' and pcc02='" & .Fields("PCC02") & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                     PUB_DelFtpFile2 RsTemp.Fields("IBF01") & "-" & RsTemp.Fields("IBF02") & "-" & RsTemp.Fields("IBF03") & "-" & RsTemp.Fields("IBF04") & "-" & RsTemp.Fields("IBF05"), , UCase("ImgByteFile")
                     stSQL = "DELETE FROM IMGBYTEFILE WHERE IBF01='" & RsTemp.Fields("IBF01") & "' AND IBF02='" & RsTemp.Fields("IBF02") & "' AND IBF03='" & RsTemp.Fields("IBF03") & "' AND IBF04='" & RsTemp.Fields("IBF04") & "' AND IBF05='" & RsTemp.Fields("IBF05") & "' "
                     cnnConnection.Execute stSQL
                     RsTemp.MoveNext
                  Loop
               End If
               'end 2024/05/10
               '刪除聯絡人資料
               stSQL = "delete from potcustcont where pcc01='" & txtPCU(1) & "' and pcc02='" & .Fields("PCC02") & "'"
               Pub_SeekTbLog stSQL
               cnnConnection.Execute stSQL, intI
            End If
            .MoveNext
         Loop
      End If
      End With
      '更新(新增)連絡人
      With rsContact
      .MoveFirst
      Do While Not .EOF
         If rsContactOld.RecordCount = 0 Then
            bAddNew = True
         Else
            rsContactOld.MoveFirst
            rsContactOld.Find "PCC02='" & .Fields("PCC02") & "'"
            If rsContactOld.EOF Then
               bAddNew = True
            Else
               bAddNew = False
            End If
         End If
         
         '新增
         If bAddNew = True Then
            stCols = "PCC01"
            stValues = "'" & txtPCU(1) & "'"
            For idx = 1 To .Fields.Count - 3
               If .Fields(idx) <> "" Then
                  stCols = stCols & "," & .Fields(idx).Name
                  If .Fields(idx).Name = "PCC11" Then
                     stValues = stValues & "," & .Fields(idx)
                  Else
                     stValues = stValues & ",'" & ChgSQL(.Fields(idx)) & "'"
                  End If
               End If
            Next
            
            stSQL = "INSERT INTO PotCustCont (" & stCols & ") Values (" & stValues & ")"
            Pub_SeekTbLog stSQL
            cnnConnection.Execute stSQL, intI
         '修改
         Else
            bDifference = False
            stSet = ""
            For idx = 2 To .Fields.Count - 3
               If "" & .Fields(idx) <> "" & rsContactOld.Fields(idx) Then
                  bDifference = True
                  If .Fields(idx).Name = "PCC11" Then
                     stSet = stSet & "," & .Fields(idx).Name & "=" & CNULL(.Fields(idx), True)
                  Else
                     stSet = stSet & "," & .Fields(idx).Name & "=" & CNULL(ChgSQL(.Fields(idx)))
                  End If
               End If
            Next
            If bDifference = True Then
               stSet = Mid(stSet, 2)
               stSQL = "begin user_data.user_enabled:=1; Update PotCustCont set " & stSet & " where PCC01='" & txtPCU(1) & "' and PCC02='" & .Fields("PCC02") & "'; end;"
               Pub_SeekTbLog stSQL
               cnnConnection.Execute stSQL
            End If
         End If
         .MoveNext
      Loop
      End With
   End If
   '更新聯絡人相關編號
   UpdateRefContact txtPCU(1)
   
   cnnConnection.CommitTrans
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
      
       Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtPCU(1).SetFocus
               txtPCU_GotFocus 1
            End If
         End If
         
   End Select
End Function

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
   
   Dim stPCU01 As String
   Dim stPCU02 As String
   Dim adoRst As New ADODB.Recordset
   
   stPCU01 = Left(txtPCU(1) & "000", 8)
   stPCU02 = "0"
      
   Select Case p_iWay
      Case 0 '當筆
         If Left(stPCU01, 1) = "X" Then
            strExc(0) = "SELECT CU01 NO,CU04 CN,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) EN,CU06 JN" & _
               " FROM Customer WHERE CU01 = '" & stPCU01 & "' AND CU02 = '0'"
         Else
            strExc(0) = "SELECT FA01 NO,FA04 CN,rtrim(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
               " FROM FAgent WHERE FA01 = '" & stPCU01 & "' AND FA02 = '0'"
         End If
         
      Case -2 '首筆
         If Left(stPCU01, 1) = "X" Then
            strExc(0) = "SELECT CU01 NO,CU04 CN,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) EN,CU06 JN" & _
               " FROM Customer WHERE CU02='0' order by CU01 ASC"
         Else
            strExc(0) = "SELECT FA01 NO,FA04 CN,rtrim(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
               " FROM FAgent WHERE FA02='0' order by FA01 ASC"
         End If
         
      Case -1 '前筆
         If Left(stPCU01, 1) = "X" Then
            strExc(0) = "SELECT CU01 NO,CU04 CN,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) EN,CU06 JN" & _
               " FROM Customer WHERE CU01<'" & stPCU01 & "' AND CU02='0' order by CU01 DESC"
         Else
            strExc(0) = "SELECT FA01 NO,FA04 CN,rtrim(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
               " FROM FAgent WHERE FA01<'" & stPCU01 & "' AND FA02='0' order by FA01 DESC"
         End If
        
      Case 1 '後筆
         If Left(stPCU01, 1) = "X" Then
            strExc(0) = "SELECT CU01 NO,CU04 CN,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) EN,CU06 JN" & _
               " FROM Customer WHERE CU01 >'" & stPCU01 & "' AND CU02='0' order by CU01 ASC"
         Else
            strExc(0) = "SELECT FA01 NO,FA04 CN,rtrim(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
               " FROM FAgent WHERE FA01>'" & stPCU01 & "' AND FA02='0' order by FA01 ASC"
         End If
         
      Case 2 '末筆
         If Left(stPCU01, 1) = "X" Then
            strExc(0) = "SELECT CU01 NO,CU04 CN,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90) EN,CU06 JN" & _
               " FROM Customer WHERE CU02='0' order by CU01 DESC"
         Else
            strExc(0) = "SELECT FA01 NO,FA04 CN,rtrim(FA05||' '||FA63||' '||FA64||' '||FA65) EN,FA06 JN" & _
               " FROM FAgent WHERE FA02='0' order by FA01 DESC"
         End If
   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      txtPCU(1) = "" & adoRst.Fields("NO")
      txtPCU(1).Tag = txtPCU(1)
      textName = "中: " & adoRst.Fields("CN") & _
         vbCrLf & "英: " & adoRst.Fields("EN") & _
         vbCrLf & "日: " & adoRst.Fields("JN")
      OpenContactTable
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
      txtPCU(1).SetFocus
      txtPCU_GotFocus 1
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
   'Modified by Lydia 2024/05/10 String(10->String(6
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(6, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   
   lstUsers(p_idx).Clear
   lstUsers(p_idx).Tag = "" 'Added by Lydia 2022/01/10
   If p_stNums <> "" Then
      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         arrID = Split(p_stNums, ",")
         With RsTemp
         '照原順序排
         For intI = UBound(arrID) To LBound(arrID) Step -1
            .MoveFirst
            Do While Not .EOF
               If .Fields("st01") = arrID(intI) Then
                  lstUsers(p_idx).AddItem "" & .Fields(1), 0
                  'Modify by Morgan 2011/8/26 員工編號已可非數字需做轉換
                  'Modified by Lydia 2022/01/10 改成Form 2.0沒有ItemData屬性
                  'lstUsers(p_idx).ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
                   lstUsers(p_idx).Tag = .Fields(0) & "," & lstUsers(p_idx).Tag
                  .MoveLast
               End If
               .MoveNext
            Loop
         Next
         End With
      End If
      lstUsers(p_idx).ListIndex = 0 'Add by Sindy 2024/12/27
   End If
End Sub

Private Sub SetList(oList As ListBox, p_stData As String)
   Dim arrID
   oList.Clear
   If p_stData <> "" Then
      arrID = Split(p_stData, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

Private Sub AddlstUsers(p_idx As Integer)
   Dim idx As Integer, bFound As Boolean
   'Modify by Morgan 2011/8/26 員工編號已可非數字需做轉換
   If txtUserNo(p_idx) <> "" And lblName(p_idx) <> "" Then
      For idx = 0 To lstUsers(p_idx).ListCount - 1
         'Modified by Lydia 2022/01/10 改成Form 2.0
         'If lstUsers(p_idx).ItemData(idx) = PUB_Id2Num(txtUserNo(p_idx)) Then
         '   MsgBox "員工已存在於開發人員清單中！"
         '   txtUserNo(p_idx).SetFocus
         '   txtUserNo_GotFocus p_idx
         '   bFound = True
         '   Exit For
         'End If
         If InStr(lstUsers(p_idx).Tag, txtUserNo(p_idx)) > 0 Then
            MsgBox "員工已存在於開發人員清單中！"
            txtUserNo(p_idx).SetFocus
            txtUserNo_GotFocus p_idx
            bFound = True
            Exit For 'Add By Sindy 2024/12/27
         End If
         'end 2022/01/10
      Next
      If bFound = False Then
'         lstUsers(p_idx).AddItem lblName(p_idx), 0 'Modify By Sindy 2024/12/27 mark
'         'Modified by Lydia 2022/01/10 改成Form 2.0
'         'lstUsers(p_idx).ItemData(0) = PUB_Id2Num(txtUserNo(p_idx))
         lstUsers(p_idx).Tag = txtUserNo(p_idx) & "," & lstUsers(p_idx).Tag
         SetlstUsers 1, lstUsers(p_idx).Tag 'Modify By Sindy 2024/12/27
         txtUserNo(p_idx) = ""
         lblName(p_idx) = ""
      End If
   End If
End Sub

Private Sub RemovelstUsers(p_idx As Integer)
   'Modified by Lydia 2022/01/10 改成Form 2.0
   'Dim idx As Integer, ii As Integer
   'If lstUsers(p_idx).ListCount > 0 Then
   '   ii = 0
   '   For idx = 0 To lstUsers(p_idx).ListCount - 1
   '      If lstUsers(p_idx).Selected(ii) = True Then
   '         lstUsers(p_idx).RemoveItem ii
   '         ii = ii - 1
   '      End If
   '      ii = ii + 1
   '   Next
   'End If
   lstUsers(p_idx).Tag = PUB_RemoveListBox2(lstUsers(p_idx), lstUsers(p_idx).Tag)
   'end 2022/01/10
End Sub

Private Sub AddLstFrmCbo(oCombo As ComboBox, oList As ListBox)
   Dim idx As Integer, bFound As Boolean
   
   If oCombo <> "" Then
      For idx = 0 To oList.ListCount - 1
         If oList.List(idx) = oCombo Then
            MsgBox "資料已存在！"
            oCombo.SetFocus
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         oList.AddItem oCombo, 0
         oCombo = ""
      End If
   End If
End Sub

Private Sub RemoveList(oList As ListBox)
   Dim idx As Integer, ii As Integer
   If oList.ListCount > 0 Then
      ii = 0
      For idx = 0 To oList.ListCount - 1
         If oList.Selected(ii) = True Then
            oList.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Next
   End If
End Sub

Private Sub txtPCU_GotFocus(Index As Integer)
   TextInverse txtPCU(Index)
End Sub

Private Sub txtPCU_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPCU_Validate(Index As Integer, Cancel As Boolean)
   If Not IsEmptyText(txtPCU(1)) Then
      If Mid(txtPCU(1), 1, 1) <> "X" And Mid(txtPCU(1), 1, 1) <> "Y" Then
         Cancel = True
         MsgBox "客戶/代理人編號必須為X/Y開頭", vbCritical + vbOKOnly, "檢核資料"
         txtPCU(1).Text = ""
         txtPCU_GotFocus 1
         Exit Sub
      End If
      
      If Len(txtPCU(1)) < 6 Then
         Cancel = True
         MsgBox "客戶/代理人編號請至少輸入六碼", vbCritical + vbOKOnly, "檢核資料"
         txtPCU_GotFocus 1
         Exit Sub
      End If
   End If
End Sub

Private Sub txtUserNo_Change(Index As Integer)
   Dim strTempName As String
   If Len(txtUserNo(Index)) = 5 Then
      If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
         lblName(Index) = strTempName
      End If
   Else
      lblName(Index) = ""
   End If
End Sub

Private Sub txtUserNo_GotFocus(Index As Integer)
   TextInverse txtUserNo(Index)
End Sub

'Add By Sindy 2010/11/26
Private Sub txtUserNo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUserNo_Validate(Index As Integer, Cancel As Boolean)
   Dim strTempName As String
   If txtUserNo(Index).Visible = True Then
      If txtUserNo(Index) <> "" And lblName(Index) = "" Then
         If Len(txtUserNo(Index)) = 5 Then
            If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
               lblName(Index) = strTempName
            End If
         End If
         If lblName(Index) = "" Then
            MsgBox "員工編號輸入錯誤！", vbExclamation
            Cancel = True
         End If
      End If
   End If
End Sub

Private Sub OpenContactTable()
   
On Error GoTo Checking

   If txtPCU(1) <> "" Then
      strExc(0) = "select PCC.*,decode(pcc20,null,'　','＊')||pcc02 X1,decode(pcc20,null,'',substr(pcc20,1,8)||'-'||substr(pcc20,9)) X2 from PotCustCont PCC where pcc01='" & txtPCU(1) & "' order by pcc02"
   Else
      strExc(0) = "select PCC.*,decode(pcc20,null,'　','＊')||pcc02 X1,decode(pcc20,null,'',substr(pcc20,1,8)||'-'||substr(pcc20,9)) X2 from PotCustCont PCC where rownum<1"
   End If
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/06/16 +FormName 改暫存TB
   Set rsContact = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   Set rsContactOld = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   'end 2014/06/16
   Set Adodc1.Recordset = rsContact
   DataGrid1.col = 0
   DataGrid1.CurrentCellVisible = True
   ReDim m_arrConRefList(1, 0)
   If rsContact.RecordCount > 0 Then
      ReadContact
   End If
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
   
End Sub

Private Function getNewNo() As String
   Dim myTemp As ADODB.Recordset
   Dim iUsableNo As Integer
   
   Set myTemp = rsContact.Clone
   With myTemp
      .Sort = "PCC02 asc"
      iUsableNo = 1
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            If iUsableNo = Val("" & .Fields(1)) Then
               iUsableNo = iUsableNo + 1
            Else
               Exit Do
            End If
            .MoveNext
         Loop
      End If
      getNewNo = Format(iUsableNo, "00")
   End With
   Set myTemp = Nothing
End Function

Private Sub UpdateContact()
   With rsContact
   If txtPCC(2) = "" Then
      m_iConEditMode = 1
      txtPCC(2) = getNewNo
      .AddNew
   Else
      If .RecordCount > 0 Then
         .MoveFirst
         .Find "PCC02='" & txtPCC(2) & "'"
         If .EOF Then
            .AddNew
         End If
      Else
         .AddNew
      End If
   End If
   For Each oText In txtPCC
      If oText.Index = 11 Then
         .Fields("PCC" & Format(oText.Index, "00")) = DBDATE(oText.Text)
      Else
         .Fields("PCC" & Format(oText.Index, "00")) = oText.Text
      End If
   Next
   .Fields("X2") = txtPCC20
   If txtPCC20 <> "" Then
      .Fields("X1") = "＊" & .Fields("PCC02")
   Else
      .Fields("X1") = "　" & .Fields("PCC02")
   End If
   .UPDATE
   End With
End Sub

Private Sub ReadContact()
   Dim CUID(1 To 6) As String
   ClearField1
   With rsContact
   If Not .EOF Then
      For Each oText In txtPCC
         oText = "" & .Fields("PCC" & Format(oText.Index, "00"))
      Next
      CUID(1) = "" & .Fields("PCC14")
      CUID(2) = "" & .Fields("PCC15")
      CUID(3) = "" & .Fields("PCC16")
      CUID(4) = "" & .Fields("PCC17")
      CUID(5) = "" & .Fields("PCC18")
      CUID(6) = "" & .Fields("PCC19")
      txtPCC20 = "" & .Fields("X2")
      If Not IsNull(.Fields("PCC15")) Then
         m_iConEditMode = 2
      End If
   End If
   End With
   If txtPCC(6) <> "" Then
      SetList lstDept, txtPCC(6)
   End If
   If txtPCC(7) <> "" Then
      SetList lstTitle, txtPCC(7)
   End If
   If txtPCC(12) <> "" Then
      SetlstUsers 1, txtPCC(12)
   End If
   UpdateCUID CUID, textCUID1
   
   'Added by Lydia 2024/05/10 聯絡人相片
   If Trim(txtPCU(1)) <> "" And Trim(txtPCC(2)) <> "" Then
      Command1.Visible = True
      Call Pub_GetPCCtoIBF_2(Trim(txtPCU(1)), Trim(txtPCC(2)), Command1)
   Else
      Command1.Visible = False
   End If
   'end 2024/05/10

End Sub

Private Function TxtValidate1() As Boolean
   
   Dim Cancel As Boolean

   For Each oText In txtPCC
      If oText.Locked = False Then
         idx = oText.Index
         Cancel = False
         txtPCC_Validate idx, Cancel
         If Cancel = True Then
            txtPCC_GotFocus idx
            Exit Function
         End If
      End If
   Next
   
   If txtPCC(3) = "" And txtPCC(4) = "" And txtPCC(5) = "" Then
      ShowMsg "聯絡人中、英、日文名稱不可同時為空白 !"
      txtPCC(3).SetFocus
      Exit Function
   End If
   
   '2010/7/27 add by sonia 代理人狀態為國內同業者其聯絡人不可輸入E-MAIL,不寄台一雜誌,不寄電子報
   If Left(txtPCU(1), 1) = "Y" Then
      strExc(0) = "SELECT FA69 FROM FAgent WHERE FA01 = '" & txtPCU(1) & "' AND FA02 = '0'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Not IsNull(RsTemp(0)) And RsTemp(0) = "國內同業" Then
            If txtPCC(8) <> "" Then
               ShowMsg "此代理人為國內同業, 聯絡人不可輸入E-MAIL以免誤發電子郵件, 如有需要請加註於備註欄 !"
               txtPCC(8).SetFocus
               txtPCC_GotFocus 8
               Exit Function
            End If
            If txtPCC(9) <> "N" Then
               ShowMsg "此代理人為國內同業, 不可寄台一雜誌 !"
               txtPCC(9).SetFocus
               txtPCC_GotFocus 9
               Exit Function
            End If
            If txtPCC(10) <> "N" Then
               ShowMsg "此代理人為國內同業, 不可寄電子報 !"
               txtPCC(10).SetFocus
               txtPCC_GotFocus 10
               Exit Function
            End If
            'Add By Sindy 2021/11/30
            If txtPCC(24) <> "N" Then
               ShowMsg "此代理人為國內同業, 不可寄專利雙週報 !"
               txtPCC(24).SetFocus
               txtPCC_GotFocus 24
               Exit Function
            End If
            '2021/11/30 END
         End If
      End If
   'Added by Lydia 2020/11/25  檢查: 因為請作單只開放代理人
   ElseIf Left(txtPCU(1), 1) = "X" Then
        'Add by Amy 2021/11/26 +國內同業控制
        strExc(0) = "Select cu80 From Customer Where cu01 = '" & txtPCU(1) & "' And cu02= '0'"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            If "" & RsTemp(0) = "國內同業" Then
                If txtPCC(8) <> "" Then
                    ShowMsg "此客戶為國內同業, 聯絡人不可輸入E-MAIL以免誤發電子郵件, 如有需要請加註於備註欄 !"
                    txtPCC(8).SetFocus
                    txtPCC_GotFocus 8
                    Exit Function
                End If
                If txtPCC(9) <> "N" Then
                   ShowMsg "此客戶為國內同業, 不可寄台一雜誌 !"
                   txtPCC(9).SetFocus
                   txtPCC_GotFocus 9
                   Exit Function
                End If
                If txtPCC(10) <> "N" Then
                   ShowMsg "此客戶為國內同業, 不可寄電子報 !"
                   txtPCC(10).SetFocus
                   txtPCC_GotFocus 10
                   Exit Function
                End If
                'Add By Sindy 2021/11/30
                If txtPCC(24) <> "N" Then
                   ShowMsg "此客戶為國內同業, 不可寄專利雙週報 !"
                   txtPCC(24).SetFocus
                   txtPCC_GotFocus 24
                   Exit Function
                End If
                '2021/11/30 END
            End If
        'end 2021/11/26
        ElseIf txtPCC(10) = "Y" Then
           ShowMsg "X編號不可設定要寄電子報 !"
           txtPCC(10).SetFocus
           txtPCC_GotFocus 10
           Exit Function
        End If
   'end 2020/11/25
   End If
   '2010/7/27 end
   
   If txtPCC(11).Text = "" Then
      'Modify by Morgan 2008/3/13 因為X,Y多為舊資料故不強制
      'ShowMsg "開發日期不可空白 !"
      If MsgBox("開發日期為空白是否確定要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
         txtPCC(11).SetFocus
         Exit Function
      End If
   End If
      
   If lstUsers(1).ListCount = 0 Then
      'Modify by Morgan 2008/3/13 因為X,Y多為舊資料故不強制
      'ShowMsg "開發人員不可空白!"
      If MsgBox("開發人員為空白是否確定要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
         txtUserNo(1).SetFocus
         txtUserNo_GotFocus 1
         Exit Function
      End If
   Else
      'Modify by Morgan 2011/8/26 員工編號已可非數字需做轉換
       'Modified by Lydia 2022/01/10 改成Form 2.0
      'strExc(1) = ""
      'strExc(1) = PUB_Num2Id(lstUsers(1).ItemData(0))
      'For intI = 1 To lstUsers(1).ListCount - 1
      '   strExc(1) = strExc(1) & "," & PUB_Num2Id(lstUsers(1).ItemData(intI))
      'Next
      'txtPCC(12).Text = strExc(1)
      'Added by Lydia 2022/04/18 去掉多餘的,
      If Right(lstUsers(1).Tag, 1) = "," Then
          txtPCC(12).Text = Mid(lstUsers(1).Tag, 1, Len(lstUsers(1).Tag) - 1)
      Else
      'end 2022/04/18
          txtPCC(12).Text = lstUsers(1).Tag
      'end 2022/01/10
      End If 'Added by Lydia 2022/04/18
   End If
   
   '檢查聯絡人是否重複
   With rsContact
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            If .Fields("PCC02") <> txtPCC(2) And ( _
               (UCase("" & .Fields("PCC03")) = UCase(txtPCC(3)) And txtPCC(3) <> "") Or _
               (UCase("" & .Fields("PCC04")) = UCase(txtPCC(4)) And txtPCC(4) <> "") Or _
               (UCase("" & .Fields("PCC05")) = UCase(txtPCC(5)) And txtPCC(5) <> "")) Then
               If MsgBox("名稱與聯絡人[" & .Fields("PCC02") & "]重複！是否仍然要更新？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                  Exit Function
               End If
            End If
            .MoveNext
         Loop
      End If
   End With
   
    'Added by Lydia 2022/01/10 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If

   TxtValidate1 = True
End Function

Private Sub AddCombo(p_iID As Integer)
   Select Case p_iID
      Case 1
         PUB_AddDeptCombo cboDept
         
      Case 2
         PUB_AddTitleCombo cboTitle
   End Select
End Sub

Private Function ComposeListX(p_index As Integer) As String
   'Modified by Lydia 2022/01/10 改成Form 2.0
'   strExc(1) = ""
'   If lstUsers(p_index).ListCount > 0 Then
'      'Modify by Morgan 2011/8/26 員工編號已可非數字需做轉換
'      strExc(1) = PUB_Num2Id(lstUsers(p_index).ItemData(0))
'      For intI = 1 To lstUsers(p_index).ListCount - 1
'         strExc(1) = strExc(1) & "," & PUB_Num2Id(lstUsers(p_index).ItemData(intI))
'      Next
'   End If
'   ComposeListX = strExc(1)
   ComposeListX = lstUsers(p_index).Tag
   'end 2022/01/10
End Function

Private Function ComposeList(oList As ListBox) As String
   strExc(1) = ""
   If oList.ListCount > 0 Then
      strExc(1) = oList.List(0)
      For intI = 1 To oList.ListCount - 1
         strExc(1) = strExc(1) & "," & oList.List(intI)
      Next
   End If
   ComposeList = strExc(1)
End Function

Private Function ContNameChanged() As Boolean
   With rsContactOld
      If .RecordCount > 0 Then
         .MoveFirst
         .Find "pcc02='" & txtPCC(2) & "'"
         If Not .EOF Then
            If UCase("" & .Fields("pcc03")) <> UCase(txtPCC(3)) Or UCase("" & .Fields("pcc04")) <> UCase(txtPCC(4)) Or UCase("" & .Fields("pcc05")) <> UCase(txtPCC(5)) Then
               ContNameChanged = True
            End If
         End If
      Else
         ContNameChanged = True
      End If
   End With
End Function
'更新相關聯絡人資料
Private Sub UpdateConRefList(p_stContNo1, Optional p_stContNo2 As String)
   Dim ii As Integer, idx As Integer
   idx = 0
   For ii = 1 To UBound(m_arrConRefList, 2)
      If m_arrConRefList(0, ii) = p_stContNo1 Then
         idx = ii
         Exit For
      End If
   Next
   '若尚無本客戶聯絡人資料時新增
   If idx = 0 Then
      idx = UBound(m_arrConRefList, 2) + 1
      ReDim Preserve m_arrConRefList(1, idx)
      m_arrConRefList(0, idx) = p_stContNo1
   End If
   '紀錄相關聯絡人編號
   m_arrConRefList(1, idx) = p_stContNo2
End Sub
'移除相關聯絡人資料
Private Sub RemoveConRefList()
   Dim ii As Integer
   For ii = 1 To UBound(m_arrConRefList, 2)
      If m_arrConRefList(0, ii) = txtPCC(2) Then
         m_arrConRefList(1, ii) = ""
         m_arrConRefList(0, ii) = ""
         Exit For
      End If
   Next
End Sub
'更新相關聯絡人
Private Sub UpdateRefContact(p_stPCC01 As String)
   Dim ii As Integer
   For ii = 1 To UBound(m_arrConRefList, 2)
      If m_arrConRefList(0, ii) <> "" Then
         If m_arrConRefList(1, ii) = "" Then
            strSql = "update potcustcont set pcc20=Null where pcc20='" & p_stPCC01 & m_arrConRefList(0, ii) & "'"
         Else
            strSql = "update potcustcont set pcc20='" & p_stPCC01 & m_arrConRefList(0, ii) & "' where pcc01='" & Left(m_arrConRefList(1, ii), 8) & "' and pcc02='" & Mid(m_arrConRefList(1, ii), 9) & "'"
         End If
         adoTaie.Execute strSql, intI
      End If
   Next
End Sub

'Added by Lydia 2024/05/10
Private Sub Command1_Click()
   If Trim(txtPCC(2)) = "" Then
      MsgBox "請輸入聯絡人編號！"
      Exit Sub
   Else
      If Left(txtPCU(1), 1) <> "X" And Left(txtPCU(1), 1) <> "Y" Then
         MsgBox "請輸入客戶/代理人編號！"
         txtPCU(1).SetFocus
         txtPCU_GotFocus 1
         Exit Sub
      End If
      strExc(0) = "select pcc01,pcc02 from potcustcont where pcc01='" & txtPCU(1) & "' and pcc02='" & txtPCC(2) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         MsgBox "請先將聯絡人建檔完成！"
         Exit Sub
      End If
   End If
   
   frmPic001.oCP01 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "1")
   frmPic001.oCP02 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "2")
   frmPic001.oCP03 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "3")
   frmPic001.oCP04 = Pub_GetPCCtoIBF(Trim(txtPCU(1)), Trim(txtPCC(2)), "4")
   frmPic001.strWorkType = "1"
   frmPic001.Label11 = "聯絡人相片上傳"
   If m_EditMode <> 1 And m_EditMode <> 2 Then
      frmPic001.bolQuery = True '只查詢
   Else
      frmPic001.bolQuery = False '可存檔
   End If
   frmPic001.StrMenu
   frmPic001.SetSeekCmdok
   frmPic001.Show vbModal
   
   Call Pub_GetPCCtoIBF_2(Trim(txtPCU(1)), Trim(txtPCC(2)), Command1)

End Sub

