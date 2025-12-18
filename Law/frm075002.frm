VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075002 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '單線固定
   Caption         =   "法務案件基本資料維護"
   ClientHeight    =   6624
   ClientLeft      =   -2916
   ClientTop       =   1692
   ClientWidth     =   9468
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6621.927
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   9565.608
   Begin VB.CommandButton cmdIns 
      Caption         =   "各項指示(&N)"
      Height          =   285
      Left            =   6960
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   1187
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "相關卷號(&F)"
      Height          =   285
      Left            =   8160
      TabIndex        =   38
      Top             =   720
      Width           =   1187
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7710
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
            Picture         =   "frm075002.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075002.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075002.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075002.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075002.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075002.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075002.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075002.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075002.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075002.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075002.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   9468
      _ExtentX        =   16701
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
      Height          =   5595
      Left            =   60
      TabIndex        =   56
      Top             =   720
      Width           =   9315
      _ExtentX        =   16425
      _ExtentY        =   9864
      _Version        =   393216
      TabHeight       =   520
      TabMaxWidth     =   2269
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm075002.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbe(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label19"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label23"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbe(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label26"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(160)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label31"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cboContact"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text(7)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text(12)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text(6)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text(9)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text(13)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text(8)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text(28)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblMemo"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Frame1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Check1(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Check1(1)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Check1(2)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Check1(3)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Combo2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Check1(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Check1(6)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Frame2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Check1(4)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "FC資料"
      TabPicture(1)   =   "frm075002.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "銷卷資料"
      TabPicture(2)   =   "frm075002.frx":212C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label78"
      Tab(2).Control(1)=   "Label79"
      Tab(2).Control(2)=   "Label80"
      Tab(2).Control(3)=   "Label81"
      Tab(2).Control(4)=   "lblLC34"
      Tab(2).Control(5)=   "lblLC36"
      Tab(2).Control(6)=   "lblLC37"
      Tab(2).Control(7)=   "lblLC38"
      Tab(2).ControlCount=   8
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "一般"
         Height          =   215
         Index           =   4
         Left            =   3960
         TabIndex        =   28
         Top             =   4860
         Width           =   735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   700
         Left            =   3960
         TabIndex        =   132
         Top             =   4800
         Width           =   4200
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "勞資糾紛"
            Height          =   180
            Index           =   7
            Left            =   2760
            TabIndex        =   36
            Top             =   480
            Width           =   1425
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "買賣糾紛"
            Height          =   180
            Index           =   6
            Left            =   1440
            TabIndex        =   35
            Top             =   480
            Width           =   1245
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "醫療糾紛"
            Height          =   180
            Index           =   5
            Left            =   0
            TabIndex        =   34
            Top             =   480
            Width           =   1245
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "消費糾紛"
            Height          =   180
            Index           =   4
            Left            =   2760
            TabIndex        =   33
            Top             =   270
            Width           =   1425
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "車禍糾紛"
            Height          =   180
            Index           =   3
            Left            =   1440
            TabIndex        =   32
            Top             =   270
            Width           =   1245
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "土地爭議"
            Height          =   180
            Index           =   2
            Left            =   0
            TabIndex        =   31
            Top             =   270
            Width           =   1245
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "公寓大廈糾紛"
            Height          =   180
            Index           =   1
            Left            =   2760
            TabIndex        =   30
            Top             =   60
            Width           =   1425
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "承攬糾紛"
            Height          =   180
            Index           =   0
            Left            =   1440
            TabIndex        =   29
            Top             =   60
            Width           =   1245
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "公平交易法"
         Height          =   255
         Index           =   6
         Left            =   5400
         TabIndex        =   27
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "營業秘密法"
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   26
         Top             =   4560
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   276
         ItemData        =   "frm075002.frx":2148
         Left            =   3270
         List            =   "frm075002.frx":214A
         TabIndex        =   125
         Text            =   "ACS案件屬性"
         Top             =   3270
         Visible         =   0   'False
         Width           =   7284
      End
      Begin VB.CheckBox Check1 
         Caption         =   "其他智財權"
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   25
         Top             =   4290
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "著作權"
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   24
         Top             =   4290
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "商標"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   23
         Top             =   4290
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "專利"
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   22
         Top             =   4290
         Width           =   735
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  '沒有框線
         Height          =   2295
         Left            =   60
         TabIndex        =   57
         Top             =   330
         Width           =   9192
         Begin VB.TextBox txtcp01 
            Height          =   300
            Left            =   1056
            MaxLength       =   3
            TabIndex        =   0
            Top             =   30
            Width           =   495
         End
         Begin VB.TextBox txtcp02 
            Height          =   300
            Left            =   1656
            MaxLength       =   6
            TabIndex        =   1
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox txtcp03 
            Height          =   300
            Left            =   2856
            MaxLength       =   1
            TabIndex        =   2
            Top             =   30
            Width           =   375
         End
         Begin VB.TextBox txtcp04 
            Height          =   300
            Left            =   3336
            MaxLength       =   2
            TabIndex        =   3
            Top             =   30
            Width           =   495
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   29
            Left            =   5520
            TabIndex        =   129
            Top             =   30
            Width           =   375
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "656;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "特殊出名公司 :             ( J:智權公司 空白:系統預設)"
            Height          =   180
            Index           =   117
            Left            =   4230
            TabIndex        =   130
            Top             =   90
            Width           =   3915
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   52
            Left            =   5160
            TabIndex        =   9
            Top             =   993
            Width           =   375
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "661;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   27
            Left            =   1050
            TabIndex        =   8
            Top             =   975
            Width           =   1215
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   26
            Left            =   5160
            TabIndex        =   7
            Top             =   670
            Width           =   1215
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   25
            Left            =   1050
            TabIndex        =   6
            Top             =   670
            Width           =   1215
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   23
            Left            =   5160
            TabIndex        =   5
            Top             =   365
            Width           =   1215
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   0
            Left            =   1050
            TabIndex        =   4
            Top             =   365
            Width           =   1215
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   1
            Left            =   1416
            TabIndex        =   10
            Top             =   1280
            Width           =   7524
            VariousPropertyBits=   671105051
            MaxLength       =   160
            Size            =   "13271;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   2
            Left            =   1416
            TabIndex        =   11
            Top             =   1585
            Width           =   7524
            VariousPropertyBits=   671105051
            MaxLength       =   160
            Size            =   "13271;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   3
            Left            =   1416
            TabIndex        =   12
            Top             =   1890
            Width           =   7530
            VariousPropertyBits=   671105051
            MaxLength       =   160
            Size            =   "13282;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label52 
            Caption         =   "專案服務案：         (Y:是)"
            Height          =   180
            Left            =   4090
            TabIndex        =   124
            Top             =   1035
            Width           =   2000
         End
         Begin MSForms.Label lbe 
            Height          =   255
            Index           =   27
            Left            =   2310
            TabIndex        =   110
            Top             =   1000
            Width           =   1695
            VariousPropertyBits=   27
            Size            =   "2990;441"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label30 
            Caption         =   "當 事 人 5："
            Height          =   180
            Left            =   120
            TabIndex        =   109
            Top             =   1035
            Width           =   975
         End
         Begin MSForms.Label lbe 
            Height          =   255
            Index           =   26
            Left            =   6420
            TabIndex        =   108
            Top             =   693
            Width           =   1695
            VariousPropertyBits=   27
            Size            =   "2990;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label28 
            Caption         =   "當 事 人 4："
            Height          =   180
            Left            =   4230
            TabIndex        =   107
            Top             =   730
            Width           =   975
         End
         Begin MSForms.Label lbe 
            Height          =   255
            Index           =   25
            Left            =   2310
            TabIndex        =   106
            Top             =   693
            Width           =   1695
            VariousPropertyBits=   27
            Size            =   "2990;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label27 
            Caption         =   "當 事 人 3："
            Height          =   180
            Left            =   120
            TabIndex        =   105
            Top             =   730
            Width           =   975
         End
         Begin MSForms.Label lbe 
            Height          =   255
            Index           =   23
            Left            =   6420
            TabIndex        =   104
            Top             =   388
            Width           =   1695
            VariousPropertyBits=   27
            Size            =   "2990;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label25 
            Caption         =   "當 事 人 2："
            Height          =   180
            Left            =   4230
            TabIndex        =   103
            Top             =   425
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "當 事 人 1："
            Height          =   180
            Left            =   120
            TabIndex        =   64
            Top             =   425
            Width           =   975
         End
         Begin MSForms.Label lbe 
            Height          =   255
            Index           =   0
            Left            =   2310
            TabIndex        =   63
            Top             =   388
            Width           =   1695
            VariousPropertyBits=   27
            Size            =   "2990;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label lbeNumber 
            Height          =   252
            Left            =   1080
            TabIndex        =   62
            Top             =   360
            Width           =   1212
         End
         Begin VB.Label lblName 
            Caption         =   "案件名稱(中)："
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   61
            Top             =   1340
            Width           =   1215
         End
         Begin VB.Label lblName 
            Caption         =   "案件名稱(英)："
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   60
            Top             =   1645
            Width           =   1215
         End
         Begin VB.Label lblName 
            Caption         =   "案件名稱(日)："
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   59
            Top             =   1950
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "本所案號："
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   90
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  '沒有框線
         Height          =   4815
         Left            =   -74940
         TabIndex        =   65
         Top             =   360
         Width           =   9192
         Begin VB.ComboBox Combo5 
            Height          =   300
            ItemData        =   "frm075002.frx":214C
            Left            =   4200
            List            =   "frm075002.frx":215F
            Style           =   2  '單純下拉式
            TabIndex        =   51
            Top             =   3150
            Width           =   1470
         End
         Begin VB.ComboBox Combo4 
            Height          =   300
            ItemData        =   "frm075002.frx":2193
            Left            =   1110
            List            =   "frm075002.frx":2195
            Style           =   2  '單純下拉式
            TabIndex        =   50
            Top             =   3150
            Width           =   990
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   30
            Left            =   2010
            TabIndex        =   53
            Top             =   4140
            Width           =   2775
            VariousPropertyBits=   671105051
            MaxLength       =   20
            Size            =   "4895;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   24
            Left            =   1710
            TabIndex        =   49
            Top             =   2838
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "1931;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   11
            Left            =   1320
            TabIndex        =   40
            Top             =   382
            Width           =   2775
            VariousPropertyBits=   671105051
            MaxLength       =   50
            Size            =   "4895;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   10
            Left            =   1320
            TabIndex        =   39
            Top             =   75
            Width           =   1215
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   17
            Left            =   1320
            TabIndex        =   42
            Top             =   996
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "1931;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   16
            Left            =   1320
            TabIndex        =   43
            Top             =   1303
            Width           =   615
            VariousPropertyBits=   671105051
            MaxLength       =   2
            Size            =   "1085;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   15
            Left            =   4320
            TabIndex        =   44
            Top             =   1303
            Width           =   495
            VariousPropertyBits=   671105051
            MaxLength       =   1
            Size            =   "873;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   14
            Left            =   1320
            TabIndex        =   41
            Top             =   689
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "1931;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   630
            Index           =   22
            Left            =   1320
            TabIndex        =   52
            Top             =   3480
            Width           =   7395
            VariousPropertyBits=   -1466941413
            MaxLength       =   2000
            ScrollBars      =   2
            Size            =   "13044;1111"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   19
            Left            =   1320
            TabIndex        =   46
            Top             =   1917
            Width           =   5655
            VariousPropertyBits=   671105051
            Size            =   "9975;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   18
            Left            =   1320
            TabIndex        =   45
            Top             =   1610
            Width           =   5175
            VariousPropertyBits=   671105051
            MaxLength       =   35
            Size            =   "9128;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   20
            Left            =   1320
            TabIndex        =   47
            Top             =   2224
            Width           =   5655
            VariousPropertyBits=   671105051
            Size            =   "9975;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text 
            Height          =   300
            Index           =   21
            Left            =   1320
            TabIndex        =   48
            Top             =   2531
            Width           =   5655
            VariousPropertyBits=   671105051
            Size            =   "9975;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "CLIENT_MATTER_ID："
            Height          =   180
            Left            =   150
            TabIndex        =   123
            Top             =   4170
            Width           =   1845
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "請款單列印幣別格式："
            Height          =   180
            Left            =   2370
            TabIndex        =   122
            Top             =   3210
            Width           =   1800
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "請款幣別："
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   121
            Top             =   3210
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "與他案合併計算結餘請於案件備註欄註明""與某案號合併計算結餘""！"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   172
            Left            =   120
            TabIndex        =   120
            Top             =   4530
            Width           =   5370
         End
         Begin VB.Label Label20 
            Caption         =   "D/N固定列印對象："
            Height          =   180
            Left            =   120
            TabIndex        =   89
            Top             =   2898
            Width           =   1545
         End
         Begin MSForms.Label lbe 
            Height          =   255
            Index           =   24
            Left            =   2880
            TabIndex        =   88
            Top             =   2861
            Width           =   6180
            VariousPropertyBits=   27
            Size            =   "10901;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label4 
            Caption         =   "彼所案號："
            Height          =   180
            Left            =   120
            TabIndex        =   81
            Top             =   442
            Width           =   975
         End
         Begin MSForms.Label lbe 
            Height          =   255
            Index           =   10
            Left            =   2640
            TabIndex        =   80
            Top             =   100
            Width           =   6360
            VariousPropertyBits=   27
            Size            =   "11218;441"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label18 
            Caption         =   "FC代理人："
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   79
            Top             =   135
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "副本收受人：  "
            Height          =   180
            Left            =   120
            TabIndex        =   78
            Top             =   1056
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "折            扣："
            Height          =   180
            Left            =   120
            TabIndex        =   77
            Top             =   1363
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "(Y:是)"
            Height          =   255
            Index           =   2
            Left            =   4950
            TabIndex        =   76
            Top             =   1326
            Width           =   510
         End
         Begin MSForms.Label lbe 
            Height          =   255
            Index           =   14
            Left            =   2520
            TabIndex        =   75
            Top             =   712
            Width           =   6390
            VariousPropertyBits=   27
            Size            =   "11271;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label7 
            Caption         =   "D/N是否列印申請人："
            Height          =   180
            Left            =   2520
            TabIndex        =   74
            Top             =   1363
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "固定請款對象："
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   73
            Top             =   749
            Width           =   1335
         End
         Begin VB.Label Label22 
            Caption         =   "案件備註："
            Height          =   180
            Left            =   120
            TabIndex        =   72
            Top             =   3465
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "聯絡人(中)："
            Height          =   180
            Left            =   120
            TabIndex        =   71
            Top             =   1977
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "副本聯絡人："
            Height          =   180
            Left            =   120
            TabIndex        =   70
            Top             =   1670
            Width           =   1095
         End
         Begin MSForms.Label lbe 
            Height          =   255
            Index           =   17
            Left            =   2520
            TabIndex        =   69
            Top             =   1019
            Width           =   6360
            VariousPropertyBits=   27
            Size            =   "11218;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label15 
            Caption         =   "%"
            Height          =   255
            Left            =   2040
            TabIndex        =   68
            Top             =   1365
            Width           =   255
         End
         Begin VB.Label Label13 
            Caption         =   "聯絡人(日)："
            Height          =   180
            Left            =   120
            TabIndex        =   67
            Top             =   2591
            Width           =   1095
         End
         Begin VB.Label Label14 
            Caption         =   "聯絡人(英)："
            Height          =   180
            Left            =   120
            TabIndex        =   66
            Top             =   2284
            Width           =   1095
         End
      End
      Begin VB.Label lblMemo 
         Caption         =   "可以直接輸入，屬性之間用逗號,區隔。"
         Height          =   765
         Left            =   150
         TabIndex        =   131
         Top             =   4560
         Width           =   945
      End
      Begin MSForms.TextBox Text 
         Height          =   1170
         Index           =   28
         Left            =   1110
         TabIndex        =   21
         Top             =   4260
         Width           =   2775
         VariousPropertyBits=   -1467989989
         MaxLength       =   200
         Size            =   "4895;2064"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   8
         Left            =   1110
         TabIndex        =   17
         Top             =   3255
         Width           =   735
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   13
         Left            =   5580
         TabIndex        =   20
         Top             =   3930
         Width           =   1935
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "3413;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   9
         Left            =   1485
         TabIndex        =   18
         Top             =   3570
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4895;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   6
         Left            =   1110
         TabIndex        =   15
         Top             =   2925
         Width           =   1215
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   12
         Left            =   1110
         TabIndex        =   19
         Top             =   3930
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   4
         Left            =   1110
         TabIndex        =   13
         Top             =   2610
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   7
         Left            =   5580
         TabIndex        =   16
         Top             =   2925
         Width           =   855
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text 
         Height          =   300
         Index           =   5
         Left            =   5580
         TabIndex        =   14
         Top             =   2610
         Width           =   3405
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "6006;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboContact 
         Height          =   300
         Left            =   5580
         TabIndex        =   126
         Top             =   3570
         Width           =   1770
         VariousPropertyBits=   679495711
         DisplayStyle    =   3
         Size            =   "3122;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0FFC0&
         Caption         =   "案件屬性："
         Height          =   210
         Left            =   180
         TabIndex        =   119
         Top             =   4320
         Width           =   945
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷備註："
         Height          =   255
         Left            =   -74880
         TabIndex        =   118
         Top             =   810
         Width           =   1260
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   250
         Left            =   -68550
         TabIndex        =   117
         Top             =   510
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   180
         Left            =   -71670
         TabIndex        =   116
         Top             =   510
         Width           =   1080
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   250
         Left            =   -74880
         TabIndex        =   115
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label lblLC34 
         Height          =   250
         Left            =   -73800
         TabIndex        =   114
         Top             =   480
         Width           =   2115
      End
      Begin VB.Label lblLC36 
         Height          =   250
         Left            =   -70590
         TabIndex        =   113
         Top             =   510
         Width           =   1995
      End
      Begin VB.Label lblLC37 
         Height          =   250
         Left            =   -67500
         TabIndex        =   112
         Top             =   480
         Width           =   1665
      End
      Begin VB.Label lblLC38 
         Height          =   250
         Left            =   -73620
         TabIndex        =   111
         Top             =   812
         Width           =   7755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   180
         Index           =   160
         Left            =   4620
         TabIndex        =   102
         Top             =   3630
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "智財權案："
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   101
         Top             =   3990
         Width           =   945
      End
      Begin VB.Label Label8 
         Caption         =   "(Y:是)"
         Height          =   255
         Index           =   0
         Left            =   1530
         TabIndex        =   100
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label26 
         Caption         =   "閉卷日期："
         Height          =   180
         Left            =   180
         TabIndex        =   99
         Top             =   2985
         Width           =   1215
      End
      Begin MSForms.Label lbe 
         Height          =   255
         Index           =   7
         Left            =   6495
         TabIndex        =   98
         Top             =   2948
         Width           =   2130
         VariousPropertyBits=   27
         Size            =   "3757;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label23 
         Caption         =   "閉卷原因："
         Height          =   180
         Left            =   4620
         TabIndex        =   97
         Top             =   2985
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "客戶案件案號："
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   96
         Top             =   3630
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "分所案號："
         Height          =   180
         Left            =   4620
         TabIndex        =   95
         Top             =   2670
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "相關國家："
         Height          =   180
         Left            =   180
         TabIndex        =   94
         Top             =   3315
         Width           =   1215
      End
      Begin MSForms.Label lbe 
         Height          =   255
         Index           =   8
         Left            =   1890
         TabIndex        =   93
         Top             =   3285
         Width           =   1860
         VariousPropertyBits=   27
         Size            =   "3281;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         Caption         =   "是否閉卷："
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   92
         Top             =   2670
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "(Y:閉卷)"
         Height          =   255
         Index           =   1
         Left            =   1530
         TabIndex        =   91
         Top             =   2633
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "署名人："
         Height          =   180
         Left            =   4620
         TabIndex        =   90
         Top             =   3990
         Width           =   780
      End
   End
   Begin MSForms.Label UIDname 
      Height          =   285
      Left            =   5220
      TabIndex        =   128
      Top             =   6360
      Width           =   735
      VariousPropertyBits=   27
      Caption         =   "UIDname"
      Size            =   "1296;494"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label IDname 
      Height          =   285
      Left            =   1140
      TabIndex        =   127
      Top             =   6360
      Width           =   735
      VariousPropertyBits=   27
      Caption         =   "IDname"
      Size            =   "1296;494"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label UTM 
      Caption         =   "Label28"
      Height          =   285
      Left            =   7110
      TabIndex        =   87
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label UDT 
      Caption         =   "Label27"
      Height          =   285
      Left            =   6150
      TabIndex        =   86
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label29 
      Caption         =   "UpdateID ："
      Height          =   285
      Left            =   4110
      TabIndex        =   85
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label CTM 
      Caption         =   "Label28"
      Height          =   285
      Left            =   3030
      TabIndex        =   84
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label CDT 
      Caption         =   "Label27"
      Height          =   285
      Left            =   2070
      TabIndex        =   83
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label24 
      Caption         =   "CreateID ："
      Height          =   285
      Left            =   60
      TabIndex        =   82
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "備註："
      Height          =   255
      Left            =   -74520
      TabIndex        =   54
      Top             =   2640
      Width           =   615
   End
End
Attribute VB_Name = "frm075002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/15 改成Form2.0 ; Text(index)、lbe(index)、cboContact、IDname、UIDname
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim Rs As New ADODB.Recordset
Dim LcTmp As String
Dim m_Cpnum As String
'Add By Sindy 2011/5/31
' 第一筆資料的本所案號
Dim m_FirstRow(4) As String
' 最後一筆資料的本所案號
Dim m_LastRow(4) As String
' 目前正在顯示的本所案號
Dim m_CurrRow(4) As String
'2011/5/31 End
' 90.07.16 modify by Ken (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Public m_EditMode As Integer
Dim strChkCuAreaMail As String, strChkCuAreaMailTo As String 'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
'Added by Lydia 201/01/14 法律所案源收文
Dim m_LOS01 As String '案源總收文號
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '案源總收文號之本所案號
Dim m_LOS02 As String '案源案件類型
Dim m_CRL84 As String '接洽記錄單-法務案件屬性
Dim oObj As Control 'Added by Lydia 2022/08/10

'Add By Sindy 2011/5/31
Private Sub Check1_Click(Index As Integer)
   If Check1(Index).Value = 1 Then
      If InStr(Text(28).Text, Trim(Check1(Index).Caption)) = 0 Then
         If Text(28).Text = "" Then
            Text(28).Text = Trim(Check1(Index).Caption)
         Else
            Text(28).Text = Text(28).Text & "," & Trim(Check1(Index).Caption)
         End If
      End If
      'Added by Lydia 2023/03/14 一般案件屬性
      If Index = 4 Then
          Frame2.Enabled = True
      End If
      'end 2023/03/14
   Else
      '案件屬性=xx,xx,xx
      If Left(Text(28), Len(Trim(Check1(Index).Caption))) = Trim(Check1(Index).Caption) Then
         Text(28).Text = Replace(Text(28).Text, Trim(Check1(Index).Caption) & ",", "")
         Text(28).Text = Replace(Text(28).Text, Trim(Check1(Index).Caption), "")
      Else
         Text(28).Text = Replace(Text(28).Text, "," & Trim(Check1(Index).Caption), "")
      End If
      'Added by Lydia 2023/03/14 一般案件屬性
      If Index = 4 Then
         Frame2.Enabled = False
         For Each oObj In Check2
            If oObj.Value = 1 Then
                oObj.Value = 0
                Call Check2_Click(oObj.Index)
            End If
         Next
      End If
      'end 2023/03/14
   End If
   If InStr(Text(28).Text, "專利") > 0 Or InStr(Text(28).Text, "商標") > 0 Or _
      InStr(Text(28).Text, "著作權") > 0 Or InStr(Text(28).Text, "智財權") > 0 Then
      Text(12) = "Y"
   Else
      Text(12) = ""
   End If
End Sub

'Added by Lydia 2016/11/24 各項指示
Private Sub cmdIns_Click()
   If txtcp01.Text = "" Or txtcp02.Text = "" Then
      MsgBox "請輸入本所案號", vbInformation
      Exit Sub
   End If
   'Added by Lydia 2020/05/05
   If m_EditMode <> 0 And m_EditMode <> 4 Then
      MsgBox IIf(m_EditMode = 1, "新增中", "修改中") & "不可執行！", vbInformation
      Exit Sub
   End If
   'end 2020/05/05
   'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
   If PUB_CheckFormExist("frm12040159") Then
       MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
       Exit Sub
   End If
   'end 2020/05/05
   
   frm12040159.SetParent "E", Trim(txtcp01.Text & txtcp02.Text & txtcp03.Text & txtcp04.Text), Me
   frm12040159.Show
End Sub

Private Sub cmdOther_Click()
Dim i As Integer
Dim strNum As String
Dim strTmp As String
   
   If txtcp01.Text = "" And txtcp02.Text = "" Then
      MsgBox "請輸入本所案號", vbInformation, "顧問案件資料維護"
      Exit Sub
   End If
   Set frm1103_2.m_form = Me
   frm1103_2.intWhereComeFrom = 1
   frm1103_2.lblSystem = txtcp01.Text
   frm1103_2.lblCode(0) = txtcp02.Text
   frm1103_2.lblCode(1) = txtcp03.Text
   frm1103_2.lblCode(2) = txtcp04.Text
   frm1103_2.Show
   Me.Hide
End Sub

'Add By Sindy 2016/11/23
Private Sub Combo4_Click()
   Call GetCurrType
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo4_Validate(Cancel As Boolean)
   If Combo4 = MsgText(601) Then
      Combo4.Tag = Combo4.Text
      Combo5.ListIndex = 0
      Combo5.Enabled = False
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo4, Label11(0)) = False Then
      Cancel = True
      Combo4.SetFocus
   End If
   If Combo4 <> "USD" Then
      If ExistCheck("DebitNoteRate", "DNR01", Combo4, Label11(0) & "匯率") = False Then
         Cancel = True
         Combo4.SetFocus
         Exit Sub
      End If
   End If
   Call GetCurrType
End Sub
Private Sub GetCurrType()
Dim intType As Integer
   
   If Combo4 = MsgText(601) Then
      Combo4.Tag = Combo4.Text
      Combo5.ListIndex = 0
      Combo5.Enabled = False
      Exit Sub
   End If
   '若更改請款幣別
   If Me.Combo4.Text <> Me.Combo4.Tag Then
      Me.Combo4.Tag = Me.Combo4.Text
      '請款幣別變更要重新預設列印幣別
      '台幣
      If Me.Combo4.Text = "NTD" Then
         intType = 1 '純台幣
      '人民幣
      ElseIf Me.Combo4.Text = "RMB" Then
         intType = 4 '外幣+美金合計
      '其他幣別
      Else
         intType = 2 '台幣+外幣合計
      End If
      Combo5.ListIndex = intType
      '若為台幣時則格式欄位鎖住不可修改
      If Me.Combo4.Text = "NTD" Then
         Combo5.Enabled = False
      Else
         Combo5.Enabled = True
      End If
   End If
End Sub
'2016/11/23 END


'add by nickc 2006/11/10 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case 13:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
   End Select
End Sub

Private Sub Form_Load()
   
   SSTab1.Tab = 0
   
   ' 90.07.16 modify by Ken (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm075002", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm075002", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm075002", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm075002", strFind, False)
   ' Ken 90.07.16 -- End
   
   m_EditMode = 0
   MoveFormToCenter Me
   
   txtcp01.Text = strSysKind  '2011/5/20 add by sonia
   
   If Not IsEmptyText(m_CurrRow(0)) And Not IsEmptyText(m_CurrRow(1)) And Not IsEmptyText(m_CurrRow(2)) And Not IsEmptyText(m_CurrRow(3)) Then
      ShowCurrRecord m_CurrRow(0), m_CurrRow(1), m_CurrRow(2), m_CurrRow(3)
      UpdateToolbarState
      SetCtrlReadOnly True
   Else
      m_EditMode = 4
      SetCtrlReadOnly True
      SetKeyReadOnly False
      UpdateToolbarState
   End If
   
   'Add By Sindy 2016/11/23
   '抓有輸入過匯率的請款幣別
   Combo4.Clear
   Combo4.AddItem ""
   Combo4.AddItem "USD"
   If RsTemp.State <> adStateClosed Then RsTemp.Close
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open "select distinct DNR01 from DebitNoteRate order by DNR01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While RsTemp.EOF = False
      Combo4.AddItem RsTemp.Fields("DNR01").Value
      RsTemp.MoveNext
   Loop
   RsTemp.Close
   '2016/11/23 End
   
   'Added by Lydia 2020/03/30 事務所合併日起取消( J:智權公司 空白:系統預設)的標題
   'Modified by Lydia 2020/05/29 非ACS案才取消。
   If strSrvDate(1) >= 事務所合併日 And strSysKind <> "ACS" Then
       Label1(117).Visible = False
       Text(29).Visible = False
   End If
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdIns.Visible = True
   Else
      cmdIns.Visible = False
   End If
   'end 2020/05/05
   'Added by Lydia 2023/03/14 一般案件屬性
   If strSysKind = "ACS" Then
      Frame2.Visible = False
      lblMemo.Visible = False
   End If
   
End Sub

' 按下按鍵
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
      ' 確定, 取消
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
      ' 取消或離開
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

Private Sub Text_Change(Index As Integer)
   Select Case Index
      'Modify By Sindy 2011/1/14 +23,25,26,27
      Case 0, 7, 8, 10, 14, 17, 24, 23, 25, 26, 27
         If Text(Index) = "" Then lbe(Index) = ""
   End Select
End Sub

Private Sub Text_GotFocus(Index As Integer)
   TextInverse Text(Index)
   Select Case Index
      Case 1, 3, 13, 19, 21, 22
         'edit by nickc 2007/06/11  切換輸入法改用API
         'Text(Index).IMEMode = 1
         OpenIme
      Case Else
         'edit by nickc 2007/06/11  切換輸入法改用API
         'Text(Index).IMEMode = 2
         CloseIme
   End Select

End Sub

'Modified by Lydia 2021/09/15 改成Form 2.0
'Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
   'Add by Amy 2018/08/15 只能輸空白及Y
   Select Case Index
      Case 4, 12, 52
        If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
           KeyAscii = 0
           Beep
        End If
   End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
Dim blnIsEmpty As Boolean
   
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text(Index).IMEMode = 2
   CloseIme
   If Index = 3 Then
      For Index = 1 To 3
         If Text(Index) <> "" Then
            blnIsEmpty = False
            Exit For
         Else
            blnIsEmpty = True
         End If
      Next
      If blnIsEmpty Then
         MsgBox "案件名稱不可同時為空", vbCritical
         Text(1).SetFocus
         Exit Sub
      End If
   End If
End Sub

'Added by Lydia 2021/09/15
Private Sub Text_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Index = 28 Then
      If Button = 2 Then Forms(0).PopupMenu2 Text(Index)
   End If
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim strTempName As String, blnIsEmpty As Boolean, i As Integer, strTemp As String
Dim arrTemp As Variant 'Added by Lydia 2023/03/14

   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text(Index).IMEMode = 2
   CloseIme
   'Modify By Sindy 2010/12/16 Mark
   'If m_click <> True Then
   '   Exit Sub
   'End If
   Select Case Index
      Case 1:
         If CheckLengthIsOK(Text(1), 160) = False Then
            Cancel = True
            Text(1).SetFocus
            TextInverse Text(1)
            Exit Sub
         End If
      Case 3:
         If CheckLengthIsOK(Text(3), 160) = False Then
            Cancel = True
            Text(3).SetFocus
            TextInverse Text(3)
            Exit Sub
         End If
      Case 5:
         If CheckLengthIsOK(Text(5), 50) = False Then
            Cancel = True
            Text(5).SetFocus
            TextInverse Text(5)
            Exit Sub
         End If
      Case 13:
         If CheckLengthIsOK(Text(13), 10) = False Then
            Cancel = True
            Text(13).SetFocus
           ' TextInverse Text(13)
            'Exit Sub
         End If
      Case 19:
         'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
         'If CheckLengthIsOK(Text(19), 10) = False Then
         If CheckLengthIsOK(Text(19), 30) = False Then
            Cancel = True
            Text(19).SetFocus
            TextInverse Text(19)
            Exit Sub
         End If
      'Added by Lydia 2017/06/14 聯絡人(英)
      Case 20:
         If CheckLengthIsOK(Text(20), 35) = False Then
            Cancel = True
            Text(20).SetFocus
            TextInverse Text(20)
            Exit Sub
         End If
      'end 2017/06/14
      Case 21:
         'Modifed by Lydia 2017/06/14  聯絡人(日)為60字
         'If CheckLengthIsOK(Text(21), 20) = False Then
         If CheckLengthIsOK(Text(21), 60) = False Then
            Cancel = True
            Text(21).SetFocus
            TextInverse Text(21)
            Exit Sub
         End If
      Case 22:
         If CheckLengthIsOK(Text(22), 2000) = False Then
            Cancel = True
            Text(22).SetFocus
            TextInverse Text(22)
            Exit Sub
         End If
      'Add By Sindy 2012/4/13
      Case 8:
         If (txtcp01 = "FCL" Or txtcp01 = "LIN") And Text(Index).Text <> "" And Text(Index).Text <> 台灣國家代號 Then
            ShowMsg MsgText(9219)
            Cancel = True
            Text(Index).SetFocus
            TextInverse Text(Index)
            Exit Sub
         End If
   End Select
   
   Select Case Index
   'Modify By Sindy 2011/1/14 +23,25,26,27
   Case 0, 23, 25, 26, 27 '當事人
      If Text(Index) <> "" Then
         strTemp = UCase(Text(Index))
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCustomer(strTemp, strTempName) Then
         If ClsPDGetCustomer(strTemp, strTempName) Then
            Text(Index) = strTemp
            lbe(Index) = strTempName
         Else
            Cancel = True
            lbe(Index) = ""
         End If
         'Add By Sindy 2011/1/14 檢查輸入當事人的順序
         If (Text(23) <> "" And Text(0) = "") Or _
            (Text(25) <> "" And Text(23) = "") Or _
            (Text(26) <> "" And Text(25) = "") Or _
            (Text(27) <> "" And Text(26) = "") Then
            MsgBox "請依序輸入當事人!", vbExclamation, "法務案件基本資料維護"
            If Text(23) <> "" And Text(0) = "" Then Text(23).SetFocus: Call Text_GotFocus(23)
            If Text(25) <> "" And Text(23) = "" Then Text(25).SetFocus: Call Text_GotFocus(25)
            If Text(26) <> "" And Text(25) = "" Then Text(26).SetFocus: Call Text_GotFocus(26)
            If Text(27) <> "" And Text(26) = "" Then Text(27).SetFocus: Call Text_GotFocus(27)
            Cancel = True
            Exit Sub
         End If
         'Add By Sindy 2011/1/14 檢查當事人不可重複
         If Index = 0 Then
            If Text(Index) = Text(23) Or _
               Text(Index) = Text(25) Or _
               Text(Index) = Text(26) Or _
               Text(Index) = Text(27) Then
               MsgBox "當事人不可重複!", vbExclamation, "法務案件基本資料維護"
               Text(Index).SetFocus
               Text_GotFocus (Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 23 Then
            If Text(Index) = Text(0) Or _
               Text(Index) = Text(25) Or _
               Text(Index) = Text(26) Or _
               Text(Index) = Text(27) Then
               MsgBox "當事人不可重複!", vbExclamation, "法務案件基本資料維護"
               Text(Index).SetFocus
               Text_GotFocus (Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 25 Then
            If Text(Index) = Text(0) Or _
               Text(Index) = Text(23) Or _
               Text(Index) = Text(26) Or _
               Text(Index) = Text(27) Then
               MsgBox "當事人不可重複!", vbExclamation, "法務案件基本資料維護"
               Text(Index).SetFocus
               Text_GotFocus (Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 26 Then
            If Text(Index) = Text(0) Or _
               Text(Index) = Text(23) Or _
               Text(Index) = Text(25) Or _
               Text(Index) = Text(27) Then
               MsgBox "當事人不可重複!", vbExclamation, "法務案件基本資料維護"
               Text(Index).SetFocus
               Text_GotFocus (Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 27 Then
            If Text(Index) = Text(0) Or _
               Text(Index) = Text(23) Or _
               Text(Index) = Text(25) Or _
               Text(Index) = Text(26) Then
               MsgBox "當事人不可重複!", vbExclamation, "法務案件基本資料維護"
               Text(Index).SetFocus
               Text_GotFocus (Index)
               Cancel = True
               Exit Sub
            End If
         End If
      End If
   Case 1, 2, 3, 5, 13, 18, 19, 20, 21, 22
      If Text(Index) <> "" Then Text(Index) = UCase(Text(Index))
    Case 4, 12, 15 'edit by nickc 2006/07/12 , 23
        If Text(Index) <> "" Then
           Text(Index) = UCase(Text(Index))
           If Text(Index) <> "Y" Then DataErrorMessage 1, "": Cancel = True
        End If
   Case 6
      If Text(Index) <> "" Then
         If Text(Index) = 0 Then Text(Index) = ""
         If Not CheckIsTaiwanDate(Text(Index)) Then
            Cancel = True
         End If
      End If
   Case 7
         If Text(Index) <> "" Then
         ' If objLawDll.GetReasonOfRelief(Text(Index), strTempName) Then lbe(Index) = strTempName Else DataErrorMessage 1, "閉卷原因": Cancel = True
              If GetReasonOfRelief(Text(Index), strTempName) Then
                 lbe(Index) = strTempName
              Else
                 DataErrorMessage 1, "閉卷原因"
                 Cancel = True
                 lbe(Index) = ""
              End If
          End If
      
   Case 8
       If Text(Index) <> "" Then
          'edit by nickc 2007/02/07 不用 dll 了
          'If objPublicData.GetNation(Text(Index), strTempName) Then Lbe(Index) = strTempName Else Cancel = True: Lbe(Index) = ""
          If ClsPDGetNation(Text(Index), strTempName) Then lbe(Index) = strTempName Else Cancel = True: lbe(Index) = ""
       Else
          lbe(Index) = ""
       End If
   Case 10
       If Text(Index) <> "" Then
   '2010/2/2 CANCEL BY SONIA
   '           If Len(Text(Index)) < 9 Then
            Text(Index) = UCase(Text(Index))
            strTemp = Text(Index)
             'edit by nickc 2007/02/07 不用 dll 了
             'If objPublicData.GetAgent(strTemp, strTempName) Then
             If ClsPDGetAgent(strTemp, strTempName) Then
                Text(Index) = strTemp
                lbe(Index) = strTempName
             Else
                Cancel = True
             End If
   '           Else
   '             DataErrorMessage 1, ""
   '             Cancel = True
   '           End If
   '2010/2/2 END
       Else
         lbe(Index) = ""
       End If
   Case 14, 17, 24
       If Text(Index) <> "" Then
          Text(Index) = UCase(Text(Index))
          If Left(Text(Index).Text, 1) <> "X" And Left(Text(Index).Text, 1) <> "Y" Then
             MsgBox "代碼輸入錯誤!", vbExclamation, "法務案件基本資料維護"
             Cancel = True
             Exit Sub
          End If
   '2010/2/2 CANCEL BY SONIA
   '           If Len(Text(Index)) < 9 Then
            Text(Index) = UCase(Text(Index))
            strTemp = Text(Index)
            If Left(strTemp, 1) = "X" Then
                 'edit by nickc 2007/02/07 不用 dll 了
                 'If objPublicData.GetCustomer(strTemp, strTempName) Then
                 If ClsPDGetCustomer(strTemp, strTempName) Then
                    Text(Index) = strTemp
                    lbe(Index) = strTempName
                 Else
                    Cancel = True
                 End If
            
            ElseIf Left(strTemp, 1) = "Y" Then
                 'edit by nickc 2007/02/07 不用 dll 了
                 'If objPublicData.GetAgent(strTemp, strTempName) Then
                 If ClsPDGetAgent(strTemp, strTempName) Then
                    Text(Index) = strTemp
                    lbe(Index) = strTempName
                 Else
                    Cancel = True
                 End If
            End If
   '           Else
   '             DataErrorMessage 1, ""
   '             Cancel = True
   '           End If
   '2010/2/2 END
       Else
         lbe(Index) = ""
       End If
   Case 16
        If Text(16).Text <> "" Then
            If CInt(Text(16)) > 99 Then
               Cancel = True
               MsgBox "折扣不可大於99!", vbExclamation, "法務案件基本資料維護"
               Text(16).SetFocus
               Exit Sub
            End If
        End If
      'Added by Lydia 2023/03/14 案件屬性：可直接輸入
      Case 28
          If Text(Index).Locked = False Then
             If Trim(Text(Index)) = "" Then
                For Each oObj In Check1
                   oObj.Value = 0
                Next
                For Each oObj In Check2
                   oObj.Value = 0
                Next
             Else
               strExc(1) = PUB_StringFilter(Text(Index))
               arrTemp = Split(strExc(1), ",")
               For i = 0 To UBound(arrTemp)
                  If Trim(arrTemp(i)) <> "" Then
                      '案件屬性
                      For Each oObj In Check1
                         If Trim(arrTemp(i)) = oObj.Caption And oObj.Value = 0 Then
                            oObj.Value = 1
                         ElseIf InStr(strExc(1), oObj.Caption) = 0 Then
                            oObj.Value = 0
                         End If
                      Next
                      '一般案件屬性
                      For Each oObj In Check2
                        If Trim(arrTemp(i)) = oObj.Caption And oObj.Value = 0 Then
                           oObj.Value = 1
                           If Check1(4).Value = 0 Then
                               Check1(4).Value = 1
                               strExc(1) = "一般," & strExc(1)
                           End If
                        ElseIf InStr(strExc(1), oObj.Caption) = 0 Then
                           oObj.Value = 0
                        End If
                      Next
                  End If
               Next i
               Text(Index) = strExc(1)
             End If
          End If
      'end 2023/03/14
   'Add By Sindy 2013/12/13
   Case 29
         Text(29) = Trim(Text(29))
         If Text(29) <> "" And Text(29) <> "J" Then
            Cancel = True
            MsgBox "請輸入J或空白!", vbExclamation, "法務案件基本資料維護"
            Call Text_GotFocus(29)
            Text(29).SetFocus
            Exit Sub
         End If
   End Select
   
'   If Not (blnIsNew Or blnisEdit) Then Cancel = False
   If Cancel Then TextInverse Text(Index)
End Sub

' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub txtcp01_GotFocus()
   InverseTextBox txtcp01
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   txtcp01.Text = UCase(txtcp01)
   If IsEmptyText(txtcp01) = False Then
      '2011/5/20 MODIFY BY SONIA
      'If Not IsCorrectSysKindLaw(txtcp01) Then
      If CheckSys(txtcp01) <> "3" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "系統類別不正確"
         MsgBox strMsg, vbOKOnly, strTit
         txtcp01_GotFocus
         Exit Sub
      End If
      
      ' 檢查使用者是否有使用該系統類別的權限
      If IsUserHasRightOfSystem(strUserNum, txtcp01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使有此系統別的權限"
         MsgBox strMsg, vbOKOnly, strTit
         txtcp01_GotFocus
         Exit Sub
      End If
   End If
   
   'add by sonia 2019/8/14
   If txtcp01 = "ACS" Then
      'Remove by Lydia 2020/11/06 取消ACS的案件屬性欄;
      'Combo2.Enabled = True
      'Combo2.Visible = True
      'Combo2.Left = 1116
      'Combo2.Top = 3852
      Label31.Visible = False
      'end 2020/11/06
      Text(28) = ""
      Text(28).Enabled = False
      Text(28).Visible = False
      'Modified by Lydia 2022/08/10
      For Each oObj In Check1
          oObj.Visible = False
          oObj.Enabled = False
      Next
      'end 2022/08/10
      'Added by Lydia 2023/03/14 一般案件屬性
      Frame2.Visible = False
      lblMemo.Visible = False
   Else
      Combo2.Enabled = False
      Combo2.Visible = False
      Label31.Visible = True 'Added by Lydia 2020/11/06
      Text(28).Enabled = True
      Text(28).Visible = True
      Combo2 = ""
      'Modified by Lydia 2022/08/10
      For Each oObj In Check1
          oObj.Visible = True
          oObj.Enabled = True
      Next
      'end 2022/08/10
      'Added by Lydia 2023/03/14 一般案件屬性
      Frame2.Visible = True
      lblMemo.Visible = True
   End If
   'end 2019/8/14
   
   'Added by Lydia 2020/06/05 ACS案才顯示 ( J:智權公司 空白:系統預設)的標題
   If strSrvDate(1) >= 事務所合併日 And txtcp01 = "ACS" Then
       Label1(117).Visible = True
       Text(29).Visible = True
   Else
       Label1(117).Visible = False
       Text(29).Visible = False
   End If
End Sub

'add by sonia 2019/8/1
Private Sub Combo2_Click()
   Combo2.Tag = Combo2
End Sub

Private Sub Combo2_DropDown()
   PUB_SetCasePTM "4", , Me.Combo2
End Sub
'end 2019/8/1

Private Sub txtcp02_GotFocus()
   InverseTextBox txtcp02
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
Dim strTemp As String
   
   If txtcp02 <> "" Then
      strTemp = GiveSymbol(txtcp01, txtcp02, txtcp03, txtcp04, LcTmp)
      m_Cpnum = strTemp
      'edit by nickc 2007/02/07 不用 dll 了
      'If objLawDll.ChkCaseNum(txtcp01, txtcp02) Then
      If ClsPDChkCaseNum(txtcp01, txtcp02) Then
         TextInverse txtcp02
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtcp02
End Sub

Private Sub txtcp03_GotFocus()
   TextInverse txtcp03
End Sub

Private Sub txtcp03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub ChgType(i As Integer)
Dim strTempName As String, strTemp1 As String, strTemp2 As String
Dim j As Integer, n As Integer
   
   Select Case i
    'Modify By Sindy 2011/1/14 +當事人2,3,4,5
    Case 0, 23, 25, 26, 27
        If Text(i) <> "" Then
          strTemp1 = Text(i)
          'edit by nickc 2007/02/07 不用 dll 了
          'If objPublicData.GetCustomer(strTemp1, strTempName) Then
          If ClsPDGetCustomer(strTemp1, strTempName) Then
              lbe(i) = strTempName
              Text(i) = strTemp1
          End If
        End If
    Case 6
       If Text(i) = "0" Then Text(i) = ""
       If Text(i) <> "" Then Text(i) = ChangeWStringToTString(Text(i))
    Case 7
       If Text(i) <> "" Then
           'edit by nickc 2007/02/07 不用 dll 了
           'If objLawDll.GetReasonOfRelief(Text(i), strTempName) Then Lbe(i) = strTempName
           If ClsLawGetReasonOfRelief(Text(i), strTempName) Then lbe(i) = strTempName
       End If
    Case 8
        If Text(i) <> "" Then
           'edit by nickc 2007/02/07 不用 dll 了
           'If objPublicData.GetNation(Text(i), strTempName) Then Lbe(i) = strTempName
           If ClsPDGetNation(Text(i), strTempName) Then lbe(i) = strTempName
        End If
        If Text(i) = "" Then lbe(i) = ""
    Case 9
         If Text(i) <> "" Then Text(i) = UCase(Text(i))
        
   Case 10
        If Text(i) <> "" Then
           strTemp1 = Text(i)
           'edit by nickc 2007/02/07 不用 dll 了
           'If objPublicData.GetAgent(strTemp1, strTempName) Then
           If ClsPDGetAgent(strTemp1, strTempName) Then
              lbe(i) = strTempName
              Text(i) = strTemp1
            End If
         Else
            lbe(i) = ""
        End If
    '2010/2/2 ADD BY SONIA 自前一CASE 10抽出
    Case 14, 17, 24
        If Text(i) <> "" Then
           strTemp1 = Text(i)
           If Left(strTemp1, 1) = "X" Then
               If ClsPDGetCustomer(strTemp1, strTempName) Then
                  lbe(i) = strTempName
                  Text(i) = strTemp1
                End If
           Else
               If ClsPDGetAgent(strTemp1, strTempName) Then
                  lbe(i) = strTempName
                  Text(i) = strTemp1
                End If
           End If
         Else
            lbe(i) = ""
        End If
    '2010/2/2 END
    End Select
End Sub

Private Sub txtcp04_GotFocus()
   TextInverse txtcp04
End Sub

Private Sub txtcp04_LostFocus()
'   If txtcp04 <> "" Then
'      blnCom4 = True
'   End If
'   If blnIsNew Then
   ' 新增模式下檢查資料是否已存在資料庫中
   If m_EditMode = 1 Then
         'edit by nickc 2007/02/07 不用 dll 了
         'If objLawDll.CheckIsExistCaseNum(1, LcTmp, m_Cpnum) Then
         If ClsPDCheckIsExistCaseNum(1, LcTmp, m_Cpnum) Then
         'If CheckIsExistCaseNum(1, LcTmp, m_Cpnum) Then
'            blnCom2 = True
            'TextInverse txtcp02
            'txtcp02.SetFocus
            tlbar.Buttons(11).Enabled = True
         Else
            MsgBox "" + m_Cpnum + "", vbCritical
            TextInverse txtcp02
            txtcp02.SetFocus
         End If
   End If
'   tlbar.Buttons(11).Enabled = True
End Sub

Private Function GetReasonOfRelief(ByVal strNum As String, ByRef strReason As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "select ror02 from reasonofrelief where ror01='" + strNum + "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.EOF = False Then
      GetReasonOfRelief = True
      If Not IsNull(rsTmp.Fields("ROR02")) Then
         strReason = rsTmp.Fields("ROR02")
      End If
   Else
      GetReasonOfRelief = False
   End If
End Function

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   'Add By Sindy 2010/12/16
   Cancel = False
   'Modify By Sindy 2011/1/14 +23,25,26,27
   'Modify By Sindy 2013/12/13 28==>29
   For ii = 0 To 29 '28 '24
      Select Case ii
         Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 24, 23, 25, 26, 27, 29
            Call Text_Validate(ii, Cancel)
            If Cancel = True Then
               Exit Function
            End If
      End Select
   Next ii
   '2010/12/16 End
   
   If Me.txtcp01.Enabled = True Then
      Cancel = False
      txtcp01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtcp02.Enabled = True Then
      Cancel = False
      txtcp02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
'   If Me.txtcp03.Enabled = True Then
'      Cancel = False
'      txtcp03_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
   
   'Add by Morgan 2007/5/10
   If Not ((Text(4).Text = "" And Text(6).Text = "" And Text(7).Text = "") Or (Text(4).Text <> "" And Text(6).Text <> "" And Text(7).Text <> "")) Then
      MsgBox "是否閉卷、閉卷日期、閉卷原因三個欄位須同時空白或有值！", vbExclamation
      Exit Function
   End If
   'end 2007/5/10
   
   'Add By Sindy 2016/11/23
   If Me.Combo4.Enabled = True Then
      Cancel = False
      Combo4_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2016/11/23 End
   
   'add by sonia 2019/8/1
   'Remove by Lydia 2020/11/06 取消ACS的案件屬性欄;
   'If (m_EditMode = 1 Or m_EditMode = 2) And txtcp01 = "ACS" And Combo2.Text = "" Then
   '   MsgBox "請勾選案件屬性!", vbExclamation + vbOKOnly
   '   Exit Function
   'End If
   'If txtcp01 = "ACS" Then 'Added by Morgan 2020/5/22
   '   Text(28) = Combo2
   'End If 'Added by Morgan 2020/5/22
   ''end 2019/8/1
   'end 2020/11/06
   
   'Added by Lydia 2024/06/14 對申請人1~5的重複輸入檢查
   If Pub_ChkAppList(strExc(0), Text(0) & "," & Text(23) & "," & Text(25) & "," & Text(26) & "," & Text(27)) = False Then
      Me.SSTab1.Tab = 0
      Select Case Val(strExc(0))
         Case 1
             Text(0).SetFocus
             Text_GotFocus 0
         Case 2
             Text(23).SetFocus
             Text_GotFocus 23
         Case Else
             Text(Val(strExc(0)) + 22).SetFocus
             Text_GotFocus Val(strExc(0)) + 22
      End Select
      Exit Function
   End If
   'end 2024/06/14
   
   
   'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
   For ii = 0 To 27
      strExc(1) = ""
      Select Case ii
         Case 0, 23, 25, 26, 27, 10 '申請人1~5, FC代理人
            strExc(1) = ChangeCustomerL(Text(ii))
            strExc(2) = ChangeCustomerL(Text(ii).Tag)
      End Select
      If strExc(1) <> "" And strExc(1) <> strExc(2) Then
         If Left(strExc(1), 1) = "X" Then
            If GetCustomerAndState(strExc(1), strExc(3), , , , txtcp01, strExc(8), False, Me.Name, txtcp02, txtcp03, txtcp04) = False Then
               Me.SSTab1.Tab = 0
               Text(ii).SetFocus
               Text_GotFocus ii
               Exit Function
            End If
         Else
            If GetAgentAndState(strExc(1), strExc(3), , , , txtcp01, strExc(2), False) = False Then
               Me.SSTab1.Tab = 1
               Text(ii).SetFocus
               Text_GotFocus ii
               Exit Function
            End If
         End If
      End If
   Next
   'end 2024/06/13
   'Added by Lydia 2021/09/15 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If

   TxtValidate = True
End Function

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   'Add By Sindy 2011/5/31
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyF5 Then
      m_CP01 = txtcp01.Text
      m_CP02 = txtcp02.Text
      If txtcp03.Text <> Empty Then
         m_CP03 = txtcp03.Text
      Else
         m_CP03 = "0"
      End If
      If txtcp04.Text <> Empty Then
         m_CP04 = txtcp04.Text
      Else
         m_CP04 = "00"
      End If
      
      ' 檢查記錄是否不存在
      If IsRecordExist(m_CP01, m_CP02, m_CP03, m_CP04) = False Then
         strTit = "檢查"
         strMsg = "無此資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Exit Sub
      End If
   End If
   
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         If IsCaseProgressExist(txtcp01, txtcp02, txtcp03, txtcp04) = True Then
            strTit = "檢核資料"
            strMsg = "此本所案號在案件進度檔中仍有資料, 不可刪除!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Else
            'Add By Sindy 2010/7/1
            If ChkCaseCode("NP", txtcp01, txtcp02, txtcp03, txtcp04) = False Then Exit Sub
            '2010/7/1 End
            strTit = "詢問"
            strMsg = "是否要刪除此筆資料?"
            nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
            If nResponse = vbYes Then
               m_EditMode = 3
               OnWork
               UpdateToolbarState
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
'edit by nickc 2008/03/28 還沒檢查完資料就先更新，有些資料在檢查時才上，會更新不到
'         UpdateFieldNewData
         OnWork
         UpdateToolbarState
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
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            tlbar.Buttons(1).Enabled = True
         Else
            tlbar.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            tlbar.Buttons(2).Enabled = True
         Else
            tlbar.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            tlbar.Buttons(3).Enabled = True
         Else
            tlbar.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(4).Enabled = True
         Else
            tlbar.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            If Not IsEmptyText(m_FirstRow(0)) And Not IsEmptyText(m_FirstRow(1)) And Not IsEmptyText(m_FirstRow(2)) And Not IsEmptyText(m_FirstRow(3)) Then
               tlbar.Buttons(6).Enabled = True
               tlbar.Buttons(7).Enabled = True
            Else
               tlbar.Buttons(6).Enabled = False
               tlbar.Buttons(7).Enabled = False
            End If
            If Not IsEmptyText(m_LastRow(0)) And Not IsEmptyText(m_LastRow(1)) And Not IsEmptyText(m_LastRow(2)) And Not IsEmptyText(m_LastRow(3)) Then
               tlbar.Buttons(8).Enabled = True
               tlbar.Buttons(9).Enabled = True
            Else
               tlbar.Buttons(8).Enabled = False
               tlbar.Buttons(9).Enabled = False
            End If
         Else
            tlbar.Buttons(6).Enabled = False
            tlbar.Buttons(7).Enabled = False
            tlbar.Buttons(8).Enabled = False
            tlbar.Buttons(9).Enabled = False
         End If
         tlbar.Buttons(11).Enabled = False
         tlbar.Buttons(12).Enabled = False
         tlbar.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         tlbar.Buttons(1).Enabled = False
         tlbar.Buttons(2).Enabled = False
         tlbar.Buttons(3).Enabled = False
         tlbar.Buttons(4).Enabled = False
         tlbar.Buttons(6).Enabled = False
         tlbar.Buttons(7).Enabled = False
         tlbar.Buttons(8).Enabled = False
         tlbar.Buttons(9).Enabled = False
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         tlbar.Buttons(14).Enabled = False
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   m_FirstRow(0) = Empty
   m_FirstRow(1) = Empty
   m_FirstRow(2) = Empty
   m_FirstRow(3) = Empty
   m_CurrRow(0) = Empty
   m_CurrRow(1) = Empty
   m_CurrRow(2) = Empty
   m_CurrRow(3) = Empty
   m_LastRow(0) = Empty
   m_LastRow(1) = Empty
   m_LastRow(2) = Empty
   m_LastRow(3) = Empty
   m_EditMode = 0
   Set frm075002 = Nothing
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   ' 設定 Query 的命令
   strSql = "SELECT LC01,LC02,LC03,LC04 FROM LawCase " & _
            "WHERE LC01 = '" & txtcp01 & "' AND " & _
                  "LC02 = (SELECT MIN(LC02) FROM LawCase WHERE LC01 = '" & txtcp01 & "') AND " & _
                  "LC03 = (SELECT MIN(LC03) FROM LawCase WHERE LC01 = '" & txtcp01 & "' AND LC02 = (SELECT MIN(LC02) FROM LawCase WHERE LC01 = '" & txtcp01 & "' )) AND " & _
                  "LC04 = (SELECT MIN(LC04) FROM LawCase WHERE LC01 = '" & txtcp01 & "' AND LC02 = (SELECT MIN(LC02) FROM LawCase WHERE LC01 = '" & txtcp01 & "' ) AND LC03 = (SELECT MIN(LC03) FROM LawCase WHERE LC01 = '" & txtcp01 & "' AND LC02 = (SELECT MIN(LC02) FROM LawCase WHERE LC01 = '" & txtcp01 & "' ))) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LC01")) = False Then: m_FirstRow(0) = rsTmp.Fields("LC01")
      If IsNull(rsTmp.Fields("LC02")) = False Then: m_FirstRow(1) = rsTmp.Fields("LC02")
      If IsNull(rsTmp.Fields("LC03")) = False Then: m_FirstRow(2) = rsTmp.Fields("LC03")
      If IsNull(rsTmp.Fields("LC04")) = False Then: m_FirstRow(3) = rsTmp.Fields("LC04")
   End If
   rsTmp.Close

   ' 設定 Query 的命令
   strSql = "SELECT LC01,LC02,LC03,LC04 FROM LawCase " & _
            "WHERE LC01 = '" & txtcp01 & "' AND " & _
                  "LC02 = (SELECT MAX(LC02) FROM LawCase WHERE LC01 = '" & txtcp01 & "') AND " & _
                  "LC03 = (SELECT MAX(LC03) FROM LawCase WHERE LC01 = '" & txtcp01 & "' AND LC02 = (SELECT MAX(LC02) FROM LawCase WHERE LC01 = '" & txtcp01 & "' )) AND " & _
                  "LC04 = (SELECT MAX(LC04) FROM LawCase WHERE LC01 = '" & txtcp01 & "' AND LC02 = (SELECT MAX(LC02) FROM LawCase WHERE LC01 = '" & txtcp01 & "' ) AND LC03 = (SELECT MAX(LC03) FROM LawCase WHERE LC01 = '" & txtcp01 & "' AND LC02 = (SELECT MAX(LC02) FROM LawCase WHERE LC01 = '" & txtcp01 & "' ))) "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LC01")) = False Then: m_LastRow(0) = rsTmp.Fields("LC01")
      If IsNull(rsTmp.Fields("LC02")) = False Then: m_LastRow(1) = rsTmp.Fields("LC02")
      If IsNull(rsTmp.Fields("LC03")) = False Then: m_LastRow(2) = rsTmp.Fields("LC03")
      If IsNull(rsTmp.Fields("LC04")) = False Then: m_LastRow(3) = rsTmp.Fields("LC04")
   End If
   rsTmp.Close
  
   Set rsTmp = Nothing
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strLC01 As String, ByVal strLC02 As String, ByVal strLC03 As String, ByVal strLC04 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strLC01, strLC02, strLC03, strLC04) = True Then
      m_CurrRow(0) = strLC01
      m_CurrRow(1) = strLC02
      m_CurrRow(2) = strLC03
      m_CurrRow(3) = strLC04
   Else
      strSql = "SELECT LC01,LC02,LC03,LC04 FROM LawCase " & _
               "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                     "LC02 = '" & m_CurrRow(1) & "' AND " & _
                     "LC03 = '" & m_CurrRow(2) & "' AND " & _
                     "LC04 = (SELECT MIN(LC04) FROM LawCase " & _
                             "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                   "LC03 = '" & m_CurrRow(2) & "' AND " & _
                                   "LC04 > '" & m_CurrRow(3) & "' )"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("LC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("LC01")
         If IsNull(rsTmp.Fields("LC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("LC02")
         If IsNull(rsTmp.Fields("LC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("LC03")
         If IsNull(rsTmp.Fields("LC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("LC04")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
   
      strSql = "SELECT LC01,LC02,LC03,LC04 FROM LawCase " & _
               "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                     "LC02 = '" & m_CurrRow(1) & "' AND " & _
                     "LC03 = (SELECT MIN(LC03) FROM LawCase " & _
                             "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                   "LC03 > '" & m_CurrRow(2) & "') AND " & _
                     "LC04 = (SELECT MIN(LC04) FROM LawCase " & _
                             "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                   "LC03 = (SELECT MIN(LC03) FROM LawCase " & _
                                           "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                                 "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                                 "LC03 > '" & m_CurrRow(2) & "'))"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("LC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("LC01")
         If IsNull(rsTmp.Fields("LC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("LC02")
         If IsNull(rsTmp.Fields("LC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("LC03")
         If IsNull(rsTmp.Fields("LC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("LC04")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
                                
      strSql = "SELECT LC01,LC02,LC03,LC04 FROM LawCase " & _
               "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                     "LC02 = (SELECT MIN(LC02) FROM LawCase " & _
                             "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "LC02 > '" & m_CurrRow(1) & "') AND " & _
                     "LC03 = (SELECT MIN(LC03) FROM LawCase " & _
                             "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "LC02 = (SELECT MIN(LC02) FROM LawCase " & _
                                           "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                                 "LC02 > '" & m_CurrRow(1) & "')) AND " & _
                     "LC04 = (SELECT MIN(LC04) FROM LawCase " & _
                             "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                   "LC02 = (SELECT MIN(LC02) FROM LawCase " & _
                                           "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                                 "LC02 > '" & m_CurrRow(1) & "') AND " & _
                                                 "LC03 = (SELECT MIN(LC03) FROM LawCase " & _
                                                         "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                                               "LC02 = (SELECT MIN(LC02) FROM LawCase " & _
                                                                       "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                                                             "LC02 > '" & m_CurrRow(1) & "'))) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("LC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("LC01")
         If IsNull(rsTmp.Fields("LC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("LC02")
         If IsNull(rsTmp.Fields("LC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("LC03")
         If IsNull(rsTmp.Fields("LC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("LC04")
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
   m_CurrRow(0) = m_FirstRow(0)
   m_CurrRow(1) = m_FirstRow(1)
   m_CurrRow(2) = m_FirstRow(2)
   m_CurrRow(3) = m_FirstRow(3)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrRow(0) = m_FirstRow(0) And m_CurrRow(1) = m_FirstRow(1) And m_CurrRow(2) = m_FirstRow(2) And m_CurrRow(3) = m_FirstRow(3) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT LC01,LC02,LC03,LC04 FROM LawCase " & _
            "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                  "LC02 = '" & m_CurrRow(1) & "' AND " & _
                  "LC03 = '" & m_CurrRow(2) & "' AND " & _
                  "LC04 = (SELECT MAX(LC04) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                "LC03 = '" & m_CurrRow(2) & "' AND " & _
                                "LC04 < '" & m_CurrRow(3) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("LC01")
      If IsNull(rsTmp.Fields("LC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("LC02")
      If IsNull(rsTmp.Fields("LC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("LC03")
      If IsNull(rsTmp.Fields("LC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("LC04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT LC01,LC02,LC03,LC04 FROM LawCase " & _
            "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                  "LC02 = '" & m_CurrRow(1) & "' AND " & _
                  "LC03 = (SELECT MAX(LC03) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                "LC03 < '" & m_CurrRow(2) & "') AND " & _
                  "LC04 = (SELECT MAX(LC04) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                "LC03 = (SELECT MAX(LC03) FROM LawCase " & _
                                        "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                              "LC03 < '" & m_CurrRow(2) & "'))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("LC01")
      If IsNull(rsTmp.Fields("LC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("LC02")
      If IsNull(rsTmp.Fields("LC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("LC03")
      If IsNull(rsTmp.Fields("LC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("LC04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT LC01,LC02,LC03,LC04 FROM LawCase " & _
            "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                  "LC02 = (SELECT MAX(LC02) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 < '" & m_CurrRow(1) & "') AND " & _
                  "LC03 = (SELECT MAX(LC03) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 = (SELECT MAX(LC02) FROM LawCase " & _
                                        "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "LC02 < '" & m_CurrRow(1) & "')) AND " & _
                  "LC04 = (SELECT MAX(LC04) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 = (SELECT MAX(LC02) FROM LawCase " & _
                                        "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "LC02 < '" & m_CurrRow(1) & "') AND " & _
                                              "LC03 = (SELECT MAX(LC03) FROM LawCase " & _
                                                      "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                                            "LC02 = (SELECT MAX(LC02) FROM LawCase " & _
                                                                    "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                                                          "LC02 < '" & m_CurrRow(1) & "'))) "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("LC01")
      If IsNull(rsTmp.Fields("LC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("LC02")
      If IsNull(rsTmp.Fields("LC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("LC03")
      If IsNull(rsTmp.Fields("LC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("LC04")
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
   
   If m_CurrRow(0) = m_LastRow(0) And m_CurrRow(1) = m_LastRow(1) And m_CurrRow(2) = m_LastRow(2) And m_CurrRow(3) = m_LastRow(3) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT LC01,LC02,LC03,LC04 FROM LawCase " & _
            "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                  "LC02 = '" & m_CurrRow(1) & "' AND " & _
                  "LC03 = '" & m_CurrRow(2) & "' AND " & _
                  "LC04 = (SELECT MIN(LC04) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                "LC03 = '" & m_CurrRow(2) & "' AND " & _
                                "LC04 > '" & m_CurrRow(3) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("LC01")
      If IsNull(rsTmp.Fields("LC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("LC02")
      If IsNull(rsTmp.Fields("LC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("LC03")
      If IsNull(rsTmp.Fields("LC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("LC04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT LC01,LC02,LC03,LC04 FROM LawCase " & _
            "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                  "LC02 = '" & m_CurrRow(1) & "' AND " & _
                  "LC03 = (SELECT MIN(LC03) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                "LC03 > '" & m_CurrRow(2) & "') AND " & _
                  "LC04 = (SELECT MIN(LC04) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                "LC03 = (SELECT MIN(LC03) FROM LawCase " & _
                                        "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "LC02 = '" & m_CurrRow(1) & "' AND " & _
                                              "LC03 > '" & m_CurrRow(2) & "'))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("LC01")
      If IsNull(rsTmp.Fields("LC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("LC02")
      If IsNull(rsTmp.Fields("LC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("LC03")
      If IsNull(rsTmp.Fields("LC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("LC04")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
                                
   strSql = "SELECT LC01,LC02,LC03,LC04 FROM LawCase " & _
            "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                  "LC02 = (SELECT MIN(LC02) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 > '" & m_CurrRow(1) & "') AND " & _
                  "LC03 = (SELECT MIN(LC03) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 = (SELECT MIN(LC02) FROM LawCase " & _
                                        "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "LC02 > '" & m_CurrRow(1) & "')) AND " & _
                  "LC04 = (SELECT MIN(LC04) FROM LawCase " & _
                          "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                "LC02 = (SELECT MIN(LC02) FROM LawCase " & _
                                        "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                              "LC02 > '" & m_CurrRow(1) & "') AND " & _
                                              "LC03 = (SELECT MIN(LC03) FROM LawCase " & _
                                                      "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                                            "LC02 = (SELECT MIN(LC02) FROM LawCase " & _
                                                                    "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                                                                          "LC02 > '" & m_CurrRow(1) & "'))) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("LC01")) = False Then: m_CurrRow(0) = rsTmp.Fields("LC01")
      If IsNull(rsTmp.Fields("LC02")) = False Then: m_CurrRow(1) = rsTmp.Fields("LC02")
      If IsNull(rsTmp.Fields("LC03")) = False Then: m_CurrRow(2) = rsTmp.Fields("LC03")
      If IsNull(rsTmp.Fields("LC04")) = False Then: m_CurrRow(3) = rsTmp.Fields("LC04")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrRow(0) = m_LastRow(0)
   m_CurrRow(1) = m_LastRow(1)
   m_CurrRow(2) = m_LastRow(2)
   m_CurrRow(3) = m_LastRow(3)
   
   UpdateCtrlData
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
Dim i As Integer
   For i = 0 To 30
      Text(i) = Empty
   Next i
   Text(52) = Empty 'Add by Amy 2018/08/15
   txtcp01 = Empty
   txtcp02 = Empty
   txtcp03 = Empty
   txtcp04 = Empty
   IDname = Empty
   CDT = Empty
   CTM = Empty
   UIDname = Empty
   UDT = Empty
   UTM = Empty
   
   If IsEmpty(m_CurrRow(0)) = False Then
      txtcp01 = m_CurrRow(0)
   End If
   cboContact.Clear 'Add by Morgan 2008/8/4
   'Add By Sindy 2011/5/31
   'Modified by Lydia 2022/08/10
   For Each oObj In Check1
       oObj = Empty
   Next
   'end 2022/08/10
   'Added by Lydia 2023/03/14
   For Each oObj In Check2
       oObj = Empty
   Next
   'end 2023/03/14
   'Add By Sindy 2016/11/23
   Me.Combo4.ListIndex = 0
   Me.Combo5.ListIndex = 0
   '2016/11/23 End
   
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   txtcp01.Locked = bEnable: txtcp02.Locked = bEnable: txtcp03.Locked = bEnable: txtcp04.Locked = bEnable
   For i = 0 To 29
      'Modify by Amy 2018/07/03 只有電腦中心才可改 特殊出名公司
      If i = 29 Then
        Text(i).Locked = True
        If Pub_StrUserSt03 = "M51" Then Text(i).Locked = False
      Else
        Text(i).Locked = bEnable
      End If
   Next i
   
   Text(52).Locked = bEnable 'Add by Amy 2018/08/15 +專案服務案
   'Modified by Lydia 2022/08/10
   For Each oObj In Check1
       If bEnable = False Then
           oObj.Enabled = True
       Else
           oObj.Enabled = False
       End If
   Next
   'end 2022/08/10
   'Added by Lydia 2023/03/14
   For Each oObj In Check2
       If bEnable = False Then
           oObj.Enabled = True
       Else
           oObj.Enabled = False
       End If
   Next
   'end 2023/03/14
   'Add By Sindy 2016/11/23
   Combo4.Locked = bEnable
   Combo5.Locked = bEnable
   '2016/11/23 End
   cboContact.Locked = bEnable 'Added by Lydia 2021/09/15
   
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   txtcp01.Locked = bEnable: txtcp02.Locked = bEnable: txtcp03.Locked = bEnable: txtcp04.Locked = bEnable
End Sub

'Add By Sindy 2011/5/31
' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, strTemp As String
   
   'add by sonia 2019/8/1
   If txtcp01 = "ACS" Then
      'Modified by Lydia 2020/11/06 取消ACS的案件屬性欄;
      'Combo2.Enabled = True
      'Combo2.Visible = True
      'Combo2 = ""
      'Combo2.Left = 1116
      'Combo2.Top = 3852
      Label31.Visible = False
      'end 2020/11/06
      Text(28).Enabled = False
      Text(28).Visible = False
      'Modified by Lydia 2022/08/10
      For Each oObj In Check1
          oObj.Visible = False
          oObj.Enabled = False
      Next
      'end 2022/08/10
      'Added by Lydia 2023/03/14 一般案件屬性
      Frame2.Visible = False
      Frame2.Enabled = False
   Else
      Combo2.Enabled = False
      Combo2.Visible = False
      Label31.Visible = True 'Added by Lydia 2020/11/06
      Text(28).Enabled = True
      Text(28).Visible = True
      Text(28) = ""
      'Modified by Lydia 2022/08/10
      For Each oObj In Check1
          oObj.Visible = True
          oObj.Enabled = True
      Next
      'end 2022/08/10
   End If
   'end 2019/8/1
   
   'Add By Sindy 2013/12/13 +lc48
   'Add By Sindy 2016/11/23 +lc49,lc50
   'Modified by Morgan 2018/4/11 +lc51
   'Modify by Amy 2018/08/15 +lc52
   strSql = "SELECT lc11,lc05,lc06,lc07,lc08,lc16,lc09,lc10,lc15,lc17," + _
                   "lc22,lc23,lc13,lc14,lc26,lc25,lc24,lc12,lc21,lc18,lc19,lc20," + _
                   "lc27,lc43,LC35,lc44,lc45,lc46,lc28,lc29,lc30,lc31,lc32,lc33," + _
                   "lc01,lc02,lc03,lc04,lc34,lc36,lc37,lc38,lc42,lc47,lc48,lc49,lc50,lc51,lc52 " + _
             "FROM LawCase " & _
            "WHERE LC01 = '" & m_CurrRow(0) & "' AND " & _
                  "LC02 = '" & m_CurrRow(1) & "' AND " & _
                  "LC03 = '" & m_CurrRow(2) & "' AND " & _
                  "LC04 = '" & m_CurrRow(3) & "' "
               
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      txtcp01 = rsTmp.Fields("LC01")
      txtcp02 = rsTmp.Fields("LC02")
      txtcp03 = rsTmp.Fields("LC03")
      txtcp04 = rsTmp.Fields("LC04")
      lbeNumber.Tag = GiveSymbol(txtcp01, txtcp02, txtcp03, txtcp04, LcTmp)
      For i = 0 To 29 '28
         If i = 28 Then
            'modify by sonia 2019/8/14
            'Text(i) = IIf(IsNull(rsTmp.Fields("lc47")), "", rsTmp.Fields("lc47"))
            'Text(i).Tag = Text(i) 'Add by Amy 2018/07/30
            If txtcp01 <> "ACS" Then
               Text(i) = IIf(IsNull(rsTmp.Fields("lc47")), "", rsTmp.Fields("lc47"))
               Text(i).Tag = Text(i)
            Else
               Combo2 = IIf(IsNull(rsTmp.Fields("lc47")), "", rsTmp.Fields("lc47"))
               Combo2.Tag = Text(i)
            End If
            'end 2019/8/14
         ElseIf i = 29 Then
            Text(i) = IIf(IsNull(rsTmp.Fields("lc48")), "", rsTmp.Fields("lc48"))
         Else
            Text(i) = IIf(IsNull(rsTmp.Fields(i)), "", rsTmp.Fields(i))
         End If
         'Added by Lydia 2024/06/13
         If i <> 28 Then
            Text(i).Tag = Text(i)
         End If
         'end 2024/06/13
         ChgType (i)
      Next
      
      Text(30) = "" & rsTmp.Fields("lc51") 'Added by Morgan 2018/4/11
      Text(52) = "" & rsTmp.Fields("lc52") 'Add by Amy 2018/08/15
      
      '案件屬性
      'Modified by Lydia 2022/08/10
      For Each oObj In Check1
         If InStr(Text(28).Text, Trim(oObj.Caption)) > 0 Then
            oObj.Value = 1
         End If
      Next
      'end 2022/08/10
      'Added by Lydia 2023/03/14 一般案件屬性
      If txtcp01 <> "ACS" Then
         If Check1(4).Value = 1 Then
            Frame2.Enabled = True
         Else
            Frame2.Enabled = False
         End If
         For Each oObj In Check2
            If InStr(Text(28).Text, Trim(oObj.Caption)) > 0 Then
               oObj.Value = 1
               If Frame2.Enabled = False Then
                  Frame2.Enabled = True
               End If
            End If
         Next
      End If
      'end 2023/03/14
      'Add By Sindy 2016/11/23
      If IsNull(rsTmp.Fields("lc49")) = False Then
         For i = 0 To Combo4.ListCount - 1
            Combo4.ListIndex = i
            If InStr(Combo4.Text, rsTmp.Fields("lc49")) > 0 Then
               Exit For
            End If
         Next
      Else
         Combo4.ListIndex = 0
      End If
      If IsNull(rsTmp.Fields("lc50")) = False Then
         Combo5.ListIndex = rsTmp.Fields("lc50")
      Else
         Combo5.ListIndex = 0
      End If
      '2016/11/23 End
      
      IDname = IIf(IsNull(rsTmp.Fields!lc28), "", rsTmp.Fields!lc28)
      IDname = GetStaffName(IDname, True)
      
      '新增日期
      strTemp = ""
      CDT = IIf(IsNull(rsTmp.Fields!lc29), "", rsTmp.Fields!lc29)
      strTemp = TAIWANDATE(CDT)
      CDT = Format(strTemp, "###/##/##")
      '新增時間
      CTM = IIf(IsNull(rsTmp.Fields!lc30), "", rsTmp.Fields!lc30)
      CTM = Format(CTM, "##:##")
      
      UIDname = IIf(IsNull(rsTmp.Fields!lc31), "", rsTmp.Fields!lc31)
      UIDname = GetStaffName(UIDname, True)
      
      '修改日期
      strTemp = ""
      UDT = IIf(IsNull(rsTmp.Fields!lc32), "", rsTmp.Fields!lc32)
      strTemp = TAIWANDATE(UDT)
      UDT = Format(strTemp, "###/##/##")
      '新增時間
      UTM = IIf(IsNull(rsTmp.Fields!lc33), "", rsTmp.Fields!lc33)
      UTM = Format(UTM, "##:##")
      
      strTemp = ""
      lblLC34 = IIf(IsNull(rsTmp.Fields!lc34), "", rsTmp.Fields!lc34)
      strTemp = TAIWANDATE(lblLC34)
      lblLC34 = Format(strTemp, "###/##/##")
      strTemp = ""
      lblLC36 = IIf(IsNull(rsTmp.Fields!lc36), "", rsTmp.Fields!lc36)
      strTemp = TAIWANDATE(lblLC36)
      lblLC36 = Format(strTemp, "###/##/##")
      lblLC37 = IIf(IsNull(rsTmp.Fields!lc37), "", rsTmp.Fields!lc37)
      lblLC37 = GetStaffName(lblLC37, True)
      lblLC38 = IIf(IsNull(rsTmp.Fields!lc38), "", rsTmp.Fields!lc38)
      
      'Modified by Lydia 2021/09/15 改成Form 2.0
      'PUB_AddContact "" & rsTmp.Fields!LC11, cboContact, "" & rsTmp.Fields!lc42 'Add by Morgan 2008/8/4
      PUB_AddContact "" & rsTmp.Fields!LC11, cboContact, "" & rsTmp.Fields!lc42, , True
      
      m_CP01 = txtcp01 '2011/5/20 add by sonia
      
      'Added by Lydia 2021/01/14 法律所案源收文：讀取案源
      If InStr(txtcp01, "L") > 0 And txtcp01 <> "ACS" Then
          Call ReadLOS
      End If
      'end 2021/01/14
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   QueryRecord = False
   
   If IsEmptyText(txtcp03) = True Then: txtcp03 = "0"
   If IsEmptyText(txtcp04) = True Then: txtcp04 = "00"
   
   If IsRecordExist(txtcp01, txtcp02, txtcp03, txtcp04) = True Then
      m_CurrRow(0) = txtcp01
      m_CurrRow(1) = txtcp02
      m_CurrRow(2) = txtcp03
      m_CurrRow(3) = txtcp04
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   ' 當系統別不為原先所輸入的系統別時則需重新取得範圍
   If txtcp01 <> m_CurrRow(0) Then
      RefreshRange
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
Dim strMsg As String
Dim strTit As String
Dim nResponse
Dim StrSQLa As String            '2009/8/19 ADD BY SONIA
Dim rsA As New ADODB.Recordset   '2009/8/19 ADD BY SONIA
   
   Select Case m_EditMode
      Case 1:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            'add by nickc 2008/03/28  更新欄位
'            UpdateFieldNewData
            'edit by nickc 2006/06/08
            'AddRecord
            If AddRecord = False Then Exit Sub
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            'add by nickc 2008/03/28  更新欄位
'            UpdateFieldNewData
            'edit by nickc 2006/06/08
            'ModRecord
            
            'Added by Lydia 2017/06/19 (存檔前)檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
            strChkCuAreaMail = PUB_ChkSameCustSales(Trim(txtcp01), Trim(txtcp02), Trim(txtcp03), Trim(txtcp04), "", Trim(Text(0)), Trim(Text(23)), Trim(Text(25)), Trim(Text(26)), Trim(Text(27)), strChkCuAreaMailTo)
            
            If ModRecord = False Then Exit Sub
            
            'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
            If strChkCuAreaMail <> "" Then
               PUB_SendMail strUserNum, strChkCuAreaMailTo, "", "案件收文通知--此案收文非原智權人員(區)！", strChkCuAreaMail
            End If
            'end 2017/06/19
         Else
            GoTo EXITSUB
         End If
      Case 3:
         'edit by nickc 2006/06/08
         'DelRecord
         If DelRecord = False Then Exit Sub
        'add by nickc 2008/03/28  更新欄位
'         UpdateFieldNewData
         
         RefreshRange
      Case 4:
         If TxtValidate = False Then Exit Sub
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: txtcp01.SetFocus
      Case 2: Text(0).SetFocus
      'modify by sonia 2020/11/19
      'Case 4: txtCP01.SetFocus
      Case 4:
         If txtcp01 = "ACS" Then
            txtcp02.SetFocus
         Else
            txtcp01.SetFocus
         End If
      'end 2020/11/19
   End Select
End Sub

' 案件進度檔
Private Function IsCaseProgressExist(ByVal strLC01 As String, ByVal strLC02 As String, ByVal strLC03 As String, ByVal strLC04 As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   IsCaseProgressExist = False
   strSql = "SELECT * from CaseProgress " & _
            "WHERE CP01 = '" & strLC01 & "' AND " & _
                  "CP02 = '" & strLC02 & "' AND " & _
                  "CP03 = '" & strLC03 & "' AND " & _
                  "CP04 = '" & strLC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsCaseProgressExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strLC01 As String, ByVal strLC02 As String, ByVal strLC03 As String, ByVal strLC04 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM LawCase " & _
            "WHERE LC01 = '" & strLC01 & "' AND " & _
                  "LC02 = '" & strLC02 & "' AND " & _
                  "LC03 = '" & strLC03 & "' AND " & _
                  "LC04 = '" & strLC04 & "'"
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function CheckDataValid() As Boolean
   CheckDataValid = False
   
   If txtcp01 = "" Then
       DataErrorMessage 5, "本所案號"
       txtcp01.SetFocus
       Exit Function
   End If
   If txtcp02 = "" Then
       DataErrorMessage 5, "本所案號"
       txtcp02.SetFocus
       Exit Function
   End If
   
   '若為新增狀態
   If m_EditMode = 1 Then
       '若本所案號的流水號非6碼
       If Len(Me.txtcp02.Text) <> 6 Then
           MsgBox "本所案號的流水號必須為 6 碼，不滿 6 碼者請在前面補零!!!", vbExclamation + vbOKOnly
           txtcp02.SetFocus
           txtcp02_GotFocus
           Exit Function
       End If
   End If
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If CheckLengthIsOK(Text(1), 160) = False Then
         Text(1).SetFocus
         TextInverse Text(1)
         Exit Function
      End If
      If CheckLengthIsOK(Text(3), 160) = False Then
         Text(3).SetFocus
         TextInverse Text(3)
         Exit Function
      End If
      If CheckLengthIsOK(Text(13), 10) = False Then
         Text(13).SetFocus
         TextInverse Text(13)
         Exit Function
      End If
      'Modified by Lydia 2017/06/14
      'If CheckLengthIsOK(Text(19), 10) = False Then
      If CheckLengthIsOK(Text(19), 30) = False Then
         Text(19).SetFocus
         TextInverse Text(19)
         Exit Function
      End If
      'Added by Lydia 2017/06/14 聯絡人(英)
      If CheckLengthIsOK(Text(20), 35) = False Then
         Text(20).SetFocus
         TextInverse Text(20)
         Exit Function
      End If
      'Modified by Lydia 2017/06/14
      'If CheckLengthIsOK(Text(21), 20) = False Then
      If CheckLengthIsOK(Text(21), 60) = False Then
         Text(21).SetFocus
         TextInverse Text(21)
         Exit Function
      End If
      If CheckLengthIsOK(Text(22), 2000) = False Then
         Text(22).SetFocus
         TextInverse Text(22)
         Exit Function
      End If
   
       If Text(1).Text = "" And Text(2).Text = "" And Text(3).Text = "" Then
           MsgBox "案件名稱不可同時為空", vbCritical
           Text(1).SetFocus
           Exit Function
       End If
   
       If Text(8) = "" Or IsNull(Text(8)) Then
           DataErrorMessage 5, "相關國家"
           Text(8).SetFocus
           Exit Function
       End If
      
      'Add By Sindy 2011/6/20
      If (InStr(Text(28).Text, "專利") > 0 Or InStr(Text(28).Text, "商標") > 0 Or _
         InStr(Text(28).Text, "著作權") > 0 Or InStr(Text(28).Text, "智財權") > 0) And Text(12) <> "Y" Then
         MsgBox "案件屬性欄位出現專利,商標,著作權或智財權字樣時,智財權案須為Y!!!", vbExclamation + vbOKOnly
         Text(12).SetFocus
         Exit Function
      End If
      
      'Add By Sindy 2016/11/23
      If Trim(Me.Combo4.Text) <> "" Then
         '若輸入幣別就一定要選格式
         If Trim(Me.Combo5.Text) = "" Then
            ShowMsg "請款單列印幣別格式不可空白 !"
            Me.Combo5.SetFocus
            Exit Function
         End If
         '請款幣別<>NTD時不可輸入1
         If Trim(Me.Combo4.Text) <> "NTD" And Me.Combo5.ListIndex = 1 Then
            ShowMsg "請款幣別<>NTD時，請款單列印幣別格式不可選純台幣 !"
            Me.Combo5.SetFocus
            Exit Function
         End If
      End If
      '2016/11/23 ENd
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

' 新增記錄
Private Function AddRecord() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   AddRecord = False
   
   If txtcp03 = "" Then txtcp03 = "0"
   If txtcp04 = "" Then txtcp04 = "00"
   m_CP01 = txtcp01.Text
   m_CP02 = txtcp02.Text
   If txtcp03.Text <> "" Then
      m_CP03 = txtcp03.Text
   Else
      m_CP03 = "0"
   End If
   If txtcp04.Text <> "" Then
      m_CP04 = txtcp04.Text
   Else
      m_CP04 = "00"
   End If
   
   ' 檢查記錄是否已存在
   If IsRecordExist(m_CP01, m_CP02, m_CP03, m_CP04) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   'Add By Cheng 2003/02/25
   '若當事人有輸入則補滿9碼
   If Me.Text(0).Text <> "" Then
       Me.Text(0).Text = Left(Me.Text(0).Text & "000000000", 9)
   End If
   'Modify By Sindy 2011/1/14 若當事人2-5有輸入則補滿9碼
   If Me.Text(23).Text <> "" Then
       Me.Text(23).Text = Left(Me.Text(23).Text & "000000000", 9)
   End If
   If Me.Text(25).Text <> "" Then
       Me.Text(25).Text = Left(Me.Text(25).Text & "000000000", 9)
   End If
   If Me.Text(26).Text <> "" Then
       Me.Text(26).Text = Left(Me.Text(26).Text & "000000000", 9)
   End If
   If Me.Text(27).Text <> "" Then
       Me.Text(27).Text = Left(Me.Text(27).Text & "000000000", 9)
   End If
   '2011/1/14 End
   '若FC代理人有輸入則補滿9碼
   If Me.Text(10).Text <> "" Then
       Me.Text(10).Text = Left(Me.Text(10).Text & "000000000", 9)
   End If
   '若固定請款對象有輸入則補滿9碼
   If Me.Text(14).Text <> "" Then
       Me.Text(14).Text = Left(Me.Text(14).Text & "000000000", 9)
   End If
   
   'Add By Sindy 2011/1/14 +當事人2-5
   'Add By Sindy 2011/5/31 +案件屬性
   'Add By Sindy 2013/12/13 +LC48
   'Add By Sindy 2016/11/23 +LC49,LC50
   'Modified by Morgan 2018/4/11 +LC51
   'Modify by Amy 2018/08/15 +LC52
   'Modified by Lydia 2021/02/08 debug 案件屬性 CNULL(ChangeCustomerL(Text(28))) => CNULL(Text(28))
   strExc(1) = "insert into lawcase (lc01,lc02 ,lc03,lc04 ,lc05,lc06,lc07,lc08,lc09,lc10,lc11,lc12,lc13," + _
     "lc14,lc15,lc16,lc17,lc18,lc19,lc20,lc21,lc22,lc23,lc24,lc25,lc26,lc27,LC35,LC43,LC44,LC45,LC46,LC47,LC48,LC49,LC50,LC51,LC52) " + _
     "values (" + CNULL(txtcp01) + "," + _
     CNULL(txtcp02) + "," + CNULL(txtcp03) + "," + CNULL(txtcp04) + "," + CNULL(ChgSQL(Text(1))) + "," + CNULL(ChgSQL(Text(2))) + "," + CNULL(ChgSQL(Text(3))) + _
     "," + CNULL(Text(4)) + "," + CNULL(IIf(Text(6) = "", "", ChangeTStringToWString(Text(6)))) + "," + CNULL(Text(7)) + "," + CNULL(ChangeCustomerL(Text(0))) + _
     "," + CNULL(ChangeCustomerL(Text(17), True)) + "," + CNULL(Text(12)) + "," + CNULL(Text(13)) + "," + CNULL(Text(8)) + "," + CNULL(Text(5)) + _
     "," + CNULL(ChgSQL(Text(9))) + "," + CNULL(ChgSQL(Text(19))) + "," + CNULL(ChgSQL(Text(20))) + " ," + CNULL(ChgSQL(Text(21))) + "," + CNULL(Text(18)) + _
     "," + CNULL(ChangeCustomerL(Text(10), True)) + "," + CNULL(ChgSQL(Text(11))) + "," + CNULL(Text(16)) + "," + CNULL(Text(15)) + _
     "," + CNULL(ChangeCustomerL(Text(14), True)) + "," + CNULL(ChgSQL(Text(22))) + "," + CNULL(ChangeCustomerL(Text(24))) + _
     "," + CNULL(ChangeCustomerL(Text(23))) + "," + CNULL(ChangeCustomerL(Text(25))) + "," + CNULL(ChangeCustomerL(Text(26))) + _
     "," + CNULL(ChangeCustomerL(Text(27))) + "," + CNULL(Text(28)) + "," + CNULL(Text(29)) + _
     "," + CNULL(Combo4.Text) + "," + CNULL(IIf(Combo5.Text <> "", Combo5.ListIndex, "")) + ",'" & ChgSQL(Text(30)) & "','" + ChgSQL(Text(52)) + "')"
   
   On Error GoTo oErr
   cnnConnection.BeginTrans
   '紀錄分析語法
   Pub_SeekTbLog strExc(1)
   cnnConnection.Execute strExc(1)
   cnnConnection.CommitTrans
   
   If ((m_CP01 & m_CP02 & m_CP03 & m_CP04) < (m_FirstRow(0) & m_FirstRow(1) & m_FirstRow(2) & m_FirstRow(3))) Or ((m_CP01 & m_CP02 & m_CP03 & m_CP04) > (m_LastRow(0) & m_LastRow(1) & m_LastRow(2) & m_LastRow(3))) Then
      RefreshRange
   End If
   
   ShowCurrRecord m_CP01, m_CP02, m_CP03, m_CP04
   AddRecord = True
EXITSUB:
Exit Function
oErr:
    cnnConnection.RollbackTrans
    MsgBox Err.Description
End Function

' 修改記錄
Private Function ModRecord() As Boolean
   ModRecord = False
   
   m_CP01 = txtcp01.Text
   m_CP02 = txtcp02.Text
   If txtcp03.Text <> "" Then
      m_CP03 = txtcp03.Text
   Else
      m_CP03 = "0"
   End If
   If txtcp04.Text <> "" Then
      m_CP04 = txtcp04.Text
   Else
      m_CP04 = "00"
   End If
      
   'Add By Cheng 2003/02/25
   '若當事人有輸入則補滿9碼
   If Me.Text(0).Text <> "" Then
       Me.Text(0).Text = Left(Me.Text(0).Text & "000000000", 9)
   End If
   'Modify By Sindy 2011/1/14 若當事人2-5有輸入則補滿9碼
   If Me.Text(23).Text <> "" Then
       Me.Text(23).Text = Left(Me.Text(23).Text & "000000000", 9)
   End If
   If Me.Text(25).Text <> "" Then
       Me.Text(25).Text = Left(Me.Text(25).Text & "000000000", 9)
   End If
   If Me.Text(26).Text <> "" Then
       Me.Text(26).Text = Left(Me.Text(26).Text & "000000000", 9)
   End If
   If Me.Text(27).Text <> "" Then
       Me.Text(27).Text = Left(Me.Text(27).Text & "000000000", 9)
   End If
   '2011/1/14 End
   '若FC代理人有輸入則補滿9碼
   If Me.Text(10).Text <> "" Then
       Me.Text(10).Text = Left(Me.Text(10).Text & "000000000", 9)
   End If
   '若固定請款對象有輸入則補滿9碼
   If Me.Text(14).Text <> "" Then
       Me.Text(14).Text = Left(Me.Text(14).Text & "000000000", 9)
   End If
   UDT = ChangeWDateStringToWString(Date)
   UTM = Format(time, "HHMM")
   'Modify By Sindy 2011/1/14 +當事人2,3,4,5
   'Add By Sindy 2011/5/31 +案件屬性
   'Add By Sindy 2013/12/13 +LC48
   'Add By Sindy 2016/11/23 +LC49,LC50
   'Modified by Morgan 2018/4/11 +LC51
   'Modify by Amy 2018/08/15 +LC52
   strExc(1) = " update lawcase set lc05=" + CNULL(ChgSQL(Text(1))) + ",lc06=" + CNULL(ChgSQL(Text(2))) + ",lc07=" + CNULL(Text(3)) + _
        ",lc08=" + CNULL(Text(4)) + ",lc09=" + CNULL(IIf(Text(6) = "", "", ChangeTStringToWString(Text(6)))) + _
        ",lc10=" + CNULL(Text(7)) + ",lc11=" + CNULL(ChangeCustomerL(Text(0))) + _
        ",lc12=" + CNULL(ChangeCustomerL(Text(17))) + ",lc13=" + CNULL(Text(12)) + ",lc14=" + CNULL(Text(13)) + ",lc15=" + CNULL(Text(8)) + _
        ",lc16=" + CNULL(Text(5)) + ",lc17=" + CNULL(ChgSQL(Text(9))) + ",lc18=" + CNULL(ChgSQL(Text(19))) + ",lc19=" + CNULL(ChgSQL(Text(20))) + _
        ",lc20=" + CNULL(Text(21)) + ",lc21=" + CNULL(ChgSQL(Text(18))) + ",lc22=" + CNULL(ChangeCustomerL(Text(10))) + ",lc23=" + CNULL(ChgSQL(Text(11))) + _
        ",lc24=" + CNULL(Text(16)) + ",lc25=" + CNULL(Text(15)) + ",lc26=" + CNULL(ChangeCustomerL(Text(14))) + ",lc27=" + CNULL(ChgSQL(Text(22))) + _
        ",lc35=" + CNULL(ChangeCustomerL(Text(24))) + _
        ",lc43=" + CNULL(ChangeCustomerL(Text(23))) + ",lc44=" + CNULL(ChangeCustomerL(Text(25))) + ",lc45=" + CNULL(ChangeCustomerL(Text(26))) + _
        ",lc46=" + CNULL(ChangeCustomerL(Text(27))) + ",lc47=" + CNULL(Text(28)) + ",lc48=" + CNULL(Text(29)) + _
        ",LC49=" + CNULL(Combo4.Text) + ",LC50=" + CNULL(IIf(Combo5.Text <> "", Combo5.ListIndex, "")) + _
        ",LC51=" + CNULL(ChgSQL(Text(30))) + ",LC52=" + CNULL(ChgSQL(Text(52))) & _
        " where " & ChgLawcase(LcTmp) + " "
   
   On Error GoTo ErrHand
   cnnConnection.BeginTrans
   '紀錄分析語法
   Pub_SeekTbLog strExc(1)
   cnnConnection.Execute "begin user_data.user_enabled:=1; " & strExc(1) & "; end;"
   cnnConnection.CommitTrans
   
   'Added by Lydia 2021/01/14 (10/5) 若案件性質或案件屬性有改時Email通知秀玲提醒確認案源及金額是否需調整。案件屬性第1次設定時要與接洽單檔比較是否不同。
   If m_LOS01 <> "" Then
       strExc(0) = "": strExc(1) = ""
       If Text(28).Visible = True And Text(28).Enabled = True Then '判斷可維護才檢查
            'Modified by Lydia 2021/09/09 不用與接洽單檔比較; ex. L-006229-1-00先是在基本檔維護拿掉案件屬性有發通知, 又在分案設定案件屬性因為與接洽單一致所以沒有發通知
            'If Text(28).Tag = "" Then  '與接洽單檔比較
            '    If PUB_ChkTwoStrLst(m_CRL84, Text(28).Text) = False Then
            '        strExc(1) = strExc(1) & "、案件屬性"
            '        strExc(0) = strExc(0) & vbCrLf & "原案件屬性：" & m_CRL84 & vbCrLf & "現案件屬性：" & Text(28).Text
            '    End If
            'ElseIf Text(28).Tag <> Text(28).Text Then
            If Text(28).Tag <> Text(28).Text Then
            'end 2021/09/09
                If PUB_ChkTwoStrLst(Text(28).Tag, Text(28).Text) = False Then
                    strExc(1) = strExc(1) & "、案件屬性"
                    'Modified by Lydia 2021/09/09
                    'strExc(0) = strExc(0) & vbCrLf & "原案件屬性：" & m_CRL84 & vbCrLf & "現案件屬性：" & Text(28).Text
                    strExc(0) = strExc(0) & vbCrLf & "原案件屬性：" & IIf(Trim(Text(28).Tag) <> "", Trim(Text(28).Tag), "(空白)") & vbCrLf & "現案件屬性：" & IIf(Trim(Text(28).Text) <> "", Trim(Text(28).Text), "(空白)")
                End If
            End If
       End If
       If strExc(0) <> "" Then
           '主旨
           'Modified by Lydia 2021/09/09 法務分案=>法務案件基本資料維護
           strExc(1) = "法務案件基本資料維護" & m_CP01 & "-" & m_CP02 & IIf(m_CP03 <> "0", "-" & m_CP03, "") & IIf(m_CP04 <> "00", "-" & m_CP04, "") & "，改變" & Mid(strExc(1), 2)
           '內文
           strExc(2) = "法律所案號：" & m_CP01 & "-" & m_CP02 & IIf(m_CP03 <> "0", "-" & m_CP03, "") & IIf(m_CP04 <> "00", "-" & m_CP04, "") & vbCrLf & _
                            "專業部案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & IIf(m_LOS01cp03 <> "0", "-" & m_LOS01cp03, "") & IIf(m_LOS01cp04 <> "00", "-" & m_LOS01cp04, "") & "(" & m_LOS01 & ")" & vbCrLf & _
                             strExc(0)
           strExc(2) = strExc(2) & vbCrLf & vbCrLf & "請確認案源及金額是否需調整。" 'Added by Lydia 2021/09/09 加提醒
           'Modified by Lydia 2023/02/01 改成系統特殊設定
           'PUB_SendMail strUserNum, "83002", "", strExc(1), strExc(2)
           PUB_SendMail strUserNum, Pub_GetSpecMan("程式管理人員"), "", strExc(1), strExc(2)
       End If
   End If
   'end 2021/01/14
   
   '紀錄修改案號
   pub_ModifyCaseNum = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04
   
   ShowCurrRecord m_CP01, m_CP02, m_CP03, m_CP04
   ModRecord = True

   Exit Function
ErrHand:
   MsgBox (Err.Description)
   cnnConnection.RollbackTrans
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strlc As String
   
   DelRecord = False
   
   m_CP01 = txtcp01.Text
   m_CP02 = txtcp02.Text
   If txtcp03.Text <> Empty Then
      m_CP03 = txtcp03.Text
   Else
      m_CP03 = "0"
   End If
   If txtcp04.Text <> Empty Then
      m_CP04 = txtcp04.Text
   Else
      m_CP04 = "00"
   End If
   
   'Add By Sindy 2010/7/1
   If ChkCaseCode("CP", m_CP01, m_CP02, m_CP03, m_CP04) = False Then Exit Function
   If ChkCaseCode("NP", m_CP01, m_CP02, m_CP03, m_CP04) = False Then Exit Function
   '2010/7/1 End
   
   strlc = txtcp01 & txtcp02 & txtcp03 & txtcp04
'   If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
      If OnDataDeleteRecord(0, m_CP01 & m_CP02 & m_CP03 & m_CP04) <> 0 Then
         GoTo EXITSUB
      End If
      'add by nickc 2006/06/08
      On Error GoTo oErr
      cnnConnection.BeginTrans
      strExc(1) = "delete lawcase where " & ChgLawcase(strlc)
      'add by nickc 2006/06/08
      Pub_SeekTbLog strExc(1)
      
      'edit by nickc 2006/06/08
      cnnConnection.Execute strExc(1)
      
      'Added by Lydia 2016/11/24 一併刪除各項指示
      strSql = "DELETE FROM INSTRUCTIONS WHERE ITS01=" & CNULL(Pub_GetITS01Type(m_CP01)) & " AND ITS02=" & CNULL(m_CP01 & m_CP02 & m_CP03 & m_CP04)
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      'end 2016/11/24
    
      cnnConnection.CommitTrans
'   End If
   
   DelRecord = True
   
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (m_CP01 = m_LastRow(0) And m_CP02 = m_LastRow(1) And m_CP03 = m_LastRow(2) And m_CP04 = m_LastRow(3)) Or (m_CP01 = m_FirstRow(0) And m_CP02 = m_FirstRow(1) And m_CP03 = m_FirstRow(2) And m_CP04 = m_FirstRow(3)) Then
      RefreshRange
   End If
   ShowCurrRecord m_CP01, m_CP02, m_CP03, m_CP04
   
EXITSUB:
   Exit Function
oErr:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

'Added by Lydia 2021/01/14 法律所案源收文：讀取法務案源檔
Private Sub ReadLOS()
Dim stSQL As String, intQ As Integer
Dim RsQ As ADODB.Recordset
 
   m_LOS01 = "": m_LOS01cp01 = "": m_LOS01cp02 = "": m_LOS01cp03 = "": m_LOS01cp04 = ""
   m_LOS02 = ""
   m_CRL84 = ""
   '曾經收過案源，排除A3類案源。
   'Modified by Lydia 2021/09/09 有收過案源就算; 拿掉AND X.LOS02<>'A3'
   stSQL = "SELECT X.LOS01,X.LOS02,X.LOS04,X.LOS06,X.LOS10,X.LOS15,CP01,CP02,CP03,CP04,NVL(X.LOS17,X.LOS18) CRLNO " & _
                "FROM LAWOFFICESOURCE X,CASEPROGRESS WHERE X.LOS06 IN (SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & txtcp01 & "' AND CP02='" & txtcp02 & "'  AND CP03='" & txtcp03 & "'  AND CP04='" & txtcp04 & "' ) " & _
                "AND X.LOS01=CP09(+) AND X.LOS07 IS NULL ORDER BY X.LOS12 "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      '案源總收文號
      m_LOS01 = "" & RsQ.Fields("los01")
      '案源總收文號之本所案號
      m_LOS01cp01 = "" & RsQ.Fields("cp01")
      m_LOS01cp02 = "" & RsQ.Fields("cp02")
      m_LOS01cp03 = "" & RsQ.Fields("cp03")
      m_LOS01cp04 = "" & RsQ.Fields("cp04")
      '(原)案源案件類型
      m_LOS02 = "" & RsQ.Fields("LOS02")
      '接洽記錄單-法務案件屬性
      If "" & RsQ.Fields("CRLNO") <> "" Then
           stSQL = "select crl84 from Consultrecordlist where crl01=" & CNULL(RsQ.Fields("CRLNO"))
           intQ = 1
           Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
           If intQ = 1 Then
                m_CRL84 = "" & RsQ.Fields("crl84")
           End If
      End If
   End If
   Set RsQ = Nothing
End Sub

'Added by Lydia 2023/03/14
Private Sub Check2_Click(Index As Integer)
   If Check2(Index).Value = 1 Then
      If InStr(Text(28).Text, Trim(Check2(Index).Caption)) = 0 Then
         If Text(28).Text = "" Then
            Text(28).Text = Trim(Check2(Index).Caption)
         Else
            Text(28).Text = Text(28).Text & "," & Trim(Check2(Index).Caption)
         End If
      End If
   Else
      '案件屬性=xx,xx,xx
      If Left(Text(28), Len(Trim(Check2(Index).Caption))) = Trim(Check2(Index).Caption) Then
         Text(28).Text = Replace(Text(28).Text, Trim(Check2(Index).Caption) & ",", "")
         Text(28).Text = Replace(Text(28).Text, Trim(Check2(Index).Caption), "")
      Else
         Text(28).Text = Replace(Text(28).Text, "," & Trim(Check2(Index).Caption), "")
      End If
   End If
End Sub


