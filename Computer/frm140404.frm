VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140404 
   BorderStyle     =   1  '單線固定
   Caption         =   "往來記錄維護"
   ClientHeight    =   6576
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   9228
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6576
   ScaleWidth      =   9228
   Begin VB.CheckBox chkCR09 
      Caption         =   "財務處告知有產生國外交際餐費"
      Height          =   408
      Left            =   7524
      TabIndex        =   20
      Top             =   2190
      Width           =   1596
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "搬檔"
      Height          =   375
      Left            =   7080
      TabIndex        =   49
      Top             =   3930
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCF 
      Height          =   300
      Index           =   6
      Left            =   4590
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4170
      Visible         =   0   'False
      Width           =   4560
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8370
      Top             =   3810
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      Height          =   255
      Left            =   8445
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   5490
      Width           =   735
   End
   Begin VB.CommandButton cmdAddAtt 
      Caption         =   "<- 新增"
      Height          =   285
      Left            =   8445
      TabIndex        =   21
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdRemAtt 
      Caption         =   "-> 移除"
      Height          =   255
      Left            =   8445
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6030
      Width           =   735
   End
   Begin VB.CommandButton cmdRemCont 
      Caption         =   "移除↓"
      Height          =   285
      Left            =   6660
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1890
      Width           =   735
   End
   Begin VB.CommandButton cmdRemSort 
      Caption         =   "移除↓"
      Height          =   285
      Left            =   6630
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAddSort 
      Caption         =   "新增↑"
      Height          =   285
      Left            =   6630
      TabIndex        =   9
      Top             =   3180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAddCont 
      Caption         =   "新增↑"
      Height          =   285
      Left            =   6660
      TabIndex        =   5
      Top             =   2190
      Width           =   735
   End
   Begin VB.TextBox txtCF 
      Height          =   300
      Index           =   2
      Left            =   60
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5670
      Visible         =   0   'False
      Width           =   1020
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8460
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7740
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
            Picture         =   "frm140404.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140404.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140404.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140404.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140404.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140404.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140404.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140404.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140404.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140404.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140404.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9228
      _ExtentX        =   16277
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
   Begin VB.Frame Frame1 
      Height          =   2565
      Left            =   8664
      TabIndex        =   37
      Top             =   1176
      Width           =   1635
      Begin VB.CommandButton cmdReply 
         Caption         =   "回覆記錄編號"
         Height          =   285
         Left            =   90
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1455
      End
      Begin VB.ListBox lstCR18 
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
         ItemData        =   "frm140404.frx":20F4
         Left            =   90
         List            =   "frm140404.frx":20FB
         MultiSelect     =   1  '簡易多重選取
         Sorted          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ListBox lstCR11 
         BackColor       =   &H8000000F&
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
         ItemData        =   "frm140404.frx":2108
         Left            =   90
         List            =   "frm140404.frx":210F
         MultiSelect     =   1  '簡易多重選取
         Sorted          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   540
         Width           =   1455
      End
      Begin MSForms.TextBox txtCR 
         Height          =   300
         Index           =   18
         Left            =   855
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2190
         Visible         =   0   'False
         Width           =   720
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "1270;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "被回覆記錄編號："
         Height          =   180
         Index           =   9
         Left            =   90
         TabIndex        =   42
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2670
      TabIndex        =   44
      Top             =   3750
      Width           =   1875
      Begin VB.TextBox txtUserNo 
         Height          =   264
         Index           =   0
         Left            =   810
         MaxLength       =   6
         TabIndex        =   14
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "<- 新增"
         Height          =   285
         Index           =   0
         Left            =   45
         TabIndex        =   46
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "移除 ->"
         Height          =   285
         Index           =   0
         Left            =   45
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   420
         Width           =   735
      End
      Begin MSForms.Label lblName 
         Height          =   300
         Index           =   0
         Left            =   840
         TabIndex        =   52
         Top             =   420
         Width           =   960
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1693;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.ListBox lstAtt 
      Height          =   1050
      Left            =   1080
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5460
      Width           =   7320
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "12912;1852"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstUsers 
      Height          =   645
      Index           =   0
      Left            =   1080
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3810
      Width           =   1560
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2752;1138"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboPlace 
      Height          =   330
      Left            =   1080
      TabIndex        =   13
      Top             =   3450
      Width           =   5550
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "9790;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstSort 
      Height          =   360
      Left            =   -270
      TabIndex        =   10
      Top             =   3030
      Visible         =   0   'False
      Width           =   5550
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "9790;635"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboSort 
      Height          =   330
      Left            =   1080
      TabIndex        =   8
      Top             =   2520
      Width           =   5550
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "9790;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboContact 
      Height          =   330
      Left            =   1080
      TabIndex        =   4
      Top             =   2190
      Width           =   5550
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "9790;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstContact 
      Height          =   600
      Left            =   1080
      TabIndex        =   6
      Top             =   1590
      Width           =   5565
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "9816;1058"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   19
      Left            =   315
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4110
      Visible         =   0   'False
      Width           =   720
      VariousPropertyBits=   671105051
      MaxLength       =   70
      Size            =   "1270;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   1
      Top             =   990
      Width           =   1125
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1984;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   10
      Left            =   4680
      TabIndex        =   2
      Top             =   990
      Width           =   1125
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1984;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   4
      Left            =   315
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   720
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "1270;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   330
      Index           =   5
      Left            =   60
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   1020
      VariousPropertyBits=   671105051
      MaxLength       =   200
      Size            =   "1799;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   930
      Index           =   8
      Left            =   1080
      TabIndex        =   18
      Top             =   4500
      Width           =   8025
      VariousPropertyBits=   -1466941413
      MaxLength       =   4000
      ScrollBars      =   2
      Size            =   "14155;1640"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   7
      Left            =   645
      TabIndex        =   17
      Top             =   3450
      Visible         =   0   'False
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   180
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   570
      Index           =   6
      Left            =   1080
      TabIndex        =   12
      Top             =   2880
      Width           =   5550
      VariousPropertyBits=   -1466941413
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "9790;1005"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   1290
      Width           =   1092
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1926;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   1
      Left            =   1080
      TabIndex        =   0
      Top             =   690
      Width           =   1092
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1926;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Left            =   2220
      TabIndex        =   51
      Top             =   1290
      Width           =   5130
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "9049;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   2220
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   660
      Width           =   6585
      VariousPropertyBits=   -2147467233
      BackColor       =   16777215
      Size            =   "11615;529"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "主旨："
      Height          =   180
      Index           =   4
      Left            =   135
      TabIndex        =   28
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "接洽同仁："
      Height          =   180
      Index           =   10
      Left            =   135
      TabIndex        =   43
      Top             =   3900
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "往來日期：                           ( 西元 )"
      Height          =   180
      Index           =   13
      Left            =   135
      TabIndex        =   36
      Top             =   1035
      Width           =   2685
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "回覆期限：                           ( 西元 )"
      Height          =   180
      Index           =   8
      Left            =   3735
      TabIndex        =   35
      Top             =   1035
      Width           =   2685
   End
   Begin VB.Label Label1 
      Caption         =   "附件："
      Height          =   180
      Index           =   7
      Left            =   135
      TabIndex        =   33
      Top             =   5460
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "內容："
      Height          =   180
      Index           =   6
      Left            =   135
      TabIndex        =   30
      Top             =   4530
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "場合："
      Height          =   180
      Index           =   5
      Left            =   135
      TabIndex        =   29
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "往來類別："
      Height          =   180
      Index           =   3
      Left            =   135
      TabIndex        =   27
      Top             =   2610
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "聯絡人："
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   26
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "記錄編號："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   24
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "往來對象："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   25
      Top             =   1320
      Width           =   900
   End
End
Attribute VB_Name = "frm140404"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modified by Lydia 2022/01/11 改成Form 2.0; lstUsers(0)、lblName(0)、lstSort、lbl1、textCUID、txtCR(index)、cboContact、lstContact、lstAtt、cboPlace
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2007/11/29
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_bLanguage As String      '2008/12/10 ADD BY SONIA 加語文權限控制
Dim m_CR12 As String           '2008/12/10 ADD BY SONIA
Dim m_CR19 As String           '2019/7/15 ADD BY Sindy

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type

Dim m_FieldList() As FIELDITEM

Dim TF_CR As Integer
Dim strTmp As String
Dim oText As Control
Dim idx As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_SHOWDROPDOWN = &H14F
Dim iLanguage As Integer '1:中 2:英 3:日
'Modify By Sindy 2019/2/25 "CONTACTRECORD" 改為 "CONTACTFILE"
Private Const cTableName As String = "CONTACTFILE" 'Added by Lydia 2017/08/09 指定FTP資料夾名稱
Dim m_PCU51 As String 'Add By Sindy 2019/7/26
Dim varTmp As Variant
'Add By Sindy 2022/6/15
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_RDate As String
Dim m_PrevForm As Form
'2022/6/15 END
Dim intRepType As Integer 'Add by Amy 2025/03/20 取代及權限狀態
Dim m_strCRexcept As String 'Added by Lydia 2025/08/08

'Add By Sindy 2022/6/15
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cboContact_GotFocus()
   If cboContact.Locked = False Then
      CloseIme
      'Modified by Lydia 2022/01/11 改成Form 2.0 =>  自動下拉選單
      'SendMessage cboContact.hWnd, CB_SHOWDROPDOWN, 1, 0
      cboContact.DropDown
   End If
End Sub

Private Sub cboSort_Click()
   Dim iPos As Integer
   iPos = InStr(cboSort.Text, Chr(1))
   If iPos > 0 Then
      cboSort.Text = Left(cboSort.Text, iPos - 1)
   End If
   'Add By Sindy 2019/3/8
   'm_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If cboSort.Text <> "" Then
         varTmp = Split(cboSort.Text, " ")
         txtCR(5) = Trim(varTmp(0)) 'Left(Trim(cboSort.Text), 3)
      End If
   End If
   '2019/3/8 END
End Sub

Private Sub cboSort_GotFocus()
   If cboSort.Locked = False Then
      CloseIme
      'Modified by Lydia 2022/01/11 改成Form 2.0 =>  自動下拉選單
      'SendMessage cboSort.hWnd, CB_SHOWDROPDOWN, 1, 0
      cboSort.DropDown
   End If
End Sub

'新增接洽同仁
Private Sub cmdAdd_Click(Index As Integer)
   AddlstUsers Index
   txtCR(19) = ComposeListX(Index)
   txtUserNo(Index).SetFocus
End Sub
'Add by Morgan 2009/5/18
'開啟附件
Private Sub cmdOpenAtt_Click()
'Added by Lydia 2017/08/09
Dim tmpArr As Variant, ii As Integer
Dim stFileName As String
Dim hLocalFile As Long
'end 2017/08/09

   'Modified by Lydia 2022/01/11 改成Form 2.0元件
   'If lstAtt.Text = "" Then
   If lstAtt.ListIndex = -1 Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      'Added by Lydia 2017/08/09 判斷移檔日期
      If strSrvDate(1) >= CR_NewDate And txtCF(6).Text <> "" Then
         tmpArr = Empty
         tmpArr = Split(txtCF(6).Text, ",")
         ii = lstAtt.ListIndex
         If ii > UBound(tmpArr) Then Exit Sub
         If Trim(tmpArr(ii)) <> "" Then
            strExc(1) = Trim(Mid(lstAtt.Text, 1, InStrRev(lstAtt.Text, "(") - 1))
            'stFileName = App.path & "\$$" & strExc(1)
            'Modify By Sindy 2022/7/5 檢查是否已有完整路徑
            If InStr(strExc(1), App.path) > 0 Then
               ShellExecute hLocalFile, "open", strExc(1), vbNullString, vbNullString, 1
            Else
            '2022/7/5 END
               stFileName = App.path & "\$$" & strExc(1)
               If PUB_GetFtpFile(Trim(tmpArr(ii)), stFileName, cTableName) Then
                  ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
               End If
            End If
         End If
'      Else
'      'end 2017/08/09
'         PUB_OpenFtpFile txtCR(1), lstAtt.Text, Winsock1
      End If 'end 2017/08/09
   End If
End Sub
'移除接洽同仁
Private Sub cmdRemove_Click(Index As Integer)
   RemovelstUsers Index
   txtCR(19) = ComposeListX(Index)
   txtUserNo(Index).SetFocus
End Sub
'Modify by Morgan 2009/6/3 +可多選,+顯示檔案大小
Private Sub cmdAddAtt_Click()
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s
   Dim strMid As String, strList As String 'Added by Lydia 2017/08/09
   
On Error GoTo ErrHnd
   
   stFileName = "*.*"
   strList = txtCF(6).Text  'Added by Lydia 2017/08/09
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            For ii = 1 To UBound(sFile)
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               'Modified by Lydia 2017/08/09 存FTP檔名
               'AddListX lstAtt, stFileName & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)"
               strMid = PUB_GetNewFileNameSec(Mid(stFileName, InStrRev(stFileName, "\") + 1), , strList)
               AddListX lstAtt, PUB_StringFilter(stFileName) & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)", strMid
               'end 2017/08/09
            Next
         Else
            stFileName = .FileName
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            'Modified by Lydia 2017/08/09 存FTP檔名
            'AddListX lstAtt, stFileName & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)"
            strMid = PUB_GetNewFileNameSec(Mid(stFileName, InStrRev(stFileName, "\") + 1), , strList)
            AddListX lstAtt, PUB_StringFilter(stFileName) & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)", strMid
            'end 2017/08/09
         End If
         'Modify by Morgan 2009/5/19
         '改上傳到FTP,故只需留檔名
         'txtCF(2) = ComposeList(lstAtt)
         txtCF(2) = ComposeAttList(lstAtt)
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub cmdAddCont_Click()
   If AddList(lstContact, cboContact, 1) = True Then
      txtCR(4) = ComposeList(lstContact, 1)
      cboContact = ""
   End If
   If cboContact.Enabled = True Then cboContact.SetFocus
End Sub

Private Sub cmdAddSort_Click()
   If AddList(lstSort, cboSort) = True Then
      txtCR(5) = ComposeList(lstSort)
      cboSort = ""
   End If
   cboSort.SetFocus
End Sub

Private Sub cmdRemAtt_Click()
   'Add by Morgan 2009/5/19
   If InStr(lstAtt, "\") = 0 And Pub_StrUserSt03 <> "M51" Then
         MsgBox "已上傳檔案不可移除！"
   'end 2009/5/19
   'Modified by Lydia 2022/01/11
   'ElseIf RemoveList(lstAtt, 1) = True Then
   '   txtCF(2) = ComposeList(lstAtt)
   ElseIf RemoveList(lstAtt, 1) = True Then
      txtCF(2) = ComposeAttList(lstAtt)
   'end 2022/01/11
      cmdAddAtt.SetFocus
   End If
End Sub

Private Sub cmdRemCont_Click()
   'Modified by Lydia 2022/01/11 + pOpt=0
   If RemoveList(lstContact, 0) = True Then
      txtCR(4) = ComposeList(lstContact, 1)
      cboContact.SetFocus
   End If
End Sub

Private Sub cmdRemSort_Click()
   'Modified by Lydia 2022/01/11 + pOpt=0
   If RemoveList(lstSort, 0) = True Then
      txtCR(5) = ComposeList(lstSort)
      cboSort.SetFocus
   End If
End Sub

Private Sub cmdReply_Click()
   Dim stCon As String, rsReply As ADODB.Recordset
   If txtCR(3) <> "" Then
      If txtCR(1) = "" Then
         strExc(0) = "select '' C01,CR01 C02,CR02 C03,CR10 C04,CR06 C05,CR05 C06 from contactrecord where cr03='" & txtCR(3) & "'"
      Else
         strExc(0) = "select '' C01,CR01 C02,CR02 C03,CR10 C04,CR06 C05,CR05 C06 from contactrecord where cr03='" & txtCR(3) & "' and cr01<>'" & txtCR(1) & "' and (cr11 is null or instr(cr11,'" & txtCR(1) & "')=0)" & _
            " UNION select 'V',CR01,CR02,CR10,CR06,CR05 from contactrecord where cr03='" & txtCR(3) & "' and cr01<>'" & txtCR(1) & "' and instr(cr11,'" & txtCR(1) & "')>0"
      End If
      intI = 1
      Set rsReply = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Me.Tag = txtCR(18)
         Set frm140404_1.grdDataList.Recordset = rsReply
         Set frm140404_1.fmParent = Me
         frm140404_1.Show vbModal
         txtCR(18) = Me.Tag
         SetList lstCR18, txtCR(18)
      Else
         MsgBox "無可回覆之往來記錄！", vbExclamation
      End If
   End If
End Sub

Private Sub Form_Initialize()
   strExc(0) = "select * from ContactRecord where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   TF_CR = RsTemp.Fields.Count
   ReDim m_FieldList(TF_CR) As FIELDITEM
End Sub

Private Sub Form_Load()
Dim stFileName As String
Dim fs, f
Dim strMid As String

   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   '2008/12/10 ADD BY SONIA
   m_bLanguage = IsUserHasRightOfLanguage
   '有值才可查潛在客戶往來記錄 Y不限語文 J限日文 E限非日文

   MoveFormToCenter Me
   
   'Add by Morgan 2008/10/27 回覆功能先鎖住以後再用
   Frame1.Visible = False
   Label1(8).Visible = False
   txtCR(10).Visible = False
   
   textCUID.BackColor = &H8000000F
   AddCombo cboSort
   InitialField
   m_EditMode = 0
   ShowRecord -2
   SetInputEntry
   UpdateToolbarState
   
   'Added by Lydia 2017/08/09
   If Pub_StrUserSt03 <> "M51" Then cmd1.Visible = False
   
   'Add By Sindy 2022/6/15
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
      Call OnAction(vbKeyF2) '新增
      
      stFileName = m_PrevForm.m_strFullFileName
      If PUB_ChkDir(stFileName) = False Then
         '下載信件檔
         Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", , stFileName, True)
      End If
      Set fs = CreateObject("Scripting.FileSystemObject")
      Set f = fs.GetFile(stFileName)
      strMid = Mid(stFileName, InStrRev(stFileName, "\") + 1)
      'strMid = PUB_GetNewFileNameSec(Mid(strReName, InStrRev(strReName, "\") + 1), , "") '存FTP檔名
      AddListX lstAtt, PUB_StringFilter(stFileName) & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)", strMid
      txtCF(2) = strMid & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)"
      txtCF(6) = strMid
      cmdOpenAtt.Enabled = True
   End If
   '2022/6/15 END
   
   m_strCRexcept = Pub_GetCRExceptNo(Me.Name) 'Added by Lydia 2025/08/08
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   If Me.Visible = True Then
      Select Case m_EditMode
         Case 1 '新增
            txtCR(1).Locked = True
            txtCR(2).SetFocus
            
         Case 2 '修改
            txtCR(1).Locked = True
            txtCR(2).SetFocus
         
         Case 4 '查詢
            txtCR(1).Locked = False
            txtCR(1).SetFocus
            
         Case Else
            txtCR(1).Locked = True
            txtCR(1).SetFocus
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
         If m_bUpdate And txtCR(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtCR(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtCR(1) <> "" Then
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
   'Add by Amy 2025/03/20 有符號且無權限不可修改
   If intRepType = 2 And TBar1.Buttons(2).Enabled = True Then
      TBar1.Buttons(2).Enabled = False
   End If
   'end 2025/03/20
End Sub

'Modified by Lydia 2022/01/11 ComboBox=> Control
Private Sub AddCombo(oCombo As Control)
   With oCombo
      .Clear
      .Tag = "" 'Added by Lydia 2022/01/11
      'Modified by Morgan 2018/3/23 調整內容及順序 --David
      '.AddItem "IP諮詢"
      '.AddItem "非IP法律諮詢"
      '.AddItem "詢價"
      '.AddItem "申請所需文件"
      '.AddItem "利益衝突"
      '.AddItem "互惠"
      '.AddItem "詢問IP侵害"
      '.AddItem "訪談" & Chr(1) & "(包括來所訪問、出國拜訪、國際會議)"
      '.AddItem "客戶特別指示" & Chr(1) & "(譬如:不要寄confirmation copy, 聯絡方式只限fax or e-mail, 付款方式只限credit card or cheque…)"
      '.AddItem "報價"   'add by sonia 2018/1/12
      '.AddItem "慰問"   'add by sonia 2018/2/9
      
      'Modify By Sindy 2019/2/26 Mark
'      .AddItem "總指示"
'      .AddItem "來函詢價/申請文件"
'      .AddItem "利益衝突調查"
'      .AddItem "合作契約"
'      .AddItem "催款"
'      .AddItem "投標/固定價"
'      .AddItem "網路平台/使用費"
'      .AddItem "問卷"
'      .AddItem "互惠"
'      .AddItem "IP諮詢"
'      .AddItem "訪談" & Chr(1) & "(包括來所訪問、出國拜訪、國際會議)"
'      .AddItem "代理人提供報價"
'      .AddItem "慰問"
'      'end 2018/3/23
      'Add By Sindy 2019/2/26
      strExc(0) = "select ac02,ac03 from allcode where ac01='11' order by ac01 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            .AddItem RsTemp.Fields("ac02") & " " & RsTemp.Fields("ac03")
            .Tag = RsTemp.Fields("ac02") & "," & .Tag 'Added by Lydia 2022/01/11
            RsTemp.MoveNext
         Loop
      End If
      '2019/2/26 END
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHand
   'Add by Morgan 2009/5/19
   '清除暫存檔
   If Dir(App.path & "\$$*.*") <> "" Then
      Kill App.path & "\$$*.*"
   End If
   'end 2009/5/19
   
   'Add By Sindy 2022/6/15
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2022/6/15 END
   
   PUB_SendMailCache 'Added by Lydia 2023/02/02
   
   Set frm140404 = Nothing
   Exit Sub
   
ErrHand:
   MsgBox Err.Number & " : Kill App.path & \$$*.*" & vbCrLf & Err.Description & vbCrLf & vbCrLf & "應該是有信件或暫存檔案，開啟中", vbInformation
   Set frm140404 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   For Each oText In txtCR
      idx = oText.Index
      m_FieldList(idx).fiName = "CR" & Format(idx, "00")
   Next
End Sub

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
   Dim adoRst As New ADODB.Recordset
   Dim stCR08Rep As String 'Add by Amy 2025/03/20
      
Top:
   intRepType = 0 'Add by Amy 2025/03/20
   '2008/12/10 MODIFY BY SONIA 因做語文控制,故此處txtCR(1)改為txtCR(1).Tag,否則最後筆或第一筆時會因權限控制而改txtCR(1)值
   Select Case p_iWay
      Case 0
         
         strExc(0) = "SELECT * FROM ContactRecord" & _
            " WHERE CR01 = '" & txtCR(1).Tag & "'"
      Case -2
         'Modified by Morgan 2019/2/14 有點慢改語法
         'strExc(0) = "SELECT * FROM ContactRecord order by CR01 ASC"
         strExc(0) = "SELECT * FROM ContactRecord a where CR01=(select min(b.cr01) from ContactRecord b)"
      Case -1
         'Modified by Morgan 2019/2/14 有點慢改語法
         'strExc(0) = "SELECT * FROM ContactRecord WHERE CR01 <'" & txtCR(1).Tag & "' order by CR01 DESC"
         strExc(0) = "SELECT * FROM ContactRecord a where CR01=(select max(b.cr01) from ContactRecord b where b.cr01<'" & txtCR(1).Tag & "')"
         
      Case 1
         'Modified by Morgan 2019/2/14 有點慢改語法
         'strExc(0) = "SELECT * FROM ContactRecord WHERE CR01 >'" & txtCR(1).Tag & "' order by CR01 ASC"
         strExc(0) = "SELECT * FROM ContactRecord a where CR01=(select min(b.cr01) from ContactRecord b where b.cr01>'" & txtCR(1).Tag & "')"
      Case 2
         'Modified by Morgan 2019/2/14 有點慢改語法
         'strExc(0) = "SELECT * FROM ContactRecord order by CR01 DESC"
         strExc(0) = "SELECT * FROM ContactRecord a where CR01=(select max(b.cr01) from ContactRecord b)"
   End Select
   '2008/12/10 END
      
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modify by Amy 2025/03/20 +內容中若特殊符號前後框住,則依權限顯示 or 不顯示
      'ex: KA2000178/Y2005000 其客戶「#(Nabtesco)#」相關商標案件
      '      [有]權限(建立者、接洽人員及其最高主管)顯示:其客戶「Nabtesco」相關商標案件
      '      [無]權限顯示:其客戶***相關商標案件
      stCR08Rep = ChkLimitAndReplace(Me.Name, "" & adoRst.Fields("CR08"), "" & adoRst.Fields("CR19"), intRepType, "" & adoRst.Fields("CR12"))
      If intRepType <> 2 Then
         '無符號 or 有特殊符號且有權限 ->顯示原始資料(含符號)
         stCR08Rep = ""
      End If
      'end 2025/03/20
      '2008/12/10 ADD BY SONIA 加語文權限控制,無權限繼續讀下一筆
      'If GetCustData(adoRst.Fields("CR03")) = False Then
      'Modify By Sindy 2009/04/30
      If Left(Trim(adoRst.Fields("CR03")), 1) = "R" Then
         'Modify By Sindy 2019/7/26
         'If PUB_CheckModifyLimit_frm140402(adoRst.Fields("CR12"), "A") = False Then
         'Modify By Sindy 2024/10/1 傳入建檔人
         If PUB_CheckModifyLimit_frm140402(m_PCU51, adoRst.Fields("CR12")) = False Then
         '2019/7/26 END
            txtCR(1).Tag = adoRst.Fields("CR01")
            Set adoRst = Nothing
            If p_iWay = -2 Then p_iWay = 1
            If p_iWay = 2 Then p_iWay = -1
            If p_iWay = 0 Then
               'MsgBox "您沒有維護此筆潛在客戶的往來記錄權限！", vbInformation
               p_iWay = 1
            End If
            GoTo Top
         End If
      End If
      '2008/12/10 END
      'Added by Lydia 2025/08/08 國外往來記錄的維護及查詢限制
      If InStr(m_strCRexcept, adoRst.Fields("CR01")) > 0 Then
          'MsgBox "限閱往來記錄！", vbInformation '秀玲:維護不用彈訊息
          txtCR(1).Tag = adoRst.Fields("CR01")
          Set adoRst = Nothing
          If p_iWay = -2 Then p_iWay = 1
          If p_iWay = 2 Then p_iWay = -1
          If p_iWay = 0 Then
             p_iWay = 1
          End If
          GoTo Top
      End If
      'end 2025/08/08
      
      'Modify by Amy 2025/03/20 +stCR08Rep及有符號且無權限不可修改
      UpdateCtrlData adoRst, stCR08Rep
      '有符號且無權限不可修改
      If intRepType = 2 Then
         TBar1.Buttons(2).Enabled = False
      Else
         TBar1.Buttons(2).Enabled = True
      End If
      'end 2025/03/20
      ShowRecord = True
   Else
      If p_iWay = 0 Then
         ClearField
         MsgBox "查無資料！", vbInformation
      ElseIf p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
         '2008/12/10 ADD BY SONIA
         Set adoRst = Nothing
         txtCR(1).Tag = txtCR(1)
         p_iWay = 0
         GoTo Top
         '2008/12/10 END
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
         '2008/12/10 ADD BY SONIA
         Set adoRst = Nothing
         txtCR(1).Tag = txtCR(1)
         p_iWay = 0
         GoTo Top
         '2008/12/10 END
      Else
         ClearField
         MsgBox "查無資料！", vbInformation
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtCR(1).SetFocus
      txtCR_GotFocus 1
   End If
End Function

' 將資料庫中的資料更新到所有欄位中
'Modify by Amy 2025/03/20 +stCR08Rep (有值不能修改)
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset, ByVal stCR08Rep As String)
Dim CUID(1 To 6) As String
Dim AdoRs As New ADODB.Recordset 'Add By Sindy 2019/2/25
   
   ClearField
   With p_Rst
      If .RecordCount > 0 Then
         For Each oText In txtCR
            idx = oText.Index
            'Modify by Amy 2025/03/20 +if
            If idx = 8 And stCR08Rep <> MsgText(601) Then
               m_FieldList(idx).fiOldData = stCR08Rep
            Else
               m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
            End If
            'end 2025/03/20
            m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
            'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
            'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
               m_FieldList(idx).fiType = 0
            'Else
            '   m_FieldList(idx).fiType = 1
            'End If
            'end 2017/06/29
            oText.Text = m_FieldList(idx).fiOldData
            oText.Tag = oText.Text 'Added by Lydia 2023/02/02
         Next
         CUID(1) = "" & .Fields("CR12")
         m_CR12 = "" & .Fields("CR12")   '2008/12/10 ADD BY SONIA
         m_CR19 = "" & .Fields("CR19")   '2019/7/15 ADD BY Sindy
         CUID(2) = "" & .Fields("CR13")
         CUID(3) = "" & .Fields("CR14")
         CUID(4) = "" & .Fields("CR15")
         CUID(5) = "" & .Fields("CR16")
         CUID(6) = "" & .Fields("CR17")
         txtCR_Validate 3, False
         
         'Modify By Sindy 2023/8/10
         If "" & .Fields("CR09") = "Y" Then
            chkCR09.Value = 1
            chkCR09.Tag = "Y"
         Else
            chkCR09.Value = 0
            chkCR09.Tag = ""
         End If
         '2023/8/10 END
         
         'Add By Sindy 2019/7/26 國內外權限
         strExc(0) = "SELECT pcu51 FROM potcustomer where pcu01='" & Left(txtCR(3), 8) & "' and pcu02='0'"
         intI = 1
         Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
         m_PCU51 = ""
         If intI = 1 Then
            m_PCU51 = "" & AdoRs.Fields("pcu51")
         End If
         '2019/7/26 END
         
         'Modify By Sindy 2019/2/25
         'SetList lstSort, txtCR(5)
         strExc(0) = "SELECT AC01,AC02,AC03 FROM allcode where AC01='11' and AC02='" & txtCR(5) & "'"
         intI = 1
         Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            cboSort.Text = AdoRs.Fields("ac02") & " " & AdoRs.Fields("ac03")
            cboSort.Tag = cboSort.Text 'Added by Lydia 2023/02/02
         End If
                  
         SetList lstCR11, "" & .Fields("CR11")
         SetList lstCR18, txtCR(18)
         txtCR(18).Tag = txtCR(18)
         'Add by Morgan 2009/2/6
         SetCboPlace txtCR(7)
         SetlstUsers 0, txtCR(19)
         'END 2009/2/6
         'Add By Sindy 2019/2/25
         strExc(0) = "SELECT cf02,cf06,cf07 FROM ContactFile where CF01='" & p_Rst.Fields("cr01") & "'"
         intI = 1
         Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            AdoRs.MoveFirst
            Do While Not AdoRs.EOF
               txtCF(2) = txtCF(2) & "," & AdoRs.Fields("cf02") & IIf("" & AdoRs.Fields("cf07") <> "", " (" & AdoRs.Fields("cf07") & " KB)", "")
               txtCF(6) = txtCF(6) & "," & AdoRs.Fields("cf06")
               AdoRs.MoveNext
            Loop
            txtCF(2) = Mid(txtCF(2), 2)
            txtCF(6) = Mid(txtCF(6), 2)
         Else
            txtCF(2) = ""
            txtCF(6) = ""
         End If
         '2019/2/25 END
         SetList lstAtt, txtCF(2)
         txtCF(6).Tag = txtCF(6).Text 'Added by Lydia 2017/08/09
      End If
   End With
   UpdateCUID CUID, textCUID
   txtCR(1).Tag = txtCR(1)
   
   Set AdoRs = Nothing 'Add By Sindy 2019/2/25
End Sub

Private Sub ClearField()
   Dim oLabel As Label
   For Each oText In txtCR
      oText.Text = Empty
      'Modified by Lydia 2020/11/03 排除流水號
      'oText.Tag = "" 'Added by Lydia 2019/09/10
      If oText.Index <> 1 Then oText.Tag = ""
   Next
   lbl1 = Empty
   
   If m_EditMode = 1 Then
      '新增時開發日期預設當天
      txtCR(2) = strSrvDate(1)
   End If
   For intI = 1 To TF_CR
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   cboContact.Clear
   lstContact.Clear
   lstSort.Clear
   lstAtt.Clear
   'Added by Lydia 2022/01/11
   cboContact.Tag = ""
   lstContact.Tag = ""
   lstSort.Tag = ""
   lstAtt.Tag = ""
   'end 2022/01/11
   lstCR11.Clear
   lstCR18.Clear
   cboPlace.Clear
   cboPlace.Tag = "" 'Added by Lydia 2022/01/11
   txtUserNo(0) = ""
   lblName(0) = ""
   lstUsers(0).Clear
   lstUsers(0).Tag = "" 'Added by Lydia 2022/01/11
   'Add By Sindy 2019/2/26
   For Each oText In txtCF
      oText.Text = Empty
      oText.Tag = "" 'Add By Sindy 2020/5/18
   Next
   '2019/2/26 END
   cboSort.ListIndex = -1 'Add By Sindy 2019/3/8
   cboSort.Tag = "" 'Added by Lydia 2023/02/02
   chkCR09.Value = 0 'Add By Sindy 2023/8/10
End Sub

'Modified by Lydia 2022/01/11 ListBox=>Control
Private Sub SetList(oList As Control, p_stList As String)
   Dim arrID
   oList.Clear
   oList.Tag = "" 'Added by Lydia 2022/01/11
   If p_stList <> "" Then
      arrID = Split(p_stList, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

'Modified by Lydia 2022/01/11
'Private Sub setContact(oCombo As ComboBox, oList As ListBox, Optional p_stList As String)
Private Sub setContact(oCombo As Control, oList As Control, Optional p_stList As String)
   Dim arrID
   Dim stPCC01 As String
   stPCC01 = Left(txtCR(3), 8)
   
   oCombo.Clear
   oCombo.Tag = "" 'Added by Lydia 2022/01/11
   Select Case iLanguage
      Case 1 '中 -> 英 -> 日
         strExc(0) = "select pcc02 c1,nvl(pcc05,nvl(pcc03,pcc04)) c2 from potcustcont where pcc01='" & stPCC01 & "' order by 1 desc"
      
      Case 3 '日 -> 英 -> 中
         strExc(0) = "select pcc02 c1,nvl(pcc04,nvl(pcc03,pcc05)) c2 from potcustcont where pcc01='" & stPCC01 & "' order by 1 desc"
         
      Case Else '英 -> 日 -> 中
         strExc(0) = "select pcc02 c1,nvl(pcc03,nvl(pcc04,pcc05)) c2 from potcustcont where pcc01='" & stPCC01 & "' order by 1 desc"
   End Select
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      '設定聯絡人選單
      .MoveFirst
      Do While Not .EOF
         oCombo.AddItem "" & .Fields(1), 0
         'Modified by Lydia 2022/01/11 改成Form 2.0沒有ItemData屬性
         'oCombo.ItemData(0) = .Fields(0)
         oCombo.Tag = .Fields(0) & "," & oCombo.Tag
         .MoveNext
      Loop
      '設定聯絡人清單
      If p_stList <> "" Then
         oList.Clear
         oList.Tag = "" 'Added by Lydia 2022/01/11
         arrID = Split(p_stList, ",")
         '照原順序排
         For intI = UBound(arrID) To LBound(arrID) Step -1
            .MoveFirst
            Do While Not .EOF
               If .Fields("C1") = arrID(intI) Then
                  oList.AddItem "" & .Fields(1), 0
                  'Modified by Lydia 2022/01/11 改成Form 2.0沒有ItemData屬性
                  'oList.ItemData(0) = .Fields(0)
                  oList.Tag = .Fields(0) & "," & oList.Tag
                  Exit Do
               End If
               .MoveNext
            Loop
         Next
      End If
      End With
   End If
End Sub

'Modified by Lydia 2022/01/11 ListBox=>Control
Private Function ComposeList(oList As Control, Optional p_iOpt As Integer = 0) As String
'Modified by Lydia 2022/01/11 改成Form 2.0
'   Dim iPos As Integer, stItem As String
'   strExc(1) = ""
'   If oList.ListCount > 0 Then
'      For intI = 0 To oList.ListCount - 1
'         If p_iOpt = 0 Then
'            iPos = InStr(oList.List(intI), Chr(1))
'            If iPos > 0 Then
'               stItem = Left(oList.List(intI), iPos - 1)
'            Else
'               stItem = oList.List(intI)
'            End If
'         Else
'            stItem = Format(oList.ItemData(intI), "00")
'         End If
'         stItem = GetFileName(stItem) 'Add By Sindy 2012/3/21
'         If intI = 0 Then
'            strExc(1) = stItem
'         Else
'            strExc(1) = strExc(1) & "," & stItem
'         End If
'      Next
'   End If
'   ComposeList = strExc(1)
   ComposeList = oList.Tag
'end 2022/01/11
End Function

'Private Function GetCustData(p_stCust As String) As Boolean
'   Dim aiOrder(1 To 3) As Integer
'   '2008/12/10 modify by sonia 加國籍才能判斷語文權限
'   Select Case Left(p_stCust, 1)
'      Case "X"
'         strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3 from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "'"
'      Case "Y"
'         strExc(0) = "select fa31,fa04,rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) fa05,fa06,FA10 N3 from fagent where fa01='" & Left(p_stCust, 8) & "' and fa02='" & Right(p_stCust, 1) & "'"
'      Case "R"
'         strExc(0) = "select pcu36,pcu08,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06) pcu03,pcu07,PCU09 N3 from potcustomer where pcu01='" & Left(p_stCust, 8) & "' and pcu02='" & Right(p_stCust, 1) & "'"
'      Case Else
'         MsgBox "往來對象必須為 X、Y 或 R 開頭", vbCritical + vbOKOnly, "檢核資料"
'         Exit Function
'   End Select
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   lbl1 = ""
'   If intI = 1 Then
'      '2008/12/10 ADD BY SONIA 加語文權限
'      If m_EditMode = "1" Or m_EditMode = "2" Then
'         If m_bLanguage = "" And Left(p_stCust, 1) = "R" Then
'            MsgBox "您沒有維護潛在客戶的往來記錄權限 !!!", vbInformation
'            Exit Function
'         ElseIf m_bLanguage = "J" And Left(p_stCust, 1) = "R" And Mid(RsTemp.Fields("N3"), 1, 3) <> "011" Then
'            MsgBox "您沒有維護英文組潛在客戶的往來記錄權限 !!!", vbInformation
'            Exit Function
'         ElseIf m_bLanguage = "E" And Left(p_stCust, 1) = "R" And Mid(RsTemp.Fields("N3"), 1, 3) = "011" Then
'            MsgBox "您沒有維護日文組潛在客戶的往來記錄權限 !!!", vbInformation
'            Exit Function
'         End If
'      Else
'         If m_bLanguage = "" And Left(p_stCust, 1) = "R" Then
'            Exit Function
'         ElseIf m_bLanguage = "J" And Left(p_stCust, 1) = "R" And Mid(RsTemp.Fields("N3"), 1, 3) <> "011" Then
'            Exit Function
'         ElseIf m_bLanguage = "E" And Left(p_stCust, 1) = "R" And Mid(RsTemp.Fields("N3"), 1, 3) = "011" Then
'            Exit Function
'         End If
'      End If
'      '2008/12/10 END
'
'      iLanguage = Val("" & RsTemp(0))
'      Select Case iLanguage
'         Case 1 '中 -> 英 -> 日
'            aiOrder(1) = 1
'            aiOrder(2) = 2
'            aiOrder(3) = 3
'
'         Case 3 '日 -> 中 -> 英
'            aiOrder(1) = 3
'            aiOrder(2) = 1
'            aiOrder(3) = 2
'
'         Case Else '英 -> 中 -> 日
'            aiOrder(1) = 2
'            aiOrder(2) = 1
'            aiOrder(3) = 3
'      End Select
'      For intI = 1 To 3
'         If Not IsNull(RsTemp(aiOrder(intI))) Then
'            lbl1 = RsTemp(aiOrder(intI))
'            Exit For
'         End If
'      Next
'      GetCustData = True
'   '2008/12/10 ADD BY SONIA
'   Else
'      MsgBox "往來對象輸入錯誤！"
'   '2008/12/10 END
'   End If
'End Function
'Private Function GetCustData(p_stCust As String) As Boolean
'Dim strName As String
'
'   GetCustData = False
'
'   Select Case Left(p_stCust, 1)
'      Case "X"
'         strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3,CU81 from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "'"
'      Case "Y"
'         strExc(0) = "select fa31,fa04,rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) fa05,fa06,FA10 N3,FA46 from fagent where fa01='" & Left(p_stCust, 8) & "' and fa02='" & Right(p_stCust, 1) & "'"
'      Case "R"
'         strExc(0) = "select pcu36,pcu08,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06) pcu03,pcu07,PCU09 N3,PCU41 from potcustomer where pcu01='" & Left(p_stCust, 8) & "' and pcu02='" & Right(p_stCust, 1) & "'"
'      Case Else
'         MsgBox "往來對象必須為 X、Y 或 R 開頭", vbCritical + vbOKOnly, "檢核資料"
'         Exit Function
'   End Select
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   lbl1 = ""
'   If intI = 1 Then
'      For intI = 1 To 3
'         If Not IsNull(RsTemp(intI)) Then
'            strName = RsTemp(intI)
'            Exit For
'         End If
'      Next
'
'      '依LoginUser和輸入人員之部門第一碼判斷部門權限, 相同者才可輸入查詢
'      '但M51不受限制
'      strExc(0) = "SELECT A.ST03,B.ST03 FROM STAFF A,STAFF B " & _
'                         "WHERE A.ST01 = '" & strUserNum & "' " & _
'                              "AND B.ST01 = '" & Trim(RsTemp(5)) & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If Trim(RsTemp(0)) <> "M51" And _
'            Left(Trim(RsTemp(0)), 1) <> Left(Trim(RsTemp(1)), 1) Then
'            MsgBox "您沒有維護此筆潛在客戶的往來記錄權限！"
'            Exit Function
'         End If
'      End If
'   Else
'      MsgBox "往來對象輸入錯誤！"
'      Exit Function
'   End If
'   lbl1 = strName
'
'   GetCustData = True
'End Function
'Add by Morgan 2009/5/20
'Modified by Lydia 2022/01/11
'Private Sub lstAtt_DblClick()
Private Sub lstAtt_DblClick(Cancel As MSForms.ReturnBoolean)
   If cmdOpenAtt.Enabled = True Then
      cmdOpenAtt.Value = True
   End If
End Sub

Private Sub txtCR_Change(Index As Integer)
   If Index = 3 Then
      If txtCR(3) <> txtCR(3).Tag Then
         cboContact.Clear
         lstContact.Clear
         'Added by Lydia 2022/01/11
         cboContact.Tag = ""
         lstContact.Tag = ""
         'end 2022/01/11
      End If
      txtCR(3).Tag = txtCR(3).Text
   End If
End Sub

Private Sub txtCR_GotFocus(Index As Integer)
   Select Case Index
      Case 6, 7, 8
         OpenIme
         
      Case Else
         CloseIme
         
   End Select
   TextInverse txtCR(Index)
End Sub

'Modified by Lydia 2022/01/11 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txtCR_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCR_Validate(Index As Integer, Cancel As Boolean)
Dim AdoRs As New ADODB.Recordset 'Add By Sindy 2024/10/1
Dim iLen As Integer
Dim strName As String
   
   Select Case Index
      Case 3
         If txtCR(Index) <> "" Then
            'Modify By Sindy 2021/3/25 平台編號會是4位的流水號
            If Len(txtCR(Index)) > 5 Or (Len(txtCR(Index)) = 4 And IsNumeric(txtCR(Index)) = True) Then
               'Add By Sindy 2021/3/25
               If Left(txtCR(Index), 1) = "X" Or Left(txtCR(Index), 1) = "Y" Or Left(txtCR(Index), 1) = "R" Then
               '2021/3/25 END
                  txtCR(Index) = Left(txtCR(Index) & "000", 9)
               End If
               lbl1 = ""
               If PUB_GetCustData(txtCR(Index), strName) = False Then
                  '2008/12/10 MODIFY BY SONIA
                  'Cancel = True
                  'MsgBox "往來對象輸入錯誤！"
                  'txtCR_GotFocus Index
                  If m_EditMode = "1" Or m_EditMode = "2" Then
                     txtCR(Index) = "" 'Add By Sindy 2024/10/9 不清除欄位值,若無權限時,會一直重覆彈訊息
                     Cancel = True
                     txtCR_GotFocus Index
                  End If
                  '2008/12/10 END
               Else
                  lbl1 = strName
                  'Add By Sindy 2021/3/25
                  If Len(txtCR(Index)) = 4 And IsNumeric(txtCR(Index)) = True Then
                     cboContact.Enabled = False
                  Else
                     cboContact.Enabled = True
                  '2021/3/25 END
                     setContact cboContact, lstContact, txtCR(4)
                  End If
               End If
            Else
               Cancel = True
               MsgBox "往來對象編號請至少輸入六碼，或是4碼平台編號。", vbCritical + vbOKOnly, "檢核資料"
               txtCR_GotFocus Index
            End If
            'Add By Sindy 2024/10/1 國內外權限
            If m_EditMode = 1 Then
               strExc(0) = "SELECT pcu51 FROM potcustomer where pcu01='" & Left(txtCR(3), 8) & "' and pcu02='0'"
               intI = 1
               Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
               m_PCU51 = ""
               If intI = 1 Then
                  m_PCU51 = "" & AdoRs.Fields("pcu51")
               End If
            End If
            '2024/10/1 END
         End If
         
      Case 2, 10
         If txtCR(Index) <> "" Then
            If CheckIsDate(txtCR(Index)) = False Then
               txtCR_GotFocus Index
               Cancel = True
            End If
         End If
   End Select
   
   If Cancel = False Then
      If txtCR(Index).MaxLength > 0 Then
         Select Case Index
            '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
            Case 6, 7, 8
               iLen = txtCR(Index).MaxLength - 1
            Case Else
               iLen = txtCR(Index).MaxLength
         End Select
         If Not CheckLengthIsOK(txtCR(Index), iLen) Then
            Cancel = True
         End If
      End If
   End If
   
   Set AdoRs = Nothing 'Add By Sindy 2024/10/1
End Sub

'Modified by Lydia 2017/08/09 +存FTP檔名 stFtpName
'Modified by Lydia 2022/01/11  ListBox => Control
Private Function AddListX(oList As Control, stNewItem As String, stFtpName As String) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String
   If InStr(stNewItem, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請重新命名！", vbExclamation
      cmdAddAtt.SetFocus
      Exit Function
   End If
   
   If stNewItem <> "" Then
      For idx = 0 To oList.ListCount - 1
         stFileName = GetFileName(oList.List(idx))
         If GetFileName(stNewItem) = stFileName Then
            MsgBox "附件[" & stFileName & "]已存在！"
            AddListX = False
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         oList.AddItem stNewItem, 0
         AddListX = True
         
         'Added by Lydia 2017/08/09 存FTP檔名 (堆疊)
         'Modified by Lydia 2022/01/11
         'txtCF(6) = stFtpName & IIf(txtCF(6) <> "", ",", "") & txtCF(6)
         txtCF(6) = stFtpName & "," & txtCF(6)
      End If
   End If
End Function
'Modified by Lydia 2022/01/11
'Private Function AddList(oList As ListBox, oCombo As ComboBox, Optional p_iOpt As Integer = 0) As Boolean
Private Function AddList(oList As Control, oCombo As Control, Optional p_iOpt As Integer = 0) As Boolean
   Dim idx As Integer, bFound As Boolean, stNewItem As String
   Dim stSort As String, iPos As Integer
   Dim iNewItemData As String 'Modified by Lydia 2022/01/11 Integer 改成String
   
   If oCombo.Text = "" Then
      Exit Function
   End If
   
   'Modified by Lydia 2022/01/11 0=>""
   iNewItemData = ""
   If p_iOpt = 1 Then
      If oCombo.ListIndex = -1 Then
         MsgBox "聯絡人資料不存在！"
         Exit Function
      Else
         'Modify by Morgan 2008/12/9 從下面移上來,因為聯絡人才是存代碼
         'Modified by Lydia 2022/01/11 改成Form 2.0
         'iNewItemData = oCombo.ItemData(oCombo.ListIndex)
         iNewItemData = PUB_GetItemData(oCombo.Tag, oCombo.ListIndex)
      End If
   End If
   
   '若有控制字元時後面為說明文字不抓
   iPos = InStr(oCombo, Chr(1))
   If iPos > 0 Then
      stNewItem = Left(oCombo, iPos - 1)
   Else
      stNewItem = oCombo
   End If
      
   '2008/11/10 modify by sonia 原用逗號,改用分號;因為聯絡人名有逗號,R00020000
   If InStr(stNewItem, ";") > 0 Then
      MsgBox "分號[;]為系統保留字，請改用其他符號！", vbExclamation
      oCombo.SetFocus
      Exit Function
   End If

   If stNewItem <> "" Then
      'Modified by Lydia 2022/01/11 改成Form 2.0
'      For idx = 0 To oList.ListCount - 1
'         If oList.List(idx) = stNewItem And oList.ItemData(idx) = iNewItemData Then
'            MsgBox "資料已存在！"
'            AddList = False
'            bFound = True
'            Exit For
'         End If
'      Next
         If InStr(oList.Tag, iNewItemData) > 0 Then
            MsgBox "資料已存在！"
            AddList = False
            bFound = True
         End If
         'end 2022/01/11
      If bFound = False Then
         oList.AddItem stNewItem, 0
         'Modified by Lydia 2022/01/11 改成Form 2.0
         'If p_iOpt <> 0 Then
         '   oList.ItemData(0) = oCombo.ItemData(oCombo.ListIndex)
         'End If
         oList.Tag = iNewItemData & "," & oList.Tag
         AddList = True
      End If
   End If
End Function

'Modified by Lydia 2022/01/11
'Private Function RemoveList(oList As ListBox) As Boolean
Private Function RemoveList(oList As Control, pOpt As Integer) As Boolean
'Modified by Lydia 2022/01/11
'   Dim ii As Integer
'   Dim tmpArr As Variant 'Added by Lydia 2017/08/09
'
'   If oList.ListCount > 0 Then
'      ii = 0
'      Do While ii < oList.ListCount
'         If oList.Selected(ii) = True Then
'            RemoveList = True
'            oList.RemoveItem ii
'            'Added by Lydia 2017/08/09 移除FTP檔名
'            'Modified by Lydia 2022/01/11 判斷處理類型
'            If txtCF(6) <> "" Then
'            'If pOpt = 1 Then
'               txtCF(6) = Replace(txtCF(6), ",,", ",")
'               If Left(txtCF(6), 1) = "," Then txtCF(6) = Mid(txtCF(6), 2)
'               If Right(txtCF(6), 1) = "," Then txtCF(6) = Mid(txtCF(6), 1, Len(txtCF(6)) - 1)
'               tmpArr = Empty
'               tmpArr = Split(txtCF(6), ",")
'               If Trim(tmpArr(ii)) <> "" Then txtCF(6) = Replace(txtCF(6), Trim(tmpArr(ii)), "")
'            End If
'            'end 2017/08/09
'            ii = ii - 1
'         End If
'         ii = ii + 1
'      Loop
'
'      'If pOpt = 1 Then 'Added by Lydia 2022/01/11 判斷處理類型
'        'Added by Lydia 2017/08/09 重整FTP檔名
'        txtCF(6) = Replace(txtCF(6), ",,", ",")
'        If Left(txtCF(6), 1) = "," Then txtCF(6) = Mid(txtCF(6), 2)
'        If Right(txtCF(6), 1) = "," Then txtCF(6) = Mid(txtCF(6), 1, Len(txtCF(6)) - 1)
'        'end 2017/08/09
'      'End If 'Added by Lydia 2022/01/11
'   End If
   If pOpt = 0 Then
       oList.Tag = PUB_RemoveListBox2(oList, oList.Tag)
   ElseIf pOpt = 1 Then
       txtCF(6) = PUB_RemoveListBox2(oList, txtCF(6))
   End If
   RemoveList = True
   'end 2022/01/11
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

' 執行指令
'Private Sub OnAction(ByVal KeyCode As Integer)
Public Sub OnAction(ByVal KeyCode As Integer)
   Screen.MousePointer = vbHourglass
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         intRepType = 0 'Add by Amy 2025/03/20
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         SetCboPlace
         txtCR(19) = strUserNum
         SetlstUsers 0, txtCR(19)
         
      Case vbKeyF3 ' 修改
         intRepType = 0 'Add by Amy 2025/03/20
'         '2008/12/10 ADD BY SONIA 輸入者或語文權限為'Y'者才可修改
'         If m_bLanguage <> "Y" And strUserNum <> m_CR12 Then
'            MsgBox "您沒有此筆潛在客戶往來記錄修改權限 !!!", vbInformation
'            GoTo ExitSub
'         End If
'         '2008/12/10 END
         'Modify By Sindy 2019/7/26
         'If PUB_CheckModifyLimit_frm140402(m_CR12, "M") = False Then GoTo EXITSUB 'Add By Sindy 2009/04/30
         'Modify By Sindy 2024/10/1 傳入建檔人
         If PUB_CheckModifyLimit_frm140402(m_PCU51, m_CR12) = False Then GoTo EXITSUB
         '2019/7/26 END
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF5 ' 刪除
         intRepType = 0 'Add by Amy 2025/03/20
         'If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
         'Modify By Sindy 2019/7/26
         'If PUB_CheckModifyLimit_frm140402(m_CR12, "D") = False Then GoTo EXITSUB 'Add By Sindy 2019/6/5
         'Modify By Sindy 2024/10/1 傳入建檔人
         If PUB_CheckModifyLimit_frm140402(m_PCU51, m_CR12) = False Then GoTo EXITSUB
         '2019/7/26 END
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                GoTo EXITSUB
            End If
         'End If
         
      Case vbKeyF4 ' 查詢
         intRepType = 0 'Add by Amy 2025/03/20
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
         If m_EditMode = 4 Then txtCR(1).Tag = txtCR(1) '2008/12/10 ADD BY SONIA
         
         If OnWork = True Then
            UpdateToolbarState
         Else
            GoTo EXITSUB
         End If
         SetInputEntry
         
         'Add By Sindy 2022/6/15
         If Me.m_strIR01 <> "" Then
            If Not m_PrevForm Is Nothing Then
               Call m_PrevForm.GoNext
            End If
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
         End If
         '2022/6/15 END
         
      Case vbKeyF10 ' 取消
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  txtCR(1) = txtCR(1).Tag
                  m_EditMode = 0
                  SetInputEntry
                  ShowRecord
                  UpdateToolbarState
                  
                  'Add By Sindy 2022/6/15
                  If Me.m_strIR01 <> "" Then
                     If Not m_PrevForm Is Nothing Then
                        Call m_PrevForm.GoNext
                     End If
                     Screen.MousePointer = vbDefault
                     Unload Me
                     Exit Sub
                  End If
                  '2022/6/15 END
               End If
            Case Else
               txtCR(1) = txtCR(1).Tag
               m_EditMode = 0
               SetInputEntry
               ShowRecord
               UpdateToolbarState
         End Select
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
EXITSUB:
   Screen.MousePointer = vbDefault
   
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
               txtCR(1).SetFocus
               txtCR_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean, ii As Integer, jj As Integer
   Dim tmpArr1 As Variant, tmpArr2 As Variant 'Added by Lydia 2017/08/09
   
   For Each oText In txtCR
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         Cancel = False
         txtCR_Validate oText.Index, Cancel
         If Cancel = True Then
            oText.SetFocus
            txtCR_GotFocus oText.Index
            Exit Function
         End If
      End If
   Next
   '查詢
   If m_EditMode = 4 Then
      If txtCR(1) = "" Then
         ShowMsg "請輸入欲查詢之往來記錄編號 !"
         txtCR(1).SetFocus
         txtCR_GotFocus 1
         Exit Function
      End If
   '維護
   Else
      If txtCR(2).Text = "" Then
         ShowMsg "往來日期不可為空白 !"
         txtCR(2).SetFocus
         Exit Function
      End If
     
      If txtCR(3).Text = "" Then
         ShowMsg "往來對象不可為空白 !"
         txtCR(3).SetFocus
         Exit Function
      End If
      
      'Add By Sindy 2019/7/4
      'Add By Sindy 2021/3/25 + Not (Len(txtCR(3).Text) = 4 And IsNumeric(txtCR(3).Text) = True)
      If Right(txtCR(3).Text, 1) <> "0" And _
         Not (Len(txtCR(3).Text) = 4 And IsNumeric(txtCR(3).Text) = True) Then
         ShowMsg "往來對象第9碼只能為0 !"
         txtCR(3).SetFocus
         Exit Function
      End If
      
      'Modify By Sindy 2019/3/8
      'If lstSort.ListCount = 0 Then
      If txtCR(5).Text = "" Then
      '2019/3/8 END
         ShowMsg "往來類別不可為空白 !"
         'cboSort.SetFocus
'         txtCR(5).SetFocus
         Exit Function
      End If
      
      If txtCR(6).Text = "" Then
         ShowMsg "主旨不可為空白 !"
         txtCR(6).SetFocus
         Exit Function
      End If
      'Modify by Morgan 2009/2/6
      'If txtCR(7).Text = "" Then
      If cboPlace.Text = "" Then
         ShowMsg "場合不可為空白，若不在選項內請自行輸入 !"
         'txtCR(7).SetFocus
         cboPlace.SetFocus
         Exit Function
      ElseIf GetTextLength(cboPlace) > txtCR(7).MaxLength Then
         ShowMsg "場合長度超過限制(" & txtCR(7).MaxLength & "個字元)!"
         cboPlace.SetFocus
         Exit Function
      End If

      If lstUsers(0).ListCount = 0 Then
         ShowMsg "接洽同仁不可空白!"
         txtUserNo(0).SetFocus
         txtUserNo_GotFocus 0
         Exit Function
      End If
         
      'Added by Lydia 2017/08/09 檢查長度
      If CheckLengthIsOK(txtCF(2), 800, False) = False Then
         MsgBox "全部的附件檔名超過最大長度！" & vbCrLf & "(1個中文=2個字元)", vbCritical
         Exit Function
      End If
      
      'Added by Lydia 2017/08/09 檢查List和FTP檔名的數量是否一致
      strExc(1) = "附件順序有誤，請全部移除後再新增附件"
      If (txtCF(2) = "" And txtCF(6) <> "") Or (txtCF(2) <> "" And txtCF(6) = "") Then
          ShowMsg strExc(1)
          Exit Function
      End If
      
      tmpArr1 = Empty: tmpArr2 = Empty
      tmpArr1 = Split(txtCF(2), ",")
      'Modified by Lydia 2022/01/17 去掉CF06尾端的, ; txtCF(6)=>Mid(txtCF(6), 1, Len(txtCF(6)) - 1)
      If txtCF(6) <> "" Then 'Added by Lydia 2022/01/24 排除沒附件
          tmpArr2 = Split(Mid(txtCF(6), 1, Len(txtCF(6)) - 1), ",")
      'Added by Lydia 2022/01/24
      Else
          tmpArr2 = Split(txtCF(6), ",")
      End If
      'end 2022/01/24
      If UBound(tmpArr1) <> UBound(tmpArr2) Then
          ShowMsg strExc(1)
          Exit Function
      End If
      
      '預估一個ftp路徑約50字
      If UBound(tmpArr2) > Format(1100 / 50, "0") Then
         MsgBox "附件數量超過最大上限(" & Format(1100 / 50, "0") & ")！", vbCritical
         Exit Function
      End If
      For intI = 0 To UBound(tmpArr1)
         If (Trim(tmpArr1(intI)) = "" And Trim(tmpArr2(intI)) <> "") Or (Trim(tmpArr1(intI)) <> "" And Trim(tmpArr2(intI)) = "") Then
            ShowMsg strExc(1)
            Exit Function
         End If
      Next intI
      'end 2017/08/09
   End If
   
    'Added by Lydia 2022/01/11 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If
    
   TxtValidate = True
End Function

Private Sub UpdateFieldNewData()
   txtCR(7) = cboPlace.Text
   For Each oText In txtCR
      idx = oText.Index
      Select Case idx
         Case 2
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         'Added by Lydia 2022/04/18去掉多餘的,
         Case 4, 5, 19  '聯絡人04,往來類別05,接洽同仁19
            If Right(oText.Text, 1) = "," Then
                m_FieldList(idx).fiNewData = Mid(oText.Text, 1, Len(oText.Text) - 1)
            Else
                m_FieldList(idx).fiNewData = oText.Text
            End If
         'end 2022/04/18
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
   Dim stSQL As String, stCols As String, stValues As String
   Dim iErr As Integer, sErrMsg As String
   Dim varTemp1, varTemp2 'Add By Sindy 2019/2/25
   Dim j As Integer 'Add By Sindy 2019/2/25
   Dim strNewFile As String, longSize As Long
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans

   If txtCR(1) = "" Then
      m_FieldList(1).fiNewData = AutoNo("K", 6)
   End If

   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtCR
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
'         'Added by Lydia 2017/08/09 跳過FTP路徑
'         If idx = 20 Then
'            stCols = stCols & "," & m_FieldList(idx).fiName
'            stValues = stValues & ", NULL"
'         Else
'         'end 2017/08/09
            stCols = stCols & "," & m_FieldList(idx).fiName
            '文字
            If m_FieldList(idx).fiType = 0 Then
               stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
            '數字
            Else
               stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
            End If
'         End If 'end 2017/08/09
      End If
   Next
   'Modify By Sindy 2023/8/10 +CR09
   stCols = Mid(stCols, 2) & ",CR09"
   stValues = Mid(stValues, 2) & IIf(chkCR09.Value = 1, ",'Y'", ",''")
   stSQL = "INSERT INTO ContactRecord (" & stCols & ") Values (" & stValues & ")"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL
   
   UpdateRefNo m_FieldList(1).fiNewData
   
   'Added by Lydia 2017/08/09 判斷移檔日期
   If strSrvDate(1) >= CR_NewDate Then
      If PUB_UpdateCRAttFile(m_FieldList(1).fiNewData, txtCF(6), txtCF(6).Tag, lstAtt, iErr, sErrMsg) = False Then
         GoTo ErrHand
      Else
         If txtCF(6).Text <> txtCF(6).Tag Then
'            stSQL = "UPDATE ContactRecord SET CR20='" & txtCF(6).Text & "' WHERE CR01='" & m_FieldList(1).fiNewData & "' "
'            cnnConnection.Execute stSQL
            
            'Add By Sindy 2019/2/25 先全部刪除附件,再新增附件資訊
            varTemp1 = Split(txtCF(2), ",")
            varTemp2 = Split(txtCF(6), ",")
            If UBound(varTemp1) = UBound(varTemp2) Then
               For j = 0 To UBound(varTemp1)
                  longSize = GetFileSize(varTemp1(j), strNewFile)
                  If longSize = 0 Then
                     MsgBox "檔案大小為 0 有問題，請確認附件內容！"
                     GoTo ErrHand
                  End If
                  stSQL = "INSERT INTO CONTACTFILE(cf01,cf02,cf06,cf07)" & _
                           "VALUES('" & m_FieldList(1).fiNewData & "','" & ChgSQL(strNewFile) & "','" & ChgSQL(varTemp2(j)) & "','" & longSize & "')"
                  Pub_SeekTbLog stSQL
                  cnnConnection.Execute stSQL
               Next j
            Else
               MsgBox "附件資料檔名和路徑個數有誤，無法儲存！"
               GoTo ErrHand
            End If
            '2019/2/25 END
         End If
      End If
'   Else
'   'end 2017/08/09
'      'Add by Morgan 2009/5/19
'      '上傳附件檔
'      If UploadAtt(m_FieldList(1).fiNewData, iErr, sErrMsg) = False Then
'         GoTo ErrHand
'      End If
   End If 'end 2017/08/09
   
   'Add by Sindy 2022/6/15
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm140404"
   End If
   '2022/6/15 END
   
   'Added by Lydia 2023/02/02 異動通知：A14客戶名稱資訊不得宣傳
   txtCR(1) = m_FieldList(1).fiNewData
   Call GetMailToA14(m_EditMode)
   'end 2023/02/02
   
   cnnConnection.CommitTrans
   AddRecord = True
   
   txtCR(1) = m_FieldList(1).fiNewData
   txtCR(1).Tag = txtCR(1)     '2008/12/10 ADD BY SONIA
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   Dim iErr As Integer, sErrMsg As String
   
On Error GoTo ErrHand
   
   'Add By Sindy 2011/8/11
   If MsgBox(IIf(txtCF(2) <> "", "有附件", "") & "是否要刪除此筆往來記錄資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
   '2011/8/11 End
      cnnConnection.BeginTrans
      
      'Added by Lydia 2017/08/09 判斷移檔日期
      'If m_FieldList(20).fiOldData <> "" And strSrvDate(1) >= CR_NewDate Then
      If txtCF(6) <> "" Then
         txtCF(6) = "" '刪除全部附件
         If PUB_UpdateCRAttFile(m_FieldList(1).fiNewData, txtCF(6), txtCF(6).Tag, lstAtt, iErr, sErrMsg) = False Then
            GoTo ErrHand
         End If
      End If
      'end 2017/08/09
      'Add By Sindy 2019/2/25 刪除附件
      stSQL = "delete from ContactFile where cf01='" & txtCR(1) & "'"
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL
      '2019/2/25 END
      
      stSQL = "delete from ContactRecord where cr01='" & txtCR(1) & "'"
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL
      
'      'Add by Morgan 2009/6/3
'      'Modifie by Lydia 2017/08/09 判斷移檔日期之前
'      'If m_FieldList(9).fiOldData <> "" Then
'      If m_FieldList(9).fiOldData <> "" And strSrvDate(1) < CR_NewDate Then
'         If RemoveAtt(m_FieldList(1).fiNewData, m_FieldList(9).fiOldData, iErr, sErrMsg) = False Then
'            GoTo ErrHand
'         End If
'      End If
      
      Call GetMailToA14(m_EditMode) 'Added by Lydia 2023/02/02 異動通知：A14客戶名稱資訊不得宣傳

      cnnConnection.CommitTrans
   
      DelRecord = True
      ClearField
      txtCR(1).Tag = ""
   End If
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Function

'Add By Sindy 2019/2/26 解析檔案大小
Private Function GetFileSize(ByVal strFileOName As String, ByRef strFileNName As String) As Long
Dim strCF02 As String
Dim intEnd As Integer, intStar As Integer
   
   GetFileSize = 0
   strCF02 = UCase(strFileOName)
   If InStr(strCF02, "KB)") > 0 Then
      intEnd = InStrRev(strCF02, "KB)")
      intStar = InStrRev(strCF02, "(")
      strFileNName = Trim(Mid(strFileOName, 1, intStar - 1))
      GetFileSize = Val(Mid(strCF02, intStar + 1, Len(strCF02) - intEnd + 1))
   End If
End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   Dim iErr As Integer, sErrMsg As String
   Dim arrFile1
   Dim ii As Integer, bolRemove As Boolean
   Dim arrTmp, arrOldTmp, varTemp1 'Add By Sindy 2019/2/25
   Dim j As Integer 'Add By Sindy 2019/2/25
   Dim strNewFile As String, longSize As Long 'Add By Sindy 2019/2/25
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   stSQL = "begin user_data.user_enabled:=1; UPDATE ContactRecord SET "
   stSet = ""
   For Each oText In txtCR
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
'         'Added by Lydia 2017/08/09 跳過FTP路徑
'         If idx = 20 Then
'         Else
'         'end 2017/08/09
            '文字
            If m_FieldList(idx).fiType = 0 Then
               stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
            '數字
            Else
               stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
            End If
'         End If 'end 2017/08/09
      End If
   Next
   
   'Modify By Sindy 2023/8/10
   If chkCR09.Value = 1 Then
      If chkCR09.Tag <> "Y" Then
         bDifference = True
         stSet = stSet & ",CR09='Y'"
      End If
   Else
      If chkCR09.Tag = "Y" Then
         bDifference = True
         stSet = stSet & ",CR09=null"
      End If
   End If
   '2023/8/10 END
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where cr01='" & txtCR(1) & "'; end; "
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   
   UpdateRefNo txtCR(1)
   
   'Added by Lydia 2017/08/09 判斷移檔日期
   If strSrvDate(1) >= CR_NewDate And txtCF(6).Text <> txtCF(6).Tag Then
      If PUB_UpdateCRAttFile(m_FieldList(1).fiNewData, txtCF(6), txtCF(6).Tag, lstAtt, iErr, sErrMsg) = False Then
         GoTo ErrHand
      Else
'         stSQL = "UPDATE ContactRecord SET CR20='" & txtCF(6).Text & "' WHERE CR01='" & txtCR(1) & "' "
'         cnnConnection.Execute stSQL
         
         'Add By Sindy 2019/2/25 異動附件資訊
         'Add By Sindy 2019/3/6
         If bDifference = False Then
            'Update人,日,時
            stSQL = "UPDATE ContactRecord SET cr15='" & strUserNum & "',cr16=" & strSrvDate(1) & ",cr17=" & Left(Format(ServerTime, "000000"), 4) & " WHERE CR01='" & txtCR(1) & "'"
            cnnConnection.Execute stSQL
         End If
         '2019/3/6 END
         arrTmp = Empty: arrOldTmp = Empty: varTemp1 = Empty
         varTemp1 = Split(txtCF(2), ",")
         arrTmp = Split(txtCF(6), ",")
         arrOldTmp = Split(txtCF(6).Tag, ",")
         '先：刪除附件
         If txtCF(6).Tag <> "" Then
            For ii = 0 To UBound(arrOldTmp)
               If Trim(arrOldTmp(ii)) <> "" And InStr(txtCF(6), Trim(arrOldTmp(ii))) = 0 Then
                  stSQL = "delete from CONTACTFILE where cf01='" & txtCR(1) & "' and upper(cf06)='" & ChgSQL(UCase(arrOldTmp(ii))) & "'"
                  Pub_SeekTbLog stSQL
                  cnnConnection.Execute stSQL
               End If
            Next ii
         End If
         '後：新增附件
         If txtCF(6) <> "" Then
            For ii = 0 To UBound(arrTmp)
               If Trim(arrTmp(ii)) <> "" And InStr(txtCF(6).Tag, Trim(arrTmp(ii))) = 0 Then
                  longSize = GetFileSize(varTemp1(ii), strNewFile)
                  If longSize = 0 Then
                     MsgBox "檔案大小為 0 有問題，請確認附件內容！"
                     GoTo ErrHand
                  End If
                  stSQL = "INSERT INTO CONTACTFILE(cf01,cf02,cf06,cf07)" & _
                          "VALUES('" & txtCR(1) & "','" & ChgSQL(strNewFile) & "','" & ChgSQL(arrTmp(ii)) & "','" & longSize & "')"
                  Pub_SeekTbLog stSQL
                  cnnConnection.Execute stSQL
               End If
            Next ii
         End If
'         'Add By Sindy 2019/2/25 先全部刪除附件,再新增附件資訊
'         stSQL = "delete from ContactFile where cf01='" & txtCR(1) & "'"
'         Pub_SeekTbLog stSQL
'         cnnConnection.Execute stSQL
'         varTemp1 = Split(txtCF(2), ",")
'         varTemp2 = Split(txtCF(6), ",")
'         If UBound(varTemp1) = UBound(varTemp2) Then
'            For j = 0 To UBound(varTemp1)
'               stSQL = "INSERT INTO CONTACTFILE(cf01,cf02,cf06)" & _
'                        "VALUES('" & txtCR(1) & "','" & ChgSQL(varTemp1(j)) & "','" & ChgSQL(varTemp2(j)) & "')"
'               Pub_SeekTbLog stSQL
'               cnnConnection.Execute stSQL
'            Next j
'         Else
'            MsgBox "附件資料檔名和路徑個數有誤，無法儲存！"
'            GoTo ErrHand
'         End If
         '2019/2/25 END
      End If
'   Else
'   'end 2017/08/09
'        'Add by Morgan 2009/5/19
'        '上傳附件檔
'        If UploadAtt(m_FieldList(1).fiNewData, iErr, sErrMsg) = False Then
'           GoTo ErrHand
'        End If
'        '檔案有異動時，移掉的要刪除
'        bolRemove = False
'        If m_FieldList(9).fiNewData <> m_FieldList(9).fiOldData Then
'           arrFile1 = Split(m_FieldList(9).fiOldData, ",")
'           For ii = LBound(arrFile1) To UBound(arrFile1)
'              If InStr(m_FieldList(9).fiNewData & ",", arrFile1(ii) & ",") > 0 Then
'                 arrFile1(ii) = ""
'              Else
'                 bolRemove = True
'              End If
'           Next
'           If bolRemove = True Then
'              If RemoveAtt(m_FieldList(1).fiNewData, Join(arrFile1, ","), iErr, sErrMsg) = False Then
'                 GoTo ErrHand
'              End If
'           End If
'        End If
   End If 'end 2017/08/09
   
   Call GetMailToA14(m_EditMode) 'Added by Lydia 2023/02/02 異動通知：A14客戶名稱資訊不得宣傳
   
   cnnConnection.CommitTrans
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Function

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtCR
      oText.Locked = bLocked
   Next
   cboContact.Locked = bLocked
   cmdAddCont.Enabled = Not bLocked
   cmdRemCont.Enabled = Not bLocked
   cboSort.Locked = bLocked
   cmdAddSort.Enabled = Not bLocked
   cmdRemSort.Enabled = Not bLocked
   
   'Add by Morgan 2009/5/18
   cmdOpenAtt.Enabled = bLocked
      
   cmdAddAtt.Enabled = Not bLocked
   cmdRemAtt.Enabled = Not bLocked
   
   cmdReply.Enabled = Not bLocked
   cboPlace.Locked = bLocked
   Frame2.Visible = Not bLocked
   chkCR09.Enabled = Not bLocked 'Add By Sindy 2023/8/10
End Sub

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
         
'Removed by Morgan 2019/2/14 取消,內容分段需要輸跳行--Widen
'      Case vbKeyReturn
'         '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
'         KeyCode = 0
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
   End Select
End Sub

Private Sub UpdateRefNo(p_stCR01 As String)
   Dim ii As Integer, stRefNo As String, bolAdd As Boolean, bolRemove As Boolean
   Dim arRefNoOld
   For ii = 0 To lstCR18.ListCount - 1
      stRefNo = lstCR18.List(ii)
      bolAdd = False
      '原來未回覆
      If txtCR(18).Tag = "" Then
         bolAdd = True
      '原來有回覆但本單號為新加入
      ElseIf InStr(txtCR(18).Tag, stRefNo) = 0 Then
         bolAdd = True
      End If
      If bolAdd = True Then
         strSql = "update contactrecord set cr11=cr11||decode(cr11,null,'',',')||'" & p_stCR01 & "'" & _
            " where cr01='" & stRefNo & "' and (cr11 is null or instr(cr11,'" & p_stCR01 & "')=0)"
         adoTaie.Execute strSql, intI
      End If
   Next
   If txtCR(18).Tag <> "" Then
      arRefNoOld = Split(txtCR(18).Tag, ",")
      For ii = LBound(arRefNoOld) To UBound(arRefNoOld)
         stRefNo = arRefNoOld(ii)
         bolRemove = False
         If InStr(txtCR(18), stRefNo) = 0 Then
            bolRemove = True
         End If
         If bolRemove = True Then
            strSql = "update contactrecord set cr11=replace(replace(cr11,'," & p_stCR01 & "',''),'" & p_stCR01 & "','')" & _
               " where cr01='" & stRefNo & "' and instr(cr11,'" & p_stCR01 & "')>0"
            adoTaie.Execute strSql, intI
         End If
      Next
   End If
End Sub

'Add by Morgan 2009/2/6
Private Sub SetCboPlace(Optional sPlace As String)
   cboPlace.Clear
   cboPlace.Tag = "" 'Added by Lydia 2022/01/11
   cboPlace.AddItem "線上會議", 0 'Add By Sindy 2022/1/20
   cboPlace.AddItem "Email", 0    'add by sonia 2018/1/12
   cboPlace.AddItem "會議場合", 0
   cboPlace.AddItem "彼所/公司", 0
   cboPlace.AddItem "台一", 0
   If sPlace <> "" Then
      cboPlace.AddItem sPlace, 0
      cboPlace.ListIndex = 0
   End If
End Sub

Private Sub AddlstUsers(p_idx As Integer)
   Dim idx As Integer, bFound As Boolean
   
   If txtUserNo(p_idx) <> "" And lblName(p_idx) <> "" Then
      'Modify by Morgan 2011/8/26 員工編號已可非數字需做轉換
      'Modified by Lydia 2022/01/11 改成Form 2.0
'      For idx = 0 To lstUsers(p_idx).ListCount - 1
'         If lstUsers(p_idx).ItemData(idx) = PUB_Id2Num(txtUserNo(p_idx)) Then
'            MsgBox "員工已存在於接洽同仁清單中！"
'            txtUserNo(p_idx).SetFocus
'            txtUserNo_GotFocus p_idx
'            bFound = True
'            Exit For
'         End If
'      Next
         If InStr(lstUsers(p_idx).Tag, txtUserNo(p_idx)) > 0 Then
            MsgBox "員工已存在於開發人員清單中！"
            txtUserNo(p_idx).SetFocus
            txtUserNo_GotFocus p_idx
            bFound = True
         End If
         'end 2022/01/11
      If bFound = False Then
         lstUsers(p_idx).AddItem lblName(p_idx), 0
         'Modified by Lydia 2022/01/11 改成Form 2.0
         'lstUsers(p_idx).ItemData(0) = PUB_Id2Num(txtUserNo(p_idx))
         lstUsers(p_idx).Tag = txtUserNo(p_idx) & "," & lstUsers(p_idx).Tag
         txtUserNo(p_idx) = ""
         lblName(p_idx) = ""
      End If
   End If
End Sub

Private Function ComposeListX(p_index As Integer) As String
   'Modified by Lydia 2022/01/11 改成Form 2.0
'   strExc(1) = ""
'   If lstUsers(p_index).ListCount > 0 Then
'      strExc(1) = PUB_Num2Id(lstUsers(p_index).ItemData(0))
'      For intI = 1 To lstUsers(p_index).ListCount - 1
'         strExc(1) = strExc(1) & "," & PUB_Num2Id(lstUsers(p_index).ItemData(intI))
'      Next
'   End If
'   ComposeListX = strExc(1)
   ComposeListX = lstUsers(p_index).Tag
   'end 2022/01/11
End Function

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   
   lstUsers(p_idx).Clear
   lstUsers(p_idx).Tag = "" 'Added by Lydia 2022/01/11
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
                  'Modified by Lydia 2022/01/11 改成Form 2.0沒有ItemData屬性
                  'lstUsers(p_idx).ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
                   lstUsers(p_idx).Tag = .Fields(0) & "," & lstUsers(p_idx).Tag
                  .MoveLast
               End If
               .MoveNext
            Loop
         Next
         End With
      End If
   End If
End Sub

Private Sub RemovelstUsers(p_idx As Integer)
   'Modified by Lydia 2022/01/11 改成Form 2.0
'   Dim idx As Integer, ii As Integer
'   If lstUsers(p_idx).ListCount > 0 Then
'      ii = 0
'      For idx = 0 To lstUsers(p_idx).ListCount - 1
'         If lstUsers(p_idx).Selected(ii) = True Then
'            lstUsers(p_idx).RemoveItem ii
'            ii = ii - 1
'         End If
'         ii = ii + 1
'      Next
'   End If
   lstUsers(p_idx).Tag = PUB_RemoveListBox2(lstUsers(p_idx), lstUsers(p_idx).Tag)
   'end 2022/01/11
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

''Add By Sindy 2009/04/30
''檢查維護權限
''strModify:M.修改
''          D.刪除
''          A.新增
'Private Function CheckModifyLimit(strChkID As String, strModify As String) As Boolean
'   CheckModifyLimit = True
'   If Trim(strUserNum) = "" Or Trim(strChkID) = "" Then Exit Function
'   strModify = UCase(strModify)
'
'   '依LoginUser和輸入人員之部門第一碼判斷部門權限, 相同者才可輸入查詢
'   '但M51不受限制
'   strExc(0) = "SELECT A.ST03,B.ST03 FROM STAFF A,STAFF B " & _
'                      "WHERE A.ST01 = '" & strUserNum & "' " & _
'                           "AND B.ST01 = '" & Trim(strChkID) & "' "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If Trim(RsTemp(0)) = "M51" Then Exit Function
'      If strModify = "M" Then
'         If Trim(RsTemp(0)) = Trim(RsTemp(1)) Then Exit Function
'      Else
'         If Left(Trim(RsTemp(0)), 1) = Left(Trim(RsTemp(1)), 1) Then Exit Function
'      End If
'   Else
'      Exit Function
'   End If
'
'   CheckModifyLimit = False
'   If strModify = "M" Then
'      MsgBox "無修改權限 !!!", vbInformation
'   ElseIf strModify = "D" Then
'      MsgBox "無刪除權限 !!!", vbInformation
'   End If
'End Function

'Add by Morgan 2009/5/19
'附件
'Modified by Lydia 2022/01/11 ListBox=>Control
Private Function ComposeAttList(oList As Control) As String
   Dim iPos As Integer, stItem As String, stRtn As String, idx As Integer
   If oList.ListCount > 0 Then
      stItem = oList.List(0)
      stRtn = GetFileName(stItem)
      For idx = 1 To oList.ListCount - 1
         stItem = oList.List(idx)
         stRtn = stRtn & "," & GetFileName(stItem)
      Next
   End If
   ComposeAttList = stRtn
End Function

''Add by Morgan 2009/5/19
''上傳附件檔
'Private Function UploadAtt(ByVal stKey As String, Optional iErrNo As Integer, Optional stErrMsg As String) As Boolean
'   Dim hOpen As Long
'   Dim hConnection As Long
'   Dim hDir As Long
'   Dim bReturn As Boolean
'   Dim dwInternetFlags As Integer
'   Dim stDir As String
'   Dim stRemoteFile As String
'   Dim stLocalFile As String
'   Dim stItem As String
'   Dim idx As Integer
'   Dim iPos As Integer
'   Dim IsTimeOut As Boolean
'   Dim SeekTimer
'   Dim ACT_FTP_IP As String
'   Dim arrIP
'   Dim ii As Integer
'
'   iErrNo = 0
'   stErrMsg = ""
'
'   stDir = 往來記錄附件存放路徑
'   If lstAtt.ListCount > 0 Then
'      For idx = 0 To lstAtt.ListCount - 1
'         stItem = lstAtt.List(idx)
'         iPos = InStr(stItem, "\")
'         If iPos > 0 Then
'            If InStrRev(stItem, " (") > 0 Then
'               stLocalFile = Left(stItem, InStrRev(stItem, " (") - 1)
'            Else
'               stLocalFile = stItem
'            End If
'            stRemoteFile = GetFileName(stLocalFile)
'
'            If hOpen = 0 Then
'               hOpen = InternetOpen("Taie FTP", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
'               If hOpen = 0 Then
'                  iErrNo = 1
'                  stErrMsg = "網路錯誤！"
'                  GoTo OutPort
'               Else
'                  IsTimeOut = True
'                  If GOOD_FTP_IP <> "" Then
'                     arrIP = Split(GOOD_FTP_IP & ";" & FTP_IP, ";")
'                  Else
'                     arrIP = Split(FTP_IP, ";")
'                  End If
'                  For ii = LBound(arrIP) To UBound(arrIP)
'                     ACT_FTP_IP = arrIP(ii)
'                     If ACT_FTP_IP <> "" Then
'                        '偵測 FTPServer 是否存在
'                        If Winsock1.State Then Winsock1.Close
'                        Winsock1.Connect ACT_FTP_IP, 21
'                        IsTimeOut = False
'                        SeekTimer = Timer
'                        Do While Winsock1.State = 6 And IsTimeOut = False
'                           DoEvents
'                           If Timer - SeekTimer > 1 Then
'                              IsTimeOut = True
'                           End If
'                        Loop
'                        If Winsock1.State Then Winsock1.Close
'                        If IsTimeOut = False Then
'                           Exit For
'                        End If
'                     End If
'                  Next
'
'                  '若是超過時間
'                  If IsTimeOut = True Then
'                     iErrNo = 2
'                     stErrMsg = "無法與FTP Server建立連線！"
'                     GoTo OutPort
'                  Else
'                     GOOD_FTP_IP = ACT_FTP_IP
'                  End If
'
'                  hConnection = InternetConnect(hOpen, ACT_FTP_IP, FTP_Port, _
'                     "pgmid", "pgmpwd", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
'                  If hConnection = 0 Then
'                     iErrNo = 3
'                     stErrMsg = "無法與FTP Server建立連線！"
'                     GoTo OutPort
'                  ElseIf FtpSetCurrentDirectory(hConnection, stDir) = False Then
'                     iErrNo = 4
'                     stErrMsg = "切換至往來記錄目錄失敗！"
'                     GoTo OutPort
'                  '切換至往來記錄單號目錄
'                  ElseIf FtpSetCurrentDirectory(hConnection, stKey) = False Then
'                     hDir = FtpCreateDirectory(hConnection, stKey)
'                     If hDir = 0 Then
'                        iErrNo = 5
'                        stErrMsg = "建立往來記錄單號目錄失敗！"
'                        GoTo OutPort
'                     ElseIf FtpSetCurrentDirectory(hConnection, stKey) = False Then
'                        iErrNo = 6
'                        stErrMsg = "切換至往來記錄單號目錄失敗！"
'                        GoTo OutPort
'                     End If
'                  End If
'               End If
'            End If
'
'            dwInternetFlags = FTP_TRANSFER_TYPE_BINARY
'            bReturn = FtpPutFile(hConnection, stLocalFile, stRemoteFile, dwInternetFlags, 0)
'            ' Upload successfully
'            If bReturn = False Then
'               iErrNo = 7
'               stErrMsg = "檔案上傳失敗！"
'               GoTo OutPort
'            End If
'         End If
'      Next
'   End If
'
'   UploadAtt = True
'
'OutPort:
'   If hOpen <> 0 Then InternetCloseHandle (hOpen)
'   If hConnection <> 0 Then InternetCloseHandle (hConnection)
'   If Winsock1.State Then Winsock1.Close
'
'End Function

'Add by Morgan 2009/5/19
'刪除附件檔
'Removed by Morgan 2024/8/2 沒用了
'Private Function RemoveAtt(ByVal stKey As String, stFiles As String, Optional iErrNo As Integer, Optional stErrMsg As String) As Boolean
'   Dim hOpen As Long
'   Dim hConnection As Long
'   Dim bReturn As Boolean
'   Dim stDir As String
'   Dim IsTimeOut As Boolean
'   Dim SeekTimer
'   Dim ACT_FTP_IP As String
'   Dim arrIP
'   Dim ii As Integer, jj As Integer
'   Dim arrFile
'   Dim stRemoteFile As String
'
'   iErrNo = 0
'   stErrMsg = ""
'
'   stDir = 往來記錄附件存放路徑
'   arrFile = Split(stFiles, ",")
'   For jj = LBound(arrFile) To UBound(arrFile)
'      If arrFile(jj) <> "" Then
'         If hOpen = 0 Then
'            hOpen = InternetOpen("Taie FTP", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
'            If hOpen = 0 Then
'               iErrNo = 1
'               stErrMsg = "網路錯誤！"
'               GoTo OutPort
'            Else
'               IsTimeOut = True
'               If GOOD_FTP_IP <> "" Then
'                  arrIP = Split(GOOD_FTP_IP & ";" & FTP_IP, ";")
'               Else
'                  arrIP = Split(FTP_IP, ";")
'               End If
'               For ii = LBound(arrIP) To UBound(arrIP)
'                  ACT_FTP_IP = arrIP(ii)
'                  If ACT_FTP_IP <> "" Then
'                     '偵測 FTPServer 是否存在
'                     If Winsock1.State Then Winsock1.Close
'                     Winsock1.Connect ACT_FTP_IP, 21
'                     IsTimeOut = False
'                     SeekTimer = Timer
'                     Do While Winsock1.State = 6 And IsTimeOut = False
'                        DoEvents
'                        If Timer - SeekTimer > 1 Then
'                           IsTimeOut = True
'                        End If
'                     Loop
'                     If Winsock1.State Then Winsock1.Close
'                     If IsTimeOut = False Then
'                        Exit For
'                     End If
'                  End If
'               Next
'
'               '若是超過時間
'               If IsTimeOut = True Then
'                  iErrNo = 2
'                  stErrMsg = "無法與FTP Server建立連線！"
'                  GoTo OutPort
'               Else
'                  GOOD_FTP_IP = ACT_FTP_IP
'               End If
'
'               hConnection = InternetConnect(hOpen, ACT_FTP_IP, FTP_Port, _
'                  "pgmid", "pgmpwd", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
'               If hConnection = 0 Then
'                  iErrNo = 3
'                  stErrMsg = "無法與FTP Server建立連線！"
'                  GoTo OutPort
'               ElseIf FtpSetCurrentDirectory(hConnection, stDir) = False Then
'                  iErrNo = 4
'                  stErrMsg = "切換至往來記錄目錄失敗！"
'                  GoTo OutPort
'               '切換至往來記錄單號目錄
'               ElseIf FtpSetCurrentDirectory(hConnection, stKey) = False Then
'                  '無法切換時當作已刪除
'                  'iErrNo = 6
'                  'stErrMsg = "切換至往來記錄單號目錄失敗！"
'                  'GoTo OutPort
'                  Exit For
'               End If
'            End If
'         End If
'         If InStrRev(arrFile(jj), " (") > 0 Then
'            stRemoteFile = Left(arrFile(jj), InStrRev(arrFile(jj), " (") - 1)
'         Else
'            stRemoteFile = arrFile(jj)
'         End If
'         '刪除檔案不控制成功與否
'         bReturn = FtpDeleteFile(hConnection, stRemoteFile)
'      End If
'   Next
'
'   RemoveAtt = True
'
'OutPort:
'   If hOpen <> 0 Then InternetCloseHandle (hOpen)
'   If hConnection <> 0 Then InternetCloseHandle (hConnection)
'   If Winsock1.State Then Winsock1.Close
'
'End Function
'Add by Morgan 2009/5/19
Private Function GetFileName(ByVal FullPath As String) As String
   Dim stItem As String, iPos As Integer
   stItem = FullPath
   iPos = InStr(stItem, "\")
   Do While iPos > 0
      stItem = Mid(stItem, iPos + 1)
      iPos = InStr(stItem, "\")
   Loop
   GetFileName = stItem
End Function

'Modify By Sindy 2018/5/30 改共用Func
''Added by Lydia 2017/08/09 新增／刪除附件
'Private Function UpdateAttFile(ByVal stKey As String, Optional iErrNo As Integer, Optional stErrMsg As String) As Boolean
'Dim arrTmp As Variant, arrOldTmp As Variant
'Dim stFtpPath As String
'Dim ii As Integer
'Dim strMid As String
'Dim stFileName As String
'
'On Error GoTo OutPort
'
'   iErrNo = 0
'   stErrMsg = ""
'
'   arrTmp = Empty: arrOldTmp = Empty
'   arrTmp = Split(txtCF(6).Text, ",")
'   arrOldTmp = Split(txtCF(6).Tag, ",")
'
'   '先：刪除附件
'   If txtCF(6).Tag <> "" Then
'    For ii = 0 To UBound(arrOldTmp)
'       If Trim(arrOldTmp(ii)) <> "" And InStr(txtCF(6).Text, Trim(arrOldTmp(ii))) = 0 Then
'          If PUB_DelFtpFile2(stKey, Trim(arrOldTmp(ii)), cTableName) = False Then
'             GoTo OutPort
'          End If
'       End If
'    Next ii
'   End If
'
'   '後：新增附件
'   If txtCF(6).Text <> "" Then
'    For ii = 0 To UBound(arrTmp)
'       If Trim(arrTmp(ii)) <> "" And InStr(txtCF(6).Tag, Trim(arrTmp(ii))) = 0 Then
'          stFileName = Trim(Mid(lstAtt.List(ii), 1, InStrRev(lstAtt.List(ii), "(") - 1))
'          If PUB_PutFtpFile(stFileName, stKey, IIf(InStr(Trim(arrTmp(ii)), stKey & "_") = 0, stKey & "_", "") & Trim(arrTmp(ii)), stFtpPath, cTableName) = False Then
'             GoTo OutPort
'          Else
'             strMid = strMid & IIf(strMid <> "", ",", "") & stFtpPath
'          End If
'       ElseIf Trim(arrTmp(ii)) <> "" Then
'          strMid = strMid & IIf(strMid <> "", ",", "") & Trim(arrTmp(ii))
'       End If
'    Next ii
'    txtCF(6).Text = strMid
'   End If
'
'   UpdateAttFile = True
'
'   Exit Function
'
'OutPort:
'   iErrNo = Err.Number
'   stErrMsg = Err.Description
'
'End Function

'Added by Lydia 2017/08/09 搬檔
'Removed by Morgan 2024/8/2 不用的標記為註解，檢查程式碼才知時可略過
'Private Sub Cmd1_Click()
'Dim stSQL As String, intR As Integer
'Dim rsQuery As ADODB.Recordset
'Dim stOldDir As String, stNewDir As String, stNewPath As String
'Dim oFileName As String, mFileName As String
'Dim strGrp As String, strList As String, strNameList As String
'Dim tmpArr As Variant
'Dim strTmpExc As String
'Dim stDownFile As String
'Dim strLost As String, strLostId As String
'
'   stOldDir = 往來記錄附件存放路徑
'   stNewDir = PUB_GetFtpTableDir(stNewDir) & cTableName
'   stSQL = "select CR01,CR09 from CONTACTRECORD where NVL(CR09,'N') <> 'N' AND NVL(CR20,'N')='N' order by cr01 "
'   intR = 0
'   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
'   If intR = 1 Then
'      With rsQuery
'         .MoveFirst
'         MsgBox "開始工作，共" & .RecordCount & "筆記錄!"
'         Do While Not .EOF
'            '清除暫存檔
'            PUB_KillTempFile "$$*.*"
'
'            If strGrp <> "" & .Fields("CR01") Then
'               If strGrp <> "" Then
'                  strTmpExc = strTmpExc & "UPDATE CONTACTRECORD SET CR20='" & strList & "' WHERE CR01='" & strGrp & "' ;"
'               End If
'               strList = "": strNameList = ""
'               strGrp = "" & .Fields("CR01")
'               tmpArr = Empty
'               tmpArr = Split("" & .Fields("CR09"), ",")
'            End If
'
'            For intR = 0 To UBound(tmpArr)
'               If Trim(tmpArr(intR)) <> "" Then
'                   '先下載檔案
'                   stDownFile = ""
'                   '因為有附件檔名有包含刮號,直接到模組處理舊檔名
'                   strExc(1) = PUB_StringFilter(Trim(tmpArr(intR)))
'                   If InStr(strExc(1), "(") > 0 And InStr(strExc(1), " (") = 0 Then
'                      strExc(1) = Mid(strExc(1), 1, InStrRev(strExc(1), "(") - 1) & " " & Mid(strExc(1), InStrRev(strExc(1), "("))
'                   End If
'                   PUB_OpenFtpFile "" & .Fields("CR01"), strExc(1), Winsock1, "1", False, stDownFile
'
'                   If stDownFile = "" Then
'                       strLostId = strLostId & .Fields("CR01") & "," & IIf(Len(strLostId) > 50, vbCrLf, "")
'                       strLost = strLost & .Fields("CR01") & "_" & Trim(tmpArr(intR)) & vbCrLf
'                   Else
'                        oFileName = Trim(tmpArr(intR))
'                        oFileName = Trim(Mid(oFileName, 1, InStrRev(oFileName, "(") - 1))
'                        '新-FTP檔名(非中文)
'                        mFileName = PUB_GetNewFileNameSec(oFileName, "2", strNameList, "" & .Fields("CR01"))
'
'                        If PUB_PutFtpFile(stDownFile, strGrp, mFileName, stNewPath, cTableName) = True Then
'                           strList = strList & IIf(strList <> "", ",", "") & stNewPath
'                        Else
'                           MsgBox "Error !"
'                           Exit Sub
'                        End If
'                   End If
'               End If
'            Next intR
'            .MoveNext
'         Loop
'
'         '最後一筆
'         strTmpExc = strTmpExc & "UPDATE CONTACTRECORD SET CR20='" & strList & "' WHERE CR01='" & strGrp & "' ;"
'      End With
'
'      '清除暫存檔
'      PUB_KillTempFile "$$*.*"
'
'      If strTmpExc <> "" Then
'         tmpArr = Empty
'         tmpArr = Split(strTmpExc, ";")
'         cnnConnection.BeginTrans
'           For intR = 0 To UBound(tmpArr)
'              If Trim(tmpArr(intR)) <> "" Then
'                 cnnConnection.Execute Trim(tmpArr(intR)), intI
'              End If
'           Next intR
'         cnnConnection.CommitTrans
'         MsgBox "工作結束!"
'      End If
'   End If
'
'   If strLost <> "" Then
'      PUB_SendMail "QPGMR", "A3034", "", 往來記錄附件存放路徑 & "在NT2缺少檔案", "資料夾:" & strLostId & vbCrLf & vbCrLf & "檔案名稱:" & strLost
'   End If
'
'   Set rsQuery = Nothing
'   Exit Sub
'
'ErrHandle:
'   cnnConnection.RollbackTrans
'
'OutPort:
'   Exit Sub
'
'End Sub

'Added by Lydia 2023/02/02 異動通知：A14客戶名稱資訊不得宣傳
Private Sub GetMailToA14(ByVal pMode As Integer)
Dim strTitle As String, strTo As String, strContent As String

    If InStr(txtCR(5).Text & "," & txtCR(5).Tag, "A14") > 0 Then
        Select Case pMode
            Case 1
                 strTitle = "新增"
            Case 2
                 If txtCR(5).Text <> txtCR(5).Tag Then strTitle = "修改"
            Case 3
                 strTitle = "刪除"
        End Select
        If strTitle <> "" Then
            strTitle = "不得宣傳之異動記錄：" & strTitle & txtCR(3) & " " & lbl1.Caption
            '預設業拓部門人員(F41)
            strSql = "select st01 from staff where st03='F41' and st04='1' and substr(st01,4,1)<>'9' and substr(st01,1,1) <> 'F' order by st01 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strTo = RsTemp.GetString(adClipString, , , ";")
               If Right(strTo, 1) = ";" Then strTo = Mid(strTo, 1, Len(strTo) - 1)
            End If
            If strTo <> "" Then
               strContent = "記錄編號：" & txtCR(1) & vbCrLf & _
                                 "往來日期：" & ChangeWStringToWDateString(txtCR(2)) & vbCrLf
               If pMode = 2 Then
                   strContent = strContent & "異動前類別：" & cboSort.Tag & vbCrLf & _
                                 "異動後類別：" & cboSort.Text & vbCrLf
               Else
                   strContent = strContent & "往來類別：" & cboSort.Text & vbCrLf
               End If
               strContent = strContent & "主旨：" & txtCR(6) & vbCrLf & _
                                "內容：" & txtCR(8)
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                           "values('" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                            ",'" & ChgSQL(strTitle) & "','" & ChgSQL(strContent) & "')"
               cnnConnection.Execute strSql
            End If
        End If
    End If
End Sub

