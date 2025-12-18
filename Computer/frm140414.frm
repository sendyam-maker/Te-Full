VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140414 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件表單簽核人員設定"
   ClientHeight    =   6100
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6100
   ScaleWidth      =   8950
   Begin VB.TextBox txtST04 
      Height          =   270
      Left            =   5790
      MaxLength       =   1
      TabIndex        =   64
      Top             =   570
      Width           =   492
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140414.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140414.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140414.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140414.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140414.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140414.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140414.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140414.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140414.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140414.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140414.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   8950
      _ExtentX        =   15787
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
      Height          =   5160
      Left            =   60
      TabIndex        =   27
      Top             =   870
      Width           =   8870
      _ExtentX        =   15646
      _ExtentY        =   9102
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm140414.frx":20F4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtF0101"
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "LblST04"
      Tab(0).Control(6)=   "txtF0101_2"
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(8)=   "Label1(1)"
      Tab(0).Control(9)=   "Label1(15)"
      Tab(0).Control(10)=   "LabDept"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm140414.frx":2110
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label10"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "grd1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt1(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt1(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdok"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txt1(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txt1(3)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "CboF0102"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin VB.ComboBox CboF0102 
         Height          =   260
         ItemData        =   "frm140414.frx":212C
         Left            =   1110
         List            =   "frm140414.frx":2139
         TabIndex        =   23
         Top             =   720
         Width           =   1515
      End
      Begin VB.TextBox txtF0101 
         Height          =   315
         Left            =   -73980
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '平面
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H00004000&
         Height          =   1272
         Left            =   -74370
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   59
         Text            =   "frm140414.frx":215F
         Top             =   3810
         Width           =   8025
      End
      Begin VB.Frame Frame3 
         Caption         =   "接洽記錄單"
         Height          =   2655
         Left            =   -69120
         TabIndex        =   51
         Top             =   870
         Width           =   2895
         Begin VB.Label Label6 
            Caption         =   "特例簽核"
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
            Height          =   730
            Left            =   2610
            TabIndex        =   67
            Top             =   1230
            Width           =   220
         End
         Begin VB.Label Label5 
            Caption         =   "一般簽核"
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
            Height          =   730
            Left            =   2610
            TabIndex        =   66
            Top             =   300
            Width           =   220
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00404000&
            BorderWidth     =   2
            X1              =   60
            X2              =   2830
            Y1              =   1140
            Y2              =   1150
         End
         Begin MSForms.TextBox textCUID 
            Height          =   470
            Index           =   3
            Left            =   90
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   2130
            Width           =   2720
            VariousPropertyBits=   -2147467233
            BackColor       =   16777215
            ForeColor       =   4194368
            Size            =   "4789;820"
            Caption         =   "LblFM2"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.ComboBox Combo2 
            Height          =   300
            Index           =   18
            Left            =   1080
            TabIndex        =   18
            Top             =   1770
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.ComboBox Combo2 
            Height          =   300
            Index           =   17
            Left            =   1080
            TabIndex        =   17
            Top             =   1470
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.ComboBox Combo2 
            Height          =   300
            Index           =   16
            Left            =   1080
            TabIndex        =   16
            Top             =   1170
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
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
            Left            =   1080
            TabIndex        =   15
            Top             =   840
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
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
            Left            =   1080
            TabIndex        =   14
            Top             =   540
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
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
            Left            =   1080
            TabIndex        =   13
            Top             =   240
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員1："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   19
            Left            =   90
            TabIndex        =   57
            Top             =   330
            Width           =   1040
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員2："
            Height          =   180
            Index           =   18
            Left            =   90
            TabIndex        =   56
            Top             =   630
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員3："
            Height          =   180
            Index           =   17
            Left            =   90
            TabIndex        =   55
            Top             =   930
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員4："
            Height          =   180
            Index           =   16
            Left            =   90
            TabIndex        =   54
            Top             =   1260
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員5："
            Height          =   180
            Index           =   14
            Left            =   90
            TabIndex        =   53
            Top             =   1560
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員6："
            Height          =   180
            Index           =   13
            Left            =   90
            TabIndex        =   52
            Top             =   1860
            Width           =   990
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "銷案/帳單"
         Height          =   2655
         Left            =   -72030
         TabIndex        =   44
         Top             =   870
         Width           =   2895
         Begin MSForms.TextBox textCUID 
            Height          =   465
            Index           =   2
            Left            =   90
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   2130
            Width           =   2715
            VariousPropertyBits=   -2147467233
            BackColor       =   16777215
            ForeColor       =   4194368
            Size            =   "4789;820"
            Caption         =   "LblFM2"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.ComboBox Combo2 
            Height          =   300
            Index           =   12
            Left            =   1080
            TabIndex        =   12
            Top             =   1740
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.ComboBox Combo2 
            Height          =   300
            Index           =   11
            Left            =   1080
            TabIndex        =   11
            Top             =   1440
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
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
            Left            =   1080
            TabIndex        =   10
            Top             =   1140
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
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
            Left            =   1080
            TabIndex        =   9
            Top             =   840
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
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
            Left            =   1080
            TabIndex        =   8
            Top             =   540
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
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
            Left            =   1080
            TabIndex        =   7
            Top             =   240
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員1："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   12
            Left            =   90
            TabIndex        =   50
            Top             =   330
            Width           =   1040
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員2："
            Height          =   180
            Index           =   11
            Left            =   90
            TabIndex        =   49
            Top             =   630
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員3："
            Height          =   180
            Index           =   6
            Left            =   90
            TabIndex        =   48
            Top             =   930
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員4："
            Height          =   180
            Index           =   5
            Left            =   90
            TabIndex        =   47
            Top             =   1230
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員5："
            Height          =   180
            Index           =   4
            Left            =   90
            TabIndex        =   46
            Top             =   1530
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員6："
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   45
            Top             =   1830
            Width           =   990
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "結案單"
         Height          =   2655
         Left            =   -74940
         TabIndex        =   37
         Top             =   870
         Width           =   2895
         Begin MSForms.TextBox textCUID 
            Height          =   465
            Index           =   1
            Left            =   90
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   2130
            Width           =   2715
            VariousPropertyBits=   -2147467233
            BackColor       =   16777215
            ForeColor       =   4194368
            Size            =   "4789;820"
            Caption         =   "LblFM2"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.ComboBox Combo2 
            Height          =   300
            Index           =   6
            Left            =   1080
            TabIndex        =   6
            Top             =   1740
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2628;529"
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
            Left            =   1080
            TabIndex        =   5
            Top             =   1440
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2628;529"
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
            Left            =   1080
            TabIndex        =   4
            Top             =   1140
            Width           =   1485
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
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
            Left            =   1080
            TabIndex        =   3
            Top             =   840
            Width           =   1490
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2628;529"
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
            Left            =   1080
            TabIndex        =   2
            Top             =   540
            Width           =   1485
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
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
            Left            =   1080
            TabIndex        =   1
            Top             =   240
            Width           =   1485
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2619;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員6："
            Height          =   180
            Index           =   2
            Left            =   90
            TabIndex        =   43
            Top             =   1830
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員5："
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   42
            Top             =   1530
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員4："
            Height          =   180
            Index           =   7
            Left            =   90
            TabIndex        =   41
            Top             =   1230
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員3："
            Height          =   180
            Index           =   8
            Left            =   90
            TabIndex        =   40
            Top             =   930
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員2："
            Height          =   180
            Index           =   9
            Left            =   90
            TabIndex        =   39
            Top             =   630
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "簽核人員1："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   90
            TabIndex        =   38
            Top             =   330
            Width           =   1040
         End
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   5310
         MaxLength       =   3
         TabIndex        =   22
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   4290
         MaxLength       =   3
         TabIndex        =   21
         Top             =   390
         Width           =   915
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   315
         Left            =   6960
         TabIndex        =   24
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   20
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   1110
         MaxLength       =   6
         TabIndex        =   19
         Top             =   390
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Bindings        =   "frm140414.frx":236B
         Height          =   3615
         Left            =   -74970
         TabIndex        =   29
         Top             =   750
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6368
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Bindings        =   "frm140414.frx":2380
         Height          =   4010
         Left            =   60
         TabIndex        =   25
         Top             =   1050
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   7073
         _Version        =   393216
         Cols            =   17
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
         _Band(0).Cols   =   17
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "表單類別："
         Height          =   180
         Left            =   180
         TabIndex        =   68
         Top             =   780
         Width           =   900
      End
      Begin VB.Label LblST04 
         AutoSize        =   -1  'True
         Caption         =   "(已離職)"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   -73950
         TabIndex        =   60
         Top             =   690
         Visible         =   0   'False
         Width           =   660
      End
      Begin MSForms.TextBox txtF0101_2 
         Height          =   300
         Left            =   -73140
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   380
         Width           =   1395
         VariousPropertyBits=   671105055
         Size            =   "2461;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "註：若非下拉選單之人員亦可自行輸入員工編號"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74700
         TabIndex        =   58
         Top             =   3570
         Width           =   5805
      End
      Begin VB.Line Line2 
         X1              =   5100
         X2              =   5430
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label3 
         Caption         =   "部門："
         Height          =   255
         Left            =   3690
         TabIndex        =   36
         Top             =   420
         Width           =   585
      End
      Begin VB.Line Line1 
         X1              =   1920
         X2              =   2250
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label2 
         Caption         =   "員工代號："
         Height          =   255
         Left            =   180
         TabIndex        =   35
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   34
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "部門別："
         Height          =   180
         Index           =   15
         Left            =   -71610
         TabIndex        =   33
         Top             =   420
         Width           =   720
      End
      Begin VB.Label LabDept 
         AutoSize        =   -1  'True
         Height          =   240
         Left            =   -70800
         TabIndex        =   32
         Top             =   390
         Width           =   1965
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "日期起："
         Height          =   180
         Left            =   -71760
         TabIndex        =   31
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74910
         TabIndex        =   30
         Top             =   420
         Width           =   900
      End
      Begin VB.Line Line4 
         X1              =   -73290
         X2              =   -72600
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line5 
         X1              =   -70320
         X2              =   -69720
         Y1              =   510
         Y2              =   510
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否含離職人員：           （Y：含離職）"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   4320
      TabIndex        =   65
      Top             =   620
      Width           =   3140
   End
End
Attribute VB_Name = "frm140414"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2023/12/22 修改抓新部門程式
'Memo by Sonia 2022/1/14 改成Form2.0(txtF0101_2,Combo2(1)~Combo2(18),grd1改Fonts)
'Create by Sindy 2015/1/6
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


Private Sub cmdok_Click()
   Call doQuery
End Sub

Private Sub Combo2_GotFocus(Index As Integer)
   InverseTextBox Combo2(Index)
End Sub

'modify by sonia 2022/1/14
'Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
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
   If m_EditMode <> 0 And Combo2(Index) <> "" Then
      If Index <> 1 And Index <> 7 And Index <> 13 And Left(Combo2(Index), 5) = txtF0101 Then
         MsgBox "不可為本人！", vbExclamation
         Call Combo2_GotFocus(Index)
         Cancel = True
         Exit Sub
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
      '檢查輸入順序
'      If Index >= 1 And Index <= 6 Then
'         If (Trim(Combo2(2)) <> "" And Trim(Combo2(1)) = "") Or _
'            (Trim(Combo2(3)) <> "" And Trim(Combo2(2)) = "") Or _
'            (Trim(Combo2(4)) <> "" And Trim(Combo2(3)) = "") Or _
'            (Trim(Combo2(5)) <> "" And Trim(Combo2(4)) = "") Or _
'            (Trim(Combo2(6)) <> "" And Trim(Combo2(5)) = "") Then
      If Index >= 1 And Index <= 4 Then
         If (Trim(Combo2(2)) <> "" And Trim(Combo2(1)) = "") Or _
            (Trim(Combo2(3)) <> "" And Trim(Combo2(2)) = "") Or _
            (Trim(Combo2(4)) <> "" And Trim(Combo2(3)) = "") Then
            MsgBox "請依序輸入簽核人員！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         'For i = 1 To 6
         For i = 1 To 4
            If i <> Index Then
               If Trim(Combo2(i)) <> "" And Left(Trim(Combo2(i)), 5) = Left(Trim(Combo2(Index)), 5) Then
                  MsgBox "資料重覆！", vbExclamation
                  Combo2(Index).SetFocus
                  Call Combo2_GotFocus(Index)
                  Cancel = True
                  Exit Sub
               End If
            End If
         Next i
      End If
'      If Index >= 7 And Index <= 12 Then
'         If (Trim(Combo2(8)) <> "" And Trim(Combo2(7)) = "") Or _
'            (Trim(Combo2(9)) <> "" And Trim(Combo2(8)) = "") Or _
'            (Trim(Combo2(10)) <> "" And Trim(Combo2(9)) = "") Or _
'            (Trim(Combo2(11)) <> "" And Trim(Combo2(10)) = "") Or _
'            (Trim(Combo2(12)) <> "" And Trim(Combo2(11)) = "") Then
      If Index >= 7 And Index <= 10 Then
         If (Trim(Combo2(8)) <> "" And Trim(Combo2(7)) = "") Or _
            (Trim(Combo2(9)) <> "" And Trim(Combo2(8)) = "") Or _
            (Trim(Combo2(10)) <> "" And Trim(Combo2(9)) = "") Then
            MsgBox "請依序輸入簽核人員！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         'For i = 7 To 12
         For i = 7 To 10
            If i <> Index Then
               If Trim(Combo2(i)) <> "" And Left(Trim(Combo2(i)), 5) = Left(Trim(Combo2(Index)), 5) Then
                  MsgBox "資料重覆！", vbExclamation
                  Combo2(Index).SetFocus
                  Call Combo2_GotFocus(Index)
                  Cancel = True
                  Exit Sub
               End If
            End If
         Next i
      End If
'      If Index >= 13 And Index <= 18 Then
'         If (Trim(Combo2(14)) <> "" And Trim(Combo2(13)) = "") Or _
'            (Trim(Combo2(15)) <> "" And Trim(Combo2(14)) = "") Or _
'            (Trim(Combo2(16)) <> "" And Trim(Combo2(15)) = "") Or _
'            (Trim(Combo2(17)) <> "" And Trim(Combo2(16)) = "") Or _
'            (Trim(Combo2(18)) <> "" And Trim(Combo2(17)) = "") Then
      'Modify By Sindy 2022/9/26
      If Index >= 13 And Index <= 15 Then
         If (Trim(Combo2(14)) <> "" And Trim(Combo2(13)) = "") Or _
            (Trim(Combo2(15)) <> "" And Trim(Combo2(14)) = "") Then
            MsgBox "請依序輸入簽核人員！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         'For i = 13 To 18
         For i = 13 To 15 '16
            If i <> Index Then
               If Trim(Combo2(i)) <> "" And Left(Trim(Combo2(i)), 5) = Left(Trim(Combo2(Index)), 5) Then
                  MsgBox "資料重覆！", vbExclamation
                  Combo2(Index).SetFocus
                  Call Combo2_GotFocus(Index)
                  Cancel = True
                  Exit Sub
               End If
            End If
         Next i
      End If
      'Add By Sindy 2022/9/26
      If Index >= 16 And Index <= 18 Then
         If (Trim(Combo2(17)) <> "" And Trim(Combo2(16)) = "") Or _
            (Trim(Combo2(18)) <> "" And Trim(Combo2(17)) = "") Then
            MsgBox "請依序輸入簽核人員！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         For i = 16 To 18
            If i <> Index Then
               If Trim(Combo2(i)) <> "" And Left(Trim(Combo2(i)), 5) = Left(Trim(Combo2(Index)), 5) Then
                  MsgBox "資料重覆！", vbExclamation
                  Combo2(Index).SetFocus
                  Call Combo2_GotFocus(Index)
                  Cancel = True
                  Exit Sub
            End If
               End If
         Next i
      End If
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn: 'vbKeyReturn / 13
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
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
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
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
   'OnAction vbKeyF4
   OnAction vbKeyF10
   
   'Add By Sindy 2023/2/13
   textCUID(1).BackColor = &H8000000F
   textCUID(2).BackColor = &H8000000F
   textCUID(3).BackColor = &H8000000F
   '2023/2/13 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140414 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub

'Add By Sindy 2025/11/12
Private Sub grd1_DblClick()
   If m_CurrKEY(0) <> "" And m_CurrKEY(1) <> "" Then
      '查詢目前資料列
      UpdateCtrlData
   End If
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim k As Integer

If dblPrevRow > GRD1.Rows - 1 Then dblPrevRow = GRD1.Rows - 1
GRD1.Visible = False
tmpMouseRow = GRD1.row
GRD1.Visible = True
If tmpMouseRow <> 0 Then
   GRD1.row = tmpMouseRow
   GRD1.col = 0
   If GRD1.CellBackColor <> &HFFC0C0 Then
      If dblPrevRow > 0 Then
         GRD1.row = dblPrevRow
         For j = 0 To GRD1.Cols - 1
            '員工姓名
            If j = 1 Then
               GRD1.col = 1
               If GRD1.TextMatrix(dblPrevRow, 10) <> "1" And GRD1.TextMatrix(dblPrevRow, 10) <> "" Then '已離職
                  GRD1.CellBackColor = &HFF& '紅色
               Else
                  GRD1.CellBackColor = QBColor(15)
               End If
            '簽核人員1~6
            ElseIf (j >= 3 And j <= 8) Then
               GRD1.col = j
               k = 8 + j
               If GRD1.TextMatrix(dblPrevRow, k) <> "1" And GRD1.TextMatrix(dblPrevRow, k) <> "" Then '已離職
                  GRD1.CellBackColor = &HFF& '紅色
               Else
                  GRD1.CellBackColor = QBColor(15)
               End If
            '其他欄位
            Else
               GRD1.col = j
               GRD1.CellBackColor = QBColor(15)
            End If
         Next j
      End If
      GRD1.row = tmpMouseRow
      dblPrevRow = tmpMouseRow
      For i = 0 To GRD1.Cols - 1
          GRD1.col = i
          GRD1.CellBackColor = &HFFC0C0
      Next i
      '記錄目前資料列
      m_CurrKEY(0) = GRD1.TextMatrix(GRD1.row, 9)
      m_CurrKEY(1) = GetStaffDepartment(m_CurrKEY(0), , True) 'Modify By Sindy 2025/2/25 + bolST93=True
      'UpdateCtrlData '查詢目前資料列
   End If
End If
GRD1.Visible = True
End Sub

'Add By Sindy 2025/11/12
Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 1 Then
      If m_CurrKEY(0) <> "" And m_CurrKEY(1) <> "" Then
         '查詢目前資料列
         UpdateCtrlData
      End If
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

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim bolChk As Boolean

TxtValidate = False

If txtF0101.Text = "" Then
   MsgBox "員工代號不可以空白！", vbExclamation
   txtF0101.SetFocus
   Exit Function
End If

If m_EditMode = 1 Then
   ' 檢查記錄是否已存在
   If IsRecordExist(txtF0101) = True Then
      MsgBox "該筆記錄已存在", vbOKOnly, "更新資料"
      txtF0101.SetFocus
      Exit Function
   End If
End If

bolChk = False
For i = 1 To 18
   If Trim(Combo2(i).Text) <> "" Then
      bolChk = True
      Exit For
   End If
Next i
If bolChk = False Then
   MsgBox "簽核人員不可全部空白！", vbExclamation
   Exit Function
End If

For i = 1 To Combo2.UBound
   Cancel = False
   Combo2_Validate i, Cancel
   If Cancel = True Then
      Exit Function
   End If
Next i

TxtValidate = True
End Function

' 更新資料
Private Function SaveData(strEditMode As Integer) As Boolean
Dim strKEY01 As String, strKEY02 As String, bolModify As Boolean

On Error GoTo ErrHand
   
   SaveData = False
   
   strKEY01 = txtF0101
   strKEY02 = GetStaffDepartment(txtF0101, , True) 'Modify By Sindy 2025/2/25 + bolST93=True
   
   cnnConnection.BeginTrans
   
   For i = 1 To 3
      strExc(0) = "Select F0101 From FLOW001 Where F0101='" & strKEY01 & "' and F0102='" & i & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      bolModify = False
      If intI = 1 Then
        bolModify = True
      End If
      '有記錄但簽核人員均無設定,則視為刪除
      If bolModify = True Then
         If (i = 1 And Trim(Combo2(1)) = "") Or _
            (i = 2 And Trim(Combo2(7)) = "") Or _
            (i = 3 And Trim(Combo2(13)) = "") Then
            strSql = "DELETE FROM FLOW001 WHERE F0101 = " & CNULL(strKEY01) & " and F0102='" & i & "'"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
            GoTo goStep
         End If
      End If
      If i = 1 And Trim(Combo2(1)) = "" Then GoTo goStep
      If i = 2 And Trim(Combo2(7)) = "" Then GoTo goStep
      If i = 3 And Trim(Combo2(13)) = "" Then GoTo goStep
      
      If i = 1 Then j = 1
      If i = 2 Then j = 7
      If i = 3 Then j = 13
      '新增
      If bolModify = False Then
         strSql = "INSERT INTO FLOW001(F0101,F0102,F0103,F0104,F0105,F0106,F0107,F0108) VALUES(" & CNULL(strKEY01) & "," & CNULL(CStr(i)) & _
                     "," & CNULL(Left(Trim(Combo2(j)), 5)) & "," & CNULL(Left(Trim(Combo2(j + 1)), 5)) & _
                     "," & CNULL(Left(Trim(Combo2(j + 2)), 5)) & "," & CNULL(Left(Trim(Combo2(j + 3)), 5)) & _
                     "," & CNULL(Left(Trim(Combo2(j + 4)), 5)) & "," & CNULL(Left(Trim(Combo2(j + 5)), 5)) & _
                     ")"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      '修改
      Else
         If Left(Trim(Combo2(j).Text), 5) <> Left(Trim(Combo2(j).Tag), 5) Or _
            Left(Trim(Combo2(j + 1).Text), 5) <> Left(Trim(Combo2(j + 1).Tag), 5) Or _
            Left(Trim(Combo2(j + 2).Text), 5) <> Left(Trim(Combo2(j + 2).Tag), 5) Or _
            Left(Trim(Combo2(j + 3).Text), 5) <> Left(Trim(Combo2(j + 3).Tag), 5) Or _
            Left(Trim(Combo2(j + 4).Text), 5) <> Left(Trim(Combo2(j + 4).Tag), 5) Or _
            Left(Trim(Combo2(j + 5).Text), 5) <> Left(Trim(Combo2(j + 5).Tag), 5) Then
            strSql = "update FLOW001 set " & _
                        " F0103=" & CNULL(Left(Trim(Combo2(j)), 5)) & ",F0104=" & CNULL(Left(Trim(Combo2(j + 1)), 5)) & _
                        ",F0105=" & CNULL(Left(Trim(Combo2(j + 2)), 5)) & ",F0106=" & CNULL(Left(Trim(Combo2(j + 3)), 5)) & _
                        ",F0107=" & CNULL(Left(Trim(Combo2(j + 4)), 5)) & ",F0108=" & CNULL(Left(Trim(Combo2(j + 5)), 5)) & _
                     " where F0101=" & CNULL(strKEY01) & " and F0102='" & i & "'"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
      End If
goStep:
   Next i
   
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

   strKEY01 = txtF0101

   cnnConnection.BeginTrans

   strSql = "DELETE FROM FLOW001 WHERE F0101 = " & CNULL(strKEY01)
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
   strKEY01 = txtF0101
   strKEY02 = GetStaffDepartment(txtF0101, , True) 'Modify By Sindy 2025/2/25 + bolST93=True
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
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
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         If SaveData(m_EditMode) = True Then
             RefreshRange
             doQuery
             SetKeyReadOnly True
         Else
             Exit Function
         End If
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         If SaveData(m_EditMode) = False Then Exit Function
         doQuery
         SetKeyReadOnly True
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
            doQuery
            SetKeyReadOnly True
         Else
            Exit Function
         End If
      Case 4: '查詢
         If txtF0101 <> "" Then
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
      Case 0, 1, 4: If Me.txtF0101.Visible = True Then txtF0101.SetFocus
      Case 2: If Me.Combo2(1).Visible = True Then Combo2(1).SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String

   IsRecordExist = False
   '只查詢在職或留職停薪人員的資料... Modify By Sindy 2016/4/27 取消離職控管 and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=F0101 and sc02=(select max(sc02) from Staff_Change where sc01=F0101)))
   strSql = "SELECT * FROM FLOW001,staff WHERE F0101=st01(+) and F0101=" & CNULL(strKEY01)
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
      '只查詢在職或留職停薪人員的資料
      'Modify By Sindy 2016/4/27 取消離職控管 and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=F0101 and sc02=(select max(sc02) from Staff_Change where sc01=F0101)))
      strSql = "SELECT F0101,nvl(A0921,A0901) FROM FLOW001,STAFF,ACC090,ACC090NEW" & _
               " WHERE F0101=ST01(+) and ST03=A0901(+) and ST93=A0921(+)" & _
               " and nvl(A0921,A0901)||F0101='" & m_CurrKEY(1) & m_CurrKEY(0) & "' group by nvl(A0921,A0901),F0101 order by nvl(A0921,A0901) asc,F0101 asc"
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
      
      '只查詢在職或留職停薪人員的資料
      'Modify By Sindy 2016/4/27 取消離職控管 and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=F0101 and sc02=(select max(sc02) from Staff_Change where sc01=F0101)))
      'Modify By Sindy 2023/11/2 +ST04
      strSql = "SELECT F0101,nvl(A0921,A0901) FROM FLOW001,STAFF,ACC090,ACC090NEW" & _
               " WHERE F0101=ST01(+) and ST03=A0901(+) and ST93=A0921(+)" & _
               IIf(txtST04 = "Y", "", " and ST04='1'") & _
               " group by nvl(A0921,A0901),F0101 order by nvl(A0921,A0901) asc,F0101 asc"
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

   '只查詢在職或留職停薪人員的資料
   'Modify By Sindy 2016/4/27 取消離職控管 and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=F0101 and sc02=(select max(sc02) from Staff_Change where sc01=F0101)))
   'Modify By Sindy 2023/11/2 +ST04
   strSql = "SELECT F0101,nvl(A0921,A0901) FROM FLOW001,STAFF,ACC090,ACC090NEW" & _
            " WHERE F0101=ST01(+) and ST03=A0901(+) and ST93=A0921(+)" & _
            " and nvl(A0921,A0901)||F0101<'" & m_CurrKEY(1) & m_CurrKEY(0) & "'" & _
            IIf(txtST04 = "Y", "", " and ST04='1'") & _
            " group by nvl(A0921,A0901),F0101 order by nvl(A0921,A0901) desc,F0101 desc "
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

   '只查詢在職或留職停薪人員的資料
   'Modify By Sindy 2016/4/27 取消離職控管 and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=F0101 and sc02=(select max(sc02) from Staff_Change where sc01=F0101)))
   'Modify By Sindy 2023/11/2 +ST04
   strSql = "SELECT F0101,nvl(A0921,A0901) FROM FLOW001,STAFF,ACC090,ACC090NEW" & _
            " WHERE F0101=ST01(+) and ST03=A0901(+) and ST93=A0921(+)" & _
            IIf(txtST04 = "Y", "", " and ST04='1'") & _
            " group by nvl(A0921,A0901),F0101 order by nvl(A0921,A0901) asc,F0101 asc "
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

   '只查詢在職或留職停薪人員的資料
   'Modify By Sindy 2016/4/27 取消離職控管 and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=F0101 and sc02=(select max(sc02) from Staff_Change where sc01=F0101)))
   'Modify By Sindy 2023/11/2 +ST04
   strSql = "SELECT F0101,nvl(A0921,A0901) FROM FLOW001,STAFF,ACC090,ACC090NEW" & _
            " WHERE F0101=ST01(+) and ST03=A0901(+) and ST93=A0921(+)" & _
            " and nvl(A0921,A0901)||F0101>'" & m_CurrKEY(1) & m_CurrKEY(0) & "'" & _
            IIf(txtST04 = "Y", "", " and ST04='1'") & _
            " group by nvl(A0921,A0901),F0101 order by nvl(A0921,A0901) asc,F0101 asc "
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

   '只查詢在職或留職停薪人員的資料
   'Modify By Sindy 2016/4/27 取消離職控管 where (st04='1' or '04'=(select sc03 from Staff_Change where sc01=F0101 and sc02=(select max(sc02) from Staff_Change where sc01=F0101)))
   'Modify By Sindy 2023/11/2 +ST04
   strSql = "SELECT F0101,nvl(A0921,A0901) FROM FLOW001,STAFF,ACC090,ACC090NEW" & _
            " WHERE F0101=ST01(+) and ST03=A0901(+) and ST93=A0921(+)" & _
            IIf(txtST04 = "Y", "", " and ST04='1'") & _
            " group by nvl(A0921,A0901),F0101 order by nvl(A0921,A0901) asc,F0101 asc "
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
         SSTab1.Tab = 0
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         'Add By Sindy 2022/9/26
         '檢查人員是否存在或離職
         If LblST04.Visible = True Then
            MsgBox "人員已離職！", vbExclamation
            Exit Sub
         End If
         '2022/9/26 END
         m_EditMode = 2
         Call txtF0101_LostFocus
         SetCtrlReadOnly False
         SetKeyReadOnly True
         SSTab1.Tab = 0
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         SSTab1.Tab = 0
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
         SSTab1.Tab = 0
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
      Case vbKeyF9: ', vbKeyReturn
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
   txtF0101.Locked = bEnable
   If bEnable Then txtF0101.BackColor = &H8000000F Else txtF0101.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   For i = 1 To Combo2.UBound
      Combo2(i).Locked = bEnable
      If bEnable Then Combo2(i).BackColor = &H8000000F Else Combo2(i).BackColor = &H80000005
   Next i
End Sub

Private Sub ClearField()
   LabDept.Caption = Empty
   txtF0101 = Empty
   txtF0101_2 = Empty
   For i = 1 To Combo2.UBound
      'modify by sonia 2022/1/14
      'Combo2(i).Clear
      Combo2(i) = ""
      Combo2(i).Tag = ""
   Next i
   LblST04.Visible = False 'Add By Sindy 2022/9/26
   
   'Add By Sindy 2023/2/13
   textCUID(1).Text = ""
   textCUID(2).Text = ""
   textCUID(3).Text = ""
   '2023/2/13 END
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub doQuery()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCon As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   dblPrevRow = 0
   strCon = ""
   'Modify By Sindy 2016/4/27 員工代號查詢,簽核人員為查詢代號的資料也要出現,就算是離職也要可查出
'   If txt1(0) <> "" Then
'      strCon = strCon & "and F0101>='" & txt1(0) & "' "
'   End If
'   If txt1(1) <> "" Then
'      strCon = strCon & "and F0101<='" & txt1(1) & "' "
'   End If
   If txt1(0) <> "" And txt1(1) <> "" Then
      strCon = strCon & "and (" & _
                        "(F0101>='" & txt1(0) & "' and F0101<='" & txt1(1) & "') " & _
                        "or (F0103>='" & txt1(0) & "' and F0103<='" & txt1(1) & "') " & _
                        "or (F0104>='" & txt1(0) & "' and F0104<='" & txt1(1) & "') " & _
                        "or (F0105>='" & txt1(0) & "' and F0105<='" & txt1(1) & "') " & _
                        "or (F0106>='" & txt1(0) & "' and F0106<='" & txt1(1) & "') " & _
                        "or (F0107>='" & txt1(0) & "' and F0107<='" & txt1(1) & "') " & _
                        "or (F0108>='" & txt1(0) & "' and F0108<='" & txt1(1) & "') " & _
                        ") "
   End If
   '2016/4/27 END
   If txt1(2) <> "" Then
      'Memo By Sindy 2023/12/22 修改抓新部門程式
      'strCon = strCon & "and s0.st03>='" & txt1(2) & "' "
      strCon = strCon & "and s0.st93>='" & txt1(2) & "' "
   End If
   If txt1(3) <> "" Then
      'Memo By Sindy 2023/12/22 修改抓新部門程式
      'strCon = strCon & "and s0.st03<='" & txt1(3) & "' "
      strCon = strCon & "and s0.st93<='" & txt1(3) & "' "
   End If
   'Modify By Sindy 2022/11/16 秀玲:請剔除離職人員的資料
   'Modify By Sindy 2023/11/2
   'If Not (txt1(0) <> "" And txt1(1) <> "") Then
   If txtST04 <> "Y" Then
      strCon = strCon & "and s0.st04='1' "
   End If
   '2022/11/16 END
   
   'Add By Sindy 2025/5/15
   If Trim(CboF0102.Text) <> "" Then
      strCon = strCon & "and F0102='" & Left(Trim(CboF0102.Text), 1) & "' "
   End If
   
   GRD1.Rows = 2
   GRD1.Clear
   '只查詢在職或留職停薪人員的資料
   'Modify By Sindy 2016/4/27 取消離職控管 "and (s0.st04='1' or '04'=(select sc03 from Staff_Change where sc01=F0101 and sc02=(select max(sc02) from Staff_Change where sc01=F0101))) "
   'Modify By Sindy 2022/11/21 + Order + ,F0102
   strSql = "select nvl(A0922,A0902),s0.ST02,decode(F0102,'1','結案單','2','銷案銷帳單','3','接洽單')" & _
            ",s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02,F0101" & _
            ",s0.ST04,s1.ST04,s2.ST04,s3.ST04,s4.ST04,s5.ST04,s6.ST04 " & _
            "from FLOW001,ACC090,STAFF s0,ACC090NEW " & _
            ",STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6 " & _
            "where F0101=s0.ST01(+) " & _
            "and s0.ST03=A0901(+) and s0.ST93=A0921(+) " & _
            "and F0103=s1.ST01(+) " & _
            "and F0104=s2.ST01(+) " & _
            "and F0105=s3.ST01(+) " & _
            "and F0106=s4.ST01(+) " & _
            "and F0107=s5.ST01(+) " & _
            "and F0108=s6.ST01(+) " & strCon & _
            "order by nvl(A0921,A0901),F0101,F0102"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
   End If
   rsTmp.Close
   SetDataListWidth
   Call GetSelChage 'Add By Sindy 2016/4/27
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault

EXITSUB:
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2016/4/27 已離職人員標示紅色
Private Sub GetSelChage()
Dim k As Integer
   
   GRD1.Visible = False
   If GRD1.Rows - 1 > 0 Then
      For j = 1 To GRD1.Rows - 1
         GRD1.row = j
         '員工姓名
         GRD1.col = 1
         If GRD1.TextMatrix(j, 10) <> "1" And GRD1.TextMatrix(j, 10) <> "" Then '已離職
            GRD1.CellBackColor = &HFF& '紅色
         End If
         '簽核人員1~6
         For i = 3 To 8
            GRD1.col = i
            k = 8 + i
            If GRD1.TextMatrix(j, k) <> "1" And GRD1.TextMatrix(j, k) <> "" Then '已離職
               GRD1.CellBackColor = &HFF& '紅色
            End If
         Next i
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

   '只查詢在職或留職停薪人員的資料
   'Modify By Sindy 2016/4/27 取消離職控管 "and (s0.st04='1' or '04'=(select sc03 from Staff_Change where sc01=F0101 and sc02=(select max(sc02) from Staff_Change where sc01=F0101))) "
   'Modify By Sindy 2023/2/13 語法有問題
'   strSql = "SELECT FLOW001.*,A0901,A0902,A0911,s0.ST02 s0_ST02" & _
'            ",s1.ST02 s1_ST02,s2.ST02 s2_ST02,s3.ST02 s3_ST02,s4.ST02 s4_ST02 " & _
'            ",s5.ST02 s5_ST02,s6.ST02 s6_ST02 " & _
'            "FROM FLOW001,STAFF s0,ACC090" & _
'            ",STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF " & _
'            "WHERE F0101=s0.ST01(+) and s0.ST03=A0901(+) and F0101='" & m_CurrKEY(0) & "' " & _
'            "and F0103=s1.ST01(+) and F0104=s2.ST01(+) and F0105=s3.ST01(+) " & _
'            "and F0106=s4.ST01(+) and F0107=s5.ST01(+) and F0108=s6.ST01(+) " & _
'            "order by F0102 asc"
   strSql = "SELECT FLOW001.*,nvl(A0921,A0901) A0901,nvl(A0922,A0902) A0902,A0911,s0.ST02 s0_ST02" & _
            " FROM FLOW001,STAFF s0,ACC090,ACC090NEW" & _
            " WHERE F0101=s0.ST01(+) and s0.ST03=A0901(+) and s0.ST93=A0921(+) and F0101='" & m_CurrKEY(0) & "'" & _
            "order by F0102 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("A0902")) = False Then LabDept.Caption = rsTmp.Fields("A0901") & "  " & GetPrjSalesBlack(rsTmp.Fields("A0901"))
      If IsNull(rsTmp.Fields("F0101")) = False Then txtF0101 = rsTmp.Fields("F0101"): txtF0101_2 = rsTmp.Fields("s0_ST02")
      'Add By Sindy 2022/9/26
      '檢查人員是否存在或離職
      If ChkStaffST04(Trim(txtF0101), False) = True Then
         LblST04.Visible = True
      End If
      '2022/9/26 END
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         j = rsTmp.Fields("F0102")
         If j = 1 Then j = 1
         If j = 2 Then j = 7
         If j = 3 Then j = 13
         If IsNull(rsTmp.Fields("F0103")) = False Then
            Combo2(j).Text = Left(Trim(rsTmp.Fields("F0103")) & Space(5), 7) & GetPrjSalesNM(rsTmp.Fields("F0103"))
            Combo2(j).Tag = Combo2(j).Text
         End If
         If IsNull(rsTmp.Fields("F0104")) = False Then
            Combo2(j + 1).Text = Left(Trim(rsTmp.Fields("F0104")) & Space(5), 7) & GetPrjSalesNM(rsTmp.Fields("F0104"))
            Combo2(j + 1).Tag = Combo2(j + 1).Text
         End If
         If IsNull(rsTmp.Fields("F0105")) = False Then
            Combo2(j + 2).Text = Left(Trim(rsTmp.Fields("F0105")) & Space(5), 7) & GetPrjSalesNM(rsTmp.Fields("F0105"))
            Combo2(j + 2).Tag = Combo2(j + 2).Text
         End If
         If IsNull(rsTmp.Fields("F0106")) = False Then
            Combo2(j + 3).Text = Left(Trim(rsTmp.Fields("F0106")) & Space(5), 7) & GetPrjSalesNM(rsTmp.Fields("F0106"))
            Combo2(j + 3).Tag = Combo2(j + 3).Text
         End If
         If IsNull(rsTmp.Fields("F0107")) = False Then
            Combo2(j + 4).Text = Left(Trim(rsTmp.Fields("F0107")) & Space(5), 7) & GetPrjSalesNM(rsTmp.Fields("F0107"))
            Combo2(j + 4).Tag = Combo2(j + 4).Text
         End If
         If IsNull(rsTmp.Fields("F0108")) = False Then
            Combo2(j + 5).Text = Left(Trim(rsTmp.Fields("F0108")) & Space(5), 7) & GetPrjSalesNM(rsTmp.Fields("F0108"))
            Combo2(j + 5).Tag = Combo2(j + 5).Text
         End If
         
         'Add By Sindy 2023/2/13
         ' 更新CUID
         UpdateCUID rsTmp, "" & rsTmp.Fields("F0102")
         '2023/2/13 END
         
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   SSTab1.Tab = 0
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2023/2/13
' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset, Index As Integer)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("f0109")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("f0109")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("f0109"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("f0110")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("f0110")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("f0110"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("f0111")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("f0111")) = False Then
         strTemp = rsSrcTmp.Fields("f0111")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("f0112")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("f0112")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("f0112"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("f0113")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("f0113")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("f0113"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("f0114")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("f0114")) = False Then
         strTemp = rsSrcTmp.Fields("f0114")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID(Index) = "CREATE:" & strCName & " " & _
              strCDate & " " & _
              strCTime & String(10, " ") & vbCrLf & _
              "UPDATE:" & strUName & " " & _
              strUDate & " " & _
              strUTime
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset

   '只查詢在職或留職停薪人員的資料
   'Modify By Sindy 2016/4/27 取消離職控管 and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=F0101 and sc02=(select max(sc02) from Staff_Change where sc01=F0101)))
   'Modify By Sindy 2023/11/2 +ST04
   strSql = "select F0101,nvl(A0921,A0901) from FLOW001,STAFF,ACC090,ACC090NEW" & _
            " where F0101=ST01(+) and ST03=A0901(+) and ST93=A0921(+)" & _
            IIf(txtST04 = "Y", "", " and ST04='1'") & _
            " group by nvl(A0921,A0901),F0101 order by nvl(A0921,A0901) asc,F0101 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then m_FirstKEY(0) = Trim(rsTmp.Fields(0))
      If IsNull(rsTmp.Fields(1)) = False Then m_FirstKEY(1) = Trim(rsTmp.Fields(1))
   End If
   rsTmp.Close

   '只查詢在職或留職停薪人員的資料
   'Modify By Sindy 2016/4/27 取消離職控管 and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=F0101 and sc02=(select max(sc02) from Staff_Change where sc01=F0101)))
   'Modify By Sindy 2023/11/2 +ST04
   strSql = "select F0101,nvl(A0921,A0901) from FLOW001,STAFF,ACC090,ACC090NEW" & _
            " where F0101=ST01(+) and ST03=A0901(+) and ST93=A0921(+)" & _
            IIf(txtST04 = "Y", "", " and ST04='1'") & _
            " group by nvl(A0921,A0901),F0101 order by nvl(A0921,A0901) desc,F0101 desc "
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
GRD1.col = 2: GRD1.Text = "表單類別"
GRD1.ColWidth(2) = 1000
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 3: GRD1.Text = "簽核人員1"
GRD1.ColWidth(3) = 900
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 4: GRD1.Text = "簽核人員2"
GRD1.ColWidth(4) = 900
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 5: GRD1.Text = "簽核人員3"
GRD1.ColWidth(5) = 900
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 6: GRD1.Text = "簽核人員4"
GRD1.ColWidth(6) = 900
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 7: GRD1.Text = "簽核人員5"
GRD1.ColWidth(7) = 900
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 8: GRD1.Text = "簽核人員6"
GRD1.ColWidth(8) = 900
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 9: GRD1.Text = "F0101"
GRD1.ColWidth(9) = 0
GRD1.CellAlignment = flexAlignLeftCenter
'Add By Sindy 2016/4/27
GRD1.col = 10: GRD1.Text = "s0.ST04"
GRD1.ColWidth(10) = 0
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 11: GRD1.Text = "s1.ST04"
GRD1.ColWidth(11) = 0
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 12: GRD1.Text = "s2.ST04"
GRD1.ColWidth(12) = 0
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 13: GRD1.Text = "s3.ST04"
GRD1.ColWidth(13) = 0
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 14: GRD1.Text = "s4.ST04"
GRD1.ColWidth(14) = 0
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 15: GRD1.Text = "s5.ST04"
GRD1.ColWidth(15) = 0
GRD1.CellAlignment = flexAlignLeftCenter
GRD1.col = 16: GRD1.Text = "s6.ST04"
GRD1.ColWidth(16) = 0
GRD1.CellAlignment = flexAlignLeftCenter
'2016/4/27 END
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         'Modify By Sindy 2025/5/15 mark
'         '檢查員工編號規則
'         'Modify By Sindy 2022/11/28 + And Left(txt1(Index).Text, 1) <> "W" And Left(txt1(Index).Text, 1) <> "M"
'         If txt1(Index).Text <> "" And Left(txt1(Index).Text, 1) <> "W" And Left(txt1(Index).Text, 1) <> "M" Then
'            If ChkStaffID(txt1(Index)) Then
'               Call txt1_GotFocus(Index)
'               Cancel = True
'               Exit Sub
'            End If
'         End If
         '2025/5/15 END
         
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 2, 3 '部門
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub

Private Sub txtF0101_GotFocus()
   InverseTextBox txtF0101
End Sub

Private Sub txtF0101_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtF0101_LostFocus()
Dim Rs As New ADODB.Recordset
Dim strText(18) As String
Dim strA0911 As String, strA0925 As String 'Add By Sindy 2023/12/22
   
   If m_EditMode <> 0 And txtF0101 <> "" Then
      txtF0101_2 = GetStaffName(txtF0101, True)
      'Modify By Sindy 2025/2/25 + bolST93=True
      LabDept.Caption = GetStaffDepartment(txtF0101, , True) & "  " & GetPrjSalesBlack(GetStaffDepartment(txtF0101, , True))
      strA0911 = GetStaffA0911(txtF0101, strA0925)
      
      For i = 1 To Combo2.UBound
         strText(i) = Combo2(i).Text
         'modify by sonia 2022/1/14
         'Combo2(i).Clear
         'Modify By Sindy 2023/12/22
'         Combo2(i) = ""
'         'end 2022/1/14
'         Combo2(i).AddItem ""
         Call SetB1003Combo(Combo2(i), strA0911, strA0925)
         Combo2(i).Text = strText(i)
         '2023/12/22 END
      Next i
'      Rs.CursorLocation = adUseClient
'      strSql = "select ST01,ST02 " & _
'               "From staff " & _
'               "where substr(st01,1,1) in (" & ST01CodeNum1 & ") " & _
'               "and st04='1' " & _
'               "and substr(st01,4,1)<>'9' " & _
'               "and st01 not in('60000','96029','96030') " & _
'               "and ST03 in(Select A0901 From ACC090 Where A0911='" & m_A0911 & "') " & _
'               "order by st01 "
'      Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      While Not Rs.EOF
'         For i = 1 To Combo2.UBound
'            Combo2(i).AddItem Left(Rs.Fields(0).Value & Space(7), 7) & Rs.Fields(1).Value
'         Next i
'         Rs.MoveNext
'      Wend
'      If Rs.State <> adStateClosed Then Rs.Close
'      Set Rs = Nothing
'      For i = 1 To Combo2.UBound
'         Combo2(i).Text = strText(i)
'      Next i
   End If
End Sub

Private Sub txtF0101_Validate(Cancel As Boolean)
Dim Rs As New ADODB.Recordset

   If txtF0101.Text = "" Then txtF0101_2 = ""

   If m_EditMode <> 0 And txtF0101 <> "" Then
      'Modify By Sindy 2025/5/15 mark
      ' 檢查員工編號規則
'      'Add By Sindy 2022/11/28
'      If Left(txtF0101, 1) <> "W" And Left(txtF0101, 1) <> "M" Then
'      '2022/11/28 END
'         If ChkStaffID(txtF0101) Then
'            Call txtF0101_GotFocus
'            Cancel = True
'            Exit Sub
'         End If
'      End If
      '2025/5/15 END
      
      txtF0101_2 = GetStaffName(txtF0101, True)
      If txtF0101_2 = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         Call txtF0101_GotFocus
         Cancel = True
         Exit Sub
      End If

      '只查詢在職或留職停薪人員的資料
      Rs.CursorLocation = adUseClient
      strSql = "select ST01,ST02 " & _
               "From staff " & _
               "where st01='" & txtF0101 & "' " & _
               "and (st04='1' or '04'=(select sc03 from Staff_Change where sc01=ST01 and sc02=(select max(sc02) from Staff_Change where sc01=ST01)))"
      Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If Rs.RecordCount = 0 Then
         MsgBox "此員工已離職！", vbInformation
         Call txtF0101_GotFocus
         Cancel = True
         If Rs.State <> adStateClosed Then Rs.Close
         Exit Sub
      End If
      If Rs.State <> adStateClosed Then Rs.Close
      
      If m_EditMode = 1 Then
         ' 檢查記錄是否已存在
         If IsRecordExist(txtF0101) = True Then
            MsgBox "該筆記錄已存在", vbInformation
            Call txtF0101_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

'Add By Sindy 2023/11/2
Private Sub txtST04_GotFocus()
   txtST04.SelStart = 0
   txtST04.SelLength = Len(txtST04)
   CloseIme
End Sub
Private Sub txtST04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtST04_LostFocus()
Dim s As Integer
   
   If m_EditMode <> 0 And txtF0101 <> "" Then
      If InStr(1, "Yy ", txtST04) = 0 Then
          s = MsgBox("請輸入 Y 或空白!!", , "輸入錯誤")
          txtST04.SetFocus
          Exit Sub
'      'Add By Sindy 2023/12/22
'      Else
'         SetDataListWidth
'         ClearField
'         RefreshRange
'         ShowFirstRecord
'         UpdateToolbarState
'         SetCtrlReadOnly True
'         'OnAction vbKeyF4
'         OnAction vbKeyF10
'      '2023/12/22 END
      End If
   End If
End Sub
'2023/11/2 END

