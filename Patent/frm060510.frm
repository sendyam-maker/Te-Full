VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060510 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "³qª¾§i­ã¥[µù/EmailºûÅ@"
   ClientHeight    =   6840
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8292
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8292
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
            Picture         =   "frm060510.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060510.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060510.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060510.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060510.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060510.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060510.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060510.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060510.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060510.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060510.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Height          =   576
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8292
      _ExtentX        =   14626
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
            Caption         =   "·s¼W"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "­×§ï"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "§R°£"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "¬d¸ß"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "²Ä¤@µ§"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«e¤@µ§"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«á¤@µ§"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "³Ì«áµ§"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "½T©w"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "¨ú®ø"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "µ²§ô"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6090
      Left            =   90
      TabIndex        =   11
      Top             =   720
      Width           =   8115
      _ExtentX        =   14309
      _ExtentY        =   10732
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "³æµ§¸ê®Æ"
      TabPicture(0)   =   "frm060510.frx":20F4
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
      Tab(0).Control(8)=   "Label1(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "textCUID"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtDB(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtDB(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtDB(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDB(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtDB(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtDB(12)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtDB(15)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(12)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(13)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(14)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Cmd1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Frame1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "¦hµ§¬d¸ß"
      TabPicture(1)   =   "frm060510.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdQuery"
      Tab(1).Control(1)=   "GRD1"
      Tab(1).Control(2)=   "Label1(20)"
      Tab(1).Control(3)=   "Label1(19)"
      Tab(1).Control(4)=   "Label1(18)"
      Tab(1).Control(5)=   "txtFM2(4)"
      Tab(1).Control(6)=   "Label1(16)"
      Tab(1).Control(7)=   "Label1(15)"
      Tab(1).Control(8)=   "txtFM2(3)"
      Tab(1).Control(9)=   "Label1(17)"
      Tab(1).Control(10)=   "lblPS"
      Tab(1).Control(11)=   "txtFM2(2)"
      Tab(1).Control(12)=   "Label1(11)"
      Tab(1).Control(13)=   "Label1(9)"
      Tab(1).Control(14)=   "Label1(10)"
      Tab(1).Control(15)=   "txtFM2(0)"
      Tab(1).Control(16)=   "txtFM2(1)"
      Tab(1).Control(17)=   "lblFM2(1)"
      Tab(1).Control(18)=   "lblFM2(2)"
      Tab(1).ControlCount=   19
      Begin VB.CommandButton cmdQuery 
         Caption         =   "¬d¸ß(&Q)"
         Height          =   300
         Left            =   -72270
         TabIndex        =   17
         Top             =   390
         Width           =   885
      End
      Begin VB.Frame Frame1 
         Caption         =   "³qª¾¤uµ{®vEmail³]©w"
         Height          =   1365
         Left            =   120
         TabIndex        =   29
         Top             =   3840
         Width           =   7755
         Begin MSForms.ComboBox Combo1 
            Height          =   300
            Left            =   1140
            TabIndex        =   8
            Top             =   270
            Width           =   6435
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "11351;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "·s²Ó©úÅé-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtDB 
            Height          =   640
            Index           =   14
            Left            =   1140
            TabIndex        =   9
            Top             =   600
            Width           =   5910
            VariousPropertyBits=   -1466941413
            MaxLength       =   500
            ScrollBars      =   2
            Size            =   "10425;1129"
            FontName        =   "·s²Ó©úÅé-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label4 
            Caption         =   "Email¤º¤å¡G"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   31
            Top             =   630
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Email¥D¦®¡G"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   30
            Top             =   300
            Width           =   945
         End
      End
      Begin VB.CommandButton Cmd1 
         Caption         =   "»¡©ú"
         Height          =   255
         Left            =   6768
         TabIndex        =   7
         Top             =   3450
         Width           =   705
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm060510.frx":212C
         Height          =   3255
         Left            =   -74910
         TabIndex        =   12
         Top             =   2760
         Width           =   7905
         _ExtentX        =   13949
         _ExtentY        =   5736
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "¬y¤ô¸¹|³Æµù¤º®e|¥»©Ò®×¸¹|¥N²z¤H|¥Ó½Ð¤H"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "·s²Ó©úÅé-ExtB"
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
      Begin VB.Label Label1 
         Caption         =   "¤é¤å©w½Z"
         Height          =   240
         Index           =   20
         Left            =   -74880
         TabIndex        =   50
         Top             =   1890
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "¥[µù¤º®e¡G"
         Height          =   240
         Index           =   19
         Left            =   -74880
         TabIndex        =   49
         Top             =   2130
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¼Ò½k¤ñ¹ï"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   18
         Left            =   -67620
         TabIndex        =   48
         Top             =   1920
         Visible         =   0   'False
         Width           =   720
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   525
         Index           =   4
         Left            =   -73860
         TabIndex        =   21
         Top             =   1890
         Width           =   6195
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "10936;917"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "­^¤å©w½Z"
         Height          =   240
         Index           =   16
         Left            =   -74880
         TabIndex        =   47
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "¥[µù¤º®e¡G"
         Height          =   240
         Index           =   15
         Left            =   -74880
         TabIndex        =   46
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "¥[µù¤º®e¡G"
         Height          =   240
         Index           =   14
         Left            =   105
         TabIndex        =   45
         Top             =   1830
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "¤é¤å©w½Z"
         Height          =   240
         Index           =   13
         Left            =   105
         TabIndex        =   44
         Top             =   1620
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "¥[µù¤º®e¡G"
         Height          =   240
         Index           =   12
         Left            =   105
         TabIndex        =   43
         Top             =   990
         Width           =   915
      End
      Begin MSForms.TextBox txtDB 
         Height          =   840
         Index           =   15
         Left            =   1050
         TabIndex        =   2
         Top             =   1590
         Width           =   5580
         VariousPropertyBits=   -1466941413
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "9842;1482"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   520
         Index           =   3
         Left            =   -73860
         TabIndex        =   20
         Top             =   1320
         Width           =   6200
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "10936;917"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¼Ò½k¤ñ¹ï"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   17
         Left            =   -67620
         TabIndex        =   42
         Top             =   1350
         Visible         =   0   'False
         Width           =   720
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   12
         Left            =   1770
         TabIndex        =   6
         Top             =   3450
         Width           =   450
         VariousPropertyBits=   671105055
         MaxLength       =   1
         Size            =   "794;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   1
         Left            =   1050
         TabIndex        =   0
         Top             =   390
         Width           =   630
         VariousPropertyBits=   671105055
         MaxLength       =   4
         Size            =   "1111;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   840
         Index           =   2
         Left            =   1050
         TabIndex        =   1
         Top             =   708
         Width           =   5580
         VariousPropertyBits=   -1466941413
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "9842;1482"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   4
         Left            =   1050
         TabIndex        =   4
         Top             =   2790
         Width           =   1170
         VariousPropertyBits=   671105055
         MaxLength       =   8
         Size            =   "2064;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   5
         Left            =   1050
         TabIndex        =   5
         Top             =   3120
         Width           =   1170
         VariousPropertyBits=   671105055
         MaxLength       =   8
         Size            =   "2064;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   3
         Left            =   1050
         TabIndex        =   3
         Top             =   2460
         Width           =   1575
         VariousPropertyBits=   671105055
         MaxLength       =   12
         Size            =   "2778;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   41
         Top             =   3150
         Width           =   5595
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "9878;450"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   40
         Top             =   2820
         Width           =   5595
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "9878;450"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID 
         Height          =   285
         Left            =   90
         TabIndex        =   39
         Top             =   5760
         Width           =   7860
         VariousPropertyBits=   671105055
         Size            =   "13864;503"
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblPS 
         Caption         =   "P.S. ¿é¤J¥»©Ò®×¸¹·|¥t¥~±a¸Ó®×¥N²z¤H©M¥Ó½Ð¤Hªº¨ä¥L³]©w"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74880
         TabIndex        =   38
         Top             =   2520
         Width           =   4845
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   285
         Index           =   2
         Left            =   -73860
         TabIndex        =   19
         Top             =   1020
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1931;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥N²z¤H¡G"
         Height          =   180
         Index           =   11
         Left            =   -74880
         TabIndex        =   37
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H¡G"
         Height          =   180
         Index           =   9
         Left            =   -74880
         TabIndex        =   36
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥»©Ò®×¸¹¡G"
         Height          =   180
         Index           =   10
         Left            =   -74880
         TabIndex        =   35
         Top             =   435
         Width           =   900
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   285
         Index           =   0
         Left            =   -73860
         TabIndex        =   16
         Top             =   390
         Width           =   1515
         VariousPropertyBits=   671105051
         MaxLength       =   12
         Size            =   "2672;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   285
         Index           =   1
         Left            =   -73860
         TabIndex        =   18
         Top             =   705
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1940;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   255
         Index           =   1
         Left            =   -72720
         TabIndex        =   34
         Top             =   720
         Width           =   5595
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "9878;450"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   255
         Index           =   2
         Left            =   -72720
         TabIndex        =   33
         Top             =   1035
         Width           =   5595
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "9878;450"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¡°±Æ°£³]­p®×"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   8
         Left            =   2820
         TabIndex        =   32
         Top             =   2520
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1.»Ýªþ¤¤¤å 2.»Ýªþ­ì¤å(§t­^¤å¤Î¤é¤å) 3.»Ýªþ¤¤¤å¤Î­ì¤å"
         Height          =   180
         Index           =   7
         Left            =   2280
         TabIndex        =   28
         Top             =   3516
         Width           =   4440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¤é¤å©w½Z½Ð¨D¶µ¡G"
         Height          =   180
         Index           =   4
         Left            =   285
         TabIndex        =   27
         Top             =   3510
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¼Ò½k¤ñ¹ï"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   6
         Left            =   6720
         TabIndex        =   26
         Top             =   750
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¬y¤ô¸¹¡G"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   25
         Top             =   435
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "­^¤å©w½Z"
         Height          =   240
         Index           =   1
         Left            =   105
         TabIndex        =   24
         Top             =   750
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥N²z¤H¡G"
         Height          =   180
         Index           =   2
         Left            =   315
         TabIndex        =   23
         Top             =   2850
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H¡G"
         Height          =   180
         Index           =   3
         Left            =   315
         TabIndex        =   22
         Top             =   3180
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥»©Ò®×¸¹¡G"
         Height          =   180
         Index           =   5
         Left            =   105
         TabIndex        =   15
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ª`·N¡G1.¥N²z¤H/¥Ó½Ð¤H¥i¿é¤J6½X©Î8½X¡A6½X¥Nªí§tÃö«Y¥ø·~¡C112/1/9¨ú®ø"
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
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
         Left            =   1110
         TabIndex        =   14
         Top             =   5280
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "¡@¡@¡@2.¥N²z¤H/¥Ó½Ð¤HµL½×6½X©Î8½X§¡¥]§t§ó¦W«e½s¸¹¡C112/1/9¨ú®ø"
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
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
         Left            =   1110
         TabIndex        =   13
         Top             =   5520
         Visible         =   0   'False
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frm060510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/02/18 ¾ã¦X¯S®í³ÆµùºûÅ@¡G1.±N²{¦³¸ê®Æ6½XY/X½s¸¹¸É¨¬¬°8½X¡F2.¦b¿é¤JY/X½s¸¹­Y¬°6½X¡A²Î¤@¸É¨¬¬°8½X¡C
'Memo by Morgan 2022/10/26 ¤é¤å¤w§ï§ìTable
'Memo by Lydia 2021/11/01 §ï¦¨Form2.0 ; GRD1§ï¦r«¬=·s²Ó©úÅé-ExtB¡BtxtDB(index)¡BLabel2(index)¡BtextCUID¡BtxtFM2(index)¡BlblFM2(index)
'Memo by Lydia 2021/11/01 µe­±­¶ÅÒ§ï¦¨¡u³æµ§¸ê®Æ¡v©M¡u¦hµ§¬d¸ß¡v¡G¤W¤è¤u¨ã¦Cªº¡u¬d¸ß¡v±a¥X²Ä¤@µ§²Å¦Xªº¸ê®Æ¡A¦b¦hµ§¬d¸ßªº­¶ÅÒ¥i¥H¿é¤J±ø¥ó¶i¦æ¬d¸ß¡A¨Ã¥B¦b¤U¤èªºGrid§e²{¦hµ§¸ê®Æ¡C
'Memo by Lydia 2021/02/02 §ó¦W¬°¡u³qª¾§i­ã¥[µù/EmailºûÅ@¡v
'Created by Lydia 2019/03/11 ·s¼W-³qª¾§i­ã¥[µùºûÅ@(ApprovalPS)
Option Explicit

Dim m_EditMode As Integer '0:ÂsÄý 1:·s¼W 2:­×§ï 3:§R°£ 4:¬d¸ß
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_blnColOrderAsc As Boolean 'Added by Lydia 2021/11/11 Äæ¦ì¸ê®Æ¥Ñ¤p¨ì¤j±Æ§Ç
Dim oText As Control, oLabel As Control
Dim stCon As String, stSQL As String, intR As Integer
Dim rsRead As New ADODB.Recordset
Dim stLanPA As String, stLanY As String, stLanX As String 'Added by Lydia 2020/12/30 ­Ó®× / Y/ X½s¸¹ªº©w½Z»y¤å

'Added by Lydia 2021/11/01
Private Sub cmdQuery_Click()
   
   stCon = ""
   If txtFM2(0) <> "" Then
      If Trim(txtFM2(1).Tag & txtFM2(2).Tag) = "" Then
          stCon = stCon & " and aps03='" & txtFM2(0) & "'"
      Else
          '¥t¥~§ì¥»©Ò®×¸¹ªº¬ÛÃöY½s¸¹¡BX½s¸¹±ø¥ó
          stCon = stCon & " and (aps03='" & txtFM2(0) & "'"
          If txtFM2(1).Tag <> "" Then stCon = stCon & " or instr(" & CNULL(txtFM2(1).Tag) & ", aps04) > 0 "
          If txtFM2(2).Tag <> "" Then stCon = stCon & " or instr(" & CNULL(txtFM2(2).Tag) & ", aps05) > 0 "
          stCon = stCon & ") "
      End If
   Else
      txtFM2(1).Tag = "": txtFM2(2).Tag = ""   '²MªÅ¥»©Ò®×¸¹ªº¬ÛÃöY½s¸¹¡BX½s¸¹±ø¥ó
   End If
   If txtFM2(1) <> "" Then
      stCon = stCon & " and aps04 like '" & txtFM2(1) & "%'"
   End If
   If txtFM2(2) <> "" Then
      stCon = stCon & " and aps05 like '" & txtFM2(2) & "%'"
   End If
   'Added by Lydia 2022/10/03 ¼W¥["­^¤å©w½Z¥[µù¤º®e"¬d¸ß
   If txtFM2(3) <> "" Then
       stCon = stCon & " and upper(aps02) like '%" & ChgSQL(UCase(txtFM2(3))) & "%' "
   End If
   'end 2022/10/03
   'Added by Lydia 2022/10/05 ¼W¥["¤é¤å©w½Z¥[µù¤º®e"¬d¸ß
   If txtFM2(4) <> "" Then
       stCon = stCon & " and upper(aps15) like '%" & ChgSQL(UCase(txtFM2(4))) & "%' "
   End If
   'end 2022/10/05
   
   'Modified by Lydia 2022/10/05 +APS15
   stSQL = "SELECT APS01,APS02,APS15,APS03,APS04,APS05,APS12,APS13,APS14 FROM APPROVALPS WHERE 1=1 " & stCon
   stSQL = stSQL & " ORDER BY aps01"
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Memo by Lydia 2021/11/01 ­ìµ{¦¡·h¨ìForm_KeyUp

End Sub

'Added by Lydia 2021/11/01
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Memo by Lydia 2021/11/01 ±qForm_KeyDown·h¨Ó
   Screen.MousePointer = vbHourglass
   Select Case KeyCode
      Case vbKeyF2 '·s¼W
         KeyCode = 0: Action 1
      Case vbKeyF3 '­×§ï
         KeyCode = 0: Action 2
      Case vbKeyF4: '¬d¸ß
         KeyCode = 0: Action 4
      Case vbKeyF5 '§R°£
         KeyCode = 0: Action 3
      Case vbKeyHome '²Ä¤@µ§
         KeyCode = 0: Action 6
      Case vbKeyPageUp '¤W¤@µ§
         KeyCode = 0: Action 7
      Case vbKeyPageDown '¤U¤@µ§
         KeyCode = 0: Action 8
      Case vbKeyEnd: '³Ì«áµ§
         KeyCode = 0: Action 9
      'Modified by Lydia 2021/11/22 Lydia 2021/11/22 ¨ú®ø¥HENTER±±¨î¬°´«¦æªº¥\¯à (Form2.0­×§ï¤§ºûÅ@¸ê®Æ¥\¯àToolbar¤§­×§ï²Î¤@)
'      Case vbKeyF9, vbKeyReturn '½T©w
      Case vbKeyF9 '½T©w
         KeyCode = 0: Action 11
    
      Case vbKeyF10 '¨ú®ø
         KeyCode = 0: Action 12
      Case vbKeyEscape 'µ²§ô
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            KeyCode = 0: Action 14
         End If
   End Select
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()
   '¨ú±o¨Ï¥ÎªÌ°õ¦æ¦U¶µ¥\¯àªºÅv­­
   m_bInsert = IsUserHasRightOfFunction("frm060510", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm060510", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm060510", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm060510", strFind, False)
  
   MoveFormToCenter Me
   
   'Added by Lydia 2021/11/01
   For Each oLabel In lblFM2
       oLabel.BackColor = &H8000000F
   Next
   Call SetGrd(True)
   'end 2021/11/01
   
   textCUID.BackColor = &H8000000F
   Action 6 '¹w³]²Ä¤@µ§
   UpdateToolbarState
   
   'Added by Lydia 2021/02/02 ¹w³]¤U©Ô¿ï³æ
   Combo1.Clear
   Combo1.AddItem "½Ð´£¨Ñ³Ì·sª©¥»¤§­ì¤å½Ð¨D¶µWORDÀÉ", 0
   Combo1.AddItem "½Ð´£¨Ñ¤w­ã½Ð¨D¶µªº¤¤¤å¥»+­^¤å¥»WORDÀÉ", 1
   
   Me.SSTab1.Tab = 1 'Added by Lydia 2021/11/01 §ï±q¦hµ§¬d¸ß­¶ÅÒ¶}©l
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060510 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
   Action Button.Index
   Screen.MousePointer = vbDefault
End Sub

'¨Ì·ÓÅv­­³]©w¨ä¤u¨ã¦Cªº«ö¯Ãª¬ºA
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' µL¥ô¦ó°Ê§@
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
      
      Case 1, 2, 3, 4 'ºûÅ@
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
   Case 0 'ÂsÄý
      For Each oText In txtDB
         oText.Locked = True
      Next
      Combo1.Enabled = False 'Added by Lydia 2021/02/02
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
      Combo1.Enabled = True 'Added by Lydia 2021/02/02
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
      Case 1 '«ö¤U·s¼W
        m_EditMode = 1
        FormReset
        
      Case 2 '«ö¤U­×§ï
         m_EditMode = 2

      Case 3 '«ö¤U§R°£
         If txtDB(1).Text = "" Then
             MsgBox "µL¸ê®Æ¥i§R°£!!!", vbExclamation + vbOKOnly
             Exit Sub
         End If

         If DelMsg() = True Then
            If FormDelete() = False Then
               MsgBox "§R°£¥¢±Ñ!", vbCritical
               Exit Sub
            '§R°£«á²¾¨ì³Ì¥½µ§
            Else
               ShowRecord 3
            End If
         End If

      Case 4 '«ö¤U¬d¸ß
         FormReset
         m_EditMode = 4
         txtDB(1).Enabled = True
         txtDB(1).SetFocus
         Label1(6).Visible = True
         
      Case 6 '²Ä¤@µ§
         ShowRecord 0
      Case 7 '«e¤@µ§
         ShowRecord 1
      Case 8 '«á¤@µ§
         ShowRecord 2
      Case 9 '³Ì«áµ§
         ShowRecord 3
      Case 11 '«ö¤U½T©w
         'Added by Lydia 2019/05/20 ¨Ï¥ÎªÌ¿é¤J®×¸¹«á¡Aª½±µ«öEnterµLªkÄ²µoÀË¬d®×¸¹¤§¥\¯à (by Winfrey)
         If Val(m_EditMode) > 0 And Trim(txtDB(3)) <> "" And ((Left(Trim(txtDB(3)), 1) = "P" And Len(Trim(txtDB(3))) < 10) Or (Left(Trim(txtDB(3)), 3) = "FCP" And Len(Trim(txtDB(3))) < 12)) Then
             Call txtDB_Validate(3, bCancel)
             If bCancel = True Then
                 Exit Sub
             End If
         End If
         
         Select Case m_EditMode
            '·s¼W,­×§ï
            Case 1, 2
               'Modified by Lydia 2021/11/01 ·s¼W,­×§ï³£­n§PÂ_
               'If m_EditMode = 1 Then
               '   If RecIsExist = True Then Exit Sub
               'End If
               If RecIsExist = True Then Exit Sub
               
               If TxtValidate = False Then
                  Exit Sub
               Else
                  If FormSave() = False Then
                     MsgBox "¦sÀÉ¥¢±Ñ!", vbCritical
                     Exit Sub
                  Else
                     strKind = m_EditMode 'Added by Lydia 2021/11/01 °O¿ý·s¼W¼Ò¦¡
                     m_EditMode = 0
                     If txtDB(1) = "" Then
                        ShowRecord 3
                     Else
                        ReadData txtDB(1)
                     End If
                  End If
                    'Added by Lydia 2021/11/01 ¦b·s¼W¦sÀÉ«á¦Û°Ê±a¤J¦hµ§¬d¸ßÅã¥Ü¥»¦¸·s¼W°O¿ý
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
            '¬d¸ß
            Case 4
               If ReadData(txtDB(1)) = False Then
                  MsgBox "µL¸ê®Æ!", vbExclamation
                  Exit Sub
               Else
                  m_EditMode = 0
               End If
         End Select
      Case 12 '«ö¤U¨ú®ø
         m_EditMode = 0
         txtDB(1) = txtDB(1).Tag
         If txtDB(1) <> "" Then
            If ReadData(txtDB(1)) = False Then
               ShowRecord 3
            End If
         End If
      Case 14 'µ²§ô
         Unload Me
         Exit Sub
   End Select
   UpdateToolbarState
   TxtLock
   Exit Sub
   
ErrHand:
   ShowMsg "¿ù»~ : " & Err.Description
End Sub

' Åã¥Ü¸ê®Æ
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
 Dim stKey As String
    
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case p_iWay
      Case 0 '²Ä¤@µ§
         strExc(0) = "SELECT nvl(min(APS01),0) FROM ApprovalPS"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKey = RsTemp.Fields(0)
            End If
         End If
         
      Case 1 '«e¤@µ§
         strExc(0) = "SELECT nvl(max(APS01),0) FROM ApprovalPS where APS01<" & txtDB(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 6
            Else
               stKey = RsTemp.Fields(0)
            End If
         End If
         
      Case 2 '«á¤@µ§
         strExc(0) = "SELECT nvl(min(APS01),0) FROM ApprovalPS where APS01>" & txtDB(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 7
            Else
               stKey = RsTemp.Fields(0)
            End If
         End If
         
      Case 3 '³Ì«áµ§
         strExc(0) = "SELECT nvl(max(APS01),0) FROM ApprovalPS"
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
   MsgBox "¿ù»~ : " & Err.Description, vbCritical
End Function

Private Function ReadData(Optional ByVal pKey As String) As Boolean

   stCon = ""
   '³æµ§
   If pKey <> "" Then
      stCon = " and APS01=" & pKey
   '¦hµ§
   Else
      Call SetGrd(True) 'Added by Lydia 2021/02/02 ²MªÅ
      If txtDB(2) <> "" Then
         stCon = stCon & " and APS02 like '%" & ChgSQL(txtDB(2)) & "%'"
      End If
      If txtDB(3) <> "" Then
         stCon = stCon & " and APS03='" & txtDB(3) & "'"
      End If
      If txtDB(4) <> "" Then
         stCon = stCon & " and APS04 like '" & txtDB(4) & "%'"
      End If
      If txtDB(5) <> "" Then
         stCon = stCon & " and APS05 like '" & txtDB(5) & "%'"
      End If
      'Added by Lydia 2020/12/30
      If txtDB(12) <> "" Then
           stCon = stCon & " and APS12='" & txtDB(12) & "'"
      End If
      'Added by Lydia 2021/02/02
      If Trim(Combo1.Text) <> "" Then
          stCon = stCon & " and APS13 like '%" & Trim(Combo1.Text) & "%'"
      End If
      If txtDB(14) <> "" Then
           stCon = stCon & " and APS14 like '%" & txtDB(14) & "%'"
      End If
      'end 2021/02/02
      'Added by Lydia 2022/10/05 ¤é¤å©w½Z¥[µù
      If txtDB(15) <> "" Then
         stCon = stCon & " and APS15 like '%" & ChgSQL(txtDB(15)) & "%'"
      End If
      'end 2022/10/05
   End If
   
   FormReset

   strExc(0) = "select * from ApprovalPS where 1=1 " & stCon & " order by APS01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If m_EditMode = 4 Then
         'Modified by Lydia 2021/02/02
         'Set GRD1.Recordset = RsTemp.Clone
         'GRD1.FormatString = GRD1.FormatString
         'GRD1.ColWidth(1) = 2775
         'GRD1.ColWidth(2) = 1290
         'GRD1.ColWidth(3) = 1500
         'GRD1.ColWidth(4) = 1500
         'For intI = 5 To GRD1.Cols - 1
         '   GRD1.ColWidth(intI) = 0
         'Next
         'Modified by Lydia 2021/11/01 §ï¦¨³æµ§¬d¸ß
         'Set GRD1.Recordset = RsTemp
         'Call SetGrd
         ''end 2021/02/02
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
        oText = "" & .Fields("APS" & Format(oText.Index, "00"))
        oText.Tag = oText.Text 'Added by Lydia 2020/12/30 ¼È¦s
     Next
     'Added by Lydia 2021/02/02 Email¥D¦®
     Combo1.Text = "" & rsQuery.Fields("APS13")
   End With
   UpdateCUID rsQuery
   
   'txtDB(1).Tag = txtDB(1) 'Remove by Lydia 2020/12/30
   If txtDB(4) <> "" Then txtDB_Validate 4, False
   If txtDB(5) <> "" Then txtDB_Validate 5, False
   If txtDB(3) <> "" Then txtDB_Validate 3, False 'Added by Lydia 2020/12/30
End Sub

' §ó·s Create ¤Î Update ªº¤H
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   If IsNull(rsSrcTmp.Fields("APS06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("APS06")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("APS06"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("APS07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("APS07")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("APS07"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("APS08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("APS08")) = False Then
         strTemp = rsSrcTmp.Fields("APS08")
         strCTime = Format(strTemp, "00:00:00")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("APS09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("APS09")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("APS09"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("APS10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("APS10")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("APS10"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("APS11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("APS11")) = False Then
         strTemp = rsSrcTmp.Fields("APS11")
         strUTime = Format(strTemp, "00:00:00")
      End If
   End If
   
   ' ³]©wCUID¤¤ªº¤å¦r
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
   
   'Added by Lydia 2021/02/02
   Combo1.Text = ""
   Combo1.Tag = ""
End Sub

Private Sub txtDB_GotFocus(Index As Integer)
   TextInverse txtDB(Index)
   If Index = 2 Then
      OpenIme
   Else
      CloseIme
   End If
End Sub

'Modified by Lydia 2021/11/01 §ï¦¨Form 2.0
'Private Sub txtDB_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtDB_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index <> 2 Then
      KeyAscii = UpperCase(KeyAscii)
      'Added by Lydia 2020/12/30 ¤é¤å©w½Z½Ð¨D¶µ
      If Index = 12 Then
         'Modified by Lydia 2024/03/18 +½Ð¨D¶µ3
         'If (KeyAscii < 49 Or KeyAscii > 50) And KeyAscii <> 8 Then
         If (KeyAscii < 49 Or KeyAscii > 51) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      End If
      'end 2020/12/30
   End If
End Sub

Private Sub txtDB_Validate(Index As Integer, Cancel As Boolean)
   Dim strCusTemp As String, strTemp As String
   
   Select Case Index
   Case 3 '¥»©Ò®×¸¹
      stLanPA = "" 'Added by Lydia 2020/12/30
      If txtDB(Index) <> "" Then
         'Modifie by Lydia 2021/09/03 +PA08
         strExc(0) = "select PA01||PA02||PA03||PA04 as CaseNo,PA01,PA02,PA03,PA04,PA08 from patent where " & ChgPatent(txtDB(Index))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            If m_EditMode <> 0 Then 'Added by Lydia 2021/11/01 ±Æ°£«D½s¿è¼Ò¦¡: ¦]¬°®×¸¹¦³¥i¯à§R°£
                 MsgBox "¥»©Ò®×¸¹¿é¤J¿ù»~!", vbExclamation
                 Cancel = True
            End If 'Added by Lydia 2021/11/01
            'If m_EditMode <> 0 Then Cancel = True 'Remove by Lydia 2021/11/01
         Else
            'Added by Lydia 2021/09/03 ±Æ°£³]­p®×
            If "" & RsTemp.Fields("PA08") = "3" Then
                If m_EditMode <> 0 Then  'Added by Lydia 2021/11/01 ±Æ°£«D½s¿è¼Ò¦¡: ¦]¬°®×¸¹¦³¥i¯à§R°£
                    MsgBox "¤£¥i³]©w¬°³]­p®×!", vbExclamation
                    Cancel = True
                End If 'Added by Lydia 2021/11/01
                'If m_EditMode <> 0 Then Cancel = True 'Remove by Lydia 2021/11/01
                Exit Sub
            End If
            'end 2021/09/03
            txtDB(Index) = "" & RsTemp.Fields("CaseNo")
            'Added by Lydia 2020/12/30 »P¹ï¥~³qª¾-®Ö­ã¨ç(frm060317_1)¥Î¬Û¦P¼Ò²Õ§PÂ_»y¤å
            stLanPA = GetLetterLanguage("" & RsTemp.Fields("PA01"), "" & RsTemp.Fields("PA02"), "" & RsTemp.Fields("PA03"), "" & RsTemp.Fields("PA04"))
         End If
      End If
   Case 4 '¥N²z¤H
      Label2(1).Caption = ""
      stLanY = "" 'Added by Lydia 2020/12/30
      If txtDB(Index) <> "" Then
         'Added by Lydia 2022/10/05
         If Left(txtDB(Index), 1) <> "Y" Then
            MsgBox "¥N²z¤H½s¸¹¥u¥i¿é¤JY½s¸¹¡I", vbCritical
            If m_EditMode <> 0 Then Cancel = True
         Else
         'end 2022/10/05
            'Modified by Morgan 2019/7/25 ¥[½X¼ÆÀË¬d
            If Len(txtDB(Index)) = 6 Or Len(txtDB(Index)) = 8 Then
               strCusTemp = ChangeCustomerL(txtDB(Index))
               'Modified by Lydia 2020/12/30 ¸Ó¼Ò²Õ¦³©w½Z»y¤å,¦ý¬O¼u°T®§¬°"©¹¨Ó¹ï¶H"(§tX,Y,R)
               'If ClsPDGetAgent(strCusTemp, strTemp) Then
               If PUB_GetCustData(strCusTemp, strTemp, , stLanY) = True Then
                  Label2(1).Caption = strTemp
                  'Added by Lydia 2023/02/18 ¾ã¦X¯S®í³ÆµùºûÅ@¡G¦b¿é¤JY/X½s¸¹­Y¬°6½X¡A²Î¤@¸É¨¬¬°8½X¡C
                  If m_EditMode <> 0 Then
                     txtDB(Index) = Left(ChangeCustomerL(txtDB(Index)), 8)
                  End If
                  'end 2023/02/18
               Else
                  'MsgBox "¥N²z¤H½s¸¹¿é¤J¿ù»~¡I", vbCritical  'Remove by Lydia 2021/11/01 ¼Ò²Õ¤w¼u°T®§
                  If m_EditMode <> 0 Then Cancel = True
               End If
            Else
               MsgBox "¥N²z¤H½s¸¹¥u¥i¿é¤J6½X©Î8½X¡I", vbCritical
               If m_EditMode <> 0 Then Cancel = True
            End If
         End If 'Added by Lydia 2022/10/05
      End If
   Case 5 '¥Ó½Ð¤H
      Label2(2).Caption = ""
      stLanX = "" 'Added by Lydia 2020/12/30
      If txtDB(Index) <> "" Then
         'Added by Lydia 2022/10/05
         If Left(txtDB(Index), 1) <> "X" Then
            MsgBox "«È¤á½s¸¹¥u¥i¿é¤JX½s¸¹¡I", vbCritical
            If m_EditMode <> 0 Then Cancel = True
         Else
         'end 2022/10/05
            'Modified by Morgan 2019/7/25 ¥[½X¼ÆÀË¬d
            If Len(txtDB(Index)) = 6 Or Len(txtDB(Index)) = 8 Then
               strCusTemp = ChangeCustomerL(txtDB(Index))
               'Modified by Lydia 2020/12/30
               'If ClsPDGetCustomer(strCusTemp, strTemp) Then
               If PUB_GetCustData(strCusTemp, strTemp, , stLanX) = True Then
                  Label2(2).Caption = strTemp
                  'Added by Lydia 2023/02/18 ¾ã¦X¯S®í³ÆµùºûÅ@¡G¦b¿é¤JY/X½s¸¹­Y¬°6½X¡A²Î¤@¸É¨¬¬°8½X¡C
                  If m_EditMode <> 0 Then
                     txtDB(Index) = Left(ChangeCustomerL(txtDB(Index)), 8)
                  End If
                  'end 2023/02/18
               Else
                  'MsgBox "«È¤á½s¸¹¿é¤J¿ù»~¡I", vbCritical 'Remove by Lydia 2021/11/01 ¼Ò²Õ¤w¼u°T®§
                  If m_EditMode <> 0 Then Cancel = True
               End If
            Else
               MsgBox "«È¤á½s¸¹¥u¥i¿é¤J6½X©Î8½X¡I", vbCritical
               If m_EditMode <> 0 Then Cancel = True
            End If
            If Cancel = True Then Label2(2).Caption = "" 'Added by Lydia 2022/10/05
         End If 'Added by Lydia 2022/10/05
      End If
   End Select
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean, idx As Integer
      
   'Modified by Lydia 2022/10/05
   'If txtDB(2) = "" Then
   If txtDB(2) & txtDB(15) = "" Then
      If MsgBox("³Æµù¤º®e¥¼¿é¤J¡A¬O§_½T©w¬°¤£­n¹w³]³qª¾§i­ã¥[µù¡H", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         txtDB(2).SetFocus
         Exit Function
      End If
   End If
   
   If txtDB(3) & txtDB(4) & txtDB(5) = "" Then
      MsgBox "½Ð¿é¤J¥»©Ò®×¸¹¡B¥N²z¤H©Î¥Ó½Ð¤H¡I", vbExclamation
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
   
   'Added by Lydia 2020/12/30 ¤é¤å©w½Z½Ð¨D¶µ¡GÀË¬d
   'Modified by Lydia 2022/10/05 +¤é¤å©w½Z¥[µù  txtDB(15)
   If txtDB(12) & txtDB(15) <> "" Then
      If Left(Trim(stLanPA & stLanY & stLanX), 1) <> "3" Then
        strExc(1) = ""
        If Trim(stLanPA & stLanY & stLanX) = "" Then
            If txtDB(3) <> "" Then strExc(1) = strExc(1) & "¥»©Ò®×¸¹¡G" & txtDB(3) & vbCrLf
            If txtDB(4) <> "" Then strExc(1) = strExc(1) & "¥N²z¤H¡G" & txtDB(4) & vbCrLf
            If txtDB(5) <> "" Then strExc(1) = strExc(1) & "¥Ó½Ð¤H¡G" & txtDB(5) & vbCrLf
        Else
            strExc(1) = strExc(1) & IIf(stLanPA <> "", "¥»©Ò®×¸¹¡G" & txtDB(3), IIf(stLanY <> "", "¥N²z¤H¡G" & txtDB(4), "¥Ó½Ð¤H¡G" & txtDB(5))) & vbCrLf
        End If
        If strExc(1) <> "" Then
            MsgBox strExc(1) & "¤£¬O¤é¤å©w½Z¤£¥i³]©w¡I", vbExclamation, "ÀË¬d"
            Exit Function
        End If
      End If
   End If
   'end 2020/12/30
   
   'Added by Lydia 2021/02/02 ¼W¥[³qª¾¤uµ{®vEmail³]©w©Ò»Ýªº¡uEmail¥D¦®¡v¡B¡uEmail¤º¤å¡v
   If Trim(Combo1.Text) <> "" And Trim(txtDB(14).Text) = "" Then
        MsgBox "½Ð¤@¨Ö¿é¤JEmail¤º¤å¡I", vbExclamation, "ÀË¬d"
        txtDB(14).SetFocus
        txtDB_GotFocus 14
        Exit Function
   ElseIf Trim(Combo1.Text) = "" And Trim(txtDB(14).Text) <> "" Then
        MsgBox "½Ð¤@¨Ö¿é¤JEmail¥D¦®¡I", vbExclamation, "ÀË¬d"
        Combo1.SetFocus
        Exit Function
   End If
   If Len(Combo1.Text) > 100 Then
       MsgBox "Email¥D¦®¶W¹L100­Ó¦r¡I", vbExclamation, "ÀË¬d"
       Combo1.SetFocus
       Exit Function
   End If
   'Email¤º¤å¤w³]©wMaxLength0
   'end 2021/02/02
   
   'Added by Lydia 2021/11/01 ÀË¬dµe­±ªº TextBox, ComboBox ¬O§_§t¦³Unicode¤å¦r
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
       Exit Function
   End If
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
'On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
   
   'Create©MUpdate¥ÑTrigger³]©w
   If m_EditMode = 1 Then
      'Modified by Lydia 2020/12/30 +APS12
      'Modified by Lydia 2021/02/02 +APS13,APS14
      'Modified by Lydia 2022/10/05 +APS15 ; ®³±¼Trigger
      'strSql = "insert into ApprovalPS(APS01,APS02,APS03,APS04,APS05,APS12,APS13,APS14,APS15) " & _
                   "VALUES ('" & Pub_GetDefColMaxNo("ApprovalPS", "APS01") & "'," & CNULL(ChgSQL(txtDB(2))) & "," & CNULL(txtDB(3)) & " ," & CNULL(txtDB(4)) & " ," & CNULL(txtDB(5)) & " ," & CNULL(txtDB(12)) & _
                   "," & CNULL(ChgSQL(Trim(Combo1.Text))) & "," & CNULL(ChgSQL(txtDB(14))) & "," & CNULL(ChgSQL(txtDB(15))) & " ) "
      strSql = "insert into ApprovalPS(APS01,APS02,APS03,APS04,APS05,APS06,APS07,APS08,APS12,APS13,APS14,APS15) " & _
                   "VALUES ('" & Pub_GetDefColMaxNo("ApprovalPS", "APS01") & "'," & CNULL(ChgSQL(txtDB(2))) & "," & CNULL(txtDB(3)) & " ," & CNULL(txtDB(4)) & " ," & CNULL(txtDB(5)) & _
                   "," & CNULL(strUserNum) & ",to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'), " & CNULL(txtDB(12)) & _
                   "," & CNULL(ChgSQL(Trim(Combo1.Text))) & "," & CNULL(ChgSQL(txtDB(14))) & "," & CNULL(ChgSQL(txtDB(15))) & " ) "
   Else
      'Modified by Lydia 2020/12/30 +APS12
      'Modified by Lydia 2021/02/02 +APS13,APS14
      'Modified by Lydia 2022/10/05 +APS15 ; ®³±¼Trigger
      'strSql = "update ApprovalPS set APS02=" & CNULL(ChgSQL(txtDB(2))) & " ,APS03=" & CNULL(txtDB(3)) & _
         ",APS04=" & CNULL(txtDB(4)) & " ,APS05=" & CNULL(txtDB(5)) & ",APS12=" & CNULL(txtDB(12)) & _
         ",APS13=" & CNULL(ChgSQL(Trim(Combo1.Text))) & ",APS14=" & CNULL(txtDB(14)) & ", APS15=" & CNULL(ChgSQL(txtDB(15))) & _
         " where APS01=" & txtDB(1)
      strSql = "update ApprovalPS set APS02=" & CNULL(ChgSQL(txtDB(2))) & " ,APS03=" & CNULL(txtDB(3)) & _
         ",APS04=" & CNULL(txtDB(4)) & " ,APS05=" & CNULL(txtDB(5)) & ",APS12=" & CNULL(txtDB(12)) & _
         ",APS13=" & CNULL(ChgSQL(Trim(Combo1.Text))) & ",APS14=" & CNULL(txtDB(14)) & ", APS15=" & CNULL(ChgSQL(txtDB(15))) & _
         ",APS09=" & CNULL(strUserNum) & ",APS10=to_char(sysdate,'yyyymmdd'),APS11=to_char(sysdate,'hh24miss')" & _
         " where APS01=" & txtDB(1)
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
   'Modified by Lydia 2021/11/11 ¦]¬°¥[¤WGrid±Æ§Ç,©Ò¥H§ï¼gªk
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

'Added by Lydia 2021/11/11
Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If InStr("¬y¤ô¸¹,½Ð¨D¶µ", Me.GRD1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '¼Æ­Èª@¾­
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '¼Æ­È­°¾­
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '¦r¦êª@¾­
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '¦r¦ê­°¾­
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Function FormDelete() As Boolean
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   strSql = "delete from ApprovalPS where APS01=" & txtDB(1)
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
   stCon = stCon & "and APS03='" & Trim(txtDB(3)) & "' "
End If
If Trim(txtDB(4)) <> "" Then
   'Modified by Lydia 2019/07/31 §ï¦¨9½X§PÂ_; ¦]¬°µLªk¥ý¿é¤J8½X«á¦A¿é¤J6½X
   'stcon = stcon & "and instr(APS04,'" & Trim(txtDB(4)) & "') > 0 "
    stCon = stCon & "and aps04='" & Trim(txtDB(4)) & "' "
   '°Ï§O¥u¦³¥N²z¤H©Î«È¤áªº±ø¥ó
   If Trim(txtDB(5)) = "" Then stCon = stCon & "and APS05 is null "
End If
If Trim(txtDB(5)) <> "" Then
   'Modified by Lydia 2019/07/31 §ï¦¨9½X§PÂ_; ¦]¬°µLªk¥ý¿é¤J8½X«á¦A¿é¤J6½X
   'stcon = stcon & "and instr(APS05,'" & Trim(txtDB(5)) & "') > 0 "
   stCon = stCon & "and aps05='" & Trim(txtDB(5)) & "' "
   '°Ï§O¥u¦³¥N²z¤H©Î«È¤áªº±ø¥ó
   If Trim(txtDB(4)) = "" Then stCon = stCon & "and APS04 is null "
End If

If Left(stCon, 3) = "and" Then
   stCon = Mid(stCon, 4, Len(stCon) - 4)
ElseIf stCon = "" Then
   Exit Function
End If

   stSQL = " select * from ApprovalPS where " & stCon
   intR = 1
   Set rsRead = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      'Added by Lydia 2021/11/01 ±Æ°£²{¦b­×§ïªº°O¿ý
      If rsRead.RecordCount = 1 And Trim(rsRead.Fields("APS01")) = Trim(txtDB(1)) Then
         RecIsExist = False
      Else
      'end 2021/11/01
         RecIsExist = True
         MsgBox "¤w¦s¦b¦P¼Ë±ø¥óªº°O¿ý(¬y¤ô¸¹ " & rsRead(0) & " )¡A½Ð¥ý¬d¸ß!!", vbCritical
      End If 'Added by Lydia 2021/11/01
   Else
      RecIsExist = False
   End If
   Set rsRead = Nothing
   
End Function

'Added by Lydia 2020/12/30 ¤é¤å©w½Z½Ð¨D¶µªº»¡©ú
Private Sub Cmd1_Click()
    'Modified by Morgan 2022/10/26
    'strExc(1) = "1.  »Ýªþ¤¤¤å½Ð¨D¶µ¡G¸m´«¬°¡uþ÷þàÇeþêþùÇV¡N³\¥i¬d©w®ÑÇU‡Àþê¤ÎÇZþðÇU©M“Õ¤å¡N¨ÃÇZÇR³\¥iþèÇsþò¤¤üÂ»yÇ«Çè¡ÐÇÜÇy²K¥IþÝ°eÇæ­PþêÇeþìÇUþú¡Nþç¬d’ÚÇUµ{¡N©yþêþâþÝÄ@Æê­PþêÇeþì¡C¡v"
    'strExc(1) = strExc(1) & vbCrLf & vbCrLf & "2.  »Ýªþ­ì¤å(§t­^¤å¤Î¤é¤å)½Ð¨D¶µ¡G¸m´«¬°¡uþ÷þàÇeþêþùÇV¡N³\¥i¬d©w®ÑÇU‡Àþê¤ÎÇZþðÇU©M“Õ¤å¡N¨ÃÇZÇR³\¥iþèÇsþòÇ«Çè¡ÐÇÜÇU­ì¤åÇy²K¥IþÝ°eÇæ­PþêÇeþìÇUþú¡Nþç¬d’ÚÇUµ{¡N©yþêþâþÝÄ@Æê­PþêÇeþì¡C¡v"
    'MsgBox strExc(1), vbInformation + vbOKOnly, "¤é¤å©w½Z½Ð¨D¶µªº»¡©ú"
    strExc(1) = PUB_GetUniText(Me.Name, "¤é¤å©w½Z½Ð¨D¶µªº»¡©ú")
    MsgBoxU strExc(1), vbInformation + vbOKOnly, "¤é¤å©w½Z½Ð¨D¶µªº»¡©ú"
    'end 2022/10/26
End Sub

'Added by Lydia 2021/02/02
Private Sub Combo1_LostFocus()

   If Combo1.Tag <> Combo1.Text And (Combo1.Text = Combo1.List(0) Or Combo1.Text = Combo1.List(1)) Then
        '­Y¿é¤J¬°¹w³]¥D¦®¡A¥ý¹w³]¬Û¦PEmail¤º¤å¡C
        txtDB(14).Text = "¦¹®×¤w®Ö­ã¡A" & Combo1.Text & "¡C"
   ElseIf Combo1.Tag <> "" And Combo1.Text = "" Then
        txtDB(14).Text = ""
   End If
   Combo1.Tag = Combo1.Text
End Sub

'Added by Lydia 2021/02/02
Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer, iR As Integer
   
   'Modified by Lydia 2021/11/01 ®³±¼APS06~APS11
    'arrGridHeadText = Array("¬y¤ô¸¹", "³Æµù¤º®e", "¥»©Ò®×¸¹", "¥N²z¤H", "¥Ó½Ð¤H", _
                                        "APS06", "APS07", "APS08", "APS09", "APS10", "APS11", _
                                        "½Ð¨D¶µ", "Email¥D¦®", "Email¤º¤å")
    'arrGridHeadWidth = Array(800, 1200, 1200, 1000, 1000, 0, 0, 0, 0, 0, 0, 800, 1000, 1000)
    'Modified by Lydia 2022/10/05 +APS15 ¤é¤å©w½Z¥[µù¤º®e
    'arrGridHeadText = Array("¬y¤ô¸¹", "³Æµù¤º®e", "¥»©Ò®×¸¹", "¥N²z¤H", "¥Ó½Ð¤H", _
                                        "½Ð¨D¶µ", "Email¥D¦®", "Email¤º¤å")
    'arrGridHeadWidth = Array(800, 1200, 1200, 1000, 1000, 800, 1000, 1000)
    'end 2021/11/01
    arrGridHeadText = Array("¬y¤ô¸¹", "­^¤å©w½Z¥[µù", "¤é¤å©w½Z¥[µù", "¥»©Ò®×¸¹", "¥N²z¤H", "¥Ó½Ð¤H", _
                                        "½Ð¨D¶µ", "Email¥D¦®", "Email¤º¤å")
    arrGridHeadWidth = Array(800, 1200, 1200, 1200, 1000, 1000, 800, 1000, 1000)
    'end 2022/10/05
    
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
       GRD1.CellAlignment = flexAlignCenterCenter
       GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
    Next

   For intI = 1 To GRD1.Rows - 1
        GRD1.row = intI
        For iRow = 0 To GRD1.Cols - 1
           GRD1.col = iRow
           GRD1.CellAlignment = flexAlignLeftCenter '¤º®e¾a¥ª
        Next iRow
   Next intI
   GRD1.Visible = True
   
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
   Case 0 '¥»©Ò®×¸¹
      txtFM2(1).Tag = "": txtFM2(2).Tag = ""   '²MªÅ¥»©Ò®×¸¹ªº¬ÛÃöY½s¸¹¡BX½s¸¹±ø¥ó
      If txtFM2(Index) <> "" Then
         strExc(0) = "select PA01||PA02||PA03||PA04,PA75, PA26||','||PA27||','||PA28||','||PA29||','||PA30 AS appno from patent where " & ChgPatent(txtFM2(Index))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            MsgBox "¥»©Ò®×¸¹¿é¤J¿ù»~!", vbExclamation
         Else
            txtFM2(Index) = RsTemp(0)
            txtFM2(1).Tag = "" & RsTemp.Fields("pa75")
            txtFM2(2).Tag = "" & RsTemp.Fields("appno")
         End If
      End If
   Case 1 '¥N²z¤H
      lblFM2(Index).Caption = ""
      If txtFM2(Index) <> "" Then
         If Len(txtFM2(Index)) = 6 Or Len(txtFM2(Index)) = 8 Then
            stCon = Left(txtFM2(Index) & "000", 9)
            If ClsPDGetAgent(stCon, strTemp) Then
               lblFM2(1).Caption = strTemp
            Else
               '¼Ò²Õ¤w¼u°T®§
            End If
         End If
      End If
   Case 2 '¥Ó½Ð¤H
      lblFM2(Index).Caption = ""
      If txtFM2(Index) <> "" Then
         If Len(txtFM2(Index)) = 6 Or Len(txtFM2(Index)) = 8 Then
            stCon = Left(txtFM2(Index) & "000", 9)
            If ClsPDGetCustomer(stCon, strTemp) Then
               lblFM2(2).Caption = strTemp
            Else
               '¼Ò²Õ¤w¼u°T®§
            End If
         End If
      End If
   End Select
End Sub
