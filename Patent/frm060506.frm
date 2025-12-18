VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060506 
   BorderStyle     =   1  '單線固定
   Caption         =   "核准函輸入備註維護"
   ClientHeight    =   5700
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8232
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8232
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
            Picture         =   "frm060506.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060506.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060506.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060506.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060506.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060506.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060506.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060506.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060506.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060506.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060506.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8232
      _ExtentX        =   14520
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
      Height          =   4920
      Left            =   60
      TabIndex        =   9
      Top             =   690
      Width           =   8115
      _ExtentX        =   14309
      _ExtentY        =   8678
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm060506.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(7)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label3(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textCUID"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label2(4)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label2(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtAM(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtAM(5)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtAM(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtAM(2)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtAM(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtAM(6)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtAM(7)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Frame1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdMsg"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm060506.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdQuery"
      Tab(1).Control(1)=   "GRD1"
      Tab(1).Control(2)=   "txtFM2(3)"
      Tab(1).Control(3)=   "Label1(13)"
      Tab(1).Control(4)=   "Label1(9)"
      Tab(1).Control(5)=   "lblPS"
      Tab(1).Control(6)=   "Label1(12)"
      Tab(1).Control(7)=   "Label1(11)"
      Tab(1).Control(8)=   "Label1(10)"
      Tab(1).Control(9)=   "txtFM2(0)"
      Tab(1).Control(10)=   "txtFM2(1)"
      Tab(1).Control(11)=   "txtFM2(2)"
      Tab(1).Control(12)=   "lblFM2(1)"
      Tab(1).Control(13)=   "lblFM2(2)"
      Tab(1).ControlCount=   14
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   300
         Left            =   -72300
         TabIndex        =   22
         Top             =   390
         Width           =   885
      End
      Begin VB.CommandButton cmdMsg 
         BackColor       =   &H00C0FFC0&
         Caption         =   "備註內容"
         Height          =   280
         Left            =   3990
         Style           =   1  '圖片外觀
         TabIndex        =   32
         Top             =   3930
         Width           =   975
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame1"
         Height          =   555
         Left            =   1080
         TabIndex        =   30
         Top             =   2850
         Width           =   6855
         Begin VB.CheckBox Check1 
            Caption         =   "核准管制分割期限"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   465
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   -60
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Caption         =   "一般核准"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   1
            Left            =   2550
            TabIndex        =   6
            Top             =   -15
            Width           =   1905
         End
         Begin VB.CheckBox Check1 
            Caption         =   "核對已准專利"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   2
            Left            =   4560
            TabIndex        =   7
            Top             =   -15
            Width           =   1875
         End
         Begin VB.Label Label6 
            Caption         =   "(只列印在核對接洽單)"
            Height          =   225
            Left            =   4710
            TabIndex        =   35
            Top             =   360
            Width           =   1995
         End
         Begin VB.Label Label5 
            Caption         =   "(彈訊息及列印在接洽單)"
            Height          =   225
            Left            =   2520
            TabIndex        =   34
            Top             =   360
            Width           =   2025
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm060506.frx":212C
         Height          =   2595
         Left            =   -74940
         TabIndex        =   10
         Top             =   2190
         Width           =   7905
         _ExtentX        =   13949
         _ExtentY        =   4572
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "流水號|備註內容|本所案號|代理人|申請人|案件性質|訊息種類"
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
         _Band(0).Cols   =   7
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   520
         Index           =   3
         Left            =   -73890
         TabIndex        =   25
         Top             =   1320
         Width           =   6255
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "11033;917"
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
         Left            =   -74910
         TabIndex        =   47
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "模糊比對"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   9
         Left            =   -74910
         TabIndex        =   46
         Top             =   1560
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblPS 
         Caption         =   "P.S. 輸入本所案號會另外帶該案代理人和申請人的其他設定"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74940
         TabIndex        =   45
         Top             =   1920
         Width           =   4845
      End
      Begin MSForms.TextBox txtAM 
         Height          =   300
         Index           =   7
         Left            =   7440
         TabIndex        =   28
         Top             =   405
         Visible         =   0   'False
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAM 
         Height          =   300
         Index           =   6
         Left            =   3600
         TabIndex        =   20
         Top             =   450
         Visible         =   0   'False
         Width           =   855
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAM 
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
      Begin MSForms.TextBox txtAM 
         Height          =   915
         Index           =   2
         Left            =   1080
         TabIndex        =   1
         Top             =   870
         Width           =   5940
         VariousPropertyBits=   -1466941413
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "10477;1614"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAM 
         Height          =   300
         Index           =   4
         Left            =   1080
         TabIndex        =   3
         Top             =   2220
         Width           =   1170
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2064;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAM 
         Height          =   300
         Index           =   5
         Left            =   1080
         TabIndex        =   4
         Top             =   2550
         Width           =   1170
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2064;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAM 
         Height          =   300
         Index           =   3
         Left            =   1080
         TabIndex        =   2
         Top             =   1875
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
         Height          =   255
         Index           =   3
         Left            =   4500
         TabIndex        =   42
         Top             =   473
         Visible         =   0   'False
         Width           =   1875
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "3307;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   2310
         TabIndex        =   44
         Top             =   2580
         Width           =   5505
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "9710;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   2310
         TabIndex        =   43
         Top             =   2235
         Width           =   5505
         VariousPropertyBits=   27
         Caption         =   "1111"
         Size            =   "9710;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID 
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Top             =   4530
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
         Index           =   12
         Left            =   -74910
         TabIndex        =   40
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   11
         Left            =   -74910
         TabIndex        =   39
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   10
         Left            =   -74910
         TabIndex        =   38
         Top             =   435
         Width           =   900
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   285
         Index           =   0
         Left            =   -73890
         TabIndex        =   21
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
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   37
         Top             =   720
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
         Index           =   2
         Left            =   -72750
         TabIndex        =   36
         Top             =   1035
         Width           =   5595
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "9869;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "　　　2.備註內容皆會帶入進度備註及承辦單。"
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
         Index           =   3
         Left            =   1080
         TabIndex        =   33
         Top             =   4230
         Width           =   4035
      End
      Begin VB.Label Label4 
         Caption         =   "若備註長度過長，請按Enter鍵折行。"
         Height          =   255
         Left            =   3960
         TabIndex        =   31
         Top             =   1875
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "初審核准："
         Height          =   180
         Index           =   7
         Left            =   6480
         TabIndex        =   29
         Top             =   450
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "訊息種類："
         Height          =   180
         Index           =   8
         Left            =   135
         TabIndex        =   27
         Top             =   2910
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "模糊比對"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   6
         Left            =   7080
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "注意：1.核准管制分割期限-(預設)                        會在核准函輸入時自動加上日期。"
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
         Index           =   2
         Left            =   1080
         TabIndex        =   19
         Top             =   3975
         Width           =   6855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "流水號："
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註內容："
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   17
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         Height          =   180
         Index           =   2
         Left            =   315
         TabIndex        =   16
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   3
         Left            =   315
         TabIndex        =   15
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   180
         Index           =   4
         Left            =   2655
         TabIndex        =   14
         Top             =   510
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   5
         Left            =   135
         TabIndex        =   13
         Top             =   1935
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "注意：1.代理人/申請人可輸入6碼或8碼，6碼代表含關係企業。112/1/9 取消"
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
         Left            =   1080
         TabIndex        =   12
         Top             =   3480
         Visible         =   0   'False
         Width           =   6300
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
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   11
         Top             =   3720
         Visible         =   0   'False
         Width           =   5820
      End
   End
End
Attribute VB_Name = "frm060506"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/02/18 整合特殊備註維護：1.將現有資料6碼Y/X編號補足為8碼；2.在輸入Y/X編號若為6碼，統一補足為8碼。
'Memo by Lydia 2021/11/01 改成Form2.0 ; GRD1改字型=新細明體-ExtB、txtAM(index)、Label2(index)、textCUID、txtFM2(index)、lblFM2(index)
'Memo by Lydia 2021/11/01 畫面頁籤改成「單筆資料」和「多筆查詢」：上方工具列的「查詢」帶出第一筆符合的資料，在多筆查詢的頁籤可以輸入條件進行查詢，並且在下方的Grid呈現多筆資料。
'Created by Lydia 2014/11/05 核准函輸入備註維護
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

'Added by Lydia 2017/08/21 初審核准-備註內容
'Modified by Lydia 2017/10/12 一併產生行事曆
'Private Const mChk1Item01 = "請管制催分割期限"
Private Const mChk1Item01 = "行事曆已管制催分割期限"
Private Const mChk1Item02 = "行事曆已管制2次催分割期限"
'end 2017/08/21

Private Const mInPty = "1001,1503,926," '案件性質
Private Const mInPtyStr = "核准函(1001)、改變原處分(1503)及核對已准專利(926)"    '案件性質message

Private Sub Check1_Click(Index As Integer)

'Modified by Lydia 2019/07/09  排除查詢
'If m_EditMode <> 0 Then '編輯
If m_EditMode = 1 Or m_EditMode = 2 Then '新增1,修改2
   If Index = 0 And Check1(0).Value = 1 Then
        'Modified by Lydia 2017/08/21
        'txtAM(2) = "請管制催分割期限"
        txtAM(2) = mChk1Item01 & " 或 " & mChk1Item02 & " (二選一)"
        
        Check1(1).Value = 0: Check1(2).Value = 0
   End If
Else  '瀏覽
    If Check1(0).Value = 1 Then
       Check1(1).Value = 0: Check1(2).Value = 0
       Check1(1).Enabled = False: Check1(2).Enabled = False
    ElseIf Check1(1).Value = 1 Or Check1(2).Value = 1 Then
       Check1(0).Value = 0: Check1(0).Enabled = False
    End If

End If

End Sub

'Added by Lydia 2021/11/01
Private Sub cmdQuery_Click()
   
   stCon = ""
   If txtFM2(0) <> "" Then
      If Trim(txtFM2(1).Tag & txtFM2(2).Tag) = "" Then
          stCon = stCon & " and am03='" & txtFM2(0) & "'"
      Else
          '另外抓本所案號的相關Y編號、X編號條件
          stCon = stCon & " and (am03='" & txtFM2(0) & "'"
          If txtFM2(1).Tag <> "" Then stCon = stCon & " or instr(" & CNULL(txtFM2(1).Tag) & ", am04) > 0 "
          If txtFM2(2).Tag <> "" Then stCon = stCon & " or instr(" & CNULL(txtFM2(2).Tag) & ", am05) > 0 "
          stCon = stCon & ") "
      End If
   Else
      txtFM2(1).Tag = "": txtFM2(2).Tag = ""   '清空本所案號的相關Y編號、X編號條件
   End If
   If txtFM2(1) <> "" Then
      stCon = stCon & " and am04 like '" & txtFM2(1) & "%'"
   End If
   If txtFM2(2) <> "" Then
      stCon = stCon & " and am05 like '" & txtFM2(2) & "%'"
   End If
   'Added by Lydia 2022/10/03 增加"備註"查詢
   If txtFM2(3) <> "" Then
       stCon = stCon & " and upper(am02) like '%" & ChgSQL(UCase(txtFM2(3))) & "%' "
   End If
   'end 2022/10/03
   
   stSQL = "SELECT AM01,AM02,AM03,AM04,AM05,AM06,AM07,DECODE(AM07,'1','一般','2','核對','3','一般＆核對','4','初審',AM07) AM07T " & _
                "FROM APPROVALMEMO2 WHERE 1=1  " & stCon & " ORDER BY AM01"
   intR = 0
   Set rsRead = ClsLawReadRstMsg(intR, stSQL)
   
   Call SetGrd(True)
   If intR = 1 Then
        grd1.FixedCols = 0
        Set grd1.Recordset = rsRead
        Call SetGrd
        grd1.FixedCols = 5
   End If
End Sub

'Added by Lydia 2021/11/01
Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("流水號", "備註內容", "本所案號", "代理人", "申請人", "AM06", "AM07", "訊息種類")
   arrGridHeadWidth = Array(800, 1200, 1200, 1000, 1000, 0, 0, 1000)
          
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
      'Case vbKeyF9, vbKeyReturn '確定
      Case vbKeyF9
         '備註欄的解開enter鍵控制
         'KeyCode = 0: Action 11
          'Modified by Lydia 2021/11/22 取消
          'If Me.ActiveControl <> txtAM(2) Then KeyCode = 0: Action 11
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
   m_bInsert = IsUserHasRightOfFunction("frm060506", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm060506", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm060506", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm060506", strFind, False)
  
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
   Set frm060506 = Nothing
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
         If m_bUpdate And txtAM(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtAM(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtAM(1) <> "" Then
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
      For Each oText In txtAM
         oText.Locked = True
      Next
      SSTab1.TabEnabled(1) = True
      'Modified by Lydia 2015/01/05
      Frame1.Enabled = False
   Case Else
      For Each oText In txtAM
         oText.Locked = False
      Next
      If m_EditMode <> 4 Then
         txtAM(1).Locked = True
         txtAM(2).SetFocus
         txtAM_GotFocus 2 '當form第一次執行,按"新增"回應會變慢的原因在於輸入法開啟
      End If
      SSTab1.TabEnabled(1) = False
      Frame1.Enabled = True
      Check1(0).Enabled = True: Check1(1).Enabled = True: Check1(2).Enabled = True
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
         If txtAM(1).Text = "" Then
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
         txtAM(1).Enabled = True
         txtAM(1).SetFocus
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
         If Val(m_EditMode) > 0 And Trim(txtAM(3)) <> "" And ((Left(Trim(txtAM(3)), 1) = "P" And Len(Trim(txtAM(3))) < 10) Or (Left(Trim(txtAM(3)), 3) = "FCP" And Len(Trim(txtAM(3))) < 12)) Then
             Call txtAM_Validate(3, bCancel)
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
                  If SetKindVaL = True Then
                     If FormSave() = False Then
                        MsgBox "存檔失敗!", vbCritical
                        Exit Sub
                     Else
                        strKind = m_EditMode 'Added by Lydia 2021/11/01 記錄新增模式
                        m_EditMode = 0
                        'Modified by Morgan 2017/9/13
                        'If m_EditMode = 1 Then
                        If txtAM(1) = "" Then
                        'end 2017/9/13
                           ShowRecord 3
                        Else
                           ReadData txtAM(1)
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
                           If txtAM(3) <> "" Then
                               txtFM2(0) = txtAM(3)
                               Call txtFM2_Validate(0, False)
                           Else
                               If txtAM(4) <> "" Then
                                  txtFM2(1) = ChangeCustomerS(txtAM(4))
                                  Call txtFM2_Validate(1, False)
                               End If
                               If txtAM(5) <> "" Then
                                  txtFM2(2) = ChangeCustomerS(txtAM(5))
                                  Call txtFM2_Validate(2, False)
                               End If
                           End If
                           SSTab1.Tab = 1
                           Call cmdQuery_Click
                       End If
                       'end 2021/11/01
                     End If
                  End If
               End If
            '查詢
            Case 4
               If ReadData(txtAM(1)) = False Then
                  MsgBox "無資料!", vbExclamation
                  Exit Sub
               Else
                  m_EditMode = 0
               End If
         End Select
      Case 12 '按下取消
         m_EditMode = 0
         txtAM(1) = txtAM(1).Tag
         If txtAM(1) <> "" Then
            If ReadData(txtAM(1)) = False Then
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
         strExc(0) = "SELECT nvl(min(am01),0) FROM ApprovalMemo2"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 1 '前一筆
         strExc(0) = "SELECT nvl(max(am01),0) FROM ApprovalMemo2 where am01<" & txtAM(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 6
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 2 '後一筆
         strExc(0) = "SELECT nvl(min(am01),0) FROM ApprovalMemo2 where am01>" & txtAM(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 7
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 3 '最後筆
         strExc(0) = "SELECT nvl(max(am01),0) FROM ApprovalMemo2"
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
      stCon = " and am01=" & pKey
   '多筆
   Else
      If txtAM(2) <> "" Then
         'Modified by Morgan 2017/9/13
         'stCon = stCon & " and am02 like '%" & txtAM(2) & "%'"
         stCon = stCon & " and am02 like '%" & ChgSQL(txtAM(2)) & "%'"
      End If
      If txtAM(3) <> "" Then
         stCon = stCon & " and am03='" & txtAM(3) & "'"
      End If
      If txtAM(4) <> "" Then
         stCon = stCon & " and am04 like '" & txtAM(4) & "%'"
      End If
      If txtAM(5) <> "" Then
         stCon = stCon & " and am05 like '" & txtAM(5) & "%'"
      End If
      If txtAM(6) <> "" Then
         stCon = stCon & " and am06='" & txtAM(6) & "'"
      End If
      'Modified by Lydia 2015/01/05 核准改為勾選
      'If txtAM(7) <> "" Then
      '   stCon = stCon & " and am07='" & txtAM(7) & "'"
      'End If
      If Check1(0).Value = 1 And Check1(1).Value = 1 And Check1(2).Value = 1 Then
      
      ElseIf Check1(1).Value = 1 And Check1(2).Value = 1 Then
         stCon = stCon & " and am07='3'"
      ElseIf Check1(0).Value = 1 Then
         stCon = stCon & " and am07='4'"
      ElseIf Check1(1).Value = 1 Then
         stCon = stCon & " and am07='1'"
      ElseIf Check1(2).Value = 1 Then
         stCon = stCon & " and am07='2'"
      End If
      'end 2015/01/05
   End If
   
   FormReset
   
   strExc(0) = "select * from ApprovalMemo2 where 1=1 " & stCon & " order by am01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If m_EditMode = 4 Then
         'Modified by Lydia 2021/11/01 改成單筆查詢
         'Set GRD1.Recordset = RsTemp.Clone
         'GRD1.FormatString = GRD1.FormatString
         'GRD1.ColWidth(1) = 2775
         'GRD1.ColWidth(2) = 1290
         'GRD1.ColWidth(3) = 1020
         'GRD1.ColWidth(4) = 1020
         'GRD1.ColWidth(5) = 810
         'GRD1.ColWidth(6) = 810
         'For intI = 7 To GRD1.Cols - 1
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
   For Each oText In txtAM
      oText = "" & .Fields("am" & Format(oText.Index, "00"))
   Next
   End With

   SetKind

   UpdateCUID rsQuery
   
   txtAM(1).Tag = txtAM(1)
   If txtAM(4) <> "" Then txtAM_Validate 4, False
   If txtAM(5) <> "" Then txtAM_Validate 5, False
   'Remove by Lydia 2015/01/05
'   If txtAM(6) <> "" Then txtAM_Validate 6, False
'   If txtAM(7) <> "" Then txtAM_Validate 7, False
   'end 2015/01/05
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
  
   If IsNull(rsSrcTmp.Fields("am08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("am08")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("am08"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("am09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("am09")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("am09"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("am10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("am10")) = False Then
         strTemp = rsSrcTmp.Fields("am10")
         strCTime = Format(strTemp, "00:00:00")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("am11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("am11")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("am11"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("am12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("am12")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("am12"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("am13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("am13")) = False Then
         strTemp = rsSrcTmp.Fields("am13")
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
   
   For Each oText In txtAM
      oText.Text = ""
   Next
   
   For Each oLabel In Label2
      oLabel.Caption = ""
   Next
   
   SetKind
   textCUID = ""
   Label1(6).Visible = False
End Sub

Private Sub txtAM_GotFocus(Index As Integer)
   TextInverse txtAM(Index)
   If Index = 2 Then
      OpenIme
   Else
      CloseIme
   End If
End Sub

'Modified by Lydia 2021/11/01 改成Form 2.0
'Private Sub txtAM_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtAM_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index <> 2 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtAM_Validate(Index As Integer, Cancel As Boolean)
   Dim strCusTemp As String, strTemp As String
   Select Case Index
   Case 3 '本所案號
      If txtAM(Index) <> "" Then
         strExc(0) = "select PA01||PA02||PA03||PA04 from patent where " & ChgPatent(txtAM(Index))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            If m_EditMode <> 0 Then 'Added by Lydia 2021/11/01 排除非編輯模式: 因為案號有可能刪除
                 MsgBox "本所案號輸入錯誤!", vbExclamation
                 Cancel = True
            End If 'Added by Lydia 2021/11/01
            'If m_EditMode <> 0 Then Cancel = True 'Remove by Lydia 2021/11/01
         Else
            txtAM(Index) = RsTemp(0)
         End If
      End If
   Case 4 '代理人
      Label2(4).Caption = ""
      If txtAM(Index) <> "" Then
         'Modified by Morgan 2019/7/25 加碼數檢查
         If Len(txtAM(Index)) = 6 Or Len(txtAM(Index)) = 8 Then
            strCusTemp = ChangeCustomerL(txtAM(Index))
            If ClsPDGetAgent(strCusTemp, strTemp) Then
               Label2(4).Caption = strTemp
               'Added by Lydia 2023/02/18 整合特殊備註維護：在輸入Y/X編號若為6碼，統一補足為8碼。
               If m_EditMode <> 0 Then
                  txtAM(Index) = Left(ChangeCustomerL(txtAM(Index)), 8)
               End If
               'end 2023/02/18
            Else
               MsgBox "代理人編號輸入錯誤！", vbCritical
               If m_EditMode <> 0 Then Cancel = True
            End If
         Else
            MsgBox "代理人編號只可輸入6碼或8碼！", vbCritical
            If m_EditMode <> 0 Then Cancel = True
         End If
      End If
   Case 5 '申請人
      Label2(5).Caption = ""
      If txtAM(Index) <> "" Then
         'Modified by Morgan 2019/7/25 加碼數檢查
         If Len(txtAM(Index)) = 6 Or Len(txtAM(Index)) = 8 Then
            strCusTemp = ChangeCustomerL(txtAM(Index))
            If ClsPDGetCustomer(strCusTemp, strTemp) Then
               Label2(5).Caption = strTemp
               'Added by Lydia 2023/02/18 整合特殊備註維護：在輸入Y/X編號若為6碼，統一補足為8碼。
               If m_EditMode = 1 <> 0 Then
                  txtAM(Index) = Left(ChangeCustomerL(txtAM(Index)), 8)
               End If
               'end 2023/02/18
            Else
               MsgBox "客戶編號輸入錯誤！", vbCritical
               If m_EditMode <> 0 Then Cancel = True
            End If
         Else
            MsgBox "客戶編號只可輸入6碼或8碼！", vbCritical
            If m_EditMode <> 0 Then Cancel = True
         End If
      End If
      'Modified by Lydia 2015/01/05 改成勾選
'   Case 6 '案件性質
'      Label2(6).Caption = ""
'      If txtAM(Index) <> "" Then
'         If InStr(mInPty, txtAM(Index)) = 0 Then
'            MsgBox "目前僅開放設定" & mInPtyStr & "！", vbExclamation
'            Cancel = True
'         Else
'            If ClsPDGetCaseProperty("FCP", txtAM(Index), strTemp) Then
'               Label2(6).Caption = strTemp
'            Else
'               If m_EditMode <> 0 Then Cancel = True
'            End If
'         End If
'      End If
'Modified by Lydia 2015/01/05 核准改為勾選
'   Case 7 '初審核准
'      If txtAM(Index) <> "" Then
'         If Not (txtAM(Index) = "Y" Or txtAM(Index) = "y") Then
'            MsgBox "只能輸入Y ！", vbExclamation
'            Cancel = True
'         End If
'      End If
   End Select
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean, idx As Integer
   
   If txtAM(2) = "" Then
      MsgBox "備註內容不可空白！", vbExclamation
      txtAM(2).SetFocus
      Exit Function
   'Added by Lydia 2017/08/21 檢查非初審核准的備註內容
   ElseIf Check1(0).Value = 0 Then
        If InStr(txtAM(2), mChk1Item01) > 0 Or InStr(txtAM(2), mChk1Item02) > 0 Then
            MsgBox "非初審核准的備註內容不可輸入:" & mChk1Item01 & "、" & mChk1Item02, vbExclamation
            txtAM(2).SetFocus
            Exit Function
        End If
   'end 2017/08/21
   End If
   If txtAM(3) & txtAM(4) & txtAM(5) = "" Then
      MsgBox "請輸入本所案號、代理人或申請人！", vbExclamation
      txtAM(3).SetFocus
      Exit Function
   End If
   
   For idx = 3 To 7
      txtAM_Validate idx, bCancel
      If bCancel = True Then
         txtAM(idx).SetFocus
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
   
   strExc(0) = strUserNum
   strExc(1) = strSrvDate(1)
   strExc(2) = Right("000000" & ServerTime, 6)
   
   cnnConnection.BeginTrans
   If m_EditMode = 1 Then
      'Modified by Lydia 2017/01/11 避免LOG語法分析錯誤
      'strSql = "insert into ApprovalMemo2(am01,am02,am03,am04,am05,am06,am07,am08,am09,am10)" & _
         " select nvl(max(am01),0)+1 am01,'" & ChgSQL(txtAM(2)) & "' am02" & _
         ",'" & txtAM(3) & "' am03,'" & txtAM(4) & "' am04,'" & txtAM(5) & "' am05,'" & txtAM(6) & "' am06,'" & txtAM(7) & "' am07 " & _
         ", '" & strExc(0) & "' am08, '" & strExc(1) & "' am09, '" & strExc(2) & "' am10 from ApprovalMemo2 "
      strSql = "insert into ApprovalMemo2(am01,am02,am03,am04,am05,am06,am07,am08,am09,am10) VALUES ('" & Pub_GetDefColMaxNo("ApprovalMemo2", "AM01") & "','" & ChgSQL(txtAM(2)) & "' " & _
         ",'" & txtAM(3) & "' ,'" & txtAM(4) & "' ,'" & txtAM(5) & "','" & txtAM(6) & "','" & txtAM(7) & "'" & _
         ", '" & strExc(0) & "', '" & strExc(1) & "', '" & strExc(2) & "') "
         
   Else
      strSql = "update ApprovalMemo2 set am02='" & ChgSQL(txtAM(2)) & "',am03='" & txtAM(3) & "'" & _
         ",am04='" & txtAM(4) & "',am05='" & txtAM(5) & "',am06='" & txtAM(6) & "',am07='" & txtAM(7) & "'" & _
         ",am11='" & strExc(0) & "',am12='" & strExc(1) & "',am13='" & strExc(2) & "' where am01=" & txtAM(1) 'NPmemo的更新日期採trigger
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
   strSql = "delete from ApprovalMemo2 where am01=" & txtAM(1)
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   FormDelete = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function
'Modified by Lydia 2015/01/05 核准改為勾選
Private Sub SetKind()
Check1(0).Enabled = True: Check1(1).Enabled = True: Check1(2).Enabled = True
Check1(0).Value = 0: Check1(1).Value = 0: Check1(2).Value = 0
If txtAM(7) = "4" Then
    Check1(0).Value = 1
ElseIf txtAM(7) = "1" Then
    Check1(1).Value = 1
ElseIf txtAM(7) = "2" Then
    Check1(2).Value = 1
ElseIf txtAM(7) = "3" Then
    Check1(1).Value = 1: Check1(2).Value = 1
End If
End Sub

Private Function SetKindVaL() As Boolean
SetKindVaL = False
If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 Then
   MsgBox "請勾選訊息種類!!", vbCritical
   Exit Function
Else
   If Check1(0).Value = 1 And (Check1(1).Value = 1 Or Check1(2).Value = 1) Then
      MsgBox "初審核准不准合併!!", vbCritical
      Check1(1).Value = 0: Check1(2).Value = 0
      Exit Function
   End If
   If Check1(0).Value = 1 Then
      'Modified by Lydia 2017/08/21 增加2次催分割
      'If txtAM(2) <> "請管制催分割期限" Then
      '   MsgBox "初審核准的訊息為-請管制催分割期限!!", vbCritical
      If txtAM(2) <> mChk1Item01 And txtAM(2) <> mChk1Item02 Then
         MsgBox "初審核准的訊息為-" & mChk1Item01 & "、" & mChk1Item02 & " !!", vbCritical
      'end 2017/08/21
         Exit Function
      Else
         txtAM(7) = "4"
      End If
   ElseIf Check1(1).Value = 1 And Check1(2).Value = 1 Then
            txtAM(7) = "3"
       ElseIf Check1(1).Value = 1 Then
                txtAM(7) = "1"
       Else
                txtAM(7) = "2"
   End If
End If
'預設案件性質
If txtAM(7) = "2" Then
   txtAM(6) = "926"
Else
   txtAM(6) = "1001"
End If

If RecIsExist = True Then
   SetKindVaL = False
Else
   SetKindVaL = True
End If
End Function

Private Function RecIsExist() As Boolean
   
strExc(0) = "am07 = '" & txtAM(7) & "' "
'本所案號
If Trim(txtAM(3)) <> "" Then
   strExc(0) = strExc(0) & "and am03='" & Trim(txtAM(3)) & "' "
End If
'代理人
If Trim(txtAM(4)) <> "" Then
   'Modified by Lydia 2019/07/31 改成=判斷; 因為無法先輸入8碼後再輸入6碼
   'strExc(0) = strExc(0) & "and instr(am04,'" & Trim(txtAM(4)) & "') > 0 "
   strExc(0) = strExc(0) & "and am04='" & Trim(txtAM(4)) & "'  "
   'Added by Lydia 2019/01/28 區別只有代理人或客戶的條件
   If Trim(txtAM(5)) = "" Then strExc(0) = strExc(0) & "and am05 is null "
End If
'客戶
If Trim(txtAM(5)) <> "" Then
   'Modified by Lydia 2019/07/31 改成=判斷; 因為無法先輸入8碼後再輸入6碼
   'strExc(0) = strExc(0) & "and instr(am05,'" & Trim(txtAM(5)) & "') > 0 "
   strExc(0) = strExc(0) & "and am05='" & Trim(txtAM(5)) & "'  "
'Added by Morgan 2016/7/21 沒有也要判斷,否則 Y+X 和 Y 就不可同時設定
'Modified by Lydia 2019/01/28 區別只有代理人或客戶的條件
'Else
'   strExc(0) = strExc(0) & "and am05 is null "
'end 2016/7/21
   If Trim(txtAM(4)) = "" Then strExc(0) = strExc(0) & "and am04 is null "
'end 2019/01/28
End If

If Left(strExc(0), 3) = "and" Then strExc(0) = Mid(strExc(0), 4, Len(strExc(0)) - 4)

   strExc(1) = " select * from ApprovalMemo2 where " & strExc(0)
   intR = 1
   Set rsRead = ClsLawReadRstMsg(intR, strExc(1))
   If intR = 1 Then
      'Added by Lydia 2016/1/11 排除現在修改的記錄
      If rsRead.RecordCount = 1 And Trim(rsRead.Fields("AM01")) = Trim(txtAM(1)) Then
         RecIsExist = False
      Else
      'end 2016/1/11
         RecIsExist = True
         MsgBox "已存在同樣條件的記錄(流水號 " & rsRead(0) & " )，請先查詢!!", vbCritical
      End If
   Else
      RecIsExist = False
   End If
   Set rsRead = Nothing
   
End Function

'Added by Lydia 2017/08/21
Private Sub CmdMsg_Click()
Dim strMsg As String
    'Modified by Lydia 2017/10/12
    'strMsg = mChk1Item01 & ": 核准函的本所收文日再加23日" & vbCrLf
    strMsg = mChk1Item01 & ": 核准函的本所收文日再加23日，若遇假日則提前至前一工作日；" & vbCrLf
    strMsg = strMsg & String(50, "-") & vbCrLf
    strMsg = strMsg & mChk1Item02 & ": 第2次催分割期限為核准函的本所收文日" & vbCrLf & _
              String(10, "　") & "再加29日，若遇假日則提前至前一工作日。" & vbCrLf
    MsgBox strMsg, vbInformation + vbOKOnly, "初審核准-備註內容"
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
