VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170014 
   BorderStyle     =   1  '單線固定
   Caption         =   "公司基本資料"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8160
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
            Picture         =   "frm170014.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170014.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170014.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170014.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170014.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170014.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170014.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170014.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170014.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170014.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170014.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4290
      Left            =   30
      TabIndex        =   15
      Top             =   630
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7567
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm170014.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(9)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(10)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblDsp(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(11)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(12)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtA08(9)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtA08(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtA08(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtA08(2)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtA08(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtA08(7)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtA08(14)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtA08(8)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtA08(15)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtA08(16)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtA08(17)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtA08(18)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtA08(19)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCUID"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtA08(20)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtA08(21)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm170014.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdok"
      Tab(1).Control(1)=   "txt1(1)"
      Tab(1).Control(2)=   "txt1(0)"
      Tab(1).Control(3)=   "GRD1"
      Tab(1).Control(4)=   "Label12"
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   400
         Left            =   -68670
         TabIndex        =   25
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -73380
         MaxLength       =   1
         TabIndex        =   24
         Top             =   405
         Width           =   400
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -74070
         MaxLength       =   1
         TabIndex        =   23
         Top             =   405
         Width           =   400
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm170014.frx":212C
         Height          =   3495
         Left            =   -74990
         TabIndex        =   26
         Top             =   840
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6165
         _Version        =   393216
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
         _Band(0).Cols   =   2
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   21
         Left            =   5880
         TabIndex        =   4
         Top             =   1020
         Width           =   1485
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "2619;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   20
         Left            =   5880
         TabIndex        =   1
         Top             =   390
         Width           =   1485
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "2619;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   4980
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3840
         Width           =   3045
         VariousPropertyBits=   671105055
         Size            =   "5371;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   19
         Left            =   1410
         TabIndex        =   14
         Top             =   3840
         Width           =   1995
         VariousPropertyBits=   671105051
         MaxLength       =   25
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   18
         Left            =   1410
         TabIndex        =   13
         Top             =   3510
         Width           =   1485
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "2619;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   17
         Left            =   1410
         TabIndex        =   12
         Top             =   3180
         Width           =   6645
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "11721;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   16
         Left            =   1410
         TabIndex        =   11
         Top             =   2880
         Width           =   6645
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "11721;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   15
         Left            =   1410
         TabIndex        =   10
         Top             =   2580
         Width           =   1485
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2619;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   8
         Left            =   5880
         TabIndex        =   9
         Top             =   2010
         Width           =   1485
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "2619;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   14
         Left            =   5880
         TabIndex        =   7
         Top             =   1650
         Width           =   1485
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "2619;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   7
         Left            =   1410
         TabIndex        =   3
         Top             =   1020
         Width           =   1305
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "2302;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   6
         Left            =   1410
         TabIndex        =   6
         Top             =   1650
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   2
         Left            =   1410
         TabIndex        =   2
         Top             =   690
         Width           =   6645
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "11721;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   1
         Left            =   1410
         TabIndex        =   0
         Top             =   390
         Width           =   315
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "556;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   4
         Left            =   1410
         TabIndex        =   5
         Top             =   1320
         Width           =   6615
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "11668;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtA08 
         Height          =   300
         Index           =   9
         Left            =   1410
         TabIndex        =   8
         Top             =   1980
         Width           =   1455
         VariousPropertyBits=   671105051
         MaxLength       =   15
         Size            =   "2566;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "健保局投保單位代號："
         Height          =   180
         Index           =   12
         Left            =   3960
         TabIndex        =   39
         Top             =   1050
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公司名稱簡稱："
         Height          =   180
         Index           =   11
         Left            =   4500
         TabIndex        =   38
         Top             =   435
         Width           =   1260
      End
      Begin VB.Label lblDsp 
         AutoSize        =   -1  'True
         Caption         =   "員工姓名"
         Height          =   180
         Index           =   1
         Left            =   2520
         TabIndex        =   37
         Top             =   1710
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "公司別：          －"
         Height          =   180
         Left            =   -74775
         TabIndex        =   35
         Top             =   450
         Width           =   1350
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1620
         X2              =   7920
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "會計師事務所"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   34
         Top             =   2340
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "電　　話："
         Height          =   180
         Index           =   10
         Left            =   450
         TabIndex        =   33
         Top             =   3900
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯  絡  人："
         Height          =   180
         Index           =   9
         Left            =   450
         TabIndex        =   32
         Top             =   3570
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "地　　址："
         Height          =   180
         Index           =   8
         Left            =   450
         TabIndex        =   31
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名　　稱："
         Height          =   180
         Index           =   7
         Left            =   450
         TabIndex        =   30
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "統一編號："
         Height          =   180
         Index           =   6
         Left            =   450
         TabIndex        =   29
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "營利事業稅籍編號："
         Height          =   180
         Index           =   5
         Left            =   4140
         TabIndex        =   28
         Top             =   2070
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "扣繳憑單媒體代號："
         Height          =   180
         Index           =   1
         Left            =   4140
         TabIndex        =   27
         Top             =   1710
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "統一編號："
         Height          =   180
         Index           =   0
         Left            =   450
         TabIndex        =   21
         Top             =   1065
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "負  責  人："
         Height          =   180
         Index           =   2
         Left            =   480
         TabIndex        =   20
         Top             =   1710
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名　　稱："
         Height          =   180
         Index           =   3
         Left            =   450
         TabIndex        =   19
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別："
         Height          =   180
         Index           =   4
         Left            =   450
         TabIndex        =   18
         Top             =   435
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "地　　址："
         Height          =   180
         Left            =   480
         TabIndex        =   17
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "房屋稅籍編號："
         Height          =   180
         Left            =   90
         TabIndex        =   16
         Top             =   2040
         Width           =   1260
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
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
End
Attribute VB_Name = "frm170014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/22 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/1/2 add by sonia
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_A08 As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim m_bConfirmCheck As Boolean
Dim m_bActived As Boolean


Private Sub cmdok_Click()
   If txt1(0) & txt1(1) <> "" Then
      If RunNick(txt1(0), txt1(1)) Then
         txt1(0).SetFocus
         Exit Sub
      End If
      GetData
   Else
      MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
      txt1(0).SetFocus
   End If
End Sub

Sub GetData()
Dim stCon As String
   
   stCon = ""
   If txt1(0) <> "" Then
      stCon = stCon & " and a0801>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
      stCon = stCon & " and a0801<='" & txt1(1) & "' "
   End If
   'Modified by Morgan 2015/1/15 +a0821
   strExc(0) = "SELECT a0801,a0802,a0807,st02,a0808,a0809,a0821,a0816,a0815 FROM ACC080,staff " & _
               " where a0806=st01(+) " & stCon & " order by a0801"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set GRD1.Recordset = RsTemp.Clone
      GRD1.FormatString = GRD1.FormatString
      SetGrd
   End If
End Sub

Private Sub Form_Activate()
   If m_bActived = False Then
      SetInputEntry
      m_bActived = True
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   
   InitialField
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      Form_KeyDown vbKeyF2, 0
   End If
   UpdateToolbarState
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170014 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from ACC080 where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_A08 = .Fields.Count
      ReDim m_FieldList(TF_A08) As FIELDITEM
      For Each oText In txtA08
         idx = oText.Index
         m_FieldList(idx).fiName = "A08" & Format(idx, "00")
         'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
         'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
            m_FieldList(idx).fiType = 0
         'Else
         '   m_FieldList(idx).fiType = 1
         'End If
         'end 2017/06/29
      Next
      End With
   End If
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
Dim stKey01 As String
Dim adoRst As New ADODB.Recordset
   
   stKey01 = txtA08(1)
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM ACC080" & _
            " WHERE a0801 = '" & stKey01 & "'"
      Case -2
         strExc(0) = "SELECT * FROM ACC080 order by 1 ASC"
      Case -1
         strExc(0) = "SELECT * FROM ACC080" & _
            " WHERE a0801 <'" & stKey01 & "' order by 1 DESC"
      Case 1
         strExc(0) = "SELECT * FROM ACC080" & _
            " WHERE a0801 >'" & stKey01 & "' order by 1 ASC"
      Case 2
         strExc(0) = "SELECT * FROM ACC080 order by 1 DESC"
   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      UpdateCtrlData adoRst
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
      txtA08(1).SetFocus
      txtA08_GotFocus 1
   End If
End Function

Private Sub GRD1_Click()
   Dim lCurRow As Long, i As Integer, j As Integer
   lCurRow = GRD1.row
   If lCurRow > 0 Then
      If GRD1.TextMatrix(lCurRow, 0) <> "" Then
         If GRD1.CellBackColor <> &HFFC0C0 Then
            GRD1.Visible = False
            For j = 1 To GRD1.Rows - 1
               GRD1.row = j
               If GRD1.CellBackColor <> QBColor(15) Then
                  For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = QBColor(15)
                  Next i
               End If
            Next j
            GRD1.row = lCurRow
            For i = 0 To GRD1.Cols - 1
                GRD1.col = i
                GRD1.CellBackColor = &HFFC0C0
            Next i
            GRD1.Visible = True
         End If
      End If
   End If
End Sub

Private Sub GRD1_DblClick()
Dim lCurRow As Long
   
   lCurRow = GRD1.row
   '呼叫查詢
   If lCurRow > 0 Then
      If GRD1.TextMatrix(lCurRow, 0) <> "" Then
         If TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
            If txtA08(1).Locked = False Then
               txtA08(1).Text = GRD1.TextMatrix(lCurRow, 0)
               If TBar1.Buttons(11).Enabled = True Then
                  Call Tbar1_ButtonClick(TBar1.Buttons(11))
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 2 Then
      txt1(0).SetFocus
      TextInverse txt1(0)
   ElseIf SSTab1.Tab = 0 And PreviousTab = 2 Then
      GRD1_DblClick
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   CloseIme
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA08_GotFocus(Index As Integer)
   TextInverse txtA08(Index)
   Select Case Index
      Case 2, 4, 6, 16, 17, 18
         OpenIme
      Case Else
         CloseIme
   End Select
End Sub

Private Sub ClearField()
   For Each oText In txtA08
      oText.Text = Empty
   Next
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_A08
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   m_bConfirmCheck = False
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 3) As String
   With p_Rst
   If .RecordCount > 0 Then
      For Each oText In txtA08
         idx = oText.Index
         ''日期轉民國
         'If idx = 4 Or idx = 5 Then
         '   m_FieldList(idx).fiOldData = TransDate("" & .Fields(m_FieldList(idx).fiName), 1)
         'Else
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         'End If
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         oText.Text = m_FieldList(idx).fiOldData
      Next
      
      If txtA08(6) <> "" Then
         If ClsPDGetStaffN(txtA08(6), strExc(1), , True) Then
            lblDsp(1) = strExc(1)
         End If
      End If
      
      CUID(1) = "" & .Fields("a0810")
      CUID(2) = "" & .Fields("a0811")
      CUID(3) = "" & .Fields("a0812")
   End If
   End With
   UpdateCUID CUID, textCUID
   txtA08(1).Tag = txtA08(1)
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtA08
      oText.Locked = bLocked
   Next
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
   Dim strTemp As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strUName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strUDate = ChangeTStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strUTime = Format(p_CUID(3), "##:##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "UPDATE : " & strUName & " " & _
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
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If

   End Select
End Sub

' 執行指令
Public Sub OnAction(ByVal KeyCode As Integer)
   Dim bCancel As Boolean
   
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         SSTab1.Tab = 0
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF3 ' 修改
         SSTab1.Tab = 0
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF5 ' 刪除
         SSTab1.Tab = 0
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         SSTab1.Tab = 0
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
         bCancel = False
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  bCancel = True
               End If
            Case Else
               bCancel = True
         End Select
         If bCancel = True Then
            txtA08(1) = txtA08(1).Tag
            m_EditMode = 0
            SetInputEntry
            ShowRecord
            UpdateToolbarState
         End If
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
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
         If m_bUpdate And txtA08(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtA08(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtA08(1) <> "" Then
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

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1
         txtA08(1).Locked = False
         If Me.Visible = True Then
            txtA08(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 2
         txtA08(1).Locked = True
         If Me.Visible = True Then
            txtA08(2).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 4
         txtA08(1).Locked = False
         If Me.Visible = True Then
            txtA08(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case Else
         txtA08(1).Locked = True
         If Me.Visible = True Then
            txtA08(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = True
   End Select
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
               txtA08(1).SetFocus
               txtA08_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   For Each oText In txtA08
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtA08_Validate idx, bCancel
         If bCancel = True Then
            txtA08(idx).SetFocus
            txtA08_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   '查詢
   If m_EditMode = 4 Then
      If txtA08(1) = "" Then
         ShowMsg "請輸入公司別 !"
         txtA08(1).SetFocus
         txtA08_GotFocus 1
         GoTo EscPoint
      End If
      
   '維護
   Else
      If txtA08(1) = "" And txtA08(1).Locked = False Then
         ShowMsg "請輸入公司別 !"
         txtA08(1).SetFocus
         txtA08_GotFocus 1
         GoTo EscPoint
      End If
      If txtA08(2) = "" And txtA08(2).Locked = False Then
         ShowMsg "請輸入名稱 !"
         txtA08(2).SetFocus
         txtA08_GotFocus 2
         GoTo EscPoint
      End If
   End If
   
   'Add by Sindy 2021/12/22 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Function
   End If

   TxtValidate = True
   
EscPoint:
   m_bConfirmCheck = False
    
End Function

Private Function AddRecord() As Boolean
Dim stCols As String, stValues As String, stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtA08
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            ''日期轉西元
            'If idx = 4 Or idx = 5 Then
            '   stValues = stValues & "," & CNULL(DBDATE(m_FieldList(idx).fiNewData), True)
            'Else
               stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
            'End If
         End If
      End If
   Next
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO ACC080 (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   
   AddRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
    
End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE ACC080 SET "
   stSet = ""
   For Each oText In txtA08
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where a0801='" & txtA08(1) & "'; end; "
      
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

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

Private Sub UpdateFieldNewData()
   For Each oText In txtA08
      idx = oText.Index
      Select Case idx
         'Case 4, 5
         '   m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub

Private Sub txtA08_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 7, 15
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtA08_Validate(Index As Integer, Cancel As Boolean)
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Select Case Index
         Case 6
            If txtA08(Index) <> "" Then
               If ChkStaffID(txtA08(Index)) = True Then
                  Cancel = True
               End If
               If Cancel = False And ClsPDGetStaffN(txtA08(Index), strExc(1), , True) = False Then
                  Cancel = True
               Else
                  lblDsp(1) = strExc(1)
               End If
            End If
         Case 7, 15
            If txtA08(Index) <> "" Then
               If GetTextLength(txtA08(Index).Text) <> 8 Then
                  MsgBox "統一編號必須是 8碼 !", vbCritical
                  Cancel = True
               End If
               If Cancel = False Then
                  If CheckID(1, txtA08(Index).Text) = False Then
                     If MsgBox("統一編號錯誤，是否確定 ?", vbYesNo + vbCritical) = vbNo Then
                        Cancel = True
                     End If
                  End If
               End If
            End If
      End Select
      
      If Cancel = True Then TextInverse txtA08(Index)
      
      '若是按確定的檢查時略過, 檢查代號檔
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
         End Select
      End If
   End If
End Sub

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '刪除
   stSQL = "delete from ACC080 where a0801='" & txtA08(1) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtA08(1).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   'Modified by Morgan 2015/1/15 +a0821
   arrGridHeadText = Array("代號", "公司名稱", "統一編號", "負責人", "營利事業稅籍", "房屋稅籍", "健保局投保單位代號", "會計師", "會計師統編")
   arrGridHeadWidth = Array(600, 1200, 800, 800, 1200, 1000, 1700, 1200, 1000)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub
