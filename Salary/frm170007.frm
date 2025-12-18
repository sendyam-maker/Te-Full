VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170007 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工貸款資料"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8175
   Begin TabDlg.SSTab SSTab1 
      Height          =   4380
      Left            =   45
      TabIndex        =   9
      Top             =   630
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm170007.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDsp(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textCUID"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtLE(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtLE(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtLE(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtLE(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtLE(5)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtLE(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtLE(8)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtLE(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtLE(9)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm170007.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtSum(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtSum(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtSum(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtSum(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txt1(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txt1(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt1(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt1(3)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdok"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "GRD1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label11"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label10"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label9"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label8"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label7"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label16"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label12"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   3
         Left            =   -68025
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   4005
         Width           =   1005
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   0
         Left            =   -73335
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   4005
         Width           =   1005
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   1
         Left            =   -71670
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   4005
         Width           =   780
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   2
         Left            =   -69825
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   4005
         Width           =   780
      End
      Begin VB.TextBox txtLE 
         Height          =   285
         Index           =   9
         Left            =   1670
         MaxLength       =   5
         TabIndex        =   8
         Text            =   "9701"
         Top             =   2760
         Width           =   600
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72715
         MaxLength       =   6
         TabIndex        =   16
         Top             =   435
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73825
         MaxLength       =   6
         TabIndex        =   15
         Top             =   435
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -70410
         MaxLength       =   7
         TabIndex        =   17
         Top             =   405
         Width           =   800
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -69205
         MaxLength       =   7
         TabIndex        =   18
         Top             =   405
         Width           =   800
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   400
         Left            =   -68300
         TabIndex        =   19
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txtLE 
         Height          =   285
         Index           =   6
         Left            =   2610
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "9801"
         Top             =   1800
         Width           =   600
      End
      Begin VB.TextBox txtLE 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   8
         Left            =   1670
         MaxLength       =   7
         TabIndex        =   7
         Text            =   "9999999"
         Top             =   2440
         Width           =   915
      End
      Begin VB.TextBox txtLE 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Index           =   7
         Left            =   1670
         MaxLength       =   7
         TabIndex        =   6
         Text            =   "9999999"
         Top             =   2120
         Width           =   915
      End
      Begin VB.TextBox txtLE 
         Height          =   285
         Index           =   5
         Left            =   1670
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "9701"
         Top             =   1800
         Width           =   600
      End
      Begin VB.TextBox txtLE 
         Height          =   285
         Index           =   2
         Left            =   1670
         MaxLength       =   7
         TabIndex        =   1
         Text            =   "960531"
         Top             =   840
         Width           =   800
      End
      Begin VB.TextBox txtLE 
         Height          =   270
         Index           =   1
         Left            =   1670
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "999999"
         Top             =   520
         Width           =   915
      End
      Begin VB.TextBox txtLE 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   4
         Left            =   1670
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "9999999"
         Top             =   1480
         Width           =   915
      End
      Begin VB.TextBox txtLE 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   3
         Left            =   1670
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "9999999"
         Top             =   1160
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm170007.frx":0038
         Height          =   3015
         Left            =   -74985
         TabIndex        =   20
         Top             =   840
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "代號|姓名|貸款日期|貸款本金|利　息|償還起|償還迄|每月金額|目前餘額|最近償還"
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
         _Band(0).Cols   =   10
      End
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3960
         Width           =   5730
         VariousPropertyBits=   671105055
         Size            =   "10107;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "目前餘額："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -68970
         TabIndex        =   37
         Top             =   4005
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "貸款本金："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74280
         TabIndex        =   35
         Top             =   4005
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "利息："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72210
         TabIndex        =   34
         Top             =   4005
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "合計："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74910
         TabIndex        =   33
         Top             =   4005
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "每月金額："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -70770
         TabIndex        =   32
         Top             =   4005
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "最近償還月份："
         Height          =   180
         Left            =   345
         TabIndex        =   28
         Top             =   2805
         Width           =   1260
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "員工代號：                      －"
         Height          =   180
         Left            =   -74790
         TabIndex        =   26
         Top             =   450
         Width           =   2070
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "貸款日期：                   －"
         Height          =   180
         Left            =   -71310
         TabIndex        =   25
         Top             =   450
         Width           =   1935
      End
      Begin MSForms.Label lblDsp 
         Height          =   285
         Index           =   1
         Left            =   2700
         TabIndex        =   24
         Top             =   555
         Width           =   840
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1482;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "目前餘額："
         Height          =   180
         Left            =   710
         TabIndex        =   23
         Top             =   2480
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "每月償還金額："
         Height          =   180
         Left            =   350
         TabIndex        =   22
         Top             =   2160
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "償還期間：                ～"
         Height          =   180
         Left            =   705
         TabIndex        =   21
         Top             =   1845
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "貸款日期："
         Height          =   180
         Left            =   710
         TabIndex        =   13
         Top             =   880
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         Height          =   180
         Index           =   4
         Left            =   710
         TabIndex        =   12
         Top             =   560
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "貸款本金："
         Height          =   180
         Index           =   3
         Left            =   710
         TabIndex        =   11
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "利　　息："
         Height          =   180
         Index           =   2
         Left            =   710
         TabIndex        =   10
         Top             =   1520
         Width           =   900
      End
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
            Picture         =   "frm170007.frx":004D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170007.frx":0369
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170007.frx":0685
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170007.frx":0861
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170007.frx":0B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170007.frx":0E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170007.frx":11B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170007.frx":14D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170007.frx":17ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170007.frx":1B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170007.frx":1E25
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
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
Attribute VB_Name = "frm170007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/20 Form2.0已修改
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
Dim TF_LE As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim m_bConfirmCheck As Boolean
Dim m_bActived As Boolean


Private Sub cmdok_Click()
   If txt1(0) & txt1(1) & txt1(2) & txt1(3) <> "" Then
      If RunNick(txt1(0), txt1(1)) Then
         txt1(0).SetFocus
         Exit Sub
      End If
      If RunNick(txt1(2), txt1(3)) Then
         txt1(2).SetFocus
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
   
   txtSum(0) = ""
   txtSum(1) = ""
   txtSum(2) = ""
   txtSum(3) = ""
   stCon = ""
   If txt1(0) <> "" Then
      stCon = stCon & " and le01>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
      stCon = stCon & " and le01<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
      stCon = stCon & " and le02>=" & DBDATE(txt1(2))
   End If
   If txt1(3) <> "" Then
      stCon = stCon & " and le02<=" & DBDATE(txt1(3))
   End If
   strExc(0) = "SELECT le01,ST02,sqldateT(le02),to_char(le03,'99G999G999') Num1,to_char(le04,'99G999G999') Num2, to_char(substr(le05,1,4))-1911||'/'||substr(le05,5,2),to_char(substr(le06,1,4))-1911||'/'||substr(le06,5,2),to_char(le07,'99G999G999') Num3,to_char(le08,'99G999G999') Num4, decode(nvl(le09,0),0,null,substr(le09,1,4)-1911||'/'||substr(le09,5,2)) FROM Loan_Employee,staff " & _
               " where le01=st01(+) " & stCon & " order by le01,le02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set grd1.Recordset = RsTemp.Clone
      grd1.FormatString = grd1.FormatString
      SetGrd
      'Added by Morgan 2013/10/17
      If RsTemp.RecordCount > 0 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            txtSum(0) = Val(txtSum(0)) + Val(Format("" & RsTemp("Num1")))
            txtSum(1) = Val(txtSum(1)) + Val(Format("" & RsTemp("Num2")))
            txtSum(2) = Val(txtSum(2)) + Val(Format("" & RsTemp("Num3")))
            txtSum(3) = Val(txtSum(3)) + Val(Format("" & RsTemp("Num4")))
            RsTemp.MoveNext
         Loop
         txtSum(0) = Format(txtSum(0), "#,##0")
         txtSum(1) = Format(txtSum(1), "#,##0")
         txtSum(2) = Format(txtSum(2), "#,##0")
         txtSum(3) = Format(txtSum(3), "#,##0")
      End If
      'end 2013/10/17
      
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
   Set frm170007 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from Loan_Employee where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_LE = .Fields.Count
      ReDim m_FieldList(TF_LE) As FIELDITEM
      For Each oText In txtLE
         idx = oText.Index
         m_FieldList(idx).fiName = "LE" & Format(idx, "00")
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
Dim stKey02 As String
Dim adoRst As New ADODB.Recordset
   
   stKey01 = txtLE(1)
   stKey02 = DBDATE(txtLE(2))
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM Loan_Employee" & _
            " WHERE le01 = '" & stKey01 & "' and le02= '" & stKey02 & "'"
      Case -2
         strExc(0) = "SELECT * FROM Loan_Employee order by 2 ASC, 1 ASC"
      Case -1
         strExc(0) = "SELECT * FROM Loan_Employee" & _
            " WHERE le02||le01 <'" & stKey02 & stKey01 & "' order by 2 DESC,1 DESC"
      Case 1
         strExc(0) = "SELECT * FROM Loan_Employee" & _
            " WHERE le02||le01 >'" & stKey02 & stKey01 & "' order by 2 ASC,1 ASC"
      Case 2
         strExc(0) = "SELECT * FROM Loan_Employee order by 2 DESC,1 DESC"
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
      txtLE(1).SetFocus
      txtLE_GotFocus 1
   End If
End Function

Private Sub GRD1_Click()
   Dim lCurRow As Long, i As Integer, j As Integer
   lCurRow = grd1.row
   If lCurRow > 0 Then
      If grd1.TextMatrix(lCurRow, 0) <> "" Then
         If grd1.CellBackColor <> &HFFC0C0 Then
            grd1.Visible = False
            For j = 1 To grd1.Rows - 1
               grd1.row = j
               If grd1.CellBackColor <> QBColor(15) Then
                  For i = 0 To grd1.Cols - 1
                     grd1.col = i
                     grd1.CellBackColor = QBColor(15)
                  Next i
               End If
            Next j
            grd1.row = lCurRow
            For i = 0 To grd1.Cols - 1
                grd1.col = i
                grd1.CellBackColor = &HFFC0C0
            Next i
            grd1.Visible = True
         End If
      End If
   End If
End Sub

Private Sub GRD1_DblClick()
Dim lCurRow As Long
   
   lCurRow = grd1.row
   '呼叫查詢
   If lCurRow > 0 Then
      If grd1.TextMatrix(lCurRow, 0) <> "" Then
         If TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
            If txtLE(1).Locked = False Then
               txtLE(1).Text = grd1.TextMatrix(lCurRow, 0)
               txtLE(2).Text = ChangeTDateStringToTString(grd1.TextMatrix(lCurRow, 2))
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

Private Sub txtLE_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtLE(Index)
End Sub

Private Sub ClearField()
   For Each oText In txtLE
      oText.Text = Empty
   Next
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_LE
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   m_bConfirmCheck = False
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   With p_Rst
   If .RecordCount > 0 Then
      For Each oText In txtLE
         idx = oText.Index
         m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         '日期轉民國
         If idx = 2 Then
            oText.Text = TransDate(m_FieldList(idx).fiOldData, 1)
         '償還年月轉民國年月
         ElseIf idx = 5 Or idx = 6 Or idx = 9 Then
            If m_FieldList(idx).fiOldData <> "" Then
               oText.Text = "" & m_FieldList(idx).fiOldData - 191100
            End If
         Else
            oText.Text = m_FieldList(idx).fiOldData
         End If
      Next
      
      If ClsPDGetStaffN(txtLE(1), strExc(1), , True) Then
         lblDsp(1) = strExc(1)
      End If
      
      CUID(1) = "" & .Fields("le10")
      CUID(2) = "" & .Fields("le11")
      CUID(3) = "" & .Fields("le12")
      CUID(4) = "" & .Fields("le13")
      CUID(5) = "" & .Fields("le14")
      CUID(6) = "" & .Fields("le15")
   End If
   End With
   UpdateCUID CUID, textCUID
   txtLE(1).Tag = txtLE(1)
   txtLE(2).Tag = txtLE(2)
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtLE
      oText.Locked = bLocked
   Next
   txtLE(9).Locked = True         '最近償還月份欄鎖住
   If m_EditMode <> 1 Then
      txtLE(7).Locked = True      '每月償還金額欄鎖住
      txtLE(8).Locked = True      '目前餘額欄鎖住
   Else
      txtLE(7).Locked = False     '每月償還金額欄鎖住
      txtLE(8).Locked = False     '目前餘額欄鎖住
   End If
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
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
            txtLE(1) = txtLE(1).Tag
            txtLE(2) = txtLE(2).Tag
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
         If m_bUpdate And txtLE(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtLE(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtLE(1) <> "" Then
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
         txtLE(1).Locked = False
         If Me.Visible = True Then
            txtLE(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 2
         txtLE(1).Locked = True
         If Me.Visible = True Then
            txtLE(3).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 4
         txtLE(1).Locked = False
         txtLE(2).Locked = False
         If Me.Visible = True Then
            txtLE(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case Else
         txtLE(1).Locked = True
         If Me.Visible = True Then
            txtLE(1).SetFocus
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
               txtLE(1).SetFocus
               txtLE_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   For Each oText In txtLE
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtLE_Validate idx, bCancel
         If bCancel = True Then
            txtLE(idx).SetFocus
            txtLE_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   '查詢
   If m_EditMode = 4 Then
      If txtLE(1) = "" Then
         ShowMsg "請輸入員工代號 !"
         txtLE(1).SetFocus
         txtLE_GotFocus 1
         GoTo EscPoint
      End If
      If txtLE(2) = "" Then
         ShowMsg "請輸入貸款日期 !"
         txtLE(2).SetFocus
         txtLE_GotFocus 2
         GoTo EscPoint
      End If
      
   '維護
   Else
      If txtLE(1) = "" And txtLE(1).Locked = False Then
         ShowMsg "請輸入員工代號 !"
         txtLE(1).SetFocus
         txtLE_GotFocus 1
         GoTo EscPoint
      End If
      If txtLE(2) = "" And txtLE(2).Locked = False Then
         ShowMsg "請輸入貸款日期 !"
         txtLE(2).SetFocus
         txtLE_GotFocus 2
         GoTo EscPoint
      End If
      If Val(txtLE(3)) = 0 And txtLE(3).Locked = False Then
         ShowMsg "請輸入貸款本金 !"
         txtLE(3).SetFocus
         txtLE_GotFocus 3
         GoTo EscPoint
      End If
      If Val(txtLE(5)) = 0 And txtLE(5).Locked = False Then
         ShowMsg "請輸入償還時間－起 !"
         txtLE(5).SetFocus
         txtLE_GotFocus 5
         GoTo EscPoint
      End If
      If Val(txtLE(6)) = 0 And txtLE(6).Locked = False Then
         ShowMsg "請輸入償還時間－迄 !"
         txtLE(6).SetFocus
         txtLE_GotFocus 6
         GoTo EscPoint
      End If
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
   For Each oText In txtLE
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO Loan_Employee (" & stCols & ") Values (" & stValues & ")"
   
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
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE Loan_Employee SET "
   stSet = ""
   For Each oText In txtLE
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
      stSQL = stSQL & stSet & " where le01='" & txtLE(1) & "' and le02=" & DBDATE(txtLE(2)) & "; end; "
      
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
   For Each oText In txtLE
      idx = oText.Index
      Select Case idx
         Case 2
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         Case 5, 6, 9
            If oText.Text <> "" Then
               m_FieldList(idx).fiNewData = Val(oText.Text) + 191100
            End If
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub

Private Sub txtLE_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 1 'Added by Morgan 2013/4/16
      
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtLE_Validate(Index As Integer, Cancel As Boolean)
Dim m_month As Variant
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Select Case Index
         Case 1
            If txtLE(Index) <> "" Then
               If ChkStaffID(txtLE(Index)) = True Then
                  Cancel = True
               End If
               If Cancel = False And ClsPDGetStaff(txtLE(Index), strExc(1)) = False Then
                  Cancel = True
               Else
                  lblDsp(1) = strExc(1)
               End If
            End If
         Case 2
            If txtLE(Index) <> "" Then
               If ChkDate(txtLE(Index)) = False Then
                  Cancel = True
               End If
            End If
         Case 5, 6
            If txtLE(Index) <> "" Then
               If Right(txtLE(Index), 2) > 12 Then
                  ShowMsg "償還月份不可超過12 !"
                  Cancel = True
               End If
               If Cancel = False And txtLE(5) <> "" And txtLE(6) <> "" Then
                  If RunNick(txtLE(5), txtLE(6)) Then
                     Cancel = True
                  End If
               End If
            End If
      End Select
      
      If Cancel = True Then TextInverse txtLE(Index)
      
      '若是按確定的檢查時略過, 檢查代號檔
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
            Case 6
               '預設目前餘額
               If Val(txtLE(8)) = 0 Then
                  txtLE(8) = Val(txtLE(3)) + Val(txtLE(4))
               End If
               '計算每月償還金額
               If Val(txtLE(7)) = 0 Then
                  m_month = 1 + (Mid(Val(txtLE(6) + 191100), 1, 4) * 12 + Mid(Val(txtLE(6) + 191100), 5, 2)) - (Mid(Val(txtLE(5) + 191100), 1, 4) * 12 + Mid(Val(txtLE(5) + 191100), 5, 2))
                  'Modify by Morgan 2010/10/12 無條件捨去--辜
                  'txtLE(7) = Round((Val(txtLE(3)) + Val(txtLE(4))) / m_month, 0)
                  txtLE(7) = (Val(txtLE(3)) + Val(txtLE(4))) \ m_month
               End If
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
   stSQL = "delete from Loan_Employee where le01='" & txtLE(1) & "' and le02=" & DBDATE(txtLE(2))
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtLE(1).Tag = ""
   txtLE(2).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("代號", "姓名", "貸款日期", "貸款本金", "利　息", "償還起", "償還迄", "每月金額", "目前餘額", "最近償還")
   arrGridHeadWidth = Array(600, 800, 800, 1000, 700, 650, 650, 850, 850, 800)
   grd1.Visible = False
   grd1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1.Cols - 1
      grd1.row = 0
      grd1.col = iRow
      grd1.Text = arrGridHeadText(iRow)
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1.CellAlignment = flexAlignCenterCenter
      'Added by Morgan 2013/10/17
      If iRow > 2 Then
         grd1.ColAlignment(iRow) = flexAlignRightCenter
      Else
         grd1.ColAlignment(iRow) = flexAlignCenterCenter
      End If
      'end 2013/10/17
   Next
   grd1.Visible = True
End Sub

