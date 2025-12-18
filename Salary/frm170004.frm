VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170004 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月獎金資料"
   ClientHeight    =   5052
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8172
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5052
   ScaleWidth      =   8172
   Begin TabDlg.SSTab SSTab1 
      Height          =   4320
      Left            =   45
      TabIndex        =   10
      Top             =   690
      Width           =   8115
      _ExtentX        =   14330
      _ExtentY        =   7620
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm170004.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(17)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDsp(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDsp(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textCUID"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtMB(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtMB(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtMB(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtMB(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtMB(12)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtMB(11)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboMB14"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtMB(13)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtNHI10"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtNet"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm170004.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkSales"
      Tab(1).Control(1)=   "txtComp"
      Tab(1).Control(2)=   "txtSum(2)"
      Tab(1).Control(3)=   "txtSum(1)"
      Tab(1).Control(4)=   "txtSum(0)"
      Tab(1).Control(5)=   "GRD1"
      Tab(1).Control(6)=   "cmdok"
      Tab(1).Control(7)=   "txt1(1)"
      Tab(1).Control(8)=   "txt1(0)"
      Tab(1).Control(9)=   "txt1(2)"
      Tab(1).Control(10)=   "txt1(3)"
      Tab(1).Control(11)=   "Label1(6)"
      Tab(1).Control(12)=   "lblComp"
      Tab(1).Control(13)=   "Label8"
      Tab(1).Control(14)=   "Label7"
      Tab(1).Control(15)=   "Label6"
      Tab(1).Control(16)=   "Label5"
      Tab(1).Control(17)=   "Label12"
      Tab(1).Control(18)=   "Label16"
      Tab(1).ControlCount=   19
      Begin VB.CheckBox chkSales 
         Caption         =   "智權部"
         Height          =   252
         Left            =   -69552
         TabIndex        =   44
         Top             =   744
         Width           =   948
      End
      Begin VB.TextBox txtComp 
         Alignment       =   2  '置中對齊
         Height          =   270
         Left            =   -73890
         MaxLength       =   1
         TabIndex        =   19
         Top             =   720
         Width           =   315
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   2
         Left            =   -68250
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   3885
         Width           =   1005
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   1
         Left            =   -70455
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3885
         Width           =   1005
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   0
         Left            =   -72750
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   3885
         Width           =   1005
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   270
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   33
         Text            =   "8888888"
         Top             =   3330
         Width           =   915
      End
      Begin VB.TextBox txtNHI10 
         Height          =   270
         Left            =   1875
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "120000"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtMB 
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2115
         MaxLength       =   7
         TabIndex        =   6
         Text            =   "1020101"
         Top             =   2370
         Width           =   765
      End
      Begin VB.ComboBox cboMB14 
         Height          =   276
         ItemData        =   "frm170004.frx":0038
         Left            =   1485
         List            =   "frm170004.frx":003F
         TabIndex        =   1
         Top             =   780
         Width           =   3075
      End
      Begin VB.TextBox txtMB 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   270
         Index           =   11
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "1"
         Top             =   2100
         Width           =   315
      End
      Begin VB.TextBox txtMB 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   270
         Index           =   12
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "666666"
         Top             =   2700
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm170004.frx":0063
         Height          =   2835
         Left            =   -75000
         TabIndex        =   26
         Top             =   1005
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   5017
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "獎金日期|員工代號|姓名　　|公司別|獎金總額|扣繳稅額|補充保費|代扣日期|獎金名目　　　"
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
         _Band(0).Cols   =   9
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   400
         Left            =   -68280
         TabIndex        =   20
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72825
         MaxLength       =   7
         TabIndex        =   16
         Top             =   405
         Width           =   825
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73890
         MaxLength       =   7
         TabIndex        =   15
         Top             =   405
         Width           =   825
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -70680
         MaxLength       =   6
         TabIndex        =   17
         Top             =   405
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -69570
         MaxLength       =   6
         TabIndex        =   18
         Top             =   405
         Width           =   915
      End
      Begin VB.TextBox txtMB 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   4
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   4
         Text            =   "2222222"
         Top             =   1815
         Width           =   915
      End
      Begin VB.TextBox txtMB 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   3
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   3
         Text            =   "8888888"
         Top             =   1500
         Width           =   915
      End
      Begin VB.TextBox txtMB 
         Height          =   270
         Index           =   2
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "999999"
         Top             =   465
         Width           =   915
      End
      Begin VB.TextBox txtMB 
         Height          =   285
         Index           =   1
         Left            =   1470
         MaxLength       =   7
         TabIndex        =   2
         Text            =   "1020101"
         Top             =   1125
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別： "
         Height          =   180
         Index           =   6
         Left            =   -74760
         TabIndex        =   43
         Top             =   765
         Width           =   945
      End
      Begin MSForms.Label lblComp 
         Height          =   315
         Left            =   -73410
         TabIndex        =   42
         Top             =   765
         Width           =   2010
         VariousPropertyBits=   27
         Caption         =   "lblComp"
         Size            =   "3545;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   195
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3960
         Width           =   5700
         VariousPropertyBits=   671105055
         Size            =   "7223;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "合計："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74595
         TabIndex        =   38
         Top             =   3900
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "補充保費："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -69240
         TabIndex        =   37
         Top             =   3900
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "扣繳稅額："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71445
         TabIndex        =   36
         Top             =   3900
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "獎金總額："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73740
         TabIndex        =   35
         Top             =   3900
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "給付淨額："
         Height          =   180
         Index           =   5
         Left            =   510
         TabIndex        =   34
         Top             =   3375
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "給付/代扣時間：                     (格式：HHMMSS)"
         Height          =   180
         Index           =   8
         Left            =   510
         TabIndex        =   32
         Top             =   3045
         Width           =   3630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "補充保費代扣日期："
         Height          =   180
         Left            =   510
         TabIndex        =   31
         Top             =   2415
         Width           =   1620
      End
      Begin MSForms.Label lblDsp 
         Height          =   315
         Index           =   2
         Left            =   1890
         TabIndex        =   30
         Top             =   2145
         Width           =   2010
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3545;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別： "
         Height          =   180
         Index           =   4
         Left            =   510
         TabIndex        =   29
         Top             =   2145
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補充保費："
         Height          =   180
         Index           =   3
         Left            =   510
         TabIndex        =   28
         Top             =   2745
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "獎金名目："
         Height          =   180
         Left            =   510
         TabIndex        =   27
         Top             =   825
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   $"frm170004.frx":0078
         ForeColor       =   &H000000FF&
         Height          =   1224
         Index           =   2
         Left            =   3648
         TabIndex        =   25
         Top             =   1236
         Width           =   4356
      End
      Begin MSForms.Label lblDsp 
         Height          =   285
         Index           =   1
         Left            =   2475
         TabIndex        =   24
         Top             =   495
         Width           =   750
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1323;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "獎金日期：                   －"
         Height          =   180
         Left            =   -74760
         TabIndex        =   22
         Top             =   450
         Width           =   1935
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "員工代號：                      －"
         Height          =   180
         Left            =   -71640
         TabIndex        =   21
         Top             =   450
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "扣繳稅額："
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   14
         Top             =   1860
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "獎金總額："
         Height          =   180
         Index           =   17
         Left            =   510
         TabIndex        =   13
         Top             =   1545
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   12
         Top             =   495
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "給付日期："
         Height          =   180
         Left            =   495
         TabIndex        =   11
         Top             =   1170
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
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
            Picture         =   "frm170004.frx":018F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170004.frx":04AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170004.frx":07C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170004.frx":09A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170004.frx":0CBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170004.frx":0FDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170004.frx":12F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170004.frx":1613
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170004.frx":192F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170004.frx":1C4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170004.frx":1F67
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8172
      _ExtentX        =   14415
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
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frm170004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/20 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2008/12/25 add by sonia
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_MB As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim m_bConfirmCheck As Boolean
Dim m_bActived As Boolean
'Added by Morgan 2013/1/25
Dim stNHI() As String
Dim m_arrMB() As String
Dim stLstAddDate As String, stLstAddMB14 As String, stLstNHI10 As String
Dim m_bolLimited As Boolean 'Added by Morgan 2013/2/21 '限制只對A公司

Private Sub cboMB14_GotFocus()
   OpenIme
End Sub

'Added by Morgan 2013/1/29
Private Sub cboMB14_Validate(Cancel As Boolean)
   If m_EditMode = 1 Or m_EditMode = 2 Then
      'Modified by Morgan 2013/2/4 台一投資固定給付時扣補充保費
      'Removed by Morgan 2015/2/9 取消--劉怡��
      'If txtMB(11) = "A" Then
      '   If txtNHI10 = "" Then
      '      txtNHI10 = ServerTime
      '   End If
      'Else
      'end 2015/2/9
         'Added by Morgan 2013/3/8
         '業務獎金預設為4,8,12月的最後一天的235940
         If cboMB14.Text = cboMB14.List(0) Then
            '新增且未設定日期狀態下
            If m_EditMode = 1 And txtMB(1) = "" Then
               intI = Val(Mid(strSrvDate(1), 5, 2))
               If intI < 5 Then
                  txtMB(1) = (strSrvDate(2) \ 10000 - 1) & "1231"
               ElseIf intI < 9 Then
                  txtMB(1) = (strSrvDate(2) \ 10000) & "0430"
               Else
                  txtMB(1) = (strSrvDate(2) \ 10000) & "0831"
               End If
               txtNHI10 = 235940
            End If
         End If
         'end 2013/3/8
         If cboMB14.Tag <> cboMB14.Text Then
            SetNHI06
         End If
      'End If 'Removed by Morgan 2015/2/9 取消--劉怡��
   End If
   cboMB14.Tag = cboMB14.Text
End Sub

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
   
   stCon = ""
   
   'Added by Morgan 2013/2/21
   If m_bolLimited Then
      stCon = stCon & " and mb11='A'"
   End If
   'end 2013/2/21
   
   'Added by Morgan 2023/1/7
   If txtComp <> "" Then
      stCon = stCon & " and mb11='" & txtComp & "'"
   End If
   'end 2023/1/7
   
   'Added by Morgan 2024/6/11
   If chkSales.Value = vbChecked Then
      stCon = stCon & " and st15 like 'S%'"
   End If
   'end 2024/6/11
   
   If txt1(0) <> "" Then
      'Modified by Morgan 2013/1/25
      'stCon = stCon & " and mb01>=" & Val(txt1(0)) + 191100
      stCon = stCon & " and mb01>=" & DBDATE(txt1(0))
   End If
   If txt1(1) <> "" Then
      'Modified by Morgan 2013/1/25
      'stCon = stCon & " and mb01<=" & Val(txt1(1)) + 191100
      stCon = stCon & " and mb01<=" & DBDATE(txt1(1))
   End If
   If txt1(2) <> "" Then
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'stCon = stCon & " and replace(mb02,'A','0')>='" & txt1(2) & "' "
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      stCon = stCon & " and substr(mb02,1,2)||replace(substr(mb02,3,1),'A','0')||substr(mb02,4)>='" & txt1(2) & "' "
   End If
   If txt1(3) <> "" Then
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'stCon = stCon & " and replace(mb02,'A','0')<='" & txt1(3) & "' "
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      stCon = stCon & " and substr(mb02,1,2)||replace(substr(mb02,3,1),'A','0')||substr(mb02,4)<='" & txt1(3) & "' "
   End If
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'strExc(0) = "SELECT mb01-191100 獎金年月,mb02 員工代號,ST02 姓名,mb03 獎金總額,mb04 扣繳稅額 FROM MonthBonus,staff " & _
               " where replace(mb02,'A','0')=st01(+) " & stCon & " order by mb01,mb02"
   'Modified by Morgan 2013/2/4 +公司別、補充保費、代扣日期、獎金名目
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   strExc(0) = "SELECT mb01-19110000 獎金日期,mb02 員工代號,ST02 姓名,mb11 公司別,mb03 獎金總額,mb04 扣繳稅額" & _
               ",mb12 補充保費,decode(sign(mb13),1,mb13-19110000) 代扣日期,mb14 獎金名目 FROM MonthBonus,staff " & _
               " where substr(mb02,1,2)||replace(substr(mb02,3,1),'A','0')||substr(mb02,4)=st01(+)" & stCon & " order by mb01,mb02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set GRD1.Recordset = RsTemp.Clone
      'Modified by Morgan 2013/2/4
      'GRD1.FormatString = GRD1.FormatString
      GRD1.FormatString = "獎金日期|員工代號|姓名　　|公司別|獎金總額|扣繳稅額|補充保費|代扣日期|獎金名目　　　"
      GRD1.ColAlignment(3) = 3
      GRD1.ColAlignment(4) = 7
      GRD1.ColAlignment(5) = 7
      GRD1.ColAlignment(6) = 7
      'end 2013/2/4
      
      'Added by Morgan 2013/9/5
      If RsTemp.RecordCount > 0 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            txtSum(0) = Val(txtSum(0)) + Val("" & RsTemp("獎金總額"))
            txtSum(1) = Val(txtSum(1)) + Val("" & RsTemp("扣繳稅額"))
            txtSum(2) = Val(txtSum(2)) + Val("" & RsTemp("補充保費"))
            RsTemp.MoveNext
         Loop
         txtSum(0) = Format(txtSum(0), "#,##0")
         txtSum(1) = Format(txtSum(1), "#,##0")
         txtSum(2) = Format(txtSum(2), "#,##0")
      End If
      'end 2013/9/5
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
   
   'Added by Morgan 2013/2/21
   If strUserNum = "86021" Then
      m_bolLimited = True
   End If
   'end 2013/2/21
   
   InitialField
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      Form_KeyDown vbKeyF2, 0
   End If
   UpdateToolbarState
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170004 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from MonthBonus where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_MB = .Fields.Count
      ReDim m_FieldList(TF_MB) As FIELDITEM
      For Each oText In txtMB
         idx = oText.Index
         m_FieldList(idx).fiName = "MB" & Format(idx, "00")
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
   
   'Added by Morgan 2013/1/25
   m_FieldList(14).fiName = "MB14"
   m_FieldList(14).fiType = 0
   
   ReDim m_arrMB(TF_MB) As String
   ReDim stNHI(TF_NHI) As String
   'end 2013/1/25
   
   lblComp = "" 'Added by Morgan 2023/1/7
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
Dim stKey01 As String
Dim stKey02 As String
Dim adoRst As New ADODB.Recordset
Dim stCon As String
   
   'Added by Morgan 2013/2/21
   If m_bolLimited Then
      stCon = " and mb11='A'"
   End If
   'end 2013/2/21
   
   'Modified by Morgan 2013/1/25
   'stKey01 = Val(txtMB(1)) + 191100
   stKey01 = DBDATE(txtMB(1))
   'end 2013/1/25
   stKey02 = txtMB(2)
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM MonthBonus,nhi2nd" & _
            " WHERE mb01 = '" & stKey01 & "' and mb02= '" & stKey02 & "' and nhi01(+)=mb02 and nhi14(+)=mb01 and nhi03(+)='50' and nhi04(+)=decode(mb13,null,'2','6')" & stCon
      Case -2
         strExc(0) = "SELECT * FROM MonthBonus,nhi2nd where nhi01(+)=mb02 and nhi14(+)=mb01 and nhi03(+)='50' and nhi04(+)=decode(mb13,null,'2','6')" & stCon & " order by mb01 ASC,mb02 ASC"
      Case -1
         strExc(0) = "SELECT * FROM MonthBonus,nhi2nd" & _
            " WHERE mb01||mb02 <'" & stKey01 & stKey02 & "' and nhi01(+)=mb02 and nhi14(+)=mb01 and nhi03(+)='50' and nhi04(+)=decode(mb13,null,'2','6')" & stCon & " order by mb01 DESC,mb02 DESC"
      Case 1
         strExc(0) = "SELECT * FROM MonthBonus,nhi2nd" & _
            " WHERE mb01||mb02 >'" & stKey01 & stKey02 & "' and nhi01(+)=mb02 and nhi14(+)=mb01 and nhi03(+)='50' and nhi04(+)=decode(mb13,null,'2','6')" & stCon & " order by mb01 ASC,mb02 ASC"
      Case 2
         strExc(0) = "SELECT * FROM MonthBonus,nhi2nd where nhi01(+)=mb02 and nhi14(+)=mb01 and nhi03(+)='50' and nhi04(+)=decode(mb13,null,'2','6')" & stCon & " order by mb01 DESC,mb02 DESC"
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
      txtMB(2).SetFocus
      txtMB_GotFocus 1
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
            If txtMB(1).Locked = False Then
               txtMB(1).Text = GRD1.TextMatrix(lCurRow, 0)
               txtMB(2).Text = GRD1.TextMatrix(lCurRow, 1)
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

Private Sub txtComp_Change()
   lblComp = ""
End Sub

Private Sub txtComp_GotFocus()
   TextInverse txtComp
End Sub

Private Sub txtComp_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
'cancel by sonia 2023/2/4 不能輸J公司故取消
'   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
'      KeyAscii = 0
'      Beep
'   End If
End Sub

Private Sub txtComp_Validate(Cancel As Boolean)
   lblComp = CompNameQuery(txtComp)
   'add by sonia 2023/2/4
   If txtComp <> "" And lblComp = "" Then
      ShowMsg "公司別錯誤 !"
      txtComp.SetFocus
      txtComp_GotFocus
      Cancel = True
   End If
   'end 2023/2/4
End Sub

Private Sub txtMB_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtMB(Index)
End Sub

Private Sub ClearField()
   For Each oText In txtMB
      oText.Text = Empty
   Next
   'Added by Morgan 2013/1/29
   cboMB14.ListIndex = -1
   txtNHI10 = ""
   'end 2013/1/29
   
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_MB
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   txtNet = ""
   
   m_bConfirmCheck = False
   
   'Added by Morgan 2013/1/25
   Erase m_arrMB
   ReDim m_arrMB(TF_MB) As String
   'end 2013/1/25
   
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   With p_Rst
   If .RecordCount > 0 Then
      For Each oText In txtMB
         idx = oText.Index
         '獎金年月轉民國年月
         If idx = 1 Then
            'Modified by Morgan 2013/1/25
            'm_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName) - 191100
            m_FieldList(idx).fiOldData = TransDate(.Fields(m_FieldList(idx).fiName), 1)
         'Added by Morgan 2013/1/31
         ElseIf idx = 13 Then
            m_FieldList(idx).fiOldData = TransDate("" & .Fields(m_FieldList(idx).fiName), 1)
         'end 2013/1/31
         Else
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         End If
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         oText.Text = m_FieldList(idx).fiOldData
         'oText.Tag = m_FieldList(idx).fiOldData
         
         m_arrMB(idx) = oText.Text 'Added by Morgan 2013/1/25
      Next
      
      If ClsPDGetStaffN(txtMB(2), strExc(1), , True) Then
         lblDsp(1) = strExc(1)
      End If
      
      'Added by Morgan 2013/1/29
      If txtMB(11) <> "" Then
         lblDsp(2) = CompNameQuery(txtMB(11))
      End If
      cboMB14.Text = "" & .Fields("MB14")
      txtNHI10 = "" & .Fields("nhi10")
      txtNHI10.Tag = txtNHI10
      'end 2013/1/29
      
      CUID(1) = "" & .Fields("mb05")
      CUID(2) = "" & .Fields("mb06")
      CUID(3) = "" & .Fields("mb07")
      CUID(4) = "" & .Fields("mb08")
      CUID(5) = "" & .Fields("mb09")
      CUID(6) = "" & .Fields("mb10")
   End If
   End With
   UpdateCUID CUID, textCUID
   txtMB(1).Tag = txtMB(1)
   txtMB(2).Tag = txtMB(2)
   SetNet 'Added by Morgan 2013/2/22
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtMB
      oText.Locked = bLocked
   Next
   'Added by Morgan 2013/1/29
   txtNHI10.Locked = bLocked
   cboMB14.Locked = bLocked
   'end 2013/1/29
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
         'Added by Morgan 2013/1/30
         '預設前一筆的日期及名目
         txtMB(1) = stLstAddDate
         cboMB14 = stLstAddMB14
         'Added by Morgan 2013/5/3
         If cboMB14 = cboMB14.List(0) Then
            txtNHI10 = stLstNHI10
         End If
         If txtMB(1) <> "" Then
         txtMB(2).SetFocus
         End If
         'end 2013/1/30

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
            txtMB(1) = txtMB(1).Tag
            txtMB(2) = txtMB(2).Tag
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
         If m_bUpdate And txtMB(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtMB(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtMB(1) <> "" Then
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
         txtMB(1).Locked = False
         'Added by Morgan 2013/1/30
         txtMB(2).Locked = False
         cboMB14.Locked = False
         txtNHI10.Locked = False
         'end 2013/1/30
         
         If Me.Visible = True Then
            'Modified by Morgan 2013/5/3
            'txtMB(1).SetFocus
            txtMB(2).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 2
         txtMB(1).Locked = True
         'Added by Morgan 2013/1/30
         txtMB(2).Locked = True
         If txtNHI10 <> "" Then
            'Modified by Morgan 2013/2/4 台一投資固定給付時扣補充保費
            'Removed by Morgan 2015/2/9 取消--劉怡��
            'If txtMB(11) <> "A" Then
               cboMB14.Locked = True
            'End If
            txtNHI10.Locked = True
         End If
         'end 2013/1/30
         If Me.Visible = True Then
            txtMB(3).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 4
         txtMB(1).Locked = False
         txtMB(2).Locked = False
         If Me.Visible = True Then
            'Modified by Morgan 2013/5/3
            'txtMB(1).SetFocus
            txtMB(2).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case Else
         txtMB(1).Locked = True
         txtMB(2).Locked = True 'Added by Morgan 2013/1/25
         If Me.Visible = True Then
            'Modified by Morgan 2013/5/3
            'txtMB(1).SetFocus
            txtMB(2).SetFocus
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
         'Added by Morgan 2013/3/7
         '代扣日的月薪資若已計算則不可再異動
         If txtMB(13) <> "" Then
            If PUB_ExistsSalaryMonth(txtMB(13)) = True Then
               ShowMsg "補充保費代扣日期之月薪資已計算不可刪除!"
               Exit Function
            End If
         
         ElseIf ChkHi2ndIsPaid(Left(DBDATE(txtMB(1)), 6)) = True Then
               MsgBox "給付日期月份的補充保費已繳納不可刪除！", vbExclamation
               Exit Function
         End If
         'end 2013/3/7
      
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
               'Modified by Morgan 2013/5/3
               'txtMB(1).SetFocus
               'txtMB_GotFocus 1
               txtMB(2).SetFocus
               txtMB_GotFocus 2
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   For Each oText In txtMB
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtMB_Validate idx, bCancel
         If bCancel = True Then
            txtMB(idx).SetFocus
            txtMB_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   '查詢
   If m_EditMode = 4 Then
      If txtMB(2) = "" Then
         ShowMsg "請輸入員工代號 !"
         txtMB(2).SetFocus
         txtMB_GotFocus 2
         GoTo EscPoint
      End If
      If txtMB(1) = "" Then
         'Modified by Morgan 2013/1/25
         'ShowMsg "請輸入獎金年月 !"
         ShowMsg "請輸入獎金給付日期 !"
         txtMB(1).SetFocus
         txtMB_GotFocus 1
         GoTo EscPoint
      End If

   '維護
   Else
      If txtMB(2) = "" And txtMB(2).Locked = False Then
         ShowMsg "請輸入員工代號 !"
         txtMB(2).SetFocus
         txtMB_GotFocus 2
         GoTo EscPoint
      End If
      
      If txtMB(1) = "" And txtMB(1).Locked = False Then
         'Modified by Morgan 2013/1/25
         'ShowMsg "請輸入獎金年月 !"
         ShowMsg "請輸入獎金給付日期 !"
         txtMB(1).SetFocus
         txtMB_GotFocus 1
         GoTo EscPoint
      End If
      
      'Added by Morgan 2023/8/24
      '檢查給付日期當天是否在職
      If ChkStaffST04(txtMB(2), False, txtMB(1)) = True Then
         MsgBox "給付日期當天必須在職！", vbCritical
         GoTo EscPoint
      End If
      'end 2023/8/24
      
      If txtMB(3) = "" And txtMB(3).Locked = False Then
         ShowMsg "請輸入獎金總額 !"
         txtMB(3).SetFocus
         txtMB_GotFocus 3
         GoTo EscPoint
      End If
      
      'Added by Morgan 2013/1/28
      If cboMB14 = "" Then
         ShowMsg "請輸入獎金名目 !"
         cboMB14.SetFocus
         GoTo EscPoint
      End If
      
      'Modified by Morgan 2013/2/4 台一投資固定給付時扣補充保費
      'Removed by Morgan 2015/2/9 取消--劉怡��
      If txtNHI10 = "" And cboMB14.Text = cboMB14.List(0) Then
      'If txtNHI10 = "" And (cboMB14.Text = cboMB14.List(0) Or txtMB(11) = "A") Then
      
         ShowMsg "請輸入給付時間以便計算補充保費 !"
         txtNHI10.SetFocus
         txtNHI10_GotFocus
         GoTo EscPoint
      End If
      
      'Added by Morgan 2013/2/21
      If cboMB14.Text = cboMB14.List(0) And txtMB(11) <> "A" Then
         'Modified by Morgan 2013/3/8 必須為 4/30,8/31,12/31
         strExc(0) = Right(txtMB(1), 4)
         If strExc(0) <> "0430" And strExc(0) <> "0831" And strExc(0) <> "1231" Then
            ShowMsg cboMB14 & "的給付月日必須是0430、0831或1231!"
            txtMB(1).SetFocus
            GoTo EscPoint
         End If
         If txtNHI10 <> "235940" Then
            ShowMsg cboMB14 & "的給付時間必須是235940!"
            GoTo EscPoint
         End If
         'end 2013/3/8
         '若為不同年度之資料則只能是前一年度的最後一天
         If Val(txtMB(1)) \ 10000 <> Val(strSrvDate(2)) \ 10000 Then
            If txtMB(1) <> (Val(strSrvDate(2)) \ 10000 - 1) & "1231" Then
               ShowMsg "不同年度之資料則只能是前一年度的最後一天!"
               txtMB(1).SetFocus
               GoTo EscPoint
            End If
         End If
      End If
      
      'Added by Morgan 2013/5/3
      If cboMB14.Text <> cboMB14.List(0) Then
         strExc(1) = GetST15(txtMB(2))
         If Left(strExc(1), 1) = "S" Then
            strExc(0) = Right(txtMB(1), 4)
            If strExc(0) = "0430" Or strExc(0) = "0831" Or strExc(0) = "1231" Then
              MsgBox "非三節競賽獎金時給付月日請勿輸入0430, 0831或1231!!", vbCritical
              txtMB(1).SetFocus
              GoTo EscPoint
            End If
         End If
      End If
      'end 2013/5/3
      
      '代扣日的月薪資若已計算則不可再異動
      If txtMB(13) <> "" Then
         If PUB_ExistsSalaryMonth(txtMB(13)) = True Then
            ShowMsg "補充保費代扣日期之月薪資已計算不可再異動!"
            GoTo EscPoint
         End If
      'Added by Morgan 2013/3/7
      ElseIf ChkHi2ndIsPaid(Left(DBDATE(txtMB(1)), 6)) = True Then
            MsgBox "給付日期月份的補充保費已繳納不可再有異動！", vbExclamation
            GoTo EscPoint
      
      End If
      'end 2013/2/21
      
      SetNHI06
      'end 2013/1/28
   End If
   TxtValidate = True
   
EscPoint:
   m_bConfirmCheck = False
    
End Function

Private Function AddRecord() As Boolean
Dim stCols As String, stValues As String, stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Added by Morgan 2013/2/6
   If txtMB(13) = "" Then
      If PUB_ChkNHi2nd(stNHI(1), stNHI(2), stNHI(10), True) = False Then
         GoTo EscPoint
      End If
      SetNHI06 '此處要重算以避免輸入過程中同時有新增較早資料而沒算到
   End If
   'end 2013/2/6
   
   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtMB
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            '獎金年月轉西元年月
            'If idx = 1 Then
            '   stValues = stValues & "," & CNULL(Val((m_FieldList(idx).fiNewData) + 191100), True)
            'Else
               stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
            'End If
         End If
      End If
   Next
   
   'Added by Morgan 2013/1/30
   stCols = stCols & "," & m_FieldList(14).fiName
   stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(14).fiNewData))
   'end 2013/1/30
   
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO MonthBonus (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
'   stSQL = "select max(mb02) from MonthBonus where mb01='" & txtMB(1) & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
'   If intI = 1 Then
'      txtMB(2) = RsTemp.Fields(0)
'   End If
   
   'Added by Morgan 2013/1/25
   If txtMB(13) = "" Then
      PUB_InsertNHI2nd stNHI
   End If
   'end 2013/1/25
   
   cnnConnection.CommitTrans
   
   AddRecord = True
   
   'Added by Morgan 2013/1/31
   stLstAddDate = txtMB(1)
   stLstAddMB14 = cboMB14
   stLstNHI10 = txtNHI10
   'end 2013/1/31
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
   
EscPoint:

End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Added by Morgan 2013/2/6
   '不可有晚於該筆資料的補充保費
   If txtMB(13) = "" Then
      If PUB_ChkNHi2nd(stNHI(1), stNHI(2), stNHI(10)) = False Then
         GoTo EscPoint
      End If
      SetNHI06
   End If
   'end 2013/2/6
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE MonthBonus SET "
   stSet = ""
   For Each oText In txtMB
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
   
   'Added by Morgan 2013/1/30
   If m_FieldList(14).fiNewData <> m_FieldList(14).fiOldData Then
      bDifference = True
      stSet = stSet & "," & m_FieldList(14).fiName & "=" & CNULL(ChgSQL(m_FieldList(14).fiNewData))
   End If
   'end 2013/1/30
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      'Modified by Morgan 2013/1/25
      'stSQL = stSQL & stSet & " where mb01='" & Val(txtMB(1)) + 191100 & "' and mb02='" & txtMB(2) & "'; end; "
      stSQL = stSQL & stSet & " where mb01='" & DBDATE(txtMB(1)) & "' and mb02='" & txtMB(2) & "'; end; "
      
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
      
      'Added by Morgan 2013/1/25
      If txtMB(13) = "" Then
         PUB_InsertNHI2nd stNHI
      End If
      'end 2013/1/25
      
   End If
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

EscPoint:

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
   For Each oText In txtMB
      idx = oText.Index
      Select Case idx
         Case 1
            'Modified by Morgan 2013/1/25
            'm_FieldList(idx).fiNewData = Val(oText.Text) + 191100
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         'Added by Morgan 2013/2/1
         Case 13
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
   m_FieldList(14).fiNewData = cboMB14.Text 'Added by Morgan 2013/1/30
End Sub

Private Sub txtMB_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 2
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtMB_Validate(Index As Integer, Cancel As Boolean)
Dim m_taxrate As String   '2010/12/30 add by sonia 非固定之薪資所得扣繳稅率
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Select Case Index
         Case 1
            If txtMB(Index) <> "" Then
               'Modified by Morgan 2013/1/25
               'If Right(txtMB(Index), 2) > 16 Then
               '   ShowMsg "獎金月份不可超過16 !"
               '   Cancel = True
               'End If
               If ChkDate(txtMB(Index)) = False Then
                  Cancel = True
               'Added by Morgan 2013/2/18
               ElseIf Val(txtMB(Index)) > Val(strSrvDate(2)) Then
                  MsgBox "給付日期不可晚於系統日！", vbExclamation
                  Cancel = True
               End If
               'end 2013/1/25
            End If
         Case 2
            If txtMB(Index) <> "" Then
               If ChkStaffID(txtMB(Index)) = True Then
                  Cancel = True
               End If
               'Modified by Morgan 2015/3/17 離職改可輸入但提醒--辜說離職當月還是會有獎金要輸入
               'If Cancel = False And ClsPDGetStaff(txtMB(Index), strExc(1), , True) = False Then
               If Cancel = False And ClsPDGetStaffN(txtMB(Index), strExc(1), , True) = False Then
               'end 2015/3/17
                  Cancel = True
               Else
                  lblDsp(1) = strExc(1)
                  
                  ClsPDGetStaff txtMB(Index), strExc(1), , True  'Added by Morgan 2015/3/17
                  
                  'Added by Morgan 2013/1/25
                  If txtMB(Index) > "F" Then
                     MsgBox "請輸入所內員工編號！", vbExclamation
                     Cancel = True
                  Else
                     '設定薪資公司別
                     strExc(0) = "select sd19 from salarydata where sd01='" & txtMB(Index) & "'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        'Added by Morgan 2013/2/21
                        If m_bolLimited And RsTemp(0) <> "A" Then
                           MsgBox "只可輸入A公司員工！", vbExclamation
                           Cancel = True
                        Else
                        'end 2013/2/21
                        
                           txtMB(11) = "" & RsTemp(0)
                           If txtMB(11) <> "" Then
                              lblDsp(2) = CompNameQuery(txtMB(11))
                           End If
                           
                        End If 'Added by Morgan 2013/2/21
                     Else
                        MsgBox "員工輸入錯誤，無法讀取公司別資料！", vbExclamation
                        Cancel = True
                     End If
                  End If
                  'End 2013/1/25
               End If
            End If
         Case 3
            'Modified by Morgan 2013/10/17 獎金總額有變都要重算--辜
            'If txtMB(index) <> "" And txtMB(4) = "" Then
            If txtMB(Index) <> txtMB(Index).Tag Then
               'Added by Morgan 2016/6/24
               '其它薪資 代號"50" 含年終/三節/翻譯等 所得73,001 以上扣稅
               'modify by sonia 2018/4/17 改84,501
               'If Val(txtMB(Index)) > 73001 Then
               'Modified by Morgan 2024/8/28 113年起扣標準為 88501元(含) --婉莘
               'If Val(txtMB(Index)) > 84501 Then
               If Val(txtMB(Index)) >= 88501 Then
               'end 2016/6/24
               
                  '2010/12/30 modify by sonia 非固定之薪資所得扣繳稅率改抓 翻譯所得oc01='01'的稅率
                  'txtMB(4) = Round(Val(txtMB(Index)) * 6 / 100, 0)
                  m_taxrate = 0
                  strExc(0) = "select oc04 from OtherSalaryCode where oc01='01'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     m_taxrate = "" & RsTemp.Fields(0)
                  End If
                  txtMB(4) = Round(Val(txtMB(Index)) * m_taxrate / 100, 0)
                  '2010/12/30 end
                  
               'Modified by Morgan 2016/6/24
                  'If txtMB(4) <= 2000 Then txtMB(4) = ""
               End If
               'end 2016/6/24
            End If
            txtMB(Index).Tag = txtMB(Index) 'Added by Morgan 2013/10/17
            
'         Case 4
'            If txtMB(Index) = "" Then
'               txtMB(Index) = Round(Val(txtMB(3)) * 6 / 100, 0)
'               If txtMB(Index) <= 2000 Then txtMB(Index) = ""
'            Else
'               If txtMB(Index) <> Round(Val(txtMB(3)) * 6 / 100, 0) Then
'                  ShowMsg "扣繳稅額不等於獎金總額*6%, 應為 " & Round(Val(txtMB(3)) * 6 / 100, 0) & " !"
'                  Cancel = True
'               End If
'            End If
      End Select
      
      If Cancel = True Then TextInverse txtMB(Index)
      
      '若是按確定的檢查時略過, 檢查代號檔
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
         'Added by Morgan 2013/1/25
         Case 1, 2, 3
            If m_arrMB(Index) <> txtMB(Index) Then
               SetNHI06
            End If
         'end 2013/1/25
         End Select
      End If
      
      'Added by Morgan 2013/1/25
      If Cancel = False Then
         m_arrMB(Index) = txtMB(Index)
      End If
      'end 2013/1/25
   End If
End Sub

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Added by Morgan 2013/2/6
   '檢查不可有晚於該筆資料的補充保費
   'SetNHI06 'Removed by Morgan 2020/8/26 刪除不必計算,否則代扣日期可能會改到而導致找不到補充保費紀錄可刪
   If PUB_ChkNHi2nd(stNHI(1), stNHI(2), stNHI(10)) = False Then
      GoTo EscPoint
   End If
   'end 2013/2/6
   
   '刪除
   'Modified by Morgan 2013/1/25
   'stSQL = "delete from MonthBonus where mb01='" & Val(txtMB(1)) + 191100 & "' and mb02='" & txtMB(2) & "'"
   stSQL = "delete from MonthBonus where mb01='" & DBDATE(txtMB(1)) & "' and mb02='" & txtMB(2) & "'"
   'end 2013/1/25
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   'Added by Morgan 2013/1/24
   '刪除補充保費
   strSql = "DELETE NHI2ND WHERE NHI01='" & txtMB(2) & "' AND NHI02=" & IIf(txtMB(13) = "", DBDATE(txtMB(1)), DBDATE(txtMB(13))) & " AND NHI03='50' AND NHI04='" & IIf(txtMB(13) = "", "2", "6") & "'"
   cnnConnection.Execute strSql, intI
   'end 2013/1/24
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtMB(1).Tag = ""
   txtMB(2).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
    
EscPoint:
End Function
'Added by Morgan 2013/1/29
'計算補充保費
Private Sub SetNHI06()
   
   txtMB(12) = ""
   txtMB(13) = ""
      
   'Modified by Morgan 2013/2/4 台一投資固定給付時扣補充保費
   'Modified by Morgan 2015/2/9 取消--劉怡文
   'If txtMB(11) <> "A" And cboMB14.Text <> cboMB14.List(0) Then
   If cboMB14.Text <> cboMB14.List(0) Then
   'end 2015/2/9
      strExc(1) = DBDATE(txtMB(1))
      strExc(2) = Left(strExc(1), 4) '年
      strExc(3) = Mid(strExc(1), 5, 2) '月
      '1~4月,4月薪資扣
      If Val(strExc(3)) <= 4 Then
         txtMB(13) = TransDate(GetLastDay(strExc(2) & "0401"), 1)
      '5~8月,8月薪資扣
      ElseIf Val(strExc(3)) <= 8 Then
         txtMB(13) = TransDate(GetLastDay(strExc(2) & "0801"), 1)
      '9~12月,12月薪資扣
      Else
         txtMB(13) = TransDate(GetLastDay(strExc(2) & "1201"), 1)
      End If
      txtNHI10 = ""
   Else
      stNHI(1) = txtMB(2)
      stNHI(2) = DBDATE(txtMB(1))
      stNHI(3) = "50"
      stNHI(4) = "2"
      stNHI(5) = ""
      stNHI(6) = ""
      stNHI(7) = txtMB(3)
      stNHI(8) = ""
      stNHI(10) = txtNHI10
      stNHI(11) = txtMB(11) 'Added by Morgan 2013/2/26
      If stNHI(1) <> "" And stNHI(2) <> "" Then
         PUB_NHI2nd stNHI(1), stNHI(2), stNHI(3), stNHI(4), stNHI(7), stNHI(5), stNHI(6), stNHI(8), stNHI(10), stNHI(11), stNHI(13) 'Modified by Morgan 2013/3/12 +NHI13
         txtMB(12) = Val(stNHI(6))
      End If
   End If
   SetNet 'Added by Morgan 2013/2/22
End Sub

Private Sub txtNHI10_GotFocus()
   TextInverse txtNHI10
   CloseIme
End Sub

Private Sub txtNHI10_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtNHI10_Validate(Cancel As Boolean)
   If txtNHI10.Tag <> txtNHI10 Then
      SetNHI06
   End If
   txtNHI10.Tag = txtNHI10
End Sub
'end 2013/1/29

'Added by Morgan 2013/2/22
Private Sub SetNet()
   If txtMB(13) = "" Or txtMB(12) <> "" Then
      txtNet = Val(txtMB(3)) - Val(txtMB(4)) - Val(txtMB(12))
   Else
      txtNet = ""
   End If
End Sub
