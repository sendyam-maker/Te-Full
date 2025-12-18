VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170028 
   BorderStyle     =   1  '單線固定
   Caption         =   "股利及退職所得資料"
   ClientHeight    =   5052
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8220
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5052
   ScaleWidth      =   8220
   Begin TabDlg.SSTab SSTab1 
      Height          =   4380
      Left            =   45
      TabIndex        =   17
      Top             =   645
      Width           =   8115
      _ExtentX        =   14309
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm170028.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDsp(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(9)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDsp(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(10)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(11)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(12)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(13)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(14)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(15)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(16)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textCUID"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtBR(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtBR(2)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtBR(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtBR(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtBR(8)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtBR(9)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtBR(7)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtBR(5)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtBR(6)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtBR(10)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtBR(11)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtBR(12)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtBR(13)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtBR(20)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtNet"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtBR(21)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtTot"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).ControlCount=   37
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm170028.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdok"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt1(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt1(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt1(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txt1(3)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "GRD1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtSum(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtSum(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.TextBox txtTot 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   270
         Left            =   6585
         MaxLength       =   10
         TabIndex        =   50
         Text            =   "8888888888"
         Top             =   2640
         Width           =   1100
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   0
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   3960
         Width           =   1005
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   1
         Left            =   -70545
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3960
         Width           =   1005
      End
      Begin VB.TextBox txtBR 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   270
         Index           =   21
         Left            =   4050
         MaxLength       =   6
         TabIndex        =   14
         Text            =   "999999"
         Top             =   3600
         Width           =   700
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   270
         Left            =   6555
         MaxLength       =   10
         TabIndex        =   15
         Text            =   "8888888888"
         Top             =   3600
         Width           =   1100
      End
      Begin VB.TextBox txtBR 
         Height          =   270
         Index           =   20
         Left            =   1530
         MaxLength       =   7
         TabIndex        =   13
         Text            =   "1020101"
         Top             =   3600
         Width           =   810
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm170028.frx":0038
         Height          =   2985
         Left            =   -74970
         TabIndex        =   40
         Top             =   840
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   5271
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "年度|月份|所得人代號|名　稱|公司別||格式|所得總額|扣繳稅額"
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
      Begin VB.TextBox txtBR 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   13
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "9999999999"
         Top             =   3305
         Width           =   1100
      End
      Begin VB.TextBox txtBR 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   12
         Left            =   4845
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "12"
         Top             =   2975
         Width           =   400
      End
      Begin VB.TextBox txtBR 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   11
         Left            =   4050
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "96"
         Top             =   2975
         Width           =   500
      End
      Begin VB.TextBox txtBR 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   10
         Left            =   1530
         MaxLength       =   5
         TabIndex        =   9
         Text            =   "99.99"
         Top             =   2975
         Width           =   600
      End
      Begin VB.TextBox txtBR 
         Height          =   270
         Index           =   6
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "12"
         Top             =   1995
         Width           =   405
      End
      Begin VB.TextBox txtBR 
         Height          =   270
         Index           =   5
         Left            =   1530
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "01"
         Top             =   1995
         Width           =   405
      End
      Begin VB.TextBox txtBR 
         Height          =   270
         Index           =   7
         Left            =   1530
         MaxLength       =   7
         TabIndex        =   6
         Text            =   "961231"
         Top             =   2325
         Width           =   735
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -69810
         MaxLength       =   6
         TabIndex        =   23
         Top             =   455
         Width           =   1425
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -71400
         MaxLength       =   6
         TabIndex        =   22
         Top             =   455
         Width           =   1425
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73920
         MaxLength       =   3
         TabIndex        =   20
         Top             =   455
         Width           =   500
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -73140
         MaxLength       =   3
         TabIndex        =   21
         Top             =   455
         Width           =   500
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢(&S)"
         Height          =   345
         Left            =   -68190
         TabIndex        =   24
         Top             =   380
         Width           =   1095
      End
      Begin VB.TextBox txtBR 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   9
         Left            =   4050
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "9999999999"
         Top             =   2645
         Width           =   1100
      End
      Begin VB.TextBox txtBR 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   8
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "9999999999"
         Top             =   2645
         Width           =   1100
      End
      Begin VB.TextBox txtBR 
         Height          =   270
         Index           =   3
         Left            =   1530
         MaxLength       =   1
         TabIndex        =   2
         Text            =   "1"
         Top             =   1365
         Width           =   405
      End
      Begin VB.TextBox txtBR 
         Height          =   270
         Index           =   1
         Left            =   1530
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "96"
         Top             =   405
         Width           =   500
      End
      Begin VB.TextBox txtBR 
         Height          =   270
         Index           =   2
         Left            =   1530
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "999999999999"
         Top             =   720
         Width           =   1425
      End
      Begin VB.TextBox txtBR 
         Height          =   270
         Index           =   4
         Left            =   1512
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "54"
         Top             =   1680
         Width           =   405
      End
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   585
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3960
         Width           =   6690
         VariousPropertyBits=   671105055
         Size            =   "7223;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "給付總額："
         Height          =   180
         Index           =   16
         Left            =   5640
         TabIndex        =   51
         Top             =   2685
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "所得總額："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73830
         TabIndex        =   49
         Top             =   3975
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "扣繳稅額："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71535
         TabIndex        =   48
         Top             =   3975
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "合計："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74685
         TabIndex        =   47
         Top             =   3975
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補充保費： "
         Height          =   180
         Index           =   15
         Left            =   3105
         TabIndex        =   44
         Top             =   3645
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "給付淨額："
         Height          =   180
         Index           =   14
         Left            =   5610
         TabIndex        =   43
         Top             =   3645
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "給付日期："
         Height          =   180
         Index           =   13
         Left            =   600
         TabIndex        =   42
         Top             =   3645
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   $"frm170028.frx":004D
         ForeColor       =   &H000000FF&
         Height          =   540
         Index           =   0
         Left            =   3600
         TabIndex        =   41
         Top             =   1995
         Width           =   4365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "退休自提： "
         Height          =   180
         Index           =   12
         Left            =   600
         TabIndex        =   39
         Top             =   3345
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "股利所屬年度："
         Height          =   180
         Index           =   11
         Left            =   2760
         TabIndex        =   38
         Top             =   3015
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "稅額扣抵比率：               % "
         Height          =   180
         Index           =   10
         Left            =   240
         TabIndex        =   37
         Top             =   3015
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "除權基準日："
         Height          =   180
         Index           =   7
         Left            =   420
         TabIndex        =   36
         Top             =   2355
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "起迄月份：          －"
         Height          =   180
         Index           =   6
         Left            =   600
         TabIndex        =   35
         Top             =   2040
         Width           =   1530
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "所得人代號：                                －"
         Height          =   180
         Left            =   -72480
         TabIndex        =   34
         Top             =   495
         Width           =   2700
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "所得年度：              －"
         Height          =   180
         Left            =   -74880
         TabIndex        =   33
         Top             =   495
         Width           =   1710
      End
      Begin MSForms.Label lblDsp 
         Height          =   300
         Index           =   2
         Left            =   2100
         TabIndex        =   31
         Top             =   1410
         Width           =   1410
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2487;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "可扣抵稅額： "
         Height          =   180
         Index           =   9
         Left            =   2940
         TabIndex        =   30
         Top             =   2685
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "　　股利淨額： "
         Height          =   180
         Index           =   8
         Left            =   240
         TabIndex        =   29
         Top             =   2685
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別： "
         Height          =   180
         Index           =   4
         Left            =   600
         TabIndex        =   28
         Top             =   1395
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所得年度： "
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   27
         Top             =   435
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名　　稱："
         Height          =   180
         Index           =   3
         Left            =   600
         TabIndex        =   26
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所得人代號："
         Height          =   180
         Index           =   2
         Left            =   420
         TabIndex        =   25
         Top             =   765
         Width           =   1080
      End
      Begin MSForms.Label lblDsp 
         Height          =   300
         Index           =   1
         Left            =   1530
         TabIndex        =   19
         Top             =   1080
         Width           =   1410
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2487;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "格式代號：            (54：股利，93：退職所得)"
         Height          =   180
         Index           =   5
         Left            =   600
         TabIndex        =   18
         Top             =   1725
         Width           =   3540
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
            Picture         =   "frm170028.frx":00E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170028.frx":03FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170028.frx":0718
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170028.frx":08F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170028.frx":0C10
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170028.frx":0F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170028.frx":1248
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170028.frx":1564
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170028.frx":1880
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170028.frx":1B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170028.frx":1EB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8220
      _ExtentX        =   14499
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
Attribute VB_Name = "frm170028"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/22 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2008/12/27 add by sonia
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_BR As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim m_bConfirmCheck As Boolean
Dim m_bActived As Boolean
Dim m_bolLimited As Boolean 'Added by Morgan 2013/2/21 '限制只對A公司
Dim stNHI() As String
Dim m_oi02 As String        'add by sonia 2016/1/20 用於判斷104年起之股利發放對象為個人或公司


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
   stCon = ""
   
   'Added by Morgan 2013/2/22
   If m_bolLimited Then
      stCon = stCon & " and br03='A'"
   End If
   'end 2013/2/22

   If txt1(0) <> "" Then
      stCon = stCon & " and br01>=" & Val(txt1(0)) + 1911
   End If
   If txt1(1) <> "" Then
      stCon = stCon & " and br01<=" & Val(txt1(1)) + 1911
   End If
   If txt1(2) <> "" Then
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'stCon = stCon & " and replace(br02,'A','0')>='" & txt1(2) & "' "
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      stCon = stCon & " and substr(br02,1,2)||replace(substr(br02,3,1),'A','0')||substr(br02,4)>='" & txt1(2) & "' "
   End If
   If txt1(3) <> "" Then
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'stCon = stCon & " and replace(br02,'A','0')<='" & txt1(3) & "' "
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      stCon = stCon & " and substr(br02,1,2)||replace(substr(br02,3,1),'A','0')||substr(br02,4)<='" & txt1(3) & "' "
   End If
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   strExc(0) = "SELECT br01-1911,br05||'~'||br06,br02,nvl(oi04,st02),br03,br04,decode(br04,'54','股利','93','退職'),br08,br09 FROM BonusRetire,Otherincomer,staff " & _
               " where br02=oi01(+) and substr(br02,1,2)||replace(substr(br02,3,1),'A','0')||substr(br02,4)=st01(+) " & stCon & " order by br01,br02,br03,br04"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set GRD1.Recordset = RsTemp.Clone
      GRD1.FormatString = GRD1.FormatString
      SetGrd
      
      'Added by Morgan 2013/10/17
      If RsTemp.RecordCount > 0 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            txtSum(0) = Val(txtSum(0)) + Val("" & RsTemp("br08"))
            txtSum(1) = Val(txtSum(1)) + Val("" & RsTemp("br09"))
            RsTemp.MoveNext
         Loop
         txtSum(0) = Format(txtSum(0), "#,##0")
         txtSum(1) = Format(txtSum(1), "#,##0")
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
   
   'Added by Morgan 2013/2/22
   If strUserNum = "86021" Then
      m_bolLimited = True
   End If
   'end 2013/2/22
   
   InitialField
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      Form_KeyDown vbKeyF2, 0
   End If
   UpdateToolbarState
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170028 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from BonusRetire where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_BR = .Fields.Count
      ReDim m_FieldList(TF_BR) As FIELDITEM
      For Each oText In txtBR
         idx = oText.Index
         m_FieldList(idx).fiName = "BR" & Format(idx, "00")
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
   
   ReDim stNHI(TF_NHI) As String
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
Dim stKey01 As String
Dim stKey02 As String
Dim stKey03 As String
Dim stKey04 As String
Dim stCon As String
Dim adoRst As New ADODB.Recordset
   
   stKey01 = Val(txtBR(1)) + 1911
   stKey02 = txtBR(2)
   stKey03 = txtBR(3)
   stKey04 = txtBR(4)
   'Added by Morgan 2013/2/22
   If m_bolLimited Then
      stCon = " and br03='A'"
   End If
   'end 2013/2/22
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM BonusRetire" & _
            " WHERE br01 = '" & stKey01 & "' and br02= '" & stKey02 & "' and br03= '" & stKey03 & "' and br04= '" & stKey04 & "'" & stCon
      Case -2
         strExc(0) = "SELECT * FROM BonusRetire where 1=1" & stCon & " order by 1 ASC,2 ASC,3 ASC,4 ASC"
      Case -1
         strExc(0) = "SELECT * FROM BonusRetire" & _
            " WHERE br01||br02||br03||br04 <'" & stKey01 & stKey02 & stKey03 & stKey04 & "'" & stCon & " order by 1 DESC,2 DESC,3 DESC,4 DESC"
      Case 1
         strExc(0) = "SELECT * FROM BonusRetire" & _
            " WHERE br01||br02||br03||br04 >'" & stKey01 & stKey02 & stKey03 & stKey04 & "'" & stCon & " order by 1 ASC,2 ASC,3 ASC,4 ASC"
      Case 2
         strExc(0) = "SELECT * FROM BonusRetire where 1=1" & stCon & " order by 1 DESC,2 DESC,3 DESC,4 DESC"
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
      txtBR(1).SetFocus
      txtBR_GotFocus 1
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
            If txtBR(1).Locked = False Then
               txtBR(1).Text = GRD1.TextMatrix(lCurRow, 0)
               txtBR(2).Text = GRD1.TextMatrix(lCurRow, 2)
               txtBR(3).Text = GRD1.TextMatrix(lCurRow, 4)
               txtBR(4).Text = GRD1.TextMatrix(lCurRow, 5)
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
   Select Case Index
      Case 2, 3
      Case 11
         KeyAscii = Pub_NumAscii(KeyAscii, True)
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtBR_Change(Index As Integer)
   If Index = 4 Then         '格式不同 label1名稱不同,輸入欄位不同
      'add by sonia 2018/6/26  107年起稅法修改為"取消可扣抵稅額",故取消"稅額扣抵比率"及"可扣抵稅額"欄位,退職時都顯示
      Label1(9).Visible = True
      Label1(10).Visible = True
      txtBR(9).Visible = True
      txtBR(10).Visible = True
      txtBR(10).Enabled = True
      'end 2018/6/26
      Select Case txtBR(Index)
         Case "54"
            Label1(8).Caption = "　　股利淨額："
            Label1(9).Caption = "可扣抵稅額："
            Label1(11).Caption = "股利所屬年度："
            txtBR(7).Locked = False
            txtBR(10).Locked = False
            txtBR(12).Locked = True
            txtBR(12).Visible = False
            txtBR(13).Locked = True
            txtBR(9).Locked = True        '2009/1/13 add by sonia 可扣抵稅額鎖住,由電腦計算
            'add by sonia 2018/6/26  107年起稅法修改為"取消可扣抵稅額",故取消"稅額扣抵比率"及"可扣抵稅額"欄位
            If Val(txtBR(1)) >= 107 Then
               Label1(9).Visible = False
               Label1(10).Visible = False
               txtBR(9).Visible = False
               txtBR(10).Visible = False
               txtBR(10).Locked = True
            End If
            'end 2018/6/26
         Case "93"
            Label1(8).Caption = "退職給付總額："
            Label1(9).Caption = "　扣繳稅額："
            Label1(11).Caption = "　　服務年資：              年          月"
            txtBR(7).Locked = True
            txtBR(10).Locked = True
            txtBR(10).Enabled = False
            txtBR(12).Locked = False
            txtBR(12).Visible = True
            txtBR(13).Locked = False
            txtBR(9).Locked = False       '2009/1/13 add by sonia
      End Select
   End If
End Sub

Private Sub txtBR_GotFocus(Index As Integer)
   TextInverse txtBR(Index)
   CloseIme
End Sub

Private Sub ClearField()
   For Each oText In txtBR
      oText.Text = Empty
   Next
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_BR
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
      For Each oText In txtBR
         idx = oText.Index
         Select Case idx
         '所得年度轉民國年
         Case 1
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName) - 1911
         '西元轉民國
         Case 7, 20
            m_FieldList(idx).fiOldData = TransDate("" & .Fields(m_FieldList(idx).fiName), 1)
         Case Else
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         End Select
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         oText.Text = m_FieldList(idx).fiOldData
      Next
      
      m_oi02 = "" 'add by sonia 2016/1/20
      If ClsPDGetOtherIncomer(txtBR(2), strExc(1), m_oi02) = True Then
         lblDsp(1) = strExc(1)
      ElseIf ClsPDGetStaffN(txtBR(2), strExc(1), , True) Then
         lblDsp(1) = strExc(1)
         m_oi02 = "0123456789" 'add by sonia 2017/7/27 要加此,否則後面判斷個人才寫補充保費會錯誤(員工檔一定是個人)
      End If
      lblDsp(2) = CompNameQuery(txtBR(3))
      
      CUID(1) = "" & .Fields("br14")
      CUID(2) = "" & .Fields("br15")
      CUID(3) = "" & .Fields("br16")
      CUID(4) = "" & .Fields("br17")
      CUID(5) = "" & .Fields("br18")
      CUID(6) = "" & .Fields("br19")
   End If
   End With
   UpdateCUID CUID, textCUID
   txtBR(1).Tag = txtBR(1)
   txtBR(2).Tag = txtBR(2)
   txtBR(3).Tag = txtBR(3)
   txtBR(4).Tag = txtBR(4)
   SetNet 'Added by Morgan 2013/3/13
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtBR
      oText.Locked = bLocked
   Next
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
            txtBR(1) = txtBR(1).Tag
            txtBR(2) = txtBR(2).Tag
            txtBR(3) = txtBR(3).Tag
            txtBR(4) = txtBR(4).Tag
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
         If m_bUpdate And txtBR(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtBR(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtBR(1) <> "" Then
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
         txtBR(1).Locked = False
         If Me.Visible = True Then
            txtBR(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 2
         txtBR(1).Locked = True
         txtBR(2).Locked = True
         txtBR(3).Locked = True
         txtBR(4).Locked = True
         
         If Me.Visible = True Then
            txtBR(3).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
         '2009/1/13 add by sonia 可扣抵稅額鎖住,由電腦計算
         If txtBR(4) = "54" Then
            txtBR(9).Locked = True
            txtBR(20).Locked = True 'Added by Morgan 2018/1/24 股利的給付日期也不可改(補充保費資料連結欄位)
         Else
            txtBR(9).Locked = False
         End If
         '2009/1/13 end
      Case 4
         txtBR(1).Locked = False
         txtBR(2).Locked = False
         txtBR(3).Locked = False
         txtBR(4).Locked = False
         If Me.Visible = True Then
            txtBR(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case Else
         txtBR(1).Locked = True
         If Me.Visible = True Then
            txtBR(1).SetFocus
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
               txtBR(1).SetFocus
               txtBR_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
Dim adoRst As New ADODB.Recordset   '2010/5/26 add by sonia

   m_bConfirmCheck = True
   
   For Each oText In txtBR
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtBR_Validate idx, bCancel
         If bCancel = True Then
            txtBR(idx).SetFocus
            txtBR_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   '查詢
   If m_EditMode = 4 Then
      If txtBR(1) = "" Then
         ShowMsg "請輸入所得年度 !"
         txtBR(1).SetFocus
         txtBR_GotFocus 1
         GoTo EscPoint
      End If
      If txtBR(2) = "" Then
         ShowMsg "請輸入所得人代號 !"
         txtBR(2).SetFocus
         txtBR_GotFocus 2
         GoTo EscPoint
      End If
      If txtBR(3) = "" Then
         ShowMsg "請輸入公司別 !"
         txtBR(3).SetFocus
         txtBR_GotFocus 3
         GoTo EscPoint
      End If
      If txtBR(4) = "" Then
         ShowMsg "請輸入格式 !"
         txtBR(4).SetFocus
         txtBR_GotFocus 4
         GoTo EscPoint
      End If
      
   '維護
   Else
      If txtBR(1) = "" And txtBR(1).Locked = False Then
         ShowMsg "請輸入所得年度 !"
         txtBR(1).SetFocus
         txtBR_GotFocus 1
         GoTo EscPoint
      End If
      If txtBR(2) = "" And txtBR(2).Locked = False Then
         ShowMsg "請輸入所得人代號 !"
         txtBR(2).SetFocus
         txtBR_GotFocus 2
         GoTo EscPoint
      End If
      If txtBR(3) = "" And txtBR(3).Locked = False Then
         ShowMsg "請輸入公司別 !"
         txtBR(3).SetFocus
         txtBR_GotFocus 3
         GoTo EscPoint
      End If
      If txtBR(4) = "" And txtBR(4).Locked = False Then
         ShowMsg "請輸入格式代號 !"
         txtBR(4).SetFocus
         txtBR_GotFocus 4
         GoTo EscPoint
      End If
      If txtBR(5) = "" And txtBR(5).Locked = False Then
         ShowMsg "請輸入起始月份 !"
         txtBR(5).SetFocus
         txtBR_GotFocus 5
         GoTo EscPoint
      End If
      If txtBR(6) = "" And txtBR(6).Locked = False Then
         ShowMsg "請輸入截止月份 !"
         txtBR(6).SetFocus
         txtBR_GotFocus 6
         GoTo EscPoint
      End If
      'Added by Morgan 2013/12/30
      If txtBR(4) = "54" Then
         If txtBR(7) = "" And txtBR(7).Locked = False Then
            MsgBox "請輸入除權基準日！"
            txtBR(7).SetFocus
            txtBR_GotFocus 7
            GoTo EscPoint
         End If
      End If
      'end 2013/12/30
      If txtBR(8) = "" And txtBR(8).Locked = False Then
         ShowMsg "請輸入所得總額 !"
         txtBR(8).SetFocus
         txtBR_GotFocus 8
         GoTo EscPoint
      End If
      If txtBR(4) = "54" And txtBR(10) = "" And txtBR(10).Locked = False Then
         If txtBR(10).Visible = True Then 'Added by Morgan 2018/7/23
            ShowMsg "請輸入稅額扣抵比率 !"
            txtBR(10).SetFocus
            txtBR_GotFocus 10
            GoTo EscPoint
         End If
      End If
      If txtBR(4) = "54" And txtBR(11) = "" And txtBR(11).Locked = False Then
         ShowMsg "請輸入股利所屬年度 !"
         txtBR(11).SetFocus
         txtBR_GotFocus 11
         GoTo EscPoint
      End If
      If txtBR(4) = "93" And txtBR(11) = "" And txtBR(12) = "" And txtBR(11).Locked = False Then
         ShowMsg "請輸入服務年資 !"
         txtBR(11).SetFocus
         txtBR_GotFocus 11
         GoTo EscPoint
      End If
      '2010/5/26 add by sonia 同一家公司同一年度,股利的稅額扣抵比率必須相同
      If (m_EditMode = 1 Or m_EditMode = 2) And txtBR(4) = "54" Then
         strExc(0) = "SELECT * FROM BonusRetire" & _
            " WHERE br01 = " & Val(txtBR(1)) + 1911 & " and br03= '" & txtBR(3) & "' and br04= '" & txtBR(4) & "' and br02<> '" & txtBR(2) & "' and br10<>'" & txtBR(10) & "' "
         intI = 1
         adoRst.MaxRecords = 1
         Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            ShowMsg "稅額扣抵比率與其他筆不符,不可輸入 !"
            txtBR(10).SetFocus
            txtBR_GotFocus 10
            GoTo EscPoint
         End If
         
         'Added by Morgan 2013/3/13
         If Val(txtBR(1)) >= 102 Then
            If txtBR(20) = "" Then
               ShowMsg "請輸入給付日期 !"
               txtBR(20).SetFocus
               GoTo EscPoint
            ElseIf txtBR(1) <> txtBR(20) \ 10000 Then
               ShowMsg "給付日期與所得年度不符 !"
               txtBR(20).SetFocus
               GoTo EscPoint
            Else
               SetNHI06
               'Modified by Morgan 2013/12/30 改提醒可繼續(補資料)
               If ChkHi2ndIsPaid(Left(DBDATE(txtBR(20)), 6), , txtBR(3)) = True Then
                  If MsgBox(txtBR(3) & "公司" & Val(txtBR(20) \ 100) & "月份的補充保費已繳納，是否確定要異動！", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                     GoTo EscPoint
                  End If
               End If
            End If
         End If
         'end 2013/3/13
         
      End If
      '2010/5/26 end
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
   For Each oText In txtBR
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
   stSQL = "INSERT INTO BonusRetire (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
'   stSQL = "select max(br02) from BonusRetire where br01='" & txtBR(1) & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
'   If intI = 1 Then
'      txtBR(2) = RsTemp.Fields(0)
'   End If
   
   'Added by Morgan 2013/3/13
   If txtBR(4) = "54" Then
      'Modified by Morgan 2017/1/23 個人才要新增補充保費記錄
      'PUB_InsertNHI2nd stNHI
      PUB_InsertNHI2nd stNHI, IIf(Len(m_oi02) = 10, True, False)
   End If
   'end 2013/3/13
   
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
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE BonusRetire SET "
   stSet = ""
   For Each oText In txtBR
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
      stSQL = stSQL & stSet & " where br01=" & Val(txtBR(1)) + 1911 & " and br02='" & txtBR(2) & "' and br03='" & txtBR(3) & "' and br04='" & txtBR(4) & "'; end; "
      
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
      
      'Added by Morgan 2013/3/13
      If txtBR(4) = "54" Then
         'Modified by Morgan 2017/1/23 個人才要新增補充保費記錄
         'PUB_InsertNHI2nd stNHI
         PUB_InsertNHI2nd stNHI, IIf(Len(m_oi02) = 10, True, False)
      ElseIf m_FieldList(4).fiOldData = "54" Then
         strSql = "DELETE NHI2ND WHERE NHI01='" & m_FieldList(2).fiOldData & "' AND NHI02=" & DBDATE(m_FieldList(20).fiOldData) & " AND NHI03='54' AND NHI04='0' and NHI11='" & m_FieldList(3).fiOldData & "'"
         cnnConnection.Execute strSql, intI
      End If
      'end 2013/3/13
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
   For Each oText In txtBR
      idx = oText.Index
      Select Case idx
         Case 1
            m_FieldList(idx).fiNewData = Val(oText.Text) + 1911
         Case 7, 20
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub

Private Sub txtBR_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 2, 3
      Case 10
         KeyAscii = Pub_NumAscii(KeyAscii, True)
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtBR_Validate(Index As Integer, Cancel As Boolean)
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Select Case Index
         Case 2
            m_oi02 = "" 'add by sonia 2016/1/20
            If txtBR(Index) <> "" Then
               'modify by sonia 2016/1/20 +m_oi02傳出身份證字號或統一編號
               If ClsPDGetOtherIncomer(txtBR(Index), strExc(1), m_oi02) = True Then
                  lblDsp(1) = strExc(1)
               Else
                  'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
                  'If ChkStaffID(Replace(txtBR(Index), "A", "0")) = True Then
                  If ChkStaffID(Left(txtBR(Index), 1) & Replace(Mid(txtBR(Index), 2), "A", "0")) = True Then
                     m_oi02 = "0123456789" 'add by sonia 2017/7/27 要加此,否則後面判斷個人才寫補充保費會錯誤(員工檔一定是個人)
                     Cancel = True
                  End If
                  If Cancel = False Then
                     If ClsPDGetStaffN(txtBR(Index), strExc(1), , True) = False Then
                        Cancel = True
                        lblDsp(1) = ""
                     Else
                        lblDsp(1) = strExc(1)
                        m_oi02 = "0123456789" 'Added by Morgan 2018/7/23
                     End If
                  End If
               End If
            End If
         Case 3
            If txtBR(Index) <> "" Then
               'Added by Morgan 2013/2/22
               If m_bolLimited And txtBR(Index) <> "A" Then
                  MsgBox "只可輸入A公司！", vbExclamation
                  Cancel = True
               Else
               'end 2013/2/22
                  lblDsp(2) = CompNameQuery(txtBR(Index))
                  If lblDsp(2) = "" Then
                     ShowMsg "公司別錯誤 !"
                     Cancel = True
                  End If
               End If 'Added by Morgan 2013/2/22
            End If
         Case 4
            If txtBR(Index) <> "" Then
               If InStr("54,93", txtBR(Index)) > 0 Then
               Else
                  ShowMsg "格式代號錯誤 !"
                  Cancel = True
               End If
            End If
         Case 5, 6
            If txtBR(Index) <> "" Then
               If Val(txtBR(Index)) < 1 Or Val(txtBR(Index)) > 12 Then
                  ShowMsg "月份不可小於 1 或超過 12 !"
                  Cancel = True
               End If
            End If
         Case 7
            If txtBR(Index) <> "" Then
               If ChkDate(txtBR(Index)) = False Then
                  Cancel = True
               End If
            End If
         '2009/1/13 add by sonia
         Case 10
            If txtBR(Index) <> "" Then
               'modify by sonia 2016/1/19 104年前起股利可扣抵稅額只能抵50%
               'txtBR(9) = Round(txtBR(8) * txtBR(Index) / 100, 0)
               'modify by sonia 2016/1/20 再加個人(Len(m_oi02) = 10)只能抵50%,公司(Len(m_oi02) = 8)則可全抵
               'If txtBR(4) = "54" And Val(txtBR(1)) >= 104 Then
               'modify by sonia 2018/1/24 應以股利所屬年度判斷Val(txtBR(1)) >= 104->Val(txtBR(11)) >= 104
               'If txtBR(4) = "54" And Val(txtBR(1)) >= 104 And Len(m_oi02) <> 8 Then
               If txtBR(4) = "54" And Val(txtBR(11)) >= 104 And Len(m_oi02) <> 8 Then
                  txtBR(9) = Round((txtBR(8) * txtBR(Index) / 100) / 2, 0)
               ElseIf txtBR(4) = "54" Then
                  txtBR(9) = Round(txtBR(8) * txtBR(Index) / 100, 0)
               End If
               SetNet
               SetNHI06
               'end 2016/1/19
            End If
         '2009/1/13 end
         
         'Added by Morgan 2013/3/13
         Case 20
            If txtBR(Index) <> "" Then
               If ChkDate(txtBR(Index)) = False Then
                  Cancel = True
               ElseIf Val(txtBR(Index)) > Val(strSrvDate(2)) Then
                  MsgBox "給付日期不可晚於系統日！", vbExclamation
                  Cancel = True
               End If
            End If
      End Select
      
      If Cancel = True Then TextInverse txtBR(Index)
      
      '若是按確定的檢查時略過, 檢查代號檔
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
         'Added by Moran 2013/3/13
         Case 2, 3, 4, 8, 9, 20
            SetNHI06
         'end 2013/3/13
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
   stSQL = "delete from BonusRetire where br01=" & Val(txtBR(1)) + 1911 & " and br02='" & txtBR(2) & "' and br03='" & txtBR(3) & "' and br04='" & txtBR(4) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   If m_FieldList(4).fiOldData = "54" Then
      strSql = "DELETE NHI2ND WHERE NHI01='" & m_FieldList(2).fiOldData & "' AND NHI02=" & DBDATE(m_FieldList(20).fiOldData) & " AND NHI03='54' AND NHI04='0' and NHI11='" & m_FieldList(3).fiOldData & "'"
      cnnConnection.Execute strSql, intI
   End If
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtBR(1).Tag = ""
   txtBR(2).Tag = ""
   txtBR(3).Tag = ""
   txtBR(4).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, iCol As Integer
   
   '格式顯示中文,代號隱藏
   arrGridHeadText = Array("年度", "月份", "所得人代號", "名　稱", "公司別", "", "格式", "所得總額", "扣繳稅額")
   arrGridHeadWidth = Array(450, 600, 1200, 1800, 600, 0, 600, 1200, 1200)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
      
      'Added by Morgan 2013/10/17
      If iRow < 7 Then
         GRD1.ColAlignment(iRow) = flexAlignLeftCenter
      Else
         GRD1.ColAlignment(iRow) = flexAlignRightCenter
      End If
   Next
   GRD1.Visible = True
End Sub
'Added by Morgan 2013/3/13
'計算補充保費
Private Sub SetNHI06()
   
   If txtBR(4) <> "54" Or txtBR(8) = "" Then
      txtBR(21) = ""
   ElseIf txtBR(20) <> "" Then
      stNHI(1) = txtBR(2)
      stNHI(2) = DBDATE(txtBR(20))
      stNHI(3) = txtBR(4)
      stNHI(4) = "0"
      stNHI(5) = ""
      stNHI(6) = ""
      'modify by sonia 2016/1/19
      'stNHI(7) = Val(txtBR(8)) + Val(txtBR(9))
      stNHI(7) = Val(txtTot)
      'end 2016/1/19
      stNHI(8) = ""
      stNHI(10) = "0"
      stNHI(11) = txtBR(3)
      PUB_NHI2nd stNHI(1), stNHI(2), stNHI(3), stNHI(4), stNHI(7), stNHI(5), stNHI(6), stNHI(8), , stNHI(11) 'Modified by Morgan 2014/5/1 +NHI11
      txtBR(21) = Val(stNHI(6))
   Else
      txtBR(21) = ""
   End If
   SetNet
End Sub
'計算給付淨額  '2016/1/19 +總額
Private Sub SetNet()
   'add by sonia 2016/1/19
   If txtBR(4) = "54" Then
      'modify by sonia 2016/1/20 再加個人(Len(m_oi02) = 10)只能抵50%,公司(Len(m_oi02) = 8)則可全抵
      'If Val(txtBR(1)) < 104 Then
      'Modified by Morgan 2016/6/24 給付總額 = 股利淨額 + 可扣抵稅額
      'If Val(txtBR(1)) < 104 Or Len(m_oi02) = 8 Then
      '   txtTot = Val(txtBR(8)) + Val(txtBR(9))
      'Else
      '   txtTot = Val(txtBR(8)) + Round(Val(txtBR(8)) * Val(txtBR(10)) / 100, 0)   '104年前起股利可扣抵稅額只能抵50%,但給付總額不變
      'End If
      txtTot = Val(txtBR(8)) + Val(txtBR(9))
      'end 2016/6/24
   Else
      txtTot = Val(txtBR(8))
   End If
   'end 2016/1/19
   txtNet = Val(txtBR(8)) - Val(txtBR(21))
End Sub
