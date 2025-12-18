VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020505 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人目標點數資料"
   ClientHeight    =   5760
   ClientLeft      =   156
   ClientTop       =   996
   ClientWidth     =   9132
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9132
   Begin TabDlg.SSTab tabCtrl 
      Height          =   4932
      Left            =   120
      TabIndex        =   50
      Top             =   720
      Width           =   8952
      _ExtentX        =   15790
      _ExtentY        =   8700
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "單筆"
      TabPicture(0)   =   "frm020505.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(13)=   "Label14"
      Tab(0).Control(14)=   "Label15"
      Tab(0).Control(15)=   "Label18"
      Tab(0).Control(16)=   "textST03"
      Tab(0).Control(17)=   "textPE01"
      Tab(0).Control(18)=   "textPE03_1"
      Tab(0).Control(19)=   "textPE03_2"
      Tab(0).Control(20)=   "textPE06"
      Tab(0).Control(21)=   "textPE08"
      Tab(0).Control(22)=   "textA0902"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textPE12"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textPE14"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textPE15"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textPE17"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textPE18"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textPE19"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textPE20(0)"
      Tab(0).Control(30)=   "textPE23(0)"
      Tab(0).Control(31)=   "textPE24(0)"
      Tab(0).Control(32)=   "textPE26(0)"
      Tab(0).Control(33)=   "textPE27(0)"
      Tab(0).Control(34)=   "textPE28(0)"
      Tab(0).Control(35)=   "textPE29(0)"
      Tab(0).Control(36)=   "textPE13"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textPE16"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textPE02"
      Tab(0).Control(39)=   "textPE20(1)"
      Tab(0).Control(40)=   "textPE20(2)"
      Tab(0).Control(41)=   "textPE20(3)"
      Tab(0).Control(42)=   "textPE25(0)"
      Tab(0).Control(43)=   "textPE21(0)"
      Tab(0).Control(44)=   "textPE21(1)"
      Tab(0).Control(45)=   "textPE21(2)"
      Tab(0).Control(46)=   "textPE21(3)"
      Tab(0).Control(47)=   "textPE22(0)"
      Tab(0).Control(48)=   "textPE22(1)"
      Tab(0).Control(49)=   "textPE22(2)"
      Tab(0).Control(50)=   "textPE22(3)"
      Tab(0).Control(51)=   "textPE23(1)"
      Tab(0).Control(52)=   "textPE23(2)"
      Tab(0).Control(53)=   "textPE23(3)"
      Tab(0).Control(54)=   "textPE24(1)"
      Tab(0).Control(55)=   "textPE24(2)"
      Tab(0).Control(56)=   "textPE24(3)"
      Tab(0).Control(57)=   "textPE25(2)"
      Tab(0).Control(58)=   "textPE25(3)"
      Tab(0).Control(59)=   "textPE26(1)"
      Tab(0).Control(60)=   "textPE26(2)"
      Tab(0).Control(61)=   "textPE26(3)"
      Tab(0).Control(62)=   "textPE27(1)"
      Tab(0).Control(63)=   "textPE27(2)"
      Tab(0).Control(64)=   "textPE27(3)"
      Tab(0).Control(65)=   "textPE28(1)"
      Tab(0).Control(66)=   "textPE28(2)"
      Tab(0).Control(67)=   "textPE28(3)"
      Tab(0).Control(68)=   "textPE29(1)"
      Tab(0).Control(69)=   "textPE29(2)"
      Tab(0).Control(70)=   "textPE29(3)"
      Tab(0).Control(71)=   "textPE25(1)"
      Tab(0).ControlCount=   72
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm020505.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label16"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label17"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "textPE03_Month"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "textPE03_Year"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdQuery"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "grdList"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   4092
         Left            =   120
         TabIndex        =   79
         Top             =   744
         Width           =   8712
         _ExtentX        =   15367
         _ExtentY        =   7218
         _Version        =   393216
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
      Begin VB.TextBox textPE25 
         Height          =   264
         Index           =   1
         Left            =   -69240
         MaxLength       =   6
         TabIndex        =   27
         Top             =   3600
         Width           =   732
      End
      Begin VB.TextBox textPE29 
         Height          =   264
         Index           =   3
         Left            =   -72120
         MaxLength       =   2
         TabIndex        =   45
         Top             =   4320
         Width           =   372
      End
      Begin VB.TextBox textPE29 
         Height          =   264
         Index           =   2
         Left            =   -72360
         MaxLength       =   1
         TabIndex        =   44
         Top             =   4320
         Width           =   252
      End
      Begin VB.TextBox textPE29 
         Height          =   264
         Index           =   1
         Left            =   -73080
         MaxLength       =   6
         TabIndex        =   43
         Top             =   4320
         Width           =   732
      End
      Begin VB.TextBox textPE28 
         Height          =   264
         Index           =   3
         Left            =   -68280
         MaxLength       =   2
         TabIndex        =   41
         Top             =   3960
         Width           =   372
      End
      Begin VB.TextBox textPE28 
         Height          =   264
         Index           =   2
         Left            =   -68520
         MaxLength       =   1
         TabIndex        =   40
         Top             =   3960
         Width           =   252
      End
      Begin VB.TextBox textPE28 
         Height          =   264
         Index           =   1
         Left            =   -69240
         MaxLength       =   6
         TabIndex        =   39
         Top             =   3960
         Width           =   732
      End
      Begin VB.TextBox textPE27 
         Height          =   264
         Index           =   3
         Left            =   -70200
         MaxLength       =   2
         TabIndex        =   37
         Top             =   3960
         Width           =   372
      End
      Begin VB.TextBox textPE27 
         Height          =   264
         Index           =   2
         Left            =   -70440
         MaxLength       =   1
         TabIndex        =   36
         Top             =   3960
         Width           =   252
      End
      Begin VB.TextBox textPE27 
         Height          =   264
         Index           =   1
         Left            =   -71160
         MaxLength       =   6
         TabIndex        =   35
         Top             =   3960
         Width           =   732
      End
      Begin VB.TextBox textPE26 
         Height          =   264
         Index           =   3
         Left            =   -72120
         MaxLength       =   2
         TabIndex        =   33
         Top             =   3960
         Width           =   372
      End
      Begin VB.TextBox textPE26 
         Height          =   264
         Index           =   2
         Left            =   -72360
         MaxLength       =   1
         TabIndex        =   32
         Top             =   3960
         Width           =   252
      End
      Begin VB.TextBox textPE26 
         Height          =   264
         Index           =   1
         Left            =   -73080
         MaxLength       =   6
         TabIndex        =   31
         Top             =   3960
         Width           =   732
      End
      Begin VB.TextBox textPE25 
         Height          =   264
         Index           =   3
         Left            =   -68280
         MaxLength       =   2
         TabIndex        =   29
         Top             =   3600
         Width           =   372
      End
      Begin VB.TextBox textPE25 
         Height          =   264
         Index           =   2
         Left            =   -68520
         MaxLength       =   1
         TabIndex        =   28
         Top             =   3600
         Width           =   252
      End
      Begin VB.TextBox textPE24 
         Height          =   264
         Index           =   3
         Left            =   -70200
         MaxLength       =   2
         TabIndex        =   25
         Top             =   3600
         Width           =   372
      End
      Begin VB.TextBox textPE24 
         Height          =   264
         Index           =   2
         Left            =   -70440
         MaxLength       =   1
         TabIndex        =   24
         Top             =   3600
         Width           =   252
      End
      Begin VB.TextBox textPE24 
         Height          =   264
         Index           =   1
         Left            =   -71160
         MaxLength       =   6
         TabIndex        =   23
         Top             =   3600
         Width           =   732
      End
      Begin VB.TextBox textPE23 
         Height          =   264
         Index           =   3
         Left            =   -72120
         MaxLength       =   2
         TabIndex        =   21
         Top             =   3600
         Width           =   372
      End
      Begin VB.TextBox textPE23 
         Height          =   264
         Index           =   2
         Left            =   -72360
         MaxLength       =   1
         TabIndex        =   20
         Top             =   3600
         Width           =   252
      End
      Begin VB.TextBox textPE23 
         Height          =   264
         Index           =   1
         Left            =   -73080
         MaxLength       =   6
         TabIndex        =   19
         Top             =   3600
         Width           =   732
      End
      Begin VB.TextBox textPE22 
         Height          =   264
         Index           =   3
         Left            =   -68280
         MaxLength       =   2
         TabIndex        =   17
         Top             =   3240
         Width           =   372
      End
      Begin VB.TextBox textPE22 
         Height          =   264
         Index           =   2
         Left            =   -68520
         MaxLength       =   1
         TabIndex        =   16
         Top             =   3240
         Width           =   252
      End
      Begin VB.TextBox textPE22 
         Height          =   264
         Index           =   1
         Left            =   -69240
         MaxLength       =   6
         TabIndex        =   15
         Top             =   3240
         Width           =   732
      End
      Begin VB.TextBox textPE22 
         Height          =   264
         Index           =   0
         Left            =   -69720
         MaxLength       =   3
         TabIndex        =   14
         Top             =   3240
         Width           =   492
      End
      Begin VB.TextBox textPE21 
         Height          =   264
         Index           =   3
         Left            =   -70200
         MaxLength       =   2
         TabIndex        =   13
         Top             =   3240
         Width           =   372
      End
      Begin VB.TextBox textPE21 
         Height          =   264
         Index           =   2
         Left            =   -70440
         MaxLength       =   1
         TabIndex        =   12
         Top             =   3240
         Width           =   252
      End
      Begin VB.TextBox textPE21 
         Height          =   264
         Index           =   1
         Left            =   -71160
         MaxLength       =   6
         TabIndex        =   11
         Top             =   3240
         Width           =   732
      End
      Begin VB.TextBox textPE21 
         Height          =   264
         Index           =   0
         Left            =   -71640
         MaxLength       =   3
         TabIndex        =   10
         Top             =   3240
         Width           =   492
      End
      Begin VB.TextBox textPE25 
         Height          =   264
         Index           =   0
         Left            =   -69720
         MaxLength       =   3
         TabIndex        =   26
         Top             =   3600
         Width           =   492
      End
      Begin VB.TextBox textPE20 
         Height          =   264
         Index           =   3
         Left            =   -72120
         MaxLength       =   2
         TabIndex        =   9
         Top             =   3240
         Width           =   372
      End
      Begin VB.TextBox textPE20 
         Height          =   264
         Index           =   2
         Left            =   -72360
         MaxLength       =   1
         TabIndex        =   8
         Top             =   3240
         Width           =   252
      End
      Begin VB.TextBox textPE20 
         Height          =   264
         Index           =   1
         Left            =   -73080
         MaxLength       =   6
         TabIndex        =   7
         Top             =   3240
         Width           =   732
      End
      Begin VB.TextBox textPE02 
         Height          =   264
         Left            =   -69960
         MaxLength       =   8
         TabIndex        =   3
         Top             =   720
         Width           =   612
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   400
         Left            =   7740
         TabIndex        =   49
         Top             =   300
         Width           =   1092
      End
      Begin VB.TextBox textPE03_Year 
         Height          =   264
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   47
         Top             =   360
         Width           =   612
      End
      Begin VB.TextBox textPE03_Month 
         Height          =   264
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   48
         Top             =   360
         Width           =   492
      End
      Begin VB.TextBox textPE16 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -69960
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1212
      End
      Begin VB.TextBox textPE13 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -69960
         Locked          =   -1  'True
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1212
      End
      Begin VB.TextBox textPE29 
         Height          =   264
         Index           =   0
         Left            =   -73560
         MaxLength       =   3
         TabIndex        =   42
         Top             =   4320
         Width           =   492
      End
      Begin VB.TextBox textPE28 
         Height          =   264
         Index           =   0
         Left            =   -69720
         MaxLength       =   3
         TabIndex        =   38
         Top             =   3960
         Width           =   492
      End
      Begin VB.TextBox textPE27 
         Height          =   264
         Index           =   0
         Left            =   -71640
         MaxLength       =   3
         TabIndex        =   34
         Top             =   3960
         Width           =   492
      End
      Begin VB.TextBox textPE26 
         Height          =   264
         Index           =   0
         Left            =   -73560
         MaxLength       =   3
         TabIndex        =   30
         Top             =   3960
         Width           =   492
      End
      Begin VB.TextBox textPE24 
         Height          =   264
         Index           =   0
         Left            =   -71640
         MaxLength       =   3
         TabIndex        =   22
         Top             =   3600
         Width           =   492
      End
      Begin VB.TextBox textPE23 
         Height          =   264
         Index           =   0
         Left            =   -73560
         MaxLength       =   3
         TabIndex        =   18
         Top             =   3600
         Width           =   492
      End
      Begin VB.TextBox textPE20 
         Height          =   264
         Index           =   0
         Left            =   -73560
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3240
         Width           =   492
      End
      Begin VB.TextBox textPE19 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -69960
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1212
      End
      Begin VB.TextBox textPE18 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1212
      End
      Begin VB.TextBox textPE17 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1212
      End
      Begin VB.TextBox textPE15 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1212
      End
      Begin VB.TextBox textPE14 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1212
      End
      Begin VB.TextBox textPE12 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1212
      End
      Begin VB.TextBox textA0902 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -69960
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   360
         Width           =   1212
      End
      Begin VB.TextBox textPE08 
         Height          =   264
         Left            =   -69960
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1080
         Width           =   1212
      End
      Begin VB.TextBox textPE06 
         Height          =   264
         Left            =   -73560
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1080
         Width           =   1212
      End
      Begin VB.TextBox textPE03_2 
         Height          =   264
         Left            =   -72480
         MaxLength       =   2
         TabIndex        =   2
         Top             =   720
         Width           =   492
      End
      Begin VB.TextBox textPE03_1 
         Height          =   264
         Left            =   -73560
         MaxLength       =   3
         TabIndex        =   1
         Top             =   720
         Width           =   612
      End
      Begin VB.TextBox textPE01 
         Height          =   264
         Left            =   -73560
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   852
      End
      Begin MSForms.TextBox textST03 
         Height          =   264
         Left            =   -72600
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   360
         Width           =   1332
         VariousPropertyBits=   679493663
         Size            =   "2350;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label18 
         Caption         =   "系統類別 :"
         Height          =   252
         Left            =   -71040
         TabIndex        =   46
         Top             =   720
         Width           =   972
      End
      Begin VB.Label Label17 
         Caption         =   "資料年月 :"
         Height          =   252
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label16 
         Caption         =   "/"
         Height          =   252
         Left            =   2280
         TabIndex        =   77
         Top             =   360
         Width           =   252
      End
      Begin VB.Label Label15 
         Caption         =   "/"
         Height          =   252
         Left            =   -72720
         TabIndex        =   66
         Top             =   720
         Width           =   252
      End
      Begin VB.Label Label14 
         Caption         =   "勝訴率 2 :"
         Height          =   252
         Left            =   -71040
         TabIndex        =   65
         Top             =   2880
         Width           =   852
      End
      Begin VB.Label Label13 
         Caption         =   "未輸入筆數 :"
         Height          =   252
         Left            =   -71040
         TabIndex        =   64
         Top             =   2160
         Width           =   1212
      End
      Begin VB.Label Label12 
         Caption         =   "英文筆數 :"
         Height          =   252
         Left            =   -71040
         TabIndex        =   63
         Top             =   1440
         Width           =   972
      End
      Begin VB.Label Label11 
         Caption         =   "其它發文點數 :"
         Height          =   255
         Left            =   -71400
         TabIndex        =   62
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "員工代號 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   61
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label9 
         Caption         =   "查名失誤案號 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   60
         Top             =   3240
         Width           =   1212
      End
      Begin VB.Label Label8 
         Caption         =   "勝訴率 1 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   59
         Top             =   2880
         Width           =   852
      End
      Begin VB.Label Label7 
         Caption         =   "預估準確率 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   58
         Top             =   2520
         Width           =   1212
      End
      Begin VB.Label Label6 
         Caption         =   "過期筆數 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   57
         Top             =   2160
         Width           =   972
      End
      Begin VB.Label Label5 
         Caption         =   "圖形筆數 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   56
         Top             =   1800
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "中文筆數 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   55
         Top             =   1440
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "目標點數 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   54
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "工作性質 :"
         Height          =   252
         Left            =   -71040
         TabIndex        =   53
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "資料年月 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   52
         Top             =   720
         Width           =   972
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   8520
      Top             =   660
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
            Picture         =   "frm020505.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020505.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020505.frx":0670
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020505.frx":084C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020505.frx":0B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020505.frx":0E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020505.frx":11A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020505.frx":14BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020505.frx":17D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020505.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020505.frx":1E10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgList"
      DisabledImageList=   "ImgList"
      HotImageList    =   "ImgList"
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
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
End
Attribute VB_Name = "frm020505"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2022/01/05 Form2.0已修改 textST03/grdList
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

Const MAX_FIELD = 29

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList(MAX_FIELD) As FIELDITEM

' 變數宣告區
Dim m_EditMode As Integer
' 第一筆資料的本所案號
Dim m_FirstPE(3) As String
' 最後一筆資料的本所案號
Dim m_LastPE(3) As String
' 目前正在顯示的本所案號
Dim m_CurrPE(3) As String
'
Dim m_CurrSel As Integer

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE) AND " & _
                  "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE)) AND " & _
                  "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE) AND PE02 = (SELECT MIN(PE02) FROM PERFORMANCE WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE)))"
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_FirstPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_FirstPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_FirstPE(2) = rsTmp.Fields("PE03")
   End If
   rsTmp.Close

   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE) AND " & _
                  "PE02 = (SELECT MAX(PE02) FROM PERFORMANCE WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE)) AND " & _
                  "PE03 = (SELECT MAX(PE03) FROM PERFORMANCE WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE) AND PE02 = (SELECT MAX(PE02) FROM PERFORMANCE WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE)))"
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_LastPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_LastPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_LastPE(2) = rsTmp.Fields("PE03")
   End If
   rsTmp.Close

   Set rsTmp = Nothing
End Sub

Private Sub cmdQuery_Click()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strPE01 As String
   Dim strPE02 As String
   Dim strPE03 As String
      
   strPE02 = "T"
   strPE03 = ConvertTYearToWYear(textPE03_Year) & textPE03_Month
   strSql = "SELECT * FROM Performance " & _
            "WHERE PE02 = '" & strPE02 & "' AND " & _
                  "PE03 = " & strPE03 & " "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   UpdateGridList rsTmp
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' Load Form
Private Sub Form_Load()
   tabCtrl.Tab = 1
   
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm020505", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm020505", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm020505", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm020505", strFind, False)
   
   m_EditMode = 0
   MoveFormToCenter Me
   
   textST03.BackColor = &H8000000F
   textA0902.BackColor = &H8000000F
   textPE12.BackColor = &H8000000F
   textPE13.BackColor = &H8000000F
   textPE14.BackColor = &H8000000F
   textPE15.BackColor = &H8000000F
   textPE16.BackColor = &H8000000F
   textPE17.BackColor = &H8000000F
   textPE18.BackColor = &H8000000F
   textPE19.BackColor = &H8000000F
   
   InitialField
   RefreshRange
   ShowFirstRecord
   SetCtrlReadOnly True
   UpdateToolbarState
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "PE" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, ByVal strData As String)
   Dim nIndex As Integer
   For nIndex = 0 To MAX_FIELD - 1
      If strName = m_FieldList(nIndex).fiName Then
         m_FieldList(nIndex).fiNewData = strData
         Exit For
      End If
   Next nIndex
End Sub

' 更新欄位的內容
Private Sub UpdateFieldNewData()
   SetFieldNewData "PE01", textPE01
   SetFieldNewData "PE02", textPE02
   SetFieldNewData "PE03", (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
   SetFieldNewData "PE06", textPE06
   SetFieldNewData "PE08", textPE08
    'Modify By Cheng 2004/03/11
    '取消Mark, 避免修改時, 將資料清除
   SetFieldNewData "PE12", textPE12
   SetFieldNewData "PE13", textPE13
   SetFieldNewData "PE14", textPE14
   SetFieldNewData "PE15", textPE15
   SetFieldNewData "PE16", textPE16
   SetFieldNewData "PE17", textPE17
   SetFieldNewData "PE18", textPE18
   SetFieldNewData "PE19", textPE19
    'End
   If IsEmptyText(textPE20(0)) = False Then
      SetFieldNewData "PE20", textPE20(0) & textPE20(1) & textPE20(2) & String(1 - Len(textPE20(2)), "0") & textPE20(3) & String(2 - Len(textPE20(3)), "0")
   Else
      SetFieldNewData "PE20", Empty
   End If
   If IsEmptyText(textPE21(0)) = False Then
      SetFieldNewData "PE21", textPE21(0) & textPE21(1) & textPE21(2) & String(1 - Len(textPE21(2)), "0") & textPE21(3) & String(2 - Len(textPE21(3)), "0")
   Else
      SetFieldNewData "PE21", Empty
   End If
   If IsEmptyText(textPE22(0)) = False Then
      SetFieldNewData "PE22", textPE22(0) & textPE22(1) & textPE22(2) & String(1 - Len(textPE22(2)), "0") & textPE22(3) & String(2 - Len(textPE22(3)), "0")
   Else
      SetFieldNewData "PE22", Empty
   End If
   If IsEmptyText(textPE23(0)) = False Then
      SetFieldNewData "PE23", textPE23(0) & textPE23(1) & textPE23(2) & String(1 - Len(textPE23(2)), "0") & textPE23(3) & String(2 - Len(textPE23(3)), "0")
   Else
      SetFieldNewData "PE23", Empty
   End If
   If IsEmptyText(textPE24(0)) = False Then
      SetFieldNewData "PE24", textPE24(0) & textPE24(1) & textPE24(2) & String(1 - Len(textPE24(2)), "0") & textPE24(3) & String(2 - Len(textPE24(3)), "0")
   Else
      SetFieldNewData "PE24", Empty
   End If
   If IsEmptyText(textPE25(0)) = False Then
      SetFieldNewData "PE25", textPE25(0) & textPE25(1) & textPE25(2) & String(1 - Len(textPE25(2)), "0") & textPE25(3) & String(2 - Len(textPE25(3)), "0")
   Else
      SetFieldNewData "PE25", Empty
   End If
   If IsEmptyText(textPE26(0)) = False Then
      SetFieldNewData "PE26", textPE26(0) & textPE26(1) & textPE26(2) & String(1 - Len(textPE26(2)), "0") & textPE26(3) & String(2 - Len(textPE26(3)), "0")
   Else
      SetFieldNewData "PE26", Empty
   End If
   If IsEmptyText(textPE27(0)) = False Then
      SetFieldNewData "PE27", textPE27(0) & textPE27(1) & textPE27(2) & String(1 - Len(textPE27(2)), "0") & textPE27(3) & String(2 - Len(textPE27(3)), "0")
   Else
      SetFieldNewData "PE27", Empty
   End If
   If IsEmptyText(textPE28(0)) = False Then
      SetFieldNewData "PE28", textPE28(0) & textPE28(1) & textPE28(2) & String(1 - Len(textPE28(2)), "0") & textPE28(3) & String(2 - Len(textPE28(3)), "0")
   Else
      SetFieldNewData "PE28", Empty
   End If
   If IsEmptyText(textPE29(0)) = False Then
      SetFieldNewData "PE29", textPE29(0) & textPE29(1) & textPE29(2) & String(1 - Len(textPE29(2)), "0") & textPE29(3) & String(2 - Len(textPE29(3)), "0")
   Else
      SetFieldNewData "PE29", Empty
   End If
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsSrcTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsSrcTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsSrcTmp.Fields(m_FieldList(nIndex).fiName)
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = rsSrcTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 讀取資料庫所有的資料
Private Sub QueryDB()
   'Dim strSQL As String
   'Dim rsTmp As New ADODB.Recordset
   
   'RefreshRange
   
   'strSQL = "SELECT * FROM Performance " & _
   '         "WHERE PE02 LIKE 'T%' " & _
   '         "ORDER BY PE01, PE02, PE03"
   'rsTmp.CursorLocation = adUseClient
   'rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
   'UpdateGridList rsTmp
   'rsTmp.Close

   'Set rsTmp = Nothing
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   textST03 = Empty
   textPE01 = Empty: textPE03_1 = Empty: textPE03_2 = Empty: textPE02 = Empty
   textPE06 = Empty: textPE08 = Empty: textA0902 = Empty
   textPE12 = Empty: textPE13 = Empty: textPE14 = Empty: textPE15 = Empty
   textPE16 = Empty: textPE17 = Empty: textPE18 = Empty: textPE19 = Empty
   textPE20(0) = Empty: textPE20(1) = Empty: textPE20(2) = Empty: textPE20(3) = Empty
   textPE21(0) = Empty: textPE21(1) = Empty: textPE21(2) = Empty: textPE21(3) = Empty
   textPE22(0) = Empty: textPE22(1) = Empty: textPE22(2) = Empty: textPE22(3) = Empty
   textPE23(0) = Empty: textPE23(1) = Empty: textPE23(2) = Empty: textPE23(3) = Empty
   textPE24(0) = Empty: textPE24(1) = Empty: textPE24(2) = Empty: textPE24(3) = Empty
   textPE25(0) = Empty: textPE25(1) = Empty: textPE25(2) = Empty: textPE25(3) = Empty
   textPE26(0) = Empty: textPE26(1) = Empty: textPE26(2) = Empty: textPE26(3) = Empty
   textPE27(0) = Empty: textPE27(1) = Empty: textPE27(2) = Empty: textPE27(3) = Empty
   textPE28(0) = Empty: textPE28(1) = Empty: textPE28(2) = Empty: textPE28(3) = Empty
   textPE29(0) = Empty: textPE29(1) = Empty: textPE29(2) = Empty: textPE29(3) = Empty

   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textPE01.Locked = bEnable: textPE02.Locked = bEnable: textPE03_1.Locked = bEnable: textPE03_2.Locked = bEnable
   textPE06.Locked = bEnable: textPE08.Locked = bEnable
   textPE12.Locked = True:    textPE13.Locked = True:    textPE14.Locked = True:    textPE15.Locked = True
   textPE16.Locked = True:    textPE17.Locked = True:    textPE18.Locked = True:    textPE19.Locked = True
   textPE20(0).Locked = bEnable: textPE20(1).Locked = bEnable: textPE20(2).Locked = bEnable: textPE20(3).Locked = bEnable
   textPE21(0).Locked = bEnable: textPE21(1).Locked = bEnable: textPE21(2).Locked = bEnable: textPE21(3).Locked = bEnable
   textPE22(0).Locked = bEnable: textPE22(1).Locked = bEnable: textPE22(2).Locked = bEnable: textPE22(3).Locked = bEnable
   textPE23(0).Locked = bEnable: textPE23(1).Locked = bEnable: textPE23(2).Locked = bEnable: textPE23(3).Locked = bEnable
   textPE24(0).Locked = bEnable: textPE24(1).Locked = bEnable: textPE24(2).Locked = bEnable: textPE24(3).Locked = bEnable
   textPE25(0).Locked = bEnable: textPE25(1).Locked = bEnable: textPE25(2).Locked = bEnable: textPE25(3).Locked = bEnable
   textPE26(0).Locked = bEnable: textPE26(1).Locked = bEnable: textPE26(2).Locked = bEnable: textPE26(3).Locked = bEnable
   textPE27(0).Locked = bEnable: textPE27(1).Locked = bEnable: textPE27(2).Locked = bEnable: textPE27(3).Locked = bEnable
   textPE28(0).Locked = bEnable: textPE28(1).Locked = bEnable: textPE28(2).Locked = bEnable: textPE28(3).Locked = bEnable
   textPE29(0).Locked = bEnable: textPE29(1).Locked = bEnable: textPE29(2).Locked = bEnable: textPE29(3).Locked = bEnable
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textPE01.Locked = bEnable: textPE02.Locked = bEnable: textPE03_1.Locked = bEnable: textPE03_2.Locked = bEnable
End Sub

Private Function ConvertAC(ByVal strData As String, ByRef strKey1 As String, ByRef StrKey2 As String, ByRef strKey3 As String, ByRef strKey4 As String) As Boolean
   ConvertAC = True
   Select Case Len(strData)
      Case 10:
         strKey1 = Mid(strData, 1, 1)
         StrKey2 = Mid(strData, 2, 6)
         strKey3 = Mid(strData, 8, 1)
         strKey4 = Mid(strData, 9, 2)
      Case 11:
         strKey1 = Mid(strData, 1, 2)
         StrKey2 = Mid(strData, 3, 6)
         strKey3 = Mid(strData, 9, 1)
         strKey4 = Mid(strData, 10, 2)
      Case 12:
         strKey1 = Mid(strData, 1, 3)
         StrKey2 = Mid(strData, 4, 6)
         strKey3 = Mid(strData, 10, 1)
         strKey4 = Mid(strData, 11, 2)
      Case Else:
         ConvertAC = False
         strKey1 = Empty
         StrKey2 = Empty
         strKey3 = Empty
         strKey4 = Empty
   End Select
End Function

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim bCnv As Boolean
   Dim strPENo(4) As String
   
   strSql = "SELECT * FROM Performance " & _
            "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                  "PE02 = '" & m_CurrPE(1) & "' AND " & _
                  "PE03 = '" & m_CurrPE(2) & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      textPE01 = rsTmp.Fields("PE01")
      textPE02 = rsTmp.Fields("PE02")
      If Not IsNull(rsTmp.Fields("PE03")) Then
         textPE03_1 = Val(Mid(rsTmp.Fields("PE03"), 1, 4)) - 1911
         textPE03_2 = Mid(rsTmp.Fields("PE03"), 5, 2)
      End If
      If Not IsNull(rsTmp.Fields("PE06")) Then: textPE06 = rsTmp.Fields("PE06"): 'End If
      If Not IsNull(rsTmp.Fields("PE08")) Then: textPE08 = rsTmp.Fields("PE08"): 'End If
      If Not IsNull(rsTmp.Fields("PE12")) Then: textPE12 = rsTmp.Fields("PE12"): 'End If
      If Not IsNull(rsTmp.Fields("PE13")) Then: textPE13 = rsTmp.Fields("PE13"): 'End If
      If Not IsNull(rsTmp.Fields("PE14")) Then: textPE14 = rsTmp.Fields("PE14"): 'End If
      If Not IsNull(rsTmp.Fields("PE15")) Then: textPE15 = rsTmp.Fields("PE15"): 'End If
      If Not IsNull(rsTmp.Fields("PE16")) Then: textPE16 = rsTmp.Fields("PE16"): 'End If
      If Not IsNull(rsTmp.Fields("PE17")) Then: textPE17 = rsTmp.Fields("PE17"): 'End If
      If Not IsNull(rsTmp.Fields("PE18")) Then: textPE18 = rsTmp.Fields("PE18"): 'End If
      If Not IsNull(rsTmp.Fields("PE19")) Then: textPE19 = rsTmp.Fields("PE19"): 'End If
   ' 本所案號
      strKey1 = Empty: StrKey2 = Empty: strKey3 = Empty: strKey4 = Empty
      If Not IsNull(rsTmp.Fields("PE20")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE20"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE20(0) = strKey1: textPE20(1) = StrKey2: textPE20(2) = strKey3: textPE20(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE21")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE21"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE21(0) = strKey1: textPE21(1) = StrKey2: textPE21(2) = strKey3: textPE21(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE22")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE22"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE22(0) = strKey1: textPE22(1) = StrKey2: textPE22(2) = strKey3: textPE22(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE23")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE23"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE23(0) = strKey1: textPE23(1) = StrKey2: textPE23(2) = strKey3: textPE23(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE24")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE24"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE24(0) = strKey1: textPE24(1) = StrKey2: textPE24(2) = strKey3: textPE24(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE25")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE25"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE25(0) = strKey1: textPE25(1) = StrKey2: textPE25(2) = strKey3: textPE25(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE26")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE26"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE26(0) = strKey1: textPE26(1) = StrKey2: textPE26(2) = strKey3: textPE26(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE27")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE27"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE27(0) = strKey1: textPE27(1) = StrKey2: textPE27(2) = strKey3: textPE27(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE28")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE28"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE28(0) = strKey1: textPE28(1) = StrKey2: textPE28(2) = strKey3: textPE28(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE29")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE29"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE29(0) = strKey1: textPE29(1) = StrKey2: textPE29(2) = strKey3: textPE29(3) = strKey4
      End If
      UpdateFieldOldData rsTmp
   
      ' 更新控制項中需帶出的資料
      textPE01_Validate False
   End If
EXITSUB:
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strPE01 As String, ByVal strPE02 As String, ByVal strPE03 As String)
   Dim strTemp As String
   Dim nIndex As Integer
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset

   If IsRecordExist(strPE01, strPE02, strPE03) = True Then
      m_CurrPE(0) = strPE01
      m_CurrPE(1) = strPE02
      m_CurrPE(2) = strPE03
   Else
      strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
               "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                     "PE02 = '" & m_CurrPE(1) & "' AND " & _
                     "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                             "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                   "PE02 = '" & m_CurrPE(1) & "' AND " & _
                                   "PE03 > '" & m_CurrPE(2) & "' ) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
         If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
         If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
   
      strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
               "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                     "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                             "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                   "PE02 > '" & m_CurrPE(1) & "') AND " & _
                     "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                             "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                   "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                                           "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                                 "PE02 > '" & m_CurrPE(1) & "'))"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
         If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
         If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
               "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                             "WHERE PE01 > '" & m_CurrPE(0) & "') AND " & _
                     "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                             "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                           "WHERE PE01 > '" & m_CurrPE(0) & "')) AND " & _
                     "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                             "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                           "WHERE PE01 > '" & m_CurrPE(0) & "') AND " & _
                                   "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                                           "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                                         "WHERE PE01 > '" & m_CurrPE(0) & "')))"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
         If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
         If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrPE(0) = m_FirstPE(0)
   m_CurrPE(1) = m_FirstPE(1)
   m_CurrPE(2) = m_FirstPE(2)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrPE(0) = m_FirstPE(0) And m_CurrPE(1) = m_FirstPE(1) And m_CurrPE(2) = m_FirstPE(2) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If

   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                  "PE02 = '" & m_CurrPE(1) & "' AND " & _
                  "PE03 = (SELECT MAX(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 = '" & m_CurrPE(1) & "' AND " & _
                                "PE03 < '" & m_CurrPE(2) & "')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                  "PE02 = (SELECT MAX(PE02) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 < '" & m_CurrPE(1) & "') AND " & _
                  "PE03 = (SELECT MAX(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 = (SELECT MAX(PE02) FROM PERFORMANCE " & _
                                        "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                              "PE02 < '" & m_CurrPE(1) & "'))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE " & _
                          "WHERE PE01 < '" & m_CurrPE(0) & "') AND " & _
                  "PE02 = (SELECT MAX(PE02) FROM PERFORMANCE " & _
                          "WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE " & _
                                        "WHERE PE01 < '" & m_CurrPE(0) & "')) AND " & _
                  "PE03 = (SELECT MAX(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE " & _
                                        "WHERE PE01 < '" & m_CurrPE(0) & "') AND " & _
                                "PE02 = (SELECT MAX(PE02) FROM PERFORMANCE " & _
                                        "WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE " & _
                                                      "WHERE PE01 < '" & m_CurrPE(0) & "')))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
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
   
   If m_CurrPE(0) = m_LastPE(0) And m_CurrPE(1) = m_LastPE(1) And m_CurrPE(2) = m_LastPE(2) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                  "PE02 = '" & m_CurrPE(1) & "' AND " & _
                  "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 = '" & m_CurrPE(1) & "' AND " & _
                                "PE03 > '" & m_CurrPE(2) & "' ) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                  "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 > '" & m_CurrPE(1) & "') AND " & _
                  "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                                        "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                              "PE02 > '" & m_CurrPE(1) & "'))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                          "WHERE PE01 > '" & m_CurrPE(0) & "') AND " & _
                  "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                          "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                        "WHERE PE01 > '" & m_CurrPE(0) & "')) AND " & _
                  "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                        "WHERE PE01 > '" & m_CurrPE(0) & "') AND " & _
                                "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                                        "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                                      "WHERE PE01 > '" & m_CurrPE(0) & "')))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close

   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrPE(0) = m_LastPE(0)
   m_CurrPE(1) = m_LastPE(1)
   m_CurrPE(2) = m_LastPE(2)
   
   UpdateCtrlData
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         ' 90.07.13 modify by louis (依照權限設定其工具列的按紐狀態)
         'tlbar.Buttons(1).Enabled = True
         'tlbar.Buttons(2).Enabled = True
         'tlbar.Buttons(3).Enabled = True
         'tlbar.Buttons(4).Enabled = True
         'tlbar.Buttons(6).Enabled = True
         'tlbar.Buttons(7).Enabled = True
         'tlbar.Buttons(8).Enabled = True
         'tlbar.Buttons(9).Enabled = True
         'tlbar.Buttons(11).Enabled = False
         'tlbar.Buttons(12).Enabled = False
         'tlbar.Buttons(14).Enabled = True
         
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
            tlbar.Buttons(6).Enabled = True
            tlbar.Buttons(7).Enabled = True
            tlbar.Buttons(8).Enabled = True
            tlbar.Buttons(9).Enabled = True
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

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 90.07.13 modify by louis
      ' 新增
      'Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
      '   If m_EditMode = 0 Then
      '      OnAction KeyCode
      '      KeyCode = 0
      '   End If
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
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            OnWork
            UpdateToolbarState
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         ClearField
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
         UpdateFieldNewData
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

Private Function IsTMExist(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsTMExist = False
      
   If IsEmptyText(strTM03) = True Then: strTM03 = "0"
   If IsEmptyText(strTM04) = True Then: strTM04 = "00"
   
   Select Case strTM01
      Case "T", "TF", "CFT", "FCT":
         strSql = "SELECT TM01, TM02, TM03, TM04 FROM TradeMark " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "TM02 = '" & strTM02 & "' AND " & _
                        "TM03 = '" & strTM03 & "' AND " & _
                        "TM04 = '" & strTM04 & "' "
      Case Else:
         strSql = "SELECT SP01, SP02, SP03, SP04 FROM ServicePractice " & _
                  "WHERE SP01 = '" & strTM01 & "' AND " & _
                        "SP02 = '" & strTM02 & "' AND " & _
                        "SP03 = '" & strTM03 & "' AND " & _
                        "SP04 = '" & strTM04 & "' "
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsTMExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm020505 = Nothing
End Sub

Private Sub textPE02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE20_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE21_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE22_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE23_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE24_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE25_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE26_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE27_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE28_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE29_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strPE01 As String, ByVal strPE02 As String, ByVal strPE03 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM Performance " & _
            "WHERE PE01 = '" & strPE01 & "' AND " & _
                  "PE02 = '" & strPE02 & "' AND " & _
                  "PE03 = '" & strPE03 & "'"
                  
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

' 新增記錄
Private Sub AddRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strPE01, strPE02, strPE03 As String
   
   strPE01 = textPE01
   strPE02 = textPE02
   strPE03 = (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strPE01, strPE02, strPE03) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO PERFORMANCE ("
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & m_FieldList(nIndex).fiNewData & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      QueryDB
      ShowCurrRecord strPE01, strPE02, strPE03
   End If
   
EXITSUB:
End Sub

' 修改記錄
Private Sub ModRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strPE01, strPE02, strPE03 As String
   
   strPE01 = textPE01
   strPE02 = textPE02
   strPE03 = (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
   
   strSql = "UPDATE PERFORMANCE SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_FieldList(nIndex).fiName & " = '" & m_FieldList(nIndex).fiNewData & "'"
            End If
         Else
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   
   strSql = strSql & " " & _
                  "WHERE PE01 = '" & strPE01 & "' AND " & _
                     "PE02 = '" & strPE02 & "' AND " & _
                     "PE03 = '" & strPE03 & "'"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      QueryDB
      ShowCurrRecord strPE01, strPE02, strPE03
   End If

End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strPE01, strPE02, strPE03 As String
   
   strPE01 = textPE01
   strPE02 = textPE02
   strPE03 = (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
   
   strSql = "DELETE FROM Performance " & _
            "WHERE PE01 = '" & strPE01 & "' AND " & _
                  "PE02 = '" & strPE02 & "' AND " & _
                  "PE03 = '" & strPE03 & "'"
                  
   cnnConnection.Execute strSql
   
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If strPE01 = m_LastPE(0) And strPE02 = m_LastPE(1) And strPE03 = m_LastPE(2) Then
      RefreshRange
   End If

   ShowCurrRecord strPE01, strPE02, strPE03
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strPE01 As String
   Dim strPE02 As String
   Dim strPE03 As String
   
   strPE01 = textPE01
   strPE02 = textPE02
   strPE03 = ConvertTYearToWYear(textPE03_1) & String(2 - Len(textPE03_2), "0") & textPE03_2
   
   QueryRecord = False
   
   If IsRecordExist(strPE01, strPE02, strPE03) = True Then
      m_CurrPE(0) = strPE01
      m_CurrPE(1) = strPE02
      m_CurrPE(2) = strPE03
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If
   
   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Select Case m_EditMode
      Case 1:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            AddRecord
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            ModRecord
         Else
            GoTo EXITSUB
         End If
      Case 3:
         DelRecord
         RefreshRange
      Case 4:
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

' 轉換西元年至民國年
Public Function ConvertWYearToTYear(ByVal strYear As String) As String
   Dim nYear As Integer
   
   nYear = Val(strYear)
   nYear = nYear - 1911
   ConvertWYearToTYear = nYear
End Function
' 轉換民國年至西元年
Public Function ConvertTYearToWYear(ByVal strYear As String) As String
   Dim nYear As Integer
   nYear = Val(strYear)
   nYear = nYear + 1911
   ConvertTYearToWYear = nYear
End Function

' 員工代碼
Private Sub textPE01_Validate(Cancel As Boolean)
   Dim strPE01 As String
   Dim strPE02 As String
   Dim strPE03 As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textST03 = Empty
   textA0902 = Empty
   If IsEmptyText(textPE01) = False Then
      Select Case m_EditMode
         ' 新增時必須在職
         Case 1:
            textST03 = GetStaffName(textPE01, False)
         Case Else:
            textST03 = GetStaffName(textPE01, True)
      End Select
      If IsEmptyText(textST03) = True Then
         Select Case m_EditMode
            Case 1, 2, 4:
               Cancel = True
               strTit = "資料檢核"
               strMsg = "員工代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE01_GotFocus
               GoTo EXITSUB
            Case Else:
         End Select
      End If
      
      textA0902 = GetDepartmentName(GetStaffDepartment(textPE01))
      
      ' 檢查Key是否存在
      'If m_EditMode = 1 Then
      '   strPE01 = textPE01
      '   strPE02 = textPE02
      '   strPE03 = ConvertTYearToWYear(textPE03_1) & textPE03_2
         ' 檢查記錄是否已存在
      '   If IsRecordExist(strPE01, strPE02, strPE03) = True Then
      '      strTit = "資料檢核"
      '      strMsg = "該筆記錄已存在"
      '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '      textPE01_GotFocus
      '      GoTo ExitSub
      '   End If
      'End If
   End If
EXITSUB:
End Sub

' 民國年
Private Sub textPE03_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If textPE03_1 <> Empty Then
      If IsNumeric(textPE03_1) = False Then
         Cancel = True
         strTit = "資料輸入有誤"
         strMsg = "請輸入民國年"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE03_1_GotFocus
      End If
   End If
End Sub
' 月份
Private Sub textPE03_2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If textPE03_2 <> Empty Then
      If IsNumeric(textPE03_2) = False Then
         Cancel = True
         strTit = "資料檢查"
         strMsg = "請輸入正確的月份"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE03_2_GotFocus
         GoTo EXITSUB
      End If
      If Val(textPE03_2) < 1 Or Val(textPE03_2) > 12 Then
         Cancel = True
         strTit = "資料檢查"
         strMsg = "請輸入正確的月份"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE03_2_GotFocus
         GoTo EXITSUB
      End If
      If Len(textPE03_2) = 1 Then
         textPE03_2 = "0" & textPE03_2
      End If
   End If

EXITSUB:
End Sub

' 目標點數
Private Sub textPE06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If textPE06 <> Empty Then
      If IsNumeric(textPE06) = False Then
         Cancel = True
         strTit = "資料輸入有誤"
         strMsg = "請輸入正確的數值"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE06_GotFocus
      End If
   End If
End Sub

' 其它點數
Private Sub textPE08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If textPE08 <> Empty Then
      If IsNumeric(textPE08) = False Then
         Cancel = True
         strTit = "資料輸入有誤"
         strMsg = "請輸入正確的數值"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE08_GotFocus
      End If
   End If
End Sub

' 初始化列表
Public Sub InitialGridList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 14
   
   grdList.ColWidth(0) = 300
   grdList.row = 0
      
   grdList.col = 1
   grdList.Text = "資料年月"
   grdList.ColWidth(1) = 800
   grdList.ColAlignment(1) = flexAlignCenterCenter
   grdList.col = 2
   grdList.Text = "系統類別"
   grdList.ColWidth(2) = 800
   grdList.ColAlignment(2) = flexAlignLeftCenter
   grdList.col = 3
   grdList.Text = "員工姓名"
   grdList.ColWidth(3) = 800
   grdList.ColAlignment(3) = flexAlignLeftCenter
   grdList.col = 4
   grdList.Text = "目標點數"
   grdList.ColWidth(4) = 800
   grdList.ColAlignment(4) = flexAlignLeftCenter
   grdList.col = 5
   grdList.Text = "勝訴率1"
   grdList.ColWidth(5) = 700
   grdList.ColAlignment(5) = flexAlignLeftCenter
   grdList.col = 6
   grdList.Text = "其它點數"
   grdList.ColWidth(6) = 800
   grdList.ColAlignment(6) = flexAlignLeftCenter
   grdList.col = 7
   grdList.Text = "勝訴率2"
   grdList.ColWidth(7) = 700
   grdList.ColAlignment(7) = flexAlignLeftCenter
   grdList.col = 8
   grdList.Text = "中文筆數"
   grdList.ColWidth(8) = 800
   grdList.ColAlignment(8) = flexAlignLeftCenter
   grdList.col = 9
   grdList.Text = "英文筆數"
   grdList.ColWidth(9) = 800
   grdList.ColAlignment(9) = flexAlignLeftCenter
   grdList.col = 10
   grdList.Text = "圖形筆數"
   grdList.ColWidth(10) = 800
   grdList.ColAlignment(10) = flexAlignLeftCenter
   grdList.col = 11
   grdList.Text = "過期筆數"
   grdList.ColWidth(11) = 800
   grdList.ColAlignment(11) = flexAlignLeftCenter
   grdList.col = 12
   grdList.Text = "未輸入筆數"
   grdList.ColWidth(12) = 1000
   grdList.ColAlignment(12) = flexAlignLeftCenter
   grdList.col = 13
   grdList.Text = "員工編號"
   grdList.ColWidth(13) = 0
End Sub

Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset)
   Dim strPE01, strPE02, strPE03 As String
   Dim nRow As Integer
   
   grdList.Clear
   InitialGridList
   
   If rsTmp.RecordCount > 0 Then
      strPE01 = rsTmp.Fields("PE01")
      strPE02 = rsTmp.Fields("PE02")
      strPE03 = rsTmp.Fields("PE03")
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         grdList.Rows = grdList.Rows + 1
         nRow = grdList.Rows - 1

        ' 目標年月
         If IsNull(rsTmp.Fields("PE03")) = False Then
            grdList.TextMatrix(nRow, 1) = ConvertWYearToTYear(Mid(rsTmp.Fields("PE03"), 1, 4)) & "/" & Mid(rsTmp.Fields("PE03"), 5, 2)
         End If
         ' 系統類別
         If IsNull(rsTmp.Fields("PE02")) = False Then
            grdList.TextMatrix(nRow, 2) = rsTmp.Fields("PE02")
         End If
         ' 員工姓名
         If IsNull(rsTmp.Fields("PE01")) = False Then
            grdList.TextMatrix(nRow, 13) = rsTmp.Fields("PE01")
            grdList.TextMatrix(nRow, 3) = GetStaffName(rsTmp.Fields("PE01"), True)
            If IsEmptyText(grdList.TextMatrix(nRow, 3)) = True Then: grdList.TextMatrix(nRow, 3) = rsTmp.Fields("PE01")
         End If
         ' 目標點數
         If IsNull(rsTmp.Fields("PE06")) = False Then
            grdList.TextMatrix(nRow, 4) = rsTmp.Fields("PE06")
         End If
         ' 商標勝訴率1
         If IsNull(rsTmp.Fields("PE18")) = False Then
            grdList.TextMatrix(nRow, 5) = rsTmp.Fields("PE18")
         End If
         ' 其它點數
         If IsNull(rsTmp.Fields("PE08")) = False Then
            grdList.TextMatrix(nRow, 6) = rsTmp.Fields("PE08")
         End If
         ' 商標勝訴率2
         If IsNull(rsTmp.Fields("PE19")) = False Then
            grdList.TextMatrix(nRow, 7) = rsTmp.Fields("PE19")
         End If
         ' 商標中文筆數
         If IsNull(rsTmp.Fields("PE12")) = False Then
            grdList.TextMatrix(nRow, 8) = rsTmp.Fields("PE12")
         End If
         ' 商標英文筆數
         If IsNull(rsTmp.Fields("PE13")) = False Then
            grdList.TextMatrix(nRow, 9) = rsTmp.Fields("PE13")
         End If
         ' 商標圖形筆數
         If IsNull(rsTmp.Fields("PE14")) = False Then
            grdList.TextMatrix(nRow, 10) = rsTmp.Fields("PE14")
         End If
         ' 商標過期筆數
         If IsNull(rsTmp.Fields("PE15")) = False Then
            grdList.TextMatrix(nRow, 11) = rsTmp.Fields("PE15")
         End If
         ' 商標未輸入筆數
         If IsNull(rsTmp.Fields("PE16")) = False Then
            grdList.TextMatrix(nRow, 12) = rsTmp.Fields("PE16")
         End If
            
         rsTmp.MoveNext
      Loop
      grdList.FixedRows = 1  'Added by Lydia 2023/10/16
   End If
End Sub

Private Sub grdList_Click()
   Dim strPE01 As String
   Dim strPE02 As String
   Dim strPE03 As String
   Dim nRow As Integer
   If grdList.row > 0 And grdList.row < grdList.Rows Then
      nRow = grdList.row
      strPE01 = grdList.TextMatrix(nRow, 13)
      strPE02 = grdList.TextMatrix(nRow, 2)
      strPE03 = CStr(Val(Left(grdList.TextMatrix(nRow, 1), Len(grdList.TextMatrix(nRow, 1)) - 3)) + 1911) & Right(grdList.TextMatrix(nRow, 1), 2)
      ShowCurrRecord strPE01, strPE02, strPE03
   End If
   grdList_ShowSelection
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
   ' 案件性質必須為延期的才可以選取
   If KeyCode = vbKeySpace Then
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
            grdList.Text = "V"
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub grdList_SelChange()
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textPE01.SetFocus
      Case 2: textPE06.SetFocus
      Case 4: textPE01.SetFocus
   End Select
End Sub

' 檢查本所案號是否存在
Private Function IsDataExist(ByVal strKey1 As String, ByVal StrKey2 As String, ByVal strKey3 As String, ByVal strKey4 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   IsDataExist = False
   
   If IsEmptyText(strKey3) = True Then: strKey3 = "0"
   If IsEmptyText(strKey4) = True Then: strKey4 = "00"
   
   Select Case strKey1
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         strSql = "SELECT * FROM TRADEMARK WHERE TM01 = '" & strKey1 & "' AND TM02 = '" & StrKey2 & "' AND TM03 = '" & strKey3 & "' AND TM04 = '" & strKey4 & "' "
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         strSql = "SELECT * FROM PATENT WHERE PA01 = '" & strKey1 & "' AND PA02 = '" & StrKey2 & "' AND PA03 = '" & strKey3 & "' AND PA04 = '" & strKey4 & "' "
      ' 讀取法務基本檔
      Case "L", "CFL", "FCL":
         strSql = "SELECT * FROM LAWCASE WHERE LC01 = '" & strKey1 & "' AND LC02 = '" & StrKey2 & "' AND LC03 = '" & strKey3 & "' AND LC04 = '" & strKey4 & "' "
      ' 讀取顧問案件基本檔
      Case "LA":
         strSql = "SELECT * FROM HIRECASE WHERE HC01 = '" & strKey1 & "' AND HC02 = '" & StrKey2 & "' AND HC03 = '" & strKey3 & "' AND HC04 = '" & strKey4 & "' "
      ' 讀取服務業務基本檔
      Case Else:
         strSql = "SELECT * FROM SERVICEPRACTICE WHERE SP01 = '" & strKey1 & "' AND SP02 = '" & StrKey2 & "' AND SP03 = '" & strKey3 & "' AND SP04 = '" & strKey4 & "' "
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsDataExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 檢查輸入的失誤案號是否有重覆
Private Function CheckPENoExist()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strPENo() As String
   Dim nCount As Integer
   Dim nX As Integer
   Dim nY As Integer
   Dim strDup As String
   
   CheckPENoExist = False
   nCount = 0
   If IsEmptyText(textPE20(0)) = False And IsEmptyText(textPE20(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE20(0) & textPE20(1) & textPE20(2) & String(1 - Len(textPE20(2)), "0") & textPE20(3) & String(2 - Len(textPE20(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE21(0)) = False And IsEmptyText(textPE21(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE21(0) & textPE21(1) & textPE21(2) & String(1 - Len(textPE21(2)), "0") & textPE21(3) & String(2 - Len(textPE21(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE22(0)) = False And IsEmptyText(textPE22(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE22(0) & textPE22(1) & textPE22(2) & String(1 - Len(textPE22(2)), "0") & textPE22(3) & String(2 - Len(textPE22(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE23(0)) = False And IsEmptyText(textPE23(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE23(0) & textPE23(1) & textPE23(2) & String(1 - Len(textPE23(2)), "0") & textPE23(3) & String(2 - Len(textPE23(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE24(0)) = False And IsEmptyText(textPE24(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE24(0) & textPE24(1) & textPE24(2) & String(1 - Len(textPE24(2)), "0") & textPE24(3) & String(2 - Len(textPE24(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE25(0)) = False And IsEmptyText(textPE25(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE25(0) & textPE25(1) & textPE25(2) & String(1 - Len(textPE25(2)), "0") & textPE25(3) & String(2 - Len(textPE25(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE26(0)) = False And IsEmptyText(textPE26(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE26(0) & textPE26(1) & textPE26(2) & String(1 - Len(textPE26(2)), "0") & textPE26(3) & String(2 - Len(textPE26(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE27(0)) = False And IsEmptyText(textPE27(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE27(0) & textPE27(1) & textPE27(2) & String(1 - Len(textPE27(2)), "0") & textPE27(3) & String(2 - Len(textPE27(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE28(0)) = False And IsEmptyText(textPE28(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE28(0) & textPE28(1) & textPE28(2) & String(1 - Len(textPE28(2)), "0") & textPE28(3) & String(2 - Len(textPE28(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE29(0)) = False And IsEmptyText(textPE29(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE29(0) & textPE29(1) & textPE29(2) & String(1 - Len(textPE29(2)), "0") & textPE29(3) & String(2 - Len(textPE29(3)), "0")
      nCount = nCount + 1
   End If
   
   strDup = Empty
   For nX = 0 To nCount - 1
      For nY = 0 To nCount - 1
         If nX <> nY Then
            If strPENo(nX) = strPENo(nY) Then
               strDup = strPENo(nX)
               CheckPENoExist = True
               Exit For
            End If
         End If
      Next nY
      If CheckPENoExist = True Then
         Exit For
      End If
   Next nX
   
   If CheckPENoExist = True Then
      strTit = "檢核資料"
      strMsg = "失誤案號<" & strDup & ">重覆"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
End Function

' 檢查輸入是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2, 4:
         ' 資料年
         If IsEmptyText(textPE03_1) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入資料年"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE03_1.SetFocus
            GoTo EXITSUB
         End If
         ' 資料月
         If IsEmptyText(textPE03_2) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入資料月"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE03_2.SetFocus
            GoTo EXITSUB
         End If
         ' 系統類別不可空白
         If IsEmptyText(textPE02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入系統類別"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE02.SetFocus
            GoTo EXITSUB
         End If
         ' 員工代號不可空白
         If IsEmptyText(textPE01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入員工代號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE01.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
   
    Select Case m_EditMode
      Case 1, 2:
         ' 目標點數不可空白
         If IsEmptyText(textPE06) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入目標點數"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE06.SetFocus
            GoTo EXITSUB
         End If
         '''''''''''''''''''''''''''''''''''''''''''''
         ' 檢查輸入的失誤案號是否完整
         If (IsEmptyText(textPE20(0)) = False And IsEmptyText(textPE20(1)) = True) Or (IsEmptyText(textPE20(0)) = True And IsEmptyText(textPE20(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE20(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE21(0)) = False And IsEmptyText(textPE21(1)) = True) Or (IsEmptyText(textPE21(0)) = True And IsEmptyText(textPE21(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE21(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE22(0)) = False And IsEmptyText(textPE22(1)) = True) Or (IsEmptyText(textPE22(0)) = True And IsEmptyText(textPE22(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE22(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE23(0)) = False And IsEmptyText(textPE23(1)) = True) Or (IsEmptyText(textPE23(0)) = True And IsEmptyText(textPE23(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE23(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE24(0)) = False And IsEmptyText(textPE24(1)) = True) Or (IsEmptyText(textPE24(0)) = True And IsEmptyText(textPE24(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE24(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE25(0)) = False And IsEmptyText(textPE25(1)) = True) Or (IsEmptyText(textPE25(0)) = True And IsEmptyText(textPE25(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE25(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE26(0)) = False And IsEmptyText(textPE26(1)) = True) Or (IsEmptyText(textPE26(0)) = True And IsEmptyText(textPE26(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE26(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE27(0)) = False And IsEmptyText(textPE27(1)) = True) Or (IsEmptyText(textPE27(0)) = True And IsEmptyText(textPE27(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE27(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE28(0)) = False And IsEmptyText(textPE28(1)) = True) Or (IsEmptyText(textPE28(0)) = True And IsEmptyText(textPE28(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE28(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE29(0)) = False And IsEmptyText(textPE29(1)) = True) Or (IsEmptyText(textPE29(0)) = True And IsEmptyText(textPE29(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE29(0).SetFocus
            GoTo EXITSUB
         End If
      
         '''''''''''''''''''''''''''''''''''''''''''''
         ' 檢查輸入的失誤案號是否存在於檔案中
         If IsEmptyText(textPE20(0)) = False And IsEmptyText(textPE20(1)) = False Then
            If IsDataExist(textPE20(0), textPE20(1), textPE20(2), textPE20(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE20(0) & "-" & textPE20(1) & "-" & textPE20(2) & "-" & textPE20(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE20(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE21(0)) = False And IsEmptyText(textPE21(1)) = False Then
            If IsDataExist(textPE21(0), textPE21(1), textPE21(2), textPE21(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE21(0) & "-" & textPE21(1) & "-" & textPE21(2) & "-" & textPE21(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE21(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE22(0)) = False And IsEmptyText(textPE22(1)) = False Then
            If IsDataExist(textPE22(0), textPE22(1), textPE22(2), textPE22(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE22(0) & "-" & textPE22(1) & "-" & textPE22(2) & "-" & textPE22(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE22(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE23(0)) = False And IsEmptyText(textPE23(1)) = False Then
            If IsDataExist(textPE23(0), textPE23(1), textPE23(2), textPE23(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE23(0) & "-" & textPE23(1) & "-" & textPE23(2) & "-" & textPE23(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE23(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE24(0)) = False And IsEmptyText(textPE24(1)) = False Then
            If IsDataExist(textPE24(0), textPE24(1), textPE24(2), textPE24(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE24(0) & "-" & textPE24(1) & "-" & textPE24(2) & "-" & textPE24(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE24(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE25(0)) = False And IsEmptyText(textPE25(1)) = False Then
            If IsDataExist(textPE25(0), textPE25(1), textPE25(2), textPE25(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE25(0) & "-" & textPE25(1) & "-" & textPE25(2) & "-" & textPE25(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE25(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE26(0)) = False And IsEmptyText(textPE26(1)) = False Then
            If IsDataExist(textPE26(0), textPE26(1), textPE26(2), textPE26(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE26(0) & "-" & textPE26(1) & "-" & textPE26(2) & "-" & textPE26(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE26(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE27(0)) = False And IsEmptyText(textPE27(1)) = False Then
            If IsDataExist(textPE27(0), textPE27(1), textPE27(2), textPE27(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE27(0) & "-" & textPE27(1) & "-" & textPE27(2) & "-" & textPE27(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE27(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE28(0)) = False And IsEmptyText(textPE28(1)) = False Then
            If IsDataExist(textPE28(0), textPE28(1), textPE28(2), textPE28(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE28(0) & "-" & textPE28(1) & "-" & textPE28(2) & "-" & textPE28(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE28(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE29(0)) = False And IsEmptyText(textPE29(1)) = False Then
            If IsDataExist(textPE29(0), textPE29(1), textPE29(2), textPE29(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE29(0) & "-" & textPE29(1) & "-" & textPE29(2) & "-" & textPE29(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE29(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         '''''''''''''''''''''''''''''''''''''''''''''
         ' 檢查所輸入的失誤案號是否重覆
         If CheckPENoExist() = True Then
            GoTo EXITSUB
         End If
   End Select
   
   CheckDataValid = True
EXITSUB:
End Function

' 系統類別
Private Sub textPE02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strPE01 As String
   Dim strPE02 As String
   Dim strPE03 As String
   Cancel = False
   
   If IsEmptyText(textPE02) = False Then
      ' 檢查系統種類對照表是否存在
      If m_EditMode = 1 Or m_EditMode = 4 Then
         If IsCorrectSysKind(textPE02) = False Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "系統類別不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE02_GotFocus
            GoTo EXITSUB
         End If
      End If
      ' 系統類別只可輸入T類
      If Mid(textPE02, 1, 1) <> "T" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "系統類別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE02_GotFocus
         GoTo EXITSUB
      End If
      
      ' 檢查Key是否存在
      If m_EditMode = 1 Then
         strPE01 = textPE01
         strPE02 = textPE02
         strPE03 = ConvertTYearToWYear(textPE03_1) & textPE03_2
         ' 檢查記錄是否已存在
         If IsRecordExist(strPE01, strPE02, strPE03) = True Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "該筆記錄已存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE01_GotFocus
            GoTo EXITSUB
         End If
      End If
   End If
EXITSUB:
End Sub

' 月
Private Sub textPE03_Month_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPE03_Month) = False Then
      textPE03_Month = String(2 - Len(textPE03_Month), "0") & textPE03_Month
      If IsNumeric(textPE03_Month) = False Then
         Cancel = True
         strTit = "資料輸入有誤"
         strMsg = "請輸入月份"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE03_Month_GotFocus
         GoTo EXITSUB
      End If
      If Val(textPE03_Month) < 1 Or Val(textPE03_Month) > 12 Then
         Cancel = True
         strTit = "資料輸入有誤"
         strMsg = "請輸入月份"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE03_Month_GotFocus
         GoTo EXITSUB
      End If
      If Len(textPE03_Month) = 1 Then
         textPE03_Month = "0" & textPE03_2
      End If
   End If

EXITSUB:
End Sub

' 年
Private Sub textPE03_Year_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If textPE03_Year <> Empty Then
      If IsNumeric(textPE03_Year) = False Then
         Cancel = True
         strTit = "資料輸入有誤"
         strMsg = "請輸入民國年"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE03_Year_GotFocus
      End If
   End If
End Sub

' 本所案號
Private Sub textPE20_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE20(Index)) = False Then
      Select Case Index
         Case 1:
            textPE20(Index) = String(6 - Len(textPE20(Index)), "0") & textPE20(Index)
         Case 2:
            textPE20(Index) = textPE20(Index) & String(1 - Len(textPE20(Index)), "0")
         Case 3:
            textPE20(Index) = textPE20(Index) & String(2 - Len(textPE20(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE21_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE21(Index)) = False Then
      Select Case Index
         Case 1:
            textPE21(Index) = String(6 - Len(textPE21(Index)), "0") & textPE21(Index)
         Case 2:
            textPE21(Index) = textPE21(Index) & String(1 - Len(textPE21(Index)), "0")
         Case 3:
            textPE21(Index) = textPE21(Index) & String(2 - Len(textPE21(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE22_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE22(Index)) = False Then
      Select Case Index
         Case 1:
            textPE22(Index) = String(6 - Len(textPE22(Index)), "0") & textPE22(Index)
         Case 2:
            textPE22(Index) = textPE22(Index) & String(1 - Len(textPE22(Index)), "0")
         Case 3:
            textPE22(Index) = textPE22(Index) & String(2 - Len(textPE22(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE23_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE23(Index)) = False Then
      Select Case Index
         Case 1:
            textPE23(Index) = String(6 - Len(textPE23(Index)), "0") & textPE23(Index)
         Case 2:
            textPE23(Index) = textPE23(Index) & String(1 - Len(textPE23(Index)), "0")
         Case 3:
            textPE23(Index) = textPE23(Index) & String(2 - Len(textPE23(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE24_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE24(Index)) = False Then
      Select Case Index
         Case 1:
            textPE24(Index) = String(6 - Len(textPE24(Index)), "0") & textPE24(Index)
         Case 2:
            textPE24(Index) = textPE24(Index) & String(1 - Len(textPE24(Index)), "0")
         Case 3:
            textPE24(Index) = textPE24(Index) & String(2 - Len(textPE24(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE25_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE25(Index)) = False Then
      Select Case Index
         Case 1:
            textPE25(Index) = String(6 - Len(textPE25(Index)), "0") & textPE25(Index)
         Case 2:
            textPE25(Index) = textPE25(Index) & String(1 - Len(textPE25(Index)), "0")
         Case 3:
            textPE25(Index) = textPE25(Index) & String(2 - Len(textPE25(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE26_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE26(Index)) = False Then
      Select Case Index
         Case 1:
            textPE26(Index) = String(6 - Len(textPE26(Index)), "0") & textPE26(Index)
         Case 2:
            textPE26(Index) = textPE26(Index) & String(1 - Len(textPE26(Index)), "0")
         Case 3:
            textPE26(Index) = textPE26(Index) & String(2 - Len(textPE26(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE27_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE27(Index)) = False Then
      Select Case Index
         Case 1:
            textPE27(Index) = String(6 - Len(textPE27(Index)), "0") & textPE27(Index)
         Case 2:
            textPE27(Index) = textPE27(Index) & String(1 - Len(textPE27(Index)), "0")
         Case 3:
            textPE27(Index) = textPE27(Index) & String(2 - Len(textPE27(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE28_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE28(Index)) = False Then
      Select Case Index
         Case 1:
            textPE28(Index) = String(6 - Len(textPE28(Index)), "0") & textPE28(Index)
         Case 2:
            textPE28(Index) = textPE28(Index) & String(1 - Len(textPE28(Index)), "0")
         Case 3:
            textPE28(Index) = textPE28(Index) & String(2 - Len(textPE28(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE29_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE29(Index)) = False Then
      Select Case Index
         Case 1:
            textPE29(Index) = String(6 - Len(textPE29(Index)), "0") & textPE29(Index)
         Case 2:
            textPE29(Index) = textPE29(Index) & String(1 - Len(textPE29(Index)), "0")
         Case 3:
            textPE29(Index) = textPE29(Index) & String(2 - Len(textPE29(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE01_GotFocus()
   InverseTextBox textPE01
End Sub

Private Sub textPE02_GotFocus()
   InverseTextBox textPE02
End Sub

Private Sub textPE03_1_GotFocus()
   InverseTextBox textPE03_1
End Sub

Private Sub textPE03_2_GotFocus()
   InverseTextBox textPE03_2
End Sub

Private Sub textPE03_Year_GotFocus()
   InverseTextBox textPE03_Year
End Sub

Private Sub textPE03_Month_GotFocus()
   InverseTextBox textPE03_Month
End Sub

Private Sub textPE06_GotFocus()
   InverseTextBox textPE06
End Sub

Private Sub textPE08_GotFocus()
   InverseTextBox textPE08
End Sub

Private Sub textPE20_GotFocus(Index As Integer)
   InverseTextBox textPE20(Index)
End Sub

Private Sub textPE21_GotFocus(Index As Integer)
   InverseTextBox textPE21(Index)
End Sub

Private Sub textPE22_GotFocus(Index As Integer)
   InverseTextBox textPE22(Index)
End Sub

Private Sub textPE23_GotFocus(Index As Integer)
   InverseTextBox textPE23(Index)
End Sub

Private Sub textPE24_GotFocus(Index As Integer)
   InverseTextBox textPE24(Index)
End Sub

Private Sub textPE25_GotFocus(Index As Integer)
   InverseTextBox textPE25(Index)
End Sub

Private Sub textPE26_GotFocus(Index As Integer)
   InverseTextBox textPE26(Index)
End Sub

Private Sub textPE27_GotFocus(Index As Integer)
   InverseTextBox textPE27(Index)
End Sub

Private Sub textPE28_GotFocus(Index As Integer)
   InverseTextBox textPE28(Index)
End Sub

Private Sub textPE29_GotFocus(Index As Integer)
   InverseTextBox textPE29(Index)
End Sub
'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textPE01.Enabled = True Then
   Cancel = False
   textPE01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE02.Enabled = True Then
   Cancel = False
   textPE02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE03_1.Enabled = True Then
   Cancel = False
   textPE03_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE03_2.Enabled = True Then
   Cancel = False
   textPE03_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE03_Month.Enabled = True Then
   Cancel = False
   textPE03_Month_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE03_Year.Enabled = True Then
   Cancel = False
   textPE03_Year_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE06.Enabled = True Then
   Cancel = False
   textPE06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE08.Enabled = True Then
   Cancel = False
   textPE08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

For Each objTxt In Me.textPE20
   If objTxt.Enabled = True Then
      Cancel = False
      textPE20_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In Me.textPE21
   If objTxt.Enabled = True Then
      Cancel = False
      textPE21_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In Me.textPE22
   If objTxt.Enabled = True Then
      Cancel = False
      textPE22_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In Me.textPE23
   If objTxt.Enabled = True Then
      Cancel = False
      textPE23_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In Me.textPE24
   If objTxt.Enabled = True Then
      Cancel = False
      textPE24_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In Me.textPE25
   If objTxt.Enabled = True Then
      Cancel = False
      textPE25_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In Me.textPE26
   If objTxt.Enabled = True Then
      Cancel = False
      textPE26_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In Me.textPE27
   If objTxt.Enabled = True Then
      Cancel = False
      textPE27_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In Me.textPE28
   If objTxt.Enabled = True Then
      Cancel = False
      textPE28_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In Me.textPE29
   If objTxt.Enabled = True Then
      Cancel = False
      textPE29_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function

