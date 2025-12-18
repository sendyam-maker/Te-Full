VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_C 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件資料及案件進度查詢 (案件進度資料)"
   ClientHeight    =   6360
   ClientLeft      =   950
   ClientTop       =   2150
   ClientWidth     =   8930
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8930
   Visible         =   0   'False
   Begin VB.TextBox textCP09 
      Height          =   270
      Left            =   930
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   331
      Top             =   240
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IDS清單"
      Height          =   270
      Left            =   120
      TabIndex        =   267
      Top             =   15
      Visible         =   0   'False
      Width           =   1085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "未發文原因"
      Height          =   270
      Left            =   2950
      TabIndex        =   226
      Top             =   15
      Width           =   1085
   End
   Begin VB.CommandButton Command4 
      Caption         =   "出庭人員"
      Height          =   270
      Left            =   4050
      TabIndex        =   0
      Top             =   15
      Width           =   1485
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "下一筆"
      Height          =   270
      Index           =   2
      Left            =   7080
      TabIndex        =   2
      Top             =   15
      Width           =   990
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   270
      Index           =   3
      Left            =   8085
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "相關收文資料"
      CausesValidation=   0   'False
      Height          =   270
      Index           =   4
      Left            =   5550
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   15
      Width           =   1500
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4890
      Left            =   30
      TabIndex        =   10
      Top             =   1170
      Width           =   8895
      _ExtentX        =   15681
      _ExtentY        =   8608
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   423
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm100101_C.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "textCP152"
      Tab(0).Control(1)=   "textCP113"
      Tab(0).Control(2)=   "textCP114"
      Tab(0).Control(3)=   "TextCP119"
      Tab(0).Control(4)=   "textCP118"
      Tab(0).Control(5)=   "textCP44"
      Tab(0).Control(6)=   "textCP15"
      Tab(0).Control(7)=   "textCP29"
      Tab(0).Control(8)=   "textCP84"
      Tab(0).Control(9)=   "textCP28"
      Tab(0).Control(10)=   "textCP10_2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "textCP12_2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textCP58_2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textCP12"
      Tab(0).Control(14)=   "textCP13"
      Tab(0).Control(15)=   "textCP14"
      Tab(0).Control(16)=   "textCP27"
      Tab(0).Control(17)=   "textCP06"
      Tab(0).Control(18)=   "textCP43"
      Tab(0).Control(19)=   "textCP08"
      Tab(0).Control(20)=   "textCP21"
      Tab(0).Control(21)=   "textCP48"
      Tab(0).Control(22)=   "textCP31"
      Tab(0).Control(23)=   "textCP57"
      Tab(0).Control(24)=   "textCP58"
      Tab(0).Control(25)=   "textCP22"
      Tab(0).Control(26)=   "textCP45"
      Tab(0).Control(27)=   "textCP25"
      Tab(0).Control(28)=   "textCP07"
      Tab(0).Control(29)=   "textCP05"
      Tab(0).Control(30)=   "textCP10"
      Tab(0).Control(31)=   "textCP82"
      Tab(0).Control(32)=   "textCP83"
      Tab(0).Control(33)=   "textCP24"
      Tab(0).Control(34)=   "textCP23"
      Tab(0).Control(35)=   "textCP26"
      Tab(0).Control(36)=   "lblNameAgent"
      Tab(0).Control(37)=   "textCP44_2"
      Tab(0).Control(38)=   "textCP83_2"
      Tab(0).Control(39)=   "textCP29_2"
      Tab(0).Control(40)=   "textCP14_2"
      Tab(0).Control(41)=   "textCP13_2"
      Tab(0).Control(42)=   "textCP64"
      Tab(0).Control(43)=   "Label1(1)"
      Tab(0).Control(44)=   "Label6(4)"
      Tab(0).Control(45)=   "Label1(12)"
      Tab(0).Control(46)=   "Label1(5)"
      Tab(0).Control(47)=   "Label15"
      Tab(0).Control(48)=   "Label6(2)"
      Tab(0).Control(49)=   "Label9"
      Tab(0).Control(50)=   "Label50"
      Tab(0).Control(51)=   "Label8"
      Tab(0).Control(52)=   "Label6(5)"
      Tab(0).Control(53)=   "Label4"
      Tab(0).Control(54)=   "Label47"
      Tab(0).Control(55)=   "Label7(0)"
      Tab(0).Control(56)=   "Label7(5)"
      Tab(0).Control(57)=   "Label31"
      Tab(0).Control(58)=   "Label26"
      Tab(0).Control(59)=   "Label16(1)"
      Tab(0).Control(60)=   "Label28"
      Tab(0).Control(61)=   "Label1(0)"
      Tab(0).Control(62)=   "Label1(2)"
      Tab(0).Control(63)=   "Label17"
      Tab(0).Control(64)=   "Label6(0)"
      Tab(0).Control(65)=   "Label5"
      Tab(0).Control(66)=   "Label11"
      Tab(0).Control(67)=   "Label12(0)"
      Tab(0).Control(68)=   "Label23"
      Tab(0).Control(69)=   "Label30(0)"
      Tab(0).Control(70)=   "Label32"
      Tab(0).Control(71)=   "Label13"
      Tab(0).Control(72)=   "Label21"
      Tab(0).Control(73)=   "Label35"
      Tab(0).Control(74)=   "Label16(0)"
      Tab(0).Control(75)=   "Label6(3)"
      Tab(0).Control(76)=   "Label14(1)"
      Tab(0).Control(77)=   "Label18(0)"
      Tab(0).Control(78)=   "Label19(0)"
      Tab(0).Control(79)=   "Label25"
      Tab(0).ControlCount=   80
      TabCaption(1)   =   "相關資料"
      TabPicture(1)   =   "frm100101_C.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPage"
      Tab(1).Control(1)=   "textCP167"
      Tab(1).Control(2)=   "textCP168"
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(4)=   "textCP71"
      Tab(1).Control(5)=   "textCP71_2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "fraTF"
      Tab(1).Control(7)=   "frmBill"
      Tab(1).Control(8)=   "textCP148"
      Tab(1).Control(9)=   "textCP140"
      Tab(1).Control(10)=   "textCP138"
      Tab(1).Control(11)=   "textCP137"
      Tab(1).Control(12)=   "textCP135"
      Tab(1).Control(13)=   "textCP136"
      Tab(1).Control(14)=   "textCP120"
      Tab(1).Control(15)=   "textCP121"
      Tab(1).Control(16)=   "textCP11_2"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "textCP17"
      Tab(1).Control(18)=   "textCP19"
      Tab(1).Control(19)=   "textCP33"
      Tab(1).Control(20)=   "textCP34"
      Tab(1).Control(21)=   "textCP46"
      Tab(1).Control(22)=   "textCP47"
      Tab(1).Control(23)=   "textCP32"
      Tab(1).Control(24)=   "textCP20"
      Tab(1).Control(25)=   "textCP11"
      Tab(1).Control(26)=   "textCP59"
      Tab(1).Control(27)=   "textCP30"
      Tab(1).Control(28)=   "textCP60"
      Tab(1).Control(29)=   "textCP61"
      Tab(1).Control(30)=   "textCP62"
      Tab(1).Control(31)=   "textCP63"
      Tab(1).Control(32)=   "textCP18"
      Tab(1).Control(33)=   "textCP16"
      Tab(1).Control(34)=   "textCP81"
      Tab(1).Control(35)=   "textCP88"
      Tab(1).Control(36)=   "textCP87"
      Tab(1).Control(37)=   "Frame1"
      Tab(1).Control(38)=   "lblCP168"
      Tab(1).Control(39)=   "lblCP167"
      Tab(1).Control(40)=   "textCP49"
      Tab(1).Control(41)=   "lblCP71"
      Tab(1).Control(42)=   "Label29"
      Tab(1).Control(43)=   "Label30(1)"
      Tab(1).Control(44)=   "Label2(2)"
      Tab(1).Control(45)=   "Label22"
      Tab(1).Control(46)=   "Label1(121)"
      Tab(1).Control(47)=   "lblCP137"
      Tab(1).Control(48)=   "lblCP136"
      Tab(1).Control(49)=   "lblCP135"
      Tab(1).Control(50)=   "lblCP138"
      Tab(1).Control(51)=   "Label20(5)"
      Tab(1).Control(52)=   "Label20(6)"
      Tab(1).Control(53)=   "Label10"
      Tab(1).Control(54)=   "Label6(1)"
      Tab(1).Control(55)=   "lblCP49"
      Tab(1).Control(56)=   "Label34"
      Tab(1).Control(57)=   "Label30(3)"
      Tab(1).Control(58)=   "Label30(2)"
      Tab(1).Control(59)=   "Label12(3)"
      Tab(1).Control(60)=   "Label12(1)"
      Tab(1).Control(61)=   "Label12(2)"
      Tab(1).Control(62)=   "lblCP19"
      Tab(1).Control(63)=   "Label37"
      Tab(1).Control(64)=   "Label38"
      Tab(1).Control(65)=   "Label39"
      Tab(1).Control(66)=   "Label14(0)"
      Tab(1).Control(67)=   "Label14(2)"
      Tab(1).Control(68)=   "Label14(3)"
      Tab(1).Control(69)=   "Label1(3)"
      Tab(1).Control(70)=   "Label6(12)"
      Tab(1).Control(71)=   "Label43"
      Tab(1).Control(72)=   "lblCP81"
      Tab(1).Control(73)=   "Label14(4)"
      Tab(1).Control(74)=   "Label14(5)"
      Tab(1).ControlCount=   75
      TabCaption(2)   =   "移轉/授權"
      TabPicture(2)   =   "frm100101_C.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "textCP54"
      Tab(2).Control(1)=   "textCP53"
      Tab(2).Control(2)=   "textCP72"
      Tab(2).Control(3)=   "textCP56"
      Tab(2).Control(4)=   "textCP55"
      Tab(2).Control(5)=   "textCP93"
      Tab(2).Control(6)=   "textCP94"
      Tab(2).Control(7)=   "textCP95"
      Tab(2).Control(8)=   "textCP96"
      Tab(2).Control(9)=   "textCP89"
      Tab(2).Control(10)=   "textCP90"
      Tab(2).Control(11)=   "textCP91"
      Tab(2).Control(12)=   "textCP92"
      Tab(2).Control(13)=   "textCP55_2"
      Tab(2).Control(14)=   "textCP93_2"
      Tab(2).Control(15)=   "textCP94_2"
      Tab(2).Control(16)=   "textCP95_2"
      Tab(2).Control(17)=   "textCP96_2"
      Tab(2).Control(18)=   "textCP56_2"
      Tab(2).Control(19)=   "textCP89_2"
      Tab(2).Control(20)=   "textCP90_2"
      Tab(2).Control(21)=   "textCP91_2"
      Tab(2).Control(22)=   "textCP92_2"
      Tab(2).Control(23)=   "textCP50"
      Tab(2).Control(24)=   "textCP51"
      Tab(2).Control(25)=   "textCP52"
      Tab(2).Control(26)=   "Label20(7)"
      Tab(2).Control(27)=   "Line1"
      Tab(2).Control(28)=   "Label20(2)"
      Tab(2).Control(29)=   "Label20(1)"
      Tab(2).Control(30)=   "Label33(0)"
      Tab(2).Control(31)=   "Label20(3)"
      Tab(2).Control(32)=   "Label20(4)"
      Tab(2).Control(33)=   "Label7(7)"
      Tab(2).Control(34)=   "Label7(8)"
      Tab(2).Control(35)=   "Label7(1)"
      Tab(2).Control(36)=   "Label7(2)"
      Tab(2).Control(37)=   "Label7(3)"
      Tab(2).Control(38)=   "Label7(4)"
      Tab(2).Control(39)=   "Label7(6)"
      Tab(2).Control(40)=   "Label7(9)"
      Tab(2).Control(41)=   "Label7(10)"
      Tab(2).Control(42)=   "Label7(11)"
      Tab(2).ControlCount=   43
      TabCaption(3)   =   "對造/其他"
      TabPicture(3)   =   "frm100101_C.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "textCP86"
      Tab(3).Control(1)=   "textCP117"
      Tab(3).Control(2)=   "textCP35"
      Tab(3).Control(3)=   "textCP80"
      Tab(3).Control(4)=   "textCP36"
      Tab(3).Control(5)=   "lblCP86"
      Tab(3).Control(6)=   "lblCP86_1"
      Tab(3).Control(7)=   "textCP37_1"
      Tab(3).Control(8)=   "textCP37"
      Tab(3).Control(9)=   "textCP38"
      Tab(3).Control(10)=   "textCP39"
      Tab(3).Control(11)=   "textCP40"
      Tab(3).Control(12)=   "textCP41"
      Tab(3).Control(13)=   "textCP42"
      Tab(3).Control(14)=   "Label33(2)"
      Tab(3).Control(15)=   "textCP144"
      Tab(3).Control(16)=   "Label24"
      Tab(3).Control(17)=   "Label51"
      Tab(3).Control(18)=   "Label18(5)"
      Tab(3).Control(19)=   "Label33(1)"
      Tab(3).Control(20)=   "Label18(4)"
      Tab(3).Control(21)=   "Label18(1)"
      Tab(3).Control(22)=   "Label19(3)"
      Tab(3).Control(23)=   "Label18(2)"
      Tab(3).Control(24)=   "Label18(3)"
      Tab(3).Control(25)=   "Label19(1)"
      Tab(3).Control(26)=   "Label19(2)"
      Tab(3).ControlCount=   27
      TabCaption(4)   =   "收據帳目"
      TabPicture(4)   =   "frm100101_C.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblCP79"
      Tab(4).Control(1)=   "lblCP78"
      Tab(4).Control(2)=   "lblCP77"
      Tab(4).Control(3)=   "lblCP76"
      Tab(4).Control(4)=   "lblCP75"
      Tab(4).Control(5)=   "lblCP74"
      Tab(4).Control(6)=   "lblCP73"
      Tab(4).Control(7)=   "Label49"
      Tab(4).Control(8)=   "Label48"
      Tab(4).Control(9)=   "Label46"
      Tab(4).Control(10)=   "Label45"
      Tab(4).Control(11)=   "Label44"
      Tab(4).Control(12)=   "Label27"
      Tab(4).Control(13)=   "Label40"
      Tab(4).ControlCount=   14
      TabCaption(5)   =   "發文室"
      TabPicture(5)   =   "frm100101_C.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Label16(6)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label14(7)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label16(4)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label16(2)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label14(6)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label16(7)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Label16(8)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Label16(9)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "textCP131"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "lblCP153"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "textCP129"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "textCP126"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "textCP125"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "textCP124"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "textCP123"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "textCP130"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "textCP132"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "SSTab2"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).ControlCount=   18
      Begin VB.CommandButton cmdPage 
         BackColor       =   &H00C0C0FF&
         Caption         =   "增刪頁數"
         Height          =   285
         Left            =   -69990
         Style           =   1  '圖片外觀
         TabIndex        =   336
         Top             =   3660
         Width           =   1065
      End
      Begin VB.TextBox textCP167 
         Height          =   285
         Left            =   -67995
         TabIndex        =   37
         Top             =   3360
         Width           =   420
      End
      Begin VB.TextBox textCP168 
         Height          =   285
         Left            =   -67995
         TabIndex        =   39
         Top             =   3675
         Width           =   420
      End
      Begin VB.TextBox textCP86 
         Height          =   285
         Left            =   -73020
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   280
         Top             =   3960
         Width           =   255
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Enabled         =   0   'False
         Height          =   225
         Left            =   -68040
         TabIndex        =   326
         Top             =   2190
         Width           =   2115
         Begin VB.OptionButton Option1 
            Caption         =   "之後"
            Height          =   195
            Index           =   2
            Left            =   1260
            TabIndex        =   327
            Top             =   0
            Width           =   660
         End
         Begin VB.OptionButton Option1 
            Caption         =   "之前"
            Height          =   195
            Index           =   1
            Left            =   630
            TabIndex        =   328
            Top             =   0
            Width           =   705
         End
         Begin VB.OptionButton Option1 
            Caption         =   "當天"
            Height          =   195
            Index           =   0
            Left            =   -30
            TabIndex        =   329
            Top             =   0
            Width           =   705
         End
      End
      Begin VB.TextBox textCP117 
         Height          =   285
         Left            =   -68595
         Locked          =   -1  'True
         TabIndex        =   279
         Top             =   3600
         Width           =   2310
      End
      Begin VB.TextBox textCP35 
         Height          =   285
         Left            =   -73230
         Locked          =   -1  'True
         MaxLength       =   32
         TabIndex        =   278
         Top             =   3600
         Width           =   3225
      End
      Begin VB.TextBox textCP71 
         Height          =   285
         Left            =   -70290
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   50
         Top             =   2151
         Width           =   975
      End
      Begin VB.TextBox textCP71_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -69300
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   2166
         Width           =   1220
      End
      Begin VB.Frame fraTF 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   600
         Left            =   -72870
         TabIndex        =   249
         Top             =   4260
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox txtTF23 
            Alignment       =   1  '靠右對齊
            Height          =   270
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   41
            Top             =   0
            Width           =   840
         End
         Begin VB.TextBox txtTF19 
            Alignment       =   1  '靠右對齊
            Height          =   270
            Left            =   4440
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   42
            Top             =   0
            Width           =   840
         End
         Begin VB.TextBox txtTF20 
            Height          =   270
            Left            =   7110
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   43
            Top             =   0
            Width           =   1455
         End
         Begin MSForms.TextBox txtTF37 
            Height          =   300
            Left            =   1320
            TabIndex        =   44
            Top             =   300
            Width           =   7260
            VariousPropertyBits=   -1467989985
            ScrollBars      =   2
            Size            =   "12806;529"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "翻譯瑕疵備註："
            Height          =   180
            Index           =   11
            Left            =   0
            TabIndex        =   265
            Top             =   300
            Width           =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "相似案號："
            Height          =   180
            Index           =   8
            Left            =   6150
            TabIndex        =   252
            Top             =   0
            Width           =   900
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "相似度：                         %"
            Height          =   180
            Index           =   7
            Left            =   3660
            TabIndex        =   251
            Top             =   0
            Width           =   1980
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "原文字數："
            Height          =   180
            Index           =   6
            Left            =   0
            TabIndex        =   250
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.Frame frmBill 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   315
         Left            =   -72960
         TabIndex        =   260
         Top             =   3990
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox txtBillFee 
            Alignment       =   1  '靠右對齊
            Height          =   270
            Left            =   4410
            Locked          =   -1  'True
            TabIndex        =   264
            Top             =   0
            Width           =   1290
         End
         Begin VB.TextBox txtBillWords 
            Alignment       =   1  '靠右對齊
            Height          =   270
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   262
            Top             =   0
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "翻譯社請款金額："
            Height          =   180
            Index           =   10
            Left            =   2910
            TabIndex        =   263
            Top             =   0
            Width           =   1440
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "翻譯社請款字數："
            Height          =   180
            Index           =   9
            Left            =   0
            TabIndex        =   261
            Top             =   0
            Width           =   1440
         End
      End
      Begin VB.TextBox textCP152 
         Height          =   285
         Left            =   -67200
         MaxLength       =   7
         TabIndex        =   255
         Top             =   1428
         Width           =   870
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2595
         Left            =   45
         TabIndex        =   227
         Top             =   2100
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   4568
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "客戶通知函"
         TabPicture(0)   =   "frm100101_C.frx":00A8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblLP26"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label16(15)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label16(13)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label16(10)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblLP23"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label16(3)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblLP21"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblCP127128"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label16(5)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lblLP0718"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "lblLP0517"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "lblLP11"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label16(16)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label16(14)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label16(12)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label16(11)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label16(19)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "lblLP3940"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label16(20)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Label16(21)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "lblLP4748"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Label16(22)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "lblLP04"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "lblLP06"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "lblCP154"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "lblLP38"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "lblLP46"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "lblLP20"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "lblLP22"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "txtLP37"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "txtLP12"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "txtLP49"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "txtLP24"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "txtLP25"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "cmdCancelLetter"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "cmdCancelConfirm"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "cmdCancelLP05"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).ControlCount=   37
         TabCaption(1)   =   "指示信"
         TabPicture(1)   =   "frm100101_C.frx":00C4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label16(17)"
         Tab(1).Control(1)=   "Label16(18)"
         Tab(1).Control(2)=   "lblAF0708"
         Tab(1).Control(3)=   "lblAF1112"
         Tab(1).Control(4)=   "lblAF06"
         Tab(1).Control(5)=   "lblAF14"
         Tab(1).ControlCount=   6
         Begin VB.CommandButton cmdCancelLP05 
            Caption         =   "取消"
            Height          =   285
            Left            =   4176
            TabIndex        =   337
            Top             =   504
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.CommandButton cmdCancelConfirm 
            Caption         =   "取消"
            Height          =   285
            Left            =   4170
            TabIndex        =   266
            Top             =   780
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.CommandButton cmdCancelLetter 
            Caption         =   "不通知"
            Height          =   285
            Left            =   135
            TabIndex        =   256
            Top             =   60
            Visible         =   0   'False
            Width           =   780
         End
         Begin MSForms.TextBox txtLP25 
            Height          =   405
            Left            =   6090
            TabIndex        =   325
            Top             =   2130
            Width           =   2625
            VariousPropertyBits=   -1467989989
            MaxLength       =   200
            ScrollBars      =   2
            Size            =   "4630;714"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtLP24 
            Height          =   435
            Left            =   6090
            TabIndex        =   324
            Top             =   1695
            Width           =   2625
            VariousPropertyBits=   -1467989989
            MaxLength       =   200
            ScrollBars      =   2
            Size            =   "4630;767"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtLP49 
            Height          =   435
            Left            =   6090
            TabIndex        =   323
            Top             =   1275
            Width           =   2625
            VariousPropertyBits=   -1467989989
            MaxLength       =   200
            ScrollBars      =   2
            Size            =   "4630;767"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtLP12 
            Height          =   435
            Left            =   6090
            TabIndex        =   322
            Top             =   840
            Width           =   2625
            VariousPropertyBits=   -1467989989
            MaxLength       =   200
            ScrollBars      =   2
            Size            =   "4630;767"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtLP37 
            Height          =   435
            Left            =   6090
            TabIndex        =   321
            Top             =   420
            Width           =   2625
            VariousPropertyBits=   -1467989989
            MaxLength       =   200
            ScrollBars      =   2
            Size            =   "4630;767"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblLP22 
            Height          =   255
            Left            =   1140
            TabIndex        =   319
            Top             =   2280
            Width           =   1300
            BackColor       =   -2147483643
            VariousPropertyBits=   27
            Size            =   "2293;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblLP20 
            Height          =   255
            Left            =   1140
            TabIndex        =   318
            Top             =   1995
            Width           =   1300
            BackColor       =   -2147483643
            VariousPropertyBits=   27
            Size            =   "2293;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblAF14 
            Height          =   255
            Left            =   -73770
            TabIndex        =   317
            Top             =   780
            Width           =   1300
            BackColor       =   -2147483643
            VariousPropertyBits=   27
            Size            =   "2293;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblAF06 
            Height          =   255
            Left            =   -73770
            TabIndex        =   316
            Top             =   480
            Width           =   1300
            BackColor       =   -2147483643
            VariousPropertyBits=   27
            Size            =   "2293;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblLP46 
            Height          =   255
            Left            =   1140
            TabIndex        =   315
            Top             =   1710
            Width           =   1300
            BackColor       =   -2147483643
            VariousPropertyBits=   27
            Size            =   "2293;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblLP38 
            Height          =   255
            Left            =   1140
            TabIndex        =   314
            Top             =   1425
            Width           =   1300
            BackColor       =   -2147483643
            VariousPropertyBits=   27
            Size            =   "2293;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblCP154 
            Height          =   255
            Left            =   1140
            TabIndex        =   313
            Top             =   1140
            Width           =   1300
            BackColor       =   -2147483643
            VariousPropertyBits=   27
            Size            =   "2293;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblLP06 
            Height          =   255
            Left            =   1140
            TabIndex        =   312
            Top             =   855
            Width           =   1300
            BackColor       =   -2147483643
            VariousPropertyBits=   27
            Size            =   "2293;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblLP04 
            Height          =   255
            Left            =   1140
            TabIndex        =   311
            Top             =   570
            Width           =   1300
            BackColor       =   -2147483643
            VariousPropertyBits=   27
            Size            =   "2293;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "確收人員備註："
            Height          =   210
            Index           =   22
            Left            =   4800
            TabIndex        =   270
            Top             =   1290
            Width           =   1380
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblLP4748 
            AutoSize        =   -1  'True
            Caption         =   "確收時間"
            Height          =   180
            Left            =   2655
            TabIndex        =   269
            Top             =   1747
            Width           =   720
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "確收人員："
            Height          =   180
            Index           =   21
            Left            =   180
            TabIndex        =   268
            Top             =   1747
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "判發人員意見："
            Height          =   180
            Index           =   20
            Left            =   4800
            TabIndex        =   259
            Top             =   420
            Width           =   1380
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblLP3940 
            AutoSize        =   -1  'True
            Caption         =   "EMail時間"
            Height          =   180
            Left            =   2655
            TabIndex        =   258
            Top             =   1462
            Width           =   780
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "EMail人員："
            Height          =   180
            Index           =   19
            Left            =   180
            TabIndex        =   257
            Top             =   1462
            Width           =   960
         End
         Begin VB.Label lblAF1112 
            AutoSize        =   -1  'True
            Caption         =   "寄送時間"
            Height          =   180
            Left            =   -72210
            TabIndex        =   247
            Top             =   817
            Width           =   720
         End
         Begin VB.Label lblAF0708 
            AutoSize        =   -1  'True
            Caption         =   "判發時間"
            Height          =   180
            Left            =   -72210
            TabIndex        =   246
            Top             =   517
            Width           =   720
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "寄送人員："
            Height          =   180
            Index           =   18
            Left            =   -74685
            TabIndex        =   245
            Top             =   817
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "判發人員："
            Height          =   180
            Index           =   17
            Left            =   -74685
            TabIndex        =   244
            Top             =   517
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "紙本發函方式："
            Height          =   180
            Index           =   11
            Left            =   180
            TabIndex        =   243
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "判發人員："
            Height          =   180
            Index           =   12
            Left            =   180
            TabIndex        =   242
            Top             =   607
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "確認人員："
            Height          =   180
            Index           =   14
            Left            =   180
            TabIndex        =   241
            Top             =   892
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "確認人員備註："
            Height          =   180
            Index           =   16
            Left            =   4800
            TabIndex        =   240
            Top             =   840
            Width           =   1380
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblLP11 
            AutoSize        =   -1  'True
            Caption         =   "發函方式"
            Height          =   180
            Left            =   1485
            TabIndex        =   239
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblLP0517 
            AutoSize        =   -1  'True
            Caption         =   "判發時間"
            Height          =   180
            Left            =   2655
            TabIndex        =   238
            Top             =   607
            Width           =   720
         End
         Begin VB.Label lblLP0718 
            AutoSize        =   -1  'True
            Caption         =   "確認時間"
            Height          =   180
            Left            =   2655
            TabIndex        =   237
            Top             =   892
            Width           =   720
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "發文人員："
            Height          =   180
            Index           =   5
            Left            =   180
            TabIndex        =   236
            Top             =   1177
            Width           =   900
         End
         Begin VB.Label lblCP127128 
            AutoSize        =   -1  'True
            Caption         =   "發文時間"
            Height          =   180
            Left            =   2655
            TabIndex        =   235
            Top             =   1177
            Width           =   720
         End
         Begin VB.Label lblLP21 
            AutoSize        =   -1  'True
            Caption         =   "發後補看日期"
            Height          =   180
            Left            =   2655
            TabIndex        =   234
            Top             =   2032
            Width           =   1080
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "發後補看："
            Height          =   180
            Index           =   3
            Left            =   180
            TabIndex        =   233
            Top             =   2032
            Width           =   900
         End
         Begin VB.Label lblLP23 
            AutoSize        =   -1  'True
            Caption         =   "繪圖補看日期"
            Height          =   180
            Left            =   2655
            TabIndex        =   232
            Top             =   2317
            Width           =   1080
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "繪圖補看："
            Height          =   180
            Index           =   10
            Left            =   180
            TabIndex        =   231
            Top             =   2317
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "發後補看備註："
            Height          =   180
            Index           =   13
            Left            =   4800
            TabIndex        =   230
            Top             =   1710
            Width           =   1380
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "繪圖補看備註："
            Height          =   180
            Index           =   15
            Left            =   4800
            TabIndex        =   229
            Top             =   2130
            Width           =   1380
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblLP26 
            AutoSize        =   -1  'True
            Caption         =   "是否E化"
            Height          =   180
            Left            =   2655
            TabIndex        =   228
            Top             =   360
            Width           =   645
         End
      End
      Begin VB.TextBox textCP148 
         Height          =   285
         Left            =   -66960
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   223
         Top             =   927
         Width           =   255
      End
      Begin VB.TextBox textCP113 
         Height          =   285
         Left            =   -66945
         TabIndex        =   220
         Top             =   1713
         Width           =   600
      End
      Begin VB.TextBox textCP114 
         Height          =   285
         Left            =   -66945
         MaxLength       =   4
         TabIndex        =   219
         Top             =   1998
         Width           =   600
      End
      Begin VB.TextBox textCP140 
         Height          =   285
         Left            =   -67560
         Locked          =   -1  'True
         TabIndex        =   211
         Top             =   1539
         Width           =   1185
      End
      Begin VB.TextBox textCP138 
         Height          =   285
         Left            =   -66555
         TabIndex        =   40
         Top             =   3690
         Width           =   375
      End
      Begin VB.TextBox textCP137 
         Height          =   285
         Left            =   -66555
         TabIndex        =   38
         Top             =   3375
         Width           =   375
      End
      Begin VB.TextBox textCP135 
         Height          =   285
         Left            =   -67995
         TabIndex        =   35
         Top             =   3069
         Width           =   420
      End
      Begin VB.TextBox textCP136 
         Height          =   285
         Left            =   -66555
         TabIndex        =   36
         Top             =   3069
         Width           =   375
      End
      Begin VB.TextBox textCP132 
         Height          =   264
         Left            =   5895
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   205
         Top             =   1230
         Width           =   825
      End
      Begin VB.TextBox textCP130 
         Height          =   264
         Left            =   2925
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   202
         Top             =   1515
         Width           =   5685
      End
      Begin VB.TextBox textCP123 
         Height          =   264
         Left            =   2925
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   196
         Top             =   345
         Width           =   255
      End
      Begin VB.TextBox textCP124 
         Height          =   264
         Left            =   2925
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   195
         Top             =   645
         Width           =   825
      End
      Begin VB.TextBox textCP125 
         BorderStyle     =   0  '沒有框線
         Height          =   195
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   194
         TabStop         =   0   'False
         Top             =   645
         Width           =   780
      End
      Begin VB.TextBox textCP126 
         Height          =   264
         Left            =   2925
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   193
         Top             =   930
         Width           =   255
      End
      Begin VB.TextBox textCP129 
         Height          =   264
         Left            =   2925
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   192
         Top             =   1230
         Width           =   825
      End
      Begin VB.TextBox textCP120 
         Height          =   285
         Left            =   -69885
         MaxLength       =   1
         TabIndex        =   33
         Top             =   2763
         Width           =   255
      End
      Begin VB.TextBox textCP121 
         Height          =   285
         Left            =   -67155
         MaxLength       =   1
         TabIndex        =   34
         Top             =   2763
         Width           =   255
      End
      Begin VB.TextBox TextCP119 
         Height          =   285
         Left            =   -71900
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   189
         Top             =   573
         Width           =   735
      End
      Begin VB.TextBox textCP118 
         Height          =   285
         Left            =   -68080
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   187
         Top             =   1143
         Width           =   270
      End
      Begin VB.TextBox textCP44 
         Height          =   285
         Left            =   -73920
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   1998
         Width           =   1095
      End
      Begin VB.TextBox textCP15 
         Height          =   285
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   76
         Top             =   1428
         Width           =   1092
      End
      Begin VB.TextBox textCP29 
         Height          =   285
         Left            =   -70170
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   77
         Top             =   1428
         Width           =   1092
      End
      Begin VB.TextBox textCP84 
         Height          =   285
         Left            =   -69150
         Locked          =   -1  'True
         TabIndex        =   168
         Top             =   1713
         Width           =   1170
      End
      Begin VB.TextBox textCP28 
         Height          =   285
         Left            =   -71670
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   167
         Top             =   1713
         Width           =   1515
      End
      Begin VB.TextBox textCP10_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   250
         Left            =   -69030
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   305
         Width           =   2052
      End
      Begin VB.TextBox textCP12_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   250
         Left            =   -69660
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   858
         Width           =   1260
      End
      Begin VB.TextBox textCP58_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   250
         Left            =   -70740
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   3810
         Width           =   2550
      End
      Begin VB.TextBox textCP11_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -73260
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1855
         Width           =   1530
      End
      Begin VB.TextBox textCP12 
         Height          =   285
         Left            =   -70170
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   84
         Top             =   858
         Width           =   495
      End
      Begin VB.TextBox textCP13 
         Height          =   285
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   83
         Top             =   858
         Width           =   1092
      End
      Begin VB.TextBox textCP14 
         Height          =   285
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   82
         Top             =   1143
         Width           =   1092
      End
      Begin VB.TextBox textCP27 
         Height          =   285
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   81
         Top             =   1713
         Width           =   1095
      End
      Begin VB.TextBox textCP06 
         Height          =   285
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   80
         Top             =   573
         Width           =   1092
      End
      Begin VB.TextBox textCP43 
         Height          =   270
         Left            =   -69600
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   79
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox textCP08 
         Height          =   285
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   78
         Top             =   3180
         Width           =   4905
      End
      Begin VB.TextBox textCP21 
         Height          =   285
         Left            =   -68790
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   75
         Top             =   2865
         Width           =   255
      End
      Begin VB.TextBox textCP48 
         Height          =   285
         Left            =   -70170
         MaxLength       =   7
         TabIndex        =   74
         Top             =   1143
         Width           =   1095
      End
      Begin VB.TextBox textCP31 
         Height          =   285
         Left            =   -71400
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   72
         Top             =   2865
         Width           =   255
      End
      Begin VB.TextBox textCP57 
         Height          =   270
         Left            =   -73635
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   71
         Top             =   3810
         Width           =   1095
      End
      Begin VB.TextBox textCP58 
         Height          =   270
         Left            =   -71085
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   70
         Top             =   3810
         Width           =   330
      End
      Begin VB.TextBox textCP22 
         Height          =   270
         Left            =   -67410
         MaxLength       =   1
         TabIndex        =   69
         Top             =   3520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox textCP45 
         Height          =   285
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   68
         Top             =   2283
         Width           =   2895
      End
      Begin VB.TextBox textCP25 
         Height          =   285
         Left            =   -73275
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   64
         Top             =   2865
         Width           =   1095
      End
      Begin VB.TextBox textCP07 
         Height          =   285
         Left            =   -70170
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   63
         Top             =   573
         Width           =   1095
      End
      Begin VB.TextBox textCP05 
         Height          =   285
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   62
         Top             =   288
         Width           =   1095
      End
      Begin VB.TextBox textCP10 
         Height          =   285
         Left            =   -70170
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   61
         Top             =   288
         Width           =   1092
      End
      Begin VB.TextBox textCP17 
         Height          =   270
         Left            =   -69960
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   60
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox textCP19 
         Height          =   285
         Left            =   -69960
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   59
         Top             =   621
         Width           =   1095
      End
      Begin VB.TextBox textCP33 
         Height          =   285
         Left            =   -73950
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   58
         Top             =   927
         Width           =   1095
      End
      Begin VB.TextBox textCP34 
         Height          =   285
         Left            =   -69960
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   57
         Top             =   927
         Width           =   1095
      End
      Begin VB.TextBox textCP46 
         Height          =   285
         Left            =   -72660
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   55
         Top             =   1233
         Width           =   900
      End
      Begin VB.TextBox textCP47 
         Height          =   285
         Left            =   -69960
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   54
         Top             =   1233
         Width           =   1095
      End
      Begin VB.TextBox textCP32 
         Height          =   285
         Left            =   -73470
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   53
         Top             =   1539
         Width           =   375
      End
      Begin VB.TextBox textCP20 
         Height          =   285
         Left            =   -69960
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   52
         Top             =   1539
         Width           =   375
      End
      Begin VB.TextBox textCP11 
         Height          =   285
         Left            =   -73650
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   51
         Top             =   1845
         Width           =   375
      End
      Begin VB.TextBox textCP59 
         Height          =   285
         Left            =   -73980
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   2151
         Width           =   1935
      End
      Begin VB.TextBox textCP30 
         Height          =   285
         Left            =   -71010
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   48
         Top             =   2457
         Width           =   4395
      End
      Begin VB.TextBox textCP60 
         Height          =   285
         Left            =   -73260
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   47
         Top             =   2763
         Width           =   1860
      End
      Begin VB.TextBox textCP61 
         Height          =   285
         Left            =   -73755
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   46
         Top             =   3069
         Width           =   1770
      End
      Begin VB.TextBox textCP62 
         Height          =   285
         Left            =   -73755
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   45
         Top             =   3375
         Width           =   1770
      End
      Begin VB.TextBox textCP63 
         Height          =   285
         Left            =   -73755
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   32
         Top             =   3690
         Width           =   1770
      End
      Begin VB.TextBox textCP18 
         Height          =   285
         Left            =   -73950
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   31
         Top             =   621
         Width           =   1095
      End
      Begin VB.TextBox textCP16 
         Height          =   270
         Left            =   -73950
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   30
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox textCP54 
         Height          =   285
         Left            =   -67890
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   29
         Top             =   3285
         Width           =   945
      End
      Begin VB.TextBox textCP53 
         Height          =   285
         Left            =   -69180
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   28
         Top             =   3285
         Width           =   945
      End
      Begin VB.TextBox textCP81 
         Height          =   285
         Left            =   -67710
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   56
         Top             =   1233
         Width           =   255
      End
      Begin VB.TextBox textCP80 
         Height          =   270
         Left            =   -73230
         Locked          =   -1  'True
         MaxLength       =   39
         TabIndex        =   27
         Top             =   2640
         Width           =   6855
      End
      Begin VB.TextBox textCP36 
         Height          =   270
         Left            =   -73230
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   26
         Top             =   295
         Width           =   2775
      End
      Begin VB.TextBox textCP82 
         Appearance      =   0  '平面
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   -66840
         MaxLength       =   6
         TabIndex        =   25
         Top             =   573
         Width           =   540
      End
      Begin VB.TextBox textCP83 
         Appearance      =   0  '平面
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   -67920
         TabIndex        =   24
         Top             =   573
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox textCP88 
         Height          =   285
         Left            =   -70815
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   23
         Top             =   3375
         Width           =   1770
      End
      Begin VB.TextBox textCP87 
         Height          =   285
         Left            =   -70815
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   22
         Top             =   3069
         Width           =   1770
      End
      Begin VB.TextBox textCP72 
         Height          =   285
         Left            =   -73980
         Locked          =   -1  'True
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   21
         Top             =   3285
         Width           =   1212
      End
      Begin VB.TextBox textCP56 
         Height          =   285
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   20
         Top             =   1795
         Width           =   1212
      End
      Begin VB.TextBox textCP55 
         Height          =   285
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   19
         Top             =   325
         Width           =   1212
      End
      Begin VB.TextBox textCP93 
         Height          =   285
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   18
         Top             =   625
         Width           =   1212
      End
      Begin VB.TextBox textCP94 
         Height          =   285
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   17
         Top             =   910
         Width           =   1212
      End
      Begin VB.TextBox textCP95 
         Height          =   285
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   16
         Top             =   1210
         Width           =   1212
      End
      Begin VB.TextBox textCP96 
         Height          =   285
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   15
         Top             =   1495
         Width           =   1212
      End
      Begin VB.TextBox textCP89 
         Height          =   285
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   14
         Top             =   2080
         Width           =   1212
      End
      Begin VB.TextBox textCP90 
         Height          =   285
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   13
         Top             =   2385
         Width           =   1212
      End
      Begin VB.TextBox textCP91 
         Height          =   285
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   12
         Top             =   2665
         Width           =   1212
      End
      Begin VB.TextBox textCP92 
         Height          =   285
         Left            =   -73110
         Locked          =   -1  'True
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   11
         Top             =   2965
         Width           =   1212
      End
      Begin VB.TextBox textCP24 
         Height          =   285
         Left            =   -73920
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   65
         Top             =   2565
         Width           =   255
      End
      Begin VB.TextBox textCP23 
         Height          =   270
         Left            =   -70870
         MaxLength       =   1
         TabIndex        =   66
         Top             =   2565
         Width           =   255
      End
      Begin VB.TextBox textCP26 
         Height          =   270
         Left            =   -67350
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   73
         Top             =   2565
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70440
         TabIndex        =   213
         Top             =   1740
         Width           =   4275
         Begin VB.CheckBox chkCP176 
            Caption         =   "暫不送"
            Height          =   250
            Left            =   3390
            TabIndex        =   338
            Top             =   95
            Width           =   850
         End
         Begin VB.TextBox textCP142 
            Height          =   270
            Left            =   2580
            MaxLength       =   7
            TabIndex        =   217
            Top             =   120
            Width           =   780
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "指定日期"
            Height          =   180
            Index           =   3
            Left            =   1596
            TabIndex        =   216
            Top             =   135
            Width           =   1035
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "收款後"
            Height          =   180
            Index           =   2
            Left            =   792
            TabIndex        =   215
            Top             =   135
            Width           =   850
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "不限制"
            Height          =   180
            Index           =   1
            Left            =   24
            TabIndex        =   214
            Top             =   135
            Width           =   850
         End
      End
      Begin VB.Label lblCP168 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪已審頁："
         Height          =   180
         Left            =   -68880
         TabIndex        =   335
         Top             =   3720
         Width           =   900
      End
      Begin VB.Label lblCP167 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪未審頁："
         Height          =   180
         Left            =   -68880
         TabIndex        =   334
         Top             =   3405
         Width           =   900
      End
      Begin VB.Label lblCP86 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "收到分所接洽單紀錄："
         Height          =   180
         Left            =   -74880
         TabIndex        =   333
         Top             =   4005
         Width           =   1800
      End
      Begin VB.Label lblCP86_1 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Left            =   -72675
         TabIndex        =   332
         Top             =   4005
         Width           =   465
      End
      Begin MSForms.Label lblCP153 
         Height          =   255
         Left            =   7050
         TabIndex        =   320
         Top             =   645
         Width           =   1305
         BackColor       =   -2147483643
         VariousPropertyBits=   27
         Size            =   "2302;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP131 
         Height          =   300
         Left            =   2925
         TabIndex        =   309
         Top             =   1800
         Width           =   5685
         VariousPropertyBits=   -1467989985
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "10028;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37_1 
         Height          =   975
         Left            =   -73230
         TabIndex        =   302
         Top             =   600
         Width           =   6855
         VariousPropertyBits=   -1467989985
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "12091;1720"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37 
         Height          =   300
         Left            =   -73230
         TabIndex        =   308
         Top             =   600
         Width           =   6855
         VariousPropertyBits=   -1467989985
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "12091;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP38 
         Height          =   300
         Left            =   -73230
         TabIndex        =   307
         Top             =   945
         Width           =   6855
         VariousPropertyBits=   -1467989985
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "12091;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP39 
         Height          =   300
         Left            =   -73230
         TabIndex        =   306
         Top             =   1290
         Width           =   6855
         VariousPropertyBits=   -1467989985
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "12091;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73230
         TabIndex        =   305
         Top             =   1620
         Width           =   6855
         VariousPropertyBits=   -1467989985
         MaxLength       =   60
         ScrollBars      =   2
         Size            =   "12091;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP41 
         Height          =   300
         Left            =   -73230
         TabIndex        =   304
         Top             =   1965
         Width           =   6855
         VariousPropertyBits=   -1467989985
         ScrollBars      =   2
         Size            =   "12091;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73230
         TabIndex        =   303
         Top             =   2310
         Width           =   6855
         VariousPropertyBits=   -1467989985
         MaxLength       =   60
         ScrollBars      =   2
         Size            =   "12091;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "報價備註："
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   301
         Top             =   3030
         Width           =   900
      End
      Begin MSForms.TextBox textCP144 
         Height          =   570
         Left            =   -73890
         TabIndex        =   300
         Top             =   2970
         Width           =   7665
         VariousPropertyBits=   -1467989985
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "13520;1005"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP55_2 
         Height          =   285
         Left            =   -71840
         TabIndex        =   299
         TabStop         =   0   'False
         Top             =   325
         Width           =   5660
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "9984;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP93_2 
         Height          =   285
         Left            =   -71840
         TabIndex        =   298
         TabStop         =   0   'False
         Top             =   630
         Width           =   5660
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "9984;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP94_2 
         Height          =   285
         Left            =   -71840
         TabIndex        =   297
         TabStop         =   0   'False
         Top             =   915
         Width           =   5660
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "9984;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP95_2 
         Height          =   285
         Left            =   -71840
         TabIndex        =   296
         TabStop         =   0   'False
         Top             =   1215
         Width           =   5660
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "9984;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP96_2 
         Height          =   285
         Left            =   -71840
         TabIndex        =   295
         TabStop         =   0   'False
         Top             =   1500
         Width           =   5660
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "9984;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP56_2 
         Height          =   285
         Left            =   -71840
         TabIndex        =   294
         TabStop         =   0   'False
         Top             =   1800
         Width           =   5660
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "9984;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP89_2 
         Height          =   285
         Left            =   -71840
         TabIndex        =   293
         TabStop         =   0   'False
         Top             =   2085
         Width           =   5660
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "9984;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP90_2 
         Height          =   285
         Left            =   -71840
         TabIndex        =   292
         TabStop         =   0   'False
         Top             =   2385
         Width           =   5660
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "9984;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP91_2 
         Height          =   285
         Left            =   -71840
         TabIndex        =   291
         TabStop         =   0   'False
         Top             =   2670
         Width           =   5660
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "9984;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP92_2 
         Height          =   285
         Left            =   -71840
         TabIndex        =   290
         TabStop         =   0   'False
         Top             =   2970
         Width           =   5660
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "9984;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP50 
         Height          =   285
         Left            =   -73080
         TabIndex        =   289
         Top             =   3630
         Width           =   6860
         VariousPropertyBits=   -1467989985
         MaxLength       =   60
         ScrollBars      =   2
         Size            =   "12100;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP51 
         Height          =   285
         Left            =   -73080
         TabIndex        =   288
         Top             =   3930
         Width           =   6860
         VariousPropertyBits=   -1467989985
         MaxLength       =   60
         ScrollBars      =   2
         Size            =   "12100;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP52 
         Height          =   285
         Left            =   -73080
         TabIndex        =   287
         Top             =   4230
         Width           =   6860
         VariousPropertyBits=   -1467989985
         MaxLength       =   60
         ScrollBars      =   2
         Size            =   "12100;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP49 
         Height          =   525
         Left            =   -73260
         TabIndex        =   286
         Top             =   4080
         Width           =   6930
         VariousPropertyBits=   -1467989985
         MaxLength       =   249
         ScrollBars      =   2
         Size            =   "12224;926"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "審查委員/法院案號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   282
         Top             =   3645
         Width           =   1665
      End
      Begin VB.Label Label51 
         Caption         =   "審查委員編號："
         Height          =   180
         Left            =   -69900
         TabIndex        =   281
         Top             =   3652
         Width           =   1260
      End
      Begin MSForms.Label lblNameAgent 
         Height          =   285
         Left            =   -74880
         TabIndex        =   277
         Top             =   3510
         Width           =   6495
         BackColor       =   -2147483643
         VariousPropertyBits=   27
         Caption         =   "出名代理人"
         Size            =   "11456;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   285
         Left            =   -72780
         TabIndex        =   276
         TabStop         =   0   'False
         Top             =   1998
         Width           =   4815
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "8493;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP83_2 
         Height          =   285
         Left            =   -67710
         TabIndex        =   275
         TabStop         =   0   'False
         Top             =   573
         Width           =   855
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "1508;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP29_2 
         Height          =   285
         Left            =   -69030
         TabIndex        =   274
         TabStop         =   0   'False
         Top             =   1428
         Width           =   855
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "1508;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   285
         Left            =   -72780
         TabIndex        =   273
         TabStop         =   0   'False
         Top             =   1143
         Width           =   1635
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "2884;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP13_2 
         Height          =   285
         Left            =   -72780
         TabIndex        =   272
         TabStop         =   0   'False
         Top             =   858
         Width           =   1635
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "2884;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   615
         Left            =   -73980
         TabIndex        =   271
         Top             =   4140
         Width           =   7725
         VariousPropertyBits=   -1466941409
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13626;1085"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCP71 
         AutoSize        =   -1  'True
         Caption         =   "機關代號："
         Height          =   180
         Left            =   -71205
         TabIndex        =   137
         Top             =   2203
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "扣款日期："
         Height          =   180
         Index           =   1
         Left            =   -68040
         TabIndex        =   254
         Top             =   1428
         Width           =   900
      End
      Begin VB.Label Label6 
         Caption         =   " ( Y:是, W:待確認,    A:自動扣款 )"
         Height          =   390
         Index           =   4
         Left            =   -67710
         TabIndex        =   253
         Top             =   1020
         Width           =   1380
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "第　　　　期登記期"
         Height          =   180
         Index           =   7
         Left            =   -69510
         TabIndex        =   248
         Top             =   3330
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label29 
         Caption         =   "(Y:是)"
         Height          =   180
         Left            =   -66660
         TabIndex        =   225
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label30 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "電子表單單號："
         Height          =   180
         Index           =   1
         Left            =   -68790
         TabIndex        =   212
         Top             =   1591
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(N:不收)"
         Height          =   180
         Index           =   2
         Left            =   -69510
         TabIndex        =   120
         Top             =   1591
         Width           =   645
      End
      Begin VB.Label Label22 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "是否為一申請書多件："
         Height          =   180
         Left            =   -68730
         TabIndex        =   224
         Top             =   979
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作時數："
         Height          =   180
         Index           =   12
         Left            =   -67845
         TabIndex        =   222
         Top             =   1713
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "核稿時數："
         Height          =   180
         Index           =   5
         Left            =   -67845
         TabIndex        =   221
         Top             =   1998
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "送件方式："
         Height          =   180
         Index           =   121
         Left            =   -71355
         TabIndex        =   218
         Top             =   1897
         Width           =   900
      End
      Begin VB.Label lblCP137 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪未審項："
         Height          =   180
         Left            =   -67425
         TabIndex        =   210
         Top             =   3420
         Width           =   900
      End
      Begin VB.Label lblCP136 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "增加項數："
         Height          =   180
         Left            =   -67425
         TabIndex        =   209
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label lblCP135 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "增加頁數："
         Height          =   180
         Left            =   -68880
         TabIndex        =   208
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label lblCP138 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪已審項："
         Height          =   180
         Left            =   -67425
         TabIndex        =   207
         Top             =   3735
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "取消發文日："
         Height          =   180
         Index           =   9
         Left            =   4770
         TabIndex        =   206
         Top             =   1230
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "發文室取消發文備註："
         Height          =   180
         Index           =   8
         Left            =   270
         TabIndex        =   204
         Top             =   1800
         Width           =   1800
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "主管機關名稱："
         Height          =   180
         Index           =   7
         Left            =   270
         TabIndex        =   203
         Top             =   1515
         Width           =   1260
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否算發文室件數-主管機關：            ( Y: 是  N:否 空白:未經發文室 )"
         Height          =   180
         Index           =   6
         Left            =   270
         TabIndex        =   201
         Top             =   345
         Width           =   5325
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "發文室發文日-主管機關："
         Height          =   180
         Index           =   2
         Left            =   270
         TabIndex        =   200
         Top             =   645
         Width           =   2040
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "發文室發文人員-主管機關："
         Height          =   180
         Index           =   4
         Left            =   4815
         TabIndex        =   199
         Top             =   645
         Width           =   2220
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否算發文室件數-非主管機關：        ( Y: 是  N:否 空白:未經發文室 )"
         Height          =   180
         Index           =   7
         Left            =   270
         TabIndex        =   198
         Top             =   930
         Width           =   5325
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "分所發文日："
         Height          =   180
         Index           =   6
         Left            =   270
         TabIndex        =   197
         Top             =   1230
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "說明書要電子檔：      (Y:是)"
         Height          =   180
         Index           =   5
         Left            =   -71325
         TabIndex        =   191
         Top             =   2815
         Width           =   2175
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "說明書電子檔已上傳：      (Y:是)"
         Height          =   180
         Index           =   6
         Left            =   -68955
         TabIndex        =   190
         Top             =   2815
         Width           =   2535
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "來函櫃台收文日："
         Height          =   180
         Left            =   -72600
         TabIndex        =   188
         Top             =   288
         Width           =   1440
      End
      Begin VB.Label Label6 
         Caption         =   "電子送件："
         Height          =   210
         Index           =   2
         Left            =   -69000
         TabIndex        =   186
         Top             =   1143
         Width           =   945
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "代  理  人："
         Height          =   180
         Left            =   -74880
         TabIndex        =   143
         Top             =   1998
         Width           =   900
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "繪圖人員/協辦人員："
         Height          =   180
         Left            =   -71880
         TabIndex        =   166
         Top             =   1428
         Width           =   1665
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "支援時數 ："
         Height          =   180
         Left            =   -74880
         TabIndex        =   150
         Top             =   1428
         Width           =   945
      End
      Begin VB.Label Label40 
         Alignment       =   1  '靠右對齊
         Caption         =   "已收服務費："
         Height          =   180
         Left            =   -74820
         TabIndex        =   185
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label Label27 
         Alignment       =   1  '靠右對齊
         Caption         =   "已收規費："
         Height          =   180
         Left            =   -74820
         TabIndex        =   184
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label44 
         Alignment       =   1  '靠右對齊
         Caption         =   "已收金額："
         Height          =   180
         Left            =   -74820
         TabIndex        =   183
         Top             =   1005
         Width           =   1080
      End
      Begin VB.Label Label45 
         Alignment       =   1  '靠右對齊
         Caption         =   "已扣繳金額："
         Height          =   180
         Left            =   -74820
         TabIndex        =   182
         Top             =   1275
         Width           =   1080
      End
      Begin VB.Label Label46 
         Alignment       =   1  '靠右對齊
         Caption         =   "已銷帳金額："
         Height          =   180
         Left            =   -74820
         TabIndex        =   181
         Top             =   1545
         Width           =   1080
      End
      Begin VB.Label Label48 
         Alignment       =   1  '靠右對齊
         Caption         =   "已退費金額："
         Height          =   180
         Left            =   -74820
         TabIndex        =   180
         Top             =   1830
         Width           =   1080
      End
      Begin VB.Label Label49 
         Alignment       =   1  '靠右對齊
         Caption         =   "未收金額："
         Height          =   180
         Left            =   -74820
         TabIndex        =   179
         Top             =   2100
         Width           =   1080
      End
      Begin VB.Label lblCP73 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -73740
         TabIndex        =   178
         Top             =   450
         Width           =   45
      End
      Begin VB.Label lblCP74 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -73740
         TabIndex        =   177
         Top             =   720
         Width           =   45
      End
      Begin VB.Label lblCP75 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -73740
         TabIndex        =   176
         Top             =   1005
         Width           =   45
      End
      Begin VB.Label lblCP76 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -73740
         TabIndex        =   175
         Top             =   1275
         Width           =   45
      End
      Begin VB.Label lblCP77 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -73740
         TabIndex        =   174
         Top             =   1545
         Width           =   45
      End
      Begin VB.Label lblCP78 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -73740
         TabIndex        =   173
         Top             =   1830
         Width           =   45
      End
      Begin VB.Label lblCP79 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -73740
         TabIndex        =   172
         Top             =   2100
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Index           =   5
         Left            =   -70050
         TabIndex        =   170
         Top             =   1713
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "發文字號："
         Height          =   180
         Left            =   -72630
         TabIndex        =   169
         Top             =   1713
         Width           =   900
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "業  務  區："
         Height          =   180
         Left            =   -71115
         TabIndex        =   165
         Top             =   910
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   164
         Top             =   858
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "承  辦  人："
         Height          =   180
         Index           =   5
         Left            =   -74880
         TabIndex        =   163
         Top             =   1143
         Width           =   900
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "進度備註："
         Height          =   180
         Left            =   -74880
         TabIndex        =   162
         Top             =   4140
         Width           =   900
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "相關總收文號："
         Height          =   180
         Left            =   -70920
         TabIndex        =   161
         Top             =   2310
         Width           =   1260
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "發  文  日："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   160
         Top             =   1713
         Width           =   900
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "取消收文日期："
         Height          =   180
         Left            =   -74880
         TabIndex        =   159
         Top             =   3810
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限："
         Height          =   180
         Index           =   0
         Left            =   -71115
         TabIndex        =   158
         Top             =   573
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限："
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   157
         Top             =   573
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "機關文號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   156
         Top             =   3180
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   180
         Index           =   0
         Left            =   -71115
         TabIndex        =   155
         Top             =   288
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "收  文  日："
         Height          =   180
         Left            =   -74880
         TabIndex        =   154
         Top             =   288
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "是否多國/取締案："
         Height          =   180
         Left            =   -70350
         TabIndex        =   153
         Top             =   2865
         Width           =   1485
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "是否出名：     (N:不出名)"
         Height          =   180
         Index           =   0
         Left            =   -68265
         TabIndex        =   152
         Top             =   3570
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Left            =   -68490
         TabIndex        =   149
         Top             =   2865
         Width           =   465
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Index           =   0
         Left            =   -71040
         TabIndex        =   148
         Top             =   2865
         Width           =   465
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "取消收文原因："
         Height          =   180
         Left            =   -72315
         TabIndex        =   147
         Top             =   3810
         Width           =   1260
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "新案："
         Height          =   180
         Left            =   -72000
         TabIndex        =   146
         Top             =   2865
         Width           =   540
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "承辦期限："
         Height          =   180
         Left            =   -71115
         TabIndex        =   145
         Top             =   1143
         Width           =   900
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   144
         Top             =   2310
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "准駁/勝敗/判決日："
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   140
         Top             =   2865
         Width           =   1530
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "點        數："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -74880
         TabIndex        =   138
         Top             =   673
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "案件來源代號："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   136
         Top             =   1897
         Width           =   1260
      End
      Begin VB.Label lblCP49 
         AutoSize        =   -1  'True
         Caption         =   "條款/當事人稱謂 ："
         Height          =   180
         Left            =   -74790
         TabIndex        =   135
         Top             =   4080
         Width           =   1530
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "大陸申請案號/延展註冊號數/股別/下一程序序號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   134
         Top             =   2505
         Width           =   3915
      End
      Begin VB.Label Label30 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "代理人提申日："
         Height          =   180
         Index           =   3
         Left            =   -71205
         TabIndex        =   133
         Top             =   1285
         Width           =   1260
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "代理人收達日/回執收受日："
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   132
         Top             =   1285
         Width           =   2205
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "結餘註記："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   131
         Top             =   2203
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "標  準  價："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   130
         Top             =   979
         Width           =   900
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "底        價："
         Height          =   180
         Index           =   2
         Left            =   -70845
         TabIndex        =   129
         Top             =   979
         Width           =   900
      End
      Begin VB.Label lblCP19 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "後        金："
         Height          =   180
         Left            =   -70845
         TabIndex        =   128
         Top             =   673
         Width           =   900
      End
      Begin VB.Label Label37 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "規        費："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -70845
         TabIndex        =   127
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "費        用："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -74880
         TabIndex        =   126
         Top             =   375
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "收據編號/請款單號："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -74910
         TabIndex        =   125
         Top             =   2815
         Width           =   1665
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CF帳單編號1："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   124
         Top             =   3121
         Width           =   1200
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CF帳單編號2："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   2
         Left            =   -74910
         TabIndex        =   123
         Top             =   3427
         Width           =   1200
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CF帳單編號3："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   3
         Left            =   -74910
         TabIndex        =   122
         Top             =   3742
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "是否向客戶收款："
         Height          =   180
         Index           =   3
         Left            =   -71385
         TabIndex        =   121
         Top             =   1591
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "是否開電腦收據：         (N:不開)"
         Height          =   180
         Index           =   12
         Left            =   -74880
         TabIndex        =   119
         Top             =   1591
         Width           =   2490
      End
      Begin VB.Line Line1 
         X1              =   -68010
         X2              =   -68160
         Y1              =   3390
         Y2              =   3390
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "授權期間/質權設定期間(起/迄)/聘任期間："
         Height          =   180
         Index           =   2
         Left            =   -72540
         TabIndex        =   118
         Top             =   3330
         Width           =   3315
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "被授權人(中)/收件人："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   117
         Top             =   3682
         Width           =   1785
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "被授權人："
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   116
         Top             =   3330
         Width           =   900
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "被授權人(英)："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   115
         Top             =   3982
         Width           =   1200
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "被授權人(日)："
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   114
         Top             =   4275
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "讓與人/移轉人1："
         Height          =   180
         Index           =   7
         Left            =   -74880
         TabIndex        =   113
         Top             =   367
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "受讓人/移轉申請人1："
         Height          =   180
         Index           =   8
         Left            =   -74880
         TabIndex        =   112
         Top             =   1837
         Width           =   1755
      End
      Begin VB.Label Label43 
         Caption         =   "若有收據或請款編號或帳單編號 ， 只能由電腦中心人員修改!!!"
         ForeColor       =   &H000000FF&
         Height          =   570
         Left            =   -68700
         TabIndex        =   111
         Top             =   360
         Width           =   2070
      End
      Begin VB.Label lblCP81 
         AutoSize        =   -1  'True
         Caption         =   "是否可減免：      (Y/N)"
         Height          =   180
         Left            =   -68730
         TabIndex        =   110
         Top             =   1285
         Width           =   1755
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱："
         Height          =   180
         Index           =   5
         Left            =   -74880
         TabIndex        =   109
         Top             =   623
         Width           =   1260
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "對造案件商品類別："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   108
         Top             =   2682
         Width           =   1620
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "對造號數："
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   107
         Top             =   337
         Width           =   900
      End
      Begin VB.Label Label18 
         Caption         =   "對造案件名稱(中)："
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   106
         Top             =   623
         Width           =   1575
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱(中)："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   105
         Top             =   1617
         Width           =   1200
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱(英)："
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   104
         Top             =   909
         Width           =   1560
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱(日)："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   103
         Top             =   1263
         Width           =   1560
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱(英)："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   102
         Top             =   1971
         Width           =   1200
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱(日)："
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   101
         Top             =   2325
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "發文人員時間："
         Height          =   180
         Index           =   3
         Left            =   -69000
         TabIndex        =   100
         Top             =   573
         Width           =   1260
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CF帳單編號5："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   4
         Left            =   -71985
         TabIndex        =   99
         Top             =   3420
         Width           =   1200
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CF帳單編號4："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   5
         Left            =   -71985
         TabIndex        =   98
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "讓與人/移轉人2："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   97
         Top             =   667
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "讓與人/移轉人3："
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   96
         Top             =   952
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "讓與人/移轉人4："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   95
         Top             =   1252
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "讓與人/移轉人5："
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   94
         Top             =   1537
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "受讓人/移轉申請人2："
         Height          =   180
         Index           =   6
         Left            =   -74880
         TabIndex        =   93
         Top             =   2122
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "受讓人/移轉申請人3："
         Height          =   180
         Index           =   9
         Left            =   -74880
         TabIndex        =   92
         Top             =   2385
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "受讓人/移轉申請人4："
         Height          =   180
         Index           =   10
         Left            =   -74880
         TabIndex        =   91
         Top             =   2707
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "受讓人/移轉申請人5："
         Height          =   180
         Index           =   11
         Left            =   -74880
         TabIndex        =   90
         Top             =   3007
         Width           =   1755
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否算案件數：      (N:不算)"
         Height          =   180
         Index           =   1
         Left            =   -68625
         TabIndex        =   151
         Top             =   2565
         Width           =   2175
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "預估結果：      (1:准/勝,2:駁/敗,3:部分勝敗)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   160
         Index           =   0
         Left            =   -71640
         TabIndex        =   142
         Top             =   2570
         Width           =   2940
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "實際結果："
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   141
         Top             =   2565
         Width           =   900
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "(1:准/勝,2:駁/敗,3:部分勝敗)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   160
         Left            =   -73650
         TabIndex        =   139
         Top             =   2570
         Width           =   1900
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8400
      Top             =   300
      _ExtentX        =   494
      _ExtentY        =   494
      _Version        =   393216
   End
   Begin VB.Label LblCP150Y 
      Caption         =   "有特例簽核"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   7230
      TabIndex        =   330
      Top             =   300
      Visible         =   0   'False
      Width           =   1005
   End
   Begin MSForms.TextBox textCUID 
      Height          =   285
      Left            =   60
      TabIndex        =   310
      TabStop         =   0   'False
      Top             =   6060
      Width           =   8865
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "15637;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   1080
      TabIndex        =   285
      Top             =   780
      Width           =   7635
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13462;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   2430
      TabIndex        =   284
      Top             =   510
      Width           =   6225
      BackColor       =   -2147483643
      VariousPropertyBits=   27
      Size            =   "10980;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1380
      TabIndex        =   283
      Top             =   510
      Width           =   975
      BackColor       =   -2147483643
      VariousPropertyBits=   27
      Size            =   "1720;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCancel 
      Caption         =   "lblCancel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   6540
      TabIndex        =   171
      Top             =   300
      Width           =   645
   End
   Begin VB.Label lbeNumber 
      Height          =   180
      Left            =   3240
      TabIndex        =   9
      Top             =   300
      Width           =   2175
   End
   Begin VB.Label lbl01020304 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   2310
      TabIndex        =   8
      Top             =   300
      Width           =   900
   End
   Begin VB.Label lblCP09 
      AutoSize        =   -1  'True
      Caption         =   "收文號： "
      Height          =   180
      Left            =   105
      TabIndex        =   7
      Top             =   300
      Width           =   765
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5520
      TabIndex        =   6
      Top             =   300
      Width           =   975
   End
   Begin VB.Label lblAll 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   105
      TabIndex        =   5
      Top             =   832
      Width           =   900
   End
   Begin VB.Label lblCU 
      AutoSize        =   -1  'True
      Caption         =   "申請人/當事人："
      Height          =   180
      Left            =   105
      TabIndex        =   4
      Top             =   547
      Width           =   1305
   End
End
Attribute VB_Name = "frm100101_C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modify By Sindy 2023/1/11 總收文號 lbePaperNum 改為 textCP09 物件
'Memo by Lydia 2021/10/08 改成Form2.0 ; cmbTM05、textCUID、textCP13_2、textCP14_2、textCP83_2、textCP44_2、textCP64、textCP49、textCP50~CP52
                                                                'textCP55_2、textCP56_2、textCP89_2~CP96_2、textCP144、lstNameAgent、textCP29_2、textCP131、txtTF37
                                                                'txtLP37、txtLP12、txtLP49、txtLP24、txtLP25、lblLP04、lblLP06、lblCP154、lblLP38、lblLP46、lblLP20、lblLP22、lblAF06、lblAF14、lblCP153
                                                                '「審查委員/法院案號」、「審查委員編號」從基本資料頁籤移到對造頁籤
                                                                
'end 2021/10/08
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
'2004/08/18 nick 重新改寫
Option Explicit

Dim StrToOutSystem As String, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim m_CP43 As String
Dim m_CP10 As String
Dim m_CurrDL As Integer
Dim m_CP53 As String
Dim m_CP54 As String
' 本所案號
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_CP09 As String
' 申請國家
Dim m_Nation As String
'add by nickc 2006/01/27
Dim m_CP110 As String
Dim m_strLP32 As String, m_strLP38 As String, m_strLP31 As String 'Add by Sindy 2020/2/26

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
'   textCP01 = m_CP01
'   textCP02 = m_CP02
'   textCP03 = m_CP03
'   textCP04 = m_CP04
   textCP05 = Empty
   textCP06 = Empty
   textCP07 = Empty
   textCP08 = Empty
   textCP10 = Empty
   textCP10_2 = Empty
   textCP11 = Empty
   textCP11_2 = Empty
   textCP12 = Empty
   textCP12_2 = Empty
   textCP13 = Empty
   textCP13_2 = Empty
   textCP14 = Empty
    'Add By Cheng 2003/10/29
    '記錄原承辦人欄位
    Me.textCP14.Tag = Empty
    'ENd
   textCP14_2 = Empty
   textCP15 = Empty
   textCP16 = Empty
   textCP17 = Empty
   textCP18 = Empty
   textCP19 = Empty
   textCP20 = Empty
   textCP21 = Empty
   textCP22 = Empty
   textCP23 = Empty
   textCP24 = Empty
   textCP25 = Empty
   textCP26 = Empty
   textCP27 = Empty
   textCP28 = Empty
   textCP29 = Empty
   textCP29_2 = Empty
   textCP30 = Empty
   textCP31 = Empty
   textCP32 = Empty
   textCP33 = Empty
   textCP34 = Empty
   textCP35 = Empty
   textCP36 = Empty
   textCP37 = Empty
   textCP37_1 = Empty
   textCP38 = Empty
   textCP39 = Empty
   textCP40 = Empty
   textCP41 = Empty
   textCP42 = Empty
   textCP43 = Empty
   textCP44 = Empty
   textCP44_2 = Empty
   textCP45 = Empty
   textCP46 = Empty
   textCP47 = Empty
   textCP48 = Empty
   textCP49 = Empty
   textCP50 = Empty
   textCP51 = Empty
   textCP52 = Empty
   textCP53 = Empty
   textCP54 = Empty
   textCP55 = Empty
   textCP55_2 = Empty
   textCP56 = Empty
   textCP56_2 = Empty
   textCP57 = Empty
   textCP58 = Empty
   textCP58_2 = Empty
   textCP59 = Empty
   textCP60 = Empty
   textCP61 = Empty
   textCP62 = Empty
   textCP63 = Empty
   textCP64 = Empty
   textCP71 = Empty
   textCP72 = Empty
   textCP71_2 = Empty
   textCUID = Empty
   textCP80 = Empty
   textCP81 = Empty
   textCP148 = Empty 'Add By Sindy 2014/6/24
   'add by nick 2004/08/18 加欄位
   textCP82 = Empty
   textCP83 = Empty
   textCP83_2 = Empty
   textCP84 = Empty
   'textCP85 = Empty 'Remove by Morgan 2010/12/30 目前沒用
   textCP86 = Empty
   textCP87 = Empty
   textCP88 = Empty
   textCP89 = Empty
   textCP89_2 = Empty
   textCP90 = Empty
   textCP90_2 = Empty
   textCP91 = Empty
   textCP91_2 = Empty
   textCP92 = Empty
   textCP92_2 = Empty
   textCP93 = Empty
   textCP93_2 = Empty
   textCP94 = Empty
   textCP94_2 = Empty
   textCP95 = Empty
   textCP95_2 = Empty
   textCP96 = Empty
   textCP96_2 = Empty
   textCP113 = Empty 'Added by Morgan 2012/9/6
   textCP114 = Empty 'Added by Morgan 2012/9/6
   textCP117 = Empty 'Add by Morgan 2008/5/14
   textCP118 = Empty 'Add by Morgan 2008/7/11
   textCP152 = Empty 'Added by Lydia 2019/01/15
   '2008/8/27 add by sonia
   TextCP119 = Empty
   'Add by Morgan 2008/11/10
   textCP120 = Empty
   textCP121 = Empty
   'Add by Morgan 2009/3/18
   textCP123 = Empty
   textCP124 = Empty
   textCP125 = Empty
   textCP126 = Empty
   'Modified by Morgan 2014/5/23
   'textCP127 = Empty
   'textCP128 = Empty
   lblCP127128.Caption = ""
   'end 2014/5/23
   textCP129 = Empty
   
   'Add By Sindy 2009/04/27
   textCP130 = Empty
   textCP131 = Empty
   textCP132 = Empty
   
   textCP144 = Empty 'Added by Lydia 2021/10/08
   
   'Add by Morgan 2010/1/5
   textCP135 = Empty
   textCP136 = Empty
   textCP137 = Empty
   textCP138 = Empty
   'Add By Sindy 2023/3/17
   textCP167 = Empty
   textCP168 = Empty
   '2023/3/17 END
   
   'add by nickc 2008/01/31
   lblCP73.Caption = Empty
   lblCP74.Caption = Empty
   lblCP75.Caption = Empty
   lblCP76.Caption = Empty
   lblCP77.Caption = Empty
   lblCP78.Caption = Empty
   lblCP79.Caption = Empty
   
'   'Add By Sindy 2011/6/29
'   txtCR(9).Text = Empty
'   lstAtt.Clear
   
   'Added by Morgan 2014/5/22
   lblCP153.Caption = ""
   lblCP154.Caption = ""
   lblLP11 = ""
   lblLP04 = ""
   lblLP0517 = ""
   lblLP06 = ""
   lblLP0718 = ""
   txtLP12 = ""
   'end 2014/5/22
   txtLP37 = "" 'Added by Morgan 2019/5/14
   'Added by Morgan 2015/1/14
   lblLP20 = ""
   lblLP21 = ""
   lblLP22 = ""
   lblLP23 = ""
   txtLP24 = ""
   txtLP25 = ""
   'end 2015/1/14
   lblLP26 = "" 'Added by Morgan 2015/6/26
   'Added by Moran 2019/4/23
   lblLP38 = ""
   lblLP3940 = ""
   'end 2019/4/23
   'Added by Moran 2021/4/1
   lblLP46 = ""
   lblLP4748 = ""
   txtLP49 = ""
   'end 2021/4/1
   'Add by Morgan 2016/5/18
   lblAF06 = ""
   lblAF0708 = ""
   lblAF14 = ""
   lblAF1112 = ""
   'end 2016/5/18
End Sub

' 讀取商標基本檔
Private Function QueryTradeMark(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryTradeMark = False
   strSql = "SELECT * FROM TRADEMARK " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryTradeMark = True
      '案件名稱
      'Modify By Sindy 2011/6/29
      AddCboName Combo1, "" & rsTmp.Fields("TM05"), "", ""
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         lbl1(0).Caption = CheckStr(rsTmp.Fields("TM23"))
         lbl1(1).Caption = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_Nation = rsTmp.Fields("TM10")
      End If
      ' 專用期間
      If IsNull(rsTmp.Fields("TM21")) = False Then
         m_CP53 = rsTmp.Fields("TM21")
      End If
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_CP54 = rsTmp.Fields("TM22")
      End If
        If IsNull(rsTmp.Fields("TM29")) Then
             Me.lblClose.Caption = ""
        Else
             Me.lblClose.Caption = "已閉卷"
        End If
        'add by nickc 2006/08/28
        If pub_strUserOffice = "1" Then
            If IsNull(rsTmp.Fields("tm57")) Then
                 Me.lblCancel.Caption = ""
            Else
                 Me.lblCancel.Caption = "已銷卷"
            End If
        Else
            If IsNull(rsTmp.Fields("tm73")) Then
                 Me.lblCancel.Caption = ""
            Else
                 Me.lblCancel.Caption = "已銷卷"
            End If
        End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取服務業務基本檔
Private Function QueryServicePractice(ByVal strSP01 As String, ByVal strSP02 As String, ByVal strSP03 As String, ByVal strSP04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryServicePractice = False
   strSql = "SELECT * FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & strSP01 & "' AND " & _
                  "SP02 = '" & strSP02 & "' AND " & _
                  "SP03 = '" & strSP03 & "' AND " & _
                  "SP04 = '" & strSP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryServicePractice = True
      ' 案件名稱
      'Modify By Sindy 2011/6/29
      AddCboName Combo1, "" & rsTmp.Fields("SP05"), "" & rsTmp.Fields("SP06"), "" & rsTmp.Fields("SP07")
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         lbl1(0).Caption = CheckStr(rsTmp.Fields("SP08"))
         lbl1(1).Caption = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_Nation = rsTmp.Fields("SP09")
      End If
      ' 專用期間
      If IsNull(rsTmp.Fields("SP20")) = False Then
         m_CP53 = rsTmp.Fields("SP20")
      End If
      If IsNull(rsTmp.Fields("SP21")) = False Then
         m_CP54 = rsTmp.Fields("SP21")
      End If
        If IsNull(rsTmp.Fields("SP15")) Then
             Me.lblClose.Caption = ""
        Else
             Me.lblClose.Caption = "已閉卷"
        End If
        'add by nickc 2006/08/28
        If pub_strUserOffice = "1" Then
            If IsNull(rsTmp.Fields("sp61")) Then
                 Me.lblCancel.Caption = ""
            Else
                 Me.lblCancel.Caption = "已銷卷"
            End If
        Else
            If IsNull(rsTmp.Fields("sp68")) Then
                 Me.lblCancel.Caption = ""
            Else
                 Me.lblCancel.Caption = "已銷卷"
            End If
        End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取專利基本檔
Private Function QueryPatent(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryPatent = False
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & strPA01 & "' AND " & _
                  "PA02 = '" & strPA02 & "' AND " & _
                  "PA03 = '" & strPA03 & "' AND " & _
                  "PA04 = '" & strPA04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryPatent = True
      ' 案件名稱
      'Modify By Sindy 2011/6/29
      AddCboName Combo1, "" & rsTmp.Fields("PA05"), "" & rsTmp.Fields("PA06"), "" & rsTmp.Fields("PA07")
      ' 申請人
      If IsNull(rsTmp.Fields("PA26")) = False Then
         lbl1(0).Caption = CheckStr(rsTmp.Fields("PA26"))
         lbl1(1).Caption = GetCustomerName(rsTmp.Fields("PA26"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("PA09")) = False Then
         m_Nation = rsTmp.Fields("PA09")
      End If
      ' 專用期間
      If IsNull(rsTmp.Fields("PA24")) = False Then
         m_CP53 = rsTmp.Fields("PA24")
      End If
      If IsNull(rsTmp.Fields("PA25")) = False Then
         m_CP54 = rsTmp.Fields("PA25")
      End If
        If IsNull(rsTmp.Fields("PA57")) Then
             Me.lblClose.Caption = ""
        Else
             Me.lblClose.Caption = "已閉卷"
        End If
        'add by nickc 2006/08/28
        If pub_strUserOffice = "1" Then
            If IsNull(rsTmp.Fields("pa108")) Then
                 Me.lblCancel.Caption = ""
            Else
                 Me.lblCancel.Caption = "已銷卷"
            End If
        Else
            If IsNull(rsTmp.Fields("pa136")) Then
                 Me.lblCancel.Caption = ""
            Else
                 Me.lblCancel.Caption = "已銷卷"
            End If
        End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取法務基本檔
Private Function QueryLawCase(ByVal strLC01 As String, ByVal strLC02 As String, ByVal strLC03 As String, ByVal strLC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryLawCase = False
   strSql = "SELECT * FROM LAWCASE " & _
            "WHERE LC01 = '" & strLC01 & "' AND " & _
                  "LC02 = '" & strLC02 & "' AND " & _
                  "LC03 = '" & strLC03 & "' AND " & _
                  "LC04 = '" & strLC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryLawCase = True
      ' 案件名稱
      'Modify By Sindy 2011/6/29
      AddCboName Combo1, "" & rsTmp.Fields("LC05"), "" & rsTmp.Fields("LC06"), "" & rsTmp.Fields("LC07")
      ' 申請人
      If IsNull(rsTmp.Fields("LC11")) = False Then
         lbl1(0).Caption = CheckStr(rsTmp.Fields("LC11"))
         lbl1(1).Caption = GetCustomerName(rsTmp.Fields("LC11"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("LC15")) = False Then
         m_Nation = rsTmp.Fields("LC15")
      End If
        If IsNull(rsTmp.Fields("LC08")) Then
             Me.lblClose.Caption = ""
        Else
             Me.lblClose.Caption = "已閉卷"
        End If
        'add by nickc 2006/08/28
        If pub_strUserOffice = "1" Then
            If IsNull(rsTmp.Fields("lc34")) Then
                 Me.lblCancel.Caption = ""
            Else
                 Me.lblCancel.Caption = "已銷卷"
            End If
        Else
            If IsNull(rsTmp.Fields("lc36")) Then
                 Me.lblCancel.Caption = ""
            Else
                 Me.lblCancel.Caption = "已銷卷"
            End If
        End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取顧問案件基本檔
Private Function QueryHireCase(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryHireCase = False
   strSql = "SELECT * FROM HIRECASE " & _
            "WHERE HC01 = '" & strHC01 & "' AND " & _
                  "HC02 = '" & strHC02 & "' AND " & _
                  "HC03 = '" & strHC03 & "' AND " & _
                  "HC04 = '" & strHC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryHireCase = True
      ' 案件名稱
      'Modify By Sindy 2011/6/29
      AddCboName Combo1, "" & rsTmp.Fields("HC06"), "", ""
      ' 申請人
      If IsNull(rsTmp.Fields("HC05")) = False Then
         lbl1(0).Caption = CheckStr(rsTmp.Fields("HC05"))
         lbl1(1).Caption = GetCustomerName(rsTmp.Fields("HC05"), 0)
      End If
        If IsNull(rsTmp.Fields("HC09")) Then
             Me.lblClose.Caption = ""
        Else
             Me.lblClose.Caption = "已閉卷"
        End If
        'add by nickc 2006/08/28
        If pub_strUserOffice = "1" Then
            If IsNull(rsTmp.Fields("hc19")) Then
                 Me.lblCancel.Caption = ""
            Else
                 Me.lblCancel.Caption = "已銷卷"
            End If
        Else
            If IsNull(rsTmp.Fields("hc20")) Then
                 Me.lblCancel.Caption = ""
            Else
                 Me.lblCancel.Caption = "已銷卷"
            End If
        End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function QueryCaseProgress() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim stRPTxt As String 'Add byAmy 2022/09/02 '顯示對造/相關人 label名稱
   
   QueryCaseProgress = False
   
   LblCP150Y.Visible = False 'Add By Sindy 2022/10/14
   
   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryCaseProgress = True
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then: textCP05 = ChangeWStringToTString(rsTmp.Fields("CP05"))
      ' 本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then: textCP06 = ChangeWStringToTString(rsTmp.Fields("CP06"))
      ' 法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then: textCP07 = ChangeWStringToTString(rsTmp.Fields("CP07"))
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then: textCP08 = rsTmp.Fields("CP08")
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then: textCP10 = rsTmp.Fields("CP10")
      'Add by Morgan 2004/2/11
      m_CP10 = textCP10
      'Add By Sindy 2023/3/15
      If m_CP01 = "T" And m_CP10 = "210" Then '210.陳述意見書
         Label11.Caption = "是否為快軌案件："
         Label23.Caption = "(Y/N)"
      End If
      '2023/3/15 END
      
      If IsNull(rsTmp.Fields("CP11")) = False Then: textCP11 = rsTmp.Fields("CP11")
      If IsNull(rsTmp.Fields("CP12")) = False Then: textCP12 = rsTmp.Fields("CP12")
      If IsNull(rsTmp.Fields("CP13")) = False Then: textCP13 = rsTmp.Fields("CP13")
      If IsNull(rsTmp.Fields("CP14")) = False Then: textCP14 = rsTmp.Fields("CP14")
        'Add By Cheng 2003/10/29
        '記錄原承辦人
        Me.textCP14.Tag = "" & rsTmp.Fields("CP14").Value
        'End
      If IsNull(rsTmp.Fields("CP15")) = False Then: textCP15 = rsTmp.Fields("CP15")
      If IsNull(rsTmp.Fields("CP16")) = False Then: textCP16 = rsTmp.Fields("CP16")
      If IsNull(rsTmp.Fields("CP17")) = False Then: textCP17 = rsTmp.Fields("CP17")
      If IsNull(rsTmp.Fields("CP18")) = False Then: textCP18 = rsTmp.Fields("CP18")
      If IsNull(rsTmp.Fields("CP19")) = False Then: textCP19 = rsTmp.Fields("CP19")
      If IsNull(rsTmp.Fields("CP20")) = False Then: textCP20 = rsTmp.Fields("CP20")
      If IsNull(rsTmp.Fields("CP21")) = False Then: textCP21 = rsTmp.Fields("CP21")
      If IsNull(rsTmp.Fields("CP22")) = False Then: textCP22 = rsTmp.Fields("CP22")
      If IsNull(rsTmp.Fields("CP23")) = False Then: textCP23 = rsTmp.Fields("CP23")
      If IsNull(rsTmp.Fields("CP24")) = False Then: textCP24 = rsTmp.Fields("CP24")
      ' 准駁日
      If IsNull(rsTmp.Fields("CP25")) = False Then: textCP25 = ChangeWStringToTString(rsTmp.Fields("CP25"))
      If IsNull(rsTmp.Fields("CP26")) = False Then: textCP26 = rsTmp.Fields("CP26")
      ' 發文日
      If IsNull(rsTmp.Fields("CP27")) = False Then: textCP27 = ChangeWStringToTString(rsTmp.Fields("CP27"))
      If IsNull(rsTmp.Fields("CP28")) = False Then: textCP28 = rsTmp.Fields("CP28")
      If IsNull(rsTmp.Fields("CP29")) = False Then: textCP29 = rsTmp.Fields("CP29")
      If IsNull(rsTmp.Fields("CP30")) = False Then: textCP30 = rsTmp.Fields("CP30")
      If IsNull(rsTmp.Fields("CP31")) = False Then: textCP31 = rsTmp.Fields("CP31")
      If IsNull(rsTmp.Fields("CP32")) = False Then: textCP32 = rsTmp.Fields("CP32")
      If IsNull(rsTmp.Fields("CP33")) = False Then: textCP33 = rsTmp.Fields("CP33")
      If IsNull(rsTmp.Fields("CP34")) = False Then: textCP34 = rsTmp.Fields("CP34")
      If IsNull(rsTmp.Fields("CP35")) = False Then: textCP35 = rsTmp.Fields("CP35")
      If IsNull(rsTmp.Fields("CP36")) = False Then: textCP36 = rsTmp.Fields("CP36")
        Select Case m_CP01
        Case "T", "FCT", "CFT", "TF"
            If IsNull(rsTmp.Fields("CP37")) = False Then: textCP37_1 = rsTmp.Fields("CP37")
        Case Else
            If IsNull(rsTmp.Fields("CP37")) = False Then: textCP37 = rsTmp.Fields("CP37")
            If IsNull(rsTmp.Fields("CP38")) = False Then: textCP38 = rsTmp.Fields("CP38")
            If IsNull(rsTmp.Fields("CP39")) = False Then: textCP39 = rsTmp.Fields("CP39")
        End Select
      If IsNull(rsTmp.Fields("CP40")) = False Then: textCP40 = rsTmp.Fields("CP40")
      If IsNull(rsTmp.Fields("CP41")) = False Then: textCP41 = rsTmp.Fields("CP41")
      If IsNull(rsTmp.Fields("CP42")) = False Then: textCP42 = rsTmp.Fields("CP42")
      If IsNull(rsTmp.Fields("CP43")) = False Then: textCP43 = rsTmp.Fields("CP43")
      'Add by Morgan 2004/2/12
      m_CP43 = textCP43

      'Added by Lydia 2021/10/08 比照frm075004_2，顯示報價備註
      If IsNull(rsTmp.Fields("CP144")) = False Then: textCP144 = rsTmp.Fields("CP144")
      If Pub_StrUserSt03 = "M51" Or Mid(Pub_StrUserSt03, 1, 2) = "P1" Then
         Label33(2).Visible = True
         textCP144.Visible = True
      Else
         Label33(2).Visible = False
         textCP144.Visible = False
      End If
      
      'Modify By Cheng 2002/04/22
      '若代理人代號最後面三碼為"000", 則顯示六碼即可
'      If IsNull(rsTmp.Fields("CP44")) = False Then: textCP44 = rsTmp.Fields("CP44")
      If IsNull(rsTmp.Fields("CP44")) = False Then: textCP44 = IIf(Len(rsTmp.Fields("CP44")) = 9 And Right(rsTmp.Fields("CP44"), 3) = "000", Left(rsTmp.Fields("CP44"), 6), rsTmp.Fields("CP44"))
      If IsNull(rsTmp.Fields("CP45")) = False Then: textCP45 = rsTmp.Fields("CP45")
      ' 代理人收達日
      If IsNull(rsTmp.Fields("CP46")) = False Then: textCP46 = ChangeWStringToTString(rsTmp.Fields("CP46"))
      ' 代理人提申日
      If IsNull(rsTmp.Fields("CP47")) = False Then: textCP47 = ChangeWStringToTString(rsTmp.Fields("CP47"))
      'Added by Lydia 2016/05/30 法務或顧問案件時，回執收受日為111111和110101，修改'代理人提申日'欄的Label名稱
'      If m_CP10 = "47" And (Val(textCP46) = 111111 Or Val(textCP46) = 110101) And textCP47 <> "" Then
'         If ClsPDGetSystemKind(m_CP01, intI) Then
'            If intI = 3 Or intI = 4 Then
'               Label30(2) = "回執收受日："
'               If Val(textCP46) = 111111 Then
'                  Label30(3) = "回執退件日："
'               Else
'                  Label30(3) = "回執未回郵局送達日："
'               End If
'            End If
'         End If
'      End If
      Label30(3) = "代理人提申日："
      If ClsPDGetSystemKind(m_CP01, intI) Then
         If intI = 3 Or intI = 4 Then
            If Val(textCP46) = 111111 Then
               Label30(3) = "回執退件日："
            ElseIf Val(textCP46) = 110101 Then
               Label30(3) = "回執未回郵局送達日："
            End If
         End If
      End If
      'end 2016/05/30
      
      ' 承辦期限
      If IsNull(rsTmp.Fields("CP48")) = False Then: textCP48 = ChangeWStringToTString(rsTmp.Fields("CP48"))
      If IsNull(rsTmp.Fields("CP49")) = False Then: textCP49 = rsTmp.Fields("CP49")
      If IsNull(rsTmp.Fields("CP50")) = False Then: textCP50 = rsTmp.Fields("CP50")
      If IsNull(rsTmp.Fields("CP51")) = False Then: textCP51 = rsTmp.Fields("CP51")
      If IsNull(rsTmp.Fields("CP52")) = False Then: textCP52 = rsTmp.Fields("CP52")
      
      'Added by Lydia 2017/08/24 預設顯示
      Label20(7).Visible = False
      Label20(2).Visible = True
      textCP53.Visible = True
      textCP53.Visible = True
      textCP54.Width = 945
      'end 2017/08/24
      
      'Modify By Sindy 2009/07/06
      'Modify by Morgan 2010/6/22 +P 1001
      'modify by sonia 2013/12/27 +CFP 601
      'Modify by Amy 2018/04/10 +612 年費移作次年
      'modify by Sindy 2019/12/13 +(m_CP01 = "FCP" And (m_CP10 = "601" Or m_CP10 = "605")) or
      If (m_CP01 = "P" And (m_CP10 = "601" Or m_CP10 = "1001")) Or _
         (m_CP01 = "CFP" And (m_CP10 = "601")) Or _
         (m_CP01 = "FCP" And (m_CP10 = "601" Or m_CP10 = "605")) Or _
         ((m_CP01 = "P" Or m_CP01 = "CFP") And (m_CP10 = "605" Or m_CP10 = "606" Or m_CP10 = "607" Or m_CP10 = "612" Or m_CP10 = "908")) Then
         Label20(2).Caption = "繳費年度/次數(起/迄)："
         ' 繳費年度/次數(起)
         If IsNull(rsTmp.Fields("CP53")) = False Then: textCP53 = rsTmp.Fields("CP53")
         ' 繳費年度/次數(迄)
         If IsNull(rsTmp.Fields("CP54")) = False Then: textCP54 = rsTmp.Fields("CP54")
      '2009/07/06 End
      'Added by Lydia 2017/08/24 TB條碼案繳年費708,服務業務結果1801-第?期登記期
      ElseIf m_CP01 = "TB" And (m_CP10 = "708" Or m_CP10 = "1801") Then
         Label20(2).Visible = False
         Label20(7).Visible = True
         textCP54.Visible = False
         textCP53.Width = 500
         If IsNull(rsTmp.Fields("CP53")) = False Then textCP53 = rsTmp.Fields("CP53")
         If IsNull(rsTmp.Fields("CP54")) = False Then textCP54 = rsTmp.Fields("CP54")
      'end 2017/08/24
      Else
         Label20(2).Caption = "授權期間/質權設定期間(起/迄)/聘任期間："
         ' 質權設定期間(起)
         If IsNull(rsTmp.Fields("CP53")) = False Then: textCP53 = ChangeWStringToTString(rsTmp.Fields("CP53"))
         ' 質權設定期間(迄)
         If IsNull(rsTmp.Fields("CP54")) = False Then: textCP54 = ChangeWStringToTString(rsTmp.Fields("CP54"))
      End If
      
      'Add by Morgan 2009/10/8
      If m_CP01 = "FCP" And m_CP10 = "908" Then
         lblCP19.Caption = "退費金額："
         lblCP49.Caption = "特定退款人名稱："
         lblCP86.Caption = "是否同意扣除服務費："
         lblCP86_1.Caption = "(N:不同意)"
      Else
         lblCP19.Caption = "後　　金："
         lblCP49.Caption = "條款/當事人稱謂："
         'Add by Morgan 2010/7/1
         If m_CP01 = "FCP" And m_CP10 = "202" Then
            lblCP86.Caption = "是否為複委任："
         Else
         'end 2010/7/1
            lblCP86.Caption = "收到分所接洽單紀錄："
         End If
         lblCP86_1.Caption = "(Y:是)"
      End If
      'end 2009/10/8
      
      If IsNull(rsTmp.Fields("CP55")) = False Then: textCP55 = rsTmp.Fields("CP55")
      If IsNull(rsTmp.Fields("CP56")) = False Then: textCP56 = rsTmp.Fields("CP56")
      ' 取消收文日期
      If IsNull(rsTmp.Fields("CP57")) = False Then: textCP57 = ChangeWStringToTString(rsTmp.Fields("CP57"))
      If IsNull(rsTmp.Fields("CP58")) = False Then: textCP58 = rsTmp.Fields("CP58")
      If IsNull(rsTmp.Fields("CP59")) = False Then: textCP59 = rsTmp.Fields("CP59")
      If IsNull(rsTmp.Fields("CP60")) = False Then: textCP60 = rsTmp.Fields("CP60")
      If IsNull(rsTmp.Fields("CP61")) = False Then: textCP61 = rsTmp.Fields("CP61")
      If IsNull(rsTmp.Fields("CP62")) = False Then: textCP62 = rsTmp.Fields("CP62")
      If IsNull(rsTmp.Fields("CP63")) = False Then: textCP63 = rsTmp.Fields("CP63")
      If IsNull(rsTmp.Fields("CP64")) = False Then: textCP64 = rsTmp.Fields("CP64")
      If IsNull(rsTmp.Fields("CP71")) = False Then: textCP71 = rsTmp.Fields("CP71")
      'Add by Morgan 2004/9/30
      If rsTmp.Fields("CP01") = "P" And rsTmp.Fields("CP10") = "412" Then
         lblCP71.Caption = "延緩公告月數/日期："
      'Add by Morgan 2010/6/3
      ElseIf rsTmp.Fields("CP01") = "CFP" And rsTmp.Fields("CP10") = "106" Then
         lblCP71.Caption = "是否需直譯本："
      'Added by Morgan 2012/4/25
      'modify bys sonia 2013/3/21 加436
      ElseIf rsTmp.Fields("CP01") = "P" And (rsTmp.Fields("CP10") = "405" Or rsTmp.Fields("CP10") = "436" Or rsTmp.Fields("CP10") = "437") Then
         lblCP71.Caption = "優先權份數："
      'end 2012/4/25
      'Added by Morgan 2020/2/4
      ElseIf rsTmp.Fields("CP01") = "P" And rsTmp.Fields("CP10") = "404" Then
         lblCP71.Caption = "延期月數："
      'end 2020/2/4
      'Added by Lydia 2025/02/12 P臺灣與大陸案若申請延緩審查，請於發文時讓user輸入延緩審查日期; FCP案在核准時輸入
      ElseIf (rsTmp.Fields("CP01") = "P" And rsTmp.Fields("CP10") = "245") Or (rsTmp.Fields("CP01") = "FCP" And rsTmp.Fields("CP10") = "1924") Then
         If m_Nation = "000" Then
            lblCP71.Caption = "延緩審查日期："
            If IsNull(rsTmp.Fields("CP71")) = False Then: textCP71 = TransDate(rsTmp.Fields("CP71"), 1)
         Else
            lblCP71.Caption = "延緩審查日期(年度)："
         End If
      'end 2025/02/12
      Else
         lblCP71.Caption = "機關代號："
      End If
      textCP71.Left = lblCP71.Left + lblCP71.Width + 50 'Added by Morgan 2012/4/25
      If IsNull(rsTmp.Fields("CP72")) = False Then: textCP72 = rsTmp.Fields("CP72")
      If IsNull(rsTmp.Fields("CP80")) = False Then: textCP80 = rsTmp.Fields("CP80")
      'add & edit by nick 2004/08/18
      'textCP81 = "" & rsTmp.Fields("CP81") 'Add by Morgan 2004/6/11
      If IsNull(rsTmp.Fields("CP148")) = False Then: textCP148 = rsTmp.Fields("CP148") 'Add By Sindy 2014/6/24
      'Add By Sindy 2015/6/3
      If InStr(m_CP01, "P") > 0 Then
         'Modify By Sindy 2015/9/23
         'Label22 = "是否有檢索："
         If Left(m_CP09, 1) = "C" Then
            Label22 = "是否有檢索："
         Else
            Label22 = "是否為特殊請款："
         End If
         '2015/9/23 END
      Else
         Label22 = "是否為一申請書多件:"
      End If
      '2015/6/3 END
      If IsNull(rsTmp.Fields("CP81")) = False Then: textCP81 = rsTmp.Fields("CP81")
      If IsNull(rsTmp.Fields("CP82")) = False Then: textCP82 = rsTmp.Fields("CP82")
      If IsNull(rsTmp.Fields("CP83")) = False Then: textCP83 = rsTmp.Fields("CP83")
      If IsNull(rsTmp.Fields("CP84")) = False Then: textCP84 = rsTmp.Fields("CP84")
      'If IsNull(rsTmp.Fields("CP85")) = False Then: textCP85 = rsTmp.Fields("CP85")'Remove by Morgan 2010/12/30 目前沒用
      If IsNull(rsTmp.Fields("CP86")) = False Then: textCP86 = rsTmp.Fields("CP86")
      If IsNull(rsTmp.Fields("CP87")) = False Then: textCP87 = rsTmp.Fields("CP87")
      If IsNull(rsTmp.Fields("CP88")) = False Then: textCP88 = rsTmp.Fields("CP88")
      If IsNull(rsTmp.Fields("CP89")) = False Then: textCP89 = rsTmp.Fields("CP89")
      If IsNull(rsTmp.Fields("CP90")) = False Then: textCP90 = rsTmp.Fields("CP90")
      If IsNull(rsTmp.Fields("CP91")) = False Then: textCP91 = rsTmp.Fields("CP91")
      If IsNull(rsTmp.Fields("CP92")) = False Then: textCP92 = rsTmp.Fields("CP92")
      If IsNull(rsTmp.Fields("CP93")) = False Then: textCP93 = rsTmp.Fields("CP93")
      If IsNull(rsTmp.Fields("CP94")) = False Then: textCP94 = rsTmp.Fields("CP94")
      If IsNull(rsTmp.Fields("CP95")) = False Then: textCP95 = rsTmp.Fields("CP95")
      If IsNull(rsTmp.Fields("CP96")) = False Then: textCP96 = rsTmp.Fields("CP96")
      textCP113 = "" & rsTmp.Fields("CP113") 'Add by Morgan 2012/9/6
      textCP114 = "" & rsTmp.Fields("CP114") 'Add by Morgan 2012/9/6
      'Add by Morgan 2008/5/14
      If IsNull(rsTmp.Fields("CP116")) = False Then: textCP44 = textCP44 & "-" & rsTmp.Fields("CP116")
      If IsNull(rsTmp.Fields("CP117")) = False Then: textCP117 = rsTmp.Fields("CP117")
      'Add by Morgan 2008/7/11
      If IsNull(rsTmp.Fields("CP118")) = False Then: textCP118 = rsTmp.Fields("CP118")
      
      'Add By Sindy 2022/10/14
      If strSrvDate(1) >= 接洽單電子收文啟用日 Then
         If IsNull(rsTmp.Fields("CP150")) = False Then
            If rsTmp.Fields("CP150") = "Y" Then LblCP150Y.Visible = True
         End If
      End If
      '2022/10/14 END
      
      'Added by Lydia 2019/01/15 扣款日期
      If IsNull(rsTmp.Fields("CP152")) = False Then: textCP152 = ChangeWStringToTString(rsTmp.Fields("CP152"))
      '2008/8/27 add by sonia 櫃台收文日
      If IsNull(rsTmp.Fields("CP119")) = False Then: TextCP119 = ChangeWStringToTString(rsTmp.Fields("CP119"))
      'Add by Morgan 2008/11/10
      If IsNull(rsTmp.Fields("CP120")) = False Then: textCP120 = rsTmp.Fields("CP120")
      If IsNull(rsTmp.Fields("CP121")) = False Then: textCP121 = rsTmp.Fields("CP121")
      
      'Add by Morgan 2009/3/18
      If IsNull(rsTmp.Fields("CP123")) = False Then: textCP123 = rsTmp.Fields("CP123")
      If IsNull(rsTmp.Fields("CP124")) = False Then: textCP124 = ChangeWStringToTString(rsTmp.Fields("CP124"))
      If IsNull(rsTmp.Fields("CP125")) = False Then: textCP125 = Format(rsTmp.Fields("CP125"), "00:00:00")
      If IsNull(rsTmp.Fields("CP126")) = False Then: textCP126 = rsTmp.Fields("CP126")
      'Modified by Morgan 2014/5/23
      'If IsNull(rsTmp.Fields("CP127")) = False Then: textCP127 = ChangeWStringToTString(rsTmp.Fields("CP127"))
      'If IsNull(rsTmp.Fields("CP128")) = False Then: textCP128 = Format(rsTmp.Fields("CP128"), "00:00:00")
      If Not IsNull(rsTmp.Fields("CP127")) Then
         lblCP127128 = ChangeWStringToTDateString(rsTmp.Fields("CP127"))
      End If
      If Not IsNull(rsTmp.Fields("CP128")) Then
         lblCP127128 = lblCP127128 & " " & Format(rsTmp.Fields("CP128"), "00:00:00")
      End If
      'end 2014/5/23
      If IsNull(rsTmp.Fields("CP129")) = False Then: textCP129 = ChangeWStringToTString(rsTmp.Fields("CP129"))
      
      'Add By Sindy 2009/04/27
      If IsNull(rsTmp.Fields("CP130")) = False Then: textCP130 = rsTmp.Fields("CP130")
      If IsNull(rsTmp.Fields("CP131")) = False Then: textCP131 = rsTmp.Fields("CP131")
      If IsNull(rsTmp.Fields("CP132")) = False Then: textCP132 = ChangeWStringToTString(rsTmp.Fields("CP132"))
      
      'Add by Morgan 2010/12/30
      If IsNull(rsTmp.Fields("CP140")) = False Then textCP140 = rsTmp.Fields("CP140")
      If IsNull(rsTmp.Fields("cp141")) = False Then
         'Add By Sindy 2024/5/27 取消4
         If Val("" & rsTmp.Fields("cp141")) <> 4 Then
         '2024/5/27 END
            OptSendType(Val(rsTmp.Fields("cp141"))).Value = True
         End If
      End If
      If IsNull(rsTmp.Fields("cp142")) = False Then textCP142 = TransDate(rsTmp.Fields("cp142"), 1)
      OptSendType(1).Caption = PUB_GetCP114Opt1Desc(m_CP01, textCP10)  'Added by Morgan 2024/1/22
      'Add By Sindy 2021/4/20
      If "" & rsTmp.Fields("CP164") = "1" Then
         Option1(0).Value = True
      ElseIf "" & rsTmp.Fields("CP164") = "2" Then
         Option1(1).Value = True
      'Add By Sindy 2021/10/20
      ElseIf "" & rsTmp.Fields("CP164") = "3" Then
         Option1(2).Value = True
      End If
      '2021/4/20 END
      'Add By Sindy 2024/5/27 暫不送
      If IsNull(rsTmp.Fields("cp176")) = False Then
         chkCP176.Value = 1
      Else
         chkCP176.Value = 0
      End If
      '2024/5/27 END
      
      'Add by Morgan 2010/1/5
      'Modified by Morgan 2015/8/5 +209,235 --靜芳
      'Modified by Morgan 2022/7/5 +107
      If textCP10 = "416" Or textCP10 = "201" Or textCP10 = "209" Or textCP10 = "235" Or textCP10 = "107" Then
         lblCP136 = "總項數："
         lblCP135 = "總頁數：" 'Add By Sindy 2023/3/17
      Else
         lblCP136 = "增加項數："
         lblCP135 = "增加頁數：" 'Add By Sindy 2023/3/17
      End If
      textCP135 = "" & rsTmp.Fields("CP135")
      textCP136 = "" & rsTmp.Fields("CP136")
      textCP137 = "" & rsTmp.Fields("CP137")
      textCP138 = "" & rsTmp.Fields("CP138")
      'end 2010/1/5
      'Add By Sindy 2023/3/17
      textCP167 = "" & rsTmp.Fields("CP167")
      textCP168 = "" & rsTmp.Fields("CP168")
      strSql = "select * from pagedetail where pd01='" & textCP09.Text & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         cmdPage.BackColor = &HC0C0FF '粉紅色
      Else
         cmdPage.BackColor = &H80000010
      End If
      '2023/3/17 END
      
      'add by nickc 2008/01/31
      If IsNull(rsTmp.Fields("CP73")) = False Then: lblCP73.Caption = rsTmp.Fields("CP73")
      If IsNull(rsTmp.Fields("CP74")) = False Then: lblCP74.Caption = rsTmp.Fields("CP74")
      If IsNull(rsTmp.Fields("CP75")) = False Then: lblCP75.Caption = rsTmp.Fields("CP75")
      If IsNull(rsTmp.Fields("CP76")) = False Then: lblCP76.Caption = rsTmp.Fields("CP76")
      If IsNull(rsTmp.Fields("CP77")) = False Then: lblCP77.Caption = rsTmp.Fields("CP77")
      If IsNull(rsTmp.Fields("CP78")) = False Then: lblCP78.Caption = rsTmp.Fields("CP78")
      '2012/3/15 modify by sonia 加cp60條件,婧瑄說請款單的不要出現,因為收款也不會更新
      If IsNull(rsTmp.Fields("CP79")) = False And (textCP60 = "" Or Left(textCP60, 1) <> "X") Then: lblCP79.Caption = rsTmp.Fields("CP79")
      
      'add by nickc 2006/01/27
      m_CP110 = CheckStr(rsTmp.Fields("cp110"))
      'Modify by Morgan 2007/6/12 加考慮代理人無法再出名的情形 Ex.65002
      'strSQL = "select st01,st02,OA03 from ouragent,staff where oa01='" & m_CP01 & "' and st01=oa02 order by 3 , 1 "
      If m_CP110 <> "" Then
         'Modified by Morgan 2020/3/17 排序改用資料順序不要再抓設定
         'Modified by Morgan 2020/4/14 再改回抓設定
         'strSql = "select st01,st02,instr('" & m_CP110 & "',st01) OA03 from staff,ouragent where instr('" & m_CP110 & "',st01)>0 and oa01(+)='" & m_CP01 & "' and oa02(+)=st01 order by 3 , 1 "
         strSql = "select st01,st02,OA03 from staff,ouragent where instr('" & m_CP110 & "',st01)>0 and oa01(+)='" & m_CP01 & "' and oa02(+)=st01 order by 3 , 1 "
         'end 2020/3/17
      'end 2007/6/12
         CheckOC
         lblNameAgent.Caption = "出名代理人："
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount > 0 Then
            Do While Not adoRecordset.EOF
               If InStr(m_CP110, CheckStr(adoRecordset.Fields(0))) > 0 Then
                  lblNameAgent.Caption = lblNameAgent.Caption & CheckStr(adoRecordset.Fields(1)) & "、"
               End If
               adoRecordset.MoveNext
            Loop
         End If
         CheckOC
         If Right(lblNameAgent.Caption, 1) = "、" Then
           lblNameAgent.Caption = Mid(lblNameAgent.Caption, 1, Len(lblNameAgent.Caption) - 1)
         End If
      End If
      
      'Add by Morgan 2010/5/13
      If m_CP01 = "FCP" And textCP10 = "202" And textCP86 = "Y" Then
         lblNameAgent = lblNameAgent & "( 複委任 )"
      End If
      
      'Added by Morgan 2014/5/22
      If Not IsNull(rsTmp.Fields("CP153")) Then
         lblCP153 = rsTmp.Fields("CP153") & " " & GetStaffName(rsTmp.Fields("CP153"), True)
      End If
      If Not IsNull(rsTmp.Fields("CP154")) Then
         lblCP154 = rsTmp.Fields("CP154") & " " & GetStaffName(rsTmp.Fields("CP154"), True)
      End If
      'end 2014/5/22
      
      ' 更新欄位的內容
      UpdateCUID rsTmp

      textCP10_Validate False
      textCP11_Validate False
      textCP12_Validate False
      textCP13_Validate False
      textCP14_Validate False
      textCP29_Validate False
      textCP44_Validate False
      textCP55_Validate False
      textCP56_Validate False
      textCP58_Validate False
      textCP71_Validate False
      'add & edit by nick 2004/08/18
      textCP83_Validate False
      textCP89_Validate False
      textCP90_Validate False
      textCP91_Validate False
      textCP92_Validate False
      textCP93_Validate False
      textCP94_Validate False
      textCP95_Validate False
      textCP96_Validate False
   End If
   rsTmp.Close
   
   SetLetter 'Added by Morgan 2014/5/22

   'Added by Lydia 2019/10/25 FCP新案翻譯增加"翻譯瑕疵" => 隱藏"條款/當事人稱謂CP49"
   frmBill.Left = lblCP49.Left
   frmBill.BackColor = &H8000000F
   fraTF.Left = lblCP49.Left
   fraTF.BackColor = &H8000000F
   'end 2019/10/25
   
   'Added by Lydia 2018/01/09 翻譯費用檔-原文字數、相似度、相似案號
   fraTF.Visible = False
   'textCP49.Height = 660 'Mark by Lydia 2019/10/25
   If InStr(FCPHaveEP04, m_CP10) > 0 Then '分析有tranfee的收文號案件性質：201,209,210,927
        'Modified by Lydia 2019/10/25 +TF37
        strSql = "SELECT TF01,TF23,TF19,TF20,TF37 FROM TRANSFEE WHERE TF01 = '" & m_CP09 & "' "
        intI = 1
        Set rsTmp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
             fraTF.Visible = True
             'Modified by Lydia 2019/10/25改成隱藏
             'textCP49.Height = 330
             lblCP49.Visible = False: textCP49.Visible = False
             txtTF23.Text = "" & rsTmp.Fields("TF23")
             txtTF19.Text = "" & rsTmp.Fields("TF19")
             txtTF20.Text = "" & rsTmp.Fields("TF20")
             txtTF37.Text = "" & rsTmp.Fields("TF37") 'Added by Lydia 2019/10/25
        End If
   End If
   'end 2018/01/09
   
   'Added by Morgan 2019/8/20 FCP,FG及FMP案翻譯費帳單帶出備註內的字數及金額--何淑華
   If Left(textCP12, 2) = "F2" And (m_CP01 = "FCP" Or m_CP01 = "FG" Or m_CP01 = "P") And InStr(FCPHaveEP04, m_CP10) > 0 Then
      frmBill.Visible = True
      If textCP61 <> "" Then
         strSql = "SELECT A1509 FROM ACC150 WHERE A1501 = '" & textCP61 & "' "
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            intI = InStr("" & rsTmp.Fields(0), "/")
            If intI > 0 Then
               txtBillWords = Left(rsTmp.Fields(0), intI - 1)
               txtBillFee = Mid(rsTmp.Fields(0), intI + 1)
            End If
         End If
      End If
   Else
      frmBill.Visible = False
   End If
   'end 2019/8/20
   
   'Added by Morgan 2021/1/4
   If (Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 2) = "P1") And m_CP01 = "CFP" And m_CP10 = "214" Then
      'Modified by Morgan 2021/1/18 +判斷有Form才顯示
      If Not Forms(0).GetForm("frm090401_1") Is Nothing Then
         Command2.Visible = True
      End If
   End If
   'end 2021/1/4
   
   'Added by Lydia 2021/05/05 ACS智財顧問專業分配比例管制
   'Modified by Lydia 2024/04/15 +LA之顧問聘任0
   'If m_CP01 = "ACS" And m_CP10 = "112" Then
   If (m_CP01 = "ACS" And m_CP10 = "112") Or (m_CP01 = "LA" And m_CP10 = "0") Then
       Label8.Caption = "簽約時數："
   End If
   'end 2021/05/05
   
   'Add by Amy 2022/09/02 若為「其他相關人」對造/其他 頁籤 顯示 關係案/其他,「對造」文字->對方
   stRPTxt = "對造"
   SSTab1.TabCaption(3) = "對造/其他"
   If Pub_ChkRelevantPeople(1, textCP09.Text) = True Then
        SSTab1.TabCaption(3) = "關係案/其他"
        stRPTxt = "對方"
   End If
   Call SetLabTxt(stRPTxt)
   'end 2022/09/02
   
   Set rsTmp = Nothing
EXITSUB:
End Function

' 將資料庫中的資料更新到所有欄位中
Private Function UpdateCtrlData() As Boolean
UpdateCtrlData = False
   m_CP53 = Empty
   m_CP54 = Empty
   ' 清除欄位內容
   ClearField
   ' 依本所案號讀取基本檔案
   Select Case m_CP01
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         If QueryTradeMark(m_CP01, m_CP02, m_CP03, m_CP04) = False Then
            Exit Function
         End If
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         If QueryPatent(m_CP01, m_CP02, m_CP03, m_CP04) = False Then
                Exit Function
         End If
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         If QueryLawCase(m_CP01, m_CP02, m_CP03, m_CP04) = False Then
                Exit Function
         End If
      ' 讀取顧問案件基本檔
      Case "LA":
         If QueryHireCase(m_CP01, m_CP02, m_CP03, m_CP04) = False Then
                Exit Function
         End If
      ' 讀取服務業務基本檔
      Case Else:
         If QueryServicePractice(m_CP01, m_CP02, m_CP03, m_CP04) = False Then
                Exit Function
         End If
   End Select
   ' 讀取案件進度檔
   If QueryCaseProgress = False Then
        Exit Function
   End If
   UpdateCtrlData = True
End Function

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   If IsNull(rsSrcTmp.Fields("CP65")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP65")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("CP65"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP66")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP66")) = False Then
         strTemp = ChangeWStringToTString(rsSrcTmp.Fields("CP66"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP67")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP67")) = False Then
         strTemp = rsSrcTmp.Fields("CP67")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP68")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP68")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("CP68"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP69")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP69")) = False Then
         strTemp = ChangeWStringToTString(rsSrcTmp.Fields("CP69"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP70")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP70")) = False Then
         strTemp = rsSrcTmp.Fields("CP70")
         strUTime = Format(strTemp, "##:##")
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

'Added by Morgan 2020/2/10
'目前只開放電腦中心
Private Sub cmdCancelConfirm_Click()
Dim strLP06 As String
   
   If MsgBox("確定要取消確認？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
      strSql = "update letterprogress set lp07=0,lp11=null where lp01='" & m_CP09 & "' and lp07>0"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
      
      'Add by Sindy 2020/2/26 商標處原為寄E-Mail改紙本 del:(InStr(lblCP154, "QPGMR") > 0 And m_strLP31 = "Y")
      If lblLP06 <> "" Then
         strLP06 = Trim(Mid(lblLP06, 1, InStr(lblLP06, " ")))
      End If
      '內商:確認人員要改回(P2002:商標處程序人員)
      '1.大宗發文
      '2.智權人員不是內商程序,但確認人員為內商程序人員
      'Modify By Sindy 2022/6/8 確認人員為內商程序人員時,才要改回(P2002:商標處程序人員)
      If Left(m_CP01, 1) = "T" And PUB_GetST03(strLP06) = "P22" Then
         'Added by Morgan 2025/2/18 非MCTF的案件不要改 Ex:T-196819
         'Modify By Sindy 2025/10/17 +D類
         If PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04) = "MCTF" _
            Or Left(m_CP09, 1) = "D" Then
         
'      If Left(m_CP01, 1) = "T" And _
'         (m_strLP32 = "Y" Or _
'          (textCP12 <> "P22" And PUB_GetST03(strLP06) = "P22") _
'         ) Then
      '2022/6/8 END
         'Add By Sindy 2020/7/24 商標大宗發文,未寄信誤上確認,恢復為商標處程序人員為待確認人員
         'If m_strLP32 = "Y" Then
            strSql = "update letterprogress set lp06='P2002' where lp01='" & m_CP09 & "'"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
         '2020/7/24 END
'         ElseIf Trim(m_strLP38) <> "" Then
'            If MsgBox("是否要改經發文室【紙本】送件？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
'               'Modify By Sindy 2021/1/8 取消更新cp28 => cp28=null,
'               strSql = "update caseprogress set cp127=null,cp128=null,cp154=null where cp09='" & m_CP09 & "' and cp154='QPGMR'"
'               Pub_SeekTbLog strSql
'               cnnConnection.Execute strSql, intI
'            End If
         'End If
         End If 'Added by Morgan 2025/2/18
      End If
      '2020/2/26 END
      
      SetLetter
   End If
End Sub

'Added by Morgan 2019/3/15
'改不通知客戶(目前只開放輸入人員操作)
Private Sub cmdCancelLetter_Click()
   Dim stSQL As String, intR As Integer, stLP12 As String
   
   stLP12 = InputBox("不通知原因:")
   If stLP12 = "" Then MsgBox "不通知原因不可空白！", vbCritical: Exit Sub
   If MsgBox("客戶函將設定為不通知，原因為 " & stLP12 & " 是否確定？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
      Exit Sub
   End If
   
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
   
   'Modify By Sindy 2020/2/25 + cp28=null,
   'Modify By Sindy 2021/1/8 取消更新cp28 => cp28=null,
   stSQL = "update caseprogress set cp127=null,cp128=null,cp154=null where cp09='" & m_CP09 & "' and cp154='QPGMR'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intR
   
   stSQL = "update letterprogress set lp06='" & strUserNum & "',lp07=to_char(sysdate,'yyyymmdd'),lp10='N',lp12='" & ChgSQL(stLP12) & "' where lp01='" & m_CP09 & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intR
   
   'Add By Sindy 2020/2/25 目前是商標處C類進度,才會遇到此狀況需更新
   If Left(m_CP09, 1) = "C" And Left(m_CP01, 1) = "T" Then
      stSQL = "update engineerprogress set ep11='N' where ep02='" & m_CP09 & "' and ep11='Y'"
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intR
      If intR = 1 Then
         stSQL = "update caseprogress set cp27=19221111 where cp09='" & m_CP09 & "'"
         Pub_SeekTbLog stSQL
         cnnConnection.Execute stSQL, intR
      End If
   End If
   '2020/2/25 END
   
   cnnConnection.CommitTrans
   If intR = 1 Then
      MsgBox "已改不通知客戶！" & vbCrLf & vbCrLf & "若卷宗區有客戶函，請自行刪除！", vbExclamation
      SetLetter
   End If
   Exit Sub
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Sub

Private Sub cmdCancelLP05_Click()
   If MsgBox("確定要取消判發？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
      strSql = "update letterprogress set lp05=0,lp10='Y' where lp01='" & m_CP09 & "' and lp05>0"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
      SetLetter
   End If
End Sub

'Add By Sindy 2023/3/17 增刪頁數
Private Sub cmdPage_Click()
   frm075004_3.bolModify = False
   frm075004_3.strReceiveNo = textCP09.Text
   Call frm075004_3.SetParent(Me)
   frm075004_3.Show vbModal
End Sub

''開啟附件
'Private Sub cmdOpenAtt_Click()
'   If lstAtt.Text = "" Then
'      MsgBox "請選擇欲開啟的附件！"
'   Else
'      PUB_OpenFtpFile textCP09.Text, lstAtt.Text, Winsock1, "2"
'   End If
'End Sub

Private Sub Command1_Click()
    If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Call frm090638_1.SetParent(Me)
    frm090638_1.BFormPeople = 3
    frm090638_1.m_NC01 = Me.textCP09.Text
    frm090638_1.Show
    If frm090638_1.QueryData = False Then
        ShowNoData
        frm090638_1.cmdState = 0
        frm090638_1.PubShowNextData
    End If
    Screen.MousePointer = vbDefault
    Me.Enabled = True
End Sub

'Added by Morgan 2021/1/4
'美專IDS清單
Private Sub Command2_Click()
   Dim nFrm As Form
   Set nFrm = Forms(0).GetForm("frm090401_1")
   If Not nFrm Is Nothing Then
      nFrm.m_CP09 = m_CP09
      nFrm.m_bQuery = True
      nFrm.Show vbModal
   End If
End Sub

'Add By Sindy 2011/6/24
Private Sub Command4_Click()
   'Modified by Lydia 2022/12/08
'   frm071018.Hide
'   Set frm071018.UpForm = frm100101_C
'   frm071018.lbePaperNum = Me.textCP09.Text
'   frm071018.lbeNumber = Me.lbeNumber
'   frm071018.cmdInput.Visible = False  '新增按鈕
'   frm071018.cmdCancel.Visible = False '刪除按鈕
'   'Modify By Sindy 2020/6/10
'   'frm071018.txtReceiver.Visible = False
'   frm071018.cboEmp.Visible = False
'   '2020/6/10 END
'   frm071018.Label26.Visible = False
   Call frm071018.SetParent(Me, Me.textCP09.Text, True)
   'end 2022/12/08
   Me.Hide
   frm071018.Show vbModal
End Sub

Private Sub Form_Activate()
Dim strSQL1 As String, Str01 As String
   
   'Add By Sindy 2011/6/24
   '其他出庭律師+配合開庭承辦人
   strPublicTemp = ""
   'strExc(0) = "select cl02 from caselawer where cl01='" + StringTwoString(Me.Tag, 2) + "' order by cl02 asc "
   strExc(0) = "select distinct st01 from " & _
               "(select st01 from caselawer,staff " & _
               "where cl01='" & StringTwoString(Me.Tag, 2) & "' " & _
               "and cl02=st01(+) " & _
               "Union " & _
               "select st01 from caseprogress,staff " & _
               "where cp09 in (select a0n02 from acc0n0 where a0n02<>a0n01 and a0n01 in (select cp43 from caseprogress where cp09='" & StringTwoString(Me.Tag, 2) & "')) " & _
               "and cp14=st01(+)) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If IsNull(RsTemp.Fields(0).Value) = False Then
            strPublicTemp = strPublicTemp & RsTemp.Fields(0).Value & ","
         End If
         RsTemp.MoveNext
      Loop
   End If
   If strPublicTemp = "" Then
      Command4.Enabled = False
   Else
      Command4.Enabled = True
   End If
   
   '本所案號
   lbeNumber.Caption = StringTwoString(Me.Tag, 1)
   If Left(Me.Tag, 1) = "N" Then
      strSQL1 = Right(Me.Tag, Len(Me.Tag) - 1)
   Else
      strSQL1 = Me.Tag
   End If
   Str01 = SystemNumber(StringTwoString(strSQL1, 1), 1)
'   Call QueryCourtyardPeriod(StringTwoString(Me.Tag, 2), Str01)
   '2011/6/24 End
End Sub

Private Sub Form_Initialize()
ReDim m_FieldList(TF_CP)
End Sub

' 案件性質
Private Sub textCP10_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
    Cancel = False
    textCP10_2 = Empty
    If IsEmptyText(textCP10) = False Then
        If m_Nation < "010" Then
            textCP10_2 = GetCaseTypeName(m_CP01, textCP10, 0)
        Else
            textCP10_2 = GetCaseTypeName(m_CP01, textCP10, 1)
        End If
    End If
    'Add By Cheng 2003/08/20
    Select Case m_CP01
    Case "P", "CFP", "FCP"
        'edit by nickc 2006/09/01 不鎖定只有核准，因為不續辦、取消收文、閉卷也要
        'Select Case Me.textCP10.Text
        'Case "1001", "1002"
            If Me.textCP10_2.Text <> "" Then Me.textCP10_2.Text = Me.textCP10_2.Text & PUB_GetRelateCasePropertyName(m_CP09, "1")
        'End Select
    Case "T", "TF", "CFT", "FCT"
        'edit by nickc 2006/09/01 不鎖定只有核准，因為不續辦、取消收文、閉卷也要
        'Select Case Me.textCP10.Text
        'Case "1001", "1002", "1003", "1004"
            If Me.textCP10_2.Text <> "" Then Me.textCP10_2.Text = Me.textCP10_2.Text & PUB_GetRelateCasePropertyName(m_CP09, "1")
        'End Select
    End Select
End Sub

' 案件來源代號
Private Sub textCP11_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCP11_2 = Empty
   If IsEmptyText(textCP11) = False Then
      strSql = "SELECT * FROM CASESOURCEMAP " & _
               "WHERE CSM01 = '" & textCP11 & "' "
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CSM02")) = False Then
            textCP11_2 = rsTmp.Fields("CSM02")
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
End Sub

' 業務區別
Private Sub textCP12_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP12_2 = Empty
   If IsEmptyText(textCP12) = False Then
      textCP12_2 = GetDepartmentName(textCP12)
   End If
End Sub

'Add By Sindy 2010/11/25
Private Sub textCP13_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 智權人員代號
Private Sub textCP13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP13_2 = Empty
   If IsEmptyText(textCP13) = False Then
      textCP13_2 = GetStaffName(textCP13, True)
   End If
End Sub

'Add By Sindy 2010/11/25
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14, True)
   End If
End Sub

'Add By Sindy 2010/11/25
Private Sub textCP29_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 繪圖人員/協辦人員
Private Sub textCP29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP29_2 = Empty
   If IsEmptyText(textCP29) = False Then
      textCP29_2 = GetStaffName(textCP29, True)
   End If
End Sub

' 代理人
Private Sub textCP44_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Cheng 2002/07/09
   Dim strTempName As String
   
   Cancel = False
   textCP44_2 = Empty
   If IsEmptyText(textCP44) = False Then
      'Add by Morgan 2008/5/14 +聯絡人
      If InStr(textCP44, "-") > 0 Then
         'modify by sonia 2017/11/15
         'If ClsPDGetContact(textCP44, strTempName) Then
         '   textCP44_2 = strTempName
         'End If
         If PUB_GetAgentName(m_CP01, Left(textCP44, InStr(textCP44, "-") - 1), strTempName) = True Then
            textCP44_2 = strTempName
         End If
         If ClsPDGetContact(textCP44, strTempName) Then
            textCP44_2 = textCP44_2 & "(" & strTempName & ")"
         End If
         'end 2017/11/15
      Else
      'end 2008/5/14
         If PUB_GetAgentName(m_CP01, Me.textCP44.Text, strTempName) = True Then
            textCP44_2 = strTempName
         End If
      End If
   End If
End Sub

' 移轉人
Private Sub textCP55_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP55_2 = Empty
   If IsEmptyText(textCP55) = False Then
      textCP55_2 = GetCustomerName(textCP55, 0)
   End If
End Sub

' 移轉申請人
Private Sub textCP56_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP56_2 = Empty
   If IsEmptyText(textCP56) = False Then
      textCP56_2 = GetCustomerName(textCP56, 0)
   End If
End Sub

' 取消收文原因
Private Sub textCP58_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCP58_2 = Empty
   If IsEmptyText(textCP58) = False Then
      strSql = "SELECT * FROM REASONOFRELIEF " & _
               "WHERE ROR01 = '" & textCP58 & "' "
      Set rsTmp = New ADODB.Recordset
      ' 讀取資料庫
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      ' 檢查讀取的資料筆數
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("ROR02")) = False Then
            textCP58_2 = rsTmp.Fields("ROR02")
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
End Sub

' 機關代號
Private Sub textCP71_Validate(Cancel As Boolean)
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP71_2 = Empty
   If IsEmptyText(textCP71) = False Then
      strSql = "SELECT * FROM ORGANIZATION " & _
               "WHERE OR01 = '" & textCP71 & "' "
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("OR02")) = False Then
            textCP71_2 = rsTmp.Fields("OR02")
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
End Sub

' 被授權人
Private Sub textCP72_Validate(Cancel As Boolean)
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strData As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP72) = False Then
            strData = textCP72 & String(9 - Len(textCP72), "0")
            strSql = "SELECT * FROM CUSTOMER " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '" & Mid(strData, 9, 1) & "' "
            Set rsTmp = New ADODB.Recordset
            ' 讀取資料庫
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
            ' 檢查讀取的資料 當被授權人中英日為空白時才取代
            If rsTmp.RecordCount > 0 Then
               If IsEmptyText(textCP50) = True Then
                  If IsNull(rsTmp.Fields("CU04")) = False Then
                     textCP50 = rsTmp.Fields("CU04")
                  End If
               End If
               If IsEmptyText(textCP51) = True Then
                  If IsNull(rsTmp.Fields("CU05")) = False Then
                     textCP51 = rsTmp.Fields("CU05")
                  End If
               End If
               If IsEmptyText(textCP52) = True Then
                  If IsNull(rsTmp.Fields("CU06")) = False Then
                     textCP52 = rsTmp.Fields("CU06")
                  End If
               End If
            End If
            rsTmp.Close
            Set rsTmp = Nothing
   End If
End Sub

Private Sub textCP83_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP83_2 = Empty
   If IsEmptyText(textCP83) = False Then
      textCP83_2 = GetStaffName(textCP83, True)
   End If
End Sub

Private Sub textCP89_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP89_2 = Empty
   If IsEmptyText(textCP89) = False Then
      textCP89_2 = GetCustomerName(textCP89, 0)
   End If
End Sub

Private Sub textCP90_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP90_2 = Empty
   If IsEmptyText(textCP90) = False Then
      textCP90_2 = GetCustomerName(textCP90, 0)
   End If
End Sub

Private Sub textCP91_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP91_2 = Empty
   If IsEmptyText(textCP91) = False Then
      textCP91_2 = GetCustomerName(textCP91, 0)
   End If
End Sub

Private Sub textCP92_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP92_2 = Empty
   If IsEmptyText(textCP92) = False Then
      textCP92_2 = GetCustomerName(textCP92, 0)
   End If
End Sub

Private Sub textCP93_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP93_2 = Empty
   If IsEmptyText(textCP93) = False Then
      textCP93_2 = GetCustomerName(textCP93, 0)
   End If
End Sub

Private Sub textCP94_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP94_2 = Empty
   If IsEmptyText(textCP94) = False Then
      textCP94_2 = GetCustomerName(textCP94, 0)
   End If
End Sub

Private Sub textCP95_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP95_2 = Empty
   If IsEmptyText(textCP95) = False Then
      textCP95_2 = GetCustomerName(textCP95, 0)
   End If
End Sub

Private Sub textCP96_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP96_2 = Empty
   If IsEmptyText(textCP96) = False Then
      textCP96_2 = GetCustomerName(textCP96, 0)
   End If
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Dim s

frm100101_C.Tag = "" 'Add By Sindy 2020/10/6
Select Case cmdState
Case 2
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 3
     'fnCloseAllFrm100
     'Modify By Sindy 2020/10/6
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
     If frm100101_C.Tag = "" Then
         fnCloseAllFrm100
     Else
         intI = MsgBox("尚有資料，確定要離開此作業嗎？", vbInformation + vbYesNo + vbDefaultButton1)
         If intI = vbYes Then
            fnCloseAllFrm100
         End If
     End If
     '2020/10/6 END
Case 4
     Call StrMenu1
     If IsNull(StrToOutSystem) Or Len(StrToOutSystem) = 0 Then
         s = MsgBox("已經沒有相關總收文號", , "USER 建立資料不夠完全")
     End If
Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
End Sub

'add by nick 2004/08/18  整合後的查詢資料
Sub QueryData(oStrCp09 As String)
Dim rsTmp As New ADODB.Recordset 'Added by Lydia 2020/05/21
   
If Trim(oStrCp09) = "" Then
    StrToOutSystem = ""
    Exit Sub
Else
    StrToOutSystem = oStrCp09
End If
'edit by nickc 2006/02/22 原先有秀，後來改不秀，現在又要秀了
textCP09.Text = StrToOutSystem
Dim strSql  As String, strSQL1 As String
Dim Str01 As String

If Left(Me.Tag, 1) = "N" Then
   strSQL1 = Right(Me.Tag, Len(Me.Tag) - 1)
Else
   strSQL1 = Me.Tag
End If
'Add By Cheng 2002/04/29
      m_CP01 = Empty
      m_CP02 = Empty
      m_CP03 = Empty
      m_CP04 = Empty
      m_CP09 = Empty
Me.lblClose.Caption = ""
'add by nickc 2006/08/28
Me.lblCancel.Caption = ""

'Added by Lydia 2020/05/21 相關總收文號的本所案號不同；ex.TT-999999
strSql = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04 FROM CASEPROGRESS WHERE CP09 = '" & oStrCp09 & "' "
rsTmp.CursorLocation = adUseClient
rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If rsTmp.RecordCount > 0 Then
   If pub_QL04 <> "" Then InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2025/8/7
   Str01 = "" & rsTmp.Fields(0)
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/7
End If
rsTmp.Close
Set rsTmp = Nothing
If Str01 <> "" Then
   m_CP01 = SystemNumber(Str01, 1)
   m_CP02 = SystemNumber(Str01, 2)
   m_CP03 = SystemNumber(Str01, 3)
   m_CP04 = SystemNumber(Str01, 4)
   lbeNumber.Caption = Str01
Else
'end 2020/05/21
    m_CP01 = SystemNumber(StringTwoString(strSQL1, 1), 1)
    m_CP02 = SystemNumber(StringTwoString(strSQL1, 1), 2)
    m_CP03 = SystemNumber(StringTwoString(strSQL1, 1), 3)
    m_CP04 = SystemNumber(StringTwoString(strSQL1, 1), 4)
End If 'Added by Lydia 2020/05/21
m_CP09 = oStrCp09

Select Case m_CP01
Case "T", "FCT", "CFT", "TF"
        Me.lblAll.Visible = True
        Me.Label18(1).Visible = False
        Me.textCP37.Visible = False
        Me.textCP37.Enabled = False
        Me.Label18(2).Visible = False
        Me.textCP38.Visible = False
        Me.textCP38.Enabled = False
        Me.Label18(3).Visible = False
        Me.textCP39.Visible = False
        Me.textCP39.Enabled = False
        Me.Label18(5).Visible = True
        Me.textCP37_1.Visible = True
        Me.textCP37_1.Enabled = True
Case Else
        Me.lblAll.Visible = False
        Me.Label18(1).Visible = True
        Me.textCP37.Visible = True
        Me.textCP37.Enabled = True
        Me.Label18(2).Visible = True
        Me.textCP38.Visible = True
        Me.textCP38.Enabled = True
        Me.Label18(3).Visible = True
        Me.textCP39.Visible = True
        Me.textCP39.Enabled = True
        Me.Label18(5).Visible = False
        Me.textCP37_1.Visible = False
        Me.textCP37_1.Enabled = False
End Select

UpdateCtrlData
End Sub

Sub StrMenu()
Dim strSql  As String, strSQL1 As String
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String
'Add By Cheng 2002/07/08
Dim StrSQLa As String

Str01 = ""
Str02 = ""
Str03 = ""
Str04 = ""
'總收文號
textCP09.Text = StringTwoString(Me.Tag, 2)
'本所案號
lbeNumber.Caption = StringTwoString(Me.Tag, 1)
If Left(Me.Tag, 1) = "N" Then
   strSQL1 = Right(Me.Tag, Len(Me.Tag) - 1)
Else
   strSQL1 = Me.Tag
End If
'Add By Cheng 2002/04/29
Me.lblClose.Caption = ""
'add by nickc 2006/08/28
Me.lblCancel.Caption = ""

Str01 = SystemNumber(StringTwoString(strSQL1, 1), 1)
Str02 = SystemNumber(StringTwoString(strSQL1, 1), 2)
Str03 = SystemNumber(StringTwoString(strSQL1, 1), 3)
Str04 = SystemNumber(StringTwoString(strSQL1, 1), 4)
'add by nick 2004/08/18  ***start
pub_QL05 = ";本所案號：" & Str01 & "-" & Str02 & "-" & Str03 & "-" & Str04 & ";總收文號：" & StringTwoString(Me.Tag, 2) & "(案件進度)" 'Add By Sindy 2025/8/7
QueryData StringTwoString(Me.Tag, 2)
Exit Sub '***end

End Sub

Sub StrMenu1()

'add by nick 2004/08/18  ***start
pub_QL05 = ";本所案號：" & m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & ";相關總收文號：" & textCP43.Text & "(案件進度)" 'Add By Sindy 2025/8/7
QueryData textCP43.Text
Exit Sub '***end

End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   If bolFNation = False Then
       Label9.Visible = False
       textCP44.Visible = False
       textCP44_2.Visible = False
   End If
   cmdState = -1
   'Modified by Lydia 2021/10/08 Object => Control
   Dim oObj As Control
   For Each oObj In Me.Controls
       'Modified by Morgan 2015/1/14
       'If TypeName(oObj) = "TextBox" Then
       'Modified by Morgan 2019/5/14 +txtLP37
       'Modified by Morgan 2021/4/1 +txtLP49
       If TypeName(oObj) = "TextBox" And oObj.Name <> "txtLP37" And oObj.Name <> "txtLP12" And oObj.Name <> "txtLP24" And oObj.Name <> "txtLP25" And oObj.Name <> "txtLP49" Then
       'end 2015/1/14
           oObj.Text = Empty
           'oObj.Appearance = 0 'Remove by Lydia 2021/10/08 因為Form 2.0的TextBox沒有這項屬性
           oObj.BorderStyle = 0
           oObj.BackColor = &H8000000F
           oObj.Locked = True
       End If
   Next
   SSTab1.Tab = 0
   
   'Added by Lydia 2021/10/08
   lblNameAgent.BackColor = &H8000000F
   For Each oObj In lbl1
       oObj.BackColor = &H8000000F
   Next
   lblLP04.BackColor = &H8000000F
   lblLP06.BackColor = &H8000000F
   lblCP153.BackColor = &H8000000F
   lblCP154.BackColor = &H8000000F
   lblLP38.BackColor = &H8000000F
   lblLP46.BackColor = &H8000000F
   lblLP20.BackColor = &H8000000F
   lblLP22.BackColor = &H8000000F
   lblAF06.BackColor = &H8000000F
   lblAF14.BackColor = &H8000000F
   'end 2021/10/08
   Frame2.BackColor = &H8000000F 'Added by Lydia 2021/11/09
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100101_C = Nothing
   'Add By Sindy 2011/6/24
   strPublicTemp = ""
   Unload frm071018
   '2011/6/24 End
End Sub

'Add By Sindy 2011/6/27
' 讀取法務基本檔
'Private Function QueryCourtyardPeriod(ByVal strCP09 As String, ByVal strCP01 As String) As Boolean
'Dim strSql As String, strST11 As String
'Dim bolReadY As Boolean
'Dim rsTmp As New ADODB.Recordset
'
'   QueryCourtyardPeriod = False
'
'   strExc(0) = "SELECT ST11 FROM Staff " & _
'                "WHERE ST01 = '" & strUserNum & "' "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   strST11 = ""
'   If intI = 1 Then
'      strST11 = "" & RsTemp.Fields("ST11")
'   End If
'
'   strExc(0) = "SELECT cp13,st52,st53,st54,st55,a0908 FROM caseprogress,staff,acc090 " & _
'                "WHERE cp09 = '" & strCP09 & "' " & _
'                  "and cp13=st01(+) and st03=a0901(+) "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   bolReadY = False
'   If intI = 1 Then
'      If strUserNum = "" & RsTemp.Fields("cp13") Or _
'         strUserNum = "" & RsTemp.Fields("st52") Or _
'         strUserNum = "" & RsTemp.Fields("st53") Or _
'         strUserNum = "" & RsTemp.Fields("st54") Or _
'         strUserNum = "" & RsTemp.Fields("st55") Or _
'         strUserNum = "" & RsTemp.Fields("a0908") Then
'         bolReadY = True
'      End If
'   End If
'
'   txtCR(9).Text = ""
'   lstAtt.Clear
'   If PUB_GetST05(strUserNum) = "00" Or PUB_GetST05(strUserNum) = "01" Or PUB_GetST05(strUserNum) = "09" Or _
'      ((strCP01 = "L" Or strCP01 = "LA") And strST11 = "G1") Or _
'      ((strCP01 = "CFL" Or strCP01 = "FCL" Or strCP01 = "LIN") And strST11 = "F1") Or _
'      bolReadY = True Then
'
'      lstAtt.Enabled = True
'      cmdOpenAtt.Enabled = True
'
'      strSql = "SELECT * FROM CourtyardPeriod " & _
'                "WHERE CDP01 = '" & strCP09 & "' "
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount > 0 Then
'         QueryCourtyardPeriod = True
'         ' 附件檔名
'         If IsNull(rsTmp.Fields("cdp16")) = False Then
'            txtCR(9).Text = CheckStr(rsTmp.Fields("cdp16"))
'            SetList lstAtt, txtCR(9)
'         End If
'      End If
'      rsTmp.Close
'      Set rsTmp = Nothing
'   Else
'      lstAtt.Enabled = False
'      cmdOpenAtt.Enabled = False
'   End If
'End Function

Private Sub SetList(oList As ListBox, p_stList As String)
   Dim arrID
   oList.Clear
   If p_stList <> "" Then
      arrID = Split(p_stList, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

'Added by Lydia 2016/10/26 修正Win7 輸入法問題
Private Sub textCP64_GotFocus()
   TextInverse textCP64 'Added by Lydia 2021/10/08
   OpenIme
End Sub

'Added by Morgan 2014/5/22
Private Sub SetLetter()
   Dim rsQuery As ADODB.Recordset
   
   '信函進度資料
   cmdCancelLetter.Visible = False 'Added by Morgan 2019/3/15
   cmdCancelConfirm.Visible = False 'Added by Morgan 2020/2/10
   cmdCancelLP05.Visible = False 'Added by Morgan 2023/9/15
   'Modified by Morgan 2018/10/30 +lp10 is not null(有通知函才顯示相關欄位)
   'Modified by Morgan 2019/6/26 +已發文條件
   'Modified by Morgan 2020/1/14 +lp10,LP43 工程師判發要顯示
   'Modified by Sindy 2020/2/26 +,LP32,LP38 是否大宗發文,E化寄送人員
   'Modified by Morgan 2021/4/1 +dt6,已有l.* 刪除重複的LP欄位
   'Modified by Morgan 2022/5/20 +LP51,LP52
   strSql = "SELECT l.*,sqldatet(lp05)||' '||sqltime6(lp17) dt1,sqldatet(lp07)||' '||sqltime6(lp18) dt2,cp83,sqldatet(lp21) dt3,sqldatet(lp23) dt4" & _
      ",sqldatet(lp39)||' '||sqltime6(lp40) dt5,sqldatet(lp47)||' '||sqltime6(lp48) dt6,cp154,lp51,lp52,lp43 FROM caseprogress,letterprogress l" & _
      " WHERE CP09 = '" & m_CP09 & "' and lp01(+)=cp09 and lp10||lp43 is not null"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      m_strLP32 = "" & rsQuery.Fields("LP32") 'Add by Sindy 2020/2/26
      m_strLP38 = "" & rsQuery.Fields("LP38") 'Add by Sindy 2020/2/26
      m_strLP31 = "" & rsQuery.Fields("LP31") 'Add by Sindy 2020/7/24
      'Added by Morgan 2019/3/15
      '開放發文人員可以改不通知客戶
      'Modified by Morgan 2020/2/3 改同部門可改
      'If rsQuery.Fields("cp83") = strUserNum Or Pub_StrUserSt03 = "M51" Then
      If Pub_StrUserSt03 = PUB_GetST03("" & rsQuery.Fields("cp83")) Or Pub_StrUserSt03 = "M51" Then
         '有信函,發文室未發文,未確認且非直寄
         'Modified by Morgan 2019/12/13 +E化發文人為QPGMR也可改
         'Modified by Morgan 2024/12/27 偶而還是有直寄不通知狀況,故取消此限制 ex:P-120726(CB3075962)
         If rsQuery.Fields("lp10") = "Y" And (rsQuery.Fields("lp15") = "N" Or rsQuery.Fields("cp154") = "QPGMR") And rsQuery.Fields("lp07") = 0 Then
            cmdCancelLetter.Visible = True
         End If
      End If
      'end 2019/3/15
      
      'Added by Morgan 2014/10/23
      lblLP11 = ""
      If rsQuery.Fields("lp10") = "N" Then
         'Modified by Morgan 2020/2/10 +併函通知
         If Not IsNull(rsQuery.Fields("lp42")) Then
            lblLP11 = "併函通知"
         Else
            lblLP11 = "不通知"
         End If
      'end 2014/10/23
      ElseIf rsQuery.Fields("lp11") = "Y" Then
         lblLP11 = "直寄"
      ElseIf rsQuery.Fields("lp11") = "0" Then
         lblLP11 = "親送"
      ElseIf rsQuery.Fields("lp11") = "1" Then
         lblLP11 = "寄送"
      ElseIf rsQuery.Fields("lp11") = "2" Then
         lblLP11 = "不寄"
      End If
      
      'Added by Morgan 2015/6/26
      lblLP26 = ""
      If "" & rsQuery.Fields("lp10") <> "" Then 'Added by Morgan 2024/8/22 無通知函(純工程師判發)不要顯示,以免誤以為會通知
         If rsQuery.Fields("lp26") = "Y" Then
            lblLP26 = "E化"
         'Modified by Morgan 2022/6/30 調整顯示內容
         ElseIf rsQuery.Fields("lp26") = "E" Then
            lblLP26 = "全E化"
         End If
         If rsQuery.Fields("lp52") = "Y" Then
            If rsQuery.Fields("lp11") <> "2" And lblLP26 <> "" Then
               lblLP26 = lblLP26 & " (要確收)"
            End If
         End If
         'Added by Morgan 2025/9/12
         If rsQuery.Fields("lp32") = "Y" Then
            lblLP26 = lblLP26 & " (大宗)"
            If IsNull(rsQuery.Fields("lp28")) And rsQuery.Fields("lp15") = "N" Then
               lblLP26 = lblLP26 & " (未檢核)"
            End If
         End If
         'end 2025/9/12
      End If
      'end 2022/6/30
      'end 2015/6/26
      
      lblLP04 = "": lblLP0517 = ""
      'Modified by Morgan 2018/10/24 改有判發人就帶出
      If Not IsNull(rsQuery.Fields("lp04")) Then
         lblLP04 = rsQuery.Fields("lp04") & " " & GetStaffName(rsQuery.Fields("lp04"), True)
      End If
      'end 2018/10/24
      If rsQuery.Fields("lp05") > 0 Then
         'Modified by Morgan 2019/6/26 自判但未發文時不可帶判發時間 Ex:CFP-22592 期限通知(未轉報價定稿時)
         If Not IsNull(rsQuery.Fields("lp04")) Then
            'lblLP04 = rsQuery.Fields("lp04") & " " & GetStaffName(rsQuery.Fields("lp04"), True) 'Removed by Morgan 2018/10/24 改有判發人就帶出(上面)
            lblLP0517 = "" & rsQuery.Fields("dt1")
            'Added by Morgan 2023/9/15
            If Pub_StrUserSt03 = "M51" Then
               cmdCancelLP05.Visible = True
            End If
            'end 2023/9/15
         Else
            If Not IsNull(rsQuery.Fields("cp83")) Then
               lblLP04 = rsQuery.Fields("cp83") & " " & GetStaffName(rsQuery.Fields("cp83"), True)
               lblLP0517 = "" & rsQuery.Fields("dt1")
            End If
         End If
      End If
      
      lblLP06 = "": lblLP0718 = ""
      'Modified by Morgan 2020/1/15
      'If Not IsNull(rsQuery.Fields("lp06")) Then
      If Not IsNull(rsQuery.Fields("lp06")) And rsQuery.Fields("lp10") <> "" Then
         lblLP06 = rsQuery.Fields("lp06") & " " & GetStaffName(rsQuery.Fields("lp06"), True)
      End If
      If rsQuery.Fields("lp07") > 0 Then
         lblLP0718 = "" & rsQuery.Fields("dt2")
                  
         'Added by Morgan 2020/2/10
         'Modify By Sindy 2020/7/24 + 商標處未經發文室已確認,收件人為代理人
         'Modify By Sindy 2020/11/3 + 排除併函通知
         'Modified by Morgan 2021/1/11 + 發文人為 QPGMR 時也可取消(E化已確認改發函方式)
         'If (lblCP154 = "" Or _
             (Left(m_CP01, 1) = "T" And InStr(lblCP154, "QPGMR") > 0 And "" & rsQuery.Fields("lp31") = "Y") _
            ) And Pub_StrUserSt03 = "M51" And _
            IsNull(rsQuery.Fields("lp42")) Then
         If (lblCP154 = "" Or InStr(lblCP154, "QPGMR") > 0) And Pub_StrUserSt03 = "M51" And IsNull(rsQuery.Fields("lp42")) Then
            cmdCancelConfirm.Visible = True
         End If
         'end 2020/2/10
      End If
      txtLP12 = "" & rsQuery.Fields("lp12")
      
      'Added by Morgan 2019/5/14 判發人員意見(目前開放給判發人及程序看)
      If rsQuery.Fields("lp04") = strUserNum Or Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "P12" Then
         Label16(20).Visible = True
         txtLP37.Visible = True
         txtLP37 = "" & rsQuery.Fields("lp37")
      Else
         Label16(20).Visible = False
         txtLP37.Visible = False
      End If
      'end 2019/5/14
      
      'Added by Morgan 2015/1/14
      If rsQuery.Fields("lp21") > 0 Then
         lblLP20 = rsQuery.Fields("lp20") & " " & GetStaffName(rsQuery.Fields("lp20"), True)
         lblLP21 = "" & rsQuery.Fields("dt3")
      End If
      If rsQuery.Fields("lp23") > 0 Then
         lblLP22 = rsQuery.Fields("lp22") & " " & GetStaffName(rsQuery.Fields("lp22"), True)
         lblLP23 = "" & rsQuery.Fields("dt4")
      End If
      txtLP24 = "" & rsQuery.Fields("lp24")
      txtLP25 = "" & rsQuery.Fields("lp25")
      'end 2015/1/14
      
      'Added by Morgan 2019/4/23
      If rsQuery.Fields("lp39") > 0 Then
         lblLP3940 = "" & rsQuery.Fields("dt5")
         lblLP38 = rsQuery.Fields("lp38") & " " & GetStaffName(rsQuery.Fields("lp38"), True)
      Else
         If rsQuery.Fields("lp10") = "Y" And rsQuery.Fields("lp11") <> "2" Then 'Added by Morgan 2024/8/22 要通知才顯示,以免誤會要EMail
            If Not IsNull(rsQuery.Fields("lp38")) Then
               lblLP38 = rsQuery.Fields("lp38") & " " & GetStaffName(rsQuery.Fields("lp38"), True)
            End If
         End If
      End If
      'end 2019/4/23
      
      'Added by Morgan 2021/4/1
      If Not IsNull(rsQuery.Fields("lp46")) Then
         lblLP46 = rsQuery.Fields("lp46") & " " & GetStaffName(rsQuery.Fields("lp46"), True)
      End If
      If rsQuery.Fields("lp47") > 0 Then
         lblLP4748 = "" & rsQuery.Fields("dt6")
      End If
      txtLP49 = "" & rsQuery.Fields("lp49")
      'end 2021/4/1
   End If
   'end 2014/5/22
   
   'Added by Morgan 2016/5/18
   '指示信
   strSql = "SELECT a.*,sqldatet(af07)||' '||sqltime6(af08) dt1,sqldatet(af11)||' '||sqltime6(af12) dt2 FROM caseprogress,appform a" & _
      " WHERE CP09 = '" & m_CP09 & "' and af01(+)=cp09 and af06 is not null"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      SSTab2.TabVisible(1) = True
      lblAF0708 = "" & rsQuery("dt1")
      lblAF1112 = "" & rsQuery("dt2")
      If Not IsNull(rsQuery.Fields("af06")) And rsQuery.Fields("af07") > 0 Then
         lblAF06 = rsQuery.Fields("af06") & " " & GetStaffName(rsQuery.Fields("af06"), True)
      End If
      If Not IsNull(rsQuery.Fields("af14")) Then
         lblAF14 = rsQuery.Fields("af14") & " " & GetStaffName(rsQuery.Fields("af14"), True)
      End If
      'Add By Sindy 2020/7/21 內商的指示信都在歷程判發,此處不需顯示
      If Left(m_CP01, 1) = "T" Then
         Label16(17).Visible = False
         lblAF06.Visible = False
         lblAF0708.Visible = False
      End If
      '2020/7/21 END
   Else
      SSTab2.TabVisible(1) = False
   End If
   'end 2016/5/18
   
   Set rsQuery = Nothing
End Sub

'Added by Lydia 2021/10/08 因為Form 2.0元件需要Focus到欄位，才能顯示捲軸
Private Sub textCP144_GotFocus()
   TextInverse textCP144
End Sub
'Added by Lydia 2021/10/08
Private Sub textCP49_GotFocus()
   TextInverse textCP49
End Sub

'Added by Lydia 2021/10/08
Private Sub textCP37_GotFocus()
   TextInverse textCP37
End Sub

'Added by Lydia 2021/10/08
Private Sub textCP37_1_GotFocus()
   TextInverse textCP37_1
End Sub

'Added by Lydia 2021/10/08
Private Sub textCP38_GotFocus()
   TextInverse textCP38
End Sub

'Added by Lydia 2021/10/08
Private Sub textCP39_GotFocus()
   TextInverse textCP39
End Sub

'Added by Lydia 2021/10/08
Private Sub textCP40_GotFocus()
   TextInverse textCP40
End Sub

'Added by Lydia 2021/10/08
Private Sub textCP41_GotFocus()
   TextInverse textCP41
End Sub

'Added by Lydia 2021/10/08
Private Sub textCP42_GotFocus()
   TextInverse textCP42
End Sub
'Added by Lydia 2021/10/08
Private Sub textCP50_GotFocus()
   TextInverse textCP50
End Sub

'Added by Lydia 2021/10/08
Private Sub textCP51_GotFocus()
   TextInverse textCP51
End Sub

'Added by Lydia 2021/10/08
Private Sub textCP52_GotFocus()
   TextInverse textCP52
End Sub

'Added by Lydia 2021/10/08
Private Sub textCP131_GotFocus()
   TextInverse textCP131
End Sub

'Added by Lydia 2021/10/08
Private Sub txtTF37_GotFocus()
   TextInverse txtTF37
End Sub
'Added by Lydia 2021/10/08
Private Sub txtLP37_GotFocus()
   TextInverse txtLP37
End Sub

'Added by Lydia 2021/10/08
Private Sub txtLP12_GotFocus()
   TextInverse txtLP12
End Sub

'Added by Lydia 2021/10/08
Private Sub txtLP49_GotFocus()
   TextInverse txtLP49
End Sub

'Added by Lydia 2021/10/08
Private Sub txtLP24_GotFocus()
   TextInverse txtLP24
End Sub

'Added by Lydia 2021/10/08
Private Sub txtLP25_GotFocus()
   TextInverse txtLP25
End Sub

'Add by Amy 2022/09/02 置換 Label文字,「其他相關人」顯示「對方」
Private Sub SetLabTxt(stRPTxt As String)
    Dim oLab
    
    If Left(Label18(4).Caption, 2) = stRPTxt Then Exit Sub
    
    For Each oLab In Label18
        If oLab.Index >= 1 And oLab.Index <= 5 Then
            oLab.Caption = stRPTxt & Mid(oLab.Caption, 3)
        End If
    Next
    
    For Each oLab In Label19
        If oLab.Index >= 1 And oLab.Index <= 3 Then
            oLab.Caption = stRPTxt & Mid(oLab.Caption, 3)
        End If
    Next
    Label33(1).Caption = stRPTxt & Mid(Label33(1).Caption, 3)
End Sub


