VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030202_05 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(變更事項)"
   ClientHeight    =   5790
   ClientLeft      =   5325
   ClientTop       =   2460
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8535
   Begin VB.CommandButton cmdCopyCaseno 
      Caption         =   "從右邊本所案號複製過來"
      Height          =   285
      Left            =   90
      TabIndex        =   187
      Top             =   30
      Width           =   2385
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   3
      Left            =   4170
      MaxLength       =   2
      TabIndex        =   97
      Top             =   30
      Width           =   285
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   2
      Left            =   3795
      MaxLength       =   1
      TabIndex        =   96
      Top             =   30
      Width           =   180
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   1
      Left            =   3015
      MaxLength       =   6
      TabIndex        =   95
      Top             =   30
      Width           =   645
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Index           =   0
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   94
      Text            =   "FCT"
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "複製(&C)"
      Height          =   400
      Left            =   5610
      TabIndex        =   186
      Top             =   30
      Width           =   912
   End
   Begin VB.TextBox textCE01 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   480
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   480
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7500
      TabIndex        =   99
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6555
      TabIndex        =   98
      Top             =   30
      Width           =   912
   End
   Begin TabDlg.SSTab tabCtrl 
      Height          =   4965
      Left            =   60
      TabIndex        =   100
      Top             =   780
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   8758
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm030202_05.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label19"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label17"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label14"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label16"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label18"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label20"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label23"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label27"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textCE04_2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textCE17"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCE18"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textCE19"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textCE20"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textCE21"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCE05_2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCE06_2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCE07_2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCE08_2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "checkCE03"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "checkCE09"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCE04"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCE02"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCE05"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCE06"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCE07"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCE08"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "checkCE56"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "checkCE54"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "checkCE52"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "checkCE22"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).ControlCount=   36
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm030202_05.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label22"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label21"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label29"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label31"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label33"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label37"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label39"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label41"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label43"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label45"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "textCE63"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "textCE64"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "textCE92"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "textCE93"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "textCE94"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "textCE95"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "textCE96"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "textCE97"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "textCE98"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "textCE99"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "checkCE65"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "第三頁"
      TabPicture(2)   =   "frm030202_05.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label63"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label62"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label61"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label60"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label59"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label58"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label57"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label56"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label55"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label54"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label53"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label52"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label51"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label50"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label49"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label48"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label47"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label46"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label6"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Label7"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Label8"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Label9"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Label10"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Label11"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Label64"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Label65"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Label66"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Label67"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Label68"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Label69"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "textCE10"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "textCE11"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "textCE12"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "textCE13"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "textCE14"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "textCE15"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "textCE68"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "textCE69"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "textCE70"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "textCE71"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "textCE72"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "textCE73"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "textCE74"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "textCE75"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "textCE76"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "textCE77"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "textCE78"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "textCE79"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "textCE80"
      Tab(2).Control(48).Enabled=   0   'False
      Tab(2).Control(49)=   "textCE81"
      Tab(2).Control(49).Enabled=   0   'False
      Tab(2).Control(50)=   "textCE82"
      Tab(2).Control(50).Enabled=   0   'False
      Tab(2).Control(51)=   "textCE83"
      Tab(2).Control(51).Enabled=   0   'False
      Tab(2).Control(52)=   "textCE84"
      Tab(2).Control(52).Enabled=   0   'False
      Tab(2).Control(53)=   "textCE85"
      Tab(2).Control(53).Enabled=   0   'False
      Tab(2).Control(54)=   "textCE86"
      Tab(2).Control(54).Enabled=   0   'False
      Tab(2).Control(55)=   "textCE87"
      Tab(2).Control(55).Enabled=   0   'False
      Tab(2).Control(56)=   "textCE88"
      Tab(2).Control(56).Enabled=   0   'False
      Tab(2).Control(57)=   "textCE89"
      Tab(2).Control(57).Enabled=   0   'False
      Tab(2).Control(58)=   "textCE90"
      Tab(2).Control(58).Enabled=   0   'False
      Tab(2).Control(59)=   "textCE91"
      Tab(2).Control(59).Enabled=   0   'False
      Tab(2).Control(60)=   "checkCE16"
      Tab(2).Control(60).Enabled=   0   'False
      Tab(2).ControlCount=   61
      TabCaption(3)   =   "第四頁"
      TabPicture(3)   =   "frm030202_05.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label75"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label74"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label73"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label72"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label71"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label70"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label24"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label25"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label26"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label76"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label77"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label78"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label79"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Label80"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Label81"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "textCE23"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "textCE24"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "textCE25"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "textCE26"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "textCE27"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "textCE28"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "textCE29"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "textCE30"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "textCE31"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "textCE32"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "textCE33"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "textCE34"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "textCE35"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "textCE36"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "textCE37"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "checkCE38"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).ControlCount=   31
      TabCaption(4)   =   "第五頁"
      TabPicture(4)   =   "frm030202_05.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label44"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label42"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label40"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label38"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label36"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label35"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label34"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label32"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Label30"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Label28"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "textCE41"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "textCE42"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "textCE43"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "textCE45"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "textCE41_1"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "textCE61"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "textCE49"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "textCE47"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "textCE39"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "textCE57"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "checkCE58"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "checkCE40"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "checkCE60"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "checkCE44"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "checkCE46"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "checkCE48"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "checkCE50"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "checkCE62"
      Tab(4).Control(28).Enabled=   0   'False
      Tab(4).Control(29)=   "textCE39_2"
      Tab(4).Control(29).Enabled=   0   'False
      Tab(4).Control(30)=   "cmdGoods"
      Tab(4).Control(30).Enabled=   0   'False
      Tab(4).ControlCount=   31
      Begin VB.CommandButton cmdGoods 
         Caption         =   "商品名稱"
         Height          =   315
         Left            =   -74610
         TabIndex        =   87
         Top             =   2670
         Width           =   1005
      End
      Begin VB.CheckBox checkCE16 
         Height          =   180
         Left            =   -74940
         TabIndex        =   28
         Top             =   300
         Width           =   252
      End
      Begin VB.CheckBox checkCE38 
         Height          =   180
         Left            =   -74730
         TabIndex        =   59
         Top             =   432
         Width           =   252
      End
      Begin VB.TextBox textCE39_2 
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   -71970
         Locked          =   -1  'True
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   855
         Width           =   5052
      End
      Begin VB.CheckBox checkCE62 
         Height          =   180
         Left            =   -74850
         TabIndex        =   92
         Top             =   3682
         Width           =   252
      End
      Begin VB.CheckBox checkCE50 
         Height          =   180
         Left            =   -74850
         TabIndex        =   90
         Top             =   3382
         Width           =   252
      End
      Begin VB.CheckBox checkCE48 
         Height          =   180
         Left            =   -74850
         TabIndex        =   88
         Top             =   3082
         Width           =   252
      End
      Begin VB.CheckBox checkCE46 
         Height          =   180
         Left            =   -74850
         TabIndex        =   85
         Top             =   2430
         Width           =   252
      End
      Begin VB.CheckBox checkCE44 
         Height          =   180
         Left            =   -74850
         TabIndex        =   80
         Top             =   1470
         Width           =   252
      End
      Begin VB.CheckBox checkCE60 
         Height          =   180
         Left            =   -74850
         TabIndex        =   79
         Top             =   1170
         Width           =   252
      End
      Begin VB.CheckBox checkCE40 
         Height          =   180
         Left            =   -74850
         TabIndex        =   77
         Top             =   870
         Width           =   252
      End
      Begin VB.CheckBox checkCE58 
         Height          =   180
         Left            =   -74850
         TabIndex        =   75
         Top             =   570
         Width           =   252
      End
      Begin VB.TextBox textCE57 
         Height          =   285
         Left            =   -73290
         MaxLength       =   20
         TabIndex        =   76
         Top             =   555
         Width           =   1212
      End
      Begin VB.TextBox textCE39 
         Height          =   285
         Left            =   -73290
         MaxLength       =   1
         TabIndex        =   78
         Top             =   855
         Width           =   1212
      End
      Begin VB.TextBox textCE47 
         Height          =   285
         Left            =   -73290
         MaxLength       =   395
         TabIndex        =   89
         Top             =   3030
         Width           =   6372
      End
      Begin VB.TextBox textCE49 
         Height          =   285
         Left            =   -73290
         MaxLength       =   699
         TabIndex        =   91
         Top             =   3330
         Width           =   6372
      End
      Begin VB.CheckBox checkCE22 
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   3060
         Width           =   252
      End
      Begin VB.CheckBox checkCE52 
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   252
      End
      Begin VB.CheckBox checkCE54 
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   2460
         Width           =   252
      End
      Begin VB.CheckBox checkCE56 
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   252
      End
      Begin VB.TextBox textCE08 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   5
         Top             =   1530
         Width           =   1212
      End
      Begin VB.TextBox textCE07 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   4
         Top             =   1230
         Width           =   1212
      End
      Begin VB.TextBox textCE06 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   3
         Top             =   930
         Width           =   1212
      End
      Begin VB.TextBox textCE05 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   2
         Top             =   630
         Width           =   1212
      End
      Begin VB.CheckBox checkCE65 
         Height          =   180
         Left            =   -74880
         TabIndex        =   17
         Top             =   330
         Width           =   252
      End
      Begin VB.TextBox textCE02 
         Height          =   285
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   7
         Top             =   1830
         Width           =   1212
      End
      Begin VB.TextBox textCE04 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   1
         Top             =   330
         Width           =   1212
      End
      Begin VB.CheckBox checkCE09 
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   372
         Width           =   252
      End
      Begin VB.CheckBox checkCE03 
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   1872
         Width           =   252
      End
      Begin MSForms.TextBox textCE61 
         Height          =   645
         Left            =   -73290
         TabIndex        =   93
         Top             =   3630
         Width           =   6375
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         Size            =   "11245;1138"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41_1 
         Height          =   915
         Left            =   -73290
         TabIndex        =   82
         Top             =   1470
         Width           =   6375
         VariousPropertyBits=   -1475330021
         ScrollBars      =   2
         Size            =   "11245;1614"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE45 
         Height          =   285
         Left            =   -73290
         TabIndex        =   86
         Top             =   2415
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE43 
         Height          =   285
         Left            =   -73290
         TabIndex        =   84
         Top             =   2070
         Width           =   6372
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "11239;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE42 
         Height          =   285
         Left            =   -73290
         TabIndex        =   83
         Top             =   1770
         Width           =   6372
         VariousPropertyBits=   671105051
         MaxLength       =   180
         Size            =   "11239;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41 
         Height          =   285
         Left            =   -73290
         TabIndex        =   81
         Top             =   1470
         Width           =   6372
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "11239;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE08_2 
         Height          =   285
         Left            =   2970
         TabIndex        =   189
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5055
         VariousPropertyBits=   671105055
         Size            =   "8916;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE07_2 
         Height          =   285
         Left            =   2970
         TabIndex        =   192
         TabStop         =   0   'False
         Top             =   1245
         Width           =   5055
         VariousPropertyBits=   671105055
         Size            =   "8916;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE06_2 
         Height          =   285
         Left            =   2970
         TabIndex        =   191
         TabStop         =   0   'False
         Top             =   930
         Width           =   5055
         VariousPropertyBits=   671105055
         Size            =   "8916;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE05_2 
         Height          =   285
         Left            =   2970
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   615
         Width           =   5055
         VariousPropertyBits=   671105055
         Size            =   "8916;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE37 
         Height          =   285
         Left            =   -73170
         TabIndex        =   74
         Top             =   4620
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE36 
         Height          =   285
         Left            =   -73170
         TabIndex        =   73
         Top             =   4320
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE35 
         Height          =   285
         Left            =   -73170
         TabIndex        =   72
         Top             =   4020
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE34 
         Height          =   285
         Left            =   -73170
         TabIndex        =   71
         Top             =   3720
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE33 
         Height          =   285
         Left            =   -73170
         TabIndex        =   70
         Top             =   3420
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE32 
         Height          =   285
         Left            =   -73170
         TabIndex        =   69
         Top             =   3120
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE31 
         Height          =   285
         Left            =   -73170
         TabIndex        =   68
         Top             =   2820
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE30 
         Height          =   285
         Left            =   -73170
         TabIndex        =   67
         Top             =   2520
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE29 
         Height          =   285
         Left            =   -73170
         TabIndex        =   66
         Top             =   2220
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE28 
         Height          =   285
         Left            =   -73170
         TabIndex        =   65
         Top             =   1920
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE27 
         Height          =   285
         Left            =   -73170
         TabIndex        =   64
         Top             =   1620
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE26 
         Height          =   285
         Left            =   -73170
         TabIndex        =   63
         Top             =   1320
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE25 
         Height          =   285
         Left            =   -73170
         TabIndex        =   62
         Top             =   1020
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE24 
         Height          =   285
         Left            =   -73170
         TabIndex        =   61
         Top             =   720
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE23 
         Height          =   285
         Left            =   -73170
         TabIndex        =   60
         Top             =   420
         Width           =   6405
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "11289;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE91 
         Height          =   285
         Left            =   -69630
         TabIndex        =   58
         Top             =   4470
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE90 
         Height          =   285
         Left            =   -69630
         TabIndex        =   57
         Top             =   4161
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE89 
         Height          =   285
         Left            =   -69630
         TabIndex        =   56
         Top             =   3864
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE88 
         Height          =   285
         Left            =   -69630
         TabIndex        =   55
         Top             =   3567
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE87 
         Height          =   285
         Left            =   -69630
         TabIndex        =   54
         Top             =   3270
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE86 
         Height          =   285
         Left            =   -69630
         TabIndex        =   53
         Top             =   2973
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE85 
         Height          =   285
         Left            =   -69630
         TabIndex        =   52
         Top             =   2676
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE84 
         Height          =   285
         Left            =   -69630
         TabIndex        =   51
         Top             =   2379
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE83 
         Height          =   285
         Left            =   -69630
         TabIndex        =   50
         Top             =   2082
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE82 
         Height          =   285
         Left            =   -69630
         TabIndex        =   49
         Top             =   1785
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE81 
         Height          =   285
         Left            =   -69630
         TabIndex        =   48
         Top             =   1488
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE80 
         Height          =   285
         Left            =   -69630
         TabIndex        =   47
         Top             =   1191
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE79 
         Height          =   285
         Left            =   -69630
         TabIndex        =   46
         Top             =   894
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE78 
         Height          =   285
         Left            =   -69630
         TabIndex        =   45
         Top             =   597
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE77 
         Height          =   285
         Left            =   -69630
         TabIndex        =   44
         Top             =   300
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE76 
         Height          =   285
         Left            =   -73650
         TabIndex        =   43
         Top             =   4470
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE75 
         Height          =   285
         Left            =   -73650
         TabIndex        =   42
         Top             =   4161
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE74 
         Height          =   285
         Left            =   -73650
         TabIndex        =   41
         Top             =   3864
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE73 
         Height          =   285
         Left            =   -73650
         TabIndex        =   40
         Top             =   3567
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE72 
         Height          =   285
         Left            =   -73650
         TabIndex        =   39
         Top             =   3270
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE71 
         Height          =   285
         Left            =   -73650
         TabIndex        =   38
         Top             =   2973
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE70 
         Height          =   285
         Left            =   -73650
         TabIndex        =   37
         Top             =   2676
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE69 
         Height          =   285
         Left            =   -73650
         TabIndex        =   36
         Top             =   2379
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE68 
         Height          =   285
         Left            =   -73650
         TabIndex        =   35
         Top             =   2082
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE15 
         Height          =   285
         Left            =   -73650
         TabIndex        =   34
         Top             =   1785
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE14 
         Height          =   285
         Left            =   -73650
         TabIndex        =   33
         Top             =   1488
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE13 
         Height          =   285
         Left            =   -73650
         TabIndex        =   32
         Top             =   1191
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE12 
         Height          =   285
         Left            =   -73650
         TabIndex        =   31
         Top             =   894
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE11 
         Height          =   285
         Left            =   -73650
         TabIndex        =   30
         Top             =   597
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE10 
         Height          =   285
         Left            =   -73650
         TabIndex        =   29
         Top             =   300
         Width           =   2925
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5159;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE99 
         Height          =   285
         Left            =   -73200
         TabIndex        =   27
         Top             =   3000
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE98 
         Height          =   285
         Left            =   -73200
         TabIndex        =   26
         Top             =   2700
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE97 
         Height          =   285
         Left            =   -73200
         TabIndex        =   25
         Top             =   2400
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE96 
         Height          =   285
         Left            =   -73200
         TabIndex        =   24
         Top             =   2100
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE95 
         Height          =   285
         Left            =   -73200
         TabIndex        =   23
         Top             =   1800
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE94 
         Height          =   285
         Left            =   -73200
         TabIndex        =   22
         Top             =   1500
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE93 
         Height          =   285
         Left            =   -73200
         TabIndex        =   21
         Top             =   1200
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE92 
         Height          =   285
         Left            =   -73200
         TabIndex        =   20
         Top             =   900
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE64 
         Height          =   285
         Left            =   -73200
         TabIndex        =   19
         Top             =   600
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE63 
         Height          =   285
         Left            =   -73200
         TabIndex        =   18
         Top             =   300
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE21 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   4320
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE20 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   3996
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE19 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   3674
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE18 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   3352
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE17 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   3030
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE04_2 
         Height          =   285
         Left            =   2970
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   300
         Width           =   5052
         VariousPropertyBits=   671105055
         Size            =   "8911;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(中) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   185
         Top             =   432
         Width           =   1200
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(英) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   184
         Top             =   734
         Width           =   1200
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(日) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   183
         Top             =   1036
         Width           =   1200
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(中) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   182
         Top             =   1338
         Width           =   1200
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(英) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   181
         Top             =   1640
         Width           =   1200
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(日) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   180
         Top             =   1942
         Width           =   1200
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(中) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   179
         Top             =   2244
         Width           =   1200
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(英) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   178
         Top             =   2546
         Width           =   1200
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(日) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   177
         Top             =   2848
         Width           =   1200
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(中) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   176
         Top             =   3150
         Width           =   1200
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(英) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   175
         Top             =   3452
         Width           =   1200
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(日) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   174
         Top             =   3754
         Width           =   1200
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(中) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   173
         Top             =   4056
         Width           =   1200
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(英) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   172
         Top             =   4358
         Width           =   1200
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(日) :"
         Height          =   180
         Left            =   -74430
         TabIndex        =   171
         Top             =   4672
         Width           =   1200
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "代表人1(中) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   170
         Top             =   330
         Width           =   1020
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "代表人1(英) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   169
         Top             =   629
         Width           =   1020
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "代表人1(日) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   168
         Top             =   928
         Width           =   1020
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "代表人2(中) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   167
         Top             =   1227
         Width           =   1020
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "代表人2(英) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   166
         Top             =   1526
         Width           =   1020
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "代表人2(日) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   165
         Top             =   1825
         Width           =   1020
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "代表人3(中) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   164
         Top             =   2124
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "代表人3(英) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   163
         Top             =   2423
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "代表人3(日) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   162
         Top             =   2722
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "代表人4(中) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   161
         Top             =   3021
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "代表人4(英) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   160
         Top             =   3320
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "代表人4(日) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   159
         Top             =   3619
         Width           =   1020
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "代表人5(中) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   158
         Top             =   3918
         Width           =   1020
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "代表人5(英) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   157
         Top             =   4217
         Width           =   1020
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "代表人5(日) :"
         Height          =   180
         Left            =   -74700
         TabIndex        =   156
         Top             =   4522
         Width           =   1020
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "代表人6(中) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   155
         Top             =   330
         Width           =   1020
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "代表人6(英) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   154
         Top             =   629
         Width           =   1020
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "代表人6(日) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   153
         Top             =   928
         Width           =   1020
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "代表人7(中) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   152
         Top             =   1227
         Width           =   1020
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "代表人7(英) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   151
         Top             =   1526
         Width           =   1020
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "代表人7(日) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   150
         Top             =   1825
         Width           =   1020
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "代表人8(中) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   149
         Top             =   2124
         Width           =   1020
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "代表人8(英) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   148
         Top             =   2423
         Width           =   1020
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "代表人8(日) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   147
         Top             =   2722
         Width           =   1020
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "代表人9(中) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   146
         Top             =   3021
         Width           =   1020
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "代表人9(英) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   145
         Top             =   3320
         Width           =   1020
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "代表人9(日) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   144
         Top             =   3619
         Width           =   1020
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "代表人10(中) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   143
         Top             =   3918
         Width           =   1110
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "代表人10(英) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   142
         Top             =   4217
         Width           =   1110
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "代表人10(日) :"
         Height          =   180
         Left            =   -70680
         TabIndex        =   141
         Top             =   4522
         Width           =   1110
      End
      Begin VB.Label Label45 
         Caption         =   "代表人9中譯文 :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   140
         Top             =   2730
         Width           =   1335
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "代表人10中譯文 :"
         Height          =   180
         Left            =   -74640
         TabIndex        =   139
         Top             =   3030
         Width           =   1350
      End
      Begin VB.Label Label41 
         Caption         =   "代表人7中譯文 :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   138
         Top             =   2130
         Width           =   1335
      End
      Begin VB.Label Label39 
         Caption         =   "代表人8中譯文 :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   137
         Top             =   2430
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "正商標號數 :"
         Height          =   255
         Left            =   -74610
         TabIndex        =   136
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "商標種類 :"
         Height          =   255
         Left            =   -74610
         TabIndex        =   135
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "圖樣 :"
         Height          =   255
         Left            =   -74610
         TabIndex        =   134
         Top             =   1170
         Width           =   615
      End
      Begin VB.Label Label34 
         Caption         =   "案件名稱(中) :"
         Height          =   255
         Left            =   -74580
         TabIndex        =   133
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label Label35 
         Caption         =   "案件名稱(英) :"
         Height          =   255
         Left            =   -74610
         TabIndex        =   132
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "案件名稱(日) :"
         Height          =   255
         Left            =   -74610
         TabIndex        =   131
         Top             =   2070
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "縮減商品 :"
         Height          =   255
         Left            =   -74580
         TabIndex        =   130
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "商品類別 :"
         Height          =   255
         Left            =   -74610
         TabIndex        =   129
         Top             =   3045
         Width           =   1215
      End
      Begin VB.Label Label42 
         Caption         =   "商品群組 :"
         Height          =   255
         Left            =   -74610
         TabIndex        =   128
         Top             =   3345
         Width           =   1215
      End
      Begin VB.Label Label44 
         Caption         =   "其它 :"
         Height          =   255
         Left            =   -74610
         TabIndex        =   127
         Top             =   3645
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "案件名稱 :"
         Height          =   255
         Left            =   -74610
         TabIndex        =   126
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label Label37 
         Caption         =   "代表人5中譯文 :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   124
         Top             =   1530
         Width           =   1335
      End
      Begin VB.Label Label33 
         Caption         =   "代表人6中譯文 :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   123
         Top             =   1830
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "代表人3中譯文 :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   122
         Top             =   930
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "代表人4中譯文 :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   121
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "申請人5 :"
         Height          =   180
         Left            =   360
         TabIndex        =   120
         Top             =   1572
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "申請人4 :"
         Height          =   180
         Left            =   360
         TabIndex        =   119
         Top             =   1272
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "申請人3 :"
         Height          =   180
         Left            =   360
         TabIndex        =   118
         Top             =   972
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "申請人2 :"
         Height          =   180
         Left            =   360
         TabIndex        =   117
         Top             =   672
         Width           =   720
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "申請人5中譯文 :"
         Height          =   180
         Left            =   360
         TabIndex        =   116
         Top             =   4320
         Width           =   1260
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "申請人4中譯文 :"
         Height          =   180
         Left            =   360
         TabIndex        =   115
         Top             =   4005
         Width           =   1260
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "申請人3中譯文 :"
         Height          =   180
         Left            =   360
         TabIndex        =   114
         Top             =   3690
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "申請人2中譯文 :"
         Height          =   180
         Left            =   360
         TabIndex        =   113
         Top             =   3375
         Width           =   1260
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代理人 :"
         Height          =   180
         Left            =   360
         TabIndex        =   112
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "代表人印鑑 :"
         Height          =   180
         Left            =   360
         TabIndex        =   111
         Top             =   2460
         Width           =   990
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人印鑑 :"
         Height          =   180
         Left            =   360
         TabIndex        =   110
         Top             =   2760
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "申請人1中譯文 :"
         Height          =   180
         Left            =   360
         TabIndex        =   109
         Top             =   3060
         Width           =   1260
      End
      Begin VB.Label Label21 
         Caption         =   "代表人1中譯文 :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   108
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "代表人2中譯文 :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   107
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請日 :"
         Height          =   180
         Left            =   360
         TabIndex        =   104
         Top             =   1872
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "申請人1 :"
         Height          =   180
         Left            =   360
         TabIndex        =   103
         Top             =   375
         Width           =   720
      End
   End
   Begin VB.Line Line1 
      X1              =   4290
      X2              =   2760
      Y1              =   150
      Y2              =   150
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   106
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   105
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frm030202_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/02 改成Form2.0 ; textCE04_2、textCE05_2、textCE06_2、textCE07_2、textCE08_2、textCE17~21、textCE63~64、textCE92~99、textCE10~15、textCE68~91、textCE23~37、textCE41~43、textCE45、textCE61
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 收文號
Dim m_CE01 As String
'Add By Sindy 2009/06/03
Dim m_TM23 As String
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
Dim m_CP27 As String
'2009/06/03 End

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_CEList() As FIELDITEM
Dim m_CECount As Integer

' 更新商標基本檔時所使用的變更事項檔欄位的暫存資料
' 申請日
Dim m_CE02 As String
' 申請人
Dim m_CE04 As String
'add by nickc 2007/01/29
Dim m_CE05 As String
Dim m_CE06 As String
Dim m_CE07 As String
Dim m_CE08 As String

' 商品種類代碼
Dim m_CE39 As String
' 前畫面
'Dim m_PrevForm As Form
Dim m_Parent As String
'Add By Cheng 2003/04/10
Private Type TMFIELDITEM
   tiName As String
   tiData As String
   tiType As String
End Type
Dim m_TMList() As TMFIELDITEM
Dim m_TMListCount As Integer
Private Type SRFIELDITEM
   siName As String
   siData As String
   siType As String
End Type
Dim m_SRList() As SRFIELDITEM
Dim m_SRListCount As Integer
Dim tmpOldCE04 As String
'add by nickc 2007/01/25
Dim tmpOldCE05 As String
Dim tmpOldCE06 As String
Dim tmpOldCE07 As String
Dim tmpOldCE08 As String

Dim tmpOldCE02 As String
Dim tmpOldCE10 As String
Dim tmpOldCE11 As String
Dim tmpOldCE12 As String
Dim tmpOldCE13 As String
Dim tmpOldCE14 As String
Dim tmpOldCE15 As String
'add by nickc2007/01/25
Dim tmpOldCE26 As String
Dim tmpOldCE27 As String
Dim tmpOldCE28 As String
Dim tmpOldCE29 As String
Dim tmpOldCE30 As String
Dim tmpOldCE31 As String
Dim tmpOldCE32 As String
Dim tmpOldCE33 As String
Dim tmpOldCE34 As String
Dim tmpOldCE35 As String
Dim tmpOldCE36 As String
Dim tmpOldCE37 As String
Dim tmpOldCE68 As String
Dim tmpOldCE69 As String
Dim tmpOldCE70 As String
Dim tmpOldCE71 As String
Dim tmpOldCE72 As String
Dim tmpOldCE73 As String
Dim tmpOldCE74 As String
Dim tmpOldCE75 As String
Dim tmpOldCE76 As String
Dim tmpOldCE77 As String
Dim tmpOldCE78 As String
Dim tmpOldCE79 As String
Dim tmpOldCE80 As String
Dim tmpOldCE81 As String
Dim tmpOldCE82 As String
Dim tmpOldCE83 As String
Dim tmpOldCE84 As String
Dim tmpOldCE85 As String
Dim tmpOldCE86 As String
Dim tmpOldCE87 As String
Dim tmpOldCE88 As String
Dim tmpOldCE89 As String
Dim tmpOldCE90 As String
Dim tmpOldCE91 As String

Dim tmpOldCE23 As String
Dim tmpOldCE24 As String
Dim tmpOldCE25 As String
Dim tmpOldCE39 As String
Dim tmpOldCE41 As String
Dim tmpOldCE42 As String
Dim tmpOldCE43 As String
Dim tmpOldCE47 As String
Dim tmpOldCE49 As String
Dim tmpOldCE57 As String
'add by nickc 2007/04/03
Dim m_TM09 As String
Public ChkTG As Boolean
Dim m_CP31 As String 'Add By Sindy 2011/8/23 是否新案件
Dim m_CP10 As String 'Add By Sindy 2018/2/1 案件性質


' 檢查該欄位是否存在
Private Function IsCEFieldExist(ByVal strField As String) As Boolean
   Dim nIndex As Integer
   IsCEFieldExist = False
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         IsCEFieldExist = True
         Exit For
      End If
   Next nIndex
End Function

' 設定欄位新值
Private Sub SetCEFieldData(ByVal strField As String, ByVal strNewData As String, ByVal nType As Integer)
   Dim bFind As Boolean
   Dim nIndex As Integer
   bFind = False
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         bFind = True
         m_CEList(nIndex).fiNewData = strNewData
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_CEList(m_CECount + 1)
      m_CEList(m_CECount).fiName = strField
      m_CEList(m_CECount).fiNewData = strNewData
      m_CEList(m_CECount).fiType = nType
      m_CECount = m_CECount + 1
   End If
End Sub

' 清除欄位串列
Private Sub ClearCEFields()
   Erase m_CEList
   m_CECount = 0
End Sub

' 更新欄位內容
Private Sub UpdateFieldNewData()
   SetCEFieldData "CE01", m_CE01, 0
   If checkCE03.Value = 1 Then
      SetCEFieldData "CE02", DBDATE(textCE02), 1
   End If
   If checkCE09.Value = 1 Then
      If Trim(textCE04) <> "" Then
         'Modify By Sindy 2013/1/22
         'SetCEFieldData "CE04", textCE04 & String(9 - Len(textCE04), "0"), 0
         textCE04 = textCE04 & String(9 - Len(textCE04), "0")
         SetCEFieldData "CE04", textCE04, 0
         '2013/1/22 End
      End If
      'add by nickc 2007/01/26
      If Trim(textCE05) <> "" Then
         'Modify By Sindy 2013/1/22
         'SetCEFieldData "CE05", textCE05 & String(9 - Len(textCE05), "0"), 0
         textCE05 = textCE05 & String(9 - Len(textCE05), "0")
         SetCEFieldData "CE05", textCE05, 0
         '2013/1/22 End
      End If
      If Trim(textCE06) <> "" Then
         'Modify By Sindy 2013/1/22
         'SetCEFieldData "CE06", textCE06 & String(9 - Len(textCE06), "0"), 0
         textCE06 = textCE06 & String(9 - Len(textCE06), "0")
         SetCEFieldData "CE06", textCE06, 0
         '2013/1/22 End
      End If
      If Trim(textCE07) <> "" Then
         'Modify By Sindy 2013/1/22
         'SetCEFieldData "CE07", textCE07 & String(9 - Len(textCE07), "0"), 0
         textCE07 = textCE07 & String(9 - Len(textCE07), "0")
         SetCEFieldData "CE07", textCE07, 0
         '2013/1/22 End
      End If
      If Trim(textCE08) <> "" Then
         'Modify By Sindy 2013/1/22
         'SetCEFieldData "CE08", textCE08 & String(9 - Len(textCE08), "0"), 0
         textCE08 = textCE08 & String(9 - Len(textCE08), "0")
         SetCEFieldData "CE08", textCE08, 0
         '2013/1/22 End
      End If
   End If
   If checkCE16.Value = 1 Then
      SetCEFieldData "CE10", textCE10, 0
      SetCEFieldData "CE11", textCE11, 0
      SetCEFieldData "CE12", textCE12, 0
      SetCEFieldData "CE13", textCE13, 0
      SetCEFieldData "CE14", textCE14, 0
      SetCEFieldData "CE15", textCE15, 0
      'add by nickc 2007/01/26
      SetCEFieldData "CE68", textCE68, 0
      SetCEFieldData "CE69", textCE69, 0
      SetCEFieldData "CE70", textCE70, 0
      SetCEFieldData "CE71", textCE71, 0
      SetCEFieldData "CE72", textCE72, 0
      SetCEFieldData "CE73", textCE73, 0
      SetCEFieldData "CE74", textCE74, 0
      SetCEFieldData "CE75", textCE75, 0
      SetCEFieldData "CE76", textCE76, 0
      SetCEFieldData "CE77", textCE77, 0
      SetCEFieldData "CE78", textCE78, 0
      SetCEFieldData "CE79", textCE79, 0
      SetCEFieldData "CE80", textCE80, 0
      SetCEFieldData "CE81", textCE81, 0
      SetCEFieldData "CE82", textCE82, 0
      SetCEFieldData "CE83", textCE83, 0
      SetCEFieldData "CE84", textCE84, 0
      SetCEFieldData "CE85", textCE85, 0
      SetCEFieldData "CE86", textCE86, 0
      SetCEFieldData "CE87", textCE87, 0
      SetCEFieldData "CE88", textCE88, 0
      SetCEFieldData "CE89", textCE89, 0
      SetCEFieldData "CE90", textCE90, 0
      SetCEFieldData "CE91", textCE91, 0
     
   End If
   If checkCE56.Value = 1 Then
      SetCEFieldData "CE55", "V", 0
   End If
   If checkCE54.Value = 1 Then
      SetCEFieldData "CE53", "V", 0
   End If
   If checkCE52.Value = 1 Then
      SetCEFieldData "CE51", "V", 0
   End If
   If checkCE22.Value = 1 Then
      SetCEFieldData "CE17", textCE17, 0
      'add by nickc 2007/01/26
      SetCEFieldData "CE18", textCE18, 0
      SetCEFieldData "CE19", textCE19, 0
      SetCEFieldData "CE20", textCE20, 0
      SetCEFieldData "CE21", textCE21, 0
   End If
   If checkCE65.Value = 1 Then
      SetCEFieldData "CE63", textCE63, 0
      SetCEFieldData "CE64", textCE64, 0
      'add by nickc 2007/01/26
      SetCEFieldData "CE92", textCE92, 0
      SetCEFieldData "CE93", textCE93, 0
      SetCEFieldData "CE94", textCE94, 0
      SetCEFieldData "CE95", textCE95, 0
      SetCEFieldData "CE96", textCE96, 0
      SetCEFieldData "CE97", textCE97, 0
      SetCEFieldData "CE98", textCE98, 0
      SetCEFieldData "CE99", textCE99, 0
   End If
   If checkCE38.Value = 1 Then
      SetCEFieldData "CE23", textCE23, 0
      SetCEFieldData "CE24", textCE24, 0
      SetCEFieldData "CE25", textCE25, 0
      'add by nickc 2007/01/26
      SetCEFieldData "CE26", textCE26, 0
      SetCEFieldData "CE27", textCE27, 0
      SetCEFieldData "CE28", textCE28, 0
      SetCEFieldData "CE29", textCE29, 0
      SetCEFieldData "CE30", textCE30, 0
      SetCEFieldData "CE31", textCE31, 0
      SetCEFieldData "CE32", textCE32, 0
      SetCEFieldData "CE33", textCE33, 0
      SetCEFieldData "CE34", textCE34, 0
      SetCEFieldData "CE35", textCE35, 0
      SetCEFieldData "CE36", textCE36, 0
      SetCEFieldData "CE37", textCE37, 0
   End If
   If checkCE58.Value = 1 Then
      SetCEFieldData "CE57", textCE57, 0
   End If
   If checkCE40.Value = 1 Then
      SetCEFieldData "CE39", textCE39, 0
   End If
   If checkCE60.Value = 1 Then
      SetCEFieldData "CE59", "V", 0
   End If
   If checkCE44.Value = 1 Then
        Select Case m_TM01
        Case "FCT", "S"
            SetCEFieldData "CE41", textCE41_1, 0
        Case Else
            SetCEFieldData "CE41", textCE41, 0
            SetCEFieldData "CE42", textCE42, 0
            SetCEFieldData "CE43", textCE43, 0
        End Select
   End If
   If checkCE46.Value = 1 Then
      SetCEFieldData "CE45", textCE45, 0
   End If
   If checkCE48.Value = 1 Then
      SetCEFieldData "CE47", textCE47, 0
   End If
   If checkCE50.Value = 1 Then
      SetCEFieldData "CE49", textCE49, 0
   End If
   If checkCE62.Value = 1 Then
      SetCEFieldData "CE61", textCE61, 0
   End If
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010401_4.Show
End Sub

Private Sub checkCE38_Click()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
    'Add By Cheng 2003/04/17
    '若有勾選變更申請地址, 且未勾選變更申請人
    If Me.checkCE38.Value = vbChecked And Me.checkCE09.Value = vbUnchecked Then
        'edit by nickc 2007/01/26
        'StrSQLa = "Select TM23 From Trademark Where " & ChgTradeMark(Replace(textTMKey.Text, "-", ""))
        'StrSQLa = StrSQLa & " union Select SP08 From Servicepractice Where " & ChgService(Replace(textTMKey.Text, "-", ""))
        StrSQLa = "Select TM23,tm78,tm79,tm80,tm81 From Trademark Where " & ChgTradeMark(Replace(textTMKey.Text, "-", ""))
        StrSQLa = StrSQLa & " union Select SP08,sp58,sp59,sp65,sp66 From Servicepractice Where " & ChgService(Replace(textTMKey.Text, "-", ""))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            '顯示申請人地址
            textCE23.Text = PUB_GetCustEachAdd("" & rsA.Fields(0).Value, "1")
            textCE24.Text = PUB_GetCustEachAdd("" & rsA.Fields(0).Value, "2")
            textCE25.Text = PUB_GetCustEachAdd("" & rsA.Fields(0).Value, "3")
            'add by nickc 2007/01/26
            '顯示申請人地址
            textCE26.Text = PUB_GetCustEachAdd("" & rsA.Fields(1).Value, "1")
            textCE27.Text = PUB_GetCustEachAdd("" & rsA.Fields(1).Value, "2")
            textCE28.Text = PUB_GetCustEachAdd("" & rsA.Fields(1).Value, "3")
            textCE29.Text = PUB_GetCustEachAdd("" & rsA.Fields(2).Value, "1")
            textCE30.Text = PUB_GetCustEachAdd("" & rsA.Fields(2).Value, "2")
            textCE31.Text = PUB_GetCustEachAdd("" & rsA.Fields(2).Value, "3")
            textCE32.Text = PUB_GetCustEachAdd("" & rsA.Fields(3).Value, "1")
            textCE33.Text = PUB_GetCustEachAdd("" & rsA.Fields(3).Value, "2")
            textCE34.Text = PUB_GetCustEachAdd("" & rsA.Fields(3).Value, "3")
            textCE35.Text = PUB_GetCustEachAdd("" & rsA.Fields(4).Value, "1")
            textCE36.Text = PUB_GetCustEachAdd("" & rsA.Fields(4).Value, "2")
            textCE37.Text = PUB_GetCustEachAdd("" & rsA.Fields(4).Value, "3")
            
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    '若有勾選變更申請地址, 同時勾選變更申請人
    ElseIf Me.checkCE38.Value = vbChecked And Me.checkCE09.Value = vbChecked Then
        '顯示申請人地址
        textCE23.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE04.Text), "1")
        textCE24.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE04.Text), "2")
        textCE25.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE04.Text), "3")
        'add by nickc 2007/01/26
        textCE26.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE05.Text), "1")
        textCE27.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE05.Text), "2")
        textCE28.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE05.Text), "3")
        textCE29.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE06.Text), "1")
        textCE30.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE06.Text), "2")
        textCE31.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE06.Text), "3")
        textCE32.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE07.Text), "1")
        textCE33.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE07.Text), "2")
        textCE34.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE07.Text), "3")
        textCE35.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE08.Text), "1")
        textCE36.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE08.Text), "2")
        textCE37.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE08.Text), "3")
        
    End If
End Sub

'Add By Sindy 2009/05/25
Private Sub cmdCopy_Click()
'Dim blCopy As Boolean
'   frm030202_18.txtCaseNo = textTMKey.Text
'   frm030202_18.strCP01 = m_TM01
'   frm030202_18.strCP02 = m_TM02
'   frm030202_18.strCP03 = m_TM03
'   frm030202_18.strCP04 = m_TM04
'   frm030202_18.strCP09 = m_CE01
'   If frm030202_18.CheckShowList Then
'      frm030202_18.Show vbModal
'   End If
'   strCE01 = frm030202_18.strCE01
'   blCopy = frm030202_18.BolOk
'   Unload frm030202_18
'   Set frm030202_18 = Nothing
'   If blCopy = True Then
'      m_CE01 = strCE01
'      strCE01 = Trim(textCE01.Text)
'      Call ClearCEFields
'      Call QueryData
'      textCE01.Text = strCE01
'      m_CE01 = Trim(textCE01.Text)
'   End If
   
   'Modify By Sindy 2009/06/03
   '取得同一天,同一申請人且為301.變更案的第一筆總收文號
   strSql = "SELECT CE01 From CaseProgress, ChangeEvent, Trademark " & _
                   "Where CP09 = CE01 " & _
                   "AND CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 " & _
                   "AND CP27=" & m_CP27 & " " & _
                   "AND CP01='FCT' " & _
                   "AND (instr(TM23,'" & m_TM23 & "')=1 OR instr(TM78,'" & m_TM78 & "')=1 OR instr(TM79,'" & m_TM79 & "')=1 OR instr(TM80,'" & m_TM80 & "')=1 OR instr(TM81,'" & m_TM81 & "')=1) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      Call ClearCEFields
      Call QueryData2("" & RsTemp("CE01"))
      Call CheckTrue
      MsgBox "複製完成!!!", vbExclamation + vbOKOnly
      Exit Sub
   Else
      MsgBox "無資料可複製!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
End Sub

'Add By Sindy 2009/07/02
Private Sub cmdCopyCaseno_Click()
Dim strNo1 As String, strNo2 As String, strNo3 As String, strNo4 As String
   
   strNo1 = Trim(txt2(0))
   strNo2 = Trim(txt2(1))
   If Trim(txt2(2)) <> "" Then
      strNo3 = Trim(txt2(2))
   Else
      strNo3 = "0"
   End If
   If Trim(txt2(3)) <> "" Then
      strNo4 = Trim(txt2(3))
   Else
      strNo4 = "00"
   End If
   If strNo1 = "" Or strNo2 = "" Then
      MsgBox "請輸入欲複製之案號!!!", vbExclamation + vbOKOnly
      If strNo1 = "" Then
         txt2(0).SetFocus
      ElseIf strNo2 = "" Then
         txt2(1).SetFocus
      End If
      Exit Sub
   End If
   
   '只抓該案號同一天發文之變更來預設
   strSql = "SELECT CE01 From CaseProgress, ChangeEvent " & _
                   "Where CP09 = CE01 " & _
                   "AND CP01='" & strNo1 & "' AND CP02='" & strNo2 & "' AND CP03='" & strNo3 & "' AND CP04='" & strNo4 & "' " & _
                   "AND CP27=" & m_CP27 & " " & _
                   "AND CP01='FCT' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      Call ClearCEFields
      Call QueryData2("" & RsTemp("CE01"))
      Call CheckTrue
      MsgBox "複製完成!!!", vbExclamation + vbOKOnly
      Exit Sub
   Else
      MsgBox "無資料可複製!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
End Sub

Private Sub cmdExit_Click()
   Select Case m_Parent
      Case "frm030202_03":
         Unload frm030202_03
      Case "frm030202_07":
         Unload frm030202_07
      Case "frm030202_08":
         Unload frm030202_08
      Case "frm030202_10":
         Unload frm030202_10
      Case Else
   End Select
   Unload frm030202_01
   Unload Me
End Sub

Private Sub cmdGoods_Click()
frm03010303_04.Hide
Set frm03010303_04.UpForm = Me
frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
frm03010303_04.AllClass = m_TM09
frm03010303_04.cmdOK(2).Visible = True
Me.Hide
frm03010303_04.QueryData
frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
End Sub

Private Sub cmdOK_Click()
   If CheckDataValidate = False Then Exit Sub
   UpdateFieldNewData
      'edit by  nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
   Unload Me
   'm_PrevForm.Show
   
   'Add By Sindy 2012/4/17 有變更事項資料時,必須按確定鍵
   Select Case m_Parent
      Case "frm030202_03":
         frm030202_03.m_blnClkChgButton = True
         frm030202_03.Show
         frm030202_03.QueryData
      Case "frm030202_07":
         frm030202_07.m_blnClkChgButton = True
         frm030202_07.Show
         frm030202_07.QueryData
      Case "frm030202_08":
         frm030202_08.m_blnClkChgButton = True
         frm030202_08.Show
         frm030202_08.QueryData
      Case "frm030202_10":
         frm030202_10.m_blnClkChgButton = True
         frm030202_10.Show
         frm030202_10.QueryData
      Case Else
   End Select
End Sub

Private Sub Form_Load()

   textTMKey.BackColor = &H8000000F
   textCE01.BackColor = &H8000000F
   textCE04_2.BackColor = &H8000000F
   textCE39_2.BackColor = &H8000000F
   
   'add by nickc 2007/01/25
   textCE05_2.BackColor = &H8000000F
   textCE06_2.BackColor = &H8000000F
   textCE07_2.BackColor = &H8000000F
   textCE08_2.BackColor = &H8000000F
    'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    textCE10.MaxLength = Pub_MaxCEL10
    textCE11.MaxLength = Pub_MaxCEL11
    textCE13.MaxLength = Pub_MaxCEL10
    textCE14.MaxLength = Pub_MaxCEL11
    textCE68.MaxLength = Pub_MaxCEL10
    textCE69.MaxLength = Pub_MaxCEL11
    textCE71.MaxLength = Pub_MaxCEL10
    textCE72.MaxLength = Pub_MaxCEL11
    textCE74.MaxLength = Pub_MaxCEL10
    textCE75.MaxLength = Pub_MaxCEL11
    textCE77.MaxLength = Pub_MaxCEL10
    textCE78.MaxLength = Pub_MaxCEL11
    textCE80.MaxLength = Pub_MaxCEL10
    textCE81.MaxLength = Pub_MaxCEL11
    textCE83.MaxLength = Pub_MaxCEL10
    textCE84.MaxLength = Pub_MaxCEL11
    textCE86.MaxLength = Pub_MaxCEL10
    textCE87.MaxLength = Pub_MaxCEL11
    textCE89.MaxLength = Pub_MaxCEL10
    textCE90.MaxLength = Pub_MaxCEL11
    'end 2016/09/10
    Me.tabCtrl.Tab = 0 'Added by Lydia 2021/09/02
    
   MoveFormToCenter Me
    'Add By Cheng 2003/06/24
    '設定案件名稱欄位長度
    Select Case m_TM01
    Case "CFT", "FCT", "T", "TF", "S"
        Me.textCE41_1.MaxLength = 140
        Me.textCE41.MaxLength = 40
        Me.textCE42.MaxLength = 60
        Me.textCE43.MaxLength = 40
        Me.Label3.Visible = True
        Me.textCE41_1.Visible = True
        Me.textCE41_1.Enabled = True
        Me.Label34.Visible = False
        Me.textCE41.Visible = False
        Me.textCE41.Enabled = False
        Me.Label35.Visible = False
        Me.textCE42.Visible = False
        Me.textCE42.Enabled = False
        Me.Label36.Visible = False
        Me.textCE43.Visible = False
        Me.textCE43.Enabled = False
    Case Else
        Me.textCE41.MaxLength = 60
        Me.textCE42.MaxLength = 60
        Me.textCE43.MaxLength = 60
        Me.Label3.Visible = False
        Me.textCE41_1.Visible = False
        Me.textCE41_1.Enabled = False
        Me.Label34.Visible = True
        Me.textCE41.Visible = True
        Me.textCE41.Enabled = True
        Me.Label35.Visible = True
        Me.textCE42.Visible = True
        Me.textCE42.Enabled = True
        Me.Label36.Visible = True
        Me.textCE43.Visible = True
        Me.textCE43.Enabled = True
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
   ClearCEFields
   'Add By Cheng 2002/07/19
   Set frm030202_05 = Nothing
End Sub

' 由客戶代碼取得客戶名稱
Private Function GetCustomer(ByVal strData As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCustomer = Empty
   If IsEmptyText(strData) = False Then
      Set rsTmp = New ADODB.Recordset
      If Len(strData) > 8 Then
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                        "CU02 = '" & Mid(strData, 9, 1) & "'"
      Else
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CU04")) = False Then
            GetCustomer = rsTmp.Fields("CU04")
         ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
            GetCustomer = rsTmp.Fields("CU05")
         ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
            GetCustomer = rsTmp.Fields("CU06")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

' 由客戶代號取得地址
' Input : strData ==> 客戶代號
'         nType ==> 種類
'                   0 : 表要取得的是中文地址
'                   1 : 表要取得的是英文地址
'                   2 : 表要取得的是日文地址
Private Function GetAddress(ByVal strData As String, ByVal nType As Integer) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetAddress = Empty
   If IsEmptyText(strData) = False Then
      If Len(strData) > 8 Then
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                        "CU02 = '" & Mid(strData, 9, 1) & "'"
      Else
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Select Case nType
            Case 0:
               If IsNull(rsTmp.Fields("CU23")) = False Then
                  GetAddress = rsTmp.Fields("CU23")
               End If
            Case 1:
               If IsNull(rsTmp.Fields("CU24")) = False Then
                  GetAddress = rsTmp.Fields("CU24")
               End If
            Case 2:
               If IsNull(rsTmp.Fields("CU29")) = False Then
                  GetAddress = rsTmp.Fields("CU29")
               End If
         End Select
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CE01 = Empty
      'Add By Sindy 2009/06/03
      m_TM23 = Empty
      m_TM78 = Empty
      m_TM79 = Empty
      m_TM80 = Empty
      m_TM81 = Empty
      m_CP27 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
      ' 收文號
      Case 4: m_CE01 = strData
      'Add By Sindy 2009/06/03
      Case 5: m_TM23 = strData '申請人1
      Case 6: m_TM78 = strData '申請人2
      Case 7: m_TM79 = strData '申請人3
      Case 8: m_TM80 = strData '申請人4
      Case 9: m_TM81 = strData '申請人5
      Case 10: m_CP27 = strData '發文日期
   End Select
End Sub

'Public Sub SetParent(ByRef fm As Form)
Public Sub SetParent(ByVal strParent As String)
   'm_PrevForm = fm
   m_Parent = strParent
End Sub

'Add By Sindy 2009/06/16
Public Sub QueryData2(strCE01 As String)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   Call ClearField
   
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & strCE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請日
      If IsNull(rsTmp.Fields("CE02")) = False Then
         textCE02 = ChangeWStringToTString(rsTmp.Fields("CE02"))
         'tmpOldCE02 = Me.textCE02.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("CE04")) = False Then
         textCE04 = rsTmp.Fields("CE04")
         textCE04_2 = GetCustomer(rsTmp.Fields("CE04"))
         'tmpOldCE04 = textCE04.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE05")) = False Then
         'tmpOldCE05 = CheckStr(rsTmp.Fields("CE05"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
         textCE05 = rsTmp.Fields("CE05")
         textCE05_2 = GetCustomer(rsTmp.Fields("CE05"))
      End If
      If IsNull(rsTmp.Fields("CE06")) = False Then
         'tmpOldCE06 = CheckStr(rsTmp.Fields("CE06"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
         textCE06 = rsTmp.Fields("CE06")
         textCE06_2 = GetCustomer(rsTmp.Fields("CE06"))
      End If
      If IsNull(rsTmp.Fields("CE07")) = False Then
         'tmpOldCE07 = CheckStr(rsTmp.Fields("CE07"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
         textCE07 = rsTmp.Fields("CE07")
         textCE07_2 = GetCustomer(rsTmp.Fields("CE07"))
      End If
      If IsNull(rsTmp.Fields("CE08")) = False Then
         'tmpOldCE08 = CheckStr(rsTmp.Fields("CE08"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
         textCE08 = rsTmp.Fields("CE08")
         textCE08_2 = GetCustomer(rsTmp.Fields("CE08"))
      End If
      
      'Add By Sindy 2012/3/5
      '申請人中譯文
      If IsNull(rsTmp.Fields("CE17")) = False Then
         textCE17 = rsTmp.Fields("CE17")
      End If
      If IsNull(rsTmp.Fields("CE18")) = False Then
         textCE18 = rsTmp.Fields("CE18")
      End If
      If IsNull(rsTmp.Fields("CE19")) = False Then
         textCE19 = rsTmp.Fields("CE19")
      End If
      If IsNull(rsTmp.Fields("CE20")) = False Then
         textCE20 = rsTmp.Fields("CE20")
      End If
      If IsNull(rsTmp.Fields("CE21")) = False Then
         textCE21 = rsTmp.Fields("CE21")
      End If
      
      ' 代表人
      If IsNull(rsTmp.Fields("CE10")) = False Then
         textCE10 = rsTmp.Fields("CE10")
         'tmpOldCE10 = textCE10.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE11")) = False Then
         textCE11 = rsTmp.Fields("CE11")
         'tmpOldCE11 = textCE11.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE12")) = False Then
         textCE12 = rsTmp.Fields("CE12")
         'tmpOldCE12 = textCE12.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE13")) = False Then
         textCE13 = rsTmp.Fields("CE13")
         'tmpOldCE13 = textCE13.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE14")) = False Then
         textCE14 = rsTmp.Fields("CE14")
         'tmpOldCE14 = textCE14.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE15")) = False Then
         textCE15 = rsTmp.Fields("CE15")
         'tmpOldCE15 = textCE15.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      textCE68 = CheckStr(rsTmp.Fields("CE68"))
      'tmpOldCE68 = CheckStr(rsTmp.Fields("CE68"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE69 = CheckStr(rsTmp.Fields("CE69"))
      'tmpOldCE69 = CheckStr(rsTmp.Fields("CE69"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE70 = CheckStr(rsTmp.Fields("CE70"))
      'tmpOldCE70 = CheckStr(rsTmp.Fields("CE70"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE71 = CheckStr(rsTmp.Fields("CE71"))
      'tmpOldCE71 = CheckStr(rsTmp.Fields("CE71"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE72 = CheckStr(rsTmp.Fields("CE72"))
      'tmpOldCE72 = CheckStr(rsTmp.Fields("CE72"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE73 = CheckStr(rsTmp.Fields("CE73"))
      'tmpOldCE73 = CheckStr(rsTmp.Fields("CE73"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE74 = CheckStr(rsTmp.Fields("CE74"))
      'tmpOldCE74 = CheckStr(rsTmp.Fields("CE74"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE75 = CheckStr(rsTmp.Fields("CE75"))
      'tmpOldCE75 = CheckStr(rsTmp.Fields("CE75"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE76 = CheckStr(rsTmp.Fields("CE76"))
      'tmpOldCE76 = CheckStr(rsTmp.Fields("CE76"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE77 = CheckStr(rsTmp.Fields("CE77"))
      'tmpOldCE77 = CheckStr(rsTmp.Fields("CE77"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE78 = CheckStr(rsTmp.Fields("CE78"))
      'tmpOldCE78 = CheckStr(rsTmp.Fields("CE78"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE79 = CheckStr(rsTmp.Fields("CE79"))
      'tmpOldCE79 = CheckStr(rsTmp.Fields("CE79"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE80 = CheckStr(rsTmp.Fields("CE80"))
      'tmpOldCE80 = CheckStr(rsTmp.Fields("CE80"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE81 = CheckStr(rsTmp.Fields("CE81"))
      'tmpOldCE81 = CheckStr(rsTmp.Fields("CE81"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE82 = CheckStr(rsTmp.Fields("CE82"))
      'tmpOldCE82 = CheckStr(rsTmp.Fields("CE82"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE83 = CheckStr(rsTmp.Fields("CE83"))
      'tmpOldCE83 = CheckStr(rsTmp.Fields("CE83"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE84 = CheckStr(rsTmp.Fields("CE84"))
      'tmpOldCE84 = CheckStr(rsTmp.Fields("CE84"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE85 = CheckStr(rsTmp.Fields("CE85"))
      'tmpOldCE85 = CheckStr(rsTmp.Fields("CE85"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE86 = CheckStr(rsTmp.Fields("CE86"))
      'tmpOldCE86 = CheckStr(rsTmp.Fields("CE86"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE87 = CheckStr(rsTmp.Fields("CE87"))
      'tmpOldCE87 = CheckStr(rsTmp.Fields("CE87"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE88 = CheckStr(rsTmp.Fields("CE88"))
      'tmpOldCE88 = CheckStr(rsTmp.Fields("CE88"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE89 = CheckStr(rsTmp.Fields("CE89"))
      'tmpOldCE89 = CheckStr(rsTmp.Fields("CE89"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE90 = CheckStr(rsTmp.Fields("CE90"))
      'tmpOldCE90 = CheckStr(rsTmp.Fields("CE90"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      textCE91 = CheckStr(rsTmp.Fields("CE91"))
      'tmpOldCE91 = CheckStr(rsTmp.Fields("CE91"))   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      
      ' 申請地址
      If IsNull(rsTmp.Fields("CE23")) = False Then
         textCE23 = rsTmp.Fields("CE23")
         'tmpOldCE23 = textCE23.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE24")) = False Then
         textCE24 = rsTmp.Fields("CE24")
         'tmpOldCE24 = textCE24.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE25")) = False Then
         textCE25 = rsTmp.Fields("CE25")
         'tmpOldCE25 = textCE25.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE26")) = False Then
         textCE26 = rsTmp.Fields("CE26")
         'tmpOldCE26 = textCE26.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE27")) = False Then
         textCE27 = rsTmp.Fields("CE27")
         'tmpOldCE27 = textCE27.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE28")) = False Then
         textCE28 = rsTmp.Fields("CE28")
         'tmpOldCE28 = textCE28.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE29")) = False Then
         textCE29 = rsTmp.Fields("CE29")
         'tmpOldCE29 = textCE29.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE30")) = False Then
         textCE30 = rsTmp.Fields("CE30")
         'tmpOldCE30 = textCE30.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE31")) = False Then
         textCE31 = rsTmp.Fields("CE31")
         'tmpOldCE31 = textCE31.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE32")) = False Then
         textCE32 = rsTmp.Fields("CE32")
         'tmpOldCE32 = textCE32.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE33")) = False Then
         textCE33 = rsTmp.Fields("CE33")
         'tmpOldCE33 = textCE33.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE34")) = False Then
         textCE34 = rsTmp.Fields("CE34")
         'tmpOldCE34 = textCE34.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE35")) = False Then
         textCE35 = rsTmp.Fields("CE35")
         'tmpOldCE35 = textCE35.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE36")) = False Then
         textCE36 = rsTmp.Fields("CE36")
         'tmpOldCE36 = textCE36.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE37")) = False Then
         textCE37 = rsTmp.Fields("CE37")
         'tmpOldCE37 = textCE37.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      
      ' 專利商標種類代號
      If IsNull(rsTmp.Fields("CE39")) = False Then
         textCE39 = rsTmp.Fields("CE39")
         If IsEmptyText(textCE39) = False Then: textCE39_Validate (False)
         'tmpOldCE39 = textCE39.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      ' 案件名稱
      If IsNull(rsTmp.Fields("CE41")) = False Then
        Select Case m_TM01
        Case "FCT", "S"
            textCE41_1 = rsTmp.Fields("CE41")
            'tmpOldCE41 = textCE41_1.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
        Case Else
            textCE41 = rsTmp.Fields("CE41")
            'tmpOldCE41 = textCE41.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
        End Select
      End If
      If IsNull(rsTmp.Fields("CE42")) = False Then
         textCE42 = rsTmp.Fields("CE42")
         'tmpOldCE42 = textCE42.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      If IsNull(rsTmp.Fields("CE43")) = False Then
         textCE43 = rsTmp.Fields("CE43")
         'tmpOldCE43 = textCE43.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      ' 縮減商品
      If IsNull(rsTmp.Fields("CE45")) = False Then
         textCE45 = rsTmp.Fields("CE45")
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("CE47")) = False Then
         textCE47 = rsTmp.Fields("CE47")
         'tmpOldCE47 = textCE47.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      ' 商品群組
      If IsNull(rsTmp.Fields("CE49")) = False Then
         textCE49 = rsTmp.Fields("CE49")
         'tmpOldCE49 = textCE49.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      ' 申請人印鑑
      If rsTmp.Fields("CE51") = "V" Then
         checkCE52.Value = 1
      End If
      ' 代表人印鑑
      If rsTmp.Fields("CE53") = "V" Then
         checkCE54.Value = 1
      End If
      ' 代理人
      If rsTmp.Fields("CE55") = "V" Then
         checkCE56.Value = 1
      End If
      ' 正商標號數
      If IsNull(rsTmp.Fields("CE57")) = False Then
         textCE57 = rsTmp.Fields("CE57")
         'tmpOldCE57 = textCE57.Text   '2009/12/15 CANCEL BY SONIA 複製資料不可搬暫存變數
      End If
      ' 圖樣
      If rsTmp.Fields("CE59") = "V" Then
         checkCE60.Value = 1
      End If
      ' 其它
      If IsNull(rsTmp.Fields("CE61")) = False Then
         textCE61 = rsTmp.Fields("CE61")
      End If
      ' 代表人譯文
      If IsNull(rsTmp.Fields("CE63")) = False Then
         textCE63 = rsTmp.Fields("CE63")
      End If
      If IsNull(rsTmp.Fields("CE64")) = False Then
         textCE64 = rsTmp.Fields("CE64")
      End If
      
      textCE92 = CheckStr(rsTmp.Fields("CE92"))
      textCE93 = CheckStr(rsTmp.Fields("CE93"))
      textCE94 = CheckStr(rsTmp.Fields("CE94"))
      textCE95 = CheckStr(rsTmp.Fields("CE95"))
      textCE96 = CheckStr(rsTmp.Fields("CE96"))
      textCE97 = CheckStr(rsTmp.Fields("CE97"))
      textCE98 = CheckStr(rsTmp.Fields("CE98"))
      textCE99 = CheckStr(rsTmp.Fields("CE99"))
   End If
   rsTmp.Close
   Set rsTmp = Nothing

End Sub

'2009/12/15 add by sonia
Public Sub ClearWorkField()
   ' 清除暫存變數
   tmpOldCE04 = Empty
   tmpOldCE02 = Empty
   tmpOldCE05 = Empty
   tmpOldCE06 = Empty
   tmpOldCE07 = Empty
   tmpOldCE08 = Empty
   tmpOldCE10 = Empty
   tmpOldCE11 = Empty
   tmpOldCE12 = Empty
   tmpOldCE13 = Empty
   tmpOldCE14 = Empty
   tmpOldCE15 = Empty
   tmpOldCE68 = Empty
   tmpOldCE69 = Empty
   tmpOldCE70 = Empty
   tmpOldCE71 = Empty
   tmpOldCE72 = Empty
   tmpOldCE73 = Empty
   tmpOldCE74 = Empty
   tmpOldCE75 = Empty
   tmpOldCE76 = Empty
   tmpOldCE77 = Empty
   tmpOldCE78 = Empty
   tmpOldCE79 = Empty
   tmpOldCE80 = Empty
   tmpOldCE81 = Empty
   tmpOldCE82 = Empty
   tmpOldCE83 = Empty
   tmpOldCE84 = Empty
   tmpOldCE85 = Empty
   tmpOldCE86 = Empty
   tmpOldCE87 = Empty
   tmpOldCE88 = Empty
   tmpOldCE89 = Empty
   tmpOldCE90 = Empty
   tmpOldCE91 = Empty
   tmpOldCE23 = Empty
   tmpOldCE24 = Empty
   tmpOldCE25 = Empty
   tmpOldCE26 = Empty
   tmpOldCE27 = Empty
   tmpOldCE28 = Empty
   tmpOldCE29 = Empty
   tmpOldCE30 = Empty
   tmpOldCE31 = Empty
   tmpOldCE32 = Empty
   tmpOldCE33 = Empty
   tmpOldCE34 = Empty
   tmpOldCE35 = Empty
   tmpOldCE36 = Empty
   tmpOldCE37 = Empty
   tmpOldCE39 = Empty
   tmpOldCE41 = Empty
   tmpOldCE42 = Empty
   tmpOldCE43 = Empty
   tmpOldCE47 = Empty
   tmpOldCE49 = Empty
   tmpOldCE57 = Empty

End Sub
'2009/12/15 end

'Add By Sindy 2009/06/16
Public Sub ClearField()
   ' 清除欄位值
   textCE02 = Empty
   textCE04 = Empty
   textCE04_2 = Empty
   textCE05 = Empty
   textCE05_2 = Empty
   textCE06 = Empty
   textCE06_2 = Empty
   textCE07 = Empty
   textCE07_2 = Empty
   textCE08 = Empty
   textCE08_2 = Empty
   textCE10 = Empty
   textCE11 = Empty
   textCE12 = Empty
   textCE13 = Empty
   textCE14 = Empty
   textCE15 = Empty
   textCE68 = Empty
   textCE69 = Empty
   textCE70 = Empty
   textCE71 = Empty
   textCE72 = Empty
   textCE73 = Empty
   textCE74 = Empty
   textCE75 = Empty
   textCE76 = Empty
   textCE77 = Empty
   textCE78 = Empty
   textCE79 = Empty
   textCE80 = Empty
   textCE81 = Empty
   textCE82 = Empty
   textCE83 = Empty
   textCE84 = Empty
   textCE85 = Empty
   textCE86 = Empty
   textCE87 = Empty
   textCE88 = Empty
   textCE89 = Empty
   textCE90 = Empty
   textCE91 = Empty
   textCE23 = Empty
   textCE24 = Empty
   textCE25 = Empty
   textCE26 = Empty
   textCE27 = Empty
   textCE28 = Empty
   textCE29 = Empty
   textCE30 = Empty
   textCE31 = Empty
   textCE32 = Empty
   textCE33 = Empty
   textCE34 = Empty
   textCE35 = Empty
   textCE36 = Empty
   textCE37 = Empty
   textCE39 = Empty
   textCE41_1 = Empty
   textCE41 = Empty
   textCE42 = Empty
   textCE43 = Empty
   textCE45 = Empty
   textCE47 = Empty
   textCE49 = Empty
   checkCE52.Value = vbUnchecked
   checkCE54.Value = vbUnchecked
   checkCE56.Value = vbUnchecked
   textCE57 = Empty
   checkCE60.Value = vbUnchecked
   textCE61 = Empty
   textCE63 = Empty
   textCE64 = Empty
   textCE92 = Empty
   textCE93 = Empty
   textCE94 = Empty
   textCE95 = Empty
   textCE96 = Empty
   textCE97 = Empty
   textCE98 = Empty
   textCE99 = Empty
End Sub

'Add By Sindy 2009/07/02
Public Sub CheckTrue()
      '判斷是否有其變更事項,則打勾
      If Trim(textCE04.Text) <> "" Or Trim(textCE05.Text) <> "" Or _
         Trim(textCE06.Text) <> "" Or Trim(textCE07.Text) <> "" Or _
         Trim(textCE08.Text) <> "" Then
         checkCE09.Value = 1
      Else
         checkCE09.Value = vbUnchecked
      End If
      If Trim(textCE02.Text) <> "" Then
         checkCE03.Value = 1
      Else
         checkCE03.Value = vbUnchecked
      End If
      If Trim(textCE17.Text) <> "" Or Trim(textCE18.Text) <> "" Or _
         Trim(textCE19.Text) <> "" Or Trim(textCE20.Text) <> "" Or _
         Trim(textCE21.Text) <> "" Then
         checkCE22.Value = 1
      Else
         checkCE22.Value = vbUnchecked
      End If
      If Trim(textCE63.Text) <> "" Or Trim(textCE64.Text) <> "" Or _
         Trim(textCE92.Text) <> "" Or Trim(textCE93.Text) <> "" Or _
         Trim(textCE94.Text) <> "" Or Trim(textCE95.Text) <> "" Or _
         Trim(textCE96.Text) <> "" Or Trim(textCE97.Text) <> "" Or _
         Trim(textCE98.Text) <> "" Or Trim(textCE99.Text) <> "" Then
         checkCE65.Value = 1
      Else
         checkCE65.Value = vbUnchecked
      End If
      If Trim(textCE10.Text) <> "" Or Trim(textCE11.Text) <> "" Or _
         Trim(textCE12.Text) <> "" Or Trim(textCE13.Text) <> "" Or _
         Trim(textCE14.Text) <> "" Or Trim(textCE15.Text) <> "" Or _
         Trim(textCE68.Text) <> "" Or Trim(textCE69.Text) <> "" Or _
         Trim(textCE70.Text) <> "" Or Trim(textCE71.Text) <> "" Or _
         Trim(textCE72.Text) <> "" Or Trim(textCE73.Text) <> "" Or _
         Trim(textCE74.Text) <> "" Or Trim(textCE75.Text) <> "" Or _
         Trim(textCE76.Text) <> "" Or Trim(textCE77.Text) <> "" Or _
         Trim(textCE78.Text) <> "" Or Trim(textCE79.Text) <> "" Or _
         Trim(textCE80.Text) <> "" Or Trim(textCE81.Text) <> "" Or _
         Trim(textCE82.Text) <> "" Or Trim(textCE83.Text) <> "" Or _
         Trim(textCE84.Text) <> "" Or Trim(textCE85.Text) <> "" Or _
         Trim(textCE86.Text) <> "" Or Trim(textCE87.Text) <> "" Or _
         Trim(textCE88.Text) <> "" Or Trim(textCE89.Text) <> "" Or _
         Trim(textCE90.Text) <> "" Or Trim(textCE91.Text) <> "" Then
         checkCE16.Value = 1
      Else
         checkCE16.Value = vbUnchecked
      End If
      If Trim(textCE23.Text) <> "" Or Trim(textCE24.Text) <> "" Or _
         Trim(textCE25.Text) <> "" Or Trim(textCE26.Text) <> "" Or _
         Trim(textCE27.Text) <> "" Or Trim(textCE28.Text) <> "" Or _
         Trim(textCE29.Text) <> "" Or Trim(textCE30.Text) <> "" Or _
         Trim(textCE31.Text) <> "" Or Trim(textCE32.Text) <> "" Or _
         Trim(textCE33.Text) <> "" Or Trim(textCE34.Text) <> "" Or _
         Trim(textCE35.Text) <> "" Or Trim(textCE36.Text) <> "" Or _
         Trim(textCE37.Text) <> "" Then
         checkCE38.Value = 1
      Else
         checkCE38.Value = vbUnchecked
      End If
      If Trim(textCE57.Text) <> "" Then
         checkCE58.Value = 1
      Else
         checkCE58.Value = vbUnchecked
      End If
      If Trim(textCE39.Text) <> "" Then
         checkCE40.Value = 1
      Else
         checkCE40.Value = vbUnchecked
      End If
      If Trim(textCE41.Text) <> "" Or Trim(textCE42.Text) <> "" Or _
         Trim(textCE43.Text) <> "" Then
         checkCE44.Value = 1
      Else
         checkCE44.Value = vbUnchecked
      End If
      If Trim(textCE45.Text) <> "" Then
         checkCE46.Value = 1
      Else
         checkCE46.Value = vbUnchecked
      End If
      If Trim(textCE47.Text) <> "" Then
         checkCE48.Value = 1
      Else
         checkCE48.Value = vbUnchecked
      End If
      If Trim(textCE49.Text) <> "" Then
         checkCE50.Value = 1
      Else
         checkCE50.Value = vbUnchecked
      End If
      If Trim(textCE61.Text) <> "" Then
         checkCE62.Value = 1
      Else
         checkCE62.Value = vbUnchecked
      End If
End Sub

Public Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
      
   ' 清除暫存變數
   m_CE02 = Empty
   m_CE04 = Empty
   
   Call ClearWorkField     '2009/12/15 ADD BY SONIA 將清除暫存變數及清除欄位值分開,否則複製資料之QueryData2會蓋掉暫存變數FCT-005481
   Call ClearField
   
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textCE01 = m_CE01
   
    'Add By Cheng 2003/09/03
    'Begin
    'edit by nickc 2007/01/25
    'StrSQLa = "Select TM11, TM23, TM47, TM48, TM49, TM50, TM51, TM52, TM08, TM05, TM06, TM07, TM09, TM32, TM27 From Trademark Where " & ChgTradeMark(textTMKey)
    'StrSQLa = StrSQLa & " Union Select SP10, SP08, SP42, '', '', '', '', '', '', SP05, SP06, SP07, '', '', '' From Servicepractice Where " & ChgService(textTMKey)
    
   StrSQLa = "Select TM11, TM23, TM47, TM48, TM49, TM50, TM51, TM52, TM08, TM05, TM06, TM07, TM09, TM32, TM27,tm78,tm79,tm80,tm81,tm94,tm95,tm96,tm97,tm98,tm99,tm100,tm101,tm102,tm103,tm104,tm105,tm106,tm107,tm108,tm109,tm110,tm111,tm112,tm113,tm114,tm115,tm116,tm117,TM24,TM25,TM26,TM82,TM86,TM90,TM83,TM87,TM91,TM84,TM88,TM92,TM85,TM89,TM93 From Trademark Where " & ChgTradeMark(textTMKey)
   'edit by nickc 2007/04/03
   'StrSQLa = StrSQLa & " Union Select SP10, SP08, SP42, '', '', '', '', '', '', SP05, SP06, SP07, '', '', '',sp58,sp59,sp65,sp66,'','','','','','','','','','','','','','','','','','','','','','','','' From Servicepractice Where " & ChgService(textTMKey)
   StrSQLa = StrSQLa & " Union Select SP10, SP08, SP42, '', '', '', '', '', '', SP05, SP06, SP07, sp73, sp74, '',sp58,sp59,sp65,sp66,'','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','' From Servicepractice Where " & ChgService(textTMKey)
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
     ' 申請日
       tmpOldCE02 = "" & rsA.Fields(0).Value
     ' 申請人
       m_TM23 = "" & rsA.Fields(1).Value 'Add By Sindy 2011/8/3
       tmpOldCE04 = "" & rsA.Fields(1).Value
     '申請人地址
'       tmpOldCE23 = PUB_GetCustEachAdd(tmpOldCE04, "1")
'       tmpOldCE24 = PUB_GetCustEachAdd(tmpOldCE04, "2")
'       tmpOldCE25 = PUB_GetCustEachAdd(tmpOldCE04, "3")
       'Modify By Sindy 2011/2/1
       tmpOldCE23 = Trim("" & rsA.Fields(43).Value)
       tmpOldCE24 = Trim("" & rsA.Fields(44).Value)
       tmpOldCE25 = Trim("" & rsA.Fields(45).Value)
       '2011/2/1 End
     ' 代表人
       tmpOldCE10 = "" & rsA.Fields(2).Value
       tmpOldCE11 = "" & rsA.Fields(3).Value
       tmpOldCE12 = "" & rsA.Fields(4).Value
       tmpOldCE13 = "" & rsA.Fields(5).Value
       tmpOldCE14 = "" & rsA.Fields(6).Value
       tmpOldCE15 = "" & rsA.Fields(7).Value
     ' 專利商標種類代號
        tmpOldCE39 = "" & rsA.Fields(8).Value
     ' 案件名稱
       tmpOldCE41 = "" & rsA.Fields(9).Value
       tmpOldCE42 = "" & rsA.Fields(10).Value
       tmpOldCE43 = "" & rsA.Fields(11).Value
     ' 商品類別
       tmpOldCE47 = "" & rsA.Fields(12).Value
     ' 商品群組
       tmpOldCE49 = "" & rsA.Fields(13).Value
     ' 正商標號數
       tmpOldCE57 = "" & rsA.Fields(14).Value
       'add by nickc 2007/01/25
       tmpOldCE05 = "" & rsA.Fields(15).Value
       tmpOldCE06 = "" & rsA.Fields(16).Value
       tmpOldCE07 = "" & rsA.Fields(17).Value
       tmpOldCE08 = "" & rsA.Fields(18).Value
       m_TM78 = "" & rsA.Fields(15).Value 'Add By Sindy 2011/8/3
       m_TM79 = "" & rsA.Fields(16).Value 'Add By Sindy 2011/8/3
       m_TM80 = "" & rsA.Fields(17).Value 'Add By Sindy 2011/8/3
       m_TM81 = "" & rsA.Fields(18).Value 'Add By Sindy 2011/8/3
       
'       tmpOldCE26 = PUB_GetCustEachAdd(tmpOldCE05, "1")
'       tmpOldCE27 = PUB_GetCustEachAdd(tmpOldCE05, "2")
'       tmpOldCE28 = PUB_GetCustEachAdd(tmpOldCE05, "3")
'       tmpOldCE29 = PUB_GetCustEachAdd(tmpOldCE06, "1")
'       tmpOldCE30 = PUB_GetCustEachAdd(tmpOldCE06, "2")
'       tmpOldCE31 = PUB_GetCustEachAdd(tmpOldCE06, "3")
'       tmpOldCE32 = PUB_GetCustEachAdd(tmpOldCE07, "1")
'       tmpOldCE33 = PUB_GetCustEachAdd(tmpOldCE07, "2")
'       tmpOldCE34 = PUB_GetCustEachAdd(tmpOldCE07, "3")
'       tmpOldCE35 = PUB_GetCustEachAdd(tmpOldCE08, "1")
'       tmpOldCE36 = PUB_GetCustEachAdd(tmpOldCE08, "2")
'       tmpOldCE37 = PUB_GetCustEachAdd(tmpOldCE08, "3")
       'Modify By Sindy 2011/2/1
       tmpOldCE26 = Trim("" & rsA.Fields(46).Value)
       tmpOldCE27 = Trim("" & rsA.Fields(47).Value)
       tmpOldCE28 = Trim("" & rsA.Fields(48).Value)
       tmpOldCE29 = Trim("" & rsA.Fields(49).Value)
       tmpOldCE30 = Trim("" & rsA.Fields(50).Value)
       tmpOldCE31 = Trim("" & rsA.Fields(51).Value)
       tmpOldCE32 = Trim("" & rsA.Fields(52).Value)
       tmpOldCE33 = Trim("" & rsA.Fields(53).Value)
       tmpOldCE34 = Trim("" & rsA.Fields(54).Value)
       tmpOldCE35 = Trim("" & rsA.Fields(55).Value)
       tmpOldCE36 = Trim("" & rsA.Fields(56).Value)
       tmpOldCE37 = Trim("" & rsA.Fields(57).Value)
       '2011/2/1 End
       tmpOldCE68 = "" & rsA.Fields(19).Value
       tmpOldCE69 = "" & rsA.Fields(20).Value
       tmpOldCE70 = "" & rsA.Fields(21).Value
       tmpOldCE71 = "" & rsA.Fields(22).Value
       tmpOldCE72 = "" & rsA.Fields(23).Value
       tmpOldCE73 = "" & rsA.Fields(24).Value
       tmpOldCE74 = "" & rsA.Fields(25).Value
       tmpOldCE75 = "" & rsA.Fields(26).Value
       tmpOldCE76 = "" & rsA.Fields(27).Value
       tmpOldCE77 = "" & rsA.Fields(28).Value
       tmpOldCE78 = "" & rsA.Fields(29).Value
       tmpOldCE79 = "" & rsA.Fields(30).Value
       tmpOldCE80 = "" & rsA.Fields(31).Value
       tmpOldCE81 = "" & rsA.Fields(32).Value
       tmpOldCE82 = "" & rsA.Fields(33).Value
       tmpOldCE83 = "" & rsA.Fields(34).Value
       tmpOldCE84 = "" & rsA.Fields(35).Value
       tmpOldCE85 = "" & rsA.Fields(36).Value
       tmpOldCE86 = "" & rsA.Fields(37).Value
       tmpOldCE87 = "" & rsA.Fields(38).Value
       tmpOldCE88 = "" & rsA.Fields(39).Value
       tmpOldCE89 = "" & rsA.Fields(40).Value
       tmpOldCE90 = "" & rsA.Fields(41).Value
       tmpOldCE91 = "" & rsA.Fields(42).Value
       'add by nickc 2007/04/03
       m_TM09 = "" & rsA.Fields(12).Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   'Modify By Sindy 2012/5/18 Mark暫存變數，因暫存變數是為比對基本檔和畫面上的欄位值是否相同,所以暫存變數不須再存取變更檔資料
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請日
      If IsNull(rsTmp.Fields("CE02")) = False Then
         checkCE03.Value = 1 'Add By Sindy 2012/3/7
         m_CE02 = rsTmp.Fields("CE02")
         textCE02 = ChangeWStringToTString(rsTmp.Fields("CE02"))
'         tmpOldCE02 = Me.textCE02.Text
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("CE04")) = False Then
         checkCE09.Value = 1 'Add By Sindy 2012/3/7
         m_CE04 = rsTmp.Fields("CE04")
         textCE04 = rsTmp.Fields("CE04")
         textCE04_2 = GetCustomer(rsTmp.Fields("CE04"))
'         tmpOldCE04 = textCE04.Text
         '顯示申請人地址
         textCE23.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "1")
         textCE24.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "2")
         textCE25.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "3")
'         tmpOldCE23 = textCE23.Text
'         tmpOldCE24 = textCE24.Text
'         tmpOldCE25 = textCE25.Text
      End If
      
      'add by nickc 2006/12/25
      If IsNull(rsTmp.Fields("CE05")) = False Then
         checkCE09.Value = 1 'Add By Sindy 2012/3/7
         m_CE05 = rsTmp.Fields("CE05")
'         tmpOldCE05 = CheckStr(rsTmp.Fields("CE05"))
         textCE05 = rsTmp.Fields("CE05")
         textCE05_2 = GetCustomer(rsTmp.Fields("CE05"))
         '顯示申請人地址
         textCE26.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "1")
         textCE27.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "2")
         textCE28.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "3")
'         tmpOldCE26 = textCE26.Text
'         tmpOldCE27 = textCE27.Text
'         tmpOldCE28 = textCE28.Text
      End If
      If IsNull(rsTmp.Fields("CE06")) = False Then
         checkCE09.Value = 1 'Add By Sindy 2012/3/7
         m_CE06 = rsTmp.Fields("CE06")
'         tmpOldCE06 = CheckStr(rsTmp.Fields("CE06"))
         textCE06 = rsTmp.Fields("CE06")
         textCE06_2 = GetCustomer(rsTmp.Fields("CE06"))
         '顯示申請人地址
         textCE29.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "1")
         textCE30.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "2")
         textCE31.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "3")
'         tmpOldCE29 = textCE29.Text
'         tmpOldCE30 = textCE30.Text
'         tmpOldCE31 = textCE31.Text
      End If
      If IsNull(rsTmp.Fields("CE07")) = False Then
         checkCE09.Value = 1 'Add By Sindy 2012/3/7
         m_CE07 = rsTmp.Fields("CE07")
'         tmpOldCE07 = CheckStr(rsTmp.Fields("CE07"))
         textCE07 = rsTmp.Fields("CE07")
         textCE07_2 = GetCustomer(rsTmp.Fields("CE07"))
         '顯示申請人地址
         textCE32.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "1")
         textCE33.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "2")
         textCE34.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "3")
'         tmpOldCE32 = textCE32.Text
'         tmpOldCE33 = textCE33.Text
'         tmpOldCE34 = textCE34.Text
      End If
      If IsNull(rsTmp.Fields("CE08")) = False Then
         checkCE09.Value = 1 'Add By Sindy 2012/3/7
         m_CE08 = rsTmp.Fields("CE08")
'         tmpOldCE08 = CheckStr(rsTmp.Fields("CE08"))
         textCE08 = rsTmp.Fields("CE08")
         textCE08_2 = GetCustomer(rsTmp.Fields("CE08"))
         '顯示申請人地址
         textCE35.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "1")
         textCE36.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "2")
         textCE37.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "3")
'         tmpOldCE35 = textCE35.Text
'         tmpOldCE36 = textCE36.Text
'         tmpOldCE37 = textCE37.Text
      End If
      
      'Add By Sindy 2012/3/5
      '申請人中譯文
      If IsNull(rsTmp.Fields("CE17")) = False Then
         checkCE22.Value = 1 'Add By Sindy 2012/3/5
         textCE17 = rsTmp.Fields("CE17")
      End If
      If IsNull(rsTmp.Fields("CE18")) = False Then
         checkCE22.Value = 1 'Add By Sindy 2012/3/5
         textCE18 = rsTmp.Fields("CE18")
      End If
      If IsNull(rsTmp.Fields("CE19")) = False Then
         checkCE22.Value = 1 'Add By Sindy 2012/3/5
         textCE19 = rsTmp.Fields("CE19")
      End If
      If IsNull(rsTmp.Fields("CE20")) = False Then
         checkCE22.Value = 1 'Add By Sindy 2012/3/5
         textCE20 = rsTmp.Fields("CE20")
      End If
      If IsNull(rsTmp.Fields("CE21")) = False Then
         checkCE22.Value = 1 'Add By Sindy 2012/3/5
         textCE21 = rsTmp.Fields("CE21")
      End If
            
      ' 代表人
      If IsNull(rsTmp.Fields("CE10")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE10 = rsTmp.Fields("CE10")
'         tmpOldCE10 = textCE10.Text
      End If
      If IsNull(rsTmp.Fields("CE11")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE11 = rsTmp.Fields("CE11")
'         tmpOldCE11 = textCE11.Text
      End If
      If IsNull(rsTmp.Fields("CE12")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE12 = rsTmp.Fields("CE12")
'         tmpOldCE12 = textCE12.Text
      End If
      If IsNull(rsTmp.Fields("CE13")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE13 = rsTmp.Fields("CE13")
'         tmpOldCE13 = textCE13.Text
      End If
      If IsNull(rsTmp.Fields("CE14")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE14 = rsTmp.Fields("CE14")
'         tmpOldCE14 = textCE14.Text
      End If
      If IsNull(rsTmp.Fields("CE15")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE15 = rsTmp.Fields("CE15")
'         tmpOldCE15 = textCE15.Text
      End If
      'add by nickc 2007/01/25
      If IsNull(rsTmp.Fields("CE68")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE68 = CheckStr(rsTmp.Fields("CE68"))
'         tmpOldCE68 = CheckStr(rsTmp.Fields("CE68"))
      End If
      If IsNull(rsTmp.Fields("CE69")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE69 = CheckStr(rsTmp.Fields("CE69"))
'         tmpOldCE69 = CheckStr(rsTmp.Fields("CE69"))
      End If
      If IsNull(rsTmp.Fields("CE70")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE70 = CheckStr(rsTmp.Fields("CE70"))
'         tmpOldCE70 = CheckStr(rsTmp.Fields("CE70"))
      End If
      If IsNull(rsTmp.Fields("CE71")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE71 = CheckStr(rsTmp.Fields("CE71"))
'         tmpOldCE71 = CheckStr(rsTmp.Fields("CE71"))
      End If
      If IsNull(rsTmp.Fields("CE72")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE72 = CheckStr(rsTmp.Fields("CE72"))
'         tmpOldCE72 = CheckStr(rsTmp.Fields("CE72"))
      End If
      If IsNull(rsTmp.Fields("CE73")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE73 = CheckStr(rsTmp.Fields("CE73"))
'         tmpOldCE73 = CheckStr(rsTmp.Fields("CE73"))
      End If
      If IsNull(rsTmp.Fields("CE74")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE74 = CheckStr(rsTmp.Fields("CE74"))
'         tmpOldCE74 = CheckStr(rsTmp.Fields("CE74"))
      End If
      If IsNull(rsTmp.Fields("CE75")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE75 = CheckStr(rsTmp.Fields("CE75"))
'         tmpOldCE75 = CheckStr(rsTmp.Fields("CE75"))
      End If
      If IsNull(rsTmp.Fields("CE76")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE76 = CheckStr(rsTmp.Fields("CE76"))
'         tmpOldCE76 = CheckStr(rsTmp.Fields("CE76"))
      End If
      If IsNull(rsTmp.Fields("CE77")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE77 = CheckStr(rsTmp.Fields("CE77"))
'         tmpOldCE77 = CheckStr(rsTmp.Fields("CE77"))
      End If
      If IsNull(rsTmp.Fields("CE78")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE78 = CheckStr(rsTmp.Fields("CE78"))
'         tmpOldCE78 = CheckStr(rsTmp.Fields("CE78"))
      End If
      If IsNull(rsTmp.Fields("CE79")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE79 = CheckStr(rsTmp.Fields("CE79"))
'         tmpOldCE79 = CheckStr(rsTmp.Fields("CE79"))
      End If
      If IsNull(rsTmp.Fields("CE80")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE80 = CheckStr(rsTmp.Fields("CE80"))
'         tmpOldCE80 = CheckStr(rsTmp.Fields("CE80"))
      End If
      If IsNull(rsTmp.Fields("CE81")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE81 = CheckStr(rsTmp.Fields("CE81"))
'         tmpOldCE81 = CheckStr(rsTmp.Fields("CE81"))
      End If
      If IsNull(rsTmp.Fields("CE82")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE82 = CheckStr(rsTmp.Fields("CE82"))
'         tmpOldCE82 = CheckStr(rsTmp.Fields("CE82"))
      End If
      If IsNull(rsTmp.Fields("CE83")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE83 = CheckStr(rsTmp.Fields("CE83"))
'         tmpOldCE83 = CheckStr(rsTmp.Fields("CE83"))
      End If
      If IsNull(rsTmp.Fields("CE84")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE84 = CheckStr(rsTmp.Fields("CE84"))
'         tmpOldCE84 = CheckStr(rsTmp.Fields("CE84"))
      End If
      If IsNull(rsTmp.Fields("CE85")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE85 = CheckStr(rsTmp.Fields("CE85"))
'         tmpOldCE85 = CheckStr(rsTmp.Fields("CE85"))
      End If
      If IsNull(rsTmp.Fields("CE86")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE86 = CheckStr(rsTmp.Fields("CE86"))
'         tmpOldCE86 = CheckStr(rsTmp.Fields("CE86"))
      End If
      If IsNull(rsTmp.Fields("CE87")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE87 = CheckStr(rsTmp.Fields("CE87"))
'         tmpOldCE87 = CheckStr(rsTmp.Fields("CE87"))
      End If
      If IsNull(rsTmp.Fields("CE88")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE88 = CheckStr(rsTmp.Fields("CE88"))
'         tmpOldCE88 = CheckStr(rsTmp.Fields("CE88"))
      End If
      If IsNull(rsTmp.Fields("CE89")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE89 = CheckStr(rsTmp.Fields("CE89"))
'         tmpOldCE89 = CheckStr(rsTmp.Fields("CE89"))
      End If
      If IsNull(rsTmp.Fields("CE90")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE90 = CheckStr(rsTmp.Fields("CE90"))
'         tmpOldCE90 = CheckStr(rsTmp.Fields("CE90"))
      End If
      If IsNull(rsTmp.Fields("CE91")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/7
         textCE91 = CheckStr(rsTmp.Fields("CE91"))
'         tmpOldCE91 = CheckStr(rsTmp.Fields("CE91"))
      End If
      
      ' 申請地址
      If IsNull(rsTmp.Fields("CE23")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE23 = rsTmp.Fields("CE23")
'         tmpOldCE23 = textCE23.Text
      End If
      If IsNull(rsTmp.Fields("CE24")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE24 = rsTmp.Fields("CE24")
'         tmpOldCE24 = textCE24.Text
      End If
      If IsNull(rsTmp.Fields("CE25")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE25 = rsTmp.Fields("CE25")
'         tmpOldCE25 = textCE25.Text
      End If
      'add by nickc 2007/01/25
      If IsNull(rsTmp.Fields("CE26")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE26 = rsTmp.Fields("CE26")
'         tmpOldCE26 = textCE26.Text
      End If
      If IsNull(rsTmp.Fields("CE27")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE27 = rsTmp.Fields("CE27")
'         tmpOldCE27 = textCE27.Text
      End If
      If IsNull(rsTmp.Fields("CE28")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE28 = rsTmp.Fields("CE28")
'         tmpOldCE28 = textCE28.Text
      End If
      If IsNull(rsTmp.Fields("CE29")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE29 = rsTmp.Fields("CE29")
'         tmpOldCE29 = textCE29.Text
      End If
      If IsNull(rsTmp.Fields("CE30")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE30 = rsTmp.Fields("CE30")
'         tmpOldCE30 = textCE30.Text
      End If
      If IsNull(rsTmp.Fields("CE31")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE31 = rsTmp.Fields("CE31")
'         tmpOldCE31 = textCE31.Text
      End If
      If IsNull(rsTmp.Fields("CE32")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE32 = rsTmp.Fields("CE32")
'         tmpOldCE32 = textCE32.Text
      End If
      If IsNull(rsTmp.Fields("CE33")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE33 = rsTmp.Fields("CE33")
'         tmpOldCE33 = textCE33.Text
      End If
      If IsNull(rsTmp.Fields("CE34")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE34 = rsTmp.Fields("CE34")
'         tmpOldCE34 = textCE34.Text
      End If
      If IsNull(rsTmp.Fields("CE35")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE35 = rsTmp.Fields("CE35")
'         tmpOldCE35 = textCE35.Text
      End If
      If IsNull(rsTmp.Fields("CE36")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE36 = rsTmp.Fields("CE36")
'         tmpOldCE36 = textCE36.Text
      End If
      If IsNull(rsTmp.Fields("CE37")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/7
         textCE37 = rsTmp.Fields("CE37")
'         tmpOldCE37 = textCE37.Text
      End If
      ' 專利商標種類代號
      If IsNull(rsTmp.Fields("CE39")) = False Then
         checkCE40.Value = 1 'Add By Sindy 2012/3/7
         m_CE39 = rsTmp.Fields("CE39")
         textCE39 = rsTmp.Fields("CE39")
         If IsEmptyText(textCE39) = False Then: textCE39_Validate (False)
'         tmpOldCE39 = textCE39.Text
      End If
      ' 案件名稱
      If IsNull(rsTmp.Fields("CE41")) = False Then
        Select Case m_TM01
        Case "FCT", "S"
            checkCE44.Value = 1 'Add By Sindy 2012/3/7
            textCE41_1 = rsTmp.Fields("CE41")
'            tmpOldCE41 = textCE41_1.Text
        Case Else
            checkCE44.Value = 1 'Add By Sindy 2012/3/7
            textCE41 = rsTmp.Fields("CE41")
'            tmpOldCE41 = textCE41.Text
        End Select
      End If
      If IsNull(rsTmp.Fields("CE42")) = False Then
         checkCE44.Value = 1 'Add By Sindy 2012/3/7
         textCE42 = rsTmp.Fields("CE42")
'         tmpOldCE42 = textCE42.Text
      End If
      If IsNull(rsTmp.Fields("CE43")) = False Then
         checkCE44.Value = 1 'Add By Sindy 2012/3/7
         textCE43 = rsTmp.Fields("CE43")
'         tmpOldCE43 = textCE43.Text
      End If
      ' 縮減商品
      If IsNull(rsTmp.Fields("CE45")) = False Then
         checkCE46.Value = 1 'Add By Sindy 2012/3/7
         textCE45 = rsTmp.Fields("CE45")
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("CE47")) = False Then
         checkCE48.Value = 1 'Add By Sindy 2012/3/7
         textCE47 = rsTmp.Fields("CE47")
'         tmpOldCE47 = textCE47.Text
      End If
      ' 商品群組
      If IsNull(rsTmp.Fields("CE49")) = False Then
         checkCE50.Value = 1 'Add By Sindy 2012/3/7
         textCE49 = rsTmp.Fields("CE49")
'         tmpOldCE49 = textCE49.Text
      End If
      ' 申請人印鑑
      If IsNull(rsTmp.Fields("CE51")) = False Then
         checkCE52.Value = 1 'Add By Sindy 2012/3/7
         'textCE51 = rsTmp.Fields("CE51")
      End If
      ' 代表人印鑑
      If IsNull(rsTmp.Fields("CE53")) = False Then
         checkCE54.Value = 1 'Add By Sindy 2012/3/7
         'textCE53 = rsTmp.Fields("CE53")
      End If
      ' 代理人
      If IsNull(rsTmp.Fields("CE55")) = False Then
         checkCE56.Value = 1 'Add By Sindy 2012/3/7
         'textCE55 = rsTmp.Fields("CE55")
      End If
      ' 正商標號數
      If IsNull(rsTmp.Fields("CE57")) = False Then
         checkCE58.Value = 1 'Add By Sindy 2012/3/7
         textCE57 = rsTmp.Fields("CE57")
'         tmpOldCE57 = textCE57.Text
      End If
      ' 圖樣
      If IsNull(rsTmp.Fields("CE59")) = False Then
         checkCE60.Value = 1 'Add By Sindy 2012/3/7
         'textCE59 = rsTmp.Fields("CE59")
      End If
      ' 其它
      If IsNull(rsTmp.Fields("CE61")) = False Then
         checkCE62.Value = 1 'Add By Sindy 2012/3/7
         textCE61 = rsTmp.Fields("CE61")
      End If
      ' 代表人譯文
      If IsNull(rsTmp.Fields("CE63")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE63 = rsTmp.Fields("CE63")
      End If
      If IsNull(rsTmp.Fields("CE64")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE64 = rsTmp.Fields("CE64")
      End If
      'add by nickc 2007/01/25
      If IsNull(rsTmp.Fields("CE92")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE92 = CheckStr(rsTmp.Fields("CE92"))
      End If
      If IsNull(rsTmp.Fields("CE93")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE93 = CheckStr(rsTmp.Fields("CE93"))
      End If
      If IsNull(rsTmp.Fields("CE94")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE94 = CheckStr(rsTmp.Fields("CE94"))
      End If
      If IsNull(rsTmp.Fields("CE95")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE95 = CheckStr(rsTmp.Fields("CE95"))
      End If
      If IsNull(rsTmp.Fields("CE96")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE96 = CheckStr(rsTmp.Fields("CE96"))
      End If
      If IsNull(rsTmp.Fields("CE97")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE97 = CheckStr(rsTmp.Fields("CE97"))
      End If
      If IsNull(rsTmp.Fields("CE98")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE98 = CheckStr(rsTmp.Fields("CE98"))
      End If
      If IsNull(rsTmp.Fields("CE99")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/7
         textCE99 = CheckStr(rsTmp.Fields("CE99"))
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   Call QueryCaseProgress 'Add By Sindy 2011/8/23
End Sub

Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'Add By Sindy 2011/8/23 是否新案件
      m_CP31 = ""
      If IsNull(rsTmp.Fields("CP31")) = False Then
         m_CP31 = rsTmp.Fields("CP31")
      End If
      'Add By Sindy 2018/2/1 案件性質
      m_CP10 = ""
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
      End If
      '2018/2/1 END
   End If
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strSql As String
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim rsTmp As New ADODB.Recordset
   
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' 先刪除掉已存在的資料
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.Close
      strSql = "DELETE FROM ChangeEvent " & _
               "WHERE CE01 = '" & m_CE01 & "' "
      cnnConnection.Execute strSql
   Else
      rsTmp.Close
   End If
            
   ' 新增一筆資料到變更事項檔
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO ChangeEvent ("
   For nIndex = 0 To m_CECount - 1
      strTmp = m_CEList(nIndex).fiName
      If IsEmptyText(strTmp) = False And IsEmptyText(m_CEList(nIndex).fiNewData) = False Then
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
   For nIndex = 0 To m_CECount - 1
      strTmp = Empty
      If m_CEList(nIndex).fiType = 0 Then
        'Modify By Cheng 2003/12/08
'         If IsEmptyText(m_CEList(nIndex).fiNewData) = False Then: strTmp = "'" & m_CEList(nIndex).fiNewData & "'"
         If IsEmptyText(m_CEList(nIndex).fiNewData) = False Then: strTmp = "'" & ChgSQL(m_CEList(nIndex).fiNewData) & "'"
      Else
         strTmp = m_CEList(nIndex).fiNewData
      End If
      If IsEmptyText(m_CEList(nIndex).fiName) = False And IsEmptyText(m_CEList(nIndex).fiNewData) = False Then
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
   End If
    'Add By Cheng 2003/04/11
    Select Case m_TM01
    Case "CFT", "FCT", "T", "TF"
        If OnSaveTrademark = False Then GoTo CheckingErr
    Case Else
        If OnSaveServicePractice = False Then GoTo CheckingErr
    End Select
   
   Set rsTmp = Nothing
   
 '911107 nick transation
  cnnConnection.CommitTrans
     Exit Function
CheckingErr:
     cnnConnection.RollbackTrans
    MsgBox (Err.Description)
    OnSaveData = False
End Function

' 申請日
Private Sub textCE02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(textCE02) = False Then
      If CheckIsTaiwanDate(textCE02, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請輸入正確的申請日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE02_GotFocus
      End If
   End If
End Sub

'add by nickc 2007/01/25
Private Sub textCE04_KeyPress(KeyAscii As Integer)
 KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2009/07/02
Private Sub txt2_GotFocus(Index As Integer)
   InverseTextBox txt2(Index)
End Sub
Private Sub txt2_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
   Case 0
      KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

' 申請人
Private Sub textCE04_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    textCE04_2 = Empty
    If IsEmptyText(textCE04) = False Then
        'Add By Cheng 2003/04/14
        '申請人編號補滿9碼
        Me.textCE04.Text = Left(Me.textCE04.Text & "000000000", 9)
        'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
        Dim oState As Boolean
        oState = True
        'textCE04_2 = GetCustomerName(textCE04)
        textCE04_2 = GetCustomerNameAndState(textCE04, "0", oState)
        If oState = False Then
            Cancel = True
            Exit Sub
        End If
        If textCE04_2 = Empty Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請人代碼<" & textCE04 & ">不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCE04_GotFocus
        End If
    End If
End Sub

'add by nickc 2007/01/25
Private Sub textCE05_GotFocus()
InverseTextBox textCE05
End Sub
Private Sub textCE05_KeyPress(KeyAscii As Integer)
 KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textCE05_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    textCE05_2 = Empty
    If IsEmptyText(textCE05) = False Then
        Me.textCE05.Text = Left(Me.textCE05.Text & "000000000", 9)
        Dim oState As Boolean
        oState = True
        textCE05_2 = GetCustomerNameAndState(textCE05, "0", oState)
        If oState = False Then
            Cancel = True
            Exit Sub
        End If
        If textCE05_2 = Empty Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請人代碼<" & textCE05 & ">不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCE05_GotFocus
        End If
    End If
End Sub
Private Sub textCE06_GotFocus()
InverseTextBox textCE06
End Sub
Private Sub textCE06_KeyPress(KeyAscii As Integer)
 KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textCE06_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    textCE06_2 = Empty
    If IsEmptyText(textCE06) = False Then
        Me.textCE06.Text = Left(Me.textCE06.Text & "000000000", 9)
        Dim oState As Boolean
        oState = True
        textCE06_2 = GetCustomerNameAndState(textCE06, "0", oState)
        If oState = False Then
            Cancel = True
            Exit Sub
        End If
        If textCE06_2 = Empty Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請人代碼<" & textCE06 & ">不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCE06_GotFocus
        End If
    End If
End Sub
Private Sub textCE07_GotFocus()
InverseTextBox textCE07
End Sub
Private Sub textCE07_KeyPress(KeyAscii As Integer)
 KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textCE07_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    textCE07_2 = Empty
    If IsEmptyText(textCE07) = False Then
        Me.textCE07.Text = Left(Me.textCE07.Text & "000000000", 9)
        Dim oState As Boolean
        oState = True
        textCE07_2 = GetCustomerNameAndState(textCE07, "0", oState)
        If oState = False Then
            Cancel = True
            Exit Sub
        End If
        If textCE07_2 = Empty Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請人代碼<" & textCE07 & ">不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCE07_GotFocus
        End If
    End If
End Sub
Private Sub textCE08_GotFocus()
InverseTextBox textCE08
End Sub
Private Sub textCE08_KeyPress(KeyAscii As Integer)
 KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textCE08_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    textCE08_2 = Empty
    If IsEmptyText(textCE08) = False Then
        Me.textCE08.Text = Left(Me.textCE08.Text & "000000000", 9)
        Dim oState As Boolean
        oState = True
        textCE08_2 = GetCustomerNameAndState(textCE08, "0", oState)
        If oState = False Then
            Cancel = True
            Exit Sub
        End If
        If textCE08_2 = Empty Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請人代碼<" & textCE08 & ">不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCE08_GotFocus
        End If
    End If
End Sub
Private Sub textCE18_GotFocus()
InverseTextBox textCE18
End Sub
Private Sub textCE19_GotFocus()
InverseTextBox textCE19
End Sub
Private Sub textCE20_GotFocus()
InverseTextBox textCE20
End Sub
Private Sub textCE21_GotFocus()
InverseTextBox textCE21
End Sub
'add by nickc 2007/01/25
Private Sub textCE26_GotFocus()
InverseTextBox textCE26
End Sub
Private Sub textCE27_GotFocus()
InverseTextBox textCE27
End Sub
Private Sub textCE28_GotFocus()
InverseTextBox textCE28
End Sub
Private Sub textCE29_GotFocus()
InverseTextBox textCE29
End Sub
Private Sub textCE30_GotFocus()
InverseTextBox textCE30
End Sub
Private Sub textCE31_GotFocus()
InverseTextBox textCE31
End Sub
Private Sub textCE32_GotFocus()
InverseTextBox textCE32
End Sub
Private Sub textCE33_GotFocus()
InverseTextBox textCE33
End Sub
Private Sub textCE34_GotFocus()
InverseTextBox textCE34
End Sub
Private Sub textCE35_GotFocus()
InverseTextBox textCE35
End Sub
Private Sub textCE36_GotFocus()
InverseTextBox textCE36
End Sub
Private Sub textCE37_GotFocus()
InverseTextBox textCE37
End Sub

Private Sub textCE39_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   textCE39_2 = Empty
   Cancel = False
   If IsEmptyText(textCE39) = False Then
      textCE39_2 = GetTradeMarkName(textCE39, 0)
      If IsEmptyText(textCE39_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商標種類不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE39_GotFocus
      End If
   End If
End Sub

Private Sub textCE02_GotFocus()
   InverseTextBox textCE02
End Sub

Private Sub textCE04_GotFocus()
   InverseTextBox textCE04
End Sub

Private Sub textCE10_GotFocus()
   InverseTextBox textCE10
End Sub

Private Sub textCE11_GotFocus()
   InverseTextBox textCE11
End Sub

Private Sub textCE12_GotFocus()
   InverseTextBox textCE12
End Sub

Private Sub textCE13_GotFocus()
   InverseTextBox textCE13
End Sub

Private Sub textCE14_GotFocus()
   InverseTextBox textCE14
End Sub

Private Sub textCE15_GotFocus()
   InverseTextBox textCE15
End Sub

Private Sub textCE17_GotFocus()
   InverseTextBox textCE17
End Sub

Private Sub textCE23_GotFocus()
   InverseTextBox textCE23
End Sub

Private Sub textCE24_GotFocus()
   InverseTextBox textCE24
End Sub

Private Sub textCE25_GotFocus()
   InverseTextBox textCE25
End Sub

Private Sub textCE39_GotFocus()
   InverseTextBox textCE39
End Sub

Private Sub textCE41_1_GotFocus()
    TextInverse Me.textCE41_1
End Sub

Private Sub textCE41_GotFocus()
   InverseTextBox textCE41
End Sub

Private Sub textCE42_GotFocus()
   InverseTextBox textCE42
End Sub

Private Sub textCE43_GotFocus()
   InverseTextBox textCE43
End Sub

Private Sub textCE45_GotFocus()
   InverseTextBox textCE45
End Sub

Private Sub textCE47_GotFocus()
   InverseTextBox textCE47
End Sub

Private Sub textCE47_Validate(Cancel As Boolean)
'add by nickc 2005/06/03
textCE47 = Replace(textCE47, " ", "")
End Sub

Private Sub textCE49_GotFocus()
   InverseTextBox textCE49
End Sub

Private Sub textCE57_GotFocus()
   InverseTextBox textCE57
End Sub

Private Sub textCE61_GotFocus()
   InverseTextBox textCE61
End Sub

Private Sub textCE63_GotFocus()
   InverseTextBox textCE63
End Sub

Private Sub textCE64_GotFocus()
   InverseTextBox textCE64
End Sub

'Add By Cheng 2003/04/10
Private Function OnSaveTrademark() As Boolean
Dim strSql As String
Dim bFirst As Boolean
Dim bDifference As Boolean
Dim strTmp As String
Dim nIndex As Integer
Dim tmpCp64 As String
Dim rsnick911204 As New ADODB.Recordset
   
On Error GoTo ErrorHandler
    
    OnSaveTrademark = True
    tmpCp64 = ""
    tmpCp64 = " select cp64 from caseprogress where cP09= '" & m_CE01 & "'"
    Set rsnick911204 = New ADODB.Recordset
    rsnick911204.CursorLocation = adUseClient
    rsnick911204.Open tmpCp64, cnnConnection, adOpenStatic, adLockReadOnly
    tmpCp64 = ""
    If rsnick911204.RecordCount > 0 Then
         tmpCp64 = CheckStr(rsnick911204.Fields(0).Value) & " "
    End If
    ' 申請人
'    If checkCE09.Value = True Then
    If checkCE09.Value = vbChecked Then
       If tmpOldCE04 <> textCE04 Then
             SetTMFieldData "TM23", textCE04, 0
'             tmpCp64 = tmpCp64 & "原申請人:" & tmpOldCE04 & " "
             If tmpOldCE04 <> "" Then tmpCp64 = tmpCp64 & "原申請人1:" & tmpOldCE04 & " "
            '申請地址
             SetTMFieldData "TM24", PUB_GetCustEachAdd(textCE04, 1), 0
             SetTMFieldData "TM25", PUB_GetCustEachAdd(textCE04, 2), 0
             SetTMFieldData "TM26", PUB_GetCustEachAdd(textCE04, 3), 0
       End If
       'add by nickc 2007/01/25
       If tmpOldCE05 <> textCE05 Then
             SetTMFieldData "TM78", textCE05, 0
             If tmpOldCE05 <> "" Then tmpCp64 = tmpCp64 & "原申請人2:" & tmpOldCE05 & " "
            '申請地址
             SetTMFieldData "TM82", PUB_GetCustEachAdd(textCE05, 1), 0
             SetTMFieldData "TM83", PUB_GetCustEachAdd(textCE05, 2), 0
             SetTMFieldData "TM84", PUB_GetCustEachAdd(textCE05, 3), 0
       End If
       If tmpOldCE06 <> textCE06 Then
             SetTMFieldData "TM79", textCE06, 0
             If tmpOldCE06 <> "" Then tmpCp64 = tmpCp64 & "原申請人3:" & tmpOldCE06 & " "
            '申請地址
             SetTMFieldData "TM85", PUB_GetCustEachAdd(textCE06, 1), 0
             SetTMFieldData "TM86", PUB_GetCustEachAdd(textCE06, 2), 0
             SetTMFieldData "TM87", PUB_GetCustEachAdd(textCE06, 3), 0
       End If
       If tmpOldCE07 <> textCE07 Then
             SetTMFieldData "TM80", textCE07, 0
             If tmpOldCE07 <> "" Then tmpCp64 = tmpCp64 & "原申請人4:" & tmpOldCE07 & " "
            '申請地址
             SetTMFieldData "TM88", PUB_GetCustEachAdd(textCE07, 1), 0
             SetTMFieldData "TM89", PUB_GetCustEachAdd(textCE07, 2), 0
             SetTMFieldData "TM90", PUB_GetCustEachAdd(textCE07, 3), 0
       End If
       If tmpOldCE08 <> textCE08 Then
             SetTMFieldData "TM81", textCE08, 0
             If tmpOldCE08 <> "" Then tmpCp64 = tmpCp64 & "原申請人5:" & tmpOldCE08 & " "
            '申請地址
             SetTMFieldData "TM91", PUB_GetCustEachAdd(textCE08, 1), 0
             SetTMFieldData "TM92", PUB_GetCustEachAdd(textCE08, 2), 0
             SetTMFieldData "TM93", PUB_GetCustEachAdd(textCE08, 3), 0
       End If
       
    End If
    ' 申請日
'    If checkCE03.Value = True Then
    If checkCE03.Value = vbChecked Then
       If tmpOldCE02 <> DBDATE(textCE02) Then
             SetTMFieldData "TM11", DBDATE(textCE02), 1
'             tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
             If tmpOldCE02 <> "" Then tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
       End If
    End If
    '911204 nick 新增代表人 只判斷是否有變更
    If checkCE16.Value = 1 Then
       If textCE10 <> tmpOldCE10 Then
             SetTMFieldData "TM47", textCE10, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE10 & " "
             If tmpOldCE10 <> "" Then tmpCp64 = tmpCp64 & "原代表人1(中):" & tmpOldCE10 & " "
       End If
       If textCE11 <> tmpOldCE11 Then
             SetTMFieldData "TM48", textCE11, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE11 & " "
             If tmpOldCE11 <> "" Then tmpCp64 = tmpCp64 & "原代表人1(英):" & tmpOldCE11 & " "
       End If
       If textCE12 <> tmpOldCE12 Then
             SetTMFieldData "TM49", textCE12, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE12 & " "
             If tmpOldCE12 <> "" Then tmpCp64 = tmpCp64 & "原代表人1(日):" & tmpOldCE12 & " "
       End If
       If textCE13 <> tmpOldCE13 Then
             SetTMFieldData "TM50", textCE13, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE13 & " "
             If tmpOldCE13 <> "" Then tmpCp64 = tmpCp64 & "原代表人2(中):" & tmpOldCE13 & " "
       End If
       If textCE14 <> tmpOldCE14 Then
             SetTMFieldData "TM51", textCE14, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE14 & " "
             If tmpOldCE14 <> "" Then tmpCp64 = tmpCp64 & "原代表人2(英):" & tmpOldCE14 & " "
       End If
       If textCE15 <> tmpOldCE15 Then
             SetTMFieldData "TM52", textCE15, 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE15 & " "
             If tmpOldCE15 <> "" Then tmpCp64 = tmpCp64 & "原代表人2(日):" & tmpOldCE15 & " "
       End If
       'add by nickc 2007/01/26
       If textCE68 <> tmpOldCE68 Then
             SetTMFieldData "TM94", textCE68, 0
             If tmpOldCE68 <> "" Then tmpCp64 = tmpCp64 & "原代表人3(中):" & tmpOldCE68 & " "
       End If
       If textCE69 <> tmpOldCE69 Then
             SetTMFieldData "TM95", textCE69, 0
             If tmpOldCE69 <> "" Then tmpCp64 = tmpCp64 & "原代表人3(英):" & tmpOldCE69 & " "
       End If
       If textCE70 <> tmpOldCE70 Then
             SetTMFieldData "TM96", textCE70, 0
             If tmpOldCE70 <> "" Then tmpCp64 = tmpCp64 & "原代表人3(日):" & tmpOldCE70 & " "
       End If
       If textCE71 <> tmpOldCE71 Then
             SetTMFieldData "TM97", textCE71, 0
             If tmpOldCE71 <> "" Then tmpCp64 = tmpCp64 & "原代表人4(中):" & tmpOldCE71 & " "
       End If
       If textCE72 <> tmpOldCE72 Then
             SetTMFieldData "TM98", textCE72, 0
             If tmpOldCE72 <> "" Then tmpCp64 = tmpCp64 & "原代表人4(英):" & tmpOldCE72 & " "
       End If
       If textCE73 <> tmpOldCE73 Then
             SetTMFieldData "TM99", textCE73, 0
             If tmpOldCE73 <> "" Then tmpCp64 = tmpCp64 & "原代表人4(日):" & tmpOldCE73 & " "
       End If
       If textCE74 <> tmpOldCE74 Then
             SetTMFieldData "TM100", textCE74, 0
             If tmpOldCE74 <> "" Then tmpCp64 = tmpCp64 & "原代表人5(中):" & tmpOldCE74 & " "
       End If
       If textCE75 <> tmpOldCE75 Then
             SetTMFieldData "TM101", textCE75, 0
             If tmpOldCE75 <> "" Then tmpCp64 = tmpCp64 & "原代表人5(英):" & tmpOldCE75 & " "
       End If
       If textCE76 <> tmpOldCE76 Then
             SetTMFieldData "TM102", textCE76, 0
             If tmpOldCE76 <> "" Then tmpCp64 = tmpCp64 & "原代表人5(日):" & tmpOldCE76 & " "
       End If
       If textCE77 <> tmpOldCE77 Then
             SetTMFieldData "TM103", textCE77, 0
             If tmpOldCE77 <> "" Then tmpCp64 = tmpCp64 & "原代表人6(中):" & tmpOldCE77 & " "
       End If
       If textCE78 <> tmpOldCE78 Then
             SetTMFieldData "TM104", textCE78, 0
             If tmpOldCE78 <> "" Then tmpCp64 = tmpCp64 & "原代表人6(英):" & tmpOldCE78 & " "
       End If
       If textCE79 <> tmpOldCE79 Then
             SetTMFieldData "TM105", textCE79, 0
             If tmpOldCE79 <> "" Then tmpCp64 = tmpCp64 & "原代表人6(日):" & tmpOldCE79 & " "
       End If
       If textCE80 <> tmpOldCE80 Then
             SetTMFieldData "TM106", textCE80, 0
             If tmpOldCE80 <> "" Then tmpCp64 = tmpCp64 & "原代表人7(中):" & tmpOldCE80 & " "
       End If
       If textCE81 <> tmpOldCE81 Then
             SetTMFieldData "TM107", textCE81, 0
             If tmpOldCE81 <> "" Then tmpCp64 = tmpCp64 & "原代表人7(英):" & tmpOldCE81 & " "
       End If
       If textCE82 <> tmpOldCE82 Then
             SetTMFieldData "TM108", textCE82, 0
             If tmpOldCE82 <> "" Then tmpCp64 = tmpCp64 & "原代表人7(日):" & tmpOldCE82 & " "
       End If
       If textCE83 <> tmpOldCE83 Then
             SetTMFieldData "TM109", textCE83, 0
             If tmpOldCE83 <> "" Then tmpCp64 = tmpCp64 & "原代表人8(中):" & tmpOldCE83 & " "
       End If
       If textCE84 <> tmpOldCE84 Then
             SetTMFieldData "TM110", textCE84, 0
             If tmpOldCE84 <> "" Then tmpCp64 = tmpCp64 & "原代表人8(英):" & tmpOldCE84 & " "
       End If
       If textCE85 <> tmpOldCE85 Then
             SetTMFieldData "TM111", textCE85, 0
             If tmpOldCE85 <> "" Then tmpCp64 = tmpCp64 & "原代表人8(日):" & tmpOldCE85 & " "
       End If
       If textCE86 <> tmpOldCE86 Then
             SetTMFieldData "TM112", textCE86, 0
             If tmpOldCE86 <> "" Then tmpCp64 = tmpCp64 & "原代表人9(中):" & tmpOldCE86 & " "
       End If
       If textCE87 <> tmpOldCE87 Then
             SetTMFieldData "TM113", textCE87, 0
             If tmpOldCE87 <> "" Then tmpCp64 = tmpCp64 & "原代表人9(英):" & tmpOldCE87 & " "
       End If
       If textCE88 <> tmpOldCE88 Then
             SetTMFieldData "TM114", textCE88, 0
             If tmpOldCE88 <> "" Then tmpCp64 = tmpCp64 & "原代表人9(日):" & tmpOldCE88 & " "
       End If
       If textCE89 <> tmpOldCE89 Then
             SetTMFieldData "TM115", textCE89, 0
             If tmpOldCE89 <> "" Then tmpCp64 = tmpCp64 & "原代表人10(中):" & tmpOldCE89 & " "
       End If
       If textCE90 <> tmpOldCE90 Then
             SetTMFieldData "TM116", textCE90, 0
             If tmpOldCE90 <> "" Then tmpCp64 = tmpCp64 & "原代表人10(英):" & tmpOldCE90 & " "
       End If
       If textCE91 <> tmpOldCE91 Then
             SetTMFieldData "TM117", textCE91, 0
             If tmpOldCE91 <> "" Then tmpCp64 = tmpCp64 & "原代表人10(日):" & tmpOldCE91 & " "
       End If
    End If
    ' 申請地址
'    If checkCE38.Value = True Then
    If checkCE38.Value = vbChecked Then
         If tmpOldCE23 <> textCE23 Then
             SetTMFieldData "TM24", textCE23, 0
'             tmpCp64 = tmpCp64 & "原申請中文地址:" & tmpOldCE23 & " "
             If tmpOldCE23 <> "" Then tmpCp64 = tmpCp64 & "原申請中文地址1:" & tmpOldCE23 & " "
         End If
         If tmpOldCE24 <> textCE24 Then
             SetTMFieldData "TM25", textCE24, 0
'             tmpCp64 = tmpCp64 & "原申請英文地址:" & tmpOldCE24 & " "
             If tmpOldCE24 <> "" Then tmpCp64 = tmpCp64 & "原申請英文地址1:" & tmpOldCE24 & " "
         End If
         If tmpOldCE25 <> textCE25 Then
             SetTMFieldData "TM26", textCE25, 0
'             tmpCp64 = tmpCp64 & "原申請日文地址:" & tmpOldCE25 & " "
             If tmpOldCE25 <> "" Then tmpCp64 = tmpCp64 & "原申請日文地址1:" & tmpOldCE25 & " "
         End If
         'add by nickc 2007/01/26
         If tmpOldCE26 <> textCE26 Then
             SetTMFieldData "TM82", textCE26, 0
             If tmpOldCE26 <> "" Then tmpCp64 = tmpCp64 & "原申請中文地址2:" & tmpOldCE26 & " "
         End If
         If tmpOldCE27 <> textCE27 Then
             SetTMFieldData "TM83", textCE27, 0
             If tmpOldCE27 <> "" Then tmpCp64 = tmpCp64 & "原申請英文地址2:" & tmpOldCE27 & " "
         End If
         If tmpOldCE28 <> textCE28 Then
             SetTMFieldData "TM84", textCE28, 0
             If tmpOldCE28 <> "" Then tmpCp64 = tmpCp64 & "原申請日文地址2:" & tmpOldCE28 & " "
         End If
         If tmpOldCE29 <> textCE29 Then
             SetTMFieldData "TM85", textCE29, 0
             If tmpOldCE29 <> "" Then tmpCp64 = tmpCp64 & "原申請中文地址3:" & tmpOldCE29 & " "
         End If
         If tmpOldCE30 <> textCE30 Then
             SetTMFieldData "TM86", textCE30, 0
             If tmpOldCE30 <> "" Then tmpCp64 = tmpCp64 & "原申請英文地址3:" & tmpOldCE30 & " "
         End If
         If tmpOldCE31 <> textCE31 Then
             SetTMFieldData "TM87", textCE31, 0
             If tmpOldCE31 <> "" Then tmpCp64 = tmpCp64 & "原申請日文地址3:" & tmpOldCE31 & " "
         End If
         If tmpOldCE32 <> textCE32 Then
             SetTMFieldData "TM88", textCE32, 0
             If tmpOldCE32 <> "" Then tmpCp64 = tmpCp64 & "原申請中文地址4:" & tmpOldCE32 & " "
         End If
         If tmpOldCE33 <> textCE33 Then
             SetTMFieldData "TM89", textCE33, 0
             If tmpOldCE33 <> "" Then tmpCp64 = tmpCp64 & "原申請英文地址4:" & tmpOldCE33 & " "
         End If
         If tmpOldCE34 <> textCE34 Then
             SetTMFieldData "TM90", textCE34, 0
             If tmpOldCE34 <> "" Then tmpCp64 = tmpCp64 & "原申請日文地址4:" & tmpOldCE34 & " "
         End If
         If tmpOldCE35 <> textCE35 Then
             SetTMFieldData "TM91", textCE35, 0
             If tmpOldCE35 <> "" Then tmpCp64 = tmpCp64 & "原申請中文地址5:" & tmpOldCE35 & " "
         End If
         If tmpOldCE36 <> textCE36 Then
             SetTMFieldData "TM92", textCE36, 0
             If tmpOldCE36 <> "" Then tmpCp64 = tmpCp64 & "原申請英文地址5:" & tmpOldCE36 & " "
         End If
         If tmpOldCE37 <> textCE37 Then
             SetTMFieldData "TM93", textCE37, 0
             If tmpOldCE37 <> "" Then tmpCp64 = tmpCp64 & "原申請日文地址5:" & tmpOldCE37 & " "
         End If
    End If
    '正商標號數
'    If checkCE58.Value = True Then
    If checkCE58.Value = vbChecked Then
         If tmpOldCE57 <> textCE57 Then
             SetTMFieldData "TM27", textCE57, 0
'             tmpCp64 = tmpCp64 & "原正商標號數:" & tmpOldCE57 & " "
             If tmpOldCE57 <> "" Then tmpCp64 = tmpCp64 & "原正商標號數:" & tmpOldCE57 & " "
         End If
    End If
    '商標種類
'    If checkCE40.Value = True Then
    If checkCE40.Value = vbChecked Then
         If tmpOldCE39 <> textCE39 Then
             SetTMFieldData "TM08", textCE39, 0
'             tmpCp64 = tmpCp64 & "原商標種類:" & tmpOldCE39 & " "
             If tmpOldCE39 <> "" Then tmpCp64 = tmpCp64 & "原商標種類:" & tmpOldCE39 & " "
            'Add By Cheng 2003/09/09
            '聯合商標變更為正商標, 清除基本檔的正商標號數
            If (tmpOldCE39 = "2" And Me.textCE39.Text = "1") Or (tmpOldCE39 = "5" And Me.textCE39.Text = "4") Then
                SetTMFieldData "TM27", "", 0
            End If
         End If
    End If
    ' 案件名稱
'    If checkCE44.Value = True Then
    If checkCE44.Value = vbChecked Then
        Select Case m_TM01
        Case "FCT"
             If tmpOldCE41 <> textCE41_1 Then
                 SetTMFieldData "TM05", textCE41_1, 0
                 If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
             End If
        Case Else
             '911204 nick
             If tmpOldCE41 <> textCE41 Then
                 SetTMFieldData "TM05", textCE41, 0
                 '911204 nick
    '             tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
                 If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
             End If
        End Select
         '911204 nick
         If tmpOldCE42 <> textCE42 Then
             SetTMFieldData "TM06", textCE42, 0
             '911204 nick
'             tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
             If tmpOldCE42 <> "" Then tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
         End If
         '911204 nick
         If tmpOldCE43 <> textCE43 Then
             SetTMFieldData "TM07", textCE43, 0
             '911204 nick
'             tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
             If tmpOldCE43 <> "" Then tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
         End If
    End If
    '商品類別
'    If checkCE48.Value = True Then
    If checkCE48.Value = vbChecked Then
         If tmpOldCE47 <> textCE47 Then
             SetTMFieldData "TM09", textCE47, 0
'             tmpCp64 = tmpCp64 & "原商標類別:" & tmpOldCE47 & " "
             If tmpOldCE47 <> "" Then tmpCp64 = tmpCp64 & "原商標類別:" & tmpOldCE47 & " "
         End If
    End If
    '商品群組
'    If checkCE50.Value = True Then
    If checkCE50.Value = vbChecked Then
         If tmpOldCE49 <> textCE49 Then
             SetTMFieldData "TM32", textCE49, 0
'             tmpCp64 = tmpCp64 & "原商標群組:" & tmpOldCE49 & " "
             If tmpOldCE49 <> "" Then tmpCp64 = tmpCp64 & "原商標群組:" & tmpOldCE49 & " "
         End If
    End If
    ' 更新商標基本檔
    strSql = "UPDATE Trademark SET "
    bFirst = True
    bDifference = False
    For nIndex = 0 To m_TMListCount - 1
       strTmp = Empty
       If m_TMList(nIndex).tiType = 0 Then
          'edit by nickc 2005/06/13
          'strTmp = m_TMList(nIndex).tiName & " = '" & m_TMList(nIndex).tiData & "'"
          strTmp = m_TMList(nIndex).tiName & " = '" & ChgSQL(m_TMList(nIndex).tiData) & "'"
       Else
          If m_TMList(nIndex).tiData = Empty Then
             strTmp = m_TMList(nIndex).tiName & " = " & 0
          Else
             'edit by nickc 2005/06/14
             'strTmp = m_TMList(nIndex).tiName & " = " & m_TMList(nIndex).tiData
             strTmp = m_TMList(nIndex).tiName & " = " & ChgSQL(m_TMList(nIndex).tiData)
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
    ' 組成SQL語法
    strSql = strSql & " " & _
                   "WHERE TM01 = '" & m_TM01 & "' AND " & _
                         "TM02 = '" & m_TM02 & "' AND " & _
                         "TM03 = '" & m_TM03 & "' AND " & _
                         "TM04 = '" & m_TM04 & "'"
    ' 執行SQL指令
    If bDifference = True Then
       cnnConnection.Execute strSql
       '911226 nick 更新回原本收文號的備註
       'edit by nickc 2005/06/14
       'StrSql = "update caseprogress set cp64='" & tmpCp64 & "' where cp09='" & m_CE01 & "' "
       strSql = "update caseprogress set cp64='" & ChgSQL(tmpCp64) & "' where cp09='" & m_CE01 & "' "
       cnnConnection.Execute strSql
    End If
    
    ' 清除所佔用的記憶體
    If m_TMListCount > 0 Then
       Erase m_TMList
       m_TMListCount = 0
    End If
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnSaveTrademark = False
End Function

' 設定欄位新值
Private Sub SetTMFieldData(ByVal strField As String, ByVal strNewData As String, ByVal nType As Integer)
   Dim bFind As Boolean
   Dim nIndex As Integer
   ' 搜尋是否存在該欄位
   bFind = False
   For nIndex = 0 To m_TMListCount - 1
      If m_TMList(nIndex).tiName = strField Then
         bFind = True
         m_TMList(nIndex).tiData = strNewData
         Exit For
      End If
   Next nIndex
   ' 不存在則新增該欄位
   If bFind = False Then
      ReDim Preserve m_TMList(m_TMListCount + 1)
      m_TMList(m_TMListCount).tiName = strField
      m_TMList(m_TMListCount).tiData = strNewData
      m_TMList(m_TMListCount).tiType = nType
      m_TMListCount = m_TMListCount + 1
   End If
End Sub


' 91.09.02 modify by louis
'Modify By Cheng 2002/11/06
'Private Sub OnSaveServicePractice()
Private Function OnSaveServicePractice() As Boolean
   Dim strSql As String
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim strTmp As String
   Dim nIndex As Integer
   '911204 nick
   Dim tmpCp64 As String
   Dim rsnick911204 As New ADODB.Recordset
   
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnSaveServicePractice = True

   '911204 nick
   tmpCp64 = ""
    'Modify By Cheng 2003/04/10
    '取消限制
'   ' 只有系統類別為TC及案件性質為變更301才更新基本檔
'   If m_TM01 <> "TC" Or m_CP10 <> "301" Then
'      Exit Function
'   End If
   
   '911204 nick
   tmpCp64 = " select cp64 from caseprogress where cP09= '" & m_CE01 & "'"
   Set rsnick911204 = New ADODB.Recordset
   rsnick911204.CursorLocation = adUseClient
   rsnick911204.Open tmpCp64, cnnConnection, adOpenStatic, adLockReadOnly
   tmpCp64 = ""
   If rsnick911204.RecordCount > 0 Then
        tmpCp64 = CheckStr(rsnick911204.Fields(0).Value) & " "
   End If
   
   ' 申請人
'   If checkCE09.Value = True Then
   If checkCE09.Value = vbChecked Then
      '911204 nick
      If tmpOldCE04 <> textCE04 Then
            SetSRFieldData "SP08", textCE04, 0
            '911204 nick
'            tmpCp64 = tmpCp64 & "原申請人:" & tmpOldCE04 & " "
            If tmpOldCE04 <> "" Then tmpCp64 = tmpCp64 & "原申請人1:" & tmpOldCE04 & " "
      End If
      'add by nickc 2007/01/26
      If tmpOldCE05 <> textCE05 Then
            SetSRFieldData "SP58", textCE05, 0
            If tmpOldCE05 <> "" Then tmpCp64 = tmpCp64 & "原申請人2:" & tmpOldCE05 & " "
      End If
      If tmpOldCE06 <> textCE06 Then
            SetSRFieldData "SP59", textCE06, 0
            If tmpOldCE06 <> "" Then tmpCp64 = tmpCp64 & "原申請人3:" & tmpOldCE06 & " "
      End If
      If tmpOldCE07 <> textCE07 Then
            SetSRFieldData "SP65", textCE07, 0
            If tmpOldCE07 <> "" Then tmpCp64 = tmpCp64 & "原申請人4:" & tmpOldCE07 & " "
      End If
      If tmpOldCE08 <> textCE08 Then
            SetSRFieldData "SP66", textCE08, 0
            If tmpOldCE08 <> "" Then tmpCp64 = tmpCp64 & "原申請人5:" & tmpOldCE08 & " "
      End If
   End If
   ' 申請日
'   If checkCE03.Value = True Then
   If checkCE03.Value = vbChecked Then
      '911204 nick
      If tmpOldCE02 <> DBDATE(textCE02) Then
            SetSRFieldData "SP10", DBDATE(textCE02), 1
            '911204 nick
'            tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
            If tmpOldCE02 <> "" Then tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
      End If
   End If
   '911204 nick 新增代表人 只判斷是否有變更
   If checkCE16.Value = 1 Then
      If textCE10 <> tmpOldCE10 Then
            SetSRFieldData "SP42", textCE10, 0
'            tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE10 & " "
            If tmpOldCE10 <> "" Then tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE10 & " "
      End If
   End If
   ' 案件名稱
'   If checkCE44.Value = True Then
   If checkCE44.Value = vbChecked Then
        Select Case m_TM01
        Case "S"
            If tmpOldCE41 <> textCE41_1 Then
                SetSRFieldData "SP05", textCE41_1, 0
                If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
            End If
        Case Else
            '911204 nick
            If tmpOldCE41 <> textCE41 Then
                SetSRFieldData "SP05", textCE41, 0
                '911204 nick
    '            tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
                If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
            End If
        End Select
        '911204 nick
        If tmpOldCE42 <> textCE42 Then
            SetSRFieldData "SP06", textCE42, 0
            '911204 nick
'            tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
            If tmpOldCE42 <> "" Then tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
        End If
        '911204 nick
        If tmpOldCE43 <> textCE43 Then
            SetSRFieldData "SP07", textCE43, 0
            '911204 nick
'            tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
            If tmpOldCE43 <> "" Then tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
        End If
   End If
   
   ' 更新服務業務基本檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_SRListCount - 1
      strTmp = Empty
      If m_SRList(nIndex).siType = 0 Then
         strTmp = m_SRList(nIndex).siName & " = '" & m_SRList(nIndex).siData & "'"
      Else
         If m_SRList(nIndex).siData = Empty Then
            strTmp = m_SRList(nIndex).siName & " = " & 0
         Else
            strTmp = m_SRList(nIndex).siName & " = " & m_SRList(nIndex).siData
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
   ' 組成SQL語法
   strSql = strSql & " " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "'"
   ' 執行SQL指令
   If bDifference = True Then
      cnnConnection.Execute strSql
      '911226 nick 更新回原本收文號的備註
      strSql = "update caseprogress set cp64='" & tmpCp64 & "' where cp09='" & m_CE01 & "' "
      cnnConnection.Execute strSql
   End If
   
   ' 清除所佔用的記憶體
   If m_SRListCount > 0 Then
      Erase m_SRList
      m_SRListCount = 0
   End If
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnSaveServicePractice = False
End Function


' 設定欄位新值
Private Sub SetSRFieldData(ByVal strField As String, ByVal strNewData As String, ByVal nType As Integer)
   Dim bFind As Boolean
   Dim nIndex As Integer
   ' 搜尋是否存在該欄位
   bFind = False
   For nIndex = 0 To m_SRListCount - 1
      If m_SRList(nIndex).siName = strField Then
         bFind = True
         m_SRList(nIndex).siData = strNewData
         Exit For
      End If
   Next nIndex
   ' 不存在則新增該欄位
   If bFind = False Then
      ReDim Preserve m_SRList(m_SRListCount + 1)
      m_SRList(m_SRListCount).siName = strField
      m_SRList(m_SRListCount).siData = strNewData
      m_SRList(m_SRListCount).siType = nType
      m_SRListCount = m_SRListCount + 1
   End If
End Sub

'Add By Cheng 2003/04/17
Private Function CheckDataValidate() As Boolean
    CheckDataValidate = False
    'Modify By Cheng 2003/07/01
'    If Me.checkCE09.Value = vbUnchecked And Me.checkCE03.Value = vbUnchecked _
'        And Me.checkCE16.Value = vbUnchecked And Me.checkCE22.Value = vbUnchecked _
'        And Me.checkCE65.Value = vbUnchecked And Me.checkCE38.Value = vbUnchecked _
'        And Me.checkCE58.Value = vbUnchecked And Me.checkCE40.Value = vbUnchecked _
'        And Me.checkCE44.Value = vbUnchecked And Me.checkCE46.Value = vbUnchecked _
'        And Me.checkCE48.Value = vbUnchecked And Me.checkCE50.Value = vbUnchecked _
'        And Me.checkCE62.Value = vbUnchecked Then
    'Modify By Cheng 2003/11/21
'    If Me.checkCE09.Value = vbUnchecked And Me.checkCE03.Value = vbUnchecked _
'        And Me.checkCE16.Value = vbUnchecked And Me.checkCE56.Value = vbUnchecked _
'        And Me.checkCE54.Value = vbUnchecked And Me.checkCE52.Value = vbUnchecked _
'        And Me.checkCE22.Value = vbUnchecked _
'        And Me.checkCE65.Value = vbUnchecked And Me.checkCE38.Value = vbUnchecked _
'        And Me.checkCE58.Value = vbUnchecked And Me.checkCE40.Value = vbUnchecked _
'        And Me.checkCE44.Value = vbUnchecked And Me.checkCE46.Value = vbUnchecked _
'        And Me.checkCE48.Value = vbUnchecked And Me.checkCE50.Value = vbUnchecked _
'        And Me.checkCE62.Value = vbUnchecked Then
    If Me.checkCE09.Value = vbUnchecked And Me.checkCE03.Value = vbUnchecked _
        And Me.checkCE16.Value = vbUnchecked And Me.checkCE56.Value = vbUnchecked _
        And Me.checkCE54.Value = vbUnchecked And Me.checkCE52.Value = vbUnchecked _
        And Me.checkCE22.Value = vbUnchecked And Me.checkCE65.Value = vbUnchecked _
        And Me.checkCE38.Value = vbUnchecked And Me.checkCE58.Value = vbUnchecked _
        And Me.checkCE40.Value = vbUnchecked And Me.checkCE60.Value = vbUnchecked _
        And Me.checkCE44.Value = vbUnchecked And Me.checkCE46.Value = vbUnchecked _
        And Me.checkCE48.Value = vbUnchecked And Me.checkCE50.Value = vbUnchecked _
        And Me.checkCE62.Value = vbUnchecked Then
            MsgBox "請勾選變更項目!!!", vbExclamation + vbOKOnly
            Exit Function
    End If
    If Me.checkCE09.Value = vbChecked Then
        'edit by nickc 2007/01/29
        'If Me.textCE04.Text = "" Then
        If Me.textCE04.Text = "" And textCE05.Text = "" And textCE06.Text = "" And textCE07.Text = "" And textCE08.Text = "" Then
            MsgBox "請輸入申請人代號!!!", vbExclamation + vbOKOnly
            Me.textCE04.SetFocus
            textCE04_GotFocus
            Exit Function
        End If
        'Add By Sindy 2011/8/3
        If m_CP31 <> "Y" Then 'Add By Sindy 2011/8/23 新案時不檢查
            If ChangeCustomerL(textCE04) = m_TM23 And ChangeCustomerL(textCE05) = m_TM78 And ChangeCustomerL(textCE06) = m_TM79 And ChangeCustomerL(textCE07) = m_TM80 And ChangeCustomerL(textCE08) = m_TM81 Then
                MsgBox "新申請人編號與目前相同 !", vbCritical
                Me.textCE04.SetFocus
                textCE04_GotFocus
                Exit Function
            End If
        End If
        '2011/8/3 End
    End If
    If Me.checkCE03.Value = vbChecked Then
        If Me.textCE02.Text = "" Then
            MsgBox "請輸入申請日!!!", vbExclamation + vbOKOnly
            Me.textCE02.SetFocus
            textCE02_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE16.Value = vbChecked Then
        'edit by nickc 2007/01/26
        'If Me.textCE10.Text = "" And Me.textCE11.Text = "" And Me.textCE12.Text = "" And Me.textCE13.Text = "" And Me.textCE14.Text = "" And Me.textCE15.Text = "" Then
        If Me.textCE10.Text = "" And Me.textCE11.Text = "" And Me.textCE12.Text = "" And Me.textCE13.Text = "" And Me.textCE14.Text = "" And Me.textCE15.Text = "" And textCE68.Text = "" And textCE69.Text = "" And textCE70.Text = "" And textCE71.Text = "" And textCE72.Text = "" And textCE73.Text = "" And textCE74.Text = "" And textCE75.Text = "" And textCE76.Text = "" And textCE77.Text = "" And textCE78.Text = "" And textCE79.Text = "" And textCE80.Text = "" And textCE81.Text = "" And textCE82.Text = "" And textCE83.Text = "" And textCE84.Text = "" And textCE85.Text = "" And textCE86.Text = "" And textCE87.Text = "" And textCE88.Text = "" And textCE89.Text = "" And textCE90.Text = "" And textCE91.Text = "" Then
            MsgBox "請輸入代表人名稱!!!", vbExclamation + vbOKOnly
            Me.textCE10.SetFocus
            textCE10_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE22.Value = vbChecked Then
        'edit by nickc 2007/01/26
        'If Me.textCE17.Text = "" Then
        If Me.textCE17.Text = "" And textCE18.Text = "" And textCE19.Text = "" And textCE20.Text = "" And textCE21.Text = "" Then
            MsgBox "請輸入申請人中譯文!!!", vbExclamation + vbOKOnly
            Me.textCE17.SetFocus
            textCE17_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE65.Value = vbChecked Then
        'edit by nick 2007/01/26
        'If Me.textCE63.Text = "" And Me.textCE64.Text = "" Then
        If Me.textCE63.Text = "" And Me.textCE64.Text = "" And textCE92.Text = "" And textCE93.Text = "" And textCE94.Text = "" And textCE95.Text = "" And textCE96.Text = "" And textCE97.Text = "" And textCE98.Text = "" And textCE99.Text = "" Then
            MsgBox "請輸入代表人中譯文!!!", vbExclamation + vbOKOnly
            Me.textCE63.SetFocus
            textCE63_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE38.Value = vbChecked Then
        'ediy by nickc 2007/01/26
        'If Me.textCE23.Text = "" And Me.textCE24.Text = "" And Me.textCE25.Text = "" Then
        If Me.textCE23.Text = "" And Me.textCE24.Text = "" And Me.textCE25.Text = "" And Me.textCE26.Text = "" And Me.textCE27.Text = "" And Me.textCE28.Text = "" And Me.textCE29.Text = "" And Me.textCE30.Text = "" And Me.textCE31.Text = "" And Me.textCE32.Text = "" And Me.textCE33.Text = "" And Me.textCE34.Text = "" And Me.textCE35.Text = "" And Me.textCE36.Text = "" And Me.textCE37.Text = "" Then
            MsgBox "請輸入申請地址!!!", vbExclamation + vbOKOnly
            Me.textCE23.SetFocus
            textCE23_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE58.Value = vbChecked Then
        If Me.textCE57.Text = "" Then
            MsgBox "請輸入正商標號數!!!", vbExclamation + vbOKOnly
            Me.textCE57.SetFocus
            textCE57_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE40.Value = vbChecked Then
        If Me.textCE39.Text = "" Then
            MsgBox "請輸入正商標種類!!!", vbExclamation + vbOKOnly
            Me.textCE39.SetFocus
            textCE39_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE44.Value = vbChecked Then
        Select Case m_TM01
        Case "FCT", "S"
            If Me.textCE41_1.Text = "" Then
                MsgBox "請輸入案件名稱!!!", vbExclamation + vbOKOnly
                Me.textCE41_1.SetFocus
                textCE41_1_GotFocus
                Exit Function
            End If
        Case Else
            If Me.textCE41.Text = "" And Me.textCE42.Text = "" And Me.textCE43.Text = "" Then
                MsgBox "請輸入案件名稱!!!", vbExclamation + vbOKOnly
                Me.textCE41.SetFocus
                textCE41_GotFocus
                Exit Function
            End If
        End Select
    End If
    'Add By Sindy 2018/2/1 減縮商品發文畫面frm030202_07按變更事項呼叫的frm030202_05
    '請檢查收文號之案件性質為減縮商品時, 第五頁之checkCE46一定要勾註,
    '但可不必輸入textCE45縮減商品.
    If m_CP10 = "313" And Me.checkCE46.Value = vbUnchecked Then
      MsgBox "減縮商品案，其變更事項的「減縮商品」為必須勾選項目!!!", vbExclamation + vbOKOnly
      tabCtrl.Tab = 4
      Exit Function
    End If
    '2018/2/1 END
'    If Me.checkCE46.Value = vbChecked Then
'        If Me.textCE45.Text = "" Then
'            MsgBox "請輸入縮減商品!!!", vbExclamation + vbOKOnly
'            Me.textCE45.SetFocus
'            textCE45_GotFocus
'            Exit Function
'        End If
'    End If
    If Me.checkCE48.Value = vbChecked Then
        If Me.textCE47.Text = "" Then
            MsgBox "請輸入商品類別!!!", vbExclamation + vbOKOnly
            Me.textCE47.SetFocus
            textCE47_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE50.Value = vbChecked Then
        If Me.textCE49.Text = "" Then
            MsgBox "請輸入商品群組!!!", vbExclamation + vbOKOnly
            Me.textCE49.SetFocus
            textCE49_GotFocus
            Exit Function
        End If
    End If
    If Me.checkCE62.Value = vbChecked Then
        If Me.textCE61.Text = "" Then
            MsgBox "請輸入其他!!!", vbExclamation + vbOKOnly
            Me.textCE61.SetFocus
            textCE61_GotFocus
            Exit Function
        End If
    End If
    
    'Added by Lydia 2021/09/02 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         Exit Function
    End If

    CheckDataValidate = True
End Function
'add by nickc 2007/01/25
Private Sub textCE68_GotFocus()
InverseTextBox textCE68
End Sub
Private Sub textCE69_GotFocus()
InverseTextBox textCE69
End Sub
Private Sub textCE70_GotFocus()
InverseTextBox textCE70
End Sub
Private Sub textCE71_GotFocus()
InverseTextBox textCE71
End Sub
Private Sub textCE72_GotFocus()
InverseTextBox textCE72
End Sub
Private Sub textCE73_GotFocus()
InverseTextBox textCE73
End Sub
Private Sub textCE74_GotFocus()
InverseTextBox textCE74
End Sub
Private Sub textCE75_GotFocus()
InverseTextBox textCE75
End Sub
Private Sub textCE76_GotFocus()
InverseTextBox textCE76
End Sub
Private Sub textCE77_GotFocus()
InverseTextBox textCE77
End Sub
Private Sub textCE78_GotFocus()
InverseTextBox textCE78
End Sub
Private Sub textCE79_GotFocus()
InverseTextBox textCE79
End Sub
Private Sub textCE80_GotFocus()
InverseTextBox textCE80
End Sub
Private Sub textCE81_GotFocus()
InverseTextBox textCE81
End Sub
Private Sub textCE82_GotFocus()
InverseTextBox textCE82
End Sub
Private Sub textCE83_GotFocus()
InverseTextBox textCE83
End Sub
Private Sub textCE84_GotFocus()
InverseTextBox textCE84
End Sub
Private Sub textCE85_GotFocus()
InverseTextBox textCE85
End Sub
Private Sub textCE86_GotFocus()
InverseTextBox textCE86
End Sub
Private Sub textCE87_GotFocus()
InverseTextBox textCE87
End Sub
Private Sub textCE88_GotFocus()
InverseTextBox textCE88
End Sub
Private Sub textCE89_GotFocus()
InverseTextBox textCE89
End Sub
Private Sub textCE90_GotFocus()
InverseTextBox textCE90
End Sub
Private Sub textCE91_GotFocus()
InverseTextBox textCE91
End Sub

'add by nickc 2007/01/25
Private Sub textCE92_GotFocus()
InverseTextBox textCE92
End Sub
Private Sub textCE93_GotFocus()
InverseTextBox textCE93
End Sub
Private Sub textCE94_GotFocus()
InverseTextBox textCE94
End Sub
Private Sub textCE95_GotFocus()
InverseTextBox textCE95
End Sub
Private Sub textCE96_GotFocus()
InverseTextBox textCE96
End Sub
Private Sub textCE97_GotFocus()
InverseTextBox textCE97
End Sub
Private Sub textCE98_GotFocus()
InverseTextBox textCE98
End Sub
Private Sub textCE99_GotFocus()
InverseTextBox textCE99
End Sub
