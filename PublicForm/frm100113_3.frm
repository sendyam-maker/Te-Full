VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100113_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "變更事項"
   ClientHeight    =   6495
   ClientLeft      =   30
   ClientTop       =   975
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9150
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   5940
      Style           =   1  '圖片外觀
      TabIndex        =   188
      Top             =   60
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8205
      Style           =   1  '圖片外觀
      TabIndex        =   187
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   6990
      Style           =   1  '圖片外觀
      TabIndex        =   186
      Top             =   60
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5670
      Left            =   120
      TabIndex        =   97
      Top             =   780
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   10001
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "第1頁"
      TabPicture(0)   =   "frm100113_3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4(11)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4(10)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label4(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label3(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3(16)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label3(14)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label3(13)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label3(12)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label5(9)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label5(10)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label5(12)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label11(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label3(4)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label5(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label3(17)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label5(6)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCE08_2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCE07_2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCE06_2"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCE05_2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCE04_2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCE61"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCE21"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textCE20"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCE19"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textCE18"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textCE17"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textCE02"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textCE03"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textCE04"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCE05"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textCE06"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCE07"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCE08"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textCE09"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textCE22"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textCE56"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textCE54"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textCE52"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "textCE60"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "textCE59"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "textCE55"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "textCE53"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "textCE51"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "textCE40"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "textCE39"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "textCE62"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "textCE39_2"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).ControlCount=   58
      TabCaption(1)   =   "第2頁"
      TabPicture(1)   =   "frm100113_3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textCE66"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "textCE67"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "textCE38"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "textCE23"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "textCE24"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "textCE25"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "textCE26"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "textCE27"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "textCE28"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "textCE29"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "textCE30"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "textCE31"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "textCE32"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "textCE33"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "textCE34"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "textCE35"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "textCE36"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "textCE37"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label4(34)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label3(8)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label3(3)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label4(30)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label4(29)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label4(28)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label4(27)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label4(26)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label4(25)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label4(24)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label4(23)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label4(22)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Label4(21)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Label4(20)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Label4(19)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label4(18)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Label4(17)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Label4(16)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).ControlCount=   36
      TabCaption(2)   =   "第3頁"
      TabPicture(2)   =   "frm100113_3.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "textCE65"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "textCE58"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "textCE50"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "textCE49"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "textCE48"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "textCE44"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "textCE57"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "textCE47"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "textCE46"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "textCE41"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "textCE42"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "textCE43"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "textCE45"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "textCE63"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "textCE64"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "textCE92"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "textCE93"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "textCE94"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "textCE95"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "textCE96"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "textCE97"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "textCE98"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "textCE99"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Label4(42)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Label4(41)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Label4(40)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Label4(39)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Label4(38)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Label4(37)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Label4(36)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Label4(35)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Label3(7)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Label4(33)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Label4(32)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Label3(15)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Label3(11)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Label3(10)"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Label3(9)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Label3(5)"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Label5(11)"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Label5(5)"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Label5(4)"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Label5(3)"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Label4(31)"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "Label70(5)"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "Label8"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).ControlCount=   46
      TabCaption(3)   =   "第4頁"
      TabPicture(3)   =   "frm100113_3.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "textCE16"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "textCE77"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "textCE78"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "textCE79"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "textCE80"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "textCE81"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "textCE82"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "textCE83"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "textCE84"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "textCE85"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "textCE86"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "textCE87"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "textCE88"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "textCE89"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "textCE90"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "textCE91"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "textCE10"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "textCE11"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "textCE12"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "textCE13"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "textCE14"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "textCE15"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "textCE68"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "textCE69"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "textCE70"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "textCE71"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "textCE72"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "textCE73"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "textCE74"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "textCE75"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "textCE76"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "Label4(49)"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "Label4(48)"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "Label4(47)"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "Label4(46)"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).Control(35)=   "Label4(45)"
      Tab(3).Control(35).Enabled=   0   'False
      Tab(3).Control(36)=   "Label4(44)"
      Tab(3).Control(36).Enabled=   0   'False
      Tab(3).Control(37)=   "Label4(15)"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "Label4(14)"
      Tab(3).Control(38).Enabled=   0   'False
      Tab(3).Control(39)=   "Label4(7)"
      Tab(3).Control(39).Enabled=   0   'False
      Tab(3).Control(40)=   "Label4(43)"
      Tab(3).Control(40).Enabled=   0   'False
      Tab(3).Control(41)=   "Label3(1)"
      Tab(3).Control(41).Enabled=   0   'False
      Tab(3).Control(42)=   "Label4(1)"
      Tab(3).Control(42).Enabled=   0   'False
      Tab(3).Control(43)=   "Label4(12)"
      Tab(3).Control(43).Enabled=   0   'False
      Tab(3).Control(44)=   "Label4(13)"
      Tab(3).Control(44).Enabled=   0   'False
      Tab(3).ControlCount=   45
      Begin VB.TextBox textCE39_2 
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   194
         TabStop         =   0   'False
         Top             =   4830
         Width           =   4515
      End
      Begin VB.TextBox textCE16 
         Height          =   300
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   68
         Top             =   408
         Width           =   255
      End
      Begin VB.TextBox textCE66 
         Height          =   300
         Left            =   -71940
         MaxLength       =   25
         TabIndex        =   44
         Top             =   5265
         Width           =   2172
      End
      Begin VB.TextBox textCE67 
         Height          =   300
         Left            =   -73680
         MaxLength       =   1
         TabIndex        =   43
         Top             =   5265
         Width           =   255
      End
      Begin VB.TextBox textCE62 
         Height          =   300
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   25
         Top             =   5160
         Width           =   255
      End
      Begin VB.TextBox textCE39 
         Height          =   300
         Left            =   3630
         MaxLength       =   1
         TabIndex        =   24
         Top             =   4830
         Width           =   255
      End
      Begin VB.TextBox textCE40 
         Height          =   300
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   23
         Top             =   4830
         Width           =   255
      End
      Begin VB.TextBox textCE65 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   57
         Top             =   2617
         Width           =   255
      End
      Begin VB.TextBox textCE51 
         Height          =   300
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   10
         Top             =   2136
         Width           =   255
      End
      Begin VB.TextBox textCE53 
         Height          =   300
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   12
         Top             =   2436
         Width           =   255
      End
      Begin VB.TextBox textCE55 
         Height          =   300
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   14
         Top             =   2736
         Width           =   255
      End
      Begin VB.TextBox textCE59 
         Height          =   300
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   16
         Top             =   3036
         Width           =   255
      End
      Begin VB.TextBox textCE60 
         Height          =   300
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   15
         Top             =   3036
         Width           =   255
      End
      Begin VB.TextBox textCE52 
         Height          =   300
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   9
         Top             =   2136
         Width           =   255
      End
      Begin VB.TextBox textCE54 
         Height          =   300
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   11
         Top             =   2436
         Width           =   255
      End
      Begin VB.TextBox textCE56 
         Height          =   300
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   13
         Top             =   2736
         Width           =   255
      End
      Begin VB.TextBox textCE58 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   55
         Top             =   2297
         Width           =   255
      End
      Begin VB.TextBox textCE50 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   53
         Top             =   1971
         Width           =   255
      End
      Begin VB.TextBox textCE49 
         Height          =   300
         Left            =   -71490
         MaxLength       =   699
         TabIndex        =   54
         Top             =   1956
         Width           =   4695
      End
      Begin VB.TextBox textCE48 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   51
         Top             =   1645
         Width           =   255
      End
      Begin VB.TextBox textCE44 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   45
         Top             =   345
         Width           =   255
      End
      Begin VB.TextBox textCE57 
         Height          =   300
         Left            =   -71490
         MaxLength       =   20
         TabIndex        =   56
         Top             =   2282
         Width           =   1935
      End
      Begin VB.TextBox textCE47 
         Height          =   300
         Left            =   -71490
         MaxLength       =   395
         TabIndex        =   52
         Top             =   1630
         Width           =   4695
      End
      Begin VB.TextBox textCE46 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   49
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox textCE38 
         Height          =   300
         Left            =   -73680
         MaxLength       =   1
         TabIndex        =   27
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox textCE22 
         Height          =   300
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   17
         Top             =   3336
         Width           =   255
      End
      Begin VB.TextBox textCE09 
         Height          =   300
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   3
         Top             =   636
         Width           =   255
      End
      Begin VB.TextBox textCE08 
         Height          =   300
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   8
         Top             =   1836
         Width           =   975
      End
      Begin VB.TextBox textCE07 
         Height          =   300
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   7
         Top             =   1536
         Width           =   975
      End
      Begin VB.TextBox textCE06 
         Height          =   300
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   6
         Top             =   1236
         Width           =   975
      End
      Begin VB.TextBox textCE05 
         Height          =   300
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   5
         Top             =   936
         Width           =   975
      End
      Begin VB.TextBox textCE04 
         Height          =   300
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   4
         Top             =   636
         Width           =   975
      End
      Begin VB.TextBox textCE03 
         Height          =   300
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   1
         Top             =   336
         Width           =   255
      End
      Begin VB.TextBox textCE02 
         Height          =   300
         Left            =   3240
         MaxLength       =   7
         TabIndex        =   2
         Top             =   336
         Width           =   975
      End
      Begin MSForms.TextBox textCE77 
         Height          =   300
         Left            =   -73920
         TabIndex        =   84
         Top             =   2515
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE78 
         Height          =   300
         Left            =   -71323
         TabIndex        =   85
         Top             =   2515
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE79 
         Height          =   300
         Left            =   -68730
         TabIndex        =   86
         Top             =   2515
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE80 
         Height          =   300
         Left            =   -73920
         TabIndex        =   87
         Top             =   2868
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE81 
         Height          =   300
         Left            =   -71323
         TabIndex        =   88
         Top             =   2868
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE82 
         Height          =   300
         Left            =   -68730
         TabIndex        =   89
         Top             =   2868
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE83 
         Height          =   300
         Left            =   -73920
         TabIndex        =   90
         Top             =   3221
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE84 
         Height          =   300
         Left            =   -71323
         TabIndex        =   91
         Top             =   3221
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE85 
         Height          =   300
         Left            =   -68730
         TabIndex        =   92
         Top             =   3221
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE86 
         Height          =   300
         Left            =   -73920
         TabIndex        =   93
         Top             =   3574
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE87 
         Height          =   300
         Left            =   -71323
         TabIndex        =   94
         Top             =   3574
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE88 
         Height          =   300
         Left            =   -68730
         TabIndex        =   95
         Top             =   3574
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE89 
         Height          =   300
         Left            =   -73920
         TabIndex        =   96
         Top             =   3930
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE90 
         Height          =   300
         Left            =   -71323
         TabIndex        =   195
         Top             =   3930
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE91 
         Height          =   300
         Left            =   -68730
         TabIndex        =   98
         Top             =   3930
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE10 
         Height          =   300
         Left            =   -73920
         TabIndex        =   69
         Top             =   750
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4410;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE11 
         Height          =   300
         Left            =   -71323
         TabIndex        =   70
         Top             =   750
         Width           =   2500
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "4410;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE12 
         Height          =   300
         Left            =   -68730
         TabIndex        =   71
         Top             =   750
         Width           =   2500
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "4410;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE13 
         Height          =   300
         Left            =   -73920
         TabIndex        =   72
         Top             =   1103
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE14 
         Height          =   300
         Left            =   -71323
         TabIndex        =   73
         Top             =   1103
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE15 
         Height          =   300
         Left            =   -68730
         TabIndex        =   74
         Top             =   1103
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE68 
         Height          =   300
         Left            =   -73920
         TabIndex        =   75
         Top             =   1456
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE69 
         Height          =   300
         Left            =   -71323
         TabIndex        =   76
         Top             =   1456
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE70 
         Height          =   300
         Left            =   -68730
         TabIndex        =   77
         Top             =   1456
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE71 
         Height          =   300
         Left            =   -73920
         TabIndex        =   78
         Top             =   1809
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE72 
         Height          =   300
         Left            =   -71323
         TabIndex        =   79
         Top             =   1809
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE73 
         Height          =   300
         Left            =   -68730
         TabIndex        =   80
         Top             =   1809
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE74 
         Height          =   300
         Left            =   -73920
         TabIndex        =   81
         Top             =   2162
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE75 
         Height          =   300
         Left            =   -71323
         TabIndex        =   82
         Top             =   2162
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE76 
         Height          =   300
         Left            =   -68730
         TabIndex        =   83
         Top             =   2162
         Width           =   2505
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "4419;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41 
         Height          =   300
         Left            =   -71490
         TabIndex        =   46
         Top             =   330
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE42 
         Height          =   300
         Left            =   -71490
         TabIndex        =   47
         Top             =   655
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   180
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE43 
         Height          =   300
         Left            =   -71490
         TabIndex        =   48
         Top             =   980
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE45 
         Height          =   300
         Left            =   -71490
         TabIndex        =   50
         Top             =   1305
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE63 
         Height          =   300
         Left            =   -71490
         TabIndex        =   58
         Top             =   2610
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE64 
         Height          =   300
         Left            =   -71490
         TabIndex        =   59
         Top             =   2910
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE92 
         Height          =   300
         Left            =   -71490
         TabIndex        =   60
         Top             =   3210
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE93 
         Height          =   300
         Left            =   -71490
         TabIndex        =   61
         Top             =   3510
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE94 
         Height          =   300
         Left            =   -71490
         TabIndex        =   62
         Top             =   3810
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE95 
         Height          =   300
         Left            =   -71490
         TabIndex        =   63
         Top             =   4110
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE96 
         Height          =   300
         Left            =   -71490
         TabIndex        =   64
         Top             =   4410
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE97 
         Height          =   300
         Left            =   -71490
         TabIndex        =   65
         Top             =   4710
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE98 
         Height          =   300
         Left            =   -71490
         TabIndex        =   66
         Top             =   5010
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE99 
         Height          =   300
         Left            =   -71490
         TabIndex        =   67
         Top             =   5310
         Width           =   5295
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9340;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE23 
         Height          =   300
         Left            =   -71940
         TabIndex        =   28
         Top             =   480
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE24 
         Height          =   300
         Left            =   -71940
         TabIndex        =   29
         Top             =   780
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE25 
         Height          =   300
         Left            =   -71940
         TabIndex        =   30
         Top             =   1080
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE26 
         Height          =   300
         Left            =   -71940
         TabIndex        =   31
         Top             =   1440
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE27 
         Height          =   300
         Left            =   -71940
         TabIndex        =   32
         Top             =   1740
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE28 
         Height          =   300
         Left            =   -71940
         TabIndex        =   33
         Top             =   2040
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE29 
         Height          =   300
         Left            =   -71940
         TabIndex        =   34
         Top             =   2400
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE30 
         Height          =   300
         Left            =   -71940
         TabIndex        =   35
         Top             =   2700
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE31 
         Height          =   300
         Left            =   -71940
         TabIndex        =   36
         Top             =   3000
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE32 
         Height          =   300
         Left            =   -71940
         TabIndex        =   37
         Top             =   3360
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE33 
         Height          =   300
         Left            =   -71940
         TabIndex        =   38
         Top             =   3660
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE34 
         Height          =   300
         Left            =   -71940
         TabIndex        =   39
         Top             =   3960
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE35 
         Height          =   300
         Left            =   -71940
         TabIndex        =   40
         Top             =   4320
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE36 
         Height          =   300
         Left            =   -71940
         TabIndex        =   41
         Top             =   4620
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE37 
         Height          =   300
         Left            =   -71940
         TabIndex        =   42
         Top             =   4920
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "10054;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE17 
         Height          =   300
         Left            =   3240
         TabIndex        =   18
         Top             =   3336
         Width           =   5500
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9701;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE18 
         Height          =   300
         Left            =   3240
         TabIndex        =   19
         Top             =   3636
         Width           =   5500
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9701;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE19 
         Height          =   300
         Left            =   3240
         TabIndex        =   20
         Top             =   3936
         Width           =   5500
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9701;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE20 
         Height          =   300
         Left            =   3240
         TabIndex        =   21
         Top             =   4236
         Width           =   5500
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9701;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE21 
         Height          =   300
         Left            =   3240
         TabIndex        =   22
         Top             =   4536
         Width           =   5500
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9701;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE61 
         Height          =   465
         Left            =   3240
         TabIndex        =   26
         Top             =   5160
         Width           =   5505
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         Size            =   "9710;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE04_2 
         Height          =   285
         Left            =   4260
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   636
         Width           =   4500
         VariousPropertyBits=   671105055
         Size            =   "7937;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE05_2 
         Height          =   285
         Left            =   4260
         TabIndex        =   192
         TabStop         =   0   'False
         Top             =   936
         Width           =   4500
         VariousPropertyBits=   671105055
         Size            =   "7937;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE06_2 
         Height          =   285
         Left            =   4260
         TabIndex        =   191
         TabStop         =   0   'False
         Top             =   1236
         Width           =   4500
         VariousPropertyBits=   671105055
         Size            =   "7937;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE07_2 
         Height          =   285
         Left            =   4260
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   1536
         Width           =   4500
         VariousPropertyBits=   671105055
         Size            =   "7937;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE08_2 
         Height          =   285
         Left            =   4260
         TabIndex        =   189
         TabStop         =   0   'False
         Top             =   1836
         Width           =   4500
         VariousPropertyBits=   671105055
         Size            =   "7937;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人10:"
         Height          =   180
         Index           =   49
         Left            =   -74880
         TabIndex        =   185
         Top             =   3990
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人9:"
         Height          =   180
         Index           =   48
         Left            =   -74880
         TabIndex        =   184
         Top             =   3635
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人8:"
         Height          =   180
         Index           =   47
         Left            =   -74880
         TabIndex        =   183
         Top             =   3280
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人7:"
         Height          =   180
         Index           =   46
         Left            =   -74880
         TabIndex        =   182
         Top             =   2925
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人6:"
         Height          =   180
         Index           =   45
         Left            =   -74880
         TabIndex        =   181
         Top             =   2570
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人5:"
         Height          =   180
         Index           =   44
         Left            =   -74880
         TabIndex        =   180
         Top             =   2215
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人4:"
         Height          =   180
         Index           =   15
         Left            =   -74880
         TabIndex        =   179
         Top             =   1860
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人3:"
         Height          =   180
         Index           =   14
         Left            =   -74880
         TabIndex        =   178
         Top             =   1505
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人2:"
         Height          =   180
         Index           =   7
         Left            =   -74880
         TabIndex        =   177
         Top             =   1150
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人1:"
         Height          =   180
         Index           =   43
         Left            =   -74880
         TabIndex        =   176
         Top             =   795
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文10:"
         Height          =   180
         Index           =   42
         Left            =   -72960
         TabIndex        =   175
         Top             =   5362
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文9:"
         Height          =   180
         Index           =   41
         Left            =   -72960
         TabIndex        =   174
         Top             =   5057
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文8:"
         Height          =   180
         Index           =   40
         Left            =   -72960
         TabIndex        =   173
         Top             =   4753
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文7:"
         Height          =   180
         Index           =   39
         Left            =   -72960
         TabIndex        =   172
         Top             =   4449
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文6:"
         Height          =   180
         Index           =   38
         Left            =   -72960
         TabIndex        =   171
         Top             =   4145
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文5:"
         Height          =   180
         Index           =   37
         Left            =   -72960
         TabIndex        =   170
         Top             =   3841
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文4:"
         Height          =   180
         Index           =   36
         Left            =   -72960
         TabIndex        =   169
         Top             =   3537
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文3:"
         Height          =   180
         Index           =   35
         Left            =   -72960
         TabIndex        =   168
         Top             =   3233
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   167
         Top             =   405
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人(中):"
         Height          =   180
         Index           =   1
         Left            =   -73110
         TabIndex        =   166
         Top             =   420
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人(英):"
         Height          =   180
         Index           =   12
         Left            =   -70515
         TabIndex        =   165
         Top             =   405
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人(日):"
         Height          =   180
         Index           =   13
         Left            =   -67922
         TabIndex        =   164
         Top             =   405
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "網域密碼 :"
         Height          =   180
         Index           =   34
         Left            =   -73200
         TabIndex        =   163
         Top             =   5325
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   8
         Left            =   -74760
         TabIndex        =   162
         Top             =   5325
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "其他:"
         Height          =   180
         Index           =   6
         Left            =   1860
         TabIndex        =   161
         Top             =   5160
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   17
         Left            =   180
         TabIndex        =   160
         Top             =   5160
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "專利商標種類代號 :"
         Height          =   180
         Index           =   2
         Left            =   1890
         TabIndex        =   159
         Top             =   4830
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   4
         Left            =   150
         TabIndex        =   158
         Top             =   4830
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   7
         Left            =   -74790
         TabIndex        =   153
         Top             =   2662
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文2:"
         Height          =   180
         Index           =   33
         Left            =   -72960
         TabIndex        =   152
         Top             =   2970
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "代表人中譯文1:"
         Height          =   180
         Index           =   32
         Left            =   -72960
         TabIndex        =   151
         Top             =   2662
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "代理人:"
         Height          =   180
         Index           =   1
         Left            =   1920
         TabIndex        =   150
         Top             =   2736
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "圖樣:"
         Height          =   180
         Index           =   12
         Left            =   1920
         TabIndex        =   149
         Top             =   3036
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "代表人印鑑:"
         Height          =   180
         Index           =   10
         Left            =   1920
         TabIndex        =   148
         Top             =   2436
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "申請人印鑑:"
         Height          =   180
         Index           =   9
         Left            =   1920
         TabIndex        =   147
         Top             =   2136
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   12
         Left            =   240
         TabIndex        =   146
         Top             =   2136
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   13
         Left            =   240
         TabIndex        =   145
         Top             =   2436
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   14
         Left            =   240
         TabIndex        =   144
         Top             =   2736
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   16
         Left            =   240
         TabIndex        =   143
         Top             =   3036
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   15
         Left            =   -74790
         TabIndex        =   141
         Top             =   2342
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   11
         Left            =   -74790
         TabIndex        =   140
         Top             =   2016
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   10
         Left            =   -74790
         TabIndex        =   139
         Top             =   1690
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   9
         Left            =   -74790
         TabIndex        =   138
         Top             =   1365
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   5
         Left            =   -74790
         TabIndex        =   137
         Top             =   390
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "正商標號數:"
         Height          =   180
         Index           =   11
         Left            =   -72960
         TabIndex        =   136
         Top             =   2342
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "案件日文名稱:"
         Height          =   180
         Index           =   5
         Left            =   -72960
         TabIndex        =   135
         Top             =   1040
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "案件英文名稱:"
         Height          =   180
         Index           =   4
         Left            =   -72960
         TabIndex        =   134
         Top             =   715
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "案件中文名稱:"
         Height          =   180
         Index           =   3
         Left            =   -72960
         TabIndex        =   133
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "減縮商品:"
         Height          =   180
         Index           =   31
         Left            =   -72960
         TabIndex        =   132
         Top             =   1365
         Width           =   765
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "商品類別:"
         Height          =   180
         Index           =   5
         Left            =   -72960
         TabIndex        =   131
         Top             =   1690
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "商品組群:"
         Height          =   180
         Left            =   -72960
         TabIndex        =   130
         Top             =   2016
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   3
         Left            =   -74760
         TabIndex        =   129
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(中)5:"
         Height          =   180
         Index           =   30
         Left            =   -73230
         TabIndex        =   128
         Top             =   4380
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(英)5:"
         Height          =   180
         Index           =   29
         Left            =   -73230
         TabIndex        =   127
         Top             =   4710
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(日)5:"
         Height          =   180
         Index           =   28
         Left            =   -73230
         TabIndex        =   126
         Top             =   5010
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(中)3:"
         Height          =   180
         Index           =   27
         Left            =   -73230
         TabIndex        =   125
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(英)3:"
         Height          =   180
         Index           =   26
         Left            =   -73230
         TabIndex        =   124
         Top             =   2760
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(日)3:"
         Height          =   180
         Index           =   25
         Left            =   -73230
         TabIndex        =   123
         Top             =   3060
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(中)4:"
         Height          =   180
         Index           =   24
         Left            =   -73230
         TabIndex        =   122
         Top             =   3420
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(英)4:"
         Height          =   180
         Index           =   23
         Left            =   -73230
         TabIndex        =   121
         Top             =   3690
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(日)4:"
         Height          =   180
         Index           =   22
         Left            =   -73230
         TabIndex        =   120
         Top             =   4050
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(中)1:"
         Height          =   180
         Index           =   21
         Left            =   -73230
         TabIndex        =   119
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(英)1:"
         Height          =   180
         Index           =   20
         Left            =   -73230
         TabIndex        =   118
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(日)1:"
         Height          =   180
         Index           =   19
         Left            =   -73230
         TabIndex        =   117
         Top             =   1170
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(中)2:"
         Height          =   180
         Index           =   18
         Left            =   -73230
         TabIndex        =   116
         Top             =   1470
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(英)2:"
         Height          =   180
         Index           =   17
         Left            =   -73230
         TabIndex        =   115
         Top             =   1770
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請地址(日)2:"
         Height          =   180
         Index           =   16
         Left            =   -73230
         TabIndex        =   114
         Top             =   2130
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   113
         Top             =   3336
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人1中譯文:"
         Height          =   180
         Index           =   2
         Left            =   1920
         TabIndex        =   112
         Top             =   3336
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人2中譯文:"
         Height          =   180
         Index           =   8
         Left            =   1920
         TabIndex        =   111
         Top             =   3636
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人3中譯文:"
         Height          =   180
         Index           =   9
         Left            =   1920
         TabIndex        =   110
         Top             =   3936
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人4中譯文:"
         Height          =   180
         Index           =   10
         Left            =   1920
         TabIndex        =   109
         Top             =   4236
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人5中譯文:"
         Height          =   180
         Index           =   11
         Left            =   1920
         TabIndex        =   108
         Top             =   4536
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   107
         Top             =   636
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人1:"
         Height          =   180
         Index           =   0
         Left            =   1920
         TabIndex        =   106
         Top             =   636
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人2:"
         Height          =   180
         Index           =   3
         Left            =   1920
         TabIndex        =   105
         Top             =   936
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人3:"
         Height          =   180
         Index           =   4
         Left            =   1920
         TabIndex        =   104
         Top             =   1236
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人4:"
         Height          =   180
         Index           =   5
         Left            =   1920
         TabIndex        =   103
         Top             =   1536
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人5:"
         Height          =   180
         Index           =   6
         Left            =   1920
         TabIndex        =   102
         Top             =   1836
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "准(1)/駁(2) :"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   101
         Top             =   336
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請日:"
         Height          =   180
         Left            =   1920
         TabIndex        =   100
         Top             =   336
         Width           =   585
      End
   End
   Begin VB.TextBox textCP04 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5460
      MaxLength       =   2
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   450
      Width           =   372
   End
   Begin VB.TextBox textCP03 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5220
      MaxLength       =   1
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   450
      Width           =   252
   End
   Begin VB.TextBox textCP02 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4500
      MaxLength       =   6
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   450
      Width           =   732
   End
   Begin VB.TextBox textCP01 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4020
      MaxLength       =   3
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   450
      Width           =   492
   End
   Begin VB.TextBox textCE01 
      Height          =   300
      Left            =   840
      MaxLength       =   9
      TabIndex        =   0
      Top             =   450
      Width           =   1212
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   3210
      TabIndex        =   142
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   210
      TabIndex        =   99
      Top             =   480
      Width           =   585
   End
End
Attribute VB_Name = "frm100113_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/05 改成Form2.0 ; textCE04_2、textCE05_2、textCE06_2、textCE07_2、textCE08_2、textCE17~21、textCE63~64、textCE92~99、textCE10~15、textCE68~91、textCE23~37、textCE41~43、textCE45、textCE61
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim m_arrCE01
Dim m_intCurRecord As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()
    Dim ii As Integer

    Select Case cmdState
    Case 0 '回前畫面
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Case 1 '結束
          fnCloseAllFrm100
    Case 2 '下一筆
        If Me.Tag <> "" Then
'            For ii = LBound(m_arrCE01) + 2 To UBound(m_arrCE01)
                ClearField
                m_intCurRecord = m_intCurRecord + 1
                UpdateCtrlData "" & m_arrCE01(m_intCurRecord)
'            Next ii
            If m_intCurRecord >= UBound(m_arrCE01) Then Me.cmdOK(2).Enabled = False
        Else
            Me.cmdOK(2).Enabled = False
        End If
    End Select

End Sub

Private Sub cmdOK_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
'92.04.16 nick 以下無效
    Dim ii As Integer

    Select Case Index
    Case 0 '回前畫面
        Unload Me
        frm100113_2.Show
    Case 1 '結束
        bolToEndByNick = True
        Unload Me
        Unload frm100113_2
        Unload frm100113_1
    Case 2 '下一筆
        If Me.Tag <> "" Then
'            For ii = LBound(m_arrCE01) + 2 To UBound(m_arrCE01)
                ClearField
                m_intCurRecord = m_intCurRecord + 1
                UpdateCtrlData "" & m_arrCE01(m_intCurRecord)
'            Next ii
            If m_intCurRecord >= UBound(m_arrCE01) Then Me.cmdOK(2).Enabled = False
        Else
            Me.cmdOK(2).Enabled = False
        End If
    End Select
End Sub

Private Sub Form_Activate()
    If m_intCurRecord = 0 Then
        If Me.Tag <> "" Then
            m_arrCE01 = Split(Me.Tag, ",")
            If Val(m_arrCE01(0)) <= 1 Then Me.cmdOK(2).Enabled = False
            UpdateCtrlData "" & m_arrCE01(1)
            m_intCurRecord = 1
        End If
    End If
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    
    textCE04_2.BackColor = &H8000000F
    textCE05_2.BackColor = &H8000000F
    textCE06_2.BackColor = &H8000000F
    textCE07_2.BackColor = &H8000000F
    textCE08_2.BackColor = &H8000000F
    textCE39_2.BackColor = &H8000000F
    
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
    
    EnableTextBox textCP01, False
    EnableTextBox textCP02, False
    EnableTextBox textCP03, False
    EnableTextBox textCP04, False
    
    MoveFormToCenter Me
    ClearField
    SetCtrlReadOnly True
    
    m_intCurRecord = 0
'92.04.16 nick
cmdState = -1
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   textCP01 = Empty
   textCP02 = Empty
   textCP03 = Empty
   textCP04 = Empty
   textCE01 = Empty
   textCE02 = Empty
   textCE03 = Empty
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
   textCE09 = Empty
   textCE10 = Empty
   textCE11 = Empty
   textCE12 = Empty
   textCE13 = Empty
   textCE14 = Empty
   textCE15 = Empty
   textCE16 = Empty
   textCE17 = Empty
   textCE18 = Empty
   textCE19 = Empty
   textCE20 = Empty
   textCE21 = Empty
   textCE22 = Empty
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
   textCE38 = Empty
   textCE39 = Empty
   textCE39_2 = Empty
   textCE40 = Empty
   textCE41 = Empty
   textCE42 = Empty
   textCE43 = Empty
   textCE44 = Empty
   textCE45 = Empty
   textCE46 = Empty
   textCE47 = Empty
   textCE48 = Empty
   textCE49 = Empty
   textCE50 = Empty
   textCE51 = Empty
   textCE52 = Empty
   textCE53 = Empty
   textCE54 = Empty
   textCE55 = Empty
   textCE56 = Empty
   textCE57 = Empty
   textCE58 = Empty
   textCE59 = Empty
   textCE60 = Empty
   textCE61 = Empty
   textCE62 = Empty
   textCE63 = Empty
   textCE64 = Empty
   textCE65 = Empty
   textCE66 = Empty
   textCE67 = Empty
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
   textCE92 = Empty
   textCE93 = Empty
   textCE94 = Empty
   textCE95 = Empty
   textCE96 = Empty
   textCE97 = Empty
   textCE98 = Empty
   textCE99 = Empty
      
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textCE01.Locked = bEnable
   textCE02.Locked = bEnable
   textCE03.Locked = bEnable
   textCE04.Locked = bEnable
   textCE05.Locked = bEnable
   textCE06.Locked = bEnable
   textCE07.Locked = bEnable
   textCE08.Locked = bEnable
   textCE09.Locked = bEnable
   textCE10.Locked = bEnable
   textCE11.Locked = bEnable
   textCE12.Locked = bEnable
   textCE13.Locked = bEnable
   textCE14.Locked = bEnable
   textCE15.Locked = bEnable
   textCE16.Locked = bEnable
   textCE17.Locked = bEnable
   textCE18.Locked = bEnable
   textCE19.Locked = bEnable
   textCE20.Locked = bEnable
   textCE21.Locked = bEnable
   textCE22.Locked = bEnable
   textCE23.Locked = bEnable
   textCE24.Locked = bEnable
   textCE25.Locked = bEnable
   textCE26.Locked = bEnable
   textCE27.Locked = bEnable
   textCE28.Locked = bEnable
   textCE29.Locked = bEnable
   textCE30.Locked = bEnable
   textCE31.Locked = bEnable
   textCE32.Locked = bEnable
   textCE33.Locked = bEnable
   textCE34.Locked = bEnable
   textCE35.Locked = bEnable
   textCE36.Locked = bEnable
   textCE37.Locked = bEnable
   textCE38.Locked = bEnable
   textCE39.Locked = bEnable
   textCE40.Locked = bEnable
   textCE41.Locked = bEnable
   textCE42.Locked = bEnable
   textCE43.Locked = bEnable
   textCE44.Locked = bEnable
   textCE45.Locked = bEnable
   textCE46.Locked = bEnable
   textCE47.Locked = bEnable
   textCE48.Locked = bEnable
   textCE49.Locked = bEnable
   textCE50.Locked = bEnable
   textCE51.Locked = bEnable
   textCE52.Locked = bEnable
   textCE53.Locked = bEnable
   textCE54.Locked = bEnable
   textCE55.Locked = bEnable
   textCE56.Locked = bEnable
   textCE57.Locked = bEnable
   textCE58.Locked = bEnable
   textCE59.Locked = bEnable
   textCE60.Locked = bEnable
   textCE61.Locked = bEnable
   textCE62.Locked = bEnable
   textCE63.Locked = bEnable
   textCE64.Locked = bEnable
   textCE65.Locked = bEnable
   textCE66.Locked = bEnable
   textCE67.Locked = bEnable
   textCE68.Locked = bEnable
   textCE69.Locked = bEnable
   textCE70.Locked = bEnable
   textCE71.Locked = bEnable
   textCE72.Locked = bEnable
   textCE73.Locked = bEnable
   textCE74.Locked = bEnable
   textCE75.Locked = bEnable
   textCE76.Locked = bEnable
   textCE77.Locked = bEnable
   textCE78.Locked = bEnable
   textCE79.Locked = bEnable
   textCE80.Locked = bEnable
   textCE81.Locked = bEnable
   textCE82.Locked = bEnable
   textCE83.Locked = bEnable
   textCE84.Locked = bEnable
   textCE85.Locked = bEnable
   textCE86.Locked = bEnable
   textCE87.Locked = bEnable
   textCE88.Locked = bEnable
   textCE89.Locked = bEnable
   textCE90.Locked = bEnable
   textCE91.Locked = bEnable
   textCE92.Locked = bEnable
   textCE93.Locked = bEnable
   textCE94.Locked = bEnable
   textCE95.Locked = bEnable
   textCE96.Locked = bEnable
   textCE97.Locked = bEnable
   textCE98.Locked = bEnable
   textCE99.Locked = bEnable
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm100113_3 = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(strCE01 As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM CHANGEEVENT " & _
            "WHERE CE01 = '" & strCE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("CE01")) = False Then: textCE01 = rsTmp.Fields("CE01")
      If IsNull(rsTmp.Fields("CE02")) = False Then
         If IsEmptyText(rsTmp.Fields("CE02")) = False Then
            textCE02 = TAIWANDATE(rsTmp.Fields("CE02"))
         End If
      End If
      If IsNull(rsTmp.Fields("CE03")) = False Then: textCE03 = rsTmp.Fields("CE03")
      If IsNull(rsTmp.Fields("CE04")) = False Then: textCE04 = rsTmp.Fields("CE04")
      If IsNull(rsTmp.Fields("CE05")) = False Then: textCE05 = rsTmp.Fields("CE05")
      If IsNull(rsTmp.Fields("CE06")) = False Then: textCE06 = rsTmp.Fields("CE06")
      If IsNull(rsTmp.Fields("CE07")) = False Then: textCE07 = rsTmp.Fields("CE07")
      If IsNull(rsTmp.Fields("CE08")) = False Then: textCE08 = rsTmp.Fields("CE08")
      If IsNull(rsTmp.Fields("CE09")) = False Then: textCE09 = rsTmp.Fields("CE09")
      If IsNull(rsTmp.Fields("CE10")) = False Then: textCE10 = rsTmp.Fields("CE10")
      If IsNull(rsTmp.Fields("CE11")) = False Then: textCE11 = rsTmp.Fields("CE11")
      If IsNull(rsTmp.Fields("CE12")) = False Then: textCE12 = rsTmp.Fields("CE12")
      If IsNull(rsTmp.Fields("CE13")) = False Then: textCE13 = rsTmp.Fields("CE13")
      If IsNull(rsTmp.Fields("CE14")) = False Then: textCE14 = rsTmp.Fields("CE14")
      If IsNull(rsTmp.Fields("CE15")) = False Then: textCE15 = rsTmp.Fields("CE15")
      If IsNull(rsTmp.Fields("CE16")) = False Then: textCE16 = rsTmp.Fields("CE16")
      If IsNull(rsTmp.Fields("CE17")) = False Then: textCE17 = rsTmp.Fields("CE17")
      If IsNull(rsTmp.Fields("CE18")) = False Then: textCE18 = rsTmp.Fields("CE18")
      If IsNull(rsTmp.Fields("CE19")) = False Then: textCE19 = rsTmp.Fields("CE19")
      If IsNull(rsTmp.Fields("CE20")) = False Then: textCE20 = rsTmp.Fields("CE20")
      If IsNull(rsTmp.Fields("CE21")) = False Then: textCE21 = rsTmp.Fields("CE21")
      If IsNull(rsTmp.Fields("CE22")) = False Then: textCE22 = rsTmp.Fields("CE22")
      If IsNull(rsTmp.Fields("CE23")) = False Then: textCE23 = rsTmp.Fields("CE23")
      If IsNull(rsTmp.Fields("CE24")) = False Then: textCE24 = rsTmp.Fields("CE24")
      If IsNull(rsTmp.Fields("CE25")) = False Then: textCE25 = rsTmp.Fields("CE25")
      If IsNull(rsTmp.Fields("CE26")) = False Then: textCE26 = rsTmp.Fields("CE26")
      If IsNull(rsTmp.Fields("CE27")) = False Then: textCE27 = rsTmp.Fields("CE27")
      If IsNull(rsTmp.Fields("CE28")) = False Then: textCE28 = rsTmp.Fields("CE28")
      If IsNull(rsTmp.Fields("CE29")) = False Then: textCE29 = rsTmp.Fields("CE29")
      If IsNull(rsTmp.Fields("CE30")) = False Then: textCE30 = rsTmp.Fields("CE30")
      If IsNull(rsTmp.Fields("CE31")) = False Then: textCE31 = rsTmp.Fields("CE31")
      If IsNull(rsTmp.Fields("CE32")) = False Then: textCE32 = rsTmp.Fields("CE32")
      If IsNull(rsTmp.Fields("CE33")) = False Then: textCE33 = rsTmp.Fields("CE33")
      If IsNull(rsTmp.Fields("CE34")) = False Then: textCE34 = rsTmp.Fields("CE34")
      If IsNull(rsTmp.Fields("CE35")) = False Then: textCE35 = rsTmp.Fields("CE35")
      If IsNull(rsTmp.Fields("CE36")) = False Then: textCE36 = rsTmp.Fields("CE36")
      If IsNull(rsTmp.Fields("CE37")) = False Then: textCE37 = rsTmp.Fields("CE37")
      If IsNull(rsTmp.Fields("CE38")) = False Then: textCE38 = rsTmp.Fields("CE38")
      If IsNull(rsTmp.Fields("CE39")) = False Then: textCE39 = rsTmp.Fields("CE39")
      If IsNull(rsTmp.Fields("CE40")) = False Then: textCE40 = rsTmp.Fields("CE40")
      If IsNull(rsTmp.Fields("CE41")) = False Then: textCE41 = rsTmp.Fields("CE41")
      If IsNull(rsTmp.Fields("CE42")) = False Then: textCE42 = rsTmp.Fields("CE42")
      If IsNull(rsTmp.Fields("CE43")) = False Then: textCE43 = rsTmp.Fields("CE43")
      If IsNull(rsTmp.Fields("CE44")) = False Then: textCE44 = rsTmp.Fields("CE44")
      If IsNull(rsTmp.Fields("CE45")) = False Then: textCE45 = rsTmp.Fields("CE45")
      If IsNull(rsTmp.Fields("CE46")) = False Then: textCE46 = rsTmp.Fields("CE46")
      If IsNull(rsTmp.Fields("CE47")) = False Then: textCE47 = rsTmp.Fields("CE47")
      If IsNull(rsTmp.Fields("CE48")) = False Then: textCE48 = rsTmp.Fields("CE48")
      If IsNull(rsTmp.Fields("CE49")) = False Then: textCE49 = rsTmp.Fields("CE49")
      If IsNull(rsTmp.Fields("CE50")) = False Then: textCE50 = rsTmp.Fields("CE50")
      If IsNull(rsTmp.Fields("CE51")) = False Then: textCE51 = rsTmp.Fields("CE51")
      If IsNull(rsTmp.Fields("CE52")) = False Then: textCE52 = rsTmp.Fields("CE52")
      If IsNull(rsTmp.Fields("CE53")) = False Then: textCE53 = rsTmp.Fields("CE53")
      If IsNull(rsTmp.Fields("CE54")) = False Then: textCE54 = rsTmp.Fields("CE54")
      If IsNull(rsTmp.Fields("CE55")) = False Then: textCE55 = rsTmp.Fields("CE55")
      If IsNull(rsTmp.Fields("CE56")) = False Then: textCE56 = rsTmp.Fields("CE56")
      If IsNull(rsTmp.Fields("CE57")) = False Then: textCE57 = rsTmp.Fields("CE57")
      If IsNull(rsTmp.Fields("CE58")) = False Then: textCE58 = rsTmp.Fields("CE58")
      If IsNull(rsTmp.Fields("CE59")) = False Then: textCE59 = rsTmp.Fields("CE59")
      If IsNull(rsTmp.Fields("CE60")) = False Then: textCE60 = rsTmp.Fields("CE60")
      If IsNull(rsTmp.Fields("CE61")) = False Then: textCE61 = rsTmp.Fields("CE61")
      If IsNull(rsTmp.Fields("CE62")) = False Then: textCE62 = rsTmp.Fields("CE62")
      If IsNull(rsTmp.Fields("CE63")) = False Then: textCE63 = rsTmp.Fields("CE63")
      If IsNull(rsTmp.Fields("CE64")) = False Then: textCE64 = rsTmp.Fields("CE64")
      If IsNull(rsTmp.Fields("CE65")) = False Then: textCE65 = rsTmp.Fields("CE65")
      If IsNull(rsTmp.Fields("CE66")) = False Then: textCE66 = rsTmp.Fields("CE66")
      If IsNull(rsTmp.Fields("CE67")) = False Then: textCE67 = rsTmp.Fields("CE67")
      If IsNull(rsTmp.Fields("CE68")) = False Then: textCE68 = rsTmp.Fields("CE68")
      If IsNull(rsTmp.Fields("CE69")) = False Then: textCE69 = rsTmp.Fields("CE69")
      If IsNull(rsTmp.Fields("CE70")) = False Then: textCE70 = rsTmp.Fields("CE70")
      If IsNull(rsTmp.Fields("CE71")) = False Then: textCE71 = rsTmp.Fields("CE71")
      If IsNull(rsTmp.Fields("CE72")) = False Then: textCE72 = rsTmp.Fields("CE72")
      If IsNull(rsTmp.Fields("CE73")) = False Then: textCE73 = rsTmp.Fields("CE73")
      If IsNull(rsTmp.Fields("CE74")) = False Then: textCE74 = rsTmp.Fields("CE74")
      If IsNull(rsTmp.Fields("CE75")) = False Then: textCE75 = rsTmp.Fields("CE75")
      If IsNull(rsTmp.Fields("CE76")) = False Then: textCE76 = rsTmp.Fields("CE76")
      If IsNull(rsTmp.Fields("CE77")) = False Then: textCE77 = rsTmp.Fields("CE77")
      If IsNull(rsTmp.Fields("CE78")) = False Then: textCE78 = rsTmp.Fields("CE78")
      If IsNull(rsTmp.Fields("CE79")) = False Then: textCE79 = rsTmp.Fields("CE79")
      If IsNull(rsTmp.Fields("CE80")) = False Then: textCE80 = rsTmp.Fields("CE80")
      If IsNull(rsTmp.Fields("CE81")) = False Then: textCE81 = rsTmp.Fields("CE81")
      If IsNull(rsTmp.Fields("CE82")) = False Then: textCE82 = rsTmp.Fields("CE82")
      If IsNull(rsTmp.Fields("CE83")) = False Then: textCE83 = rsTmp.Fields("CE83")
      If IsNull(rsTmp.Fields("CE84")) = False Then: textCE84 = rsTmp.Fields("CE84")
      If IsNull(rsTmp.Fields("CE85")) = False Then: textCE85 = rsTmp.Fields("CE85")
      If IsNull(rsTmp.Fields("CE86")) = False Then: textCE86 = rsTmp.Fields("CE86")
      If IsNull(rsTmp.Fields("CE87")) = False Then: textCE87 = rsTmp.Fields("CE87")
      If IsNull(rsTmp.Fields("CE88")) = False Then: textCE88 = rsTmp.Fields("CE88")
      If IsNull(rsTmp.Fields("CE89")) = False Then: textCE89 = rsTmp.Fields("CE89")
      If IsNull(rsTmp.Fields("CE90")) = False Then: textCE90 = rsTmp.Fields("CE90")
      If IsNull(rsTmp.Fields("CE91")) = False Then: textCE91 = rsTmp.Fields("CE91")
      If IsNull(rsTmp.Fields("CE92")) = False Then: textCE92 = rsTmp.Fields("CE92")
      If IsNull(rsTmp.Fields("CE93")) = False Then: textCE93 = rsTmp.Fields("CE93")
      If IsNull(rsTmp.Fields("CE94")) = False Then: textCE94 = rsTmp.Fields("CE94")
      If IsNull(rsTmp.Fields("CE95")) = False Then: textCE95 = rsTmp.Fields("CE95")
      If IsNull(rsTmp.Fields("CE96")) = False Then: textCE96 = rsTmp.Fields("CE96")
      If IsNull(rsTmp.Fields("CE97")) = False Then: textCE97 = rsTmp.Fields("CE97")
      If IsNull(rsTmp.Fields("CE98")) = False Then: textCE98 = rsTmp.Fields("CE98")
      If IsNull(rsTmp.Fields("CE99")) = False Then: textCE99 = rsTmp.Fields("CE99")
      
      ' 更新本所案號
      UpdateCPData "" & rsTmp.Fields("CE01")
            
        textCE04_Validate False
        textCE05_Validate False
        textCE06_Validate False
        textCE07_Validate False
        textCE08_Validate False
        textCE39_Validate False
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub UpdateCPData(ByVal strData As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & strData & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CP01")) = False Then
         textCP01 = rsTmp.Fields("CP01")
      End If
      If IsNull(rsTmp.Fields("CP02")) = False Then
         textCP02 = rsTmp.Fields("CP02")
      End If
      If IsNull(rsTmp.Fields("CP03")) = False Then
         textCP03 = rsTmp.Fields("CP03")
      End If
      If IsNull(rsTmp.Fields("CP04")) = False Then
         textCP04 = rsTmp.Fields("CP04")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 申請人1
Private Sub textCE04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCE04_2 = Empty
   If IsEmptyText(textCE04) = False Then
      textCE04_2 = GetCustomerName(textCE04, "0")
      If IsEmptyText(textCE04_2) = True Then
            strTit = "檢核資料"
            strMsg = "申請人1代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

' 申請人2
Private Sub textCE05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCE05_2 = Empty
   If IsEmptyText(textCE05) = False Then
      textCE05_2 = GetCustomerName(textCE05, "0")
      If IsEmptyText(textCE05_2) = True Then
            strTit = "檢核資料"
            strMsg = "申請人2代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

' 申請人3
Private Sub textCE06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCE06_2 = Empty
   If IsEmptyText(textCE06) = False Then
      textCE06_2 = GetCustomerName(textCE06, "0")
      If IsEmptyText(textCE06_2) = True Then
            strTit = "檢核資料"
            strMsg = "申請人3代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

' 申請人4
Private Sub textCE07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCE07_2 = Empty
   If IsEmptyText(textCE07) = False Then
      textCE07_2 = GetCustomerName(textCE07, "0")
      If IsEmptyText(textCE07_2) = True Then
            strTit = "檢核資料"
            strMsg = "申請人4代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

' 申請人5
Private Sub textCE08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCE08_2 = Empty
   If IsEmptyText(textCE08) = False Then
      textCE08_2 = GetCustomerName(textCE08, "0")
      If IsEmptyText(textCE08_2) = True Then
            strTit = "檢核資料"
            strMsg = "申請人5代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

' 專利商標種類代號
Private Sub textCE39_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strNation As String
   
   Cancel = False
   textCE39_2 = Empty
   strNation = GetPrjNation(textCP01 & "-" & textCP02 & "-" & textCP03 & "-" & textCP04) 'Add By Sindy 2015/8/13
   If IsEmptyText(textCE39) = False Then
      Select Case textCP01
         Case "CFT", "FCT", "T", "TF":
            'Modify By Sindy 2015/8/13
            'textCE39_2 = GetTradeMarkName(textCE39, 0)
            textCE39_2 = GetTradeMarkName(textCE39, IIf(strNation = "020", 1, 0))
            '2015/8/13 END
            If IsEmptyText(textCE39_2) = True Then
                strTit = "檢核資料"
                strMsg = "商標種類代碼不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            End If
         Case "P", "CFP", "FCP":
            textCE39_2 = GetPatentName(textCE39, 0)
            If IsEmptyText(textCE39_2) = True Then
                strTit = "檢核資料"
                strMsg = "專利種類代碼不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            End If
         Case Else:
      End Select
   End If
End Sub

