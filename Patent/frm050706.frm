VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050706 
   BorderStyle     =   1  '單線固定
   Caption         =   "變更事項"
   ClientHeight    =   6345
   ClientLeft      =   30
   ClientTop       =   975
   ClientWidth     =   9150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9150
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   150
      TabIndex        =   100
      Top             =   840
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9631
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "第1頁"
      TabPicture(0)   =   "frm050706.frx":0000
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
      Tab(0).Control(31)=   "textCE21"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCE20"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textCE19"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCE18"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textCE17"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textCE61"
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
      TabPicture(1)   =   "frm050706.frx":001C
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
      TabPicture(2)   =   "frm050706.frx":0038
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
      Tab(2).Control(9)=   "textCE41_1"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "textCE41"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "textCE42"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "textCE43"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "textCE45"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "textCE63"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "textCE64"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "textCE92"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "textCE93"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "textCE94"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "textCE95"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "textCE96"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "textCE97"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "textCE98"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "textCE99"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Label5(0)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Label4(42)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Label4(41)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Label4(40)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Label4(39)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Label4(38)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Label4(37)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Label4(36)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Label4(35)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Label3(7)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Label4(33)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Label4(32)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Label3(15)"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Label3(11)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Label3(10)"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Label3(9)"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Label3(5)"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Label5(11)"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Label5(5)"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Label5(4)"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "Label5(3)"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "Label4(31)"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "Label70(5)"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "Label8"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).ControlCount=   48
      TabCaption(3)   =   "第4頁"
      TabPicture(3)   =   "frm050706.frx":0054
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
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   252
         Left            =   4020
         TabIndex        =   189
         Top             =   4455
         Width           =   4332
      End
      Begin VB.TextBox textCE16 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   69
         Top             =   408
         Width           =   255
      End
      Begin VB.TextBox textCE66 
         Height          =   270
         Left            =   -71910
         MaxLength       =   25
         TabIndex        =   44
         Top             =   5025
         Width           =   2172
      End
      Begin VB.TextBox textCE67 
         Height          =   270
         Left            =   -73710
         MaxLength       =   1
         TabIndex        =   43
         Top             =   5025
         Width           =   255
      End
      Begin VB.TextBox textCE62 
         Height          =   270
         Left            =   1380
         MaxLength       =   1
         TabIndex        =   25
         Top             =   4740
         Width           =   255
      End
      Begin VB.TextBox textCE39 
         Height          =   270
         Left            =   3660
         MaxLength       =   1
         TabIndex        =   24
         Top             =   4455
         Width           =   255
      End
      Begin VB.TextBox textCE40 
         Height          =   270
         Left            =   1380
         MaxLength       =   1
         TabIndex        =   23
         Top             =   4455
         Width           =   255
      End
      Begin VB.TextBox textCE65 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   58
         Top             =   2440
         Width           =   255
      End
      Begin VB.TextBox textCE51 
         Height          =   270
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1944
         Width           =   255
      End
      Begin VB.TextBox textCE53 
         Height          =   270
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   12
         Top             =   2212
         Width           =   255
      End
      Begin VB.TextBox textCE55 
         Height          =   270
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   14
         Top             =   2480
         Width           =   255
      End
      Begin VB.TextBox textCE59 
         Height          =   270
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   16
         Top             =   2748
         Width           =   255
      End
      Begin VB.TextBox textCE60 
         Height          =   270
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   15
         Top             =   2748
         Width           =   255
      End
      Begin VB.TextBox textCE52 
         Height          =   270
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   9
         Top             =   1944
         Width           =   255
      End
      Begin VB.TextBox textCE54 
         Height          =   270
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   11
         Top             =   2212
         Width           =   255
      End
      Begin VB.TextBox textCE56 
         Height          =   270
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   13
         Top             =   2480
         Width           =   255
      End
      Begin VB.TextBox textCE58 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   56
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox textCE50 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   54
         Top             =   1880
         Width           =   255
      End
      Begin VB.TextBox textCE49 
         Height          =   270
         Left            =   -71400
         MaxLength       =   699
         TabIndex        =   55
         Top             =   1880
         Width           =   4695
      End
      Begin VB.TextBox textCE48 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   52
         Top             =   1600
         Width           =   255
      End
      Begin VB.TextBox textCE44 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   45
         Top             =   420
         Width           =   255
      End
      Begin VB.TextBox textCE57 
         Height          =   270
         Left            =   -71400
         MaxLength       =   20
         TabIndex        =   57
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox textCE47 
         Height          =   270
         Left            =   -71400
         MaxLength       =   395
         TabIndex        =   53
         Top             =   1600
         Width           =   4695
      End
      Begin VB.TextBox textCE46 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   50
         Top             =   1305
         Width           =   255
      End
      Begin VB.TextBox textCE38 
         Height          =   270
         Left            =   -73680
         MaxLength       =   1
         TabIndex        =   27
         Top             =   408
         Width           =   255
      End
      Begin VB.TextBox textCE22 
         Height          =   270
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   17
         Top             =   3016
         Width           =   255
      End
      Begin VB.TextBox textCE09 
         Height          =   270
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   3
         Top             =   604
         Width           =   255
      End
      Begin VB.TextBox textCE08 
         Height          =   270
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   8
         Top             =   1676
         Width           =   975
      End
      Begin VB.TextBox textCE07 
         Height          =   270
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   7
         Top             =   1408
         Width           =   975
      End
      Begin VB.TextBox textCE06 
         Height          =   270
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   6
         Top             =   1140
         Width           =   975
      End
      Begin VB.TextBox textCE05 
         Height          =   270
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   5
         Top             =   872
         Width           =   975
      End
      Begin VB.TextBox textCE04 
         Height          =   270
         Left            =   3240
         MaxLength       =   9
         TabIndex        =   4
         Top             =   604
         Width           =   975
      End
      Begin VB.TextBox textCE03 
         Height          =   270
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   1
         Top             =   336
         Width           =   255
      End
      Begin VB.TextBox textCE02 
         Height          =   270
         Left            =   3240
         MaxLength       =   7
         TabIndex        =   2
         Top             =   336
         Width           =   975
      End
      Begin MSForms.TextBox textCE77 
         Height          =   285
         Left            =   -73320
         TabIndex        =   85
         Top             =   2300
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE78 
         Height          =   285
         Left            =   -71040
         TabIndex        =   86
         Top             =   2300
         Width           =   2052
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "3619;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE79 
         Height          =   285
         Left            =   -68880
         TabIndex        =   87
         Top             =   2300
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE80 
         Height          =   285
         Left            =   -73320
         TabIndex        =   88
         Top             =   2616
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE81 
         Height          =   285
         Left            =   -71040
         TabIndex        =   89
         Top             =   2616
         Width           =   2052
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "3619;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE82 
         Height          =   285
         Left            =   -68880
         TabIndex        =   90
         Top             =   2616
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE83 
         Height          =   285
         Left            =   -73320
         TabIndex        =   91
         Top             =   2932
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE84 
         Height          =   285
         Left            =   -71040
         TabIndex        =   92
         Top             =   2932
         Width           =   2052
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "3619;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE85 
         Height          =   285
         Left            =   -68880
         TabIndex        =   93
         Top             =   2932
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE86 
         Height          =   285
         Left            =   -73320
         TabIndex        =   94
         Top             =   3248
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE87 
         Height          =   285
         Left            =   -71040
         TabIndex        =   95
         Top             =   3248
         Width           =   2052
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "3619;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE88 
         Height          =   285
         Left            =   -68880
         TabIndex        =   96
         Top             =   3248
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE89 
         Height          =   285
         Left            =   -73320
         TabIndex        =   97
         Top             =   3570
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE90 
         Height          =   285
         Left            =   -71040
         TabIndex        =   98
         Top             =   3570
         Width           =   2052
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "3619;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE91 
         Height          =   285
         Left            =   -68880
         TabIndex        =   99
         Top             =   3570
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE10 
         Height          =   285
         Left            =   -73320
         TabIndex        =   70
         Top             =   720
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE11 
         Height          =   285
         Left            =   -71040
         TabIndex        =   71
         Top             =   720
         Width           =   2052
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "3619;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE12 
         Height          =   285
         Left            =   -68880
         TabIndex        =   72
         Top             =   720
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE13 
         Height          =   285
         Left            =   -73320
         TabIndex        =   73
         Top             =   1036
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE14 
         Height          =   285
         Left            =   -71040
         TabIndex        =   74
         Top             =   1036
         Width           =   2052
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "3619;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE15 
         Height          =   285
         Left            =   -68880
         TabIndex        =   75
         Top             =   1036
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE68 
         Height          =   285
         Left            =   -73320
         TabIndex        =   76
         Top             =   1352
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE69 
         Height          =   285
         Left            =   -71040
         TabIndex        =   77
         Top             =   1352
         Width           =   2052
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "3619;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE70 
         Height          =   285
         Left            =   -68880
         TabIndex        =   78
         Top             =   1352
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE71 
         Height          =   285
         Left            =   -73320
         TabIndex        =   79
         Top             =   1668
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE72 
         Height          =   285
         Left            =   -71040
         TabIndex        =   80
         Top             =   1668
         Width           =   2052
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "3619;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE73 
         Height          =   285
         Left            =   -68880
         TabIndex        =   81
         Top             =   1668
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE74 
         Height          =   285
         Left            =   -73320
         TabIndex        =   82
         Top             =   1984
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE75 
         Height          =   285
         Left            =   -71040
         TabIndex        =   83
         Top             =   1984
         Width           =   2052
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "3619;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE76 
         Height          =   285
         Left            =   -68880
         TabIndex        =   84
         Top             =   1984
         Width           =   2175
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41_1 
         Height          =   915
         Left            =   -71400
         TabIndex        =   46
         Top             =   420
         Width           =   5205
         VariousPropertyBits=   -1475330021
         ScrollBars      =   2
         Size            =   "9181;1614"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41 
         Height          =   285
         Left            =   -71400
         TabIndex        =   47
         Top             =   420
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE42 
         Height          =   285
         Left            =   -71400
         TabIndex        =   48
         Top             =   715
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   180
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE43 
         Height          =   285
         Left            =   -71400
         TabIndex        =   49
         Top             =   1010
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE45 
         Height          =   285
         Left            =   -71400
         TabIndex        =   51
         Top             =   1305
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE63 
         Height          =   285
         Left            =   -71400
         TabIndex        =   59
         Top             =   2440
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE64 
         Height          =   285
         Left            =   -71400
         TabIndex        =   60
         Top             =   2735
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE92 
         Height          =   285
         Left            =   -71400
         TabIndex        =   61
         Top             =   3030
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE93 
         Height          =   285
         Left            =   -71400
         TabIndex        =   62
         Top             =   3325
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE94 
         Height          =   285
         Left            =   -71400
         TabIndex        =   63
         Top             =   3620
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE95 
         Height          =   285
         Left            =   -71400
         TabIndex        =   64
         Top             =   3915
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE96 
         Height          =   285
         Left            =   -71400
         TabIndex        =   65
         Top             =   4210
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE97 
         Height          =   285
         Left            =   -71400
         TabIndex        =   66
         Top             =   4505
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE98 
         Height          =   285
         Left            =   -71400
         TabIndex        =   67
         Top             =   4800
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE99 
         Height          =   285
         Left            =   -71400
         TabIndex        =   68
         Top             =   5100
         Width           =   5205
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE61 
         Height          =   585
         Left            =   3240
         TabIndex        =   26
         Top             =   4740
         Width           =   5505
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         Size            =   "9710;1032"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE23 
         Height          =   285
         Left            =   -71910
         TabIndex        =   28
         Top             =   390
         Width           =   5595
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "9869;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE24 
         Height          =   285
         Left            =   -71910
         TabIndex        =   29
         Top             =   690
         Width           =   5595
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "9869;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE25 
         Height          =   285
         Left            =   -71910
         TabIndex        =   30
         Top             =   990
         Width           =   5595
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "9869;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE26 
         Height          =   285
         Left            =   -71910
         TabIndex        =   31
         Top             =   1320
         Width           =   5595
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "9869;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE27 
         Height          =   285
         Left            =   -71910
         TabIndex        =   32
         Top             =   1620
         Width           =   5595
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "9869;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE28 
         Height          =   285
         Left            =   -71910
         TabIndex        =   33
         Top             =   1920
         Width           =   5595
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "9869;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE29 
         Height          =   285
         Left            =   -71910
         TabIndex        =   34
         Top             =   2250
         Width           =   5595
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "9869;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE30 
         Height          =   285
         Left            =   -71910
         TabIndex        =   35
         Top             =   2550
         Width           =   5595
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "9869;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE31 
         Height          =   285
         Left            =   -71910
         TabIndex        =   36
         Top             =   2850
         Width           =   5595
         VariousPropertyBits=   675299355
         MaxLength       =   70
         Size            =   "9869;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE32 
         Height          =   285
         Left            =   -71910
         TabIndex        =   37
         Top             =   3180
         Width           =   5600
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "9878;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE33 
         Height          =   285
         Left            =   -71910
         TabIndex        =   38
         Top             =   3480
         Width           =   5600
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "9878;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE34 
         Height          =   285
         Left            =   -71910
         TabIndex        =   39
         Top             =   3780
         Width           =   5600
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "9878;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE35 
         Height          =   285
         Left            =   -71910
         TabIndex        =   40
         Top             =   4110
         Width           =   5600
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "9878;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE36 
         Height          =   285
         Left            =   -71910
         TabIndex        =   41
         Top             =   4410
         Width           =   5600
         VariousPropertyBits=   671105051
         MaxLength       =   154
         Size            =   "9878;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE37 
         Height          =   285
         Left            =   -71910
         TabIndex        =   42
         Top             =   4710
         Width           =   5600
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "9878;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE17 
         Height          =   285
         Left            =   3240
         TabIndex        =   18
         Top             =   3016
         Width           =   5505
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9710;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE18 
         Height          =   285
         Left            =   3240
         TabIndex        =   19
         Top             =   3299
         Width           =   5505
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9710;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE19 
         Height          =   285
         Left            =   3240
         TabIndex        =   20
         Top             =   3582
         Width           =   5505
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9710;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE20 
         Height          =   285
         Left            =   3240
         TabIndex        =   21
         Top             =   3865
         Width           =   5505
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9710;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE21 
         Height          =   285
         Left            =   3240
         TabIndex        =   22
         Top             =   4140
         Width           =   5505
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9710;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE04_2 
         Height          =   285
         Left            =   4290
         TabIndex        =   195
         TabStop         =   0   'False
         Top             =   604
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
         Left            =   4290
         TabIndex        =   194
         TabStop         =   0   'False
         Top             =   865
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
         Left            =   4290
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   1140
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
         Left            =   4290
         TabIndex        =   192
         TabStop         =   0   'False
         Top             =   1408
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
         Left            =   4290
         TabIndex        =   191
         TabStop         =   0   'False
         Top             =   1676
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
      Begin VB.Label Label5 
         Caption         =   "案件名稱:"
         Height          =   255
         Index           =   0
         Left            =   -72960
         TabIndex        =   190
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "代表人10:"
         Height          =   252
         Index           =   49
         Left            =   -74280
         TabIndex        =   188
         Top             =   3570
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "代表人9:"
         Height          =   252
         Index           =   48
         Left            =   -74280
         TabIndex        =   187
         Top             =   3248
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "代表人8:"
         Height          =   252
         Index           =   47
         Left            =   -74280
         TabIndex        =   186
         Top             =   2932
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "代表人7:"
         Height          =   252
         Index           =   46
         Left            =   -74280
         TabIndex        =   185
         Top             =   2616
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "代表人6:"
         Height          =   252
         Index           =   45
         Left            =   -74280
         TabIndex        =   184
         Top             =   2300
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "代表人5:"
         Height          =   252
         Index           =   44
         Left            =   -74280
         TabIndex        =   183
         Top             =   1984
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "代表人4:"
         Height          =   252
         Index           =   15
         Left            =   -74280
         TabIndex        =   182
         Top             =   1668
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "代表人3:"
         Height          =   252
         Index           =   14
         Left            =   -74280
         TabIndex        =   181
         Top             =   1352
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "代表人2:"
         Height          =   252
         Index           =   7
         Left            =   -74280
         TabIndex        =   180
         Top             =   1036
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "代表人1:"
         Height          =   252
         Index           =   43
         Left            =   -74280
         TabIndex        =   179
         Top             =   720
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "代表人中譯文10:"
         Height          =   255
         Index           =   42
         Left            =   -72960
         TabIndex        =   178
         Top             =   5100
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "代表人中譯文9:"
         Height          =   255
         Index           =   41
         Left            =   -72960
         TabIndex        =   177
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "代表人中譯文8:"
         Height          =   255
         Index           =   40
         Left            =   -72960
         TabIndex        =   176
         Top             =   4505
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "代表人中譯文7:"
         Height          =   255
         Index           =   39
         Left            =   -72960
         TabIndex        =   175
         Top             =   4210
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "代表人中譯文6:"
         Height          =   255
         Index           =   38
         Left            =   -72960
         TabIndex        =   174
         Top             =   3915
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "代表人中譯文5:"
         Height          =   255
         Index           =   37
         Left            =   -72960
         TabIndex        =   173
         Top             =   3620
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "代表人中譯文4:"
         Height          =   255
         Index           =   36
         Left            =   -72960
         TabIndex        =   172
         Top             =   3325
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "代表人中譯文3:"
         Height          =   255
         Index           =   35
         Left            =   -72960
         TabIndex        =   171
         Top             =   3030
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   252
         Index           =   1
         Left            =   -74880
         TabIndex        =   170
         Top             =   408
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "代表人(中):"
         Height          =   252
         Index           =   1
         Left            =   -72720
         TabIndex        =   169
         Top             =   408
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "代表人(英):"
         Height          =   252
         Index           =   12
         Left            =   -70440
         TabIndex        =   168
         Top             =   408
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "代表人(日):"
         Height          =   252
         Index           =   13
         Left            =   -68280
         TabIndex        =   167
         Top             =   408
         Width           =   1092
      End
      Begin VB.Label Label4 
         Caption         =   "網域密碼 :"
         Height          =   255
         Index           =   34
         Left            =   -73230
         TabIndex        =   166
         Top             =   5025
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   255
         Index           =   8
         Left            =   -74790
         TabIndex        =   165
         Top             =   5025
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "其他:"
         Height          =   252
         Index           =   6
         Left            =   1920
         TabIndex        =   164
         Top             =   4740
         Width           =   852
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   252
         Index           =   17
         Left            =   180
         TabIndex        =   163
         Top             =   4740
         Width           =   972
      End
      Begin VB.Label Label5 
         Caption         =   "專利商標種類代號 :"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   162
         Top             =   4455
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   161
         Top             =   4455
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   156
         Top             =   2440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "代表人中譯文2:"
         Height          =   255
         Index           =   33
         Left            =   -72960
         TabIndex        =   155
         Top             =   2735
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "代表人中譯文1:"
         Height          =   255
         Index           =   32
         Left            =   -72960
         TabIndex        =   154
         Top             =   2440
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "代理人:"
         Height          =   252
         Index           =   1
         Left            =   1920
         TabIndex        =   153
         Top             =   2480
         Width           =   732
      End
      Begin VB.Label Label5 
         Caption         =   "圖樣:"
         Height          =   252
         Index           =   12
         Left            =   1920
         TabIndex        =   152
         Top             =   2748
         Width           =   1092
      End
      Begin VB.Label Label5 
         Caption         =   "代表人印鑑:"
         Height          =   252
         Index           =   10
         Left            =   1920
         TabIndex        =   151
         Top             =   2212
         Width           =   1092
      End
      Begin VB.Label Label5 
         Caption         =   "申請人印鑑:"
         Height          =   252
         Index           =   9
         Left            =   1920
         TabIndex        =   150
         Top             =   1944
         Width           =   1092
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   252
         Index           =   12
         Left            =   240
         TabIndex        =   149
         Top             =   1944
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   252
         Index           =   13
         Left            =   240
         TabIndex        =   148
         Top             =   2212
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   252
         Index           =   14
         Left            =   240
         TabIndex        =   147
         Top             =   2480
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   252
         Index           =   16
         Left            =   240
         TabIndex        =   146
         Top             =   2748
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   255
         Index           =   15
         Left            =   -74760
         TabIndex        =   144
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   143
         Top             =   1880
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   255
         Index           =   10
         Left            =   -74760
         TabIndex        =   142
         Top             =   1600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   141
         Top             =   1305
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   252
         Index           =   5
         Left            =   -74760
         TabIndex        =   140
         Top             =   420
         Width           =   972
      End
      Begin VB.Label Label5 
         Caption         =   "正商標號數:"
         Height          =   255
         Index           =   11
         Left            =   -72960
         TabIndex        =   139
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "案件日文名稱:"
         Height          =   252
         Index           =   5
         Left            =   -72960
         TabIndex        =   138
         Top             =   1010
         Width           =   1212
      End
      Begin VB.Label Label5 
         Caption         =   "案件英文名稱:"
         Height          =   252
         Index           =   4
         Left            =   -72960
         TabIndex        =   137
         Top             =   715
         Width           =   1332
      End
      Begin VB.Label Label5 
         Caption         =   "案件中文名稱:"
         Height          =   252
         Index           =   3
         Left            =   -72960
         TabIndex        =   136
         Top             =   420
         Width           =   1452
      End
      Begin VB.Label Label4 
         Caption         =   "減縮商品:"
         Height          =   255
         Index           =   31
         Left            =   -72960
         TabIndex        =   135
         Top             =   1305
         Width           =   975
      End
      Begin VB.Label Label70 
         Caption         =   "商品類別:"
         Height          =   255
         Index           =   5
         Left            =   -72960
         TabIndex        =   134
         Top             =   1600
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "商品組群:"
         Height          =   255
         Left            =   -72960
         TabIndex        =   133
         Top             =   1880
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   132
         Top             =   408
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(中)5:"
         Height          =   252
         Index           =   30
         Left            =   -73200
         TabIndex        =   131
         Top             =   4110
         Width           =   1332
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(英)5:"
         Height          =   252
         Index           =   29
         Left            =   -73200
         TabIndex        =   130
         Top             =   4410
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(日)5:"
         Height          =   252
         Index           =   28
         Left            =   -73200
         TabIndex        =   129
         Top             =   4710
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(中)3:"
         Height          =   252
         Index           =   27
         Left            =   -73200
         TabIndex        =   128
         Top             =   2250
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(英)3:"
         Height          =   252
         Index           =   26
         Left            =   -73200
         TabIndex        =   127
         Top             =   2550
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(日)3:"
         Height          =   252
         Index           =   25
         Left            =   -73200
         TabIndex        =   126
         Top             =   2850
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(中)4:"
         Height          =   252
         Index           =   24
         Left            =   -73200
         TabIndex        =   125
         Top             =   3180
         Width           =   1332
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(英)4:"
         Height          =   252
         Index           =   23
         Left            =   -73200
         TabIndex        =   124
         Top             =   3480
         Width           =   1572
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(日)4:"
         Height          =   252
         Index           =   22
         Left            =   -73200
         TabIndex        =   123
         Top             =   3780
         Width           =   1812
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(中)1:"
         Height          =   255
         Index           =   21
         Left            =   -73200
         TabIndex        =   122
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(英)1:"
         Height          =   252
         Index           =   20
         Left            =   -73200
         TabIndex        =   121
         Top             =   690
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(日)1:"
         Height          =   252
         Index           =   19
         Left            =   -73200
         TabIndex        =   120
         Top             =   990
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(中)2:"
         Height          =   252
         Index           =   18
         Left            =   -73200
         TabIndex        =   119
         Top             =   1320
         Width           =   1332
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(英)2:"
         Height          =   252
         Index           =   17
         Left            =   -73200
         TabIndex        =   118
         Top             =   1620
         Width           =   1572
      End
      Begin VB.Label Label4 
         Caption         =   "申請地址(日)2:"
         Height          =   252
         Index           =   16
         Left            =   -73200
         TabIndex        =   117
         Top             =   1920
         Width           =   1812
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   116
         Top             =   3016
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "申請人1中譯文:"
         Height          =   252
         Index           =   2
         Left            =   1920
         TabIndex        =   115
         Top             =   3016
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "申請人2中譯文:"
         Height          =   252
         Index           =   8
         Left            =   1920
         TabIndex        =   114
         Top             =   3299
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "申請人3中譯文:"
         Height          =   252
         Index           =   9
         Left            =   1920
         TabIndex        =   113
         Top             =   3582
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "申請人4中譯文:"
         Height          =   252
         Index           =   10
         Left            =   1920
         TabIndex        =   112
         Top             =   3865
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "申請人5中譯文:"
         Height          =   255
         Index           =   11
         Left            =   1920
         TabIndex        =   111
         Top             =   4140
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   110
         Top             =   604
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "申請人1:"
         Height          =   252
         Index           =   0
         Left            =   1920
         TabIndex        =   109
         Top             =   604
         Width           =   732
      End
      Begin VB.Label Label4 
         Caption         =   "申請人2:"
         Height          =   252
         Index           =   3
         Left            =   1920
         TabIndex        =   108
         Top             =   881
         Width           =   732
      End
      Begin VB.Label Label4 
         Caption         =   "申請人3:"
         Height          =   252
         Index           =   4
         Left            =   1920
         TabIndex        =   107
         Top             =   1140
         Width           =   732
      End
      Begin VB.Label Label4 
         Caption         =   "申請人4:"
         Height          =   252
         Index           =   5
         Left            =   1920
         TabIndex        =   106
         Top             =   1408
         Width           =   732
      End
      Begin VB.Label Label4 
         Caption         =   "申請人5:"
         Height          =   252
         Index           =   6
         Left            =   1920
         TabIndex        =   105
         Top             =   1676
         Width           =   732
      End
      Begin VB.Label Label3 
         Caption         =   "准(1)/駁(2) :"
         Height          =   252
         Index           =   6
         Left            =   240
         TabIndex        =   104
         Top             =   336
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "申請日:"
         Height          =   252
         Left            =   1920
         TabIndex        =   103
         Top             =   336
         Width           =   732
      End
   End
   Begin VB.TextBox textCP04 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   5760
      MaxLength       =   2
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   564
      Width           =   372
   End
   Begin VB.TextBox textCP03 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   5520
      MaxLength       =   1
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   564
      Width           =   252
   End
   Begin VB.TextBox textCP02 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   4800
      MaxLength       =   6
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   564
      Width           =   732
   End
   Begin VB.TextBox textCP01 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   564
      Width           =   492
   End
   Begin VB.TextBox textCE01 
      Height          =   270
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   0
      Top             =   564
      Width           =   1212
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8460
      Top             =   600
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
            Picture         =   "frm050706.frx":0070
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050706.frx":038C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050706.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050706.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050706.frx":0BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050706.frx":0EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050706.frx":11D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050706.frx":14F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050706.frx":1810
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050706.frx":1B2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050706.frx":1E48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
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
   Begin VB.Label Label6 
      Caption         =   "本所案號:"
      Height          =   255
      Left            =   3240
      TabIndex        =   145
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "收文號:"
      Height          =   255
      Left            =   240
      TabIndex        =   102
      Top             =   630
      Width           =   975
   End
End
Attribute VB_Name = "frm050706"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/12 改成Form2.0 ; textCE04_2、textCE05_2、textCE06_2、textCE07_2、textCE08_2、textCE17~21、textCE63~64、textCE92~99、textCE10~15、textCE68~91、textCE23~37、textCE41~43、textCE45、textCE61
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit
Const MAX_FIELD = 99

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
Dim m_SubMode As Integer

' 辦識其為外商還是內商的程式
' 0 表內商
' 1 表外商
Dim m_SysKind As Integer

' 第一筆資料的本所案號
Dim m_FirstKEY As String
' 最後一筆資料的本所案號
Dim m_LastKEY As String
' 目前正在顯示的本所案號
Dim m_CurrKEY As String

'edit by nick 2004/10/14
' 90.07.13 modify by louis (執行各項功能的權限)
'Dim m_bInsert As Boolean
'Dim m_bUpdate As Boolean
'Dim m_bDelete As Boolean
'Dim m_bQuery As Boolean
Public m_bInsert As Boolean
Public m_bUpdate As Boolean
Public m_bDelete As Boolean
Public m_bQuery As Boolean
Public IsCall As Boolean
Public IsMod As Boolean
Public frmParent As Form 'Add by Morgan 2007/5/22

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT CE01 FROM CHANGEEVENT " & _
            "WHERE CE01 = (SELECT MIN(CE01) FROM CHANGEEVENT) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CE01")) = False Then: m_FirstKEY = rsTmp.Fields("CE01")
   End If
   rsTmp.Close

   strSql = "SELECT CE01 FROM CHANGEEVENT " & _
            "WHERE CE01 = (SELECT MAX(CE01) FROM CHANGEEVENT) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CE01")) = False Then: m_LastKEY = rsTmp.Fields("CE01")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' Load Form
Private Sub Form_Load()
   SSTab1.Tab = 0
   'add by nick 2004/10/14 設定是否被呼叫
   IsCall = False
   IsMod = False
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm050706", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm050706", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm050706", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm050706", strFind, False)
   
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
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
   InitialField
   
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CE" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 2:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To MAX_FIELD - 1
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 更新欄位的內容
Private Sub UpdateFieldNewData()
   SetFieldNewData "CE01", textCE01
   If IsEmptyText(textCE02) = False Then
      SetFieldNewData "CE02", DBDATE(textCE02)
   Else
      SetFieldNewData "CE02", textCE02
   End If
   SetFieldNewData "CE03", textCE03
   If IsEmptyText(textCE04) = False Then
      SetFieldNewData "CE04", textCE04 & String(9 - Len(textCE04), "0")
   Else
      SetFieldNewData "CE04", textCE04
   End If
   If IsEmptyText(textCE05) = False Then
      SetFieldNewData "CE05", textCE05 & String(9 - Len(textCE05), "0")
   Else
      SetFieldNewData "CE05", textCE05
   End If
   If IsEmptyText(textCE06) = False Then
      SetFieldNewData "CE06", textCE06 & String(9 - Len(textCE06), "0")
   Else
      SetFieldNewData "CE06", textCE06
   End If
   If IsEmptyText(textCE07) = False Then
      SetFieldNewData "CE07", textCE07 & String(9 - Len(textCE07), "0")
   Else
      SetFieldNewData "CE07", textCE07
   End If
   If IsEmptyText(textCE08) = False Then
      SetFieldNewData "CE08", textCE08 & String(9 - Len(textCE08), "0")
   Else
      SetFieldNewData "CE08", textCE08
   End If
   SetFieldNewData "CE09", textCE09
   SetFieldNewData "CE10", textCE10
   SetFieldNewData "CE11", textCE11
   SetFieldNewData "CE12", textCE12
   SetFieldNewData "CE13", textCE13
   SetFieldNewData "CE14", textCE14
   SetFieldNewData "CE15", textCE15
   SetFieldNewData "CE16", textCE16
   SetFieldNewData "CE17", textCE17
   SetFieldNewData "CE18", textCE18
   SetFieldNewData "CE19", textCE19
   SetFieldNewData "CE20", textCE20
   SetFieldNewData "CE21", textCE21
   SetFieldNewData "CE22", textCE22
   SetFieldNewData "CE23", textCE23
   SetFieldNewData "CE24", textCE24
   SetFieldNewData "CE25", textCE25
   SetFieldNewData "CE26", textCE26
   SetFieldNewData "CE27", textCE27
   SetFieldNewData "CE28", textCE28
   SetFieldNewData "CE29", textCE29
   SetFieldNewData "CE30", textCE30
   SetFieldNewData "CE31", textCE31
   SetFieldNewData "CE32", textCE32
   SetFieldNewData "CE33", textCE33
   SetFieldNewData "CE34", textCE34
   SetFieldNewData "CE35", textCE35
   SetFieldNewData "CE36", textCE36
   SetFieldNewData "CE37", textCE37
   SetFieldNewData "CE38", textCE38
   SetFieldNewData "CE39", textCE39
   SetFieldNewData "CE40", textCE40
    Select Case Me.textCP01.Text
    Case "T", "FCT", "CFT", "TF"
        SetFieldNewData "CE41", textCE41_1
        DoEvents
    Case Else
        SetFieldNewData "CE41", textCE41
        SetFieldNewData "CE42", textCE42
        SetFieldNewData "CE43", textCE43
    End Select
   SetFieldNewData "CE44", textCE44
   SetFieldNewData "CE45", textCE45
   SetFieldNewData "CE46", textCE46
   SetFieldNewData "CE47", textCE47
   SetFieldNewData "CE48", textCE48
   SetFieldNewData "CE49", textCE49
   SetFieldNewData "CE50", textCE50
   SetFieldNewData "CE51", textCE51
   SetFieldNewData "CE52", textCE52
   SetFieldNewData "CE53", textCE53
   SetFieldNewData "CE54", textCE54
   SetFieldNewData "CE55", textCE55
   SetFieldNewData "CE56", textCE56
   SetFieldNewData "CE57", textCE57
   SetFieldNewData "CE58", textCE58
   SetFieldNewData "CE59", textCE59
   SetFieldNewData "CE60", textCE60
   SetFieldNewData "CE61", textCE61
   SetFieldNewData "CE62", textCE62
   SetFieldNewData "CE63", textCE63
   SetFieldNewData "CE64", textCE64
   SetFieldNewData "CE65", textCE65
   SetFieldNewData "CE66", textCE66
   SetFieldNewData "CE67", textCE67
   SetFieldNewData "CE68", textCE68
   SetFieldNewData "CE69", textCE69
   SetFieldNewData "CE70", textCE70
   SetFieldNewData "CE71", textCE71
   SetFieldNewData "CE72", textCE72
   SetFieldNewData "CE73", textCE73
   SetFieldNewData "CE74", textCE74
   SetFieldNewData "CE75", textCE75
   SetFieldNewData "CE76", textCE76
   SetFieldNewData "CE77", textCE77
   SetFieldNewData "CE78", textCE78
   SetFieldNewData "CE79", textCE79
   SetFieldNewData "CE80", textCE80
   SetFieldNewData "CE81", textCE81
   SetFieldNewData "CE82", textCE82
   SetFieldNewData "CE83", textCE83
   SetFieldNewData "CE84", textCE84
   SetFieldNewData "CE85", textCE85
   SetFieldNewData "CE86", textCE86
   SetFieldNewData "CE87", textCE87
   SetFieldNewData "CE88", textCE88
   SetFieldNewData "CE89", textCE89
   SetFieldNewData "CE90", textCE90
   SetFieldNewData "CE91", textCE91
   SetFieldNewData "CE92", textCE92
   SetFieldNewData "CE93", textCE93
   SetFieldNewData "CE94", textCE94
   SetFieldNewData "CE95", textCE95
   SetFieldNewData "CE96", textCE96
   SetFieldNewData "CE97", textCE97
   SetFieldNewData "CE98", textCE98
   SetFieldNewData "CE99", textCE99
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
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
   'RefreshRange
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
   textCE41_1 = Empty
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
      
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
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
   textCE41_1.Locked = bEnable
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

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textCE01.Locked = bEnable
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

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   'add by nickc 2007/01/24
   ClearField
   
   strSql = "SELECT * FROM CHANGEEVENT " & _
            "WHERE CE01 = '" & m_CurrKEY & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("CE01")) = False Then: textCE01 = rsTmp.Fields("CE01")
      
      'add by nickc 2007/01/24 從下面移上來， 更新本所案號
      UpdateCPData rsTmp.Fields("CE01")
      
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
        Select Case Me.textCP01.Text
        Case "T", "FCT", "CFT", "TF"
            If IsNull(rsTmp.Fields("CE41")) = False Then: textCE41_1 = rsTmp.Fields("CE41")
        Case Else
            If IsNull(rsTmp.Fields("CE41")) = False Then: textCE41 = rsTmp.Fields("CE41")
            If IsNull(rsTmp.Fields("CE42")) = False Then: textCE42 = rsTmp.Fields("CE42")
            If IsNull(rsTmp.Fields("CE43")) = False Then: textCE43 = rsTmp.Fields("CE43")
        End Select
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
      
      'edit by nickc 2007/01/24 移到上面
      ' 更新本所案號
      'UpdateCPData rsTmp.Fields("CE01")
            
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
      
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


' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY = strKEY01
   Else
      strSql = "SELECT CE01 FROM CHANGEEVENT " & _
               "WHERE CE01 = (SELECT MIN(CE01) FROM CHANGEEVENT " & _
                             "WHERE CE01 > '" & m_CurrKEY & "' )"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CE01")) = False Then: m_CurrKEY = rsTmp.Fields("CE01")
      Else
         rsTmp.Close
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY = m_FirstKEY
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY = m_FirstKEY Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT CE01 FROM CHANGEEVENT " & _
            "WHERE CE01 = (SELECT MAX(CE01) FROM CHANGEEVENT " & _
                          "WHERE CE01 < '" & m_CurrKEY & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CE01")) = False Then: m_CurrKEY = rsTmp.Fields("CE01")
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
   
   If m_CurrKEY = m_LastKEY Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT CE01 FROM CHANGEEVENT " & _
            "WHERE CE01 = (SELECT MIN(CE01) FROM CHANGEEVENT " & _
                          "WHERE CE01 > '" & m_CurrKEY & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CE01")) = False Then: m_CurrKEY = rsTmp.Fields("CE01")
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY = m_LastKEY
   
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
            KeyCode = 0
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
   m_SubMode = 0
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
            'add by nick 2004/10/14
            If IsCall = True Then
                If IsMod = True Then
                    'Modify by Morgan 2007/5/22
                    'frm075004_2.Show
                    frmParent.Show
                    'end 2007/5/22
                    Unload Me
                Else
                    tmpBol = fnCancelNowFormAndShowParentForm(Me)
                End If
            Else
                Unload Me
            End If
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm050706 = Nothing
End Sub

Private Sub textCE41_1_GotFocus()
    TextInverse Me.textCE41_1
End Sub

Private Sub textCE41_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCE41, 140) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE41_1_GotFocus
   End If
End Sub

Private Sub textCP01_Change()
    Select Case Me.textCP01
    Case "T", "FCT", "CFT", "TF"
        Me.Label5(3).Visible = False
        Me.textCE41.Visible = False
        Me.textCE41.Enabled = False
        Me.Label5(4).Visible = False
        Me.textCE42.Visible = False
        Me.textCE42.Enabled = False
        Me.Label5(5).Visible = False
        Me.textCE43.Visible = False
        Me.textCE43.Enabled = False
        Me.Label5(0).Visible = True
        Me.textCE41_1.Visible = True
        Me.textCE41_1.Enabled = True
    Case Else
        Me.Label5(3).Visible = True
        Me.textCE41.Visible = True
        Me.textCE41.Enabled = True
        Me.Label5(4).Visible = True
        Me.textCE42.Visible = True
        Me.textCE42.Enabled = True
        Me.Label5(5).Visible = True
        Me.textCE43.Visible = True
        Me.textCE43.Enabled = True
        Me.Label5(0).Visible = False
        Me.textCE41_1.Visible = False
        Me.textCE41_1.Enabled = False
    End Select
End Sub

Private Sub textCP01_GotFocus()
  TextInverse textCP01
End Sub

Private Sub textCP02_GotFocus()
  TextInverse textCP02
End Sub

Private Sub textCP03_GotFocus()
  TextInverse textCP03
End Sub

Private Sub textCP04_GotFocus()
  TextInverse textCP04
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
Public Function IsRecordExist(ByVal strKEY01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM CHANGEEVENT " & _
            "WHERE CE01 = '" & strKEY01 & "' "
                  
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
   Dim strCE01 As String
   
   strCE01 = textCE01
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strCE01) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   
   'Modify By Sindy 2012/5/11
   'strSql = "INSERT INTO CHANGEEVENT ("
   strSql = "begin user_data.user_enabled:=1; INSERT INTO CHANGEEVENT ("
   '2012/5/11 End
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
   strSql = strSql & "); end;"
   
   Pub_SeekTbLog strSql 'Add By Sindy 2012/10/19
   cnnConnection.Execute strSql
   
   If (strCE01 < m_FirstKEY) Or (strCE01 > m_LastKEY) Then
      RefreshRange
   End If
   
   ShowCurrRecord strCE01
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
   Dim strCE01 As String
   
   strCE01 = m_CurrKEY
   
   'Modify By Sindy 2012/5/11
   'strSql = "UPDATE CHANGEEVENT SET "
   strSql = "begin user_data.user_enabled:=1; UPDATE CHANGEEVENT SET "
   '2012/5/11 End
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = NULL "
            Else
               'Modify By Sindy 2011/1/27
               'strTmp = m_FieldList(nIndex).fiName & " = '" & m_FieldList(nIndex).fiNewData & "'"
               strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
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
                  "WHERE CE01 = '" & strCE01 & "'; end;"
   
   If bDifference = True Then
      Pub_SeekTbLog strSql 'Add By Sindy 2012/10/19
      cnnConnection.Execute strSql
      ShowCurrRecord strCE01
   End If
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strCE01 As String
   
   strCE01 = m_CurrKEY

   strSql = "DELETE FROM CHANGEEVENT " & _
            "WHERE CE01 = '" & strCE01 & "' "
   
   Pub_SeekTbLog strSql 'Add By Sindy 2012/10/19
   cnnConnection.Execute strSql

   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strCE01 = m_LastKEY) Or (strCE01 = m_FirstKEY) Then
      RefreshRange
   End If
   ShowCurrRecord strCE01
   
EXITSUB:
End Sub

' 查詢記錄
Public Function QueryRecord() As Boolean
   QueryRecord = False

   If IsRecordExist(textCE01) = True Then
      m_CurrKEY = textCE01
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
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            AddRecord
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
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
         If CheckDataValid() = True Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textCE01.SetFocus
      Case 2: textCE03.SetFocus
      Case 4: textCE01.SetFocus
   End Select
End Sub

Private Sub textCE01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 總收文號
Private Sub textCE01_Validate(Cancel As Boolean)
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE01) = False Then
      Select Case m_EditMode
         Case 1:
            If IsRecordExist(textCE01) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "該筆變更事項記錄已存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCE01_GotFocus
               GoTo EXITSUB
            End If
            
            strSql = "SELECT CP09 FROM CASEPROGRESS " & _
                     "WHERE CP09 = '" & textCE01 & "' "
            Set rsTmp = New ADODB.Recordset
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <= 0 Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "該筆收文記錄不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCE01_GotFocus
            Else
               ' 更新本所案號
               UpdateCPData textCE01
            End If
            rsTmp.Close
            Set rsTmp = Nothing
      End Select
   End If
EXITSUB:
End Sub

' 申請日
Private Sub textCE02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE02) = False Then
      If CheckIsTaiwanDate(textCE02, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE02_GotFocus
      End If
   End If
End Sub

Private Sub textCE03_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE03) = False Then
      Cancel = True
      textCE08_GotFocus
   End If
End Sub

Private Sub textCE04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "申請人1代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCE04_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub

Private Sub textCE05_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "申請人2代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCE05_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub

Private Sub textCE06_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "申請人3代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCE06_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub

Private Sub textCE07_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "申請人4代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCE07_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub

Private Sub textCE08_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "申請人5代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCE08_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub

Private Sub textCE09_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE09) = False Then
      Cancel = True
      textCE09_GotFocus
   End If
End Sub

' 代表人1(中)
Private Sub textCE10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE10) = False Then
      If StrLength(textCE10) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人1(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE10_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE10.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人1(日)
Private Sub textCE12_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE12) = False Then
      If StrLength(textCE12) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人1(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE12_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE12.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人2(中)
Private Sub textCE13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE13) = False Then
      If StrLength(textCE13) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人2(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE13_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE13.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人2(日)
Private Sub textCE15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE15) = False Then
      If StrLength(textCE15) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人2(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE15_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE15.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人3(中)
Private Sub textCE68_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE68) = False Then
      If StrLength(textCE68) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人3(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE68_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE68.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人3(日)
Private Sub textCE70_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE70) = False Then
      If StrLength(textCE70) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人3(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE70_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE70.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人4(中)
Private Sub textCE71_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE71) = False Then
      If StrLength(textCE71) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人4(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE71_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE71.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人4(日)
Private Sub textCE73_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE73) = False Then
      If StrLength(textCE73) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人4(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE73_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE73.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人5(中)
Private Sub textCE74_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE74) = False Then
      If StrLength(textCE74) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人5(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE74_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE74.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人5(日)
Private Sub textCE76_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE76) = False Then
      If StrLength(textCE76) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人5(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE76_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE76.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人6(中)
Private Sub textCE77_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE77) = False Then
      If StrLength(textCE77) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人6(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE77_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE77.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人6(日)
Private Sub textCE79_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE79) = False Then
      If StrLength(textCE79) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人6(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE79_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE79.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人7(中)
Private Sub textCE80_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE80) = False Then
      If StrLength(textCE80) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人7(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE80_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE80.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人7(日)
Private Sub textCE82_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE82) = False Then
      If StrLength(textCE82) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人7(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE82_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE82.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人8(中)
Private Sub textCE83_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE83) = False Then
      If StrLength(textCE83) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人8(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE83_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE83.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人8(日)
Private Sub textCE85_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE85) = False Then
      If StrLength(textCE85) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人8(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE85_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE85.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人9(中)
Private Sub textCE86_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE86) = False Then
      If StrLength(textCE86) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人9(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE86_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE86.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人9(日)
Private Sub textCE88_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE88) = False Then
      If StrLength(textCE88) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人9(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE88_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE88.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人10(中)
Private Sub textCE89_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE89) = False Then
      If StrLength(textCE89) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人10(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE89_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE89.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人10(日)
Private Sub textCE91_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE91) = False Then
      If StrLength(textCE91) > 40 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人10(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE91_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE91.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textCE16_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE16) = False Then
      Cancel = True
      textCE16_GotFocus
   End If
End Sub

' 申請人中譯文1
Private Sub textCE17_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE17) = False Then
      If StrLength(textCE17) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人中譯文1內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE17_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE17.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請人中譯文2
Private Sub textCE18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE18) = False Then
      If StrLength(textCE18) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人中譯文2內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE18_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE18.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請人中譯文3
Private Sub textCE19_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE19) = False Then
      If StrLength(textCE19) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人中譯文3內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE19_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE19.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請人中譯文4
Private Sub textCE20_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE20) = False Then
      If StrLength(textCE20) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人中譯文4內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE20_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE20.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請人中譯文5
Private Sub textCE21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE21) = False Then
      If StrLength(textCE21) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人中譯文5內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE21_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE21.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textCE22_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE22) = False Then
      Cancel = True
      textCE22_GotFocus
   End If
End Sub

' 申請地址(中)1
Private Sub textCE23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE23) = False Then
      If StrLength(textCE23) > 80 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(中)1內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE23_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE23.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請地址(日)1
Private Sub textCE25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE25) = False Then
      If StrLength(textCE25) > 70 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(日)1內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE25_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE25.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請地址(中)2
Private Sub textCE26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE26) = False Then
      If StrLength(textCE26) > 80 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(中)2內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE26_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE23.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請地址(日)2
Private Sub textCE28_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE28) = False Then
      If StrLength(textCE28) > 70 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(日)2內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE28_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE28.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請地址(中)3
Private Sub textCE29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE29) = False Then
      If StrLength(textCE29) > 80 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(中)3內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE29_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE29.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請地址(日)3
Private Sub textCE31_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE31) = False Then
      If StrLength(textCE31) > 70 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(日)3內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE31_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE31.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請地址(中)4
Private Sub textCE32_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE32) = False Then
      If StrLength(textCE32) > 80 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(中)4內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE32_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE32.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請地址(日)4
Private Sub textCE34_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE34) = False Then
      If StrLength(textCE34) > 70 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(日)4內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE34_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE34.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請地址(中)5
Private Sub textCE35_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE35) = False Then
      If StrLength(textCE35) > 80 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(中)5內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE35_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE35.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請地址(日)5
Private Sub textCE37_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE37) = False Then
      If StrLength(textCE37) > 70 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址(日)5內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE37_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE37.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 案件中文名稱
Private Sub textCE41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE41) = False Then
      If StrLength(textCE41) > 160 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件中文名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE41_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE41.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 案件日文名稱
Private Sub textCE43_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE43) = False Then
      If StrLength(textCE43) > 160 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件日文名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE43_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE43.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textCE38_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE38) = False Then
      Cancel = True
      textCE38_GotFocus
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
               Select Case m_EditMode
                  Case 1, 2:
                     Cancel = True
                     strTit = "檢核資料"
                     strMsg = "商標種類代碼不存在"
                     nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                     textCE39_GotFocus
                  Case Else:
               End Select
            End If
         Case "P", "CFP", "FCP":
            textCE39_2 = GetPatentName(textCE39, 0)
            If IsEmptyText(textCE39_2) = True Then
               Select Case m_EditMode
                  Case 1, 2:
                     Cancel = True
                     strTit = "檢核資料"
                     strMsg = "專利種類代碼不存在"
                     nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                     textCE39_GotFocus
                  Case Else:
               End Select
            End If
         Case Else:
      End Select
   End If
   
End Sub

Private Sub textCE40_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE40) = False Then
      Cancel = True
      textCE40_GotFocus
   End If
End Sub

Private Sub textCE44_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE44) = False Then
      Cancel = True
      textCE44_GotFocus
   End If
End Sub

Private Sub textCE46_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE46) = False Then
      Cancel = True
      textCE46_GotFocus
   End If
End Sub

Private Sub textCE48_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE48) = False Then
      Cancel = True
      textCE48_GotFocus
   End If
End Sub

Private Sub textCE50_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE50) = False Then
      Cancel = True
      textCE50_GotFocus
   End If
End Sub

Private Sub textCE52_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE52) = False Then
      Cancel = True
      textCE52_GotFocus
   End If
End Sub

Private Sub textCE54_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE54) = False Then
      Cancel = True
      textCE54_GotFocus
   End If
End Sub

Private Sub textCE56_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE56) = False Then
      Cancel = True
      textCE56_GotFocus
   End If
End Sub

Private Sub textCE58_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE58) = False Then
      Cancel = True
      textCE58_GotFocus
   End If
End Sub

Private Sub textCE60_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE60) = False Then
      Cancel = True
      textCE60_GotFocus
   End If
End Sub

Private Sub textCE62_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE62) = False Then
      Cancel = True
      textCE62_GotFocus
   End If
End Sub

' 代表人中譯文1
Private Sub textCE63_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE63) = False Then
      If StrLength(textCE63) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人中譯文1內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE63_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE63.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人中譯文2
Private Sub textCE64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE64) = False Then
      If StrLength(textCE64) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人中譯文2內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE64_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE64.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人中譯文3
Private Sub textCE92_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE92) = False Then
      If StrLength(textCE92) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人中譯文3內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE92_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE92.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人中譯文4
Private Sub textCE93_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE93) = False Then
      If StrLength(textCE93) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人中譯文4內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE93_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE93.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人中譯文5
Private Sub textCE94_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE94) = False Then
      If StrLength(textCE94) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人中譯文5內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE94_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE94.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人中譯文6
Private Sub textCE95_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE95) = False Then
      If StrLength(textCE95) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人中譯文6內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE95_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE95.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人中譯文7
Private Sub textCE96_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE96) = False Then
      If StrLength(textCE96) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人中譯文7內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE96_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE96.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人中譯文8
Private Sub textCE97_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE97) = False Then
      If StrLength(textCE97) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人中譯文8內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE97_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE97.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人中譯文9
Private Sub textCE98_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE98) = False Then
      If StrLength(textCE98) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人中譯文9內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE98_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE98.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代表人中譯文10
Private Sub textCE99_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCE99) = False Then
      If StrLength(textCE99) > 60 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代表人中譯文10內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCE99_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCE99.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textCE65_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE65) = False Then
      Cancel = True
      textCE65_GotFocus
   End If
End Sub

Private Sub textCE67_Validate(Cancel As Boolean)
   Cancel = False
   If IsText1or2(textCE67) = False Then
      Cancel = True
      textCE67_GotFocus
   End If
End Sub

' 准(1)/駁(2)
Private Function IsText1or2(ByVal strData As String) As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   IsText1or2 = True
   If IsEmptyText(strData) = False Then
      Select Case m_EditMode
         Case 1, 2:
            Select Case strData
               Case "1", "2":
               Case Else:
                  IsText1or2 = False
                  strTit = "檢核資料"
                  strMsg = "准(1)/駁(2)只可輸入1或2"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            End Select
      End Select
   End If
End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2, 4:
         ' 收文號
         If IsEmptyText(textCE01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入收文號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCE01.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
      
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCE01_GotFocus()
   InverseTextBox textCE01
End Sub

Private Sub textCE02_GotFocus()
   InverseTextBox textCE02
End Sub

Private Sub textCE03_GotFocus()
   InverseTextBox textCE03
End Sub

Private Sub textCE04_GotFocus()
   InverseTextBox textCE04
End Sub

Private Sub textCE05_GotFocus()
   InverseTextBox textCE05
End Sub

Private Sub textCE06_GotFocus()
   InverseTextBox textCE06
End Sub

Private Sub textCE07_GotFocus()
   InverseTextBox textCE07
End Sub

Private Sub textCE08_GotFocus()
   InverseTextBox textCE08
End Sub

Private Sub textCE09_GotFocus()
   InverseTextBox textCE09
End Sub

Private Sub textCE10_GotFocus()
   InverseTextBox textCE10
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE10.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE11_GotFocus()
   InverseTextBox textCE11
End Sub

Private Sub textCE12_GotFocus()
   InverseTextBox textCE12
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE12.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE13_GotFocus()
   InverseTextBox textCE13
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE13.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE14_GotFocus()
   InverseTextBox textCE14
End Sub

Private Sub textCE15_GotFocus()
   InverseTextBox textCE15
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE15.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE16_GotFocus()
   InverseTextBox textCE16
End Sub

Private Sub textCE17_GotFocus()
   InverseTextBox textCE17
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE17.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE18_GotFocus()
   InverseTextBox textCE18
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE18.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE19_GotFocus()
   InverseTextBox textCE19
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE19.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE20_GotFocus()
   InverseTextBox textCE20
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE20.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE21_GotFocus()
   InverseTextBox textCE21
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE21.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE22_GotFocus()
   InverseTextBox textCE22
End Sub

Private Sub textCE23_GotFocus()
   InverseTextBox textCE23
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE23.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE24_GotFocus()
   InverseTextBox textCE24
End Sub

Private Sub textCE25_GotFocus()
   InverseTextBox textCE25
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE25.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE26_GotFocus()
   InverseTextBox textCE26
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE26.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE27_GotFocus()
   InverseTextBox textCE27
End Sub

Private Sub textCE28_GotFocus()
   InverseTextBox textCE28
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE28.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE29_GotFocus()
   InverseTextBox textCE29
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE29.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE30_GotFocus()
   InverseTextBox textCE30
End Sub

Private Sub textCE31_GotFocus()
   InverseTextBox textCE31
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE31.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE32_GotFocus()
   InverseTextBox textCE32
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE32.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE33_GotFocus()
   InverseTextBox textCE33
End Sub

Private Sub textCE34_GotFocus()
   InverseTextBox textCE34
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE34.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE35_GotFocus()
   InverseTextBox textCE35
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE35.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE36_GotFocus()
   InverseTextBox textCE36
End Sub

Private Sub textCE37_GotFocus()
   InverseTextBox textCE37
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE37.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE38_GotFocus()
   InverseTextBox textCE38
End Sub

Private Sub textCE39_GotFocus()
   InverseTextBox textCE39
End Sub

Private Sub textCE40_GotFocus()
   InverseTextBox textCE40
End Sub

Private Sub textCE41_GotFocus()
   InverseTextBox textCE41
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE41.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE42_GotFocus()
   InverseTextBox textCE42
End Sub

Private Sub textCE43_GotFocus()
   InverseTextBox textCE43
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE43.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE44_GotFocus()
   InverseTextBox textCE44
End Sub

Private Sub textCE45_GotFocus()
   InverseTextBox textCE45
End Sub

Private Sub textCE46_GotFocus()
   InverseTextBox textCE46
End Sub

Private Sub textCE47_GotFocus()
   InverseTextBox textCE47
End Sub

Private Sub textCE48_GotFocus()
   InverseTextBox textCE48
End Sub

Private Sub textCE49_GotFocus()
   InverseTextBox textCE49
End Sub

Private Sub textCE50_GotFocus()
   InverseTextBox textCE50
End Sub

Private Sub textCE51_GotFocus()
   InverseTextBox textCE51
End Sub

Private Sub textCE52_GotFocus()
   InverseTextBox textCE52
End Sub

Private Sub textCE53_GotFocus()
   InverseTextBox textCE53
End Sub

Private Sub textCE54_GotFocus()
   InverseTextBox textCE54
End Sub

Private Sub textCE55_GotFocus()
   InverseTextBox textCE55
End Sub

Private Sub textCE56_GotFocus()
   InverseTextBox textCE56
End Sub

Private Sub textCE57_GotFocus()
   InverseTextBox textCE57
End Sub

Private Sub textCE58_GotFocus()
   InverseTextBox textCE58
End Sub

Private Sub textCE59_GotFocus()
   InverseTextBox textCE59
End Sub

Private Sub textCE60_GotFocus()
   InverseTextBox textCE60
End Sub

Private Sub textCE61_GotFocus()
   InverseTextBox textCE61
End Sub

Private Sub textCE62_GotFocus()
   InverseTextBox textCE62
End Sub

Private Sub textCE63_GotFocus()
   InverseTextBox textCE63
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE63.IMEMode = 1
    OpenIme
End Sub

Private Sub textCE64_GotFocus()
   InverseTextBox textCE64
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE64.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE65_GotFocus()
   InverseTextBox textCE65
End Sub

Private Sub textCE66_GotFocus()
   InverseTextBox textCE66
End Sub

Private Sub textCE67_GotFocus()
   InverseTextBox textCE67
End Sub

Private Sub textCE68_GotFocus()
   InverseTextBox textCE68
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE68.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE69_GotFocus()
   InverseTextBox textCE69
End Sub

Private Sub textCE70_GotFocus()
   InverseTextBox textCE70
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE70.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE71_GotFocus()
   InverseTextBox textCE71
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE71.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE72_GotFocus()
   InverseTextBox textCE72
End Sub

Private Sub textCE73_GotFocus()
   InverseTextBox textCE73
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE73.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE74_GotFocus()
   InverseTextBox textCE74
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE74.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE75_GotFocus()
   InverseTextBox textCE75
End Sub

Private Sub textCE76_GotFocus()
   InverseTextBox textCE76
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE76.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE77_GotFocus()
   InverseTextBox textCE77
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE77.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE78_GotFocus()
   InverseTextBox textCE78
End Sub

Private Sub textCE79_GotFocus()
   InverseTextBox textCE79
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE79.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE80_GotFocus()
   InverseTextBox textCE80
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE80.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE81_GotFocus()
   InverseTextBox textCE81
End Sub

Private Sub textCE82_GotFocus()
   InverseTextBox textCE82
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE82.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE83_GotFocus()
   InverseTextBox textCE83
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE83.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE84_GotFocus()
   InverseTextBox textCE84
End Sub

Private Sub textCE85_GotFocus()
   InverseTextBox textCE85
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE85.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE86_GotFocus()
   InverseTextBox textCE86
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE86.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE87_GotFocus()
   InverseTextBox textCE87
End Sub

Private Sub textCE88_GotFocus()
   InverseTextBox textCE88
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE88.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE89_GotFocus()
   InverseTextBox textCE89
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE89.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE90_GotFocus()
   InverseTextBox textCE90
End Sub

Private Sub textCE91_GotFocus()
   InverseTextBox textCE91
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE91.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE92_GotFocus()
   InverseTextBox textCE92
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE92.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE93_GotFocus()
   InverseTextBox textCE93
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE93.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE94_GotFocus()
   InverseTextBox textCE94
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE94.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE95_GotFocus()
   InverseTextBox textCE95
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE95.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE96_GotFocus()
   InverseTextBox textCE96
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE96.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE97_GotFocus()
   InverseTextBox textCE97
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE97.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE98_GotFocus()
   InverseTextBox textCE98
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE98.IMEMode = 1
   OpenIme
End Sub

Private Sub textCE99_GotFocus()
   InverseTextBox textCE99
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCE99.IMEMode = 1
   OpenIme
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCE01.Enabled = True Then
   Cancel = False
   textCE01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE02.Enabled = True Then
   Cancel = False
   textCE02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE03.Enabled = True Then
   Cancel = False
   textCE03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE04.Enabled = True Then
   Cancel = False
   textCE04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE05.Enabled = True Then
   Cancel = False
   textCE05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE06.Enabled = True Then
   Cancel = False
   textCE06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE07.Enabled = True Then
   Cancel = False
   textCE07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE08.Enabled = True Then
   Cancel = False
   textCE08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE09.Enabled = True Then
   Cancel = False
   textCE09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE10.Enabled = True Then
   Cancel = False
   textCE10_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE12.Enabled = True Then
   Cancel = False
   textCE12_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE13.Enabled = True Then
   Cancel = False
   textCE13_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE15.Enabled = True Then
   Cancel = False
   textCE15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE16.Enabled = True Then
   Cancel = False
   textCE16_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE17.Enabled = True Then
   Cancel = False
   textCE17_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE18.Enabled = True Then
   Cancel = False
   textCE18_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE19.Enabled = True Then
   Cancel = False
   textCE19_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE20.Enabled = True Then
   Cancel = False
   textCE20_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE21.Enabled = True Then
   Cancel = False
   textCE20_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE22.Enabled = True Then
   Cancel = False
   textCE22_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE23.Enabled = True Then
   Cancel = False
   textCE23_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE25.Enabled = True Then
   Cancel = False
   textCE25_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE26.Enabled = True Then
   Cancel = False
   textCE26_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE28.Enabled = True Then
   Cancel = False
   textCE28_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE29.Enabled = True Then
   Cancel = False
   textCE29_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE31.Enabled = True Then
   Cancel = False
   textCE31_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE32.Enabled = True Then
   Cancel = False
   textCE32_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE34.Enabled = True Then
   Cancel = False
   textCE34_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE35.Enabled = True Then
   Cancel = False
   textCE35_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE37.Enabled = True Then
   Cancel = False
   textCE37_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE38.Enabled = True Then
   Cancel = False
   textCE38_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE39.Enabled = True Then
   Cancel = False
   textCE39_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE40.Enabled = True Then
   Cancel = False
   textCE40_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE41.Enabled = True Then
   Cancel = False
   textCE41_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textCE41_1.Enabled = True Then
   Cancel = False
   textCE41_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE43.Enabled = True Then
   Cancel = False
   textCE43_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE44.Enabled = True Then
   Cancel = False
   textCE44_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE46.Enabled = True Then
   Cancel = False
   textCE46_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE48.Enabled = True Then
   Cancel = False
   textCE48_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE50.Enabled = True Then
   Cancel = False
   textCE50_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE52.Enabled = True Then
   Cancel = False
   textCE52_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE54.Enabled = True Then
   Cancel = False
   textCE54_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE56.Enabled = True Then
   Cancel = False
   textCE56_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE58.Enabled = True Then
   Cancel = False
   textCE58_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE60.Enabled = True Then
   Cancel = False
   textCE60_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE62.Enabled = True Then
   Cancel = False
   textCE62_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE63.Enabled = True Then
   Cancel = False
   textCE63_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE64.Enabled = True Then
   Cancel = False
   textCE64_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE65.Enabled = True Then
   Cancel = False
   textCE65_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE67.Enabled = True Then
   Cancel = False
   textCE67_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE68.Enabled = True Then
   Cancel = False
   textCE68_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE70.Enabled = True Then
   Cancel = False
   textCE70_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE71.Enabled = True Then
   Cancel = False
   textCE71_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE73.Enabled = True Then
   Cancel = False
   textCE73_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE74.Enabled = True Then
   Cancel = False
   textCE74_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE76.Enabled = True Then
   Cancel = False
   textCE76_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE77.Enabled = True Then
   Cancel = False
   textCE77_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE79.Enabled = True Then
   Cancel = False
   textCE79_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE80.Enabled = True Then
   Cancel = False
   textCE80_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE82.Enabled = True Then
   Cancel = False
   textCE82_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE83.Enabled = True Then
   Cancel = False
   textCE83_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE85.Enabled = True Then
   Cancel = False
   textCE85_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE86.Enabled = True Then
   Cancel = False
   textCE86_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE88.Enabled = True Then
   Cancel = False
   textCE88_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE89.Enabled = True Then
   Cancel = False
   textCE89_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE91.Enabled = True Then
   Cancel = False
   textCE91_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE92.Enabled = True Then
   Cancel = False
   textCE92_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE93.Enabled = True Then
   Cancel = False
   textCE93_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE94.Enabled = True Then
   Cancel = False
   textCE94_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE95.Enabled = True Then
   Cancel = False
   textCE95_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE96.Enabled = True Then
   Cancel = False
   textCE96_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE97.Enabled = True Then
   Cancel = False
   textCE97_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE98.Enabled = True Then
   Cancel = False
   textCE98_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE99.Enabled = True Then
   Cancel = False
   textCE99_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Lydia 2021/10/12 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

TxtValidate = True
End Function

