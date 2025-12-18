VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010301_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-電子送件-新案"
   ClientHeight    =   6640
   ClientLeft      =   410
   ClientTop       =   1500
   ClientWidth     =   8030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6640
   ScaleWidth      =   8030
   Begin TabDlg.SSTab SSTab1 
      Height          =   6045
      Left            =   45
      TabIndex        =   76
      Top             =   540
      Width           =   7905
      _ExtentX        =   13952
      _ExtentY        =   10672
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      ForeColor       =   -2147483630
      TabCaption(0)   =   "案件名稱"
      TabPicture(0)   =   "frm06010301_1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5(1)"
      Tab(0).Control(1)=   "Label14(1)"
      Tab(0).Control(2)=   "Label5(0)"
      Tab(0).Control(3)=   "lblNameAgent"
      Tab(0).Control(4)=   "lblFavDate"
      Tab(0).Control(5)=   "lblFavReason"
      Tab(0).Control(6)=   "Label27"
      Tab(0).Control(7)=   "Label30"
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(9)=   "Label31"
      Tab(0).Control(10)=   "Label24"
      Tab(0).Control(11)=   "Label25"
      Tab(0).Control(12)=   "Label26"
      Tab(0).Control(13)=   "Label32"
      Tab(0).Control(14)=   "lstNameAgent"
      Tab(0).Control(15)=   "Text6(0)"
      Tab(0).Control(16)=   "Text6(1)"
      Tab(0).Control(17)=   "txtFavDate"
      Tab(0).Control(18)=   "cboFavReason"
      Tab(0).Control(19)=   "txtCP84"
      Tab(0).Control(20)=   "Text7"
      Tab(0).Control(21)=   "Frame2"
      Tab(0).Control(22)=   "txtTF24"
      Tab(0).Control(23)=   "txtTF25"
      Tab(0).Control(24)=   "txtForeign"
      Tab(0).Control(25)=   "chkDoc(1)"
      Tab(0).Control(26)=   "cboLagnuage"
      Tab(0).Control(27)=   "txtSimplified"
      Tab(0).Control(28)=   "chkDoc(2)"
      Tab(0).Control(29)=   "cmdOpen(1)"
      Tab(0).Control(30)=   "cmdOpen(0)"
      Tab(0).Control(31)=   "chkEexcerpt"
      Tab(0).Control(32)=   "FraPA174"
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "說明書及附件"
      TabPicture(1)   =   "frm06010301_1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label28"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Shape2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label15"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label17"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label19"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label20"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label21"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label22"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label6"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label33"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtMemo"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "chkAtt(10)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "chkAtt(9)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkAtt(8)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "chkAtt(7)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "chkAtt(6)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "chkAtt(5)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "chkAtt(4)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "chkAtt(3)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "chkAtt(2)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "chkAtt(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "chkAtt(0)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "chkDoc(0)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtDocCh(0)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtDocCh(1)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "txtDocCh(2)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "txtDocCh(3)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txtDocCh(4)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "txtDocCh(5)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txtDocCh(6)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "chkAtt(15)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "chkAtt(13)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "chkAtt(14)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "chkAtt(12)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Frame1"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "chkAtt(11)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "chkAtt(16)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "txtDocCh(7)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Check3"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "chkAtt(25)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Frame3"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).ControlCount=   45
      TabCaption(2)   =   "地址"
      TabPicture(2)   =   "frm06010301_1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label23(0)"
      Tab(2).Control(1)=   "Label5(9)"
      Tab(2).Control(2)=   "Label5(10)"
      Tab(2).Control(3)=   "Label5(11)"
      Tab(2).Control(4)=   "Label5(12)"
      Tab(2).Control(5)=   "Label5(13)"
      Tab(2).Control(6)=   "Label5(14)"
      Tab(2).Control(7)=   "Label5(15)"
      Tab(2).Control(8)=   "Label5(16)"
      Tab(2).Control(9)=   "Label5(17)"
      Tab(2).Control(10)=   "Label5(18)"
      Tab(2).Control(11)=   "Label5(19)"
      Tab(2).Control(12)=   "Label5(20)"
      Tab(2).Control(13)=   "Label5(21)"
      Tab(2).Control(14)=   "Label5(22)"
      Tab(2).Control(15)=   "Label5(23)"
      Tab(2).Control(16)=   "Label23(1)"
      Tab(2).Control(17)=   "Label23(2)"
      Tab(2).Control(18)=   "Label23(3)"
      Tab(2).Control(19)=   "Label23(4)"
      Tab(2).Control(20)=   "Text6(9)"
      Tab(2).Control(21)=   "Text6(10)"
      Tab(2).Control(22)=   "Text6(11)"
      Tab(2).Control(23)=   "Text6(12)"
      Tab(2).Control(24)=   "Text6(13)"
      Tab(2).Control(25)=   "Text6(14)"
      Tab(2).Control(26)=   "Text6(15)"
      Tab(2).Control(27)=   "Text6(16)"
      Tab(2).Control(28)=   "Text6(17)"
      Tab(2).Control(29)=   "Text6(18)"
      Tab(2).Control(30)=   "Text6(19)"
      Tab(2).Control(31)=   "Text6(20)"
      Tab(2).Control(32)=   "Text6(21)"
      Tab(2).Control(33)=   "Text6(22)"
      Tab(2).Control(34)=   "Text6(23)"
      Tab(2).ControlCount=   35
      TabCaption(3)   =   "發明人"
      TabPicture(3)   =   "frm06010301_1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label46(0)"
      Tab(3).Control(1)=   "GRD1"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "代表人 1"
      TabPicture(4)   =   "frm06010301_1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label14(3)"
      Tab(4).Control(1)=   "Label5(35)"
      Tab(4).Control(2)=   "Label5(34)"
      Tab(4).Control(3)=   "Label5(33)"
      Tab(4).Control(4)=   "Label18(1)"
      Tab(4).Control(5)=   "Label14(2)"
      Tab(4).Control(6)=   "Label5(29)"
      Tab(4).Control(7)=   "Label5(28)"
      Tab(4).Control(8)=   "Label5(27)"
      Tab(4).Control(9)=   "Label5(26)"
      Tab(4).Control(10)=   "Label5(25)"
      Tab(4).Control(11)=   "Label5(24)"
      Tab(4).Control(12)=   "Label18(0)"
      Tab(4).Control(13)=   "Label14(0)"
      Tab(4).Control(14)=   "Label5(3)"
      Tab(4).Control(15)=   "Label5(4)"
      Tab(4).Control(16)=   "Label5(5)"
      Tab(4).Control(17)=   "Label5(6)"
      Tab(4).Control(18)=   "Label5(7)"
      Tab(4).Control(19)=   "Label5(8)"
      Tab(4).Control(20)=   "Text6(3)"
      Tab(4).Control(21)=   "Text6(4)"
      Tab(4).Control(22)=   "Text6(5)"
      Tab(4).Control(23)=   "Text6(6)"
      Tab(4).Control(24)=   "Text6(7)"
      Tab(4).Control(25)=   "Text6(8)"
      Tab(4).Control(26)=   "Text6(24)"
      Tab(4).Control(27)=   "Text6(25)"
      Tab(4).Control(28)=   "Text6(26)"
      Tab(4).Control(29)=   "Text6(27)"
      Tab(4).Control(30)=   "Text6(28)"
      Tab(4).Control(31)=   "Text6(29)"
      Tab(4).Control(32)=   "Text6(30)"
      Tab(4).Control(33)=   "Text6(31)"
      Tab(4).Control(34)=   "Text6(32)"
      Tab(4).Control(35)=   "Combo2(0)"
      Tab(4).Control(36)=   "Combo2(1)"
      Tab(4).Control(37)=   "Combo2(2)"
      Tab(4).Control(38)=   "Combo2(3)"
      Tab(4).Control(39)=   "Combo2(4)"
      Tab(4).ControlCount=   40
      TabCaption(5)   =   "代表人2"
      TabPicture(5)   =   "frm06010301_1.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label18(6)"
      Tab(5).Control(1)=   "Label5(47)"
      Tab(5).Control(2)=   "Label5(46)"
      Tab(5).Control(3)=   "Label5(45)"
      Tab(5).Control(4)=   "Label18(5)"
      Tab(5).Control(5)=   "Label5(44)"
      Tab(5).Control(6)=   "Label5(43)"
      Tab(5).Control(7)=   "Label5(42)"
      Tab(5).Control(8)=   "Label18(4)"
      Tab(5).Control(9)=   "Label5(41)"
      Tab(5).Control(10)=   "Label5(40)"
      Tab(5).Control(11)=   "Label5(39)"
      Tab(5).Control(12)=   "Label18(3)"
      Tab(5).Control(13)=   "Label5(38)"
      Tab(5).Control(14)=   "Label5(37)"
      Tab(5).Control(15)=   "Label5(36)"
      Tab(5).Control(16)=   "Label18(2)"
      Tab(5).Control(17)=   "Label5(32)"
      Tab(5).Control(18)=   "Label5(31)"
      Tab(5).Control(19)=   "Label5(30)"
      Tab(5).Control(20)=   "Text6(33)"
      Tab(5).Control(21)=   "Text6(34)"
      Tab(5).Control(22)=   "Text6(35)"
      Tab(5).Control(23)=   "Text6(36)"
      Tab(5).Control(24)=   "Text6(37)"
      Tab(5).Control(25)=   "Text6(38)"
      Tab(5).Control(26)=   "Text6(39)"
      Tab(5).Control(27)=   "Text6(40)"
      Tab(5).Control(28)=   "Text6(41)"
      Tab(5).Control(29)=   "Text6(42)"
      Tab(5).Control(30)=   "Text6(43)"
      Tab(5).Control(31)=   "Text6(44)"
      Tab(5).Control(32)=   "Text6(45)"
      Tab(5).Control(33)=   "Text6(46)"
      Tab(5).Control(34)=   "Text6(47)"
      Tab(5).Control(35)=   "Combo2(5)"
      Tab(5).Control(36)=   "Combo2(6)"
      Tab(5).Control(37)=   "Combo2(7)"
      Tab(5).Control(38)=   "Combo2(8)"
      Tab(5).Control(39)=   "Combo2(9)"
      Tab(5).ControlCount=   40
      TabCaption(6)   =   "優先權資料"
      TabPicture(6)   =   "frm06010301_1.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "grdDataList2"
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   1440
         Left            =   900
         TabIndex        =   225
         Top             =   4410
         Width           =   7536
         Begin VB.CheckBox chk2 
            Caption         =   "援用原申請案優先權主張"
            Height          =   315
            Left            =   108
            TabIndex        =   227
            Top             =   450
            Value           =   1  '核取
            Width           =   2520
         End
         Begin VB.CheckBox chk1 
            Caption         =   "本案符合優惠期相關規定"
            Height          =   315
            Left            =   108
            TabIndex        =   226
            Top             =   150
            Width           =   2490
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "                    "
            Height          =   180
            Left            =   3630
            TabIndex        =   228
            Top             =   195
            Width           =   1590
         End
      End
      Begin VB.Frame FraPA174 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame1"
         Height          =   525
         Left            =   -74820
         TabIndex        =   219
         Top             =   660
         Visible         =   0   'False
         Width           =   825
         Begin VB.CommandButton CmdPA174 
            BackColor       =   &H00C0FFFF&
            Caption         =   "特殊字"
            Height          =   280
            Left            =   0
            Style           =   1  '圖片外觀
            TabIndex        =   220
            Top             =   210
            Width           =   800
         End
         Begin VB.Label lblPA174 
            Caption         =   "有特殊字"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   35
            TabIndex        =   221
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.CheckBox chkEexcerpt 
         Caption         =   "未附英文說明書，但可減收申請規費800元"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72270
         TabIndex        =   6
         Top             =   1890
         Width           =   3690
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "委任書(附譯文)"
         Height          =   195
         Index           =   25
         Left            =   5550
         TabIndex        =   40
         Top             =   2016
         Width           =   1530
      End
      Begin VB.CheckBox Check3 
         Caption         =   "個案"
         Height          =   195
         Left            =   7080
         TabIndex        =   41
         Top             =   2016
         Width           =   690
      End
      Begin VB.TextBox txtDocCh 
         Enabled         =   0   'False
         Height          =   270
         Index           =   7
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   24
         Top             =   966
         Width           =   420
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "外文本"
         Height          =   330
         Index           =   0
         Left            =   -72270
         TabIndex        =   216
         Top             =   3540
         Width           =   800
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "電子送件暫存區"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   -71310
         TabIndex        =   215
         Top             =   3540
         Width           =   1395
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "簡體字本資訊"
         Height          =   195
         Index           =   2
         Left            =   -74790
         TabIndex        =   15
         Top             =   4905
         Width           =   1455
      End
      Begin VB.TextBox txtSimplified 
         Enabled         =   0   'False
         Height          =   270
         Left            =   -71820
         TabIndex        =   16
         Top             =   4860
         Width           =   420
      End
      Begin VB.ComboBox cboLagnuage 
         Enabled         =   0   'False
         Height          =   276
         ItemData        =   "frm06010301_1.frx":00C4
         Left            =   -73350
         List            =   "frm06010301_1.frx":00E3
         Style           =   2  '單純下拉式
         TabIndex        =   14
         Top             =   4500
         Width           =   1770
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "外文本資訊"
         Height          =   195
         Index           =   1
         Left            =   -74790
         TabIndex        =   10
         Top             =   3300
         Width           =   1230
      End
      Begin VB.TextBox txtForeign 
         Enabled         =   0   'False
         Height          =   270
         Left            =   -73350
         TabIndex        =   13
         Top             =   4185
         Width           =   420
      End
      Begin VB.TextBox txtTF25 
         Height          =   280
         Left            =   -73350
         TabIndex        =   12
         Top             =   3870
         Width           =   420
      End
      Begin VB.TextBox txtTF24 
         Height          =   280
         Left            =   -73350
         TabIndex        =   11
         Top             =   3540
         Width           =   420
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  '沒有框線
         Height          =   795
         Left            =   -69000
         TabIndex        =   206
         Top             =   2190
         Width           =   1545
         Begin VB.TextBox txtCP135 
            Height          =   280
            Left            =   810
            TabIndex        =   8
            Top             =   30
            Width           =   420
         End
         Begin VB.TextBox txtCP136 
            Height          =   280
            Left            =   810
            TabIndex        =   9
            Top             =   330
            Width           =   420
         End
         Begin VB.Label lblCP135 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "總頁數:"
            Height          =   180
            Left            =   180
            TabIndex        =   208
            Top             =   60
            Width           =   585
         End
         Begin VB.Label lblCP136 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "總項數:"
            Height          =   180
            Left            =   180
            TabIndex        =   207
            Top             =   330
            Width           =   585
         End
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "其他"
         Height          =   195
         Index           =   16
         Left            =   6480
         TabIndex        =   48
         Top             =   3528
         Width           =   828
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "一案兩請聲明"
         Height          =   195
         Index           =   11
         Left            =   4680
         TabIndex        =   47
         Top             =   3528
         Width           =   1404
      End
      Begin VB.Frame Frame1 
         Height          =   1440
         Left            =   108
         TabIndex        =   198
         Top             =   3852
         Width           =   7536
         Begin VB.CheckBox chkBiology 
            Caption         =   "生物材料不須寄存   所屬技術領域中具有通常知識者易於獲得。"
            Height          =   315
            Index           =   1
            Left            =   108
            TabIndex        =   54
            Top             =   1044
            Width           =   5400
         End
         Begin VB.TextBox txtBiology 
            Enabled         =   0   'False
            Height          =   270
            Index           =   3
            Left            =   5328
            TabIndex        =   53
            Top             =   756
            Width           =   1995
         End
         Begin VB.TextBox txtBiology 
            Enabled         =   0   'False
            Height          =   270
            Index           =   2
            Left            =   2808
            MaxLength       =   8
            TabIndex        =   52
            Top             =   756
            Width           =   1275
         End
         Begin VB.TextBox txtBiology 
            Enabled         =   0   'False
            Height          =   270
            Index           =   1
            Left            =   2808
            TabIndex        =   51
            Top             =   459
            Width           =   4515
         End
         Begin VB.TextBox txtBiology 
            Enabled         =   0   'False
            Height          =   270
            Index           =   0
            Left            =   2808
            MaxLength       =   3
            TabIndex        =   50
            Top             =   156
            Width           =   735
         End
         Begin VB.CheckBox chkBiology 
            Caption         =   "主張利用生物材料"
            Height          =   315
            Index           =   0
            Left            =   108
            TabIndex        =   49
            Top             =   144
            Width           =   1815
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "寄存號碼:"
            Height          =   180
            Left            =   4524
            TabIndex        =   203
            Top             =   804
            Width           =   768
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "寄存日期:"
            Height          =   180
            Left            =   1956
            TabIndex        =   202
            Top             =   804
            Width           =   768
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "寄存機構:"
            Height          =   180
            Left            =   1956
            TabIndex        =   201
            Top             =   504
            Width           =   768
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "寄存國家:"
            Height          =   180
            Left            =   1956
            TabIndex        =   200
            Top             =   204
            Width           =   768
         End
         Begin VB.Label lblCountry 
            AutoSize        =   -1  'True
            Caption         =   "                    "
            Height          =   180
            Left            =   3630
            TabIndex        =   199
            Top             =   195
            Width           =   1590
         End
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "國際優先權證明文件"
         Height          =   195
         Index           =   12
         Left            =   4680
         TabIndex        =   42
         Top             =   2232
         Width           =   2445
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "序列表"
         Height          =   195
         Index           =   14
         Left            =   6615
         TabIndex        =   37
         Top             =   1596
         Width           =   915
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "圖式"
         Height          =   195
         Index           =   13
         Left            =   5760
         TabIndex        =   36
         Top             =   1596
         Width           =   780
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "優惠期證明文件"
         Height          =   195
         Index           =   15
         Left            =   4680
         TabIndex        =   43
         Tag             =   "ExperimentNovelty.pdf"
         Top             =   2436
         Width           =   2445
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H80000011&
         Height          =   276
         Index           =   9
         Left            =   -73800
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   152
         Top             =   4836
         Width           =   6225
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H80000011&
         Height          =   276
         Index           =   8
         Left            =   -73800
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   148
         Top             =   3717
         Width           =   6225
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H80000011&
         Height          =   276
         Index           =   7
         Left            =   -73800
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   144
         Top             =   2598
         Width           =   6225
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H80000011&
         Height          =   276
         Index           =   6
         Left            =   -73800
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   140
         Top             =   1470
         Width           =   6225
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H80000011&
         Height          =   276
         Index           =   5
         Left            =   -73800
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   136
         Top             =   360
         Width           =   6225
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H80000011&
         Height          =   276
         Index           =   4
         Left            =   -73800
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   132
         Top             =   4830
         Width           =   6225
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H80000011&
         Height          =   276
         Index           =   3
         Left            =   -73800
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   128
         Top             =   3720
         Width           =   6225
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H80000011&
         Height          =   276
         Index           =   2
         Left            =   -73800
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   124
         Top             =   2580
         Width           =   6225
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H80000011&
         Height          =   276
         Index           =   1
         Left            =   -73800
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   120
         Top             =   1470
         Width           =   6225
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H80000011&
         Height          =   276
         Index           =   0
         Left            =   -73800
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   116
         Top             =   360
         Width           =   6225
      End
      Begin VB.TextBox txtDocCh 
         Enabled         =   0   'False
         Height          =   270
         Index           =   6
         Left            =   3240
         TabIndex        =   29
         Top             =   2335
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Enabled         =   0   'False
         Height          =   270
         Index           =   5
         Left            =   3240
         TabIndex        =   28
         Top             =   2058
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Height          =   270
         Index           =   4
         Left            =   3240
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   27
         Top             =   1785
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   26
         Top             =   1512
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   25
         Top             =   1239
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   23
         Top             =   693
         Width           =   420
      End
      Begin VB.TextBox txtDocCh 
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   22
         Top             =   420
         Width           =   420
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "中文本資訊"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   21
         Top             =   465
         Width           =   1230
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "基本資料表"
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   106
         Top             =   468
         Value           =   2  '灰色
         Width           =   1230
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "發明摘要"
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   31
         Top             =   708
         Width           =   1230
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "發明說明書"
         Height          =   195
         Index           =   2
         Left            =   4680
         TabIndex        =   32
         Top             =   936
         Width           =   1230
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "發明申請專利範圍"
         Height          =   195
         Index           =   3
         Left            =   4680
         TabIndex        =   33
         Top             =   1152
         Width           =   1860
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "發明圖式"
         Height          =   195
         Index           =   4
         Left            =   4680
         TabIndex        =   34
         Top             =   1368
         Width           =   1230
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "說明書"
         Height          =   195
         Index           =   5
         Left            =   4680
         TabIndex        =   35
         Top             =   1596
         Width           =   1005
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "簡體字本"
         Height          =   195
         Index           =   6
         Left            =   4680
         TabIndex        =   38
         Top             =   1812
         Width           =   1230
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "委任書"
         Height          =   195
         Index           =   7
         Left            =   4680
         TabIndex        =   39
         Top             =   2016
         Width           =   1020
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "國內生物材料寄存證明文件"
         Height          =   195
         Index           =   8
         Left            =   4680
         TabIndex        =   44
         Top             =   2676
         Width           =   2535
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "國外生物材料寄存證明文件"
         Height          =   195
         Index           =   9
         Left            =   4680
         TabIndex        =   45
         Top             =   2916
         Width           =   2535
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "生物材料為通常知識者易於獲得證明文件"
         Height          =   405
         Index           =   10
         Left            =   4680
         TabIndex        =   46
         Top             =   3132
         Width           =   2445
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  '置中對齊
         Height          =   285
         Left            =   -73380
         TabIndex        =   5
         Top             =   1845
         Width           =   285
      End
      Begin VB.TextBox txtCP84 
         Height          =   270
         Left            =   -73860
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1530
         Width           =   1095
      End
      Begin VB.ComboBox cboFavReason 
         Height          =   276
         ItemData        =   "frm06010301_1.frx":0129
         Left            =   -71700
         List            =   "frm06010301_1.frx":0136
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   1170
         Width           =   4155
      End
      Begin VB.TextBox txtFavDate 
         Height          =   270
         Left            =   -73545
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   5055
         Left            =   -74910
         TabIndex        =   115
         Top             =   630
         Width           =   7725
         _ExtentX        =   13635
         _ExtentY        =   8908
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "V|發明人編號|中文名稱|英文名稱|日文名稱"
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
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList2 
         Height          =   5235
         Left            =   -74340
         TabIndex        =   196
         Top             =   480
         Width           =   6495
         _ExtentX        =   11448
         _ExtentY        =   9225
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         HighLight       =   0
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
         _Band(0).Cols   =   3
      End
      Begin MSForms.TextBox txtMemo 
         Height          =   1080
         Left            =   750
         TabIndex        =   30
         Top             =   2640
         Width           =   2925
         VariousPropertyBits=   -1466939365
         ScrollBars      =   2
         Size            =   "5159;1905"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   47
         Left            =   -73800
         TabIndex        =   155
         Top             =   5670
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   40
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   46
         Left            =   -73800
         TabIndex        =   154
         Top             =   5403
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   45
         Left            =   -73800
         TabIndex        =   153
         Top             =   5127
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   50
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   44
         Left            =   -73800
         TabIndex        =   151
         Top             =   4560
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   40
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   43
         Left            =   -73800
         TabIndex        =   150
         Top             =   4284
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   42
         Left            =   -73800
         TabIndex        =   149
         Top             =   4008
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   50
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   41
         Left            =   -73800
         TabIndex        =   147
         Top             =   3441
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   40
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   40
         Left            =   -73800
         TabIndex        =   146
         Top             =   3165
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   39
         Left            =   -73800
         TabIndex        =   145
         Top             =   2889
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   50
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   38
         Left            =   -73800
         TabIndex        =   143
         Top             =   2322
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   40
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   37
         Left            =   -73800
         TabIndex        =   142
         Top             =   2046
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   36
         Left            =   -73800
         TabIndex        =   141
         Top             =   1770
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   50
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   35
         Left            =   -73800
         TabIndex        =   139
         Top             =   1200
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   40
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   34
         Left            =   -73800
         TabIndex        =   138
         Top             =   927
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   33
         Left            =   -73800
         TabIndex        =   137
         Top             =   651
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   50
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   32
         Left            =   -73800
         TabIndex        =   135
         Top             =   5670
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   40
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   31
         Left            =   -73800
         TabIndex        =   134
         Top             =   5400
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   30
         Left            =   -73800
         TabIndex        =   133
         Top             =   5130
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   50
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   29
         Left            =   -73800
         TabIndex        =   130
         Top             =   4560
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   40
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   28
         Left            =   -73800
         TabIndex        =   131
         Top             =   4290
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   27
         Left            =   -73800
         TabIndex        =   129
         Top             =   4020
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   50
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   26
         Left            =   -73800
         TabIndex        =   127
         Top             =   3420
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   40
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   25
         Left            =   -73800
         TabIndex        =   126
         Top             =   3150
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   24
         Left            =   -73800
         TabIndex        =   125
         Top             =   2880
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   50
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   23
         Left            =   -73620
         TabIndex        =   69
         Top             =   4950
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   22
         Left            =   -73620
         TabIndex        =   66
         Top             =   3960
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   21
         Left            =   -73620
         TabIndex        =   63
         Top             =   2940
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   20
         Left            =   -73620
         TabIndex        =   60
         Top             =   1920
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   19
         Left            =   -73620
         TabIndex        =   57
         Top             =   870
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   18
         Left            =   -73620
         TabIndex        =   68
         Top             =   4680
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   154
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   17
         Left            =   -73620
         TabIndex        =   65
         Top             =   3690
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   154
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   16
         Left            =   -73620
         TabIndex        =   62
         Top             =   2670
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   154
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   15
         Left            =   -73620
         TabIndex        =   59
         Top             =   1650
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   154
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   14
         Left            =   -73620
         TabIndex        =   56
         Top             =   600
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   154
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   13
         Left            =   -73620
         TabIndex        =   67
         Top             =   4410
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   12
         Left            =   -73620
         TabIndex        =   64
         Top             =   3420
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   11
         Left            =   -73620
         TabIndex        =   61
         Top             =   2400
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   10
         Left            =   -73620
         TabIndex        =   58
         Top             =   1380
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   9
         Left            =   -73620
         TabIndex        =   55
         Top             =   330
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   8
         Left            =   -73800
         TabIndex        =   123
         Top             =   2310
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   40
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   7
         Left            =   -73800
         TabIndex        =   122
         Top             =   2040
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   6
         Left            =   -73800
         TabIndex        =   121
         Top             =   1770
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   50
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   5
         Left            =   -73800
         TabIndex        =   119
         Top             =   1200
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   40
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   4
         Left            =   -73800
         TabIndex        =   118
         Top             =   930
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   80
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   3
         Left            =   -73800
         TabIndex        =   117
         Top             =   660
         Width           =   6225
         VariousPropertyBits=   679495707
         MaxLength       =   50
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   1
         Left            =   -73560
         TabIndex        =   1
         Top             =   690
         Width           =   6225
         VariousPropertyBits=   679495707
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text6 
         Height          =   285
         Index           =   0
         Left            =   -73560
         TabIndex        =   0
         Top             =   390
         Width           =   6225
         VariousPropertyBits=   679495707
         Size            =   "10980;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   315
         Left            =   -73800
         TabIndex        =   7
         Top             =   2160
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;556"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "不算超頁費"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   2320
         TabIndex        =   218
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "序列表:"
         Height          =   180
         Left            =   1710
         TabIndex        =   217
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label Label32 
         Caption         =   "(P.S 外文本說明書和圖示頁數，只有當收文新案翻譯、檢視中說或核對中說格式才會回寫到新案建檔的翻譯頁籤)"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   -72270
         TabIndex        =   214
         Top             =   3900
         Width           =   4095
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "簡體字頁數總計:"
         Height          =   180
         Left            =   -73350
         TabIndex        =   213
         Top             =   4905
         Width           =   1305
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "外文本種類:"
         Height          =   180
         Left            =   -74490
         TabIndex        =   212
         Top             =   4500
         Width           =   945
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "外文頁數總計:"
         Height          =   180
         Left            =   -74520
         TabIndex        =   211
         Top             =   4230
         Width           =   1125
      End
      Begin VB.Label Label31 
         Caption         =   "外文本圖示:　              頁"
         Height          =   255
         Left            =   -74520
         TabIndex        =   210
         Top             =   3870
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "外文本說明書:              頁"
         Height          =   255
         Left            =   -74520
         TabIndex        =   209
         Top             =   3555
         Width           =   2055
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "備註:"
         Height          =   180
         Left            =   285
         TabIndex        =   205
         Top             =   2730
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "外文本"
         Height          =   180
         Left            =   4092
         TabIndex        =   197
         Top             =   1596
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   30
         Left            =   -74280
         TabIndex        =   195
         Top             =   1110
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   31
         Left            =   -74280
         TabIndex        =   194
         Top             =   870
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   32
         Left            =   -74280
         TabIndex        =   193
         Top             =   630
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人6"
         Height          =   180
         Index           =   2
         Left            =   -74640
         TabIndex        =   192
         Top             =   420
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   36
         Left            =   -74280
         TabIndex        =   191
         Top             =   2250
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   37
         Left            =   -74280
         TabIndex        =   190
         Top             =   2010
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   38
         Left            =   -74280
         TabIndex        =   189
         Top             =   1770
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人7"
         Height          =   180
         Index           =   3
         Left            =   -74640
         TabIndex        =   188
         Top             =   1530
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   39
         Left            =   -74280
         TabIndex        =   187
         Top             =   3390
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   40
         Left            =   -74280
         TabIndex        =   186
         Top             =   3150
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   41
         Left            =   -74280
         TabIndex        =   185
         Top             =   2910
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人8"
         Height          =   180
         Index           =   4
         Left            =   -74640
         TabIndex        =   184
         Top             =   2670
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   42
         Left            =   -74280
         TabIndex        =   183
         Top             =   4500
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   43
         Left            =   -74280
         TabIndex        =   182
         Top             =   4260
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   44
         Left            =   -74280
         TabIndex        =   181
         Top             =   4020
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人9"
         Height          =   180
         Index           =   5
         Left            =   -74640
         TabIndex        =   180
         Top             =   3780
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   45
         Left            =   -74280
         TabIndex        =   179
         Top             =   5640
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   46
         Left            =   -74280
         TabIndex        =   178
         Top             =   5400
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   47
         Left            =   -74280
         TabIndex        =   177
         Top             =   5160
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人10"
         Height          =   180
         Index           =   6
         Left            =   -74640
         TabIndex        =   176
         Top             =   4920
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   8
         Left            =   -74280
         TabIndex        =   175
         Top             =   2250
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   7
         Left            =   -74280
         TabIndex        =   174
         Top             =   2010
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   6
         Left            =   -74280
         TabIndex        =   173
         Top             =   1770
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   5
         Left            =   -74280
         TabIndex        =   172
         Top             =   1140
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   4
         Left            =   -74280
         TabIndex        =   171
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   3
         Left            =   -74280
         TabIndex        =   170
         Top             =   660
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人1"
         Height          =   180
         Index           =   0
         Left            =   -74640
         TabIndex        =   169
         Top             =   420
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人2"
         Height          =   180
         Index           =   0
         Left            =   -74640
         TabIndex        =   168
         Top             =   1530
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   24
         Left            =   -74280
         TabIndex        =   167
         Top             =   4500
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   25
         Left            =   -74280
         TabIndex        =   166
         Top             =   4260
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   26
         Left            =   -74280
         TabIndex        =   165
         Top             =   4020
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   27
         Left            =   -74280
         TabIndex        =   164
         Top             =   3360
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   28
         Left            =   -74280
         TabIndex        =   163
         Top             =   3120
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   29
         Left            =   -74280
         TabIndex        =   162
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人3"
         Height          =   180
         Index           =   2
         Left            =   -74640
         TabIndex        =   161
         Top             =   2640
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人4"
         Height          =   180
         Index           =   1
         Left            =   -74640
         TabIndex        =   160
         Top             =   3780
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   33
         Left            =   -74280
         TabIndex        =   159
         Top             =   5610
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   34
         Left            =   -74280
         TabIndex        =   158
         Top             =   5370
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   35
         Left            =   -74280
         TabIndex        =   157
         Top             =   5130
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人5"
         Height          =   180
         Index           =   3
         Left            =   -74640
         TabIndex        =   156
         Top             =   4890
         Width           =   630
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "圖式圖數:"
         Height          =   180
         Left            =   1710
         TabIndex        =   114
         Top             =   2380
         Width           =   765
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "申請專利範圍項數:"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   1710
         TabIndex        =   113
         Top             =   2103
         Width           =   1485
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "頁數總計:"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   1710
         TabIndex        =   112
         Top             =   1830
         Width           =   765
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "圖式頁數:"
         Height          =   180
         Left            =   1710
         TabIndex        =   111
         Top             =   1557
         Width           =   765
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請專利範圍頁數:"
         Height          =   180
         Left            =   1710
         TabIndex        =   110
         Top             =   1284
         Width           =   1485
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "說明書頁數:"
         Height          =   180
         Left            =   1710
         TabIndex        =   109
         Top             =   738
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "摘要頁數:"
         Height          =   180
         Left            =   1710
         TabIndex        =   108
         Top             =   465
         Width           =   765
      End
      Begin VB.Shape Shape2 
         Height          =   3492
         Left            =   132
         Top             =   348
         Width           =   3612
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "附件文書"
         Height          =   180
         Left            =   3876
         TabIndex        =   107
         Top             =   456
         Width           =   720
      End
      Begin VB.Shape Shape3 
         Height          =   3495
         Left            =   3780
         Top             =   345
         Width           =   4050
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "是否一併提實審:          (Y:是)"
         Height          =   180
         Left            =   -74760
         TabIndex        =   105
         Top             =   1890
         Width           =   2220
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "繳費金額:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   104
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label lblFavReason 
         AutoSize        =   -1  'True
         Caption         =   "原因:"
         Height          =   180
         Left            =   -72195
         TabIndex        =   103
         Top             =   1230
         Width           =   405
      End
      Begin VB.Label lblFavDate 
         AutoSize        =   -1  'True
         Caption         =   "優惠期發生日期:"
         Height          =   180
         Left            =   -74895
         TabIndex        =   102
         Top             =   1230
         Width           =   1305
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   101
         Top             =   2190
         Width           =   945
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5"
         Height          =   180
         Index           =   4
         Left            =   -74925
         TabIndex        =   100
         Top             =   4680
         Width           =   810
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4"
         Height          =   180
         Index           =   3
         Left            =   -74955
         TabIndex        =   99
         Top             =   3675
         Width           =   810
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3"
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   98
         Top             =   2670
         Width           =   810
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2"
         Height          =   180
         Index           =   1
         Left            =   -74910
         TabIndex        =   97
         Top             =   1650
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   23
         Left            =   -74040
         TabIndex        =   96
         Top             =   4920
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   22
         Left            =   -74040
         TabIndex        =   95
         Top             =   4680
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   21
         Left            =   -74040
         TabIndex        =   94
         Top             =   4440
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   20
         Left            =   -74040
         TabIndex        =   93
         Top             =   3900
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   19
         Left            =   -74040
         TabIndex        =   92
         Top             =   3660
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   18
         Left            =   -74040
         TabIndex        =   91
         Top             =   3420
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   17
         Left            =   -74040
         TabIndex        =   90
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   16
         Left            =   -74040
         TabIndex        =   89
         Top             =   2640
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   15
         Left            =   -74040
         TabIndex        =   88
         Top             =   2400
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   14
         Left            =   -74040
         TabIndex        =   87
         Top             =   1860
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   13
         Left            =   -74040
         TabIndex        =   86
         Top             =   1620
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   12
         Left            =   -74040
         TabIndex        =   85
         Top             =   1380
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   11
         Left            =   -74040
         TabIndex        =   84
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   10
         Left            =   -74040
         TabIndex        =   83
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   9
         Left            =   -74040
         TabIndex        =   82
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1"
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   81
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "發明人:"
         Height          =   180
         Index           =   0
         Left            =   -74820
         TabIndex        =   80
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   0
         Left            =   -73980
         TabIndex        =   79
         Top             =   465
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱"
         Height          =   180
         Index           =   1
         Left            =   -74760
         TabIndex        =   78
         Top             =   465
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   1
         Left            =   -73980
         TabIndex        =   77
         Top             =   735
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6360
      TabIndex        =   17
      Top             =   60
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   2
      Left            =   7020
      TabIndex        =   18
      Top             =   60
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   984
      MaxLength       =   3
      TabIndex        =   71
      Top             =   12
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1536
      MaxLength       =   6
      TabIndex        =   70
      Top             =   12
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2376
      MaxLength       =   1
      TabIndex        =   20
      Top             =   12
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   19
      Top             =   12
      Width           =   375
   End
   Begin MSForms.Label Label7 
      Height          =   195
      Index           =   11
      Left            =   990
      TabIndex        =   224
      Top             =   300
      Width           =   1335
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2355;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label7 
      Height          =   195
      Index           =   12
      Left            =   3960
      TabIndex        =   223
      Top             =   300
      Width           =   1335
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2355;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label7 
      Height          =   195
      Index           =   10
      Left            =   3960
      TabIndex        =   222
      Top             =   30
      Width           =   1335
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2355;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "一案兩請"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   5430
      TabIndex        =   204
      Top             =   60
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3144
      TabIndex        =   75
      Top             =   12
      Width           =   768
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3144
      TabIndex        =   74
      Top             =   312
      Width           =   768
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   144
      TabIndex        =   73
      Top             =   312
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   144
      TabIndex        =   72
      Top             =   12
      Width           =   768
   End
End
Attribute VB_Name = "frm06010301_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/10/12 Form2.0已修改
'Modified by Morgan 2013/6/17 原程式未使用,改為產生電子申請書
Option Explicit

Dim strReceiveNo As String, intWhere As Integer
Dim pa() As String, cp() As String
Dim CP10 As String, m_CaseNo As String
Dim m_DualCase As Boolean
Dim m_CPM26 As String 'Add By Sindy 2018/6/28
Dim m_TF01 As String, m_TF01pty As String 'Added by Lydia 2018/08/08 記錄中說進度和案件性質
Dim oText  As Control  'Added by Lydia 2018/12/27
Dim m_lngOverPageFee As Long, m_lngOverItemFee As Long '超頁費,超項費
Dim m_AgentName As String 'Add By Sindy 2021/5/10
'Add By Sindy 2022/5/3
Public m_PrevForm As Form '前一畫面
Dim m_PrevForm_Text6 As String
'2022/5/3 END


'Add By Sindy 2022/5/3
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'Add By Sindy 2019/1/31
Private Sub Check3_Click()
   If Check3.Value = 1 Then
      chkAtt(25).Value = 1
   End If
End Sub

Private Sub chkAtt_Click(Index As Integer)
   If Index = 0 Then
      chkAtt(0).Value = 2
   End If
End Sub

Private Sub chkBiology_Click(Index As Integer)
   'Modified by Lydia 2018/12/27
   'Dim oText As TextBox, bolEnabled As Boolean
   Dim bolEnabled As Boolean
   
   '寄存
   If Index = 0 Then
      If chkBiology(0).Value = vbChecked Then
         bolEnabled = True
         chkBiology(1).Value = vbUnchecked
      Else
         bolEnabled = False
         chkAtt(8).Value = chkBiology(Index).Value
         chkAtt(9).Value = chkBiology(Index).Value
      End If
      For Each oText In txtBiology
         oText.Enabled = bolEnabled
      Next
   '不須寄存
   Else
      chkAtt(10).Value = chkBiology(1).Value
      
      If chkBiology(1).Value = vbChecked Then
         chkBiology(0).Value = vbUnchecked
      End If
   End If
   
End Sub

Private Sub chkDoc_Click(Index As Integer)
   'Modified by Lydia 2018/12/27
   'Dim oText As TextBox, bEnabled As Boolean
   Dim bEnabled As Boolean
   
   If chkDoc(Index).Value = 1 Then
      bEnabled = True
   Else
      bEnabled = False
   End If
   
   Select Case Index
   Case 0 '中文
      For Each oText In txtDocCh
         oText.Enabled = bEnabled
      Next
   Case 1 '外文
      txtForeign.Enabled = bEnabled
      cboLagnuage.Enabled = bEnabled
      chkAtt(5).Value = chkDoc(Index).Value
      'Added by Lydia 2018/08/08
      txtTF24.Enabled = bEnabled
      txtTF25.Enabled = bEnabled
   Case 2 '簡體
      txtSimplified.Enabled = bEnabled
      chkAtt(6).Value = chkDoc(Index).Value
   End Select
End Sub

Private Sub chkEexcerpt_Click()
   'Add By Sindy 2019/2/13
   If chkEexcerpt.Value = 0 Then
      If Val(chkEexcerpt.Tag) > 0 Then
         txtCP84 = Val(txtCP84) + Val(chkEexcerpt.Tag)
         txtCP84.Tag = txtCP84.Text
         chkEexcerpt.Tag = ""
      End If
      If InStr(txtMemo, "減收申請規費800元") > 0 Then
         txtMemo = ""
      End If
   Else
      txtCP84 = Val(txtCP84) - 800
      txtCP84.Tag = txtCP84.Text
      chkEexcerpt.Tag = 800
      If InStr(txtMemo, "減收申請規費800元") = 0 Then
         txtMemo = "本案未附英文說明書，但所檢附之申請書中發明名稱、申請人姓名或名稱、發明人姓名及摘要同時附有英文翻譯，故可減收申請規費800元，英文發明名稱及英文摘要後補。"
      End If
   End If
   '2019/2/13 END
End Sub

Private Sub Form_Load()
   'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    Text6(3).MaxLength = Pub_MaxCEL10
    Text6(4).MaxLength = Pub_MaxCEL11
    Text6(6).MaxLength = Pub_MaxCEL10
    Text6(7).MaxLength = Pub_MaxCEL11
    Text6(24).MaxLength = Pub_MaxCEL10
    Text6(25).MaxLength = Pub_MaxCEL11
    Text6(27).MaxLength = Pub_MaxCEL10
    Text6(28).MaxLength = Pub_MaxCEL11
    Text6(30).MaxLength = Pub_MaxCEL10
    Text6(31).MaxLength = Pub_MaxCEL11
    Text6(33).MaxLength = Pub_MaxCEL10
    Text6(34).MaxLength = Pub_MaxCEL11
    Text6(36).MaxLength = Pub_MaxCEL10
    Text6(37).MaxLength = Pub_MaxCEL11
    Text6(39).MaxLength = Pub_MaxCEL10
    Text6(40).MaxLength = Pub_MaxCEL11
    Text6(42).MaxLength = Pub_MaxCEL10
    Text6(43).MaxLength = Pub_MaxCEL11
    Text6(45).MaxLength = Pub_MaxCEL10
    Text6(46).MaxLength = Pub_MaxCEL11
   'end 2016/09/10
   
   MoveFormToCenter Me
   intWhere = 國外_FC
   With m_PrevForm
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
      'Add By Sindy 2022/5/3
      If UCase(TypeName(m_PrevForm)) = UCase("frm090904") Then '工程師各式申請書
         m_PrevForm_Text6 = "3" '電子送件
      Else
         m_PrevForm_Text6 = .Text6
      End If
      '2022/5/3 END
   End With
   
   'Add By Sindy 2018/6/15
   If m_PrevForm_Text6 = "4" Then '新案申請書
      Frame2.Visible = True
   Else
      Frame2.Visible = False
   End If
   '2018/6/15 END
   
   ReDim pa(1 To TF_PA) As String
   ReDim cp(TF_CP)
   ReadPatent
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), cp(110), cp(10), True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1100
   lstNameAgent.Width = 1300

   'Added by Morgan 2018/1/11 設計案
   lblCountry.BackColor = SSTab1.BackColor
   If pa(8) <> "1" Then
      '是否一併提實審
      Label30.Visible = False
      Text7.Visible = False
         
      strExc(1) = IIf(pa(8) = "2", "新型", "設計")
      chkAtt(1).Caption = Replace(chkAtt(1).Caption, "發明", strExc(1)) '摘要
      chkAtt(2).Caption = Replace(chkAtt(2).Caption, "發明", strExc(1)) '說明書
      chkAtt(3).Caption = Replace(chkAtt(3).Caption, "發明", strExc(1)) '專利範圍
      chkAtt(4).Caption = Replace(chkAtt(4).Caption, "發明", strExc(1)) '圖式
      
      chkAtt(14).Visible = False '序列表
      chkAtt(8).Visible = False '國內生物材料寄存證明文件
      chkAtt(9).Visible = False '國外生物材料寄存證明文件
      chkAtt(10).Visible = False '生物材料為通常知識者易於獲得證明文件
      Frame1.Visible = False '生物寄存
      '設計
      If pa(8) = "3" Then
         '摘要頁數
         Label15.Visible = False
         txtDocCh(0).Visible = False
         '申請專利範圍頁數
         Label17.Visible = False
         txtDocCh(2).Visible = False
         '申請專利範圍項數
         Label21.Visible = False
         txtDocCh(5).Visible = False
         chkAtt(1).Visible = False '摘要
         chkAtt(3).Visible = False '專利範圍
      End If
   End If
   'end 2018/1/11
   'Add By Sindy 2022/5/3
   Frame3.Visible = False
   If cp(10) = 改請衍生設計 Then
      SSTab1.Tab = 1
      Frame3.Visible = True: Frame3.Top = 3852: Frame3.Left = 108 '援用原申請案優先權主張..等等
      chkAtt(7).Visible = False '委任書
      chkAtt(12).Visible = False '國際優先權證明文件
      chkAtt(13).Visible = False '圖式
      chkAtt(15).Visible = False '優惠期證明文件
      chkAtt(25).Visible = False '委任書 (附譯文)
      Check3.Visible = False '個案
      chkAtt(11).Visible = False '一案兩請聲明
      cmdOpen(0).Visible = False '外文本
      cmdOpen(1).Visible = False '電子送件暫存區
   End If
   '2022/5/3 END
   
   SSTab1.Tab = 0
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing 'Add By Sindy 2022/5/3
   Set frm06010301_1 = Nothing
End Sub

'Add By Sindy 2014/11/14
Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("序號", "發明人編號", "中文名稱", "英文名稱", "日文名稱")
   arrGridHeadWidth = Array(500, 1100, 2200, 2200, 2200)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

'************************************************
' 取回專利基本資料及收文資料
'
'************************************************
Private Sub ReadPatent()
Dim i As Integer, j As Integer, Lbl As Object ', strTempName As String
'Add By Sindy 2014/11/14
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'2014/11/14 END

   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   For Each Lbl In Label7
      Lbl.Caption = ""
   Next
   For i = 0 To 9
      Combo2(i).AddItem ""
   Next
   If ClsPDReadPatentDatabase(pa(), intWhere) Then
      For i = 0 To 1 '案件名稱
         Text6(i) = pa(i + 5)
      Next
      For i = 3 To 8 '代表人1.2
         Text6(i) = pa(i + 76)
      Next
      
      For i = 24 To 47 '代表人3.4.5.6.7.8.9.10
         Text6(i) = pa(i + 85)
      Next
      
      For i = 9 To 23 '申請人 12345 地址
         Text6(i) = pa(i + 22)
      Next
      
      '發明人代號
'      strTempName = ""
'      For i = 26 To 30
'         If IsEmptyText(pa(i)) = False Then
'            If Len(pa(i)) > 8 Then
'               strTempName = strTempName & "'" & Left(pa(i), 8) & "',"
'            Else
'               strTempName = strTempName & "'" & Left(pa(i), 8) & String(8 - Len(pa(i)), "0") & "',"
'            End If
'         End If
'      Next
'      If strTempName <> "" Then strTempName = Left(strTempName, Len(strTempName) - 1)
'      strExc(0) = "SELECT NVL(IN05,NVL(IN04,IN06))||'('||IN01||IN02||')' FROM INVENTOR WHERE IN01 IN (" & strTempName & ")"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      With RsTemp
'         If intI = 1 Then
'            Do While Not .EOF
'               For i = 0 To 9
'                  Combo1.AddItem .Fields(0)
'               Next
'               .MoveNext
'            Loop
'         End If
'      End With
'      For i = 0 To 9
'         If pa(i + 60) <> "" Then
'            Combo1(i).Text = GetInventorName(pa(i + 60))
'            ChgType i + 10
'         Else
'            Combo1(i).ListIndex = 0
'         End If
'      Next
      'Modify By Sindy 2014/11/14
      GRD1.Clear
      SetGrd
      StrSQLa = "SELECT pi05 as 序號,pi06 as 發明人編號,in04 as 中文名稱,in05 as 英文名稱,in06 as 日文名稱 from PatentInventor,Inventor where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
                " and substr(pi06,1,8)=in01(+) and substr(pi06,9,2)=in02(+)" & _
                " order by pi05 asc"
      If rsA.State <> adStateClosed Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         Set GRD1.Recordset = rsA
      End If
      '2014/11/14 END
      
      '代表人
      For i = 26 To 30
         If pa(i) <> "" Then
            strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(pa(i))
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               For j = 1 To 6
                  If IsNull(RsTemp.Fields(j - 1)) Then
                     strExc(0) = ""
                  Else
                     strExc(0) = "-" & RsTemp.Fields(j - 1)
                  End If
                  Combo2((i - 26) * 2).AddItem pa(i) & "-" & j & strExc(0)
                  Combo2((i - 26) * 2 + 1).AddItem pa(i) & "-" & j & strExc(0)
               Next
            End If
         End If
      Next
   End If
   
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If ClsPDGetStaff(cp(13), strExc(0)) Then Label7(12) = strExc(0)
      If ClsPDGetStaff(cp(14), strExc(0)) Then Label7(11) = strExc(0)
      If ClsPDGetCaseProperty("FCP", cp(10), strExc(0)) Then Label7(10) = strExc(0)
   End If
   
   'Add By Sindy 2018/6/28 抓副檔名
   strExc(0) = "select cpm26 from casepropertymap WHERE cpm01='" & cp(1) & "' and cpm02='" & cp(10) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_CPM26 = "" & RsTemp.Fields("cpm26")
   End If
   '2018/6/28 END
   'Add By Sindy 2018/7/2
   If m_CPM26 = "" Then
      If cp(10) = 發明申請 Then
         m_CPM26 = "inv"
      ElseIf cp(10) = 新型申請 Then
         m_CPM26 = "utl"
      ElseIf cp(10) = 設計申請 Or cp(10) = 衍生設計 Then
         m_CPM26 = "des"
      End If
   End If
   '2018/7/2 END
   
   'Add By Sindy 2019/2/27 敏莉:請將新型案、設計案的新案電子送件各式申請書扣800元規費的欄位鎖起來，因只有發明案才有扣800元
   chkEexcerpt.Visible = False
   If pa(8) = "1" Then '發明
      chkEexcerpt.Visible = True
   End If
   '2019/2/27 END
   
   strExc(0) = "select PD05 AS  優先權日,PD06 AS 優先權號,NA03 AS 優先權國家 from PRIDATE,NATION WHERE PD01='" & pa(1) & "' AND PD02='" & pa(2) & "' AND PD03='" & pa(3) & "' AND PD04 ='" & pa(4) & "' AND PD07=NA01(+) ORDER BY PD01,PD02,PD03,PD04"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set grdDataList2.Recordset = RsTemp.Clone
   
   '是否一併提實審
   If PUB_ChkCPExist(pa, "416") = True Then
      Text7 = "Y"
      'Add By Sindy 2018/6/15 一併提實審時,存頁/項數
      If m_PrevForm_Text6 = "4" Then '新案申請書
         If Val(cp(135)) > 0 Then txtCP135 = cp(135)
         If Val(cp(136)) > 0 Then txtCP136 = cp(136)
      End If
      '2018/6/15 END
   End If
   Call Text7_Validate(False) 'Add By Sindy 2018/5/28
   
   If pa(140) <> "" Then
      txtFavDate = TransDate(pa(140), 1)
   End If
   
   '一案兩請檢查
   strExc(0) = "select 1 from casemap where cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "' and cm10='3'" & _
      " union select 1 from casemap where cm05='" & pa(1) & "' and cm06='" & pa(2) & "' and cm07='" & pa(3) & "' and cm08='" & pa(4) & "' and cm10='3'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_DualCase = True
      Label4.Visible = True
      chkAtt(11).Value = vbChecked
   End If
   
   'Add By Sindy 2018/6/27
   '在"各式申請書-新案"(包含電子送件和紙本)若翻譯費用檔已有外文本說明書頁數和圖示頁數(TransFee.TF24,TF25)，
   '代入產出檔案的外文本頁數和圖示，並且加總設為合計頁數。
   'Modified by Lydia 2018/08/08 改成有收文新案翻譯201、檢視中說209或核對中說格式235才會回寫到新案建檔的翻譯頁籤(TransFee.TF24,TF25)
'   strExc(0) = "select nvl(TF24,0) TF24,nvl(TF25,0) TF25 from TransFee where TF01='" & strReceiveNo & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If RsTemp.Fields("TF24") + RsTemp.Fields("TF25") > 0 Then
'         chkDoc(1).Value = 1
'         txtForeign.Text = RsTemp.Fields("TF24") + RsTemp.Fields("TF25")
'      End If
'   End If
   '2018/6/27 END
   'Added by Lydia 2018/08/08  讀取翻譯費用檔
   strExc(0) = "SELECT CP09,CP10,CP14,CP27,B.TF24,B.TF25 FROM CaseProgress A,TransFee B " & _
                     "WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 in ('201','209','235') AND CP159=0 AND CP09=TF01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
         m_TF01 = "" & RsTemp.Fields("CP09")
         m_TF01pty = "" & RsTemp.Fields("CP10")
         txtTF24.Text = "" & RsTemp.Fields("TF24")
         txtTF25.Text = "" & RsTemp.Fields("TF25")
   Else
         m_TF01 = ""
         m_TF01pty = ""
         txtTF24.Text = ""
         txtTF25.Text = ""
   End If
   txtTF24.Tag = txtTF24.Text
   txtTF25.Tag = txtTF25.Text
   If Val(txtTF24.Text) + Val(txtTF25.Text) > 0 Then
         chkDoc(1).Value = 1
         txtForeign.Text = Val(txtTF24.Text) + Val(txtTF25.Text)
   End If
   'end 2018/08/08
   
   'Added by Lydia 2018/12/27 中文本資訊-各項頁數
    txtDocCh(0).Text = pa(64) '摘要頁數
    txtDocCh(1).Text = pa(65) '說明書頁數
    txtDocCh(7).Text = pa(66) '序列表頁數
    txtDocCh(2).Text = pa(67) '申請專利範圍頁數
    txtDocCh(3).Text = pa(68) '圖式頁數
    'Added by Lydia 2019/01/10
    txtDocCh(5).Text = pa(172) '申請專利範圍項數(最初項數)
    txtDocCh(6).Text = pa(173) '圖式圖數
    'end 2019/01/10
    If Val(pa(64)) + Val(pa(65)) + Val(pa(66)) + Val(pa(67)) + Val(pa(68)) > 0 Then
        chkDoc(0).Value = 1
        Call txtDocCh_Validate(0, False)
        Call txtDocCh_Validate(1, False)
        Call txtDocCh_Validate(2, False)
        Call txtDocCh_Validate(3, False)
    End If
    For Each oText In txtDocCh
        oText.Tag = oText.Text
    Next
   'end 2018/12/27
    
   FraPA174.Visible = False 'Added by Lydia 2020/02/21
   
   'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：判斷檔案上傳目的地
   If pa(1) = "FCP" Or pa(1) = "P" Then
       ' 已放在原始檔區
       If PUB_ChkCPExist(pa, cntEnglish_Vers, , strExc(1), , "D") = True Then 'English_Vers992
            cmdOpen(0).Caption = "原始檔"
            cmdOpen(0).Tag = strExc(1)
       Else
            cmdOpen(0).Caption = "外文本"
            cmdOpen(0).Tag = ""
       End If
      'Added by Lydia 2020/02/21 預設「名稱有特殊字」
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
      'end 2020/02/21
   End If
   'end 2020/01/20
      
   'Added by Lydia 2021/04/09 預設附件勾選
     '外文本序列表(.SEQ.)：1.English Vers原始檔區有.SEQ.檔案、2.新案建檔、工程師命名作業勾選□有序列表
     '國際優先權證明文件(.PRI.)、委任書(.POA.)：
     '1.English Vers原始檔區：在新案收文立卷搬入”國際優先權證明文件(.PRI.)、委任書(.POA.)”之檔案。
     '2.新案收文、B類收文補文件202之卷宗區檔案：有”國際優先權證明文件(.PRI.)、委任書(.POA.)”之檔案；承辦人員在”客戶提供文件”上傳檔案到卷宗區，在經過程序人員做”補文件”處理。
   If SSTab1.TabVisible(1) = True Then '電子送件
      If pa(175) = "Y" Then chkAtt(14).Value = 1
        
      strExc(0) = ""
      strExc(1) = "(UPPER(CPF02) LIKE '%.SEQ.%' OR UPPER(CPF02) LIKE '%.PRI.%' OR UPPER(CPF02) LIKE '%.POA.%') "
      '原始檔區
      If PUB_ChkCPExist(pa, cntEnglish_Vers, , strExc(2), , "D") = True Then
            strExc(0) = "SELECT '1' as pkind, CPF01 as pKeyNo,CPF02  as pFName FROM CASEPROGRESS A,CASEPAPERFILE B " & _
                             "WHERE CP09='" & strExc(2) & "' AND CP159=0 AND CP09=CPF01(+) " & _
                             "AND NVL(CPF10,'N') <> 'D' AND " & strExc(1)
      End If
      strSql = "SELECT '2' as pkind, CPP01 as pKeyNo,CPP02 as pFName FROM CASEPROGRESS A,CASEPAPERPDF B " & _
                   "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP159=0 " & _
                   " AND (CP31='Y' OR (CP10='202' AND SUBSTR(CP09,1,1)='B')) " & _
                   "AND CP09=CPP01(+) AND NVL(CPP10,'N') <> 'D' AND " & Replace(strExc(1), "CPF", "CPP") & IIf(strExc(0) <> "", " UNION ", "") & strExc(0)
      strSql = strSql & " ORDER BY 1, 2, 3"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
          RsTemp.MoveFirst
          Do While Not RsTemp.EOF
                If InStr("" & RsTemp.Fields("pfname"), ".SEQ.") > 0 And chkAtt(14).Value = 0 Then
                   chkAtt(14).Value = 1
                ElseIf InStr("" & RsTemp.Fields("pfname"), ".PRI.") > 0 And chkAtt(12).Value = 0 Then
                   chkAtt(12).Value = 1
                ElseIf InStr("" & RsTemp.Fields("pfname"), ".POA.") > 0 And chkAtt(7).Value = 0 Then
                   chkAtt(7).Value = 1
                End If
                RsTemp.MoveNext
          Loop
      End If
 End If
 'end 2021/04/09
End Sub

'Modify By Sindy 2015/12/2
'Private Sub cmdOK_Click(Index As Integer)
Public Sub cmdok_Click(Index As Integer)
'2015/12/2 END
   Dim strFolder As String, strFileName As String
   Dim Cancel As Boolean
   Dim bolHad As Boolean, ii As Integer 'Add By Sindy 2021/10/12
   
   Select Case Index
      Case 0
         Call Text7_Validate(False)
         
         'Add By Sindy 2021/10/12 點選新案(101~103)，出電子送件請書時：
         '若此案已有代表人資料，但未勾選委任書時，在按確定後，
         '請彈跳訊息：[此案已有代表人資料，是否一併附委任書？]
         '[是:委任書打勾後，直接出申請書，申請書一併帶出附送書件:委任書]
         '[否:委任書維持不勾選，直接出申請書]
         '避免漏送委任書
         bolHad = False
         If InStr("101,102,103", cp(10)) > 0 And chkAtt(7).Value = 0 Then
            For ii = 79 To 84
               If pa(ii) <> "" Then
                  bolHad = True
                  Exit For
               End If
            Next ii
            If bolHad = False Then
               For ii = 109 To 132
                  If pa(ii) <> "" Then
                     bolHad = True
                     Exit For
                  End If
               Next ii
            End If
            If bolHad = True Then
               If MsgBox("此案已有代表人資料，是否一併附委任書？", vbYesNo + vbQuestion) = vbYes Then
                  chkAtt(7).Value = 1
               End If
            End If
         End If
         '2021/10/12 END
         
         'Add By Sindy 2015/11/26
         If m_PrevForm_Text6 = "4" Then '新案申請書
            Cancel = False
            lstNameAgent_Validate Cancel
            If Cancel = True Then
               SSTab1.Tab = 0
'               lstNameAgent.SetFocus
               Exit Sub
            End If
            
            'Add By Sindy 2018/6/15
            If Text7 = "Y" Then '一併提實審時,要輸入總頁,項數
               'Modify By Sindy 2018/6/28 沒一起送中說時,可以不輸頁項數
'               If Val(txtCP135) = 0 Then
'                  MsgBox "請輸入總頁數！", vbInformation
'                  SSTab1.Tab = 1
'                  txtCP135.SetFocus
'                  Exit Sub
'               ElseIf Val(txtCP136) = 0 Then
'                  MsgBox "請輸入總項數！", vbInformation
'                  SSTab1.Tab = 1
'                  txtCP136.SetFocus
'                  Exit Sub
'               End If
            End If
            '2018/6/15 END
            
            'Added by Lydia 2020/02/21 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
            If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
                MsgBox MsgText(1111), vbInformation
                If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                     Exit Sub
                End If
            End If
            'end 2020/02/21
            
            If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
            If GetApplBook = True Then
               m_PrevForm.ClearForm
            End If
            m_PrevForm.Show
            Unload Me
         Else
         '2015/11/26 END
            
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            'Added by Lydia 2020/02/21 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
            If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
                MsgBox MsgText(1111), vbInformation
                If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                    Exit Sub
                End If
            End If
            'end 2020/02/21
            
            If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
            
            'Modified by Morgan 2018/3/8 檔名改數字第1碼0也要
            'm_CaseNo = pa(1) & IIf(Left(pa(2), 1) = "0", Mid(pa(2), 2), pa(2)) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "")
            m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
            'Ex.FCP47842 --> \\Typing2\專利案件\478\47842\FCP47842 (主管看過後會更改資料夾名稱加上 "-案件性質代碼" 如 FCP47842-101)
            'If Pub_StrUserSt03 = "M51" Then
            If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
               strFolder = PUB_Getdesktop
            Else
               strFolder = FCP電子送件檔案存放路徑
            End If
            'Modify By Sindy 2017/10/26 敏莉:請將電子送件所產生的新案及補文件申請書改路徑為:\\Typing2\電子送件暫存區\FCPXXXXX
'            strFolder = strFolder & "\" & Mid(m_CaseNo, 4, Len(m_CaseNo) - 5)
'            If Dir(strFolder, vbDirectory) = "" Then
'               MkDir strFolder
'            End If
            
'Removed by Morgan 2017/9/21 --葉敏莉
'            strFolder = strFolder & "\" & Mid(m_CaseNo, 4)
'            If Dir(strFolder, vbDirectory) = "" Then
'               MkDir strFolder
'            End If
            
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            
'            '1.基本資料
'            'Added by Lydia 2022/03/03 區分基本資料表(用在101,102,103,307)
'            If pa(1) = "FCP" And InStr("101,102,103,307", cp(10)) > 0 Then
'                StartLetterPA_EData "01", "B1", strReceiveNo, pa, cp
'                NowPrint strReceiveNo, "01", "B1", False, strUserNum, , , True, strExc(9)
'            Else
'            'end 2022/03/03
               StartLetterPA_EData "01", "14", strReceiveNo, pa, cp 'Modify By Sindy 2014/11/14 "13"==>"14"
               NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(9)
'            End If   'Added by Lydia 2022/03/03
            strFileName = strFolder & "\" & m_CaseNo & ".contact"
            If PUB_MakeDoc(strExc(9), strFileName) = True Then
               'Removed by Morgan 2017/8/9 目前 Html2Pdf 有問題,暫取消
               'If ConvertHtml2Pdf("A9801", strFileName) = False Then
               '   Exit Sub
               'End If
            End If
            
            '2.申請書
            StartLetter2 "01", "03" 'Modify By Sindy 2014/11/14 "02"==>"03"
            NowPrint strReceiveNo, "01", "03", False, strUserNum, , , True, strExc(9)
            'Modified by Morgan 2018/1/12 +設計專利申請書
            If pa(8) = "1" Then
               strFileName = strFolder & "\" & "發明專利申請書"
            ElseIf pa(8) = "2" Then
               strFileName = strFolder & "\" & "新型專利申請書"
            ElseIf pa(8) = "3" Then
               'Add By Sindy 2022/5/3 frm090904.工程師各式申請書 會呼叫此作業產出改請衍生設計
               If cp(10) = 改請衍生設計 Then
                  strFileName = strFolder & "\" & "改請衍生設計專利申請書"
               'Add By Sindy 2018/8/7
               ElseIf cp(10) = 衍生設計 Then
                  strFileName = strFolder & "\" & "衍生設計專利申請書"
               Else
               '2018/8/7 END
                  strFileName = strFolder & "\" & "設計專利申請書"
               End If
            End If
            
            If PUB_MakeDoc(strExc(9), strFileName) = True Then
               'Removed by Morgan 2017/8/9 目前 Html2Pdf 有問題,暫取消
               'If ConvertHtml2Pdf("P0" & cp(10), strFileName) = False Then
               '   Exit Sub
               'End If
            End If
            
            'Add By Sindy 2019/1/31
            Dim strCaseData As String
            Dim strBookDate As String
            If chkAtt(25).Value = 1 Then '委任書(附譯文)
               If Check3.Value = 1 Then
                  '個案
                  strCaseData = "　　有關台灣專利申請案第" & pa(11) & "號『" & pa(5) & "』之" & IIf(pa(8) = "1", "發明", IIf(pa(8) = "2", "新型", "設計")) & "專利申請案，"
               End If
               '委任日期
               strBookDate = Left(strSrvDate(2), 3) & " 年 " & Mid(strSrvDate(2), 4, 2) & " 月 " & Right(strSrvDate(2), 2) & " 日"
               If pa(26) <> "" Then
                  EndLetter "01", strReceiveNo, "31", strUserNum
                  '個案
                  If Check3.Value = 1 Then
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('01','" & strReceiveNo & "','31','" & strUserNum & "','個案','" & strCaseData & "')"
                     cnnConnection.Execute strExc(0)
                  End If
                  '委任日期
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','31','" & strUserNum & "','委任日期','" & strBookDate & "')"
                  cnnConnection.Execute strExc(0)
                  '申請人國籍
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','31','" & strUserNum & "','申請人國籍','　" & Trim(GetNationName(GetPrjNationNumber1(ChangeCustomerL(pa(26))))) & "　')"
                  cnnConnection.Execute strExc(0)
                  NowPrint strReceiveNo, "01", "31", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "委任狀中譯文1"
                  Call PUB_MakeDoc(strExc(9), strFileName, , pa(1), "01", "31")
               End If
               If pa(27) <> "" Then
                  EndLetter "01", strReceiveNo, "32", strUserNum
                  '個案
                  If Check3.Value = 1 Then
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('01','" & strReceiveNo & "','32','" & strUserNum & "','個案','" & strCaseData & "')"
                     cnnConnection.Execute strExc(0)
                  End If
                  '委任日期
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','32','" & strUserNum & "','委任日期','" & strBookDate & "')"
                  cnnConnection.Execute strExc(0)
                  '申請人國籍
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','32','" & strUserNum & "','申請人國籍','　" & Trim(GetNationName(GetPrjNationNumber1(ChangeCustomerL(pa(27))))) & "　')"
                  cnnConnection.Execute strExc(0)
                  NowPrint strReceiveNo, "01", "32", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "委任狀中譯文2"
                  Call PUB_MakeDoc(strExc(9), strFileName, , pa(1), "01", "32")
               End If
               If pa(28) <> "" Then
                  EndLetter "01", strReceiveNo, "33", strUserNum
                  '個案
                  If Check3.Value = 1 Then
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('01','" & strReceiveNo & "','33','" & strUserNum & "','個案','" & strCaseData & "')"
                     cnnConnection.Execute strExc(0)
                  End If
                  '委任日期
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','33','" & strUserNum & "','委任日期','" & strBookDate & "')"
                  cnnConnection.Execute strExc(0)
                  '申請人國籍
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','33','" & strUserNum & "','申請人國籍','　" & Trim(GetNationName(GetPrjNationNumber1(ChangeCustomerL(pa(28))))) & "　')"
                  cnnConnection.Execute strExc(0)
                  NowPrint strReceiveNo, "01", "33", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "委任狀中譯文3"
                  Call PUB_MakeDoc(strExc(9), strFileName, , pa(1), "01", "33")
               End If
               If pa(29) <> "" Then
                  EndLetter "01", strReceiveNo, "34", strUserNum
                  '個案
                  If Check3.Value = 1 Then
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('01','" & strReceiveNo & "','34','" & strUserNum & "','個案','" & strCaseData & "')"
                     cnnConnection.Execute strExc(0)
                  End If
                  '委任日期
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','34','" & strUserNum & "','委任日期','" & strBookDate & "')"
                  cnnConnection.Execute strExc(0)
                  '申請人國籍
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','34','" & strUserNum & "','申請人國籍','　" & Trim(GetNationName(GetPrjNationNumber1(ChangeCustomerL(pa(29))))) & "　')"
                  cnnConnection.Execute strExc(0)
                  NowPrint strReceiveNo, "01", "34", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "委任狀中譯文4"
                  Call PUB_MakeDoc(strExc(9), strFileName, , pa(1), "01", "34")
               End If
               If pa(30) <> "" Then
                  EndLetter "01", strReceiveNo, "35", strUserNum
                  '個案
                  If Check3.Value = 1 Then
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                                 "('01','" & strReceiveNo & "','35','" & strUserNum & "','個案','" & strCaseData & "')"
                     cnnConnection.Execute strExc(0)
                  End If
                  '委任日期
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','35','" & strUserNum & "','委任日期','" & strBookDate & "')"
                  cnnConnection.Execute strExc(0)
                  '申請人國籍
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('01','" & strReceiveNo & "','35','" & strUserNum & "','申請人國籍','　" & Trim(GetNationName(GetPrjNationNumber1(ChangeCustomerL(pa(30))))) & "　')"
                  cnnConnection.Execute strExc(0)
                  NowPrint strReceiveNo, "01", "35", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "委任狀中譯文5"
                  Call PUB_MakeDoc(strExc(9), strFileName, , pa(1), "01", "35")
               End If
            End If
            '2019/1/31 END
            
            m_PrevForm.Show
            m_PrevForm.ClearForm
            Unload Me
         End If
      Case 2
         m_PrevForm.Show
         m_PrevForm.cmdok_Click 1
         Unload Me
   End Select
End Sub

'Add By Sindy 2015/11/26 產生申請書
Private Function GetApplBook() As Boolean
Dim m_FileName As String
Dim i As Integer, int_TotCnt As Integer
Dim strName As String, strText As String, strLineText As String
Dim strApplLineText As String
Dim intAppCnt As Integer, intAppCnt2 As Integer, strApplData As String
Dim intInvCnt As Integer, strApplInventor As String
Dim strApplPriDate As String
Dim strCP84 As String
Dim bolFileExist As Boolean 'Add By Sindy 2015/12/14

On Error GoTo ErrHand
   
   GetApplBook = False
   
   m_MySt(1) = pa(1)
   m_MySt(2) = pa(2)
   m_MySt(3) = pa(3)
   m_MySt(4) = pa(4)
   m_SysKind = CheckSys(m_MySt(1))
   SetLetterSt
   
   '取得樣本檔
   Select Case cp(10)
      Case 發明申請
         m_FileName = "$$發明申請_樣本.doc"
         Call PUB_GetSampleFile(m_FileName, "M51-000200-0-02", bolFileExist)
         If bolFileExist = True Then MsgBox "請將已開啟的申請書檔案，關閉後再重新執行！": Exit Function 'Add By Sindy 2015/12/14
         'Modified by Lydia 2018/08/08 +外文本3項
         'int_TotCnt = 12
         int_TotCnt = 15
         '取得發明人資料
         strApplInventor = PUB_GetApplInventor(pa(1), pa(2), pa(3), pa(4), "發明人", intInvCnt)
      Case 新型申請
         m_FileName = "$$新型申請_樣本.doc"
         Call PUB_GetSampleFile(m_FileName, "M51-000200-0-03", bolFileExist)
         If bolFileExist = True Then MsgBox "請將已開啟的申請書檔案，關閉後再重新執行！": Exit Function 'Add By Sindy 2015/12/14
         'Modified by Lydia 2018/08/08 +外文本3項
         'int_TotCnt = 11
         int_TotCnt = 14
         '取得發明人資料
         strApplInventor = PUB_GetApplInventor(pa(1), pa(2), pa(3), pa(4), "新型創作人", intInvCnt)
      Case 設計申請
         m_FileName = "$$設計申請_樣本.doc"
         Call PUB_GetSampleFile(m_FileName, "M51-000200-0-04", bolFileExist)
         If bolFileExist = True Then MsgBox "請將已開啟的申請書檔案，關閉後再重新執行！": Exit Function 'Add By Sindy 2015/12/14
         'Modified by Lydia 2018/08/08 +外文本3項
         'int_TotCnt = 11
         int_TotCnt = 14
         '取得發明人資料
         strApplInventor = PUB_GetApplInventor(pa(1), pa(2), pa(3), pa(4), "設計人", intInvCnt)
   End Select

   If Dir(App.path & "\" & m_FileName) <> "" Then
      Screen.MousePointer = vbHourglass
      '判斷word是否已開啟
      If g_WordAp Is Nothing Then
RestarWord:
         Set g_WordAp = New Word.Application
         g_WordAp.Visible = True 'False
      End If
'         If Dir(PUB_Getdesktop & "\" & m_TempFileName) <> "" Then
'            Kill PUB_Getdesktop & "\" & m_TempFileName
'         End If
      g_WordAp.Documents.Open App.path & "\" & m_FileName
'         g_WordAp.ActiveDocument.SaveAs PUB_Getdesktop & "\" & m_TempFileName
'         g_WordAp.ActiveDocument.Close
'         g_WordAp.Documents.Open PUB_Getdesktop & "\" & m_TempFileName
      With g_WordAp
         .Selection.WholeStory
         .Selection.Copy
         
         '取得申請人資料
         strApplData = PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), , , , , , strApplLineText, intAppCnt, intAppCnt2, cp(10))
         '取得優先權資料
         strApplPriDate = PUB_GetApplPriDate(pa(1), pa(2), pa(3), pa(4), pa(8))
         
         For i = 0 To int_TotCnt
            strName = ""
            strText = ""
            strLineText = ""
            Select Case cp(10)
               Case 發明申請
                  If i = 0 Then
                     strName = "一併實審"
                     If Text7 = "Y" Then
                        strText = "■本案一併申請實體審查"
                     Else
                        strText = "□本案一併申請實體審查"
                     End If
                  ElseIf i = 1 Then
                     strName = "案件名稱"
                     strText = Text6(0) & IIf(Text6(0) <> "" And Text6(1) <> "", vbCrLf, "") & Text6(1)
                  ElseIf i = 2 Then
                     strName = "共幾人"
                     strText = intAppCnt
                  ElseIf i = 3 Then
                     strName = "申請人"
                     strText = strApplData
                     strLineText = strApplLineText
                  ElseIf i = 4 Then
                     strName = "出名代理人"
                     strText = PUB_GetAgentCP110(strReceiveNo)
                  ElseIf i = 5 Then
                     strName = "發文字號"
                     If cp(27) = "" Then
                        'Modify By Sindy 2018/5/28 敏莉說先預帶系統日
                        'strText = "發文字號： 　 年 　 月 　 日(" & Left(strSrvDate(2), 3) & ")"
                        strText = "發文字號： " & Left(strSrvDate(2), 3) & " 年 " & _
                                               Mid(strSrvDate(2), 4, 2) & " 月 " & _
                                               Right(strSrvDate(2), 2) & " 日(" & Left(strSrvDate(2), 3) & ")"
                     Else
                        strText = "發文字號： " & Val(Left(DBDATE(cp(27)), 4)) - 1911 & " 年 " & _
                                               Mid(DBDATE(cp(27)), 5, 2) & " 月 " & _
                                               Right(DBDATE(cp(27)), 2) & " 日(" & Left(DBDATE(cp(27)), 4) - 1911 & ")"
                     End If
                     strText = strText & "晉專外字第　　　　　　　號"
                  ElseIf i = 6 Then
                     strName = "發明人共幾人"
                     strText = intInvCnt
                  ElseIf i = 7 Then
                     strName = "發明人"
                     strText = strApplInventor
                  ElseIf i = 8 Then
'                     'Add By Sindy 2015/12/9 加實審規費
'                     strCP84 = ""
'                     'Modify By Sindy 2018/5/28 Mark:直接抓取規費
'                     If Text7 = "Y" Then
'                        strExc(0) = " SELECT cp84 FROM caseprogress WHERE CP01='" & pa(1) & "' and CP02='" & pa(2) & "' and CP03='" & pa(3) & "' and CP04='" & pa(4) & "' and CP10='416' and CP57 is null"
'                        intI = 1
'                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                        If intI = 1 Then
'                           strCP84 = "" & RsTemp.Fields("cp84")
'                        End If
'                     End If
'                     strCP84 = Val(strCP84) + Val(cp(84))
'                     '2018/5/28 END
'                     '2015/12/9
                     strName = "發文規費"
                     'strText = IIf(Val(strCP84) = 0, "　　　", Format(Val(strCP84), "#,##0"))
                     strText = IIf(Val(txtCP84) = 0, "　　　", Format(Val(txtCP84), "#,##0"))
                  ElseIf i = 9 Then
                     strName = "主張優先權"
                     If strApplPriDate = "" Then
                        strText = "□主張優先權："
                     Else
                        strText = "■主張優先權："
                     End If
                  ElseIf i = 10 Then
                     strName = "優先權資料"
                     strText = strApplPriDate
                  ElseIf i = 11 Then
                     strName = "一案兩請"
                     If chkAtt(11).Value = vbChecked Then
                        strText = "■聲明本人就相同創作在申請本發明專利之同日，另申請新型專利。"
                     Else
                        strText = "□聲明本人就相同創作在申請本發明專利之同日，另申請新型專利。"
                     End If
                  'Add By Sindy 2015/12/9
                  ElseIf i = 12 Then
                     strName = "案號"
                     strText = pa(2)
                  '2015/12/9 END
                  'Added by Lydia 2018/08/08
                  ElseIf i = 13 Then
                     strName = "外文本1"
                     strText = IIf(Trim(txtTF24) <> "", " " & Val(txtTF24) & " ", "    ")
                  ElseIf i = 14 Then
                     strName = "外文本2"
                     strText = IIf(Trim(txtTF25) <> "", " " & Val(txtTF25) & " ", "    ")
                  ElseIf i = 15 Then
                     strName = "外文本3"
                     strText = IIf(Trim(txtForeign) <> "", " " & Val(txtForeign) & " ", "    ")
                  'end 2018/08/08
                  End If
                  
               Case 新型申請
                  If i = 0 Then
                     strName = "案件名稱"
                     strText = Text6(0) & IIf(Text6(0) <> "" And Text6(1) <> "", vbCrLf, "") & Text6(1)
                  ElseIf i = 1 Then
                     strName = "共幾人"
                     strText = intAppCnt
                  ElseIf i = 2 Then
                     strName = "申請人"
                     strText = strApplData
                     strLineText = strApplLineText
                  ElseIf i = 3 Then
                     strName = "出名代理人"
                     strText = PUB_GetAgentCP110(strReceiveNo)
                  ElseIf i = 4 Then
                     strName = "發文字號"
'                     If cp(27) = "" Then
'                        strText = "發文字號： 　 年 　 月 　 日(" & Left(strSrvDate(2), 3) & ")"
'                     Else
'                        strText = "發文字號： " & Val(Left(DBDATE(cp(27)), 4)) - 1911 & " 年 " & _
'                                               Mid(DBDATE(cp(27)), 5, 2) & " 月 " & _
'                                               Right(DBDATE(cp(27)), 2) & " 日(" & Left(DBDATE(cp(27)), 4) - 1911 & ")"
'                     End If
                     If cp(27) = "" Then
                        'Modify By Sindy 2018/5/28 敏莉說先預帶系統日
                        'strText = "發文字號： 　 年 　 月 　 日(" & Left(strSrvDate(2), 3) & ")"
                        strText = "發文字號： " & Left(strSrvDate(2), 3) & " 年 " & _
                                               Mid(strSrvDate(2), 4, 2) & " 月 " & _
                                               Right(strSrvDate(2), 2) & " 日(" & Left(strSrvDate(2), 3) & ")"
                     Else
                        strText = "發文字號： " & Val(Left(DBDATE(cp(27)), 4)) - 1911 & " 年 " & _
                                               Mid(DBDATE(cp(27)), 5, 2) & " 月 " & _
                                               Right(DBDATE(cp(27)), 2) & " 日(" & Left(DBDATE(cp(27)), 4) - 1911 & ")"
                     End If
                     strText = strText & "晉專外字第　　　　　　　號"
                  ElseIf i = 5 Then
                     strName = "發明人共幾人"
                     strText = intInvCnt
                  ElseIf i = 6 Then
                     strName = "發明人"
                     strText = strApplInventor
                  ElseIf i = 7 Then
                     strName = "發文規費"
                     'strText = IIf(cp(84) = "", "　　　", Format(cp(84), "#,##0"))
                     strText = IIf(Val(txtCP84) = 0, "　　　", Format(Val(txtCP84), "#,##0"))
                  ElseIf i = 8 Then
                     strName = "主張優先權"
                     If strApplPriDate = "" Then
                        strText = "□主張優先權："
                     Else
                        strText = "■主張優先權："
                     End If
                  ElseIf i = 9 Then
                     strName = "優先權資料"
                     strText = strApplPriDate
                  ElseIf i = 10 Then
                     strName = "一案兩請"
                     If chkAtt(11).Value = vbChecked Then
                        strText = "■聲明本人就相同創作在申請本新型專利之同日，另申請發明專利。"
                     Else
                        strText = "□聲明本人就相同創作在申請本新型專利之同日，另申請發明專利。"
                     End If
                  'Add By Sindy 2015/12/9
                  ElseIf i = 11 Then
                     strName = "案號"
                     strText = pa(2)
                  '2015/12/9 END
                  'Added by Lydia 2018/08/08
                  ElseIf i = 12 Then
                     strName = "外文本1"
                     strText = IIf(Trim(txtTF24) <> "", " " & Val(txtTF24) & " ", "    ")
                  ElseIf i = 13 Then
                     strName = "外文本2"
                     strText = IIf(Trim(txtTF25) <> "", " " & Val(txtTF25) & " ", "    ")
                  ElseIf i = 14 Then
                     strName = "外文本3"
                     strText = IIf(Trim(txtForeign) <> "", " " & Val(txtForeign) & " ", "    ")
                  'end 2018/08/08
                  End If
                  
               Case 設計申請
                  If i = 0 Then
                     strName = "設計種類"
                     If pa(158) = "1" Then
                        strText = "設計種類：■整體□部分□圖像□成組"
                     ElseIf pa(158) = "2" Then
                        strText = "設計種類：□整體■部分□圖像□成組"
                     ElseIf pa(158) = "3" Then
                        strText = "設計種類：□整體□部分■圖像□成組"
                     ElseIf pa(158) = "4" Then
                        strText = "設計種類：□整體□部分□圖像■成組"
                     Else
                        strText = "設計種類：□整體□部分□圖像□成組"
                     End If
                  ElseIf i = 1 Then
                     strName = "案件名稱"
                     strText = Text6(0) & IIf(Text6(0) <> "" And Text6(1) <> "", vbCrLf, "") & Text6(1)
                  ElseIf i = 2 Then
                     strName = "共幾人"
                     strText = intAppCnt
                  ElseIf i = 3 Then
                     strName = "申請人"
                     strText = strApplData
                     strLineText = strApplLineText
                  ElseIf i = 4 Then
                     strName = "出名代理人"
                     strText = PUB_GetAgentCP110(strReceiveNo)
                  ElseIf i = 5 Then
                     strName = "發文字號"
'                     If cp(27) = "" Then
'                        strText = "發文字號： 　 年 　 月 　 日(" & Left(strSrvDate(2), 3) & ")"
'                     Else
'                        strText = "發文字號： " & Val(Left(DBDATE(cp(27)), 4)) - 1911 & " 年 " & _
'                                               Mid(DBDATE(cp(27)), 5, 2) & " 月 " & _
'                                               Right(DBDATE(cp(27)), 2) & " 日(" & Left(DBDATE(cp(27)), 4) - 1911 & ")"
'                     End If
                     If cp(27) = "" Then
                        'Modify By Sindy 2018/5/28 敏莉說先預帶系統日
                        'strText = "發文字號： 　 年 　 月 　 日(" & Left(strSrvDate(2), 3) & ")"
                        strText = "發文字號： " & Left(strSrvDate(2), 3) & " 年 " & _
                                               Mid(strSrvDate(2), 4, 2) & " 月 " & _
                                               Right(strSrvDate(2), 2) & " 日(" & Left(strSrvDate(2), 3) & ")"
                     Else
                        strText = "發文字號： " & Val(Left(DBDATE(cp(27)), 4)) - 1911 & " 年 " & _
                                               Mid(DBDATE(cp(27)), 5, 2) & " 月 " & _
                                               Right(DBDATE(cp(27)), 2) & " 日(" & Left(DBDATE(cp(27)), 4) - 1911 & ")"
                     End If
                     strText = strText & "晉專外字第　　　　　　　號"
                  ElseIf i = 6 Then
                     strName = "發明人共幾人"
                     strText = intInvCnt
                  ElseIf i = 7 Then
                     strName = "發明人"
                     strText = strApplInventor
                  ElseIf i = 8 Then
                     strName = "發文規費"
                     'strText = IIf(cp(84) = "", "　　　", Format(cp(84), "#,##0"))
                     strText = IIf(Val(txtCP84) = 0, "　　　", Format(Val(txtCP84), "#,##0"))
                  ElseIf i = 9 Then
                     strName = "主張優先權"
                     If strApplPriDate = "" Then
                        strText = "□主張優先權："
                     Else
                        strText = "■主張優先權："
                     End If
                  ElseIf i = 10 Then
                     strName = "優先權資料"
                     strText = strApplPriDate
                  'Add By Sindy 2015/12/9
                  ElseIf i = 11 Then
                     strName = "案號"
                     strText = pa(2)
                  '2015/12/9 END
                  'Added by Lydia 2018/08/08
                  ElseIf i = 12 Then
                     strName = "外文本1"
                     strText = IIf(Trim(txtTF24) <> "", " " & Val(txtTF24) & " ", "    ")
                  ElseIf i = 13 Then
                     strName = "外文本2"
                     strText = IIf(Trim(txtTF25) <> "", " " & Val(txtTF25) & " ", "    ")
                  ElseIf i = 14 Then
                     strName = "外文本3"
                     strText = IIf(Trim(txtForeign) <> "", " " & Val(txtForeign) & " ", "    ")
                  'end 2018/08/08
                  End If
            End Select
            'Find並且置換
            If Trim(strName) <> "" Then
               .Selection.Find.ClearFormatting
               .Selection.Find.Text = "|#" & strName & "#|"
               .Selection.Find.Replacement.Text = ""
               .Selection.Find.Forward = True
               .Selection.Find.Wrap = wdFindContinue
               .Selection.Find.Format = False
               .Selection.Find.MatchCase = False
               .Selection.Find.MatchWholeWord = False
               .Selection.Find.MatchWildcards = False
               .Selection.Find.MatchSoundsLike = False
               .Selection.Find.MatchAllWordForms = False
               .Selection.Find.MatchByte = True
               .Selection.Find.Execute
               .Selection.Delete
               .Selection.TypeText strText
               If strLineText <> "" Then
                  .Selection.HomeKey
                  .Selection.Find.ClearFormatting
                  With .Selection.Find
                      .Text = strLineText
                      .Replacement.Text = ""
                      .Forward = True
                      .Wrap = wdFindContinue
                      .Format = False
                      .MatchCase = False
                      .MatchWholeWord = False
                      .MatchWildcards = False
                      .MatchSoundsLike = False
                      .MatchAllWordForms = False
                      .MatchByte = True
                  End With
                  .Selection.Find.Execute
                  .Selection.Font.Underline = wdUnderlineSingle
               End If
               'Modify By Sindy 2016/1/11
               If InStr("出名代理人;優先權資料", strName) > 0 Then
               '2016/1/11 END
                  ChgWordFormat g_WordAp.Application, strText
               End If
            End If
ReadNext:
         Next i
         '功能變數更新
         '.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '頁首
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter '頁尾
         '.NormalTemplate.AutoTextEntries("第 X 頁，共 Y 頁").Insert Where:=Selection.Range
         '.Selection.Fields.ToggleShowCodes
         .Selection.MoveRight Unit:=wdCharacter, Count:=8
         .Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
         .Selection.Fields.UPDATE
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
      End With
      Screen.MousePointer = vbDefault
'         g_WordAp.ActiveDocument.Save
'         g_WordAp.ActiveDocument.Close
'         MsgBox "檔案已存放在：" & PUB_Getdesktop & "\" & m_TempFileName
      MsgBox "資料已產生完畢!!!"
      GetApplBook = True
   Else
      MsgBox "無申請書的樣本!!!"
   End If

   Exit Function
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   End If
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function

'Private Function ConvertNameFormat(pName As String) As String
'   Dim strTmp As String
'
'   If InStr(pName, ",") = 0 Then
'      strTmp = Left(pName, 1) & "," & Mid(pName, 2)
'   Else
'      strTmp = pName
'   End If
'   ConvertNameFormat = strTmp
'End Function

'申請書
Private Sub StartLetter2(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(110) As String, strTmp As String
   Dim ii As Integer, jj As Integer
   Dim strInventor As String 'Add By Sindy 2014/11/14
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   'Add By Sindy 2019/10/18
   '一併提實審也要加.2
   If Text7 = "Y" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','電子稽核防止漏發文','♀')"
   End If
   '2019/10/18 END
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','是否一併提實審','" & IIf(Text7 = "Y", "是", "否") & "')"
      
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
               
   'Added by Morgan 2018/1/12
   If pa(8) = "3" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','設計種類','" & PUB_GetCaseAttributeName(pa(158), pa(8)) & "')"
   End If
   If chkAtt(11).Value = vbChecked Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一案兩請','♀')"
   End If
   'end 2018/1/12
   
   'Modify By Sindy 2017/11/15
   Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa())
'   '申請人
'   For jj = 1 To 5
'      If pa(25 + jj) <> "" Then
'         '申請人
'         strExc(0) = " SELECT C.*,N1.NA72 X1,N2.NA72 X2" & _
'            " FROM CUSTOMER C,NATION N1,NATION N2 WHERE CU01='" & Left(ChangeCustomerL(pa(25 + jj)), 8) & "'" & _
'            " and cu02='" & Mid(ChangeCustomerL(pa(25 + jj)), 9) & "' AND N1.NA01(+)=CU10 AND N2.NA01(+)=CU87"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-國籍','" & RsTemp("X1") & "')"
'
'            ii = ii + 1
'            If RsTemp("CU10") < "011" Then
'               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-ID','" & RsTemp("CU11") & "')"
'            End If
'
'            ii = ii + 1
'            If RsTemp("CU15") = "0" Then
'               strTmp = "申請人" & jj & "-中文姓名"
'            Else
'               strTmp = "申請人" & jj & "-中文名稱"
'            End If
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL("" & RsTemp("CU04")) & "')"
'
'            ii = ii + 1
'            If RsTemp("CU15") = "0" Then
'               strTmp = "申請人" & jj & "-英文姓名"
'            Else
'               strTmp = "申請人" & jj & "-英文名稱"
'            End If
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL(RTrim(Trim("" & RsTemp("CU05")) & " " & Trim("" & RsTemp("CU88")) & " " & Trim("" & RsTemp("CU89")) & " " & Trim("" & RsTemp("CU90")))) & "')"
'         End If
'      End If
'   Next
   
   '出名代理人
   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      jj = 1
      Do While Not .EOF
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
         jj = jj + 1
         .MoveNext
      Loop
      End With
   End If
   
   'Add By Sindy 2018/12/22
   '相關卷號檔
   If cp(10) = "125" Then
      strExc(0) = "select cr05,cr06,cr07,cr08,pa11 from caserelation1,patent where cr01='" & pa(1) & "' and cr02='" & pa(2) & "' and cr03='" & pa(3) & "' and cr04='" & pa(4) & "'" & _
                  " and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','原設計申請案號','" & RsTemp.Fields("pa11") & "')"
      End If
      
   'Add By Sindy 2022/5/3
   ElseIf cp(10) = 改請衍生設計 Then
      strExc(0) = "select pa01,pa02,pa03,pa04,pa11 from patent where pa01||pa02||pa03||pa04='" & cp(30) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','原設計申請案號','" & RsTemp.Fields("pa11") & "')"
      End If
      '若有收文分割且未發文則案號後面加.2
      If PUB_ChkCPExist(pa, "307", 1) = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','電子稽核防止漏發文','♀')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','分割同時改請','是')"
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','分割同時改請','否')"
      End If
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本案符合優惠期相關規定','" & IIf(Chk1.Value = 1, "是", "否") & "')"
      If Chk1.Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','列印優惠期','♀')"
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','援用原申請案優先權主張','" & IIf(Chk2.Value = 1, "是", "否") & "')"
   End If
   '2018/12/22 END
   
'   '讀取發明人資料
'   'Modify By Sindy 2014/11/14
'   If pa(8) = "1" Then
'      strExc(1) = "發明人"
'   ElseIf pa(8) = "2" Then
'      strExc(1) = "新型創作人"
'   Else
'      strExc(1) = "設計人"
'   End If
''   For jj = 1 To 10
''      If pa(59 + jj) <> "" Then
''         strExc(0) = " SELECT IN03,IN04,IN05,IN11,NA72" & _
''            " FROM INVENTOR,NATION WHERE IN01='" & Left(pa(59 + jj), 8) & "' AND IN02='" & Mid(pa(59 + jj), 9) & "'" & _
''            " AND NA01(+)=IN11"
''         intI = 1
''         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''         If intI = 1 Then
''            ii = ii + 1
''            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人" & jj & "-國籍','" & RsTemp("NA72") & "')"
''            ii = ii + 1
''            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人" & jj & "-中文姓名','" & ChgSQL("" & RsTemp("IN04")) & "')"
''
''            ii = ii + 1
''            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人" & jj & "-英文姓名','" & ChgSQL("" & RsTemp("IN05")) & "')"
''         End If
''      End If
''   Next
'   strInventor = ""
'   strExc(0) = " SELECT IN03,IN04,IN05,IN11,NA72" & _
'               " FROM PatentInventor,INVENTOR,NATION" & _
'               " WHERE pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
'               " AND IN01=substr(pi06,1,8) AND IN02=substr(pi06,9,2)" & _
'               " AND NA01(+)=IN11" & _
'               " order by pi05 asc"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   'Modified by Morgan 2018/1/8 發明人TAG後面加序號,取消內縮
'   If intI = 1 Then
'      RsTemp.MoveFirst
'      Do While Not RsTemp.EOF
'         If strInventor <> "" Then strInventor = strInventor & vbCrLf
'         'Modify By Sindy 2018/10/25 增加英文名稱格式化 PUB_FCPIN05Format_EName
'         strInventor = strInventor & "【" & strExc(1) & intI & "】" & vbCrLf & _
'                                     "　　【國籍】　　　　　　　　　" & RsTemp("NA72") & vbCrLf & _
'                                     "　　【中文姓名】　　　　　　　" & ChgSQL("" & RsTemp("IN04")) & vbCrLf & _
'                                     IIf("" & RsTemp("IN05") = "", "", "　　【英文姓名】　　　　　　　" & ChgSQL(PUB_FCPIN05Format_EName("" & RsTemp("IN05"), "" & RsTemp("NA72"))) & vbCrLf)
'         RsTemp.MoveNext
'         intI = intI + 1
'      Loop
'   Else
'      strInventor = "【" & strExc(1) & "1】" & vbCrLf & _
'                    "　　【國籍】　　　　　　　　　" & vbCrLf & _
'                    "　　【中文姓名】　　　　　　　" & vbCrLf
'   End If
'   If Not (pa(1) = "FCP" And InStr("101,102,103,307", cp(10)) > 0) Then  'Added by Lydia 2022/03/03 排除FCP發明申請書，改用<申請書發明人資料>；ex.FCP-066642發明申請書的發明人資料超過4000字
'        ii = ii + 1
'        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'           " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人資料','" & strInventor & "')"
'        '2014/11/14 END
'   End If

   '優惠期發生日期
   If txtFavDate <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優惠期發生日期','" & ChangeTStringToWDateString(txtFavDate) & "')"
      
      Select Case cboFavReason.ListIndex
      Case 0: strTmp = "優惠期原因-實驗"
      Case 1: strTmp = "優惠期原因-刊物"
      Case 2: strTmp = "優惠期原因-展覽"
      End Select
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','♀')"
   End If
   
   '優先權資料
   'Modify By Sindy 2019/8/8 改共用函數,不限數量
   strTmp = PUB_GetAppPridate(pa, ET01, strReceiveNo, ET03)
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優先權資料','" & strTmp & "')"
'   'Modify By Sindy 2018/7/25 + order by pd05 asc
'   strExc(0) = "SELECT sqldatew(pd05) pd05,na72,pd06,pd07,decode(pd08,'1','發明','2','新型','3','設計',pd08) pd08,pd09" & _
'      " FROM pridate,nation where pd01='" & pa(1) & "' and pd02='" & pa(2) & "' and pd03='" & pa(3) & "' and pd04='" & pa(4) & "'" & _
'      " and na01(+)=pd07" & _
'      " order by pd05 asc"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      jj = 0
'      Do While Not RsTemp.EOF
'         jj = jj + 1
'         If jj > 10 Then
'            MsgBox "優先權資料超過 10 筆，請自行維護！"
'            Exit Do
'         End If
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-日','" & RsTemp("pd05") & "')"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-國','" & RsTemp("na72") & "')"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-號','" & RsTemp("pd06") & "')"
'         'Added by Morgan 2017/8/9
'         'Modify By Sindy 2017/11/9 輸入優先權國家代碼時,代表是以電子交換檢送
'         'Modify By Sindy 2018/8/10 ex:FCP-59317 韓國,分交換和不交換
'         If RsTemp("pd07") = "012" Then '韓國
'            If RsTemp("pd07") = "" & RsTemp("pd09") Then
'         '2017/11/9 END
'               '電子交換
'               ii = ii + 1
'               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-碼','交換')"
'            Else
'               '非電子交換
'               ii = ii + 1
'               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-碼','不交換')"
'            End If
'            '2018/8/10 END
'         ElseIf Not IsNull(RsTemp("pd09")) Then
'            '非電子交換
'            'Add By Sindy 2017/11/9
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-種類','" & IIf("" & RsTemp("pd08") = "", "♀", ChgSQL("" & RsTemp("pd08"))) & "')"
'            '2017/11/9 END
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-碼','" & ChgSQL(RsTemp("pd09")) & "')"
'         End If
'         'Removed by Morgan 2018/1/12 移到下面,改只要印一個,因為不一定都有,程序自己改檔名--敏莉
'         'If chkAtt(12).Value = 1 Then
'         '   ii = ii + 1
'         '   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         '      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-國際優先權證明文件" & jj & "','" & m_CaseNo & "Priority_" & Left(RsTemp("na72"), 2) & ".pdf')"
'         'End If
'         RsTemp.MoveNext
'      Loop
'   End If
   
   If chkBiology(0).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','寄存國家','" & lblCountry & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','寄存機構','" & txtBiology(1) & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','寄存日期','" & ChangeWStringToWDateString(DBDATE(txtBiology(2))) & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','寄存號碼','" & txtBiology(3) & "')"
   End If
   
   If chkDoc(0).Value = 1 Then
      If txtDocCh(0).Visible Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','摘要頁數','" & Val(txtDocCh(0)) & "')"
      End If
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','說明書頁數','" & Val(txtDocCh(1)) & "')"
      
      If txtDocCh(2).Visible Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利範圍頁數','" & Val(txtDocCh(2)) & "')"
      End If
      If Val(txtDocCh(3)) > 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','圖式頁數','" & Val(txtDocCh(3)) & "')"
      End If
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','頁數總計','" & Val(txtDocCh(4)) & "')"
      
      If txtDocCh(5).Visible Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','專利範圍項數','" & Val(txtDocCh(5)) & "')"
      End If
      If Val(txtDocCh(6)) > 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','圖式圖數','" & Val(txtDocCh(6)) & "')"
      End If
   End If
   
   If chkDoc(1).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','外文頁數總計','" & Val(txtForeign) & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','外文本種類','" & cboLagnuage & "')"
   End If
   
   If chkDoc(2).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','簡體字頁數總計','" & Val(txtSimplified) & "')"
   End If
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
      
   '附件
   'Modify By Sindy 2018/6/28 依案件性質帶其副檔名 IIf(m_CPM26 <> "", "." & UCase(m_CPM26), ".INV")
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & ".contact.pdf')"
   If chkAtt(1).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-摘要','" & m_CaseNo & IIf(m_CPM26 <> "", "." & UCase(m_CPM26), ".INV") & "_ABSTRACT.pdf')"
   End If
   If chkAtt(2).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-說明書','" & m_CaseNo & IIf(m_CPM26 <> "", "." & UCase(m_CPM26), ".INV") & "_DESCRIPTION.pdf')"
   End If
   'Add By Sindy 2020/2/26 序列表頁數,有輸入就要顯示於附送書件
   If Val(txtDocCh(7).Text) > 0 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-中文本-序列表','" & m_CaseNo & ".SEQ.pdf')"
      txtMemo = txtMemo & "序列表" & Val(txtDocCh(7).Text) & "頁不算超頁費。"
   End If
   '2020/2/26 END
   If chkAtt(3).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-專利範圍','" & m_CaseNo & IIf(m_CPM26 <> "", "." & UCase(m_CPM26), ".INV") & "_CLAIMS.pdf')"
   End If
   If chkAtt(4).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-圖式','" & m_CaseNo & IIf(m_CPM26 <> "", "." & UCase(m_CPM26), ".INV") & "_DRAWINGS.pdf')"
   End If
   If chkAtt(5).Value = 1 Then
      ii = ii + 1
      'Modified by Morgan 2017/8/9 Foreign.pdf->ForeignSpec.pdf--葉敏莉
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-外文本','" & m_CaseNo & ".ORI.pdf')"
   End If
   'Added by Morgan 2017/11/14 +外文圖式及序列表--葉敏莉
   If chkAtt(13).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-外文本-圖式','')"
   End If
   If chkAtt(14).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-外文本-序列表','" & m_CaseNo & ".SEQ.pdf')"
   End If
   'end 2017/11/14
   If chkAtt(6).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-簡體字本','" & m_CaseNo & ".ORI.pdf')" 'Modify By Sindy 2018/6/5 敏莉提:SEP==>ORI
   End If
      
   'Added by Morgan 2018/1/12
   If chkAtt(12).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-國際優先權證明文件','" & m_CaseNo & ".PRI.pdf')"
   End If
   If chkAtt(15).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-優惠期證明文件','" & m_CaseNo & ".EXHIBITION.pdf')"
   End If
   'end 2018/1/12
   
   If chkAtt(7).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-委任書','" & m_CaseNo & ".POA.pdf')"
   End If
   If chkAtt(8).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-國內寄存證明','" & m_CaseNo & ".DOMESTICPROOF.pdf')"
   End If
   If chkAtt(9).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-國外寄存證明','" & m_CaseNo & ".FOREIGNPROOF.pdf')"
   End If
   If chkAtt(10).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-易於獲得證明','" & m_CaseNo & ".EASILYOBTAINED.pdf')"
   End If
   
   If chkAtt(16).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-其他','♀')"
   End If
   
   'Added by Morgan 2018/2/23
   'Modify By Sindy 2022/5/4 + Or cp(10) = 改請衍生設計
   If txtMemo <> "" Or cp(10) = 改請衍生設計 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & ChgSQL(txtMemo) & "')"
   End If
   'end 2018/2/23
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Private Sub Combo1_Click(Index As Integer)
'   ChgType Index + 10
'End Sub

Private Sub Combo2_Click(Index As Integer)
 Dim i As Integer, strTmp As String
   Select Case Index
      Case 0, 1
         If Combo2(Index) = "" Then
            For i = 0 To 2
               Text6(i + (Index + 1) * 3) = ""
            Next
            Exit Sub
         End If
      Case Else
         If Combo2(Index) = "" Then
            For i = 0 To 2
               Text6(i + (Index - 2) * 3 + 24) = ""
            Next
            Exit Sub
         End If
   End Select
   
   strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
   strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
   strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      For i = 0 To 2
         Select Case Index
            Case 0, 1
               If Not IsNull(RsTemp.Fields(i)) Then
                  Text6(i + (Index + 1) * 3) = RsTemp.Fields(i)
               Else
                  Text6(i + (Index + 1) * 3) = ""
               End If
            Case Else
               If Not IsNull(RsTemp.Fields(i)) Then
                  Text6(i + (Index - 2) * 3 + 24) = RsTemp.Fields(i)
               Else
                  Text6(i + (Index - 2) * 3 + 24) = ""
               End If
         End Select
      Next
   End If
End Sub

'Private Function ChgType(i As Integer) As Boolean
'   Dim strTxt(1 To 5) As String
'   ChgType = False
'   Select Case i
'      Case 10, 11, 12, 13, 14, 15, 16, 17, 18, 19
'         If ClsLawGetInventor(Replace(Right(Combo1(i - 10).Text, 11), ")", ""), strTxt) = True Then
'            Label4(i - 10) = strTxt(1)
'            Label6(i - 10) = strTxt(2)
'            Label8(i - 10) = strTxt(3)
'            ChgType = True
'         Else
'            Label4(i - 10) = ""
'            Label6(i - 10) = ""
'            Label8(i - 10) = ""
'         End If
'   End Select
'End Function

'************************************************
' 儲存專利案件資料
'
'************************************************
Private Function FormSave() As Boolean
Dim ii As Integer, m_CP110 As String
Dim strCon As String 'Add By Sindy 2018/2/14
   
'   cp(110) = ""
'   For ii = 0 To lstNameAgent.ListCount - 1
'      If lstNameAgent.Selected(ii) = True Then
'         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
'         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
'         cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
'      End If
'   Next
'   If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   
   'Add By Sindy 2018/2/14
   If m_PrevForm_Text6 = "3" Then '電子送件
      strCon = strCon & ",cp118='A'"
'      'Add By Sindy 2018/6/15 一併提實審時,存頁/項數
      'If Text7 = "Y" And chkDoc(0).Value = 1 Then
      If chkDoc(0).Value = 1 Then 'Modify By Sindy 2018/6/28
         strCon = strCon & ",cp135=" & Val(txtDocCh(4))
         strCon = strCon & ",cp136=" & Val(txtDocCh(5))
      End If
      '2018/6/15 END
   Else
      strCon = strCon & ",cp118=null"
'      'Add By Sindy 2018/6/15 一併提實審時,存頁/項數
'      If Text7 = "Y" Then
         strCon = strCon & ",cp135=" & Val(txtCP135)
         strCon = strCon & ",cp136=" & Val(txtCP136)
'      End If
      '2018/6/15 END
   End If
   '2018/2/14 END
   strCon = strCon & ",cp84=" & Val(txtCP84)
   
   'Add By Sindy 2019/2/13
   If chkEexcerpt.Value = 1 And InStr(cp(64), "減收申請規費800元") = 0 Then
      strCon = strCon & ",cp64='" & "未附英文說明書，所檢附之申請書及摘要附有英文翻譯，可減收申請規費800元。" & cp(64) & "'"
   End If
   'Add By Sindy 2018/2/14
   If lstNameAgent.Visible = True Then
      strCon = strCon & ",cp110='" & cp(110) & "'"
   End If
   
   cnnConnection.BeginTrans
   
On Error GoTo CheckingErr
      
   If strCon <> "" Then
      strCon = Mid(strCon, 2)
   '2018/2/14 END
      strSql = " UPDATE CASEPROGRESS SET " & strCon & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2023/3/24
   '若產生新案電子送件申請書時，有勾選: 未附英文說明書，但可減收申請費800元
   '按確定後，請在（201）新案翻譯、（209）檢視中說、（235）核對中說格式的進度檔加上備註：中說需附英文摘要
   If chkEexcerpt.Value = 1 Then
      strSql = " UPDATE CASEPROGRESS SET CP64=CP64||'中說需附英文摘要;'" & _
               " WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
               " AND CP10 in('201','209','235')"
      cnnConnection.Execute strSql
   End If
   '2023/3/24 END
   
   If UCase(TypeName(m_PrevForm)) <> UCase("frm090904") Then '工程師各式申請書
      'Added by Lydia 2018/08/08 回寫外文本說明書和圖示頁數
      If m_TF01 <> "" And (Val(txtTF24.Tag) <> Val(txtTF24.Text) Or Val(txtTF25.Tag) <> Val(txtTF25.Text)) Then
            strSql = "Update TransFee set TF24=" & CNULL(txtTF24.Text, True) & ", TF25=" & CNULL(txtTF25.Text, True) & " where TF01='" & m_TF01 & "' "
            cnnConnection.Execute strSql, intI
            If intI = 0 Then
                strSql = "INSERT INTO TransFee (TF01,TF24,TF25) VALUES ('" & m_TF01 & "' , " & CNULL(txtTF24.Text, True) & " , " & CNULL(txtTF25.Text, True) & ") "
                cnnConnection.Execute strSql, intI
            End If
      End If
      'end 2018/08/08
   End If
   
   'Added by Lydia 2018/12/27 存中文本資訊
   strSql = ""
   For Each oText In txtDocCh
      'Modified by Lydi 2019/01/10
      'If (oText.Index <= 3 Or oText.Index = 7) And oText.Tag <> oText.Text Then
      If (oText.Index <> 4) And oText.Tag <> oText.Text Then
          Select Case oText.Index
               Case 0 '摘要頁數
                    strSql = strSql & ", PA64=" & CNULL(oText.Text, True)
               Case 1 '說明書頁數
                    strSql = strSql & ", PA65=" & CNULL(oText.Text, True)
               Case 7 '序列表頁數
                    strSql = strSql & ", PA66=" & CNULL(oText.Text, True)
               Case 2 '申請專利範圍頁數
                    strSql = strSql & ", PA67=" & CNULL(oText.Text, True)
               Case 3 '圖式頁數
                    strSql = strSql & ", PA68=" & CNULL(oText.Text, True)
               'Added by Lydia 2019/01/10
               Case 5 '申請專利範圍項數(最初項數)
                    strSql = strSql & ", PA172=" & CNULL(oText.Text, True)
               Case 6 '圖式圖數
                    strSql = strSql & ", PA173=" & CNULL(oText.Text, True)
          End Select
      End If
   Next
   If strSql <> "" Then
        strSql = "UPDATE PATENT SET " & Mid(strSql, 2) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        strSql = "begin user_data.user_enabled:=0; " & strSql & "; end;"
        'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
        'Pub_SeekTbLog strSql '新增log
        Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
        cnnConnection.Execute strSql
   End If
   'end 2018/12/27
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   
End Function

'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   cp(110) = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

Private Sub Text6_GotFocus(Index As Integer)
   InverseTextBox Text6(Index)
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   
   If Text6(0) = "" Then MsgBox "中文案件名稱不可空白!", vbCritical: Exit Function
   
   'Added by Morgan 2014/1/16 發明人沒有英文時提醒--Elvan
   'Memo by Lydia 2021/08/17 刪除舊程式碼：專利發明人在專利基本檔60~69
   'Modify By Sindy 2014/11/14
   strExc(0) = "select pi05,pi06 INV from patentinventor,inventor where pi01='" & pa(1) & "' and pi02='" & pa(2) & "'" & _
      " and pi03='" & pa(3) & "' and pi04='" & pa(4) & "' and in01(+)=substr(pi06,1,8) and in02(+)=substr(pi06,9) and in05 is null" & _
      " order by pi05 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modified by Morgan 2018/6/4
      'If MsgBox("【發明人" & RsTemp("Num") & "】尚缺【英文名稱】，是否要繼續??", vbYesNo + vbQuestion) = vbNo Then
      If MsgBox("【發明人" & RsTemp("INV") & "】尚缺【英文名稱】，是否要繼續??", vbYesNo + vbQuestion) = vbNo Then
         Exit Function
      End If
   End If
   'end 2014/1/16
   
   Cancel = False
   lstNameAgent_Validate Cancel
   If Cancel = True Then
      SSTab1.Tab = 0
      lstNameAgent.SetFocus
      Exit Function
   End If
      
   If txtFavDate <> "" Then
      txtFavDate_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         txtFavDate.SetFocus
         Exit Function
      Else
         If cboFavReason.ListIndex < 0 Then
            MsgBox "請點選優惠期發生原因！", vbInformation
            SSTab1.Tab = 0
            cboFavReason.SetFocus
            Exit Function
         End If
      End If
   End If
   
   If chkBiology(0).Value = vbChecked Then
      If txtBiology(0) = "" Then
         MsgBox "請輸入寄存國家！", vbInformation
         SSTab1.Tab = 1
         txtBiology(0).SetFocus
         Exit Function
      Else
         txtBiology_Validate 0, Cancel
         If Cancel = True Then
            SSTab1.Tab = 1
            txtBiology(0).SetFocus
            Exit Function
         End If
      End If
      If txtBiology(1) = "" Then
         MsgBox "請輸入寄存機構！", vbInformation
         SSTab1.Tab = 1
         txtBiology(1).SetFocus
         Exit Function
      End If
      If txtBiology(2) = "" Then
         MsgBox "請輸入寄存日期！", vbInformation
         SSTab1.Tab = 1
         txtBiology(2).SetFocus
         Exit Function
      Else
         txtBiology_Validate 2, Cancel
         If Cancel = True Then
            SSTab1.Tab = 1
            txtBiology(2).SetFocus
            Exit Function
         End If
      End If
      If txtBiology(3) = "" Then
         MsgBox "請輸入寄存號碼！", vbInformation
         SSTab1.Tab = 1
         txtBiology(3).SetFocus
         Exit Function
      End If
   End If
   
   If chkDoc(0).Value + chkDoc(1).Value + chkDoc(2).Value = 0 Then
      MsgBox "外文本、中文本或簡體字本資訊至少要選擇一種！", vbCritical
      'SSTab1.Tab = 1
      SSTab1.Tab = 0
      Exit Function
   End If
   
   If chkDoc(1).Value = vbChecked Then
      If Val(txtForeign) = 0 Then
         MsgBox "請輸入外文頁數總計！", vbInformation
         SSTab1.Tab = 0
         txtForeign.SetFocus
         Exit Function
      End If
      'Modified by Lydia 2023/05/26 改請衍生設計308顯示選項會遮蓋外文本種類
      If cboLagnuage.ListIndex < 0 And Frame3.Visible = False Then
         MsgBox "請點選外文本種類！", vbInformation
         'SSTab1.Tab = 1
         SSTab1.Tab = 0
         cboLagnuage.SetFocus
         Exit Function
      End If
   End If
   
   If chkDoc(0).Value = vbChecked Then
      If Val(txtDocCh(0)) = 0 And txtDocCh(0).Visible Then
         MsgBox "請輸入摘要頁數！", vbInformation
         SSTab1.Tab = 1
         txtDocCh(0).SetFocus
         Exit Function
      End If
      If Val(txtDocCh(1)) = 0 Then
         MsgBox "請輸入說明書頁數！", vbInformation
         SSTab1.Tab = 1
         txtDocCh(1).SetFocus
         Exit Function
      End If
      If Val(txtDocCh(2)) = 0 And txtDocCh(2).Visible Then
         MsgBox "請輸入申請專利範圍頁數！", vbInformation
         SSTab1.Tab = 1
         txtDocCh(2).SetFocus
         Exit Function
      End If
      If Val(txtDocCh(5)) = 0 And txtDocCh(5).Visible Then
         MsgBox "請輸入申請專利範圍項數！", vbInformation
         SSTab1.Tab = 1
         txtDocCh(5).SetFocus
         Exit Function
      End If
   End If
   
   If chkDoc(2).Value = vbChecked Then
      If Val(txtSimplified) = 0 Then
         MsgBox "請輸入簡體字頁數總計！", vbInformation
         SSTab1.Tab = 0
         txtSimplified.SetFocus
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

'Private Function GetInventorName(strIn As String) As String
'   Dim rsA  As New ADODB.Recordset
'   Dim StrSQLa As String
'
'   GetInventorName = ""
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   StrSQLa = "SELECT NVL(IN05,NVL(IN04,IN06))||'('||IN01||IN02||')' FROM INVENTOR WHERE IN01||IN02='" & strIn & "'"
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      If Not IsNull(rsA.Fields(0).Value) Then
'         GetInventorName = "" & rsA.Fields(0).Value
'      End If
'   End If
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'End Function

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   'Modify By Sindy 2018/5/28
   If m_PrevForm_Text6 = "3" Then '電子送件
      If cp(10) = "101" Then
         If Text7 = "Y" Then
            txtCP84 = "9900"
         Else
            txtCP84 = "2900"
         End If
      'Modify By Sindy 2022/5/3 + Or cp(10) = "308"
      ElseIf cp(10) = "102" Or cp(10) = "103" Or cp(10) = "125" Or cp(10) = "308" Then
         txtCP84 = "2400"
      End If
   Else
      If Text7 = "Y" Then
         txtCP84 = Val(cp(17)) + Val(7000)
         'Add By Sindy 2018/6/15
         txtCP135.Enabled = True
         txtCP136.Enabled = True
         '2018/6/15 END
      Else
         txtCP84 = Val(cp(17))
         'Add By Sindy 2018/6/15
         txtCP135.Enabled = False
         txtCP136.Enabled = False
         txtCP135.Text = ""
         txtCP136.Text = ""
         '2018/6/15 END
      End If
      txtCP84.Tag = cp(17)
   End If
   '2018/5/28 END
   'Add By Sindy 2019/10/24
   If Text7 = "Y" Then
      '超頁費
      If Val(txtDocCh(4)) > 50 Then
         m_lngOverPageFee = 500# * ((Val(txtDocCh(4)) - 1) \ 50)
      End If
      '超項費
      If Val(txtDocCh(5)) > 10 Then
         m_lngOverItemFee = 800# * (txtDocCh(5) - 10)
      End If
      If m_lngOverPageFee > 0 Or m_lngOverItemFee > 0 Then
         txtCP84 = txtCP84 + m_lngOverPageFee + m_lngOverItemFee
      End If
   End If
   '2019/10/24 END
   
   'Add By Sindy 2019/2/18
   chkEexcerpt.Tag = ""
   Call chkEexcerpt_Click
   '2019/2/18 END
End Sub

Private Sub txtBiology_GotFocus(Index As Integer)
   TextInverse txtBiology(Index)
   If Index <> 1 Then CloseIme
End Sub

Private Sub txtBiology_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 0 '寄存國家
      lblCountry = ""
      If txtBiology(Index) <> "" Then
         strExc(0) = "select na72 from nation where na01='" & txtBiology(Index) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            lblCountry = "" & RsTemp(0)
            If txtBiology(Index) >= "011" Then
               chkAtt(8).Value = vbUnchecked
               chkAtt(9).Value = vbChecked
            Else
               chkAtt(8).Value = vbChecked
               chkAtt(9).Value = vbUnchecked
            End If
         Else
            Cancel = True
            MsgBox "寄存國家輸入錯誤！", vbExclamation
         End If
      End If
   Case 2 '寄存日期
      If txtBiology(Index) <> "" Then
         If ChkDate(txtBiology(Index)) = False Then
            txtBiology_GotFocus Index
            Cancel = True
         End If
      End If
   End Select
End Sub

Private Sub txtCP135_GotFocus()
   TextInverse txtCP135
   CloseIme
End Sub

Private Sub txtCP135_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP136_GotFocus()
   TextInverse txtCP136
   CloseIme
End Sub

Private Sub txtCP136_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
   CloseIme
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDocCh_GotFocus(Index As Integer)
   TextInverse txtDocCh(Index)
   CloseIme
End Sub

Private Sub txtDocCh_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDocCh_Validate(Index As Integer, Cancel As Boolean)
   Dim iChecked As Single
   
   If Val(txtDocCh(Index)) > 0 Then
      iChecked = vbChecked
   Else
      iChecked = vbUnchecked
   End If
   
   Select Case Index
   '摘要
   Case 0: chkAtt(1).Value = iChecked:
   '說明書
   Case 1: chkAtt(2).Value = iChecked
   '專利範圍
   Case 2: chkAtt(3).Value = iChecked
   '圖式
   Case 3: chkAtt(4).Value = iChecked
   End Select
   
   'Memo by Lydia 2018/12/27 序列表(不算超頁費=不算錢),所以不加入進度檔的總頁數
   If Index <= 3 Then
      txtDocCh(4) = Val(txtDocCh(0)) + Val(txtDocCh(1)) + Val(txtDocCh(2)) + Val(txtDocCh(3))
   End If
   
   'Add By Sindy 2019/10/24 計算超頁費超項費
   If Text7 = "Y" Then
      If Index = 3 Or Index = 5 Then
         Call Text7_Validate(False)
      End If
   End If
   '2019/10/24 END
End Sub

Private Sub txtFavDate_GotFocus()
   TextInverse txtFavDate
   CloseIme
End Sub

Private Sub txtFavDate_Validate(Cancel As Boolean)
   If txtFavDate <> "" Then
      If ChkDate(txtFavDate) = False Then
         txtFavDate_GotFocus
         Cancel = True
      End If
   End If
End Sub

'Private Function ConvertHtml2Pdf(pCode As String, pFileName As String) As Boolean
'Dim program_name As String, program_path As String
'Dim stHtmlName As String
'Dim stErrFile As String, stErrText As String
'Dim process_id As Long
'Dim process_handle As Long
'
'    stHtmlName = Mid(pFileName, InStrRev(pFileName, "\") + 1)
'    program_path = "C:\TIPOHtml2Pdf\"
'    program_name = program_path & "html2pdfcmd.exe"
'    ' Start the program.
'    On Error GoTo ShellError
'
'    If Dir(program_name) = "" Then
'      MsgBox "指定路徑找不到 Pdf 轉檔程式!!", vbCritical, " ( " & stHtmlName & " ) pdf 轉檔失敗!!"
'      Exit Function
'    End If
'
'    stErrFile = program_path & "errMessage.txt"
'    If Dir(stErrFile) <> "" Then
'      Kill stErrFile
'   End If
'
'    process_id = SHELL(program_name & " " & pCode & " """ & pFileName & """", vbNormalNoFocus)
'
'    On Error GoTo 0
'
'    ' Wait for the program to finish.
'    ' Get the process handle.
'    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
'    If process_handle <> 0 Then
'        WaitForSingleObject process_handle, INFINITE
'        CloseHandle process_handle
'    End If
'
'    If Dir(stErrFile) <> "" Then
'      frm06010301_12.TextBox1.Text = GetText(stErrFile)
'      frm06010301_12.Caption = "( " & stHtmlName & " ) " & frm06010301_12.Caption
'      frm06010301_12.Show vbModal
'    Else
'      ConvertHtml2Pdf = True
'    End If
'    Exit Function
'
'ShellError:
'    MsgBox " " & _
'        program_name & vbCrLf & _
'        Err.Description, vbOKOnly Or vbExclamation, _
'        "Error"
'End Function

Private Function GetText(pFileName As String) As String
   Dim objStream As Object
   Set objStream = CreateObject("ADODB.Stream")
   Dim stContent As String
   
On Error GoTo ErrHnd
   
   With objStream
      .Type = 2
      .Mode = 3
      .Open
      .Charset = "UTF-8" ' 或其他編碼
      '.Charset = "UTF-16"
      .LoadFromFile pFileName
      stContent = .ReadText
      ' PS : 也可透過 .SaveToFile 方法把檔案存檔
      .Close
   End With
   GetText = stContent
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set objStream = Nothing
End Function

'Added by Lydai 2018/08/08 外文本和電子送件暫件區
Private Sub cmdOpen_Click(Index As Integer)
Dim hLocalFile As Long

On Error GoTo ErrHand01
    
    'Added by Lydia 2020/01/20 開啟[原始檔區]
    If Index = 0 And InStr(cmdOpen(Index).Caption, "原始檔") > 0 Then
        If mdiMain.mnuTitle(10).Enabled = True Then
            If cmdOpen(Index).Tag = "" Then
                MsgBox pa(1) & "-" & pa(2) & "在〔原始檔區〕的English_Vers收文號不存在!", vbInformation
            Else
                frm100101_M.m_strKey = cmdOpen(Index).Tag '多筆總收文號
                frm100101_M.SetParent Me
                If frm100101_M.QueryData = True Then
                   frm100101_M.Show
                   Me.Hide
                End If
            End If
        Else
            MsgBox "請先關閉共同查詢畫面！"
        End If
    Else
    'end 2020/01/20
        strExc(1) = ""
        If Index = 0 Then '外文本=English_vers
            'Remove by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
            'strExc(1) = Pub_GetFCPcaseFilePath(pa(2), , pa(1))
        ElseIf Index = 1 Then
            'Modified by Lydia 2024/07/22 改用變數
            'strExc(1) = "\\Typing2\電子送件暫存區\" & pa(1) & pa(2)
            strExc(1) = "\\" & strTyping2Path & "\電子送件暫存區\" & pa(1) & pa(2)
        End If
        If strExc(1) = "" Then Exit Sub
    
        If Dir(strExc(1), vbDirectory) <> "" Then
             ShellExecute hLocalFile, "explore", strExc(1), vbNullString, vbNullString, 1
        Else
             MsgBox strExc(1) & " 資料夾不存在 ！", vbInformation
        End If
            
        Exit Sub
    End If 'Added by Lydia 2020/01/20
    
ErrHand01:
    If Err.Number <> 0 Then
         MsgBox "無法讀取" & strExc(1) & "，請通知電腦中心！", vbCritical
         Resume Next
    End If
End Sub

Private Sub txtTF24_KeyPress(KeyAscii As Integer)
    KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtTF24_LostFocus()
    If Trim(txtTF24.Text) <> "" Then
         txtForeign.Text = Val(txtTF24.Text) + Val(txtTF25.Text)
    End If
End Sub

Private Sub txtTF25_KeyPress(KeyAscii As Integer)
    KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtTF25_LostFocus()
    If Trim(txtTF25.Text) <> "" Then
         txtForeign.Text = Val(txtTF24.Text) + Val(txtTF25.Text)
    End If
End Sub

Private Sub txtForeign_KeyPress(KeyAscii As Integer)
    KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtSimplified_KeyPress(KeyAscii As Integer)
    KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
