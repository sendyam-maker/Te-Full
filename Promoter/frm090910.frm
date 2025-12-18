VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090910 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專-藥證號維護作業"
   ClientHeight    =   7308
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8772
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7308
   ScaleWidth      =   8772
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8760
      Top             =   1128
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
            Picture         =   "frm090910.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090910.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090910.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090910.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090910.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090910.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090910.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090910.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090910.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090910.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090910.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6612
      Left            =   24
      TabIndex        =   0
      Top             =   648
      Width           =   8676
      _ExtentX        =   15304
      _ExtentY        =   11663
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   420
      TabMaxWidth     =   2646
      TabCaption(0)   =   "資料維護"
      TabPicture(0)   =   "frm090910.frx":20F4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtDB(6)"
      Tab(0).Control(1)=   "Label1(2)"
      Tab(0).Control(2)=   "txtDB(2)"
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(4)=   "lblCUID"
      Tab(0).Control(5)=   "txtDB(1)"
      Tab(0).Control(6)=   "Label1(0)"
      Tab(0).Control(7)=   "Label1(4)"
      Tab(0).Control(8)=   "Label1(5)"
      Tab(0).Control(9)=   "txtDB(3)"
      Tab(0).Control(10)=   "Label1(6)"
      Tab(0).Control(11)=   "txtDB(4)"
      Tab(0).Control(12)=   "txtDB(5)"
      Tab(0).Control(13)=   "Label1(7)"
      Tab(0).Control(14)=   "Frame1"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm090910.frx":2110
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(16)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(15)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(13)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtFM2(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtFM2(6)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtFM2(5)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(3)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtFM2(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtFM2(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtFM2(3)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtFM2(4)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblData2(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblData2(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtFM2(7)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label1(8)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label1(9)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label1(10)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtFM2(8)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtFM2(9)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "lblData2(2)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "MGrid2"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cmdQuery"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).ControlCount=   23
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Height          =   300
         Left            =   4536
         TabIndex        =   21
         Top             =   384
         Width           =   885
      End
      Begin VB.Frame Frame1 
         Caption         =   "關聯案號"
         Height          =   3372
         Left            =   -74880
         TabIndex        =   13
         Top             =   3168
         Width           =   8436
         Begin VB.CommandButton Cmd1 
            Caption         =   "清除"
            Height          =   312
            Index           =   2
            Left            =   5496
            TabIndex        =   55
            Top             =   168
            Width           =   852
         End
         Begin VB.CommandButton Cmd1 
            Caption         =   "刪除"
            Height          =   312
            Index           =   1
            Left            =   4584
            TabIndex        =   12
            Top             =   168
            Width           =   852
         End
         Begin VB.TextBox txtCode 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   10.2
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1056
            MaxLength       =   3
            TabIndex        =   6
            Top             =   216
            Width           =   560
         End
         Begin VB.CommandButton Cmd1 
            Caption         =   "加入"
            Height          =   312
            Index           =   0
            Left            =   3672
            TabIndex        =   11
            Top             =   168
            Width           =   852
         End
         Begin VB.TextBox txtCode 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   10.2
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1664
            MaxLength       =   6
            TabIndex        =   7
            Top             =   216
            Width           =   780
         End
         Begin VB.TextBox txtCode 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   10.2
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   2488
            MaxLength       =   1
            TabIndex        =   8
            Top             =   216
            Width           =   348
         End
         Begin VB.TextBox txtCode 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   10.2
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   9
            Top             =   216
            Width           =   492
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
            Height          =   1692
            Left            =   144
            TabIndex        =   14
            Top             =   1584
            Width           =   8124
            _ExtentX        =   14330
            _ExtentY        =   2985
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            AllowUserResizing=   3
            FormatString    =   "本所案號|案件名稱|專利連結通知(Y)|代理人名稱|申請人1名稱"
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   1392
            X2              =   3048
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Label Label2 
            Caption         =   "本所案號："
            Height          =   276
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   264
            Width           =   924
         End
         Begin VB.Label Label2 
            Caption         =   "案件名稱："
            Height          =   276
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   588
            Width           =   924
         End
         Begin VB.Label Label2 
            Caption         =   "申請人1："
            Height          =   276
            Index           =   2
            Left            =   120
            TabIndex        =   31
            Top             =   1248
            Width           =   828
         End
         Begin VB.Label Label2 
            Caption         =   "代理人："
            Height          =   276
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   828
         End
         Begin MSForms.ComboBox Combo1 
            Height          =   348
            Left            =   1032
            TabIndex        =   10
            Top             =   552
            Width           =   7308
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "12890;614"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblData 
            Height          =   276
            Index           =   0
            Left            =   1032
            TabIndex        =   18
            Top             =   960
            Width           =   900
            BackColor       =   16777215
            Size            =   "1587;487"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblData 
            Height          =   276
            Index           =   1
            Left            =   2000
            TabIndex        =   17
            Top             =   960
            Width           =   6200
            BackColor       =   16777215
            VariousPropertyBits=   27
            Size            =   "10936;487"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblData 
            Height          =   276
            Index           =   2
            Left            =   1032
            TabIndex        =   16
            Top             =   1248
            Width           =   900
            BackColor       =   16777215
            Size            =   "1587;487"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblData 
            Height          =   276
            Index           =   3
            Left            =   2000
            TabIndex        =   15
            Top             =   1248
            Width           =   6200
            BackColor       =   16777215
            VariousPropertyBits=   27
            Size            =   "10936;487"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid2 
         Height          =   3700
         Left            =   96
         TabIndex        =   45
         Top             =   2760
         Width           =   8412
         _ExtentX        =   14838
         _ExtentY        =   6519
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         AllowUserResizing=   3
         FormatString    =   "流水號|藥證號|藥品名稱(中文)|藥品名稱(英文)|有效成分|本所案號|案件名稱|專利連結通知(Y)|代理人名稱|申請人1名稱|藥證號備註"
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   11
      End
      Begin MSForms.Label lblData2 
         Height          =   276
         Index           =   2
         Left            =   3456
         TabIndex        =   54
         Top             =   744
         Width           =   5028
         BackColor       =   16777215
         Size            =   "8869;487"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   300
         Index           =   9
         Left            =   1440
         TabIndex        =   30
         Top             =   2400
         Width           =   7080
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "12488;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   300
         Index           =   8
         Left            =   1440
         TabIndex        =   29
         Top             =   2076
         Width           =   7080
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "12488;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "有效成分："
         Height          =   276
         Index           =   10
         Left            =   96
         TabIndex        =   53
         Top             =   2412
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "藥品名稱(英文)"
         Height          =   204
         Index           =   9
         Left            =   96
         TabIndex        =   52
         Top             =   2124
         Width           =   1272
      End
      Begin VB.Label Label1 
         Caption         =   "藥品名稱(中文)"
         Height          =   204
         Index           =   8
         Left            =   96
         TabIndex        =   51
         Top             =   1800
         Width           =   1320
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   300
         Index           =   7
         Left            =   1440
         TabIndex        =   28
         Top             =   1752
         Width           =   7080
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "12488;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "有效成分："
         Height          =   276
         Index           =   7
         Left            =   -74940
         TabIndex        =   50
         Top             =   1776
         Width           =   960
      End
      Begin MSForms.TextBox txtDB 
         Height          =   492
         Index           =   5
         Left            =   -73884
         TabIndex        =   4
         Top             =   1752
         Width           =   7404
         VariousPropertyBits=   -1466939365
         MaxLength       =   200
         Size            =   "13060;868"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   204
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   492
         Index           =   4
         Left            =   -73716
         TabIndex        =   3
         Top             =   1236
         Width           =   7236
         VariousPropertyBits=   -1467987941
         MaxLength       =   100
         Size            =   "12763;868"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   204
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "藥品名稱(英文)"
         Height          =   204
         Index           =   6
         Left            =   -74940
         TabIndex        =   49
         Top             =   1320
         Width           =   1272
      End
      Begin MSForms.TextBox txtDB 
         Height          =   492
         Index           =   3
         Left            =   -73716
         TabIndex        =   2
         Top             =   729
         Width           =   7236
         VariousPropertyBits=   -1467987941
         MaxLength       =   100
         Size            =   "12763;868"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   204
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "藥品名稱(中文)"
         Height          =   204
         Index           =   5
         Left            =   -74940
         TabIndex        =   48
         Top             =   816
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "(系統自動編號)"
         ForeColor       =   &H000000FF&
         Height          =   276
         Index           =   4
         Left            =   -68088
         TabIndex        =   46
         Top             =   408
         Width           =   1296
      End
      Begin MSForms.Label lblData2 
         Height          =   276
         Index           =   1
         Left            =   2184
         TabIndex        =   44
         Top             =   1440
         Width           =   6396
         BackColor       =   16777215
         Size            =   "11289;494"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData2 
         Height          =   276
         Index           =   0
         Left            =   2184
         TabIndex        =   43
         Top             =   1104
         Width           =   6396
         BackColor       =   16777215
         Size            =   "11289;494"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   300
         Index           =   4
         Left            =   2856
         TabIndex        =   25
         Top             =   744
         Width           =   564
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "995;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   300
         Index           =   3
         Left            =   2470
         TabIndex        =   24
         Top             =   744
         Width           =   348
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "614;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   300
         Index           =   2
         Left            =   1653
         TabIndex        =   23
         Top             =   744
         Width           =   780
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   300
         Index           =   1
         Left            =   1056
         TabIndex        =   22
         Top             =   744
         Width           =   560
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "988;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "藥證號："
         Height          =   276
         Index           =   3
         Left            =   96
         TabIndex        =   42
         Top             =   432
         Width           =   756
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   300
         Index           =   5
         Left            =   1056
         TabIndex        =   26
         Top             =   1080
         Width           =   1092
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   300
         Index           =   6
         Left            =   1056
         TabIndex        =   27
         Top             =   1416
         Width           =   1092
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtFM2 
         Height          =   300
         Index           =   0
         Left            =   1056
         TabIndex        =   20
         Top             =   384
         Width           =   3400
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "5997;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   13
         Left            =   96
         TabIndex        =   41
         Top             =   804
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Index           =   15
         Left            =   96
         TabIndex        =   40
         Top             =   1476
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         Height          =   180
         Index           =   16
         Left            =   96
         TabIndex        =   39
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "流水號："
         Height          =   276
         Index           =   0
         Left            =   -69624
         TabIndex        =   38
         Top             =   408
         Width           =   816
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   1
         Left            =   -68700
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   396
         Width           =   588
         VariousPropertyBits=   671107099
         MaxLength       =   3
         Size            =   "1037;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   204
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCUID 
         Height          =   252
         Left            =   -74784
         TabIndex        =   36
         Top             =   2880
         Width           =   8340
         BackColor       =   16777215
         Size            =   "14711;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "藥證號："
         Height          =   276
         Index           =   1
         Left            =   -74940
         TabIndex        =   35
         Top             =   432
         Width           =   816
      End
      Begin MSForms.TextBox txtDB 
         Height          =   330
         Index           =   2
         Left            =   -73884
         TabIndex        =   1
         Top             =   396
         Width           =   3400
         VariousPropertyBits=   671107099
         MaxLength       =   20
         Size            =   "5997;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   204
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "備　註："
         Height          =   276
         Index           =   2
         Left            =   -74940
         TabIndex        =   34
         Top             =   2328
         Width           =   816
      End
      Begin MSForms.TextBox txtDB 
         Height          =   564
         Index           =   6
         Left            =   -73884
         TabIndex        =   5
         Top             =   2280
         Width           =   7404
         VariousPropertyBits=   -1466939365
         MaxLength       =   500
         Size            =   "13060;995"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   204
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   1512
         X2              =   3048
         Y1              =   912
         Y2              =   912
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   8772
      _ExtentX        =   15473
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
   End
End
Attribute VB_Name = "frm090910"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/06/28 改成Form2.0 ;txtDB(index)、lblData(index)、lblData2(index)、lblCUID、txtFM2(index)、MGrid1字型、MGrid2字型
'Create by Lydia 2023/06/28 藥證號維護+藥證號對照檔
Option Explicit

Dim m_EditMode As Integer '0:瀏覽 1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_blnColOrderAsc As Boolean '查詢:欄位資料由小到大排序
Dim oText As Control, oLabel As Control
Dim stCon As String, stSQL As String, intR As Integer
Dim rsRead As New ADODB.Recordset
Dim rsMap As New ADODB.Recordset, rsMapOld As New ADODB.Recordset
Dim m_MPKey As String
Dim colPA01 As Integer, colPA02 As Integer, colPA03 As Integer, colPA04 As Integer, colMPKey As Integer
Dim colPA177 As Integer
Private Const cntFixed1 As Integer = 8
Private Const cntFixed2 As Integer = 5
Dim intLastRow1 As Integer, intLastRow2 As Integer

Private Sub Cmd1_Click(Index As Integer)
   
   If Index = 2 Then '清除
      Call ClearCaseMap(True)
      Exit Sub
   End If
   
   If Trim(txtCode(0) & txtCode(1) & txtCode(2) & txtCode(3)) <> "" Then
      If Trim(txtCode(0) & txtCode(1) & txtCode(2) & txtCode(3)) <> Trim(txtCode(0).Tag & txtCode(1).Tag & txtCode(2).Tag & txtCode(3).Tag) Then
         MsgBox "請更新本所案號的資料！", vbExclamation
         txtCode(1).SetFocus
         Exit Sub
      End If
      If Index = 0 And Val(txtDB(1)) > 0 And m_MPKey = Val(txtDB(1)) & Trim(txtCode(0) & txtCode(1) & txtCode(2) & txtCode(3)) Then
         MsgBox "關聯案號已存在，不可加入！", vbExclamation
         txtCode(1).SetFocus
         Exit Sub
      End If
      If Combo1.Tag = "" Then
         MsgBox "請輸入正確的本所案號!", vbExclamation
         Exit Sub
      Else
         If UpdateCaseMap(IIf(Index = 0, "A", "D")) = True Then
            Call ClearCaseMap(True)
            Exit Sub
         Else
            txtCode(1).SetFocus
            Exit Sub
         End If
      End If
   End If
   

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsMap = Nothing
   Set rsMapOld = Nothing
   Set rsRead = Nothing
   
   Set frm090910 = Nothing
End Sub

Private Sub MGrid1_Click()

Dim intRow As Integer
   With MGrid1
      If .MouseRow > 0 Then
         intRow = .MouseRow
         .row = intRow
         GridClick MGrid1, intLastRow1, 0, 0
         intLastRow1 = intRow
         If intLastRow1 > 0 And .TextMatrix(intLastRow1, 0) <> "" And .TextMatrix(intLastRow1, colPA01) <> "" Then
            txtCode(0) = .TextMatrix(intLastRow1, colPA01)
            txtCode(1) = .TextMatrix(intLastRow1, colPA02)
            txtCode(2) = .TextMatrix(intLastRow1, colPA03)
            txtCode(3) = .TextMatrix(intLastRow1, colPA04)
            m_MPKey = .TextMatrix(intLastRow1, colMPKey)
         Else
            txtCode(0) = ""
            txtCode(1) = ""
            txtCode(2) = ""
            txtCode(3) = ""
            m_MPKey = ""
         End If
      End If
   End With
   Call txtCode_Validate(2, False)

End Sub

Private Sub MGrid2_Click()
Dim intRow As Integer
   With MGrid2
      If .MouseRow > 0 Then
         intRow = .MouseRow
         .row = intRow
         GridClick MGrid2, intLastRow2, 0, 0
         intLastRow2 = intRow
      End If
   End With
      
End Sub


Private Sub MGrid2_DblClick()
Dim intRow As Integer
   With MGrid2
       If .MouseRow > 0 Then
          intRow = .MouseRow
          .row = intRow
          If .row > 0 And .TextMatrix(intRow, 0) <> "" And .TextMatrix(intRow, 1) <> "" Then
              ReadData .TextMatrix(intRow, 1)
          End If
       End If
   End With
End Sub

Private Sub MGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow MGrid2, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MGrid2.col = nCol
   MGrid2.row = nRow
   If Me.MGrid2.row < 1 And Me.MGrid2.Text <> "V" Then
      If InStr("流水號,", Me.MGrid2.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.MGrid2.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid2.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MGrid2.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid2.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 1 And intLastRow2 > 0 Then
      If MGrid2.TextMatrix(intLastRow2, 0) <> "" And MGrid2.TextMatrix(intLastRow2, 1) <> "" Then
         ReadData MGrid2.TextMatrix(intLastRow2, 1)
      End If
   End If
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   TextInverse txtCode(Index)
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cmdQuery_Click()
   intLastRow2 = 0
   Call ReadCaseMap("2")
End Sub

Private Sub ReadCaseMap(ByVal pKind As String, Optional pKey As String)
'pKind: 1=資料維護, 2=多筆查詢
  
   stCon = ""
   If pKind = "1" Then
      stCon = stCon & " and MC01 = '" & Val(pKey) & "' AND PA01 IS NOT NULL"
   Else
      '藥證號
      If txtFM2(0) <> "" Then
         stCon = stCon & " and UPPER(MC02) like '%" & ChgSQL(UCase(txtFM2(0))) & "%'"
      End If
      '本所案號
      If txtFM2(1) <> "" Then
         stCon = stCon & " and PA01 = '" & txtFM2(1) & "' "
      End If
      If txtFM2(2) <> "" Then
         stCon = stCon & " and PA02 = '" & txtFM2(2) & "' "
      End If
      If txtFM2(3) <> "" Then
         stCon = stCon & " and PA03 = '" & txtFM2(3) & "' "
      End If
      If txtFM2(4) <> "" Then
         stCon = stCon & " and PA04 = '" & txtFM2(4) & "' "
      End If
      '代理人
      If txtFM2(5) <> "" Then
         stCon = stCon & " and FANO = '" & ChangeCustomerL(txtFM2(5)) & "' "
      End If
      '申請人
      If txtFM2(6) <> "" Then
         stCon = stCon & " and APP01 = '" & ChangeCustomerL(txtFM2(6)) & "' "
      End If
      '藥品名稱(中文)
      If txtFM2(7) <> "" Then
         stCon = stCon & " and UPPER(MC03) like '%" & ChgSQL(UCase(txtFM2(7))) & "%'"
      End If
      '藥品名稱(英文)
      If txtFM2(8) <> "" Then
         stCon = stCon & " and UPPER(MC04) like '%" & ChgSQL(UCase(txtFM2(8))) & "%'"
      End If
      '有效成分
      If txtFM2(9) <> "" Then
         stCon = stCon & " and UPPER(MC05) like '%" & ChgSQL(UCase(txtFM2(9))) & "%'"
      End If
   End If
   
   stSQL = "SELECT '' AS V, MC01,MC02,MC03,MC04,MC05,DECODE(PA01,NULL,NULL,PA01||'-'||PA02||DECODE(PA03||PA04,'000',NULL,'-'||PA03||'-'||PA04)) AS CASENO," & _
           " CNAME,PA177,FNAME,APPNAME1,MC06,FANO,APP01,PA01,PA02,PA03,PA04,MC01||PA01||PA02||PA03||PA04 AS MPKEY" & _
           " FROM MEDICINECODE, MEDICINECODEMAP,(SELECT PA01,PA02,PA03,PA04,NVL(PA05,NVL(PA06,PA07)) CNAME,PA177," & _
           " NVL(FA04,NVL(FA05,FA06)) FNAME, NVL(CU04,NVL(CU05,CU06)) APPNAME1 ,PA75 AS FANO,PA26 AS APP01" & _
           " FROM PATENT, FAGENT, CUSTOMER WHERE SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+)" & _
           " AND (PA01,PA02,PA03,PA04) IN (SELECT MCM02,MCM03,MCM04,MCM05 FROM MEDICINECODEMAP)" & _
           " UNION SELECT SP01,SP02,SP03,SP04,NVL(SP05,NVL(SP06,SP07)) CNAME,'' AS PA177," & _
           " NVL(FA04,NVL(FA05,FA06)) FNAME, NVL(CU04,NVL(CU05,CU06)) APPNAME1,SP26 AS FANO,SP08 AS APP01" & _
           " From SERVICEPRACTICE, FAGENT, CUSTOMER WHERE SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+)" & _
           " AND (SP01,SP02,SP03,SP04) IN (SELECT MCM02,MCM03,MCM04,MCM05 FROM MEDICINECODEMAP)" & _
           " ) WHERE MC01=MCM01(+) AND MCM02=PA01(+) AND MCM03=PA02(+) AND MCM04=PA03(+) AND MCM05=PA04(+) " & stCon
   stSQL = stSQL & " ORDER BY MC01, CASENO "
   intR = 1
   
   If pKind = "1" Then
      Call SetGrd("1", MGrid1, True)
      Set rsRead = ClsLawReadRstMsg(intR, stSQL)
      '操作資料
      Set rsMap = PUB_CreateRecordset(rsRead, , , , Me.Name)
      If intR = 1 Then
         MGrid1.FixedCols = 0
         Call SetGrd("1", MGrid1, , rsMap)
         MGrid1.FixedCols = cntFixed1
      End If

      '保留原始資料
      intR = 1
      Set rsMapOld = ClsLawReadRstMsg(intR, stSQL)
   Else
      Set rsRead = ClsLawReadRstMsg(intR, stSQL)
      Call SetGrd("2", MGrid2, True)
      If intR = 1 Then
         MGrid2.FixedCols = 0
         Call SetGrd("2", MGrid2, , rsRead)
         MGrid2.FixedCols = cntFixed2
      End If
   End If
End Sub

Private Sub SetGrd(ByVal pKind As String, ByRef pGRD As MSHFlexGrid, Optional ByVal pReset As Boolean = False, Optional ByRef pRst As ADODB.Recordset)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   If pKind = "1" Then
      '資料維護
      arrGridHeadText = Array("v", "流水號", "藥 證 號", "藥品名稱(中文)", "藥品名稱(英文)", "有效成分", "本所案號", "案 件 名 稱", "專利連結", "代理人名稱", "申請人1名稱", "備    註", "FANO", "APP01", "PA01", "PA02", "PA03", "PA04", "MPKEY")
      arrGridHeadWidth = Array(200, 0, 0, 0, 0, 0, 1200, 2000, 800, 1500, 1500, 0, 0, 0, 0, 0, 0, 0, 0)
   Else
      '多筆查詢
      arrGridHeadText = Array("v", "流水號", "藥 證 號", "藥品名稱(中文)", "藥品名稱(英文)", "有效成分", "本所案號", "案 件 名 稱", "專利連結", "代理人名稱", "申請人1名稱", "備    註", "FANO", "APP01", "PA01", "PA02", "PA03", "PA04", "MPKEY")
      arrGridHeadWidth = Array(200, 600, 1800, 1500, 1500, 1100, 1200, 1800, 800, 1100, 1100, 1100, 0, 0, 0, 0, 0, 0, 0)
   End If
   
   pGRD.Visible = False
   
   pGRD.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
      pGRD.Clear
      pGRD.Rows = 2
   Else
      pGRD.FixedCols = 0
      Set pGRD.DataSource = pRst
   End If
   For iRow = 0 To pGRD.Cols - 1
      pGRD.row = 0
      pGRD.col = iRow
      pGRD.Text = arrGridHeadText(iRow)
      pGRD.ColWidth(iRow) = arrGridHeadWidth(iRow)
      pGRD.CellAlignment = flexAlignCenterCenter
   Next
   If colPA01 = 0 And pKind = "1" Then
      colPA01 = PUB_MGridGetId("PA01", pGRD)
      colPA02 = PUB_MGridGetId("PA02", pGRD)
      colPA03 = PUB_MGridGetId("PA03", pGRD)
      colPA04 = PUB_MGridGetId("PA04", pGRD)
      colMPKey = PUB_MGridGetId("MPKEY", pGRD)
      colPA177 = PUB_MGridGetId("專利連結", pGRD)
   End If
   
   For intI = 1 To pGRD.Rows - 1
      pGRD.row = intI
      For iRow = 0 To pGRD.Cols - 1
         pGRD.col = iRow
         pGRD.CellBackColor = &H80000005
         If iRow = colPA177 Then
            pGRD.CellAlignment = flexAlignCenterCenter
         End If
      Next iRow
   Next intI
   
   pGRD.Visible = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Screen.MousePointer = vbHourglass
   Select Case KeyCode
      Case vbKeyF2 '新增
         KeyCode = 0: Action 1
      Case vbKeyF3 '修改
         KeyCode = 0: Action 2
      Case vbKeyF4: '查詢
         KeyCode = 0: Action 4
      Case vbKeyF5 '刪除
         KeyCode = 0: Action 3
      Case vbKeyHome '第一筆
         KeyCode = 0: Action 6
      Case vbKeyPageUp '上一筆
         KeyCode = 0: Action 7
      Case vbKeyPageDown '下一筆
         KeyCode = 0: Action 8
      Case vbKeyEnd: '最後筆
         KeyCode = 0: Action 9
      Case vbKeyF9 '確定
         KeyCode = 0: Action 11
      
      Case vbKeyF10 '取消
         KeyCode = 0: Action 12
      Case vbKeyEscape '結束
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            KeyCode = 0: Action 14
         End If
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction("frm090910", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm090910", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm090910", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm090910", strFind, False)
  
   MoveFormToCenter Me
   
   For Each oLabel In lblData
       oLabel.BackColor = &H8000000F
   Next
   For Each oLabel In lblData2
       oLabel.BackColor = &H8000000F
   Next
   lblCUID.BackColor = &H8000000F
   
   Call SetGrd("1", MGrid1, True)
   Call SetGrd("2", MGrid2, True)
   
   Action 6 '預設第一筆
   UpdateToolbarState
   
   Me.SSTab1.Tab = 1 '改從多筆查詢頁籤開始
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
   Action Button.Index
   Screen.MousePointer = vbDefault
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
         If m_bUpdate And txtDB(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtDB(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtDB(1) <> "" Then
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

Private Sub TxtLock()
   Select Case m_EditMode
   Case 0 '瀏覽
      For Each oText In txtDB
         oText.Locked = True
      Next
      For Each oText In txtCode
         oText.Locked = True
      Next
      Cmd1(0).Visible = False: Cmd1(1).Visible = False: Cmd1(2).Visible = False
      SSTab1.TabEnabled(1) = True
   Case Else
      For Each oText In txtDB
         oText.Locked = False
      Next
      For Each oText In txtCode
         oText.Locked = False
      Next
      If m_EditMode <> 4 Then
         txtDB(1).Locked = True
         txtDB(2).SetFocus
         txtDB_GotFocus 2
      End If
      Cmd1(0).Visible = True: Cmd1(1).Visible = True: Cmd1(2).Visible = True
      SSTab1.TabEnabled(1) = False
   End Select
End Sub

Private Sub Action(Index As Integer)
Dim bCancel As Boolean

   If TBar1.Buttons(Index).Enabled = False Then Exit Sub

On Error GoTo ErrHand

   SSTab1.Tab = 0
   Select Case Index
      Case 1 '按下新增
        m_EditMode = 1
        FormReset
        Call ReadCaseMap("1", "0")
      Case 2 '按下修改
         m_EditMode = 2

      Case 3 '按下刪除
         If txtDB(1).Text = "" Then
             MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
             Exit Sub
         End If

         If DelMsg() = True Then
            If FormDelete() = False Then
               MsgBox "刪除失敗!", vbCritical
               Exit Sub
            '刪除後移到最末筆
            Else
               ShowRecord 3
            End If
         End If

      Case 4 '按下查詢
         FormReset
         m_EditMode = 4
         txtDB(1).Enabled = True
         txtDB(1).SetFocus
         
      Case 6 '第一筆
         ShowRecord 0
      Case 7 '前一筆
         ShowRecord 1
      Case 8 '後一筆
         ShowRecord 2
      Case 9 '最後筆
         ShowRecord 3
      Case 11 '按下確定
         
         Select Case m_EditMode
            '新增,修改
            Case 1, 2
               '新增,修改都要判斷
               If RecIsExist = True Then Exit Sub
               
               If TxtValidate = False Then
                  Exit Sub
               Else
                  If FormSave() = False Then
                     MsgBox "存檔失敗!", vbCritical
                     Exit Sub
                  Else
                     m_EditMode = 0
                     If txtDB(1) = "" Then
                        ShowRecord 3
                     Else
                        ReadData txtDB(1)
                     End If

                  End If
               End If
            '查詢
            Case 4
               If ReadData(txtDB(1)) = False Then
                  MsgBox "無資料!", vbExclamation
                  Exit Sub
               Else
                  m_EditMode = 0
               End If
         End Select
      Case 12 '按下取消
         m_EditMode = 0
         txtDB(1) = txtDB(1).Tag
         If txtDB(1) <> "" Then
            If ReadData(txtDB(1)) = False Then
               ShowRecord 3
            End If
         Else
            FormReset
         End If
      Case 14 '結束
         Unload Me
         Exit Sub
   End Select
   UpdateToolbarState
   TxtLock
   Exit Sub
   
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

' 顯示資料
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
 Dim stKEY As String
    
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case p_iWay
      Case 0 '第一筆
         strExc(0) = "SELECT nvl(min(MC01),0) FROM MedicineCode"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 1 '前一筆
         strExc(0) = "SELECT nvl(max(MC01),0) FROM MedicineCode where MC01<" & txtDB(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 6
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 2 '後一筆
         strExc(0) = "SELECT nvl(min(MC01),0) FROM MedicineCode where MC01>" & txtDB(1)
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 7
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 3 '最後筆
         strExc(0) = "SELECT nvl(max(MC01),0) FROM MedicineCode"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
   End Select
      
      
   If stKEY <> "" Then
      ReadData stKEY
      ShowRecord = True
   Else
      FormReset
      ShowRecord = True
   End If
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Private Function ReadData(Optional ByVal pKey As String) As Boolean

   '單筆
   If pKey <> "" Then
      stCon = " and MC01=" & pKey
   '多筆
   Else
      If txtDB(2) <> "" Then
         stCon = stCon & " and UPPER(MC02) like '%" & ChgSQL(UCase(txtDB(2))) & "%'"
      End If
      If txtDB(3) <> "" Then
         stCon = stCon & " and UPPER(MC03) like '%" & ChgSQL(UCase(txtDB(3))) & "%'"
      End If
      If txtDB(4) <> "" Then
         stCon = stCon & " and UPPER(MC04) like '%" & ChgSQL(UCase(txtDB(4))) & "%'"
      End If
      If txtDB(5) <> "" Then
         stCon = stCon & " and UPPER(MC05) like '%" & ChgSQL(UCase(txtDB(5))) & "%'"
      End If
      If txtDB(6) <> "" Then
         stCon = stCon & " and UPPER(MC06) like '%" & ChgSQL(UCase(txtDB(6))) & "%'"
      End If
   End If
   
   FormReset
   
   strExc(0) = "select * from MedicineCode where 1=1 " & stCon & " order by MC01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If m_EditMode = 4 Then
         '單筆查詢
         RsTemp.MoveFirst
       Else
         SSTab1.Tab = 0
      End If
      SetData RsTemp
      ReadData = True
   End If
End Function

Private Sub SetData(ByRef rsQuery As ADODB.Recordset, Optional ByVal iRow As Integer)
   If iRow > 0 Then
      rsQuery.MoveFirst
      If iRow > 1 Then
         rsQuery.Move iRow - 1
      End If
      SSTab1.Tab = 0
   End If
   
   With rsQuery
   For Each oText In txtDB
      oText = "" & .Fields("MC" & Format(oText.Index, "00"))
   Next
   End With
   UpdateCUID rsQuery
   
   '讀取藥證號數對照檔
   Call ReadCaseMap("1", txtDB(1))
   
   txtDB(1).Tag = txtDB(1)
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   If IsNull(rsSrcTmp.Fields("MC07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("MC07")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("MC07"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("MC08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("MC08")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("MC08"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("MC09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("MC09")) = False Then
         strTemp = rsSrcTmp.Fields("MC09")
         strCTime = Format(strTemp, "00:00:00")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("MC10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("MC10")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("MC10"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("MC11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("MC11")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("MC11"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("MC12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("MC12")) = False Then
         strTemp = rsSrcTmp.Fields("MC12")
         strUTime = Format(strTemp, "00:00:00")
      End If
   End If
   
   lblCUID = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub FormReset()
   
   For Each oText In txtDB
      oText.Text = ""
      oText.Tag = ""
   Next
   
   lblCUID.Caption = ""
   
   Call SetGrd("1", MGrid1, True)
      
   Call ClearCaseMap(True)
   
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
Dim bolClear As Boolean

   If txtCode(Index).Text <> txtCode(Index).Tag Then
     If txtCode(0) <> "" And txtCode(1) <> "" Then
        If Trim(txtCode(2).Text) = "" Then txtCode(2).Text = "0"
        If Trim(txtCode(3).Text) = "" Then txtCode(3).Text = "00"
        strExc(0) = "SELECT PA01,PA02,PA03,PA04,PA05,PA06,PA07,PA75 AS FANO, NVL(FA04,NVL(FA05,FA06)) FANAME,PA26 AS APP01, NVL(CU04,NVL(CU05,CU06)) APP01NAME,PA177 " & _
                    "From PATENT, FAGENT, CUSTOMER Where PA01='" & txtCode(0) & "' and PA02='" & txtCode(1) & "' and PA03='" & txtCode(2) & "' and PA04='" & txtCode(3) & "' " & _
                    "AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
                    "Union SELECT SP01,SP02,SP03,SP04,SP05,SP06,SP07,SP26 AS FANO, NVL(FA04,NVL(FA05,FA06)) FANAME,SP08 AS APP01, NVL(CU04,NVL(CU05,CU06)) APP01NAME,'' as PA177 " & _
                    "From ServicePractice, FAGENT, CUSTOMER Where SP01='" & txtCode(0) & "' and SP02='" & txtCode(1) & "' and SP03='" & txtCode(2) & "' and SP04='" & txtCode(3) & "' " & _
                    "AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           AddCboName Combo1, "" & RsTemp.Fields("PA05"), "" & RsTemp.Fields("PA06"), "" & RsTemp.Fields("PA07")
           Combo1.Tag = IIf("" & RsTemp.Fields("PA05") <> "", "" & RsTemp.Fields("PA05"), IIf("" & RsTemp.Fields("PA06") <> "", "" & RsTemp.Fields("PA06"), "" & RsTemp.Fields("PA07")))
           lblData(0).Tag = "" & RsTemp.Fields("PA177") '專利連結
           lblData(0) = "" & RsTemp.Fields("FANO")
           lblData(1) = "" & RsTemp.Fields("FANAME")
           lblData(2) = "" & RsTemp.Fields("APP01")
           lblData(3) = "" & RsTemp.Fields("APP01NAME")
           txtCode(0).Tag = txtCode(0).Text
           txtCode(1).Tag = txtCode(1).Text
           txtCode(2).Tag = txtCode(2).Text
           txtCode(3).Tag = txtCode(3).Text
        Else
           MsgBox "請輸入正確的本所案號!", vbExclamation
           bolClear = True
        End If
     End If
   End If
   If bolClear = True Or (txtCode(0) = "" And txtCode(1) = "") Then
      Call ClearCaseMap(False)
      txtCode(0).Tag = txtCode(0).Text
      txtCode(1).Tag = txtCode(1).Text
      txtCode(2).Tag = txtCode(2).Text
      txtCode(3).Tag = txtCode(3).Text
   End If
End Sub

Private Sub txtDB_GotFocus(Index As Integer)
   TextInverse txtDB(Index)
End Sub

Private Sub txtDB_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index = 2 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtDB_Validate(Index As Integer, Cancel As Boolean)
   '藥證號和藥品名稱不可換行
   If Index < 4 Then
      txtDB(Index).Text = PUB_StringFilter(txtDB(Index).Text)
   End If
   '藥證號若有阿拉伯數字統一為半形數字
   If Index = 2 Then
      txtDB(Index).Text = PUB_ChgNumeralStyle(txtDB(Index).Text)
   End If

End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean, idx As Integer
   
   If txtDB(2) = "" Then
      MsgBox "請輸入藥證號!", vbExclamation
      txtDB(2).SetFocus
      Exit Function
   End If
   
   If txtDB(3) & txtDB(4) = "" Then
      If MsgBox("是否要輸入藥品名稱？", vbYesNo + vbDefaultButton1) = vbYes Then
        txtDB(3).SetFocus
        Exit Function
      End If
   End If
  
   '檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
Dim bolExist As Boolean
Dim StrCaseList As String

On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
   If m_EditMode = 1 Then
      txtDB(1) = Pub_GetDefColMaxNo("MedicineCode", "MC01")
      strSql = "insert into MedicineCode (MC01,MC02,MC03,MC04,MC05,MC06,MC07,MC08,MC09) VALUES ('" & txtDB(1) & "','" & _
               ChgSQL(txtDB(2)) & "','" & ChgSQL(txtDB(3)) & "' ,'" & ChgSQL(txtDB(4)) & "' ,'" & ChgSQL(txtDB(5)) & "','" & ChgSQL(txtDB(6)) & "','" & strUserNum & "',TO_CHAR(SYSDATE,'YYYYMMDD'),TO_CHAR(SYSDATE,'HH24MMSS')) "
   Else
      strSql = "update MedicineCode set MC02='" & ChgSQL(txtDB(2)) & "',MC03='" & ChgSQL(txtDB(3)) & "'" & _
         ",MC04='" & ChgSQL(txtDB(4)) & "',MC05='" & ChgSQL(txtDB(5)) & "',MC06='" & ChgSQL(txtDB(6)) & "',MC10='" & strUserNum & "',MC11=TO_CHAR(SYSDATE,'YYYYMMDD'),MC12=TO_CHAR(SYSDATE,'HH24MMSS')" & _
         " where MC01=" & txtDB(1)
   End If
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   
   '*******關聯案號*******
   '有資料 => 無資料
   If rsMap.RecordCount = 0 Then
      If rsMapOld.RecordCount > 0 Then
         '刪除資料
         strSql = "DELETE FROM MedicineCodeMap WHERE MCM01='" & txtDB(1) & "' "
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql, intI
      End If
   Else
      '刪除資料(原來的資料在新的資料中找不到的)
      With rsMapOld
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            rsMap.MoveFirst
            bolExist = False
            Do While Not (rsMap.EOF Or bolExist = True)
               If rsMap.Fields("MPKEY") = .Fields("MPKEY") And Len("" & .Fields("MPKEY")) >= 10 Then
                  bolExist = True
               End If
               rsMap.MoveNext
            Loop
            
            If bolExist = False And Len("" & .Fields("MPKEY")) >= 10 Then
               '刪除資料
               strSql = "DELETE FROM MedicineCodeMap WHERE MCM01='" & .Fields("MC01") & "' and MCM02='" & .Fields("PA01") & "' and MCM03='" & .Fields("PA02") & "' and MCM04='" & .Fields("PA03") & "' and MCM05='" & .Fields("PA04") & "' "
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql, intI
            End If
            .MoveNext
         Loop
      End If
      End With
      
      '新增
      With rsMap
      .MoveFirst
      Do While Not .EOF
         If rsMapOld.RecordCount = 0 Then
            bolExist = False
         Else
            rsMapOld.MoveFirst
            bolExist = False
            Do While Not (rsMapOld.EOF Or bolExist = True)
               If rsMapOld.Fields("MPKEY") = .Fields("MPKEY") Then
                  bolExist = True
               End If
               rsMapOld.MoveNext
            Loop
         End If
         
         '新增資料
         If bolExist = False Then
            strSql = "INSERT INTO MedicineCodeMap (MCM01,MCM02,MCM03,MCM04,MCM05,MCM06,MCM07,MCM08)" & _
                     "VALUES ('" & txtDB(1) & "','" & .Fields("PA01") & "','" & .Fields("PA02") & "','" & .Fields("PA03") & "'," & CNULL(.Fields("PA04")) & _
                     ",'" & strUserNum & "',TO_CHAR(SYSDATE,'YYYYMMDD'),TO_CHAR(SYSDATE,'HH24MMSS')) "
            cnnConnection.Execute strSql, intI
         End If
         .MoveNext
      Loop
      End With
   End If
   
   
   '**********************
   cnnConnection.CommitTrans
   FormSave = True
   
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

Private Function FormDelete() As Boolean
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   strSql = "delete from MedicineCode where MC01=" & txtDB(1)
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   strSql = "delete from MedicineCodeMap where MCM01=" & txtDB(1)
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   FormDelete = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

Private Function RecIsExist() As Boolean

stCon = ""
If Trim(txtDB(2)) <> "" Then
   stCon = stCon & "and MC02='" & ChgSQL(Trim(txtDB(2))) & "' "
End If

If Left(stCon, 3) = "and" Then
    stCon = Mid(stCon, 4, Len(stCon) - 4)
ElseIf stCon = "" Then
    Exit Function
End If

   strExc(1) = " select * from MedicineCode where " & stCon
   intR = 1
   Set rsRead = ClsLawReadRstMsg(intR, strExc(1))
   If intR = 1 Then
      '排除現在修改的記錄
      If rsRead.RecordCount = 1 And Trim(rsRead.Fields("MC01")) = Trim(txtDB(1)) Then
          RecIsExist = False
      Else
          RecIsExist = True
          MsgBox "已存在同樣條件的記錄(流水號 " & rsRead(0) & " )，請先查詢!!", vbCritical
      End If
   Else
      RecIsExist = False
   End If
   Set rsRead = Nothing
   
End Function

Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index <= 6 Then
       KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtFM2_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp As String

   '藥證號和藥品名稱不可換行
   If Index = 0 Or Index = 8 Or Index = 9 Then
      txtFM2(Index).Text = PUB_StringFilter(txtFM2(Index).Text)
   End If
   '藥證號若有阿拉伯數字統一為半形數字
   If Index = 0 Then
      txtFM2(Index).Text = PUB_ChgNumeralStyle(txtFM2(Index).Text)
   End If

   Select Case Index
   Case 2, 3, 4 '本所案號
      If txtFM2(1) <> "" And txtFM2(2) <> "" And txtFM2(1) & txtFM2(2) & txtFM2(3) & txtFM2(4) <> txtFM2(1).Tag & txtFM2(2).Tag & txtFM2(3).Tag & txtFM2(4).Tag Then
         lblData2(2).Caption = ""
         If txtFM2(3) = "" Then txtFM2(3) = "0"
         If txtFM2(4) = "" Then txtFM2(4) = "00"
         If ClsPDCheckCaseCodeIsExist(txtFM2(1), txtFM2(2), txtFM2(3), txtFM2(4), strExc(1), strExc(2), strExc(3)) = True Then
            If strExc(1) <> "" Then
               lblData2(2).Caption = strExc(1)
            ElseIf strExc(2) <> "" Then
               lblData2(2).Caption = strExc(2)
            ElseIf strExc(3) <> "" Then
               lblData2(2).Caption = strExc(3)
            End If
         End If
         txtFM2(1).Tag = txtFM2(1).Text
         txtFM2(2).Tag = txtFM2(2).Text
         txtFM2(3).Tag = txtFM2(3).Text
         txtFM2(4).Tag = txtFM2(4).Text
      End If
   Case 5 '代理人
      lblData2(0).Caption = ""
      If txtFM2(Index) <> "" Then
         If Len(txtFM2(Index)) <= 9 Then
            stCon = ChangeCustomerL(txtFM2(Index))
            txtFM2(Index) = stCon
            If ClsPDGetAgent(stCon, strTemp) Then
               lblData2(0).Caption = strTemp
            Else
               '模組已彈訊息
            End If
         End If
      End If
   Case 6 '申請人
      lblData2(1).Caption = ""
      If txtFM2(Index) <> "" Then
         If Len(txtFM2(Index)) <= 9 Then
            stCon = ChangeCustomerL(txtFM2(Index))
            txtFM2(Index) = stCon
            If ClsPDGetCustomer(stCon, strTemp) Then
               lblData2(1).Caption = strTemp
            Else
               '模組已彈訊息
            End If
         End If
      End If
   End Select
End Sub

Private Sub ClearCaseMap(ByVal bolAll As Boolean)

   If bolAll = True Then
      For Each oText In txtCode
         oText.Text = ""
         oText.Tag = ""
      Next
   End If
   
   For Each oLabel In lblData
      oLabel.Caption = ""
      oLabel.Tag = ""
   Next
   Combo1.Text = ""
   Combo1.Tag = ""
   Combo1.Clear
   m_MPKey = ""
   intLastRow1 = 0
End Sub

Private Function UpdateCaseMap(ByVal pType As String) As Boolean
Dim m_Rows  As Integer
Dim bFind As Boolean

   If m_MPKey <> "" And m_MPKey = Val(txtDB(1)) & Trim(txtCode(0) & txtCode(1) & txtCode(2) & txtCode(3)) Then
     m_Rows = 1
   Else
     m_MPKey = Val(txtDB(1)) & Trim(txtCode(0) & txtCode(1) & txtCode(2) & txtCode(3))
   End If
   
   With rsMap
      '-------刪除
      If pType = "D" Then
        If .RecordCount > 0 And m_Rows > 0 Then
           .MoveFirst
           '移動到要修改的資料
           Do While Not .EOF
              If .Fields("MPKEY") = m_MPKey Then
                 .Delete
                 Exit Do
              End If
              .MoveNext
           Loop
        End If
      '--------加入
      ElseIf pType = "A" Then
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("MPKEY") = m_MPKey Then
                bFind = True
                Exit Do
             End If
            .MoveNext
          Loop
        End If
        If bFind = False Then
           .AddNew
           .Fields("MC01") = Val(txtDB(1))
           .Fields("MC02") = Trim(txtDB(2))
           .Fields("MC03") = Trim(txtDB(3))
           .Fields("MC04") = Trim(txtDB(4))
           .Fields("MC05") = Trim(txtDB(5))
           .Fields("CASENO") = txtCode(0) & "-" & txtCode(1) & IIf(txtCode(2) & txtCode(3) = "000", "", "-" & txtCode(2) & "-" & txtCode(3))
           .Fields("CNAME") = Combo1.Tag
           .Fields("PA177") = lblData(0).Tag
           .Fields("FNAME") = lblData(1).Caption
           .Fields("APPNAME1") = lblData(3).Caption
           .Fields("MC06") = Trim(txtDB(6))
           .Fields("FANO") = lblData(0).Caption
           .Fields("APP01") = lblData(2).Caption
           .Fields("PA01") = txtCode(0)
           .Fields("PA02") = txtCode(1)
           .Fields("PA03") = txtCode(2)
           .Fields("PA04") = txtCode(3)
           .Fields("MPKEY") = m_MPKey
        End If
        .UPDATE
      End If
   End With
   
   '更新Grid
   Call SetGrd("1", MGrid1, True) '清空
     
   If rsMap.RecordCount > 0 Then
      Call SetGrd("1", MGrid1, , rsMap)
   End If
End Function
