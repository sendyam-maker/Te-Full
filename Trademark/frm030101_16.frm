VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030101_16 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(異議答辯, 評定答辯, 撤銷答辯, 補充答辯, 參加被評定, 撤銷禁止處分, 修正, 刊登廣告, 其它)"
   ClientHeight    =   6516
   ClientLeft      =   4788
   ClientTop       =   1752
   ClientWidth     =   9132
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6516
   ScaleWidth      =   9132
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7020
      TabIndex        =   20
      Top             =   50
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6072
      TabIndex        =   19
      Top             =   50
      Width           =   852
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8220
      TabIndex        =   21
      Top             =   50
      Width           =   852
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3432
      Left            =   120
      TabIndex        =   46
      Top             =   3000
      Width           =   8892
      _ExtentX        =   15685
      _ExtentY        =   6054
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "發文資料"
      TabPicture(0)   =   "frm030101_16.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(12)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(10)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label23"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label22"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label25"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label28"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label15"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label16"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label7"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label14(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblCP113(18)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCP44_2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textCP64"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "grdList"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textCP23"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCF09"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCP44"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCP18"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textPrint"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCP27"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textTM29"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textNP08"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textUargeDate"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtCP113"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "相關人"
      TabPicture(1)   =   "frm030101_16.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textCP37_1"
      Tab(1).Control(1)=   "textCP42"
      Tab(1).Control(2)=   "textCP41"
      Tab(1).Control(3)=   "textCP40"
      Tab(1).Control(4)=   "textCP39"
      Tab(1).Control(5)=   "textCP38"
      Tab(1).Control(6)=   "textCP37"
      Tab(1).Control(7)=   "textCP36"
      Tab(1).Control(8)=   "Label26"
      Tab(1).Control(9)=   "Label21"
      Tab(1).Control(10)=   "Label20"
      Tab(1).Control(11)=   "Label19"
      Tab(1).Control(12)=   "Label18"
      Tab(1).Control(13)=   "Label17"
      Tab(1).Control(14)=   "Label13"
      Tab(1).Control(15)=   "Label14(1)"
      Tab(1).ControlCount=   16
      Begin VB.TextBox txtCP113 
         Height          =   270
         Left            =   3690
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1257
         Width           =   540
      End
      Begin VB.TextBox textUargeDate 
         Height          =   264
         Left            =   3360
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Width           =   1092
      End
      Begin VB.TextBox textNP08 
         Height          =   264
         Left            =   6240
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1260
         Width           =   1812
      End
      Begin VB.TextBox textTM29 
         Height          =   264
         Left            =   7176
         MaxLength       =   1
         TabIndex        =   4
         Top             =   660
         Width           =   372
      End
      Begin VB.TextBox textCP27 
         Height          =   264
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   0
         Top             =   360
         Width           =   1092
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   7
         Top             =   1260
         Width           =   372
      End
      Begin VB.TextBox textCP18 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   5640
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   2412
      End
      Begin VB.ComboBox textCP44 
         Height          =   276
         Left            =   1080
         TabIndex        =   3
         Top             =   660
         Width           =   1404
      End
      Begin VB.TextBox textCF09 
         Height          =   264
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   6
         Top             =   960
         Width           =   612
      End
      Begin VB.TextBox textCP23 
         Height          =   264
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   5
         Top             =   960
         Width           =   372
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   1212
         Left            =   1128
         TabIndex        =   77
         Top             =   1560
         Width           =   7632
         _ExtentX        =   13462
         _ExtentY        =   2138
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
      Begin MSForms.TextBox textCP37_1 
         Height          =   870
         Left            =   -73320
         TabIndex        =   15
         Top             =   1560
         Width           =   6855
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "12091;1535"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   285
         Left            =   -73320
         TabIndex        =   14
         Top             =   1260
         Width           =   6855
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "12091;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP41 
         Height          =   285
         Left            =   -73320
         TabIndex        =   13
         Top             =   960
         Width           =   6855
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "12091;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   285
         Left            =   -73320
         TabIndex        =   12
         Top             =   660
         Width           =   6855
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "12091;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP39 
         Height          =   285
         Left            =   -73320
         TabIndex        =   18
         Top             =   2160
         Width           =   6855
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "12091;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP38 
         Height          =   285
         Left            =   -73320
         TabIndex        =   17
         Top             =   1860
         Width           =   6855
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "12091;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37 
         Height          =   285
         Left            =   -73320
         TabIndex        =   16
         Top             =   1560
         Width           =   6855
         VariousPropertyBits=   671105051
         MaxLength       =   140
         Size            =   "12091;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP36 
         Height          =   285
         Left            =   -73320
         TabIndex        =   11
         Top             =   360
         Width           =   6855
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "12091;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   555
         Left            =   1110
         TabIndex        =   10
         Top             =   2790
         Width           =   7575
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13361;979"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   270
         Left            =   2520
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   660
         Width           =   3630
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "6403;476"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCP113 
         AutoSize        =   -1  'True
         Caption         =   "工作時數:"
         Height          =   180
         Index           =   18
         Left            =   2730
         TabIndex        =   71
         Top             =   1302
         Width           =   765
      End
      Begin VB.Label Label14 
         Caption         =   "催審期限 :"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   69
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "對造案件名稱 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   68
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "下次繳年費期限 :"
         Height          =   255
         Left            =   4770
         TabIndex        =   67
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "對造日文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   66
         Top             =   1260
         Width           =   1572
      End
      Begin VB.Label Label20 
         Caption         =   "對造英文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   65
         Top             =   960
         Width           =   1572
      End
      Begin VB.Label Label19 
         Caption         =   "對造中文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   64
         Top             =   660
         Width           =   1572
      End
      Begin VB.Label Label18 
         Caption         =   "對造案件日文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   63
         Top             =   2160
         Width           =   1572
      End
      Begin VB.Label Label17 
         Caption         =   "對造案件英文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   62
         Top             =   1860
         Width           =   1572
      End
      Begin VB.Label Label13 
         Caption         =   "對造案件中文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   61
         Top             =   1560
         Width           =   1572
      End
      Begin VB.Label Label14 
         Caption         =   "對造號數 :"
         Height          =   252
         Index           =   1
         Left            =   -74880
         TabIndex        =   60
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "(Y:閉卷)"
         Height          =   180
         Left            =   7656
         TabIndex        =   59
         Top             =   660
         Width           =   648
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "是否閉卷 :"
         Height          =   180
         Left            =   6264
         TabIndex        =   58
         Top             =   672
         Width           =   816
      End
      Begin VB.Label Label9 
         Caption         =   "本案期限 :"
         Height          =   252
         Left            =   120
         TabIndex        =   57
         Top             =   1560
         Width           =   852
      End
      Begin VB.Label Label28 
         Caption         =   "進度備註 :"
         Height          =   252
         Left            =   120
         TabIndex        =   56
         Top             =   2820
         Width           =   972
      End
      Begin VB.Label Label25 
         Caption         =   "發文日 :"
         Height          =   252
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "代理人 :"
         Height          =   252
         Left            =   120
         TabIndex        =   54
         Top             =   660
         Width           =   972
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   252
         Left            =   120
         TabIndex        =   53
         Top             =   1260
         Width           =   972
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不印)"
         Height          =   252
         Left            =   1560
         TabIndex        =   52
         Top             =   1260
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "點數 :"
         Height          =   255
         Index           =   10
         Left            =   4770
         TabIndex        =   51
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "可接獲回音"
         Height          =   252
         Left            =   6360
         TabIndex        =   50
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "大約"
         Height          =   255
         Index           =   12
         Left            =   4770
         TabIndex        =   49
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "預估勝敗 :"
         Height          =   252
         Left            =   120
         TabIndex        =   48
         Top             =   960
         Width           =   972
      End
      Begin VB.Label Label12 
         Caption         =   "(1:勝 2:敗)"
         Height          =   252
         Left            =   1560
         TabIndex        =   47
         Top             =   960
         Width           =   972
      End
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   840
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   840
      Width           =   2532
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox textCP08 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2532
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1200
      TabIndex        =   75
      Top             =   2310
      Width           =   7875
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13891;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1200
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2535
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5760
      TabIndex        =   73
      Top             =   1740
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2535
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "經由網路查名查詢，代理人欄請輸：Y99999999"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   4800
      TabIndex        =   70
      Top             =   2700
      Width           =   3720
   End
   Begin VB.Label Label5 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   45
      Top             =   2340
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   252
      Left            =   120
      TabIndex        =   44
      Top             =   1140
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4800
      TabIndex        =   43
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   252
      Index           =   9
      Left            =   4800
      TabIndex        =   42
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   41
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   40
      Top             =   540
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   4800
      TabIndex        =   38
      Top             =   1140
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "發證日 :"
      Height          =   252
      Index           =   3
      Left            =   4800
      TabIndex        =   37
      Top             =   840
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "業務區別 :"
      Height          =   252
      Index           =   2
      Left            =   4800
      TabIndex        =   36
      Top             =   540
      Width           =   972
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   35
      Top             =   1740
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   252
      Index           =   4
      Left            =   4800
      TabIndex        =   34
      Top             =   2040
      Width           =   852
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   33
      Top             =   2040
      Width           =   852
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   32
      Top             =   2640
      Width           =   972
   End
End
Attribute VB_Name = "frm030101_16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/08/11 改成Form2.0 ; textCP13、textCP14、textTM23、cmbTM05、textCP44_2、textCP64、textCP36、textCP37、textCP37_1、textCP38、textCP39、textCP40、textCP41、textCP42、grdList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 收文號
Dim m_CP09 As String
' 申請國家
Dim m_TM10 As String
' 案件性質代號
Dim m_CP10 As String
' 智權人員
Dim m_CP13 As String
' 承辦人
Dim m_CP14 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存商標基本檔或服務業務基本檔檔案欄位的串列
Dim m_TMSPList() As FIELDITEM
Dim m_TMSPCount As Integer
' 儲存案件進度檔檔案欄位的串列
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer

' 宣告代理人內容結構
Private Type AGENTITEM
   aiCode As String
   aiName As String
End Type
Dim m_AgentList() As AGENTITEM
Dim m_AgentCount As Integer
'
Dim m_CurrSel As Integer
'add by nickc 2008/02/22
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_990CP09 As String 'Add By Sindy 2016/12/20
Dim m_CP07 As String 'Add By Sindy 2019/6/11
Dim m_CP31 As String    'add by sonia 2024/1/8

Private Sub cmdCancel_Click()
   frm030101_01.Show
   Unload Me
End Sub

Private Sub cmdCaseProgress_Click()
   frm030101_04.SetData 0, m_TM01, True
   frm030101_04.SetData 1, m_TM02, False
   frm030101_04.SetData 2, m_TM03, False
   frm030101_04.SetData 3, m_TM04, False
   frm030101_04.SetData 4, m_CP09, False
   frm030101_04.SetParent Me
   Me.Hide
   frm030101_04.Show
   frm030101_04.QueryData
End Sub

Private Sub cmdExit_Click()
   Unload frm030101_01
   Unload Me
End Sub

Private Sub cmdok_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 更新欄位輸入的內容
      OnUpdateField
      ' 存檔
      'edit by nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub

      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      'Add By Sindy 2012/4/5 CFT,FCT所有案件性質發文時,檢查代表圖是否存在
      'Modify by Amy 2018/07/31 ChkIsExistImg不使用
      'Call ChkIsExistImg(m_TM01, m_TM02, m_TM03, m_TM04)
      If m_CP10 <> "001" Then 'add by sonia 2018/9/6 查名不檢查代表圖
         If ChkImgByteFile(m_TM01, m_TM02, m_TM03, m_TM04) = False Then MsgBox "本案尚未放代表圖至系統！"
      End If  'end 2018/9/6
      
      '*********** 90.11.23   nick  清畫面
      'frm030101_01.radio(0).Value = True
      'frm030101_01.textCP09.Enabled = True
      'frm030101_01.textCP09.Text = ""
      'frm030101_01.textTM01.Enabled = False
      'frm030101_01.textTM01.Text = ""
      'frm030101_01.textTM02.Enabled = False
      'frm030101_01.textTM02.Text = ""
      'frm030101_01.textTM02_2.Enabled = False
      'frm030101_01.textTM02_2.Text = ""
      'frm030101_01.textTM03.Enabled = False
      'frm030101_01.textTM03.Text = "'"
      'frm030101_01.textTM04.Enabled = False
      'frm030101_01.textTM04.Text = ""
      'frm030101_01.grdList.Clear
      'frm030101_01.grdList.Rows = 2
      'frm030101_01.RefreshData
      '***********************************
      
      'Add By Sindy 2024/8/19
      If frm030101_01.bolIsEMPFlow = True Then
         frm090202_4.QueryData
      End If
      '2024/8/19 End
      'Add By Cheng 2002/04/30
      '若有未發文資料顯示警告
      If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = False Then
         'Add By Sindy 2024/8/19
         If frm030101_01.bolIsEMPFlow = True Then
            Unload frm030101_01
            frm090202_4.Show
            Unload Me
            Exit Sub
         End If
         '2024/8/19 End
      End If
      frm030101_01.Show
      ' 90.12.07 modify by louis
'      frm030101_01.Clear
      'Add By Cheng 2002/01/10
      frm030101_01.Clear1
      Unload Me
   End If
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

Private Sub Form_Activate()
'add by nickc 2005/08/23
'Modify By Sindy 2012/10/1 下列程式無意義Mark
'If (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
'   pub_ModifyCaseNum = ""
'   QueryData
'End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP08.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   
   SSTab1.Tab = 0
   
   MoveFormToCenter Me
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 收文號
      Case 0: m_CP09 = strData
   End Select
End Sub

Private Sub ClearAgentList()
   If m_AgentCount > 0 Then
      Erase m_AgentList
   End If
   m_AgentCount = 0
End Sub

Private Sub AddAgent(ByVal strAgentCode As String, ByVal strAgentName As String)
   Dim nIndex As Integer
   Dim bFind As Boolean
   bFind = False
   For nIndex = 0 To m_AgentCount - 1
      If m_AgentList(nIndex).aiCode = strAgentCode Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_AgentList(m_AgentCount + 1)
      m_AgentList(m_AgentCount).aiCode = strAgentCode
      m_AgentList(m_AgentCount).aiName = strAgentName
      m_AgentCount = m_AgentCount + 1
   End If
End Sub


' 清除商標基本檔檔案欄位串列
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiOldData = strFieldData
         m_TMSPList(nPos).fiNewData = strFieldData
         m_TMSPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_TMSPList(m_TMSPCount + 1)
      m_TMSPList(m_TMSPCount).fiName = strFieldName
      m_TMSPList(m_TMSPCount).fiOldData = strFieldData
      m_TMSPList(m_TMSPCount).fiNewData = strFieldData
      m_TMSPList(m_TMSPCount).fiType = nFieldType
      m_TMSPCount = m_TMSPCount + 1
   End If
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 清除案件進度檔檔案欄位串列
Private Sub ClearCPFieldList()
   If m_CPCount > 0 Then
      Erase m_CPList
   End If
   m_CPCount = 0
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiOldData = strFieldData
         m_CPList(nPos).fiNewData = strFieldData
         m_CPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_CPList(m_CPCount + 1)
      m_CPList(m_CPCount).fiName = strFieldName
      m_CPList(m_CPCount).fiOldData = strFieldData
      m_CPList(m_CPCount).fiNewData = strFieldData
      m_CPList(m_CPCount).fiType = nFieldType
      m_CPCount = m_CPCount + 1
   End If
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 取得商標基本檔的欄位內容
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSubSQL As String
   Dim rsSubTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_TM44 = CheckStr(rsTmp.Fields("TM44"))
      m_TM119 = CheckStr(rsTmp.Fields("TM119"))
      m_TM120 = CheckStr(rsTmp.Fields("TM120"))
      ' 審定號數
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         textTM20 = DBDATE(rsTmp.Fields("TM20"))
      End If
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 取得服務業務基本檔的欄位內容
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
        'add by nickc 2008/02/22
        m_TM44 = CheckStr(rsTmp.Fields("SP26"))
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP05")
      End If
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP06")
      End If
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP07")
      End If
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("SP12")) = False Then
         textTM20 = rsTmp.Fields("SP12")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgress()
Dim strSql As String
Dim strSubSQL As String
Dim rsTmp As New ADODB.Recordset
Dim rsSubTmp As New ADODB.Recordset
Dim strCP27 As String
Dim strCP43 As String
Dim strCP44 As String
Dim strCP45 As String
Dim nIndex As Integer
Dim bFind As Boolean

   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      
      'Add By Sindy 2019/6/11
      '法定期限
      m_CP07 = Empty
      If IsNull(rsTmp.Fields("CP07")) = False Then: m_CP07 = rsTmp.Fields("CP07")
      '2019/6/11 End
      'add by sonia 2024/1/8
      '是否新案
      m_CP31 = Empty
      If IsNull(rsTmp.Fields("CP31")) = False Then: m_CP31 = rsTmp.Fields("CP31")
      'end 2024/1/8
      
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 業務區別
      If IsNull(rsTmp.Fields("CP12")) = False Then
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      m_CP13 = "" 'Add By Sindy 2014/9/11
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      ' 承辦人員
      m_CP14 = Empty
      If IsNull(rsTmp.Fields("CP14")) = False Then
         m_CP14 = rsTmp.Fields("CP14")
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 發文日(預設為系統日)
      'edit by nickc 2006/03/17
      'textCP27 = DBDATE(Date)
      textCP27 = strSrvDate(1)
      strCP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         strCP27 = rsTmp.Fields("CP27")
      End If
      SetCPFieldOldData "CP27", strCP27, 1
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = rsTmp.Fields("CP44")
      End If
      SetCPFieldOldData "CP44", textCP44, 0
      ' 彼所案號
      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         strCP45 = rsTmp.Fields("CP45")
      End If
      SetCPFieldOldData "CP45", strCP45, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      ' 預估結果(預估勝敗)
      textCP23 = Empty
      If IsNull(rsTmp.Fields("CP23")) = False Then
         textCP23 = rsTmp.Fields("CP23")
      End If
      SetCPFieldOldData "CP23", textCP23, 0
      ' 對造號數
      textCP36 = Empty
      If IsNull(rsTmp.Fields("CP36")) = False Then
         textCP36 = rsTmp.Fields("CP36")
      End If
      SetCPFieldOldData "CP36", textCP36, 0
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            ' 對造案件名稱(中)
            textCP37_1 = Empty
            If IsNull(rsTmp.Fields("CP37")) = False Then
               textCP37_1 = rsTmp.Fields("CP37")
            End If
            SetCPFieldOldData "CP37", textCP37_1, 0
        Case Else
            ' 對造案件名稱(中)
            textCP37 = Empty
            If IsNull(rsTmp.Fields("CP37")) = False Then
               textCP37 = rsTmp.Fields("CP37")
            End If
            SetCPFieldOldData "CP37", textCP37, 0
            ' 對造案件名稱(英)
            textCP38 = Empty
            If IsNull(rsTmp.Fields("CP38")) = False Then
               textCP38 = rsTmp.Fields("CP38")
            End If
            SetCPFieldOldData "CP38", textCP38, 0
            ' 對造案件名稱(日)
            textCP39 = Empty
            If IsNull(rsTmp.Fields("CP39")) = False Then
               textCP39 = rsTmp.Fields("CP39")
            End If
            SetCPFieldOldData "CP39", textCP39, 0
        End Select
      ' 對造名稱(中)
      textCP40 = Empty
      If IsNull(rsTmp.Fields("CP40")) = False Then
         textCP40 = rsTmp.Fields("CP40")
      End If
      SetCPFieldOldData "CP40", textCP40, 0
      ' 對造名稱(英)
      textCP41 = Empty
      If IsNull(rsTmp.Fields("CP41")) = False Then
         textCP41 = rsTmp.Fields("CP41")
      End If
      SetCPFieldOldData "CP41", textCP41, 0
      ' 對造名稱(日)
      textCP42 = Empty
      If IsNull(rsTmp.Fields("CP42")) = False Then
         textCP42 = rsTmp.Fields("CP42")
      End If
      SetCPFieldOldData "CP42", textCP42, 0
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
      ' 代理人
      ClearAgentList
      'Add By Sindy 2013/5/23 若是原先有，也要加入
      If textCP44.Text <> "" Then
'         If InStr(textCP44, "-") > 0 Then
'            If ClsPDGetContact(textCP44, strCP44) Then
'               AddAgent textCP44, strCP44
'            End If
'         Else
            strCP44 = GetFAgentName(textCP44)
            AddAgent textCP44, strCP44
'         End If
      End If
      '2013/5/23 End
      '2009/2/3 modify by sonia B類收文之文件簽證711及申請英文證明304不要列入
      '2010/9/7 Modify by Sindy 文件簽證711及申請英文證明304不要列入
      strSubSQL = "SELECT CP44, MAX(CP27) AS CP27 FROM CASEPROGRESS " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null " & _
                        "AND CP10 NOT IN ('711','304') " & _
                  "GROUP BY CP44 " & _
                  "ORDER BY CP27 DESC "
      rsSubTmp.CursorLocation = adUseClient
      rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSubTmp.RecordCount > 0 Then
         rsSubTmp.MoveFirst
         ' 依序將代理人加入到系統串列中
         Do While rsSubTmp.EOF = False
            If IsNull(rsSubTmp.Fields("CP44")) = False Then
               strCP44 = GetFAgentName(rsSubTmp.Fields("CP44"))
               AddAgent rsSubTmp.Fields("CP44"), GetFAgentName(rsSubTmp.Fields("CP44"))
            End If
            rsSubTmp.MoveNext
         Loop
      End If
      rsSubTmp.Close
      ' 從系統串列中取得所有代理人並放入Combo Box中
      For nIndex = 0 To m_AgentCount - 1
         textCP44.AddItem m_AgentList(nIndex).aiCode
      Next nIndex
      ' 設定顯示為第一筆
      If textCP44.ListCount > 0 Then
         textCP44.ListIndex = 0
         textCP44_Validate False
      End If
   End If
   rsTmp.Close
   
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         ' 是否續辦欄位必須為空白
         If IsNull(rsTmp.Fields("NP06")) = False Then
            If IsEmptyText(rsTmp.Fields("NP06")) = False Then
               GoTo NextRecord
            End If
         End If
         
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         
         ' 收文號
         If IsNull(rsTmp.Fields("NP01")) = False Then
            grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("NP01")
         End If
         ' 下一程序
         If IsNull(rsTmp.Fields("NP07")) = False Then
            grdList.TextMatrix(grdList.row, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"))
            grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("NP07")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("NP08")) = False Then
            If IsEmptyText(rsTmp.Fields("NP08")) = False Then
               grdList.TextMatrix(grdList.row, 2) = ChangeWStringToTString(rsTmp.Fields("NP08"))
            End If
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               grdList.TextMatrix(grdList.row, 3) = ChangeWStringToTString(rsTmp.Fields("NP09"))
            End If
         End If
         ' 機關文號
         If IsNull(rsTmp.Fields("NP13")) = False Then
            grdList.TextMatrix(grdList.row, 4) = rsTmp.Fields("NP13")
         End If
         ' 相關人
         If IsNull(rsTmp.Fields("NP14")) = False Then
            grdList.TextMatrix(grdList.row, 5) = rsTmp.Fields("NP14")
         End If
         ' 備註
         If IsNull(rsTmp.Fields("NP15")) = False Then
            grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("NP15")
         End If
         ' 序號
         If IsNull(rsTmp.Fields("NP22")) = False Then
            grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("NP22")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/17
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/17
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strDay As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 先清除商標基本檔或服務業務基本檔欄位串列
   ClearTMSPFieldList
   ' 先清除案件進度檔欄位串列
   ClearCPFieldList
   
   ' 先取得本所案號
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 本所案號
      If IsNull(rsTmp.Fields("CP01")) = False Then: m_TM01 = rsTmp.Fields("CP01")
      If IsNull(rsTmp.Fields("CP02")) = False Then: m_TM02 = rsTmp.Fields("CP02")
      If IsNull(rsTmp.Fields("CP03")) = False Then: m_TM03 = rsTmp.Fields("CP03")
      If IsNull(rsTmp.Fields("CP04")) = False Then: m_TM04 = rsTmp.Fields("CP04")
   End If
   rsTmp.Close
    'Add By Cheng 2003/11/11
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        Me.Label13.Visible = False
        Me.textCP37.Visible = False
        Me.textCP37.Enabled = False
        Me.Label17.Visible = False
        Me.textCP38.Visible = False
        Me.textCP38.Enabled = False
        Me.Label18.Visible = False
        Me.textCP39.Visible = False
        Me.textCP39.Enabled = False
    Case Else
        Me.Label26.Visible = False
        Me.textCP37_1.Visible = False
        Me.textCP37_1.Enabled = False
    End Select
    'End
   ' 本所案號
'   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)

   ' 收文號
   textCP09 = m_CP09
   
   '2008/12/17 ADD BY SONIA
   If m_CP09 > "C" Then textPrint = "N"
   '2008/12/17 END
   
   ' 取得國家代碼
   m_TM10 = GetNationNo(m_TM01, m_TM02, m_TM03, m_TM04)
      
   ' 取得案件進度檔的欄位
   QueryCaseProgress
   
   'add by sonia 2018/9/6 查名預設不出定稿,本顯示Label27
   Label27.Visible = False
   If m_CP10 = "001" Then
      textPrint = "N": Label27.Visible = True
   'add by sonia 2024/1/8
      Label27.Caption = "經由網路查名查詢，代理人欄請輸：Y99999999"
   ElseIf m_TM01 = "CFT" And m_CP10 = "706" And m_CP31 = "Y" Then
      Label27.Visible = True: Label27.Caption = "若未委託國外代理人處理，代理人欄請輸Y00000000"
   ElseIf m_TM01 = "S" And (m_CP10 = "706" Or m_CP10 = "711" Or m_CP10 = "713") Then
      Label27.Visible = True: Label27.Caption = "若未委託國外代理人處理，代理人欄請輸Y00000000"
   'end 2024/1/8
   End If
   'end 2018/9/6
   
   'Added by Lydia 2025/10/27 各專業部的智財協作發文，都預設不出定稿
   If m_TM01 = "S" And m_CP10 = "737" Then
      textPrint = "N"
   End If
   'end 2025/10/27
   
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "CFT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   
   'Add By Sindy 2009/05/18
   ' 計算催審期限
   strDay = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
   If IsEmptyText(strDay) = False Then
      textUargeDate = strDay
   End If
   Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   
   ' 大約?可接獲回音(欄位)
   textCF09 = Empty
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF09")) = False Then
         textCF09 = rsTmp.Fields("CF09")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache  'Added by Lydia 2025/02/24
   
   'Add By Cheng 2002/07/19
   Set frm030101_16 = Nothing
End Sub

' 預估勝敗
Private Sub textCP23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP23) = False Then
      Select Case textCP23
         Case "1", "2":
         Case Else
            strTit = "檢核資料"
            strMsg = "預估勝敗只可輸入1或2"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP23_GotFocus
      End Select
   End If
End Sub

' 對造號數
Private Sub textCP36_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modify by Amy 2022/09/29 原20 放寬至200
   If CheckLengthIsOK(textCP36, textCP36.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造號數內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP36_GotFocus
   End If
End Sub

Private Sub textCP37_1_GotFocus()
    TextInverse Me.textCP37_1
End Sub

Private Sub textCP37_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP37_1, 140) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件名稱內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textCP37_1_GotFocus
   End If
End Sub

' 對造案件中文名稱
Private Sub textCP37_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP37, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件中文名稱內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'textCP37_GotFocus
   End If
End Sub

' 對造案件英文名稱
Private Sub textCP38_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP38, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件英文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP38_GotFocus
   End If
End Sub

' 對造案件日文名稱
Private Sub textCP39_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP39, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件日文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP39_GotFocus
   End If
End Sub

' 對造中文名稱
Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP40, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造中文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP40_GotFocus
   End If
End Sub

' 對造英文名稱
Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP41, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造英文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP41_GotFocus
   End If
End Sub

' 對造日文名稱
Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP42, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造日文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP42_GotFocus
   End If
End Sub

Private Sub textCP44_Click()
   textCP44_2 = m_AgentList(textCP44.ListIndex).aiName
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM29_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否閉卷
Private Sub textTM29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textTM29) = False Then
      Select Case textTM29
         Case "", " ":
         Case "Y":
            strTit = "閉卷"
            strMsg = "請確認是否閉卷"
            nResponse = MsgBox(strMsg, vbYesNo, strTit)
            If nResponse = vbNo Then
               textTM29 = Empty
            End If
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入Y或空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM29_GotFocus
      End Select
   End If
End Sub

' 發文日
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP27) = False Then
      ' 發文日日期不正確
      If CheckIsDate(textCP27, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' 發文日日期不可超過系統日
      If Val(DBDATE(textCP27)) > Val(strSrvDate(1)) Then 'edit by nickc 2006/03/17 原先為 DBDATE(Date)) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "發文日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      'Add By Sindy 2009/05/18
      ' 計算催審期限
      If Me.textCP27.Tag <> Me.textCP27.Text Then 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
            strDay = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
            If IsEmptyText(strDay) = False Then
               textUargeDate = strDay
            End If
      End If
      Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   End If
EXITSUB:
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCP44_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCP44.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

' 代理人
Private Sub textCP44_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTempName As String   '2010/11/24 add by sonia
   
   Cancel = False
   'Add By Cheng 2002/03/08
   If m_TM10 <> 台灣國家代號 Then
      If Len(Me.textCP44.Text) <= 0 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Cancel = True
         Exit Sub
      End If
   End If
   
   If textCP44.ListIndex >= 0 Then
      textCP44 = m_AgentList(textCP44.ListIndex).aiCode
   End If
   textCP44_2 = Empty
   'add by nickc 2007/10/22 修正錯誤
   textCP44 = Mid(textCP44, 1, 9)
   
   If IsEmptyText(textCP44) = False Then
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      '2010/11/24 modify by sonia 取消basQuery的GetFAgentNameAndState
      'Dim oState As Boolean
      'oState = True
      ''textCP44_2 = GetFAgentName(textCP44)
      'textCP44_2 = GetFAgentNameAndState(textCP44, oState)
      'If oState = False Then
      '      Cancel = True
      '      Exit Sub
      'End If
      If PUB_GetAgentNameAndState(m_TM01, textCP44.Text, strTempName) Then
         textCP44_2 = strTempName
      Else
         textCP44_2.Text = ""
         If strTempName <> "" Then
            Cancel = True
            Exit Sub
         End If
      End If
      '2010/11/24 end
      If IsEmptyText(textCP44_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44_GotFocus
      Else
         ' 依所選擇的代理人找出案件進度檔中其收文日最大的一筆其彼所案號更新到畫面上的彼所案號欄位
         strSql = "SELECT CP45 FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP44 = '" & textCP44 & "' AND " & _
                        "CP05 IN (SELECT MAX(CP05) FROM CASEPROGRESS " & _
                                 "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                       "CP02 = '" & m_TM02 & "' AND " & _
                                       "CP03 = '" & m_TM03 & "' AND " & _
                                       "CP04 = '" & m_TM04 & "' AND " & _
                                       "CP44 = '" & textCP44 & "')"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CP45")) = False Then
               textTM45 = rsTmp.Fields("CP45")
            End If
         End If
         rsTmp.Close
      End If
   End If
   Set rsTmp = Nothing
End Sub

' 下次繳年費期限
Private Sub textNP08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textNP08) = False Then
      ' 下次繳年費期限日期不正確
      If CheckIsDate(textNP08, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的下次繳年費期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP08_GotFocus
      End If
   End If
End Sub

' 列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
   ' 預估結果
   SetCPFieldNewData "CP23", textCP23
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' 對造號數
   SetCPFieldNewData "CP36", textCP36
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        ' 對造案件名稱
        SetCPFieldNewData "CP37", textCP37_1
    Case Else
        ' 對造案件名稱(中)
        SetCPFieldNewData "CP37", textCP37
        ' 對造案件名稱(英)
        SetCPFieldNewData "CP38", textCP38
        ' 對造案件名稱(日)
        SetCPFieldNewData "CP39", textCP39
    End Select
   ' 對造名稱(中)
   SetCPFieldNewData "CP40", textCP40
   ' 對造名稱(英)
   SetCPFieldNewData "CP41", textCP41
   ' 對造名稱(日)
   SetCPFieldNewData "CP42", textCP42
   ' 代理人代號
   If IsEmptyText(textCP44) = False Then
      SetCPFieldNewData "CP44", textCP44 & String(9 - Len(textCP44), "0")
      'add by nickc 2008/02/22
      m_CP44New = textCP44 & String(9 - Len(textCP44), "0")
   Else
      SetCPFieldNewData "CP44", textCP44
      'add by nickc 2008/02/22
      m_CP44New = textCP44
   End If
   ' 彼所案號
   SetCPFieldNewData "CP45", textTM45
   ' 進度備註
   SetCPFieldNewData "CP64", textCP64
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
End Sub

'edit b nick 2004/11/03
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strTmp As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim strNP07 As String
   Dim strNP08 As String
   'Add By Cheng 2002/06/05
   Dim strNP09 As String
   Dim strNP22 As String
   Dim strNP10 As String 'Add By Sindy 2014/9/11
   
'911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   'Modified by Lydia 2016/03/11 +案號
   'Call GetNP69("", m_TM10, m_CP13, strNP10) 'Add By Sindy 2014/9/11
   'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
   Call GetNA69("", m_TM10, m_CP13, strNP10, m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 更新案件進度檔
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
            strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_CPList(nIndex).fiName & " = " & m_CPList(nIndex).fiNewData
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
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
   
   'Added by Lydia 2025/02/24 TIPS分配比例管制：與ACS案有關之智財協作發文時一併產生TIPS案請款階段分配比例
   If m_TM01 = "S" Then
      Call PUB_InsertACS_TIPS_Rate(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_CP10)
   End If
   'end 2025/02/24
      
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新基本檔的是否閉卷欄位
   Select Case m_TM01
      ' 系統類別為CFT的為儲存商標基本檔
      Case "CFT":
         strSql = "UPDATE TradeMark SET TM29 = '" & textTM29 & "' " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
         cnnConnection.Execute strSql
      Case Else:
         strSql = "UPDATE ServicePractice SET SP15 = '" & textTM29 & "' " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "' "
         cnnConnection.Execute strSql
   End Select
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 有輸入下次繳年費期限時, 新增一筆繳年費記錄的資料到下一程序檔
   If IsEmptyText(textNP08) = False Then
      'Add By Cheng 2002/06/05
      '法定期限
      strNP09 = textNP08
      '本所期限 ( = 法定期限 - 1年)
        'Modify By Cheng 2003/09/02
'      strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP09)) - 1, Val(DBMONTH(strNP09)), Val(DBDAY(strNP09)))))
      strNP08 = DBDATE(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(strNP09))))
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      strNP22 = GetNextProgressNo()
      'Modify By Cheng 2002/06/05
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & "708" & "," & _
'                          DBDATE(textNP08) & "," & DBDATE(textNP08) & ",'" & m_CP13 & "'," & strNP22 & ")"
        'Modify By Cheng 2003/11/24
        '重抓智權人員
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & "708" & "," & _
'                          DBDATE(strNP08) & "," & DBDATE(strNP09) & ",'" & m_CP13 & "'," & strNP22 & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & "708" & "," & _
                          DBDATE(strNP08) & "," & DBDATE(strNP09) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 繳年費不印
      ' 列印國內案件接洽及結案記錄單
      ' g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有定義代理人收達天數時, 新增一筆收達的記錄到下一程序檔
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CF23")) = False Then
         strNP07 = "997"
         strNP08 = DBDATE(textCP27)
        'Modify By Cheng 2003/09/02
'         strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP08)), Val(DBMONTH(strNP08)), Val(DBDAY(strNP08)) + Val(rsTmp.Fields("CF23")))))
         strNP08 = DBDATE(DateAdd("d", Val(rsTmp.Fields("CF23")), ChangeWStringToWDateString(DBDATE(strNP08))))
         'Add By Sindy 2019/6/11 檢查期限是否正確
         strNP08 = PUB_T997998LimitDate(strNP08, m_CP07, 1)
         '2019/6/11 END
         strNP22 = GetNextProgressNo()
         '92.10.6 MODIFY BY SONIA
         'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
         '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
         '                   strNP08 & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
         'Modify By Sindy 2014/9/11 m_CP14=>strNP10
         'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            strNP08 & "," & strNP08 & ",'" & strNP10 & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & strNP10 & "'," & strNP22 & ")"
         '92'10'6 END
         cnnConnection.Execute strSql
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing

   'Added by Lydia 2024/07/09 判斷案件國家收費表內有設定提申期限(天)CF11，要加掛提申(998)期限；
   Call Pub_GetCF11to998(m_TM10, m_TM01, m_TM02, m_TM03, m_TM04, m_CP07, m_CP09, m_CP10, m_CP14, textCP27)
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新使用者所選取的本案期限資料
   For nIndex = 1 To grdList.Rows - 1
      ' 判斷該列是否有被選取
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         strNP07 = grdList.TextMatrix(nIndex, 8)
         strNP22 = grdList.TextMatrix(nIndex, 9)
         strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND " & _
                        "NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND " & _
                        "NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = " & strNP07 & " AND " & _
                        "NP22 = " & strNP22 & " "
         cnnConnection.Execute strSql
      End If
   Next nIndex
   
   'Add By Sindy 2009/05/18
   ' 若有輸入催審期限時, 新增一筆催審的記錄到下一程序檔
   If IsEmptyText(textUargeDate) = False Then
      strNP07 = "305"
      strNP22 = GetNextProgressNo()
      'Modify By Sindy 2014/9/11 m_CP14=>strNP10
      'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                          DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strNP10 & "'," & strNP22 & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                          PUB_GetWorkDay1(textUargeDate, True) & "," & DBDATE(textUargeDate) & ",'" & strNP10 & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
   
   Pub_UpdateEndModCash m_TM01, m_TM02, m_TM03, m_TM04 'Added by Morgan 2022/6/23
   
'911106 nick transation
    cnnConnection.CommitTrans
   
     'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
    PUB_CheckEMail m_CP44New, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22
    
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 列印定稿
   If textPrint <> "N" Then
      PrintLetter
   End If

    Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
   OnSaveData = False
End Function

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 10
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "下一程序"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "本所期限"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "法定期限"
   grdList.ColWidth(3) = 1000
   grdList.col = 4
   grdList.Text = "機關文號"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "相關人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "備註"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "收文號"
   grdList.ColWidth(7) = 0
   grdList.col = 8
   grdList.Text = "下一程序代號"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "序號"
   grdList.ColWidth(9) = 0
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

Private Sub grdList_Click()
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         If grdList.TextMatrix(grdList.row, 0) = "V" Then
            grdList.TextMatrix(grdList.row, 0) = Empty
         Else
            grdList.TextMatrix(grdList.row, 0) = "V"
         End If
      End If
   End If
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
      Dim nOldCol As Integer
      nOldCol = grdList.col
      grdList.col = 1
      If grdList.CellBackColor <> &H8000000D Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
         Next nCol
      End If
      grdList.col = nOldCol
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

' 檢查欄位是否都已輸入或是輸入的值是否正確
Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim ii As Integer   'add by sonia 2016/11/17

   CheckDataValid = False
   'add by nickc 2008/05/01
   If IsDebt(m_TM10, textCP09) Then
        strTit = "警告！禁止發文！"
        strMsg = "未收款且無 預定收款日 請轉告智權同仁！！"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        GoTo EXITSUB
   End If
   ' 發文日
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' 代理人
   If IsEmptyText(textCP44) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入代理人"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP44.SetFocus
      GoTo EXITSUB
   End If
   ' 預估勝敗不可為空白
   'If IsEmptyText(textCP23) = True Then
   '   strTit = "檢核資料"
   '   strMsg = "請輸入預估勝敗"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   GoTo ExitSub
   'End If
   
    'add by nickc 2006/03/17 加入驗證
    Dim Cancel As Boolean
    Cancel = False
    textCP27_Validate Cancel
    If Cancel = True Then GoTo EXITSUB
    
   'add by sonia 2016/11/17 +催審305(T-201684)
   If m_CP10 = "305" Then
      For ii = 1 To Me.grdList.Rows - 1
         If Me.grdList.TextMatrix(ii, 0) <> "" Then
            MsgBox "案件性質為<催審>，不可點選下一程序期限資料!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
         End If
      Next ii
   End If
   'end 2016/11/17
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textTM29_GotFocus()
   InverseTextBox textTM29
End Sub

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
End Sub

Private Sub textCP23_GotFocus()
   InverseTextBox textCP23
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP36_GotFocus()
   InverseTextBox textCP36
End Sub

Private Sub textCP37_GotFocus()
   InverseTextBox textCP37
End Sub

Private Sub textCP38_GotFocus()
   InverseTextBox textCP38
End Sub

Private Sub textCP39_GotFocus()
   InverseTextBox textCP39
End Sub

Private Sub textCP40_GotFocus()
   InverseTextBox textCP40
End Sub

Private Sub textCP41_GotFocus()
   InverseTextBox textCP41
End Sub

Private Sub textCP42_GotFocus()
   InverseTextBox textCP42
End Sub

Private Sub textCP44_GotFocus()
   InverseTextBox textCP44
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textNP08_GotFocus()
   InverseTextBox textNP08
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strSql As String
   
   ' 系統類別為CFT
   If m_TM01 = "CFT" Then
      Select Case m_CP10
         ' 申請
         Case "101":
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "01", strUserNum
            ' 回音
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                     "','回音','" & textCF09 & "')"
            cnnConnection.Execute strSql
         ' 其它
         Case Else:
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "03", strUserNum
      End Select
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   ' 系統類別為CFT
   If m_TM01 = "CFT" Then
      Select Case m_CP10
         ' 申請
         Case "101":
            ' 列印定稿
            NowPrint m_CP09, "01", "01", False, strUserNum, 0
         ' 其它
         Case Else:
            ' 列印定稿
            NowPrint m_CP09, "01", "03", False, strUserNum, 0
      End Select
   End If
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
   If Me.textCP23.Enabled = True Then
      Cancel = False
      textCP23_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP27.Enabled = True Then
      Cancel = False
      textCP27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP36.Enabled = True Then
      Cancel = False
      textCP36_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP37.Enabled = True Then
      Cancel = False
      textCP37_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP37_1.Enabled = True Then
      Cancel = False
      textCP37_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP38.Enabled = True Then
      Cancel = False
      textCP38_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP39.Enabled = True Then
      Cancel = False
      textCP39_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP40.Enabled = True Then
      Cancel = False
      textCP40_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP41.Enabled = True Then
      Cancel = False
      textCP41_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP42.Enabled = True Then
      Cancel = False
      textCP42_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP44.Enabled = True Then
      Cancel = False
      textCP44_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textNP08.Enabled = True Then
      Cancel = False
      textNP08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textPrint.Enabled = True Then
      Cancel = False
      textPrint_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM29.Enabled = True Then
      Cancel = False
      textTM29_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2009/05/18
   If Me.textUargeDate.Enabled = True Then
      Cancel = False
      textUargeDate_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2016/12/20
   '檢查有設定副本收受人需提醒並新增信函副本B類收文
   m_990CP09 = ""
   If textPrint = "N" Then '不印定稿
      If PUB_ChkCC(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_990CP09) = False Then
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   '2016/12/20 END
    'Added by Lydia 2021/06/04 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
    If Pub_ChkACS112isNull(m_TM01, m_TM02, m_TM03, m_TM04, txtCP113) = True Then
        txtCP113.SetFocus
        txtCP113_GotFocus
        Exit Function
    End If
    'end 2021/06/04
   
   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
   
   'Added by Morgan 2022/6/23
   '與CFT有關的S案件性質詢問是否可結餘--洪琬姿
   bolEndModCash = False
   If m_TM01 = "S" And m_TM10 <> "000" And (m_CP10 = "001" Or m_CP10 = "706" Or m_CP10 = "711") Then
      'Pub_EndModCashMsg 有控制要已發文故直接問
      bolEndModCash = IIf(MsgBox("此案號是否可以結餘？", vbYesNo, "上可結餘日期") = vbYes, True, False)
   End If
   'end 2022/6/23
   
   TxtValidate = True
End Function

'Add By Sindy 2009/05/18
' 催審期限
Private Sub textUargeDate_GotFocus()
   InverseTextBox textUargeDate
End Sub

Private Sub textUargeDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textUargeDate) = False Then
      If CheckIsDate(textUargeDate, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "催審期限日期不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textUargeDate_GotFocus
      End If
   End If
End Sub

'Added by Lydia 2021/06/04
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/06/04
Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
