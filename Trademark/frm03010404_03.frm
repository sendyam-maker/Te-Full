VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03010404_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "發回補理由/發回補答辨"
   ClientHeight    =   6144
   ClientLeft      =   144
   ClientTop       =   996
   ClientWidth     =   9132
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6144
   ScaleWidth      =   9132
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Left            =   6840
      TabIndex        =   18
      Top             =   0
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5880
      TabIndex        =   17
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8100
      TabIndex        =   19
      Top             =   0
      Width           =   912
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   372
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1907
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1299
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2211
      Width           =   2532
   End
   Begin VB.TextBox textCP45 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1907
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2211
      Width           =   2412
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   678
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   372
      Width           =   2532
   End
   Begin VB.TextBox textTM16 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   676
      Width           =   2532
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3195
      Left            =   150
      TabIndex        =   43
      Top             =   2910
      Width           =   8895
      _ExtentX        =   15706
      _ExtentY        =   5652
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   420
      TabCaption(0)   =   "一般資料"
      TabPicture(0)   =   "frm03010404_03.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label26"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label24"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label25"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textCP14_2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label23"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label22"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "grdList"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textCP48"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textCP14"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textCP26"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCP08"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textCP07"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textCP06"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textCF15_2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCF15"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textPrint"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "對造名稱"
      TabPicture(1)   =   "frm03010404_03.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textCP37_1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "textCP40"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "textCP41"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "textCP42"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "textCP64"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "textCP36"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "textCP37"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "textCP38"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "textCP39"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label11"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label27"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label21"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label20"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label19"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label18"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label17"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label13"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label12"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   4
         Top             =   960
         Width           =   732
      End
      Begin VB.ComboBox textCF15 
         Height          =   276
         Left            =   5400
         TabIndex        =   1
         Top             =   360
         Width           =   1332
      End
      Begin VB.TextBox textCF15_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   360
         Width           =   1812
      End
      Begin VB.TextBox textCP06 
         Height          =   264
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   2
         Top             =   660
         Width           =   2412
      End
      Begin VB.TextBox textCP07 
         Height          =   264
         Left            =   5400
         MaxLength       =   8
         TabIndex        =   3
         Top             =   660
         Width           =   2412
      End
      Begin VB.TextBox textCP08 
         Height          =   264
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   0
         Top             =   360
         Width           =   2412
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   5
         Top             =   960
         Width           =   372
      End
      Begin VB.TextBox textCP14 
         Height          =   264
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1260
         Width           =   732
      End
      Begin VB.TextBox textCP48 
         Height          =   264
         Left            =   5400
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1260
         Width           =   2412
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   1116
         Left            =   1200
         TabIndex        =   72
         Top             =   1584
         Width           =   7512
         _ExtentX        =   13250
         _ExtentY        =   1969
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
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   225
         Left            =   120
         TabIndex        =   71
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不印)"
         Height          =   225
         Left            =   2040
         TabIndex        =   70
         Top             =   990
         Width           =   855
      End
      Begin MSForms.TextBox textCP37_1 
         Height          =   870
         Left            =   -73260
         TabIndex        =   9
         Top             =   630
         Width           =   7005
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "12347;1535"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   285
         Left            =   -73260
         TabIndex        =   13
         Top             =   1530
         Width           =   7005
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "12356;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP41 
         Height          =   285
         Left            =   -73260
         TabIndex        =   14
         Top             =   1830
         Width           =   7005
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   285
         Left            =   -73260
         TabIndex        =   15
         Top             =   2130
         Width           =   7005
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   645
         Left            =   -73260
         TabIndex        =   16
         Top             =   2460
         Width           =   7005
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12347;1138"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP36 
         Height          =   285
         Left            =   -73260
         TabIndex        =   8
         Top             =   330
         Width           =   7005
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37 
         Height          =   285
         Left            =   -73260
         TabIndex        =   10
         Top             =   630
         Width           =   7000
         VariousPropertyBits=   671105051
         MaxLength       =   140
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP38 
         Height          =   285
         Left            =   -73260
         TabIndex        =   11
         Top             =   930
         Width           =   7000
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP39 
         Height          =   285
         Left            =   -73260
         TabIndex        =   12
         Top             =   1230
         Width           =   7000
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "12347;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   285
         Left            =   1980
         TabIndex        =   69
         Top             =   1260
         Width           =   1695
         VariousPropertyBits=   671105055
         Size            =   "2990;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         Caption         =   "對造案件名稱 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   62
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label Label27 
         Caption         =   "進度備註 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   61
         Top             =   2460
         Width           =   972
      End
      Begin VB.Label Label21 
         Caption         =   "對造日文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   60
         Top             =   2160
         Width           =   1572
      End
      Begin VB.Label Label20 
         Caption         =   "對造英文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   59
         Top             =   1860
         Width           =   1572
      End
      Begin VB.Label Label19 
         Caption         =   "對造中文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   58
         Top             =   1560
         Width           =   1572
      End
      Begin VB.Label Label18 
         Caption         =   "對造案件日文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   57
         Top             =   1260
         Width           =   1572
      End
      Begin VB.Label Label17 
         Caption         =   "對造案件英文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   56
         Top             =   960
         Width           =   1572
      End
      Begin VB.Label Label13 
         Caption         =   "對造案件中文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   55
         Top             =   660
         Width           =   1572
      End
      Begin VB.Label Label12 
         Caption         =   "對造號數 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   54
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label10 
         Caption         =   "下一程序 :"
         Height          =   252
         Left            =   4440
         TabIndex        =   53
         Top             =   360
         Width           =   852
      End
      Begin VB.Label Label7 
         Caption         =   "本所期限 :"
         Height          =   252
         Left            =   120
         TabIndex        =   52
         Top             =   660
         Width           =   852
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   252
         Left            =   4440
         TabIndex        =   51
         Top             =   660
         Width           =   852
      End
      Begin VB.Label Label8 
         Caption         =   "機關文號 :"
         Height          =   252
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label9 
         Caption         =   "(N:不算)"
         Height          =   225
         Left            =   6240
         TabIndex        =   49
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "是否算案件數 :"
         Height          =   225
         Left            =   4440
         TabIndex        =   48
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "承辦人 :"
         Height          =   252
         Left            =   120
         TabIndex        =   47
         Top             =   1260
         Width           =   852
      End
      Begin VB.Label Label26 
         Caption         =   "承辦期限 :"
         Height          =   252
         Left            =   4440
         TabIndex        =   46
         Top             =   1260
         Width           =   852
      End
      Begin VB.Label Label5 
         Caption         =   "本案期限 :"
         Height          =   252
         Left            =   120
         TabIndex        =   45
         Top             =   1620
         Width           =   852
      End
   End
   Begin MSForms.TextBox textCP40S 
      Height          =   285
      Left            =   1200
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   1596
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
      Left            =   5700
      TabIndex        =   67
      Top             =   1608
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
   Begin MSForms.TextBox textCP14S 
      Height          =   285
      Left            =   1200
      TabIndex        =   66
      Top             =   2504
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
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   1290
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
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1200
      TabIndex        =   64
      Top             =   990
      Width           =   7605
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13414;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
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
      Left            =   3780
      TabIndex        =   63
      Top             =   420
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   42
      Top             =   372
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   41
      Top             =   984
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   40
      Top             =   1290
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   39
      Top             =   1902
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   4740
      TabIndex        =   38
      Top             =   1323
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   252
      Index           =   8
      Left            =   4740
      TabIndex        =   37
      Top             =   2227
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   252
      Index           =   9
      Left            =   4740
      TabIndex        =   36
      Top             =   1925
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Index           =   10
      Left            =   120
      TabIndex        =   35
      Top             =   2208
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4740
      TabIndex        =   34
      Top             =   1624
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   13
      Left            =   120
      TabIndex        =   33
      Top             =   678
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   32
      Top             =   372
      Width           =   852
   End
   Begin VB.Label Label4 
      Caption         =   "目前准駁 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   31
      Top             =   673
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "對造名稱 :"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   30
      Top             =   1596
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人 :"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   2520
      Width           =   852
   End
End
Attribute VB_Name = "frm03010404_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/20 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/08/13 改成Form2.0 ;textTM23、cmbTM05、textCP13、textCP14_2、textCP14S、textCP64、textCP40、textCP40S、textCP36、textCP37、textCP37_1、textCP38、textCP39、grd1改字型=新細明體-ExtB
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
' 來函收文日
Dim m_CP05 As String
' 原本所期限
Dim m_CP06 As String
' 收文號
Dim m_CP09 As String
' 原案件性質
Dim m_CP10 As String
' 原智權人員代號
Dim m_CP13 As String
' 原業務區
Dim m_CP12 As String
' 原承辦人員代號
Dim m_CP14 As String
' 國家代碼
Dim m_TM10 As String
'
Dim m_CurrSel As Integer
'add by nickc 2005/05/31
Dim IsAppNpData As Boolean
Dim SeekNewCp09 As String
Dim oClsPrtForm001 As New ClsPrtForm001
'Add By Sindy 2023/5/2
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2023/5/2 END


'Add By Sindy 2023/5/2
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm03010404_02.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm03010404_02
   Unload frm03010404_01
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 存檔
      'edit by  nick 2004/11/03
      'OnSaveData
   'add by nickc 2005/05/31
   IsAppNpData = False
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
      'Added by Lydia 2021/12/08
      ' 列印定稿
      If textPrint <> "N" Then
         PrintLetter
      End If
      'end 2021/12/08
      
      'add by nickc 2005/05/31
      If IsAppNpData Then
         'add by nickc 2005/09/27
         If MsgBox("準備列印回覆單!!!", vbExclamation + vbOKCancel) = vbOK Then
            Call oClsPrtForm001.PrintReturnSheet(SeekNewCp09, textCF15, DBDATE(textCP07), False)
         End If
      End If
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      'Add By Sindy 2023/5/2
      If Me.m_strIR01 <> "" Then
         Unload frm03010404_02
         Unload frm03010404_01
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
      Else
      '2023/5/2 END
         Unload Me
         Unload frm03010404_02
         frm03010404_01.Show
      End If
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM16.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP14S.BackColor = &H8000000F
   textCP40S.BackColor = &H8000000F
   textCP45.BackColor = &H8000000F
   
   textCF15_2.BackColor = &H8000000F
   
   ' 下一程序
   textCF15.AddItem "補充理由"
   textCF15.AddItem "補充答辯"
   
   'Added by Lydia 2021/12/08 CFT預設出通用定稿
   If m_TM01 = "CFT" Then
      textPrint = ""
   Else
      textPrint = "N"
   End If
   'end 2021/12/08
   
   MoveFormToCenter Me
   
   Me.SSTab1.Tab = 0 'Added by Lydia 2021/09/03
   
   'Add By Sindy 2023/5/2
   m_strIR01 = frm03010404_02.m_strIR01
   m_strIR02 = frm03010404_02.m_strIR02
   m_strIR03 = frm03010404_02.m_strIR03
   m_strIR04 = frm03010404_02.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2023/5/2 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      m_CP09 = Empty
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
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
   End Select
End Sub

' 取得承辦期限
Private Function GetCP48() As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strDate As String
   Dim strDay As String
   Dim strTemp As String
   
   GetCP48 = Empty
   ' 承辦期限的日期應為來函收文日加上工作天數
   ' 工作天數由系統別+國家代碼+案件性質(勝訴)搜尋案件收費表的工作天數
   ' 若有值才做檢查
   If IsEmptyText(textCP48) = False Then
''''edit by nickc 2007/10/12 改抓有時效性的
''''      strSQL = "SELECT * FROM CaseFee " & _
''''               "WHERE CF01 = '" & m_TM01 & "' AND " & _
''''                     "CF02 = '" & m_TM10 & "' AND " & _
''''                     "CF03 = '" & "1003" & "' "
''''      rsTmp.CursorLocation = adUseClient
''''      rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
''''      If rsTmp.RecordCount > 0 Then
''''         rsTmp.MoveFirst
''''         If IsNull(rsTmp.Fields("CF04")) = False Then
''''            If IsEmptyText(rsTmp.Fields("CF04")) = False Then
''''               strDay = rsTmp.Fields("CF04")
''''               strDate = DBDATE(m_CP05)
''''               ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''               'strTemp = DBDATE(Format(DateSerial(Val(DBYEAR(strDate)), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) + Val(strDay))))
''''               strTemp = DBDATE(CompWorkDay(Val(strDay), DBDATE(strDate), 0))
''''               GetCP48 = strTemp
''''            End If
''''         End If
''''      End If
''''      rsTmp.Close
        textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1003", DBDATE(m_CP05), DBDATE(textCP06), textCP09)
   End If
   Set rsTmp = Nothing
End Function

' 讀取商標基本檔
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim strSub As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSub As ADODB.Recordset
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 審定號數
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' 目前准駁
      If IsNull(rsTmp.Fields("TM16")) = False Then
         Select Case rsTmp.Fields("TM16")
            Case "1": textTM16 = "准"
            Case "2": textTM16 = "駁"
         End Select
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("tm29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取案件進度檔
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSubSQL As String
   Dim rsSubTmp As ADODB.Recordset
   Dim bCP40 As Boolean
   Dim strTemp As String
   
   ' 來函收文日
   textCP05S = m_CP05
   
   ' 先設定承辦期限為可輸入
   textCP48.BackColor = &H80000005
   textCP48.Locked = False
   textCP48.TabStop = True
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 原本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then
         m_CP06 = rsTmp.Fields("CP06")
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         textCP09 = rsTmp.Fields("CP09")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 承辦人
      'Modified by Lydia 2016/03/11 CFT改成模組判斷
      'Modified by Lydia 2016/03/25 全部套用
      'If m_TM01 = "CFT" Then
        m_CP14 = rsTmp.Fields("CP14")
        Dim strNA69 As String
        'Modified by Lydia 2016/08/17 改抓NA69
        'Call GetNP69("", "", "", strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        'Modified by Lydia 2016/11/21 傳入智權人員代號
        'Call GetNP69("", m_TM10, "", strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
        Call GetNA69("", m_TM10, "" & rsTmp.Fields("CP13"), strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        
        textCP14S = GetStaffName(m_CP14, True)
        textCP14 = strNA69
        textCP14_2 = GetStaffName(strNA69)
'      Else
'      'end 2016/03/11
'        If IsNull(rsTmp.Fields("CP14")) = False Then
'           m_CP14 = rsTmp.Fields("CP14")
'           textCP14S = GetStaffName(rsTmp.Fields("CP14"))
'           textCP14 = rsTmp.Fields("CP14")
'           textCP14_2 = GetStaffName(rsTmp.Fields("CP14"))
'        End If
'      End If
      'end 2016/03/25
      ' 彼所案號
      If IsNull(rsTmp.Fields("CP45")) = False Then
         textCP45 = rsTmp.Fields("CP45")
      End If
      ' 智權人員
      m_CP13 = Empty
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True) 'Modified by Lydia 2016/03/25 離職人員也顯示
      End If
      '業務區   nick 91.08.22
      m_CP12 = Empty
      If IsNull(rsTmp.Fields("cp12")) = False Then
        m_CP12 = rsTmp.Fields("cp12")
      End If
      ' 對照名稱 (無中文取英文, 無英文取日文)
      bCP40 = False
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP40")) = False Then
            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
               textCP40S = rsTmp.Fields("CP40")
               bCP40 = True
            End If
         End If
      End If
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP41")) = False Then
            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
               textCP40S = rsTmp.Fields("CP41")
               bCP40 = True
            End If
         End If
      End If
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP42")) = False Then
            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
               textCP40S = rsTmp.Fields("CP42")
               bCP40 = True
            End If
         End If
      End If
      ' 對造號數
      If IsNull(rsTmp.Fields("CP36")) = False Then
         textCP36 = rsTmp.Fields("CP36")
      End If
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            ' 對造案件名稱(中)
            If IsNull(rsTmp.Fields("CP37")) = False Then
               textCP37_1 = rsTmp.Fields("CP37")
            End If
        Case Else
            ' 對造案件名稱(中)
            If IsNull(rsTmp.Fields("CP37")) = False Then
               textCP37 = rsTmp.Fields("CP37")
            End If
            ' 對造案件名稱(英)
            If IsNull(rsTmp.Fields("CP38")) = False Then
               textCP38 = rsTmp.Fields("CP38")
            End If
            ' 對造案件名稱(日)
            If IsNull(rsTmp.Fields("CP39")) = False Then
               textCP39 = rsTmp.Fields("CP39")
            End If
        End Select
      ' 對造名稱(中)
      If IsNull(rsTmp.Fields("CP40")) = False Then
         textCP40 = rsTmp.Fields("CP40")
      End If
      ' 對造名稱(英)
      If IsNull(rsTmp.Fields("CP41")) = False Then
         textCP41 = rsTmp.Fields("CP41")
      End If
      ' 對造名稱(日)
      If IsNull(rsTmp.Fields("CP42")) = False Then
         textCP42 = rsTmp.Fields("CP42")
      End If
      
      ' 承辦期限(若計算結果超過本所期限), 則設定為本所期限且不可輸入
      strTemp = GetCP48()
      If IsEmptyText(strTemp) = False And IsEmptyText(m_CP06) = False Then
         If Val(strTemp) > Val(m_CP06) Then
            textCP48 = m_CP06
            textCP48.BackColor = &H8000000F
            textCP48.Locked = True
            textCP48.TabStop = False
         End If
      End If
      
      Set rsSubTmp = Nothing
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Public Sub QueryData()
   Dim strDay As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   
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
        Me.Label11.Visible = False
        Me.textCP37_1.Visible = False
        Me.textCP37_1.Enabled = False
    End Select
   ' 讀取基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "CFT":
         QueryTradeMark
   End Select
   
   ' 讀取案件進度檔
   QueryCaseProgress
   
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
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
      'Added by Lydia 2023/10/20
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/20
   End If
   rsTmp.Close
   
   ' 以發回補理由或發回補答辯計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''   strDay = Empty
   Select Case frm03010404_02.GetSelectResult
      Case "1":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1612")
            textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1612", DBDATE(m_CP05), DBDATE(textCP06), textCP09)
      Case "2":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1613")
            textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1613", DBDATE(m_CP05), DBDATE(textCP06), textCP09)
   End Select
''''   If IsEmptyText(strDay) = False Then
''''      ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''      'textCP48 = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''      textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''      If IsEmptyText(textCP06) = False Then
''''         If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
''''            textCP48 = DBDATE(textCP06)
''''         End If
''''      End If
''''   End If
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
      
   Set rsTmp = Nothing
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim nIndex As Integer
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strCP27 As String
   Dim strNP07 As String
   Dim strNP14 As String
   Dim strNP22 As String
   
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 新增資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質
   ' 案件性質
   strCP10 = "612"
   Select Case frm03010404_02.GetSelectResult
      Case "1": strCP10 = "612"
      Case "2": strCP10 = "613"
   End Select
   
   'Add By Cheng 2002/01/03
   '若來函性質屬於爭議程序(16XX), 應更新商標基本檔是否有爭議程序欄(TM19)為"Y"
   If Left(strCP10, 2) = "16" Then
      strSql = "UPDATE TradeMark SET TM19='Y'" & _
               " WHERE TM01 = '" & m_TM01 & "'" & _
               " And TM02 = '" & m_TM02 & "'" & _
               " And TM03 = '" & m_TM03 & "'" & _
               " And TM04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
   End If
   
   ' 業務區別 91.8.
   'strCP12 = GetST15(m_CP13)
   ' 92.6.14 ADD BY SONIA 發文日
   '2008/12/11 MODIFY BY SONIA 有期限則不上發文日
   'strCP27 = DBDATE(SystemDate())
   If IsEmptyText(textCP06) = False Then
      strCP27 = ""
   Else
      strCP27 = DBDATE(SystemDate())
   End If
   '2008/12/11 END
   ' 新增案件進度資料
   ' 91.03.25 modify by louis (單引號)
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
   '92.6.14 MODIFY BY SONIA 加發文日
   'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP64) " & _
   '         "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
   '                 "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
   '                 "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & ChgSQL(textCP36) & "'," & _
   '                 "'" & ChgSQL(textCP37) & "','" & ChgSQL(textCP38) & "','" & ChgSQL(textCP39) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
   '                 "'" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        'Modify By Cheng 2004/02/04
        '業務區為最近收文A類接洽記錄單智權人員的業務區
'        strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP27,CP32,CP36,CP37,CP40,CP41,CP42,CP43,CP64) " & _
'                 "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                         "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
'                         "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & ChgSQL(textCP36) & "'," & _
'                         "'" & ChgSQL(textCP37_1) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
'                         "'" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
        '2008/12/11 modify by sonia 合併CP14,CP48,CP06,CP07
        'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP27,CP32,CP36,CP37,CP40,CP41,CP42,CP43,CP64) " & _
                 "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                         "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                         "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & ChgSQL(textCP36) & "'," & _
                         "'" & ChgSQL(textCP37_1) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
                         "'" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
        strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP27,CP32,CP36,CP37,CP40,CP41,CP42,CP43,CP64,CP14,CP48,CP06,CP07) " & _
                 "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                         "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                         "'" & "N" & "','" & textCP26 & "'," & CNULL(DBDATE(strCP27)) & ",'" & "N" & "','" & ChgSQL(textCP36) & "'," & _
                         "'" & ChgSQL(textCP37_1) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
                         "'" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & CNULL(textCP14) & "," & CNULL(DBDATE(textCP48)) & "," & CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & ")"
        '2008/12/11 END
        'End
    Case Else
        'Modify By Cheng 2004/02/04
        '業務區為最近收文A類接洽記錄單智權人員的業務區
'        strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP64) " & _
'                 "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                         "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
'                         "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & ChgSQL(textCP36) & "'," & _
'                         "'" & ChgSQL(textCP37) & "','" & ChgSQL(textCP38) & "','" & ChgSQL(textCP39) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
'                         "'" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
        '2008/12/11 modify by sonia 合併CP14,CP48,CP06,CP07
        'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP64) " & _
                 "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                         "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                         "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & ChgSQL(textCP36) & "'," & _
                         "'" & ChgSQL(textCP37) & "','" & ChgSQL(textCP38) & "','" & ChgSQL(textCP39) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
                         "'" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
        strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP64,CP14,CP48,CP06,CP07) " & _
                 "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                         "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                         "'" & "N" & "','" & textCP26 & "'," & CNULL(DBDATE(strCP27)) & ",'" & "N" & "','" & ChgSQL(textCP36) & "'," & _
                         "'" & ChgSQL(textCP37) & "','" & ChgSQL(textCP38) & "','" & ChgSQL(textCP39) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
                         "'" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & CNULL(textCP14) & "," & CNULL(DBDATE(textCP48)) & "," & CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & ")"
        '2008/12/11 END
        'End
    End Select
   '92.6.14 END
   cnnConnection.Execute strSql
   
   'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
   Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   '2008/12/11 CANCEL BY SONIA 移至INSERT INTO CaseProgress
   '' 若有輸入承辦人時
   'If IsEmptyText(textCP14) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP14 = '" & textCP14 & "' " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 若有輸入承辦期限時
   'If IsEmptyText(textCP48) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP48 = " & DBDATE(textCP48) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 有輸入本所期限時
   'If IsEmptyText(textCP06) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP06 = " & DBDATE(textCP06) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 有輸入法定期限時
   'If IsEmptyText(textCP06) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP07 = " & DBDATE(textCP07) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '2008/12/11 END
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 有輸入下一程序時
   If IsEmptyText(textCF15) = False Then
      strNP22 = GetNextProgressNo()
      strNP14 = Empty
      strNP14 = GetRelatedPerson(strCP09)
      'Modify By Cheng 2002/09/25
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
'                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "','" & textCP64 & "'," & strNP22 & ")"
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
      'Modified by Lydia 2024/07/02 +ChgSQL
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "','" & ChgSQL(textCP64) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      Select Case textCF15
         Case "102", "105", "702", "708", "305", "998", "997":
         Case Else:
            'add by nickc 2005/05/31
            IsAppNpData = True
            SeekNewCp09 = strCP09
            'Modify By Cheng 2003/07/09
'            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   End If
   
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
   
   'Add By Sindy 2009/08/17 CFT故將收達及提申期限一併上Y
   If m_TM01 = "CFT" Then
      strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
                     "WHERE NP01 = '" & m_CP09 & "' AND " & _
                           "NP02 = '" & m_TM01 & "' AND " & _
                           "NP03 = '" & m_TM02 & "' AND " & _
                           "NP04 = '" & m_TM03 & "' AND " & _
                           "NP05 = '" & m_TM04 & "' AND " & _
                           "NP07 in (997,998) AND " & _
                           "(NP06 IS NULL OR NP06 <> 'Y') "
      cnnConnection.Execute strSql
   End If
   '2009/08/17 End
   
   'Add by Sindy 2023/5/2
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm03010404_01", strCP09
   End If
   '2023/5/2 END
   
   '911107 nick transation
   cnnConnection.CommitTrans
   Exit Function
   
CheckingErr:
   MsgBox (Err.Description)
   cnnConnection.RollbackTrans
   OnSaveData = False
End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   ' 下一程序不可空白
   If IsEmptyText(textCF15) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入下一程序"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCF15.SetFocus
      GoTo EXITSUB
   End If
   ' 本所期限不可空白
   If IsEmptyText(textCP06) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入本所期限"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP06.SetFocus
      GoTo EXITSUB
   End If
   'Add By Cheng 2002/03/11
   If Me.textCP06.Text <> "" Then
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Me.textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      End If
   End If
   ' 法定期限不可空白
   If IsEmptyText(textCP07) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入法定期限"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP07.SetFocus
      GoTo EXITSUB
   End If
   ' 本所期限不可超過法定期限
   If Val(textCP06) > Val(textCP07) Then
      strTit = "資料檢核"
      strMsg = "本所期限的日期不可超過法定期限的日期"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP06.SetFocus
      GoTo EXITSUB
   End If
   ' 承辦期限不可空白
   If IsEmptyText(textCP48) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入承辦期限"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP48.SetFocus
      GoTo EXITSUB
   End If
   'Add By Cheng 2002/05/07
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         textCP48.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
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

Private Sub Form_Unload(Cancel As Integer)
    'Add By Cheng 2002/07/19
   Set frm03010404_03 = Nothing
End Sub



Private Sub grdList_Click()
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
            grdList.Text = "V"
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

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCF15) = False Then
      Select Case textCF15.Text
         Case "補充理由":
            textCF15 = "612"
         Case "補充答辯":
            textCF15 = "613"
      End Select
   
      ' 只取得國內的案件性質名稱
      If m_TM10 < "010" Then
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
      Else
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 1)
      End If
      If IsEmptyText(textCF15_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      ' 檢查是否為西元日期
      If CheckIsDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = PUB_GetWorkDay1(textCP06, True)
      'end 2020/07/09
      End If
      'Add By Cheng 2002/03/11
      'Modify By Sindy 2009/09/18
      'If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
      If Val(Me.textCP06.Text) < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Cancel = True
         textCP06_GotFocus
         GoTo EXITSUB
      End If
      
      ' 以發回補理由或發回補答辯計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''      strDay = Empty
      Select Case frm03010404_02.GetSelectResult
         Case "1":
''''            strDay = GetWorkDays(m_TM01, m_TM10, "1612")
                textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1612", DBDATE(m_CP05), DBDATE(textCP06), textCP09)
         Case "2":
''''            strDay = GetWorkDays(m_TM01, m_TM10, "1613")
                textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1613", DBDATE(m_CP05), DBDATE(textCP06), textCP09)
      End Select
''''      If IsEmptyText(strDay) = False Then
''''         ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''         'textCP48 = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''         textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''         If IsEmptyText(textCP06) = False Then
''''            If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
''''               textCP48 = DBDATE(textCP06)
''''            End If
''''         End If
''''      End If
   End If
EXITSUB:
End Sub

' 法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      ' 檢查是否為民國年
      If CheckIsDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2010/11/29
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14)
      If IsEmptyText(textCP14_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "承辦人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP14_GotFocus
      End If
   End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否算案件數
Private Sub textCP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP26) = False Then
      Select Case textCP26
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
   End If
End Sub

Private Sub textCP37_1_GotFocus()
    TextInverse Me.textCP37_1
End Sub

' 承辦人期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為西元日期
      If CheckIsDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的承辦期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
         Exit Sub
      End If
   End If
   'Add By Cheng 2002/05/07
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         Cancel = True
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         textCP48_GotFocus
         Exit Sub
      End If
   End If
   
End Sub

Private Sub textCF15_GotFocus()
   textCF15.SelStart = 0
   textCF15.SelLength = Len(textCF15.Text)
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP08_GotFocus()
   InverseTextBox textCP08
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
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

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCF15.Enabled = True Then
   Cancel = False
   textCF15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP06.Enabled = True Then
   Cancel = False
   textCP06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP07.Enabled = True Then
   Cancel = False
   textCP07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP14.Enabled = True Then
   Cancel = False
   textCP14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP26.Enabled = True Then
   Cancel = False
   textCP26_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP48.Enabled = True Then
   Cancel = False
   textCP48_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Lydia 2021/12/08
If Me.textPrint.Enabled = True Then
   Cancel = False
   textPrint_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If

TxtValidate = True
End Function

'Added by Lydia 2021/12/08
Private Sub textPrint_GotFocus()
   TextInverse textPrint
End Sub

'Added by Lydia 2021/12/08
' 是否列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case "", " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

'Added by Lydia 2021/12/08
' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   If m_TM01 = "CFT" Then
      ' 清除定稿例外欄位檔原有資料(定稿別 /收文號或本所案號000 /處理狀況 /使用者編號)
      EndLetter "04", m_CP09, "00", strUserNum
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Added by Lydia 2021/12/08 CFT 沒設定都出通用定稿(不列印)
   If m_TM01 = "CFT" Then
      NowPrint m_CP09, "04", "00", False, strUserNum, 0, , , , , , , , , , , , , , , , , True
   End If
End Sub

