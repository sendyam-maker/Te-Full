VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010403_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "審查報告輸入"
   ClientHeight    =   6120
   ClientLeft      =   156
   ClientTop       =   972
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9324
   Begin VB.CommandButton Command2 
      Caption         =   "部份核駁商品資料異動(&I)"
      Height          =   375
      Left            =   4110
      TabIndex        =   25
      Top             =   0
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4212
      Left            =   60
      TabIndex        =   46
      Top             =   1860
      Width           =   9228
      _ExtentX        =   16277
      _ExtentY        =   7430
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "一般"
      TabPicture(0)   =   "frm02010403_4.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label23"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label22"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label21"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label24"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label25"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label26"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label14"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label32"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textCP64"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCP14_2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtTM67"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "grdList"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textCP26"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textPrint"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCP06"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCP14"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCP07"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCP48"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCP49"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Frame2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Frame1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCP08"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCF15"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCF15_2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "對造名稱"
      TabPicture(1)   =   "frm02010403_4.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label35"
      Tab(1).Control(1)=   "Label34"
      Tab(1).Control(2)=   "Label31"
      Tab(1).Control(3)=   "Label29"
      Tab(1).Control(4)=   "Label28"
      Tab(1).Control(5)=   "Label30"
      Tab(1).Control(6)=   "Label27"
      Tab(1).Control(7)=   "textCP40"
      Tab(1).Control(8)=   "textCP42"
      Tab(1).Control(9)=   "textCP37_1"
      Tab(1).Control(10)=   "textCP36"
      Tab(1).Control(11)=   "textCP41"
      Tab(1).Control(12)=   "textCP80"
      Tab(1).ControlCount=   13
      Begin VB.TextBox textCF15_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   6510
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   330
         Width           =   1692
      End
      Begin VB.TextBox textCF15 
         Height          =   264
         Left            =   5670
         MaxLength       =   4
         TabIndex        =   1
         Top             =   330
         Width           =   732
      End
      Begin VB.TextBox textCP08 
         Height          =   264
         Left            =   1170
         MaxLength       =   40
         TabIndex        =   0
         Top             =   330
         Width           =   2532
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   1170
         TabIndex        =   70
         Top             =   540
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "文到次日"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   3
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "文到當日"
            Height          =   180
            Index           =   0
            Left            =   144
            TabIndex        =   2
            Top             =   180
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   4020
         TabIndex        =   69
         Top             =   540
         Width           =   4215
         Begin VB.TextBox Text12 
            Height          =   252
            Left            =   2760
            MaxLength       =   7
            TabIndex        =   9
            Top             =   150
            Width           =   975
         End
         Begin VB.TextBox Text11 
            Height          =   270
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   7
            Top             =   150
            Width           =   375
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Left            =   840
            MaxLength       =   2
            TabIndex        =   5
            Top             =   150
            Width           =   375
         End
         Begin VB.OptionButton Option4 
            Caption         =   "                      日"
            Height          =   225
            Index           =   2
            Left            =   2520
            TabIndex        =   8
            Top             =   180
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "        月"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   6
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "文到          天"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   180
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox textCP80 
         Height          =   264
         Left            =   -73530
         MaxLength       =   39
         TabIndex        =   24
         Top             =   2430
         Width           =   3495
      End
      Begin VB.TextBox textCP41 
         Height          =   264
         Left            =   -73530
         TabIndex        =   22
         Top             =   1860
         Width           =   7092
      End
      Begin VB.TextBox textCP36 
         Height          =   264
         Left            =   -73530
         MaxLength       =   200
         TabIndex        =   19
         Top             =   480
         Width           =   7092
      End
      Begin VB.TextBox textCP49 
         Height          =   264
         Left            =   1170
         MaxLength       =   300
         TabIndex        =   14
         Top             =   2784
         Width           =   7992
      End
      Begin VB.TextBox textCP48 
         Height          =   264
         Left            =   5670
         MaxLength       =   7
         TabIndex        =   13
         Top             =   1350
         Width           =   2532
      End
      Begin VB.TextBox textCP07 
         Height          =   264
         Left            =   5670
         MaxLength       =   7
         TabIndex        =   11
         Top             =   1080
         Width           =   2532
      End
      Begin VB.TextBox textCP14 
         Height          =   264
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   12
         Top             =   1350
         Width           =   732
      End
      Begin VB.TextBox textCP06 
         Height          =   264
         Left            =   1170
         MaxLength       =   7
         TabIndex        =   10
         Top             =   1050
         Width           =   2532
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   1170
         MaxLength       =   1
         TabIndex        =   15
         Top             =   3096
         Width           =   372
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   6024
         MaxLength       =   1
         TabIndex        =   16
         Top             =   3096
         Width           =   372
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   1116
         Left            =   1176
         TabIndex        =   72
         Top             =   1632
         Width           =   7900
         _ExtentX        =   13928
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
      Begin MSForms.TextBox txtTM67 
         Height          =   300
         Left            =   1170
         TabIndex        =   17
         Top             =   3396
         Width           =   7992
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "14097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37_1 
         Height          =   792
         Left            =   -73530
         TabIndex        =   20
         Top             =   750
         Width           =   7092
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "12509;1397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73530
         TabIndex        =   23
         Top             =   2130
         Width           =   7095
         VariousPropertyBits=   679493659
         Size            =   "12515;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73530
         TabIndex        =   21
         Top             =   1560
         Width           =   7092
         VariousPropertyBits=   679493659
         Size            =   "12509;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   264
         Left            =   2010
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1692
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         MaxLength       =   20
         Size            =   "2984;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   456
         Left            =   1170
         TabIndex        =   18
         Top             =   3708
         Width           =   7992
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "14097;804"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label32 
         Caption         =   "來函期限:"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "對造商品類別 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   68
         Top             =   2430
         Width           =   1575
      End
      Begin VB.Label Label30 
         Caption         =   "對造案件名稱 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   67
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label Label28 
         Caption         =   "對造日文名稱 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   66
         Top             =   2130
         Width           =   1575
      End
      Begin VB.Label Label29 
         Caption         =   "對造英文名稱 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   65
         Top             =   1890
         Width           =   1575
      End
      Begin VB.Label Label31 
         Caption         =   "對造中文名稱 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   64
         Top             =   1605
         Width           =   1575
      End
      Begin VB.Label Label34 
         Caption         =   "對造案件中文名稱 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   63
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label Label35 
         Caption         =   "對造號數 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   62
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "條款 :"
         Height          =   252
         Left            =   120
         TabIndex        =   61
         Top             =   2784
         Width           =   852
      End
      Begin VB.Label Label26 
         Caption         =   "承辦期限 :"
         Height          =   255
         Left            =   4710
         TabIndex        =   60
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   255
         Left            =   4710
         TabIndex        =   59
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "承辦人 :"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "本所期限 :"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "下一程序 :"
         Height          =   255
         Left            =   4710
         TabIndex        =   56
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "機關文號 :"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "進度備註 :"
         Height          =   252
         Left            =   120
         TabIndex        =   54
         Top             =   3732
         Width           =   972
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   252
         Left            =   120
         TabIndex        =   53
         Top             =   3096
         Width           =   972
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印)"
         Height          =   180
         Left            =   1644
         TabIndex        =   52
         Top             =   3096
         Width           =   648
      End
      Begin VB.Label Label16 
         Caption         =   "是否算案件數 :"
         Height          =   252
         Left            =   4710
         TabIndex        =   51
         Top             =   3096
         Width           =   1212
      End
      Begin VB.Label Label15 
         Caption         =   "(N:不算)"
         Height          =   252
         Left            =   6540
         TabIndex        =   50
         Top             =   3096
         Width           =   972
      End
      Begin VB.Label Label7 
         Caption         =   "本案期限 :"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "放棄專用權 :"
         Height          =   252
         Left            =   120
         TabIndex        =   48
         Top             =   3396
         Width           =   1128
      End
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   390
      Width           =   2532
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   375
      Left            =   7245
      TabIndex        =   27
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   6420
      TabIndex        =   26
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8475
      TabIndex        =   28
      Top             =   0
      Width           =   800
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1290
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1590
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   990
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   390
      Width           =   2532
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1170
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   660
      Width           =   7992
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14097;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5700
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1290
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1200
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   990
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   255
      Index           =   1
      Left            =   4740
      TabIndex        =   44
      Top             =   390
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "審定號 :"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4740
      TabIndex        =   42
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   41
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   255
      Index           =   3
      Left            =   4740
      TabIndex        =   40
      Top             =   990
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   390
      Width           =   975
   End
End
Attribute VB_Name = "frm02010403_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/28 Form2.0已修改 cmbTM05/textTM23/textCP13/textCP14_2/textCP44/textCP37_1/textCP40/textCP42/grdList/txtTM67(111/8/8 Lydia)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 來函收文日
Dim m_CP05 As String
' 收文號
Dim m_CP09 As String
' 原案件性質
Dim m_CP10 As String
' 原智權人員代號
Dim m_CP13 As String
Dim m_CP12 As String
' 原移轉申請人代號
Dim m_CP56 As String
' 商標種類代碼
Dim m_TM08 As String
' 國家代碼
Dim m_TM10 As String
' 原專用期限起日
Dim m_TM21 As String
' 原專用期限止日
Dim m_TM22 As String
' 原申請人代號
Dim m_TM23 As String
' 申請國家的延展年度
Dim m_NA14 As Integer

Dim m_CurrSel As Integer

'add by Sindy 2009/06/02 檢查是否已經有商品及服務
Public ChkTG As Boolean
Dim strRvType As String 'Add By Sindy 2012/4/26

'Added by Morgan 2017/4/17 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/4/17
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/10 END
Dim strFromDate As String 'Added by Lydia 2019/06/21 期限起算日
Dim strLD18 As String 'Add By Sindy 2020/1/7 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2020/1/7 FC代理人


'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm02010403_3.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010403_3
   Unload frm02010403_2
   Unload frm02010403_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
        'Modify By Cheng 2002/11/07
'      'OnSaveData
        If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      'Unload Me
      Unload frm02010403_3
      Unload frm02010403_2
      'Add By Sindy 2019/5/10
      If Me.m_strIR01 <> "" Then
        Unload frm02010403_1
        If Not m_PrevForm Is Nothing Then
           Call m_PrevForm.GoNext
        End If
        Unload Me
      '2019/5/10 END
      'Modified by Morgan 2017/4/17 電子公文
      'frm02010403_1.Show
      ElseIf m_DocNo <> "" Then
         Unload Me
         Unload frm02010403_1
         frm02010412.GoNext
      Else
         frm02010403_1.Show
         Unload Me
      End If
      'end 2017/4/17
   End If
End Sub

'Add By Sindy 2009/06/02
Private Sub Command2_Click()
frm02010403_5.Hide
Set frm02010403_5.UpForm = Me
frm02010403_5.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
frm02010403_5.AllClass = textTM09.Text
frm02010403_5.textCP64 = textCP64.Text
frm02010403_5.Tag = m_TM10 & m_CP09  'Added by Lydia 2024/11/21
Me.Hide
frm02010403_5.Show 'vbModal
frm02010403_5.QueryData
'If frm02010403_5.BolOk = True Then
'   textCP64 = frm02010403_5.strCP64
'End If
'Unload frm02010403_5
'Set frm02010403_5 = Nothing
End Sub

Private Sub Form_Activate()
   If SSTab1.Tab = 0 Then
      If textCP08.Visible = True Then 'Modify By Sindy 2009/06/03
         textCP08.SetFocus
      End If
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   'textTM08.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   'textTM45.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
'   textCP05S.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCF15_2.BackColor = &H8000000F
  
   MoveFormToCenter Me
   
   'Add By Sindy 2009/06/02
   '3.部分核駁
   If frm02010403_3.GetSelectResult = "3" Then
      Command2.Visible = True
   Else
      Command2.Visible = False
      '2.核駁前先行通知
      'Modify by Amy 2022/09/26 +1.審查報告
      If frm02010403_3.GetSelectResult = "2" Or frm02010403_3.GetSelectResult = "1" Then
            SSTab1.TabCaption(1) = "關係案"
            strExc(1) = "對方"
            Label35.Caption = strExc(1) & Mid(Label35.Caption, 3)
            Label30.Caption = strExc(1) & Mid(Label30.Caption, 3)
            Label31.Caption = strExc(1) & Mid(Label31.Caption, 3)
            Label29.Caption = strExc(1) & Mid(Label29.Caption, 3)
            Label28.Caption = strExc(1) & Mid(Label28.Caption, 3)
            Label27.Caption = strExc(1) & Mid(Label27.Caption, 3)
      End If
   End If
   
   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010403_1.m_strIR01
   m_strIR02 = frm02010403_1.m_strIR02
   m_strIR03 = frm02010403_1.m_strIR03
   m_strIR04 = frm02010403_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/10 END
   
   strFromDate = DBDATE(frm02010403_1.textCP05)  'Added by Lydia 2019/06/21
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

Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
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
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         Label2.Caption = "審定號 :" 'Add By Sindy 2011/2/25
         textTM15 = rsTmp.Fields("TM15")
      Else
         'Add By Sindy 2011/2/25
         If IsNull(rsTmp.Fields("TM12")) = False Then
            Label2.Caption = "申請案號 :"
            textTM15 = rsTmp.Fields("TM12")
         End If
         '2011/2/25 End
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
      ' 商標種類
'      If IsNull(rsTmp.Fields("TM08")) = False Then
'         m_TM08 = rsTmp.Fields("TM08")
'         If m_TM10 < "010" Then
'            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
'         Else
'            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
'         End If
'      End If
        'Add By Cheng 2004/02/10
        '商品類別
        Me.textTM09.Text = "" & rsTmp.Fields("TM09").Value
        'End
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      
      'Add By Sindy 2020/1/7
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("TM44")) = False Then
         m_TM44 = rsTmp.Fields("TM44")
      End If
      '2020/1/7 END
      
      ' 彼所案號
'      If IsNull(rsTmp.Fields("TM45")) = False Then
'         textTM45 = rsTmp.Fields("TM45")
'      End If
       'Add By Cheng 2003/03/27
       '放棄專用權
       Me.txtTM67.Text = "" & rsTmp("TM67").Value
       
      'Added by Lydia 2019/06/21 台-大核駁案期限管制:取消來函期限
      If m_TM10 <> "000" And frm02010403_3.GetSelectResult = "3" Then
          Label32.Caption = "來函類別:"
          Option1(0).Caption = "紙本公文"
          Option1(1).Caption = "電子公文"
          Frame2.Visible = False
      End If
      
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 取得服務業務基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 審定號
      Label2.Caption = "審定號 :" 'Add By Sindy 2011/2/25
      textTM15 = Empty
      ' 案件名稱(中)
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP05")
      End If
      ' 案件名稱(英)
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP06")
      End If
      ' 案件名稱(日)
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP07")
      End If
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 商標種類
'      m_TM08 = Empty
'      textTM08 = Empty
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      
      'Add By Sindy 2020/1/7
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("SP26")) = False Then
         m_TM44 = rsTmp.Fields("SP26")
      End If
      '2020/1/7 END
      
      ' 彼所案號
'      If IsNull(rsTmp.Fields("SP27")) = False Then
'         textTM45 = rsTmp.Fields("SP27")
'      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then
         'Modify By Sindy 2012/5/31 Mark
         'textCP08 = rsTmp.Fields("CP08")
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
      ' 智權人員
      'Add By Cheng 2002/07/17
      m_CP13 = Empty
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      '業務區   nick 91.08.22
      m_CP12 = Empty
      If IsNull(rsTmp.Fields("cp12")) = False Then
         m_CP12 = rsTmp.Fields("cp12")
      End If
      ' 移轉申請人代號
      If IsNull(rsTmp.Fields("CP56")) = False Then
         m_CP56 = rsTmp.Fields("CP56")
      End If
      ' 下一程序
      'MODIFY BY SONIA
      'textCF15 = GetNextProgress(m_TM01, m_TM10, m_CP10)
      Select Case frm02010403_3.GetSelectResult
         Case "1": textCF15 = GetNextProgress(m_TM01, m_TM10, "1201")
         Case "2": textCF15 = GetNextProgress(m_TM01, m_TM10, "1202")
         '92.8.1 ADD BY SONIA
         Case "3": textCF15 = GetNextProgress(m_TM01, m_TM10, "1205")
      End Select
      '92.8.1 add by sonia
      If textCF15 = "401" And m_TM10 = "020" Then
         textCP64 = "部分核駁商品："
      End If
      '92.8.1 end
      
      If IsEmptyText(textCF15) = False Then
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
      End If
      ' 本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then
         If IsEmptyText(rsTmp.Fields("CP06")) = False Then
            textCP06 = TAIWANDATE(rsTmp.Fields("CP06"))
         End If
      End If
      ' 法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then
         If IsEmptyText(rsTmp.Fields("CP07")) = False Then
            textCP07 = TAIWANDATE(rsTmp.Fields("CP07"))
         End If
      End If
      ' 承辦人
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = rsTmp.Fields("CP14")
         textCP14_2 = GetStaffName(rsTmp.Fields("CP14"))
      End If
   End If
   rsTmp.Close
   
   ' 90.11.19 modify by sonia
   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   If m_TM10 < "010" Then
      If textCP08 = "" Then
         textCP08 = "（" & strTmp & "）慧商字第號"
      End If
   End If
   
End Sub

Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strDay As String
   
   m_CP10 = Empty
   m_CP56 = Empty
   m_TM08 = Empty
   m_TM10 = Empty
   m_TM21 = Empty
   m_TM22 = Empty
   m_TM23 = Empty
   
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   
   ' 來函收文日
'   textCP05S = m_CP05
   
   ' 讀取基本檔的資料
   Select Case m_TM01
      Case "T", "TF", "FCT", "CFT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   ' 讀取案件進度檔的資料
   QueryCaseProgress
   Call ChgType 'Add By Sindy 2012/4/17 讀取來函期限
   
''''edit by nickc 2007/10/12 改抓有時效的
''''   strDay = Empty
   Select Case frm02010403_3.GetSelectResult
      Case "1":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1201")
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1201", DBDATE(m_CP05), DBDATE(textCP06)))
      Case "2":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1202")
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1202", DBDATE(m_CP05), DBDATE(textCP06)))
   End Select
''''   If IsEmptyText(strDay) = False Then
''''    'Modify By Cheng 2003/09/01
'''''      textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''      textCP48 = TAIWANDATE(DateAdd("d", Val(strDay), ChangeWStringToWDateString(DBDATE(m_CP05))))
''''   End If
   
   'Added by Morgan 2017/4/17 電子公文
   If m_DocWord <> "" Then
      textCP08 = m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號"
   ElseIf m_DocNo <> "" Then
      textCP08 = Replace(textCP08, "第號", "第" & PUB_GetEDocNo(m_DocNo) & "號")
   End If
   '期限
   If m_DeadLine <> "" Then
      Option1(1).Value = True
      If Len(m_DeadLine) >= 7 Then
         Option4(2).Value = True
         Text12 = m_DeadLine
         Text12_Validate False
      ElseIf Right(m_DeadLine, 1) = "日" Then
         Option4(0).Value = True
         Text10 = Val(m_DeadLine)
         Text10_Validate False
      ElseIf Right(m_DeadLine, 1) = "月" Then
         Option4(1).Value = True
         Text11 = Val(m_DeadLine)
         Text11_Validate False
      End If
   End If
   'end 2017/4/17
   
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
            '92.7.4 MODIFY BY SONIA
            'grdList.TextMatrix(grdList.Row, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"))
            If m_TM10 < "010" Then
               grdList.TextMatrix(grdList.row, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"), 0)
            Else
               grdList.TextMatrix(grdList.row, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"), 1)
            End If
            '92.7.4 END
            grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("NP07")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("NP08")) = False Then
            If IsEmptyText(rsTmp.Fields("NP08")) = False Then
               grdList.TextMatrix(grdList.row, 2) = TAIWANDATE(rsTmp.Fields("NP08"))
            End If
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               grdList.TextMatrix(grdList.row, 3) = TAIWANDATE(rsTmp.Fields("NP09"))
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
      'Added by Lydia 2023/10/18
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/18
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
   '90.08.16 modify by sonia 90.12.26全都預設N
   'If frm02010403_3.textResult = "1" Then
      textPrint = "N"
   'End If
End Sub

'Modify By Cheng 2002/11/07
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim bUpdate As Boolean
Dim strSubTMSQL As String
Dim strSubCPSQL As String
Dim strCP09 As String
Dim strCP12 As String
Dim strNP07 As String
Dim strNP08 As String
Dim strNP09 As String
Dim strNP14 As String
Dim strNP22 As String
Dim nIndex As Integer
'Add By Cheng 2003/03/27
Dim StrSQLa As String
'Add by Amy 2017/11/13
Dim m_CP06 As String, m_CP07 As String, st_CP09 As String, m_CP14 As String, strMsg As String
Dim bolUpdCP As Boolean '是否更新進度檔
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount <= 0 Then
      GoTo EXITSUB
   End If
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   rsTmp.MoveFirst
   
   ' 設定SQL中Update TradeMark的語法
   strSubTMSQL = "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
    'Add By Cheng 2003/03/27
    '更新商標基標檔
    StrSQLa = "Update Trademark Set TM67='" & ChgSQL(Me.txtTM67.Text) & "' " & strSubTMSQL
    cnnConnection.Execute StrSQLa
   ' 設定SQL中CaseProgress的語法
   strSubCPSQL = "WHERE CP09 = '" & m_CP09 & "' "
   '  新增資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質
   strRvType = "1002"
   Select Case frm02010403_3.GetSelectResult
      Case "1": strRvType = "1201"
      Case "2": strRvType = "1202"
      '92.8.1 ADD
      Case "3": strRvType = "1205"
   End Select
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 91.03.25 modify by louis (單引號)
    '承辦人為原程序承辦人, 不上發文日
    'Modify By Cheng 2003/04/04
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng  2004/02/03
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP49,CP64) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
'                    "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "') "
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP49,CP64) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
'                    "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "') "
'業務區為最近收文A類接洽記錄單智權人員的業務區
'2008/10/02 增加對造名稱 CP37,CP40,CP41,CP42 ADD BY TONI
'Modify By Sindy 2010/7/6 增加對造商品類別 CP80
strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP36,CP37,CP40,CP41,CP42,CP80,CP43,CP49,CP64) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
                    "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & ChgSQL(textCP36) & "','" & ChgSQL(textCP37_1) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & ChgSQL(textCP42) & "','" & ChgSQL(textCP80) & "','" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "') "
'End
   cnnConnection.Execute strSql
    
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   ' 有輸入本所期限
   If IsEmptyText(textCP06) = False Then
      strSql = "UPDATE CaseProgress SET CP06 = " & DBDATE(textCP06) & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   ' 有輸入法定期限
   If IsEmptyText(textCP07) = False Then
      strSql = "UPDATE CaseProgress SET CP07 = " & DBDATE(textCP07) & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   ' 有輸入承辦期限時
   If IsEmptyText(textCP48) = False Then
      strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(textCP48) & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   'add by nickc 2008/01/10 FCT 加判斷，有期限用期限判斷(第三或第五個工作天)，無期限以第三個工作日(當日不算)，寫入承辦期限
   ElseIf m_TM01 = "FCT" Then
        If Trim(textCP07) = "" Then
            strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(4, DBDATE(m_CP05), 0)) & " " & _
                     "WHERE CP09 = '" & strCP09 & "' "
            cnnConnection.Execute strSql
        Else
            If DateDiff("d", ChangeWStringToWDateString(DBDATE(m_CP05)), ChangeWStringToWDateString(DBDATE(textCP07))) <= 30 Then    '無法與上句合併，因為沒有日期時，datediff  會發生  型態不符 的錯誤
                strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(4, DBDATE(m_CP05), 0)) & " " & _
                         "WHERE CP09 = '" & strCP09 & "' "
                cnnConnection.Execute strSql
            Else
                strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(6, DBDATE(m_CP05), 0)) & " " & _
                         "WHERE CP09 = '" & strCP09 & "' "
                cnnConnection.Execute strSql
            End If
        End If
    End If
    
    'Add By Sindy 2012/4/26 儲存官方發文日及官方期限月數
    If Trim(Text11) <> "" Then
      strSql = "UPDATE CaseProgress SET CP133=" & DBDATE(m_CP05) & ",CP134=" & Text11 & " " & _
               "WHERE CP09='" & strCP09 & "' "
      cnnConnection.Execute strSql
    End If
    
   ' 有輸入下一程序時, 新增資料到下一程序檔
   If IsEmptyText(textCF15) = False Then
    'Modify by Amy 2017/11/13 +if 判斷進度檔已有相同未發文未取消收文之案件性質,則判斷是否更新本限及法限
    If ChkSameCaseProgress(m_TM01, m_TM02, m_TM03, m_TM04, textCF15, m_CP06, m_CP07, st_CP09, m_CP14) = True Then
      If m_CP06 = MsgText(601) Or m_CP07 = MsgText(601) Then
        If MsgBox("下一程序已收文但無期限，是否要代入新期限？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
            bolUpdCP = True
        End If
      ElseIf Val(textCP06) + 19110000 <> Val(m_CP06) Or Val(textCP07) + 19110000 <> Val(m_CP07) Then
        strMsg = "下一程序已收文且期限不同" & vbCrLf & _
                 "已收文本所期限：" & IIf(m_CP06 <> "", Val(m_CP06) - 19110000, "") & " 來函本所期限：" & textCP06 & vbCrLf & _
                 "已收文法定期限：" & IIf(m_CP07 <> "", Val(m_CP07) - 19110000, "") & " 來函法定期限：" & textCP07 & vbCrLf
        
        If MsgBox(strMsg & "是否要更新為來函期限？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
            bolUpdCP = True
        End If
      End If
    End If
      
    '更新進度檔,並發Mail通知承辦人
    If bolUpdCP = True Then
        strSql = "Update CaseProgress Set CP06=" & Val(textCP06) + 19110000 & ",CP07=" & Val(textCP07) + 19110000 & " Where CP09='" & st_CP09 & "'"
        cnnConnection.Execute strSql
        
        If m_CP14 = MsgText(601) Then m_CP14 = GetDeptMan("P20") '無承辦人發給P20部門之A0908
        strMsg = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "收到" & "" & GetCaseTypeName(m_TM01, textCF15, IIf(m_TM10 = "000", 0, 1)) & "前已收文,請辦理後續！"
        PUB_SendMail strUserNum, m_CP14, "", strMsg, "本所期限：" & textCP06 & "　　法定期限：" & textCP07
      
    '進度檔未有相同未發文未取消收文之案件性質或上述不更新期限,才新增下一程序
    Else
        strNP08 = Empty
        If IsEmptyText(textCP06) = False Then: strNP08 = DBDATE(textCP06)
        strNP09 = Empty
        If IsEmptyText(textCP07) = False Then: strNP09 = DBDATE(textCP07)
        strNP14 = Empty
        '2008/12/4 modify by sonia 加放對造名稱於下一程序備註T-158424
        'strNP14 = GetRelatedPerson(m_CP09)
        strNP14 = GetRelatedPerson(strCP09)
        '2008/12/4 end
        ' 序號
        strNP22 = GetNextProgressNo()
        'Modify By Cheng 2002/09/25
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
'                          strNP08 & "," & strNP09 & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "','" & textCP64 & "'," & strNP22 & ")"
        'Modify By Cheng 2003/04/04
        '智權人員存最近收文A類接洽記錄單的智權人員
        'Modified by Lydia 2024/07/02 +ChgSQL
        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                            strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "','" & ChgSQL(textCP64) & "'," & strNP22 & ")"
        cnnConnection.Execute strSql
    End If
    'end 2017/11/13
    
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
'      '92.6.8 SONIA 加 言詞辯論, 準備程序
      Select Case textCF15
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
         '2008/5/7 modify by sonia 補正201也不印
         '2010/4/20 MODIFY BY SONIA 台->大的201要印回覆單
         'Case "102", "105", "702", "708", "305", "998", "997", "201"
         Case "102", "105", "702", "708", "305", "998", "997"
         Case "201"
            If m_TM10 = "020" Then
                'Modify by Amy 2017/11/16 未更新進度檔才印回覆單
                If bolUpdCP = False Then
                    Call g_PrtForm001.PrintReturnSheet(strCP09, textCF15, DBDATE(strNP09), False, , , , m_TM01 & m_TM02 & m_TM03 & m_TM04)
                End If
            End If
         '2010/4/20 END
         Case Else:
            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
            'add by nickc 2008/04/23  加入案件回覆單
            '2008/5/20 MODIFY BY SONIA FCT案不印回覆單
            'modify by sonia 2018/9/28 再加控制大->台案件不印案件回覆單Left(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), 3) <> "MCT"
            If m_TM01 <> "FCT" And Left(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), 3) <> "MCT" Then
                'Modify by Amy 2017/11/16 未更新進度檔才印回覆單
                If bolUpdCP = False Then
                    Call g_PrtForm001.PrintReturnSheet(strCP09, textCF15, DBDATE(strNP09), False, , , , m_TM01 & m_TM02 & m_TM03 & m_TM04)
                End If
            End If
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
   
   'Add By Sindy 2019/12/19 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = strCP09
      If Val(textCP06) > 0 Then '有期限者,為掛號
         PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", True, m_TM23, strRvType, m_TM44
      Else
         PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", False, m_TM23, strRvType, m_TM44
      End If
   End If
   '2019/12/19 END
   
   'Add By Sindy 2009/09/24
   '因為有些來函由內商輸入，內商有自行控管之承辦期限及發文日。改為內商輸入所有C類來函，
   '若業務區為F字頭者，除爭議受理外，自動產生B類收文，案件性質為外商發文722，不上發文日，不向客戶請款
   Dim strCP48 As String, strCP09B As String
   If Left(GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)), 1) = "F" And _
      ((m_TM01 = "T" And m_TM10 = "020") Or (m_TM01 = "FCT" And m_TM10 = "000")) Then
      strCP09B = AutoNo("B", 6)
      '承辦期限為系統日加4個工作天
      strCP48 = DBDATE(Pub_GetHandleDay(m_TM01, m_TM10, "722", strSrvDate(1), , m_CP09))
      '2011/4/28 modify by sonia 智權人員原抓點選收文號之智權人員,改抓該案最後收文在職智權人員
      strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp48,cp20,cp26,cp32,cp43) " & _
                     "values (" & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & _
                     "," & CNULL(m_TM04) & "," & CNULL(strSrvDate(1)) & "," & CNULL(strCP09B) & ",722," & _
                     CNULL(GetSalesArea(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "," & CNULL(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "," & CNULL(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "," & CNULL(strCP48) & ",'N','N','N'," & CNULL(strCP09) & ")"
      cnnConnection.Execute strSql
   End If
   '2009/09/24 End
   
   '2010/3/25 add by sonia 部分核駁1205來函,更新點選收文號的催審期限為來函收文日+6個月T-156759
   If strRvType = "1205" Then
      strNP08 = CompDate(1, 6, ChangeTStringToWString(frm02010403_1.textCP05))
      'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天 +PUB_GetWorkDay1()
       strSql = "UPDATE NextProgress SET NP08 = " & PUB_GetWorkDay1(strNP08, True) & ",NP09=" & strNP08 & " " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP01 = '" & m_CP09 & "' AND NP07=305 AND NP06 IS NULL "
      cnnConnection.Execute strSql
   End If
   '2010/3/25 end
   'Added by Lydia 2016/03/15 台灣案T案為'2'核駁前先行通知(1202) 時,存檔時將該案號申請101那一道的下一程序催審305期限(NP06 IS NULL)更新 NP08=NP09=畫面上之法定期限+8個月.
   'Modifid by Lydia 2023/12/01 修改：A. 加入'1'審查報告(1201)也要更新；B. 不更新申請101的催審期限，改為更新點選收文號的催審期限；C. 改為來函收文日+原點選案件性質之催審天數
   'If m_TM01 = "T" And m_TM10 = "000" And strRvType = "1202" Then
   '   strNP08 = CompDate(1, 8, ChangeTStringToWString(textCP07))
   '   strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
   '   strSql = "UPDATE NEXTPROGRESS SET NP08=" & CNULL(strNP08, True) & ",NP09=" & CNULL(strNP08, True) & _
               " WHERE (NP01,NP22) IN (SELECT NP01,NP22 FROM NEXTPROGRESS,CASEPROGRESS C1" & _
               " WHERE NP02='" & m_TM01 & "' AND NP03='" & m_TM02 & "' AND NP04='" & m_TM03 & "' AND NP05='" & m_TM04 & "'" & _
               " AND NP07='305' AND NP06 IS NULL AND NP01=C1.CP09(+) AND C1.CP10='101')"
   If m_TM01 = "T" And m_TM10 = "000" And (strRvType = "1202" Or strRvType = "1201") Then
      strNP08 = GetUrgeDate(m_TM01, m_TM10, m_CP10, strFromDate)
      
      If Val(strNP08) > 0 Then 'Added by Morgan 2024/12/18 增加判斷有設催審天數才要更新 Ex:T-246593 613補充答辯不必催審 --桂英
         
         strSql = "UPDATE NEXTPROGRESS SET NP08=" & PUB_GetWorkDay1(strNP08, True) & ",NP09=" & strNP08 & _
               " WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND " & _
                      "NP01 = '" & m_CP09 & "' AND NP07=305 AND NP06 IS NULL "
   'end 2023/12/01
         cnnConnection.Execute strSql, intI
         
      End If 'Added by Morgan 2024/12/18
   End If
   'end 2016/03/15
   
   'Added by Lydia 2024/11/21 內商大陸之部份核駁商品異動，記錄各類的比對結果; 先用相關收文號做為PK,等到來函收文再轉為C類收文號
   If m_TM10 = "020" And strRvType = "1205" Then
      strSql = "update tmgoods set tg01='" & Mid(strCP09, 1, 3) & "', tg02='" & Mid(strCP09, 4, 6) & "' where tg01='" & Mid(m_CP09, 1, 3) & "' and tg02='" & Mid(m_CP09, 4, 6) & "' "
      cnnConnection.Execute strSql, intI
   End If
   'end 2024/11/21
   
   'Added by Morgan 2017/4/17 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strRvType
   End If
   'end 2017/4/17
   
   'Add by Sindy 2019/5/10
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010403_1", strCP09
   End If
   '2019/5/10 END
   
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
   
    cnnConnection.RollbackTrans
    OnSaveData = False
EXITSUB:
   rsTmp.Close
   Set rsTmp = Nothing
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
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   'Add By Cheng 2002/07/18
   Set frm02010403_4 = Nothing
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

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then
      textCP08.SetFocus
   Else
      textCP36.SetFocus
   End If
   
End Sub

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCF15_2 = Empty
   If IsEmptyText(textCF15) = False Then
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
      End If
   End If
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/07
      End If
      'Add By Cheng 2002/03/11
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Cancel = True
         textCP06_GotFocus
         GoTo EXITSUB
      End If
        'Modify By Cheng 2002/11/18
        '按確定時才檢查
'      ' 申請國家為台灣時需檢查來函記錄檔
'      If m_TM10 < "010" Then
'         strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR16")
'         If IsEmptyText(strDate) = False Then
'            If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
'               strTit = "資料檢核"
'               strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
'               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'               If nResponse = vbCancel Then
'                  Cancel = True
'                  textCP06_GotFocus
'                  GoTo EXITSUB
'               End If
'            End If
'         Else
'            strTit = "資料檢核"
'            strMsg = "來函記錄中無該筆記錄"
'            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'            If nResponse = vbCancel Then
'               Cancel = True
'               textCP06_GotFocus
'               GoTo EXITSUB
'            End If
'         End If
'      End If
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
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
         GoTo EXITSUB
      End If
        'Modify By Cheng 2002/11/18
'      ' 申請國家為台灣時需檢查來函記錄檔
'      If m_TM10 < "010" Then
'         strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR17")
'         If IsEmptyText(strDate) = False Then
'            If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
'               strTit = "資料檢核"
'               strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
'               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'               If nResponse = vbCancel Then
'                  Cancel = True
'                  textCP07_GotFocus
'                  GoTo EXITSUB
'               End If
'            End If
'         Else
'            strTit = "資料檢核"
'            strMsg = "來函記錄中無該筆記錄"
'            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'            If nResponse = vbCancel Then
'               Cancel = True
'               textCP07_GotFocus
'               GoTo EXITSUB
'            End If
'         End If
'      End If
   End If
EXITSUB:
End Sub

'Add By Sindy 2010/11/26
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

Private Sub textCP36_GotFocus()
   InverseTextBox textCP36
   CloseIme

End Sub

Private Sub textCP37_1_GotFocus()
   InverseTextBox textCP37_1
   OpenIme
End Sub

'Add by Amy 2025/01/17
Private Sub textCP37_1_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造案件名稱"
   Cancel = False
   If CheckLengthIsOK(textCP37_1, 160, True, strMsg) = False Then
      Cancel = True
      textCP37_1_GotFocus
   End If
End Sub

Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造案件名稱"
   Cancel = False
   If CheckLengthIsOK(textCP40, 600, True, strMsg) = False Then
      Cancel = True
      textCP40_GotFocus
   End If
End Sub

Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造案件名稱"
   Cancel = False
   If CheckLengthIsOK(textCP41, 600, True, strMsg) = False Then
      Cancel = True
      textCP41_GotFocus
   End If
End Sub

Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造案件名稱"
   Cancel = False
   If CheckLengthIsOK(textCP42, 600, True, strMsg) = False Then
      Cancel = True
      textCP42_GotFocus
   End If
End Sub
'end 2025/01/17

Private Sub textCP40_GotFocus()
   InverseTextBox textCP40
   OpenIme
End Sub

Private Sub textCP41_GotFocus()
   InverseTextBox textCP41
   CloseIme

End Sub

Private Sub textCP42_GotFocus()
   InverseTextBox textCP42
   OpenIme
End Sub

'Add By Sindy 2010/7/6
Private Sub textCP80_GotFocus()
   InverseTextBox textCP80
   CloseIme
End Sub

'Add By Sindy 2010/7/6
'對造商品類別
Private Sub textCP80_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP80, textCP80.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造商品類別欄位內容太長"
      textCP80_GotFocus
   End If
End Sub

' 承辦人期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的承辦期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
         Exit Sub
      End If
   End If
   'Add By Cheng 2002/05/06
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         Cancel = True
         textCP48_GotFocus
         Exit Sub
      End If
   End If
   
End Sub

Private Sub textCP49_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 條款
Private Sub textCP49_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textCP49) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textCP49)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textCP49, nIndex)
      'Add By Cheng 2002/07/22
      '條款每項可輸入1~3碼
'      If Len(strTemp) > 4 Then
      'Modify By Sindy 2012/7/5
      'If Len(strTemp) > 3 Or Len(strTemp) < 1 Then
      If Len(strTemp) > 4 Or Len(strTemp) < 1 Then
      '2012/7/5 End
         Cancel = True
         strTit = "條款"
         strMsg = "條款內容<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP49_GotFocus
         GoTo EXITSUB
      End If
      
      ' 90.08.12 modify by sonia
      ' 檢查主張內容分類表
      'StrSQL = "SELECT * FROM ClaimContents " & _
      '         "WHERE CC01 = '" & Right(strTemp, 1) & "'"
      'rsTmp.CursorLocation = adUseClient
      'rsTmp.Open StrSQL, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      'If rsTmp.RecordCount <= 0 Then
      '   Cancel = True
      '   strTit = "條款"
      '   strMsg = "條款內容<" & strTemp & ">不正確"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCP49_GotFocus
      '   rsTmp.Close
      '   GoTo EXITSUB
      'End If
      'rsTmp.Close
      
      ' 檢查
      'Modify By Sindy 2012/7/5
'      strSql = "SELECT * FROM LAW " & _
'               "WHERE LW01 = '" & Mid(strTemp, 1, 3) & "' "
      strSql = "SELECT * FROM LAW " & _
               "WHERE LW01 = '" & Trim(strTemp) & "' "
      '2012/7/5 End
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount <= 0 Then
         Cancel = True
         strTit = "條款"
         strMsg = "條款代號<" & strTemp & ">不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP49_GotFocus
         rsTmp.Close
         GoTo EXITSUB
      End If
      rsTmp.Close
   Next nIndex
EXITSUB:
   Set rsTmp = Nothing
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
      strMsg = "進度備註欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/28檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        GoTo EXITSUB
   End If

   'Add By Sindy 2012/4/17
   '檢查來函期限--日期
   If m_TM10 = 台灣國家代號 Then
      If Me.Option4(2).Value = True Then
         If Me.Text12.Text = "" Then
            MsgBox "請輸入來函期限!!!", vbExclamation + vbOKOnly
            Me.Text12.SetFocus
            GoTo EXITSUB
         End If
      End If
   'Added by Lydia 2019/06/21 台-大核駁案期限管制: 檢查法限和所限
   ElseIf textCP07.Tag <> "" And frm02010403_3.GetSelectResult = "3" Then
       If textCP07.Text > textCP07.Tag Then
            MsgBox "法定期限不可早於" & textCP07.Tag, vbExclamation + vbOKOnly
            Me.textCP07.SetFocus
            GoTo EXITSUB
       End If
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
   ' 本所期限的日期不可超過法定期限的日期
   If Val(textCP06) > Val(textCP07) Then
      strTit = "資料檢核"
      strMsg = "本所期限的日期不可超過法定期限的日期"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP06.SetFocus
      GoTo EXITSUB
   End If
   ' 下一程序不可空白
   If IsEmptyText(textCF15) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入下一程序"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCF15.SetFocus
      GoTo EXITSUB
   End If
   ' 機關文號(申請國家為台灣時不可為空白)
   If IsEmptyText(textCP08) = True Then
      If m_TM10 < "010" Then
         strTit = "資料檢核"
         strMsg = "申請國家為台灣時機關文號不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP08.SetFocus
         GoTo EXITSUB
      End If
   End If
   'Add By Cheng 2002/05/06
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         Me.textCP48.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP08_GotFocus()
   'Modify By Cheng 2002/04/22
   '將游標停在"字"的前面
'   InverseTextBox textCP08
Dim intPos As Integer
With Me.textCP08
   If Len("" & .Text) > 0 Then
      intPos = InStr("" & .Text, "字")
      If intPos - 1 >= 0 Then
         .SelStart = intPos - 1
         .SelLength = 0
      End If
   End If
End With
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textCP49_GotFocus()
   InverseTextBox textCP49
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
'Add By Cheng 2002/11/18
Dim strTit As String
Dim strMsg As String
Dim nResponse

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

'Add by Amy 2025/01/17 輸完直接按Enter鍵不會檢查
If Me.textCP37_1.Enabled = True Then
   Cancel = False
   textCP37_1_Validate Cancel
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
'end 2025/01/17

If Me.textCP48.Enabled = True Then
   Cancel = False
   textCP48_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP49.Enabled = True Then
   Cancel = False
   textCP49_Validate Cancel
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

If Me.textPrint.Enabled = True Then
   Cancel = False
   textPrint_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
    'Add By Cheng 2002/11/18
    ' 申請國家為台灣時需檢查來函記錄檔
    If m_TM10 < "010" Then
       strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR16")
       If IsEmptyText(strDate) = False Then
          If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
             strTit = "資料檢核"
             strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
             nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
             If nResponse = vbCancel Then
                Cancel = True
                textCP06_GotFocus
                Exit Function
             End If
          End If
       '2008/11/27 CANCEL BY SONIA
       'Else
       '   strTit = "資料檢核"
       '   strMsg = "來函記錄中無該筆記錄"
       '   nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
       '   If nResponse = vbCancel Then
       '      Cancel = True
       '      textCP06_GotFocus
       '     Exit Function
       '   End If
       '2008/11/27 END
       '2011/6/15 ADD BY SONIA
       Else
         If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then
         Else
            If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
               strTit = "資料檢核"
               strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  Cancel = True
                  textCP06_GotFocus
                  Exit Function
               End If
            End If
         End If
         '2011/6/15 END
       End If
    End If
    ' 申請國家為台灣時需檢查來函記錄檔
    If m_TM10 < "010" Then
       strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR17")
       If IsEmptyText(strDate) = False Then
          If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
             strTit = "資料檢核"
             strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
             nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
             If nResponse = vbCancel Then
                Cancel = True
                textCP07_GotFocus
                Exit Function
             End If
          End If
       Else
         If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then  '2011/6/15 ADD BY SONIA
            'modify by sonia 2018/2/8 電子公文都不檢查來函記錄檔
            'If m_DocNo = "" Or textCP07 <> "" Then 'Added by Morgan 2017/4/17 電子公文
            If m_DocNo = "" And textCP07 <> "" Then 'Added by Morgan 2017/4/17 電子公文
               strTit = "資料檢核"
               strMsg = "來函記錄中無該筆記錄"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  Cancel = True
                  textCP07_GotFocus
                 Exit Function
               End If
            End If
         '2011/6/15 ADD BY SONIA
         Else
            If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
               strTit = "資料檢核"
               strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  Cancel = True
                  textCP07_GotFocus
                  Exit Function
               End If
            End If
         End If
         '2011/6/15 END
       End If
    End If

TxtValidate = True
End Function

'Add By Sindy 2012/4/17
Private Sub Option1_Click(Index As Integer)
   If m_TM10 = 台灣國家代號 Then 'Addecd by Lydia 2019/06/21
        If Me.Option4(0).Value Then
           Text10_Validate False
        ElseIf Me.Option4(1).Value Then
           Text11_Validate False
        ElseIf Me.Option4(2).Value Then
           Text12_Validate False
        End If
   'Added by Lydia 2019/06/21 台-大核駁案期限管制: 先預設法限和所限，可人工變更；所限=法限-3個工作天。
   Else
        Call GetTime
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme
End Sub

Private Sub Text10_LostFocus()
   '非台灣"天"跳離時到"本所期限"欄位
   If m_TM10 <> 台灣國家代號 Then
      If textCP06.Enabled = True Then textCP06.SetFocus
   End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
   CloseIme
End Sub

Private Sub Text11_LostFocus()
   '非台灣"月"跳離時到"本所期限"欄位
   'If m_TM10 <> 台灣國家代號 Then
   '   If textCP06.Enabled = True Then textCP06.SetFocus
   'End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_LostFocus()
   '非台灣"日"跳離時到"本所期限"欄位
   If m_TM10 <> 台灣國家代號 Then
      If textCP06.Enabled = True Then textCP06.SetFocus
   End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
   Else
      If ChkDate(Text12) Then
         If m_TM10 = 台灣國家代號 Then
            If Val(Text12) < Val(strSrvDate(2)) Then
               MsgBox "來函期限不可小於系統日 !", vbCritical
               Cancel = True
            Else
               textCP07 = Text12
               'Modify By Sindy 2014/10/6 台灣案之本所期限設定
               If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                  textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
               Else
               '2014/10/6 END
                  textCP06 = TransDate(CompDate(2, -2, TransDate(textCP07, 2)), 1)
               End If
               textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
            End If
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub GetTime()
   Dim i As Integer
   'Dim strFromDate As String '期限起算日  'Remove by Lydia 2019/06/21
   
   'Add By Sindy 2012/8/30
   If Option4(0).Value = False And Option4(1).Value = False Then Exit Sub
   '2012/8/30 End
   
   'strFromDate = DBDATE(textCP05)
   'strFromDate = DBDATE(frm02010403_1.textCP05)  'Remove by Lydia 2019/06/21
   
   If m_TM10 = 台灣國家代號 Then
      '文到天數
      If Option4(0).Value = True Then
         textCP07 = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
         If Option1(0).Value = True Then textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
         If Val(Text10) >= 60 Then
            i = -4
         Else
            i = -2
         End If
      '文到月數
      ElseIf Option4(1).Value = True Then
         textCP07 = TAIWANDATE(AddMonth(strFromDate, Val(Text11)))
         If Option1(0).Value = True Then textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
         If Val(Text11) >= 2 Then
            i = -4
         Else
            i = -2
         End If
      End If
      If textCP07 <> "" Then
         'Modify By Sindy 2014/10/6 台灣案之本所期限設定
         If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
            textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
         Else
         '2014/10/6 END
            textCP06 = TransDate(CompDate(2, i, TransDate(textCP07, 2)), 1)
         End If
         textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      End If

   'Added by Lydia 2019/06/21 台-大核駁案期限管制: 先預設法限和所限，可人工變更；所限=法限-3個工作天。
   ElseIf frm02010403_3.GetSelectResult = "3" Then
      textCP06.Tag = ""
      textCP07.Tag = ""
      If Option1(0).Value = True Then ' 紙本公文不可大於15日曆天
           i = 15
      Else  '電子公文不可大於30日曆天
           i = 30
      End If
      strExc(1) = CompDate(2, i, strFromDate)
      If strExc(1) <> "" Then
         strExc(2) = CompWorkDay(4, strExc(1), 1)
         If strExc(2) < strSrvDate(1) Then
             strExc(2) = strSrvDate(1)
         End If
      End If
      textCP06.Text = TransDate(strExc(2), 1)
      textCP06.Tag = textCP06.Text
      textCP07.Text = TransDate(strExc(1), 1)
      textCP07.Tag = textCP07.Text
   End If
End Sub

'讀取來函期限
Private Function ChgType() As Boolean
Dim strTempName As String, bolTmp As Boolean
Dim i As Integer
'Remove by Lydia 2019/06/21
'Dim strFromDate As String '期限起算日
   
'   'strFromDate = DBDATE(textCP05)
'   strFromDate = DBDATE(frm02010403_1.textCP05)
'end 2019/06/21
   ChgType = False
   If m_TM10 = 台灣國家代號 Then
      bolTmp = False
   Else
      bolTmp = True
   End If
   
   ' 案件性質
   strRvType = "1002"
   Select Case frm02010403_3.GetSelectResult
      Case "1": strRvType = "1201"
      Case "2": strRvType = "1202"
      '92.8.1 ADD
      Case "3": strRvType = "1205"
   End Select
   If strRvType = "" Then Exit Function
   
   If ClsPDGetCaseProperty(m_TM01, strRvType, strTempName, bolTmp) Then
      textCP06 = ""
      textCP07 = ""
      
      If m_TM10 = 台灣國家代號 Then
         strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & m_TM01 & "' AND CPM02='" & strRvType & "'"
         If strExc(0) <> "" Then
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            With RsTemp
               If intI = 1 Then
                  If Not IsNull(.Fields(1)) Then
                     '文到天數
                     Option4(0).Value = True
                     Text10 = .Fields(1)
                     textCP07 = TransDate(CompDate(2, Text10, TransDate(strFromDate, 2)), 1)
                  ElseIf Not IsNull(.Fields(2)) Then
                     '文到月數
                     Option4(1).Value = True
                     Text11 = .Fields(2)
                     textCP07 = TransDate(CompDate(1, .Fields(2), TransDate(strFromDate, 2)), 1)
                  Else
                     '文到天數
                     Option4(0).Value = True
                     Text10 = ""
                     Text11 = ""
                  End If
                  If textCP07 <> "" And Not IsNull(.Fields(0)) Then
                     '文到當日
                     If .Fields(0) = "1" Then
                        Option1(0).Value = True
                        textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
                     '文到次日
                     Else
                        Option1(1).Value = True
                     End If
                  End If
                  '文到天數
                  If Text10 <> "" Then
                     If Val(Text10) >= 60 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  '文到月數
                  ElseIf Not IsNull(.Fields(2)) Then
                     If Val(.Fields(2)) >= 2 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  End If
                  If textCP07 <> "" Then
                     'Modify By Sindy 2014/10/6 台灣案之本所期限設定
                     If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                        textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
                     Else
                     '2014/10/6 END
                        textCP06 = TransDate(CompDate(2, i, TransDate(textCP07, 2)), 1)
                     End If
                     textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                  End If

               End If
            End With
         End If
      'Added by Lydia 2019/06/21 台-大核駁案期限管制: 先預設法限和所限，可人工變更；所限=法限-3個工作天。
      Else
           Call GetTime
      End If
      ChgType = True
   End If
End Function




