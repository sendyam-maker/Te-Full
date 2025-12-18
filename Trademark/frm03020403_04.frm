VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020403_04 
   BorderStyle     =   1  '單線固定
   Caption         =   "審查報告輸入"
   ClientHeight    =   6384
   ClientLeft      =   -3252
   ClientTop       =   4836
   ClientWidth     =   9132
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6384
   ScaleWidth      =   9132
   Begin VB.TextBox textAddDate 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   18
      Top             =   1798
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   510
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   510
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1476
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1154
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1476
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1798
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   2
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5880
      TabIndex        =   0
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6840
      TabIndex        =   1
      Top             =   30
      Width           =   1212
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4200
      Left            =   30
      TabIndex        =   22
      Top             =   2130
      Width           =   9080
      _ExtentX        =   16002
      _ExtentY        =   7408
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "一般"
      TabPicture(0)   =   "frm03020403_04.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "textCP14_2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "textCP64"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label32"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label26"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label25"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label24"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label21"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label22"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label23"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label16"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label15"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "grdList"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textCP49"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCP48"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCP07"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCP14"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCP06"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCF15_2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCF15"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCP08"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textPrint"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCP26"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text10"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text11"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text12"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "對造/其他"
      TabPicture(1)   =   "frm03020403_04.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textCP36"
      Tab(1).Control(1)=   "textCP41"
      Tab(1).Control(2)=   "textCP80"
      Tab(1).Control(3)=   "Label35"
      Tab(1).Control(4)=   "Label31"
      Tab(1).Control(5)=   "Label29"
      Tab(1).Control(6)=   "Label28"
      Tab(1).Control(7)=   "Label30"
      Tab(1).Control(8)=   "Label2"
      Tab(1).Control(9)=   "textCP40"
      Tab(1).Control(10)=   "textCP42"
      Tab(1).Control(11)=   "textCP37_1"
      Tab(1).ControlCount=   12
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   6780
         MaxLength       =   7
         TabIndex        =   70
         Top             =   810
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   5790
         MaxLength       =   2
         TabIndex        =   69
         Top             =   810
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   68
         Top             =   810
         Width           =   375
      End
      Begin VB.TextBox textCP36 
         Height          =   264
         Left            =   -73500
         MaxLength       =   200
         TabIndex        =   58
         Top             =   390
         Width           =   7092
      End
      Begin VB.TextBox textCP41 
         Height          =   264
         Left            =   -73500
         TabIndex        =   57
         Top             =   1770
         Width           =   7092
      End
      Begin VB.TextBox textCP80 
         Height          =   264
         Left            =   -73500
         MaxLength       =   39
         TabIndex        =   56
         Top             =   2370
         Width           =   3495
      End
      Begin VB.TextBox textCP26 
         Height          =   285
         Left            =   6000
         MaxLength       =   1
         TabIndex        =   39
         Top             =   3195
         Width           =   372
      End
      Begin VB.TextBox textPrint 
         Height          =   285
         Left            =   1140
         MaxLength       =   1
         TabIndex        =   38
         Top             =   3195
         Width           =   732
      End
      Begin VB.TextBox textCP08 
         Height          =   285
         Left            =   1140
         MaxLength       =   40
         TabIndex        =   37
         Top             =   360
         Width           =   2532
      End
      Begin VB.TextBox textCF15 
         Height          =   285
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   36
         Top             =   360
         Width           =   732
      End
      Begin VB.TextBox textCF15_2 
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   360
         Width           =   1692
      End
      Begin VB.TextBox textCP06 
         Height          =   285
         Left            =   1140
         MaxLength       =   7
         TabIndex        =   34
         Top             =   1215
         Width           =   2532
      End
      Begin VB.TextBox textCP14 
         Height          =   285
         Left            =   1140
         MaxLength       =   6
         TabIndex        =   33
         Top             =   2865
         Width           =   732
      End
      Begin VB.TextBox textCP07 
         Height          =   285
         Left            =   5640
         MaxLength       =   7
         TabIndex        =   32
         Top             =   1215
         Width           =   2532
      End
      Begin VB.TextBox textCP48 
         Height          =   285
         Left            =   5640
         MaxLength       =   7
         TabIndex        =   31
         Top             =   2865
         Width           =   2532
      End
      Begin VB.TextBox textCP49 
         Height          =   285
         Left            =   1140
         MaxLength       =   300
         TabIndex        =   30
         Top             =   2550
         Width           =   7812
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   3990
         TabIndex        =   26
         Top             =   675
         Width           =   4215
         Begin VB.OptionButton Option4 
            Caption         =   "文到          天"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   180
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "        月"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   28
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "                      日"
            Height          =   225
            Index           =   2
            Left            =   2520
            TabIndex        =   27
            Top             =   180
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   1140
         TabIndex        =   23
         Top             =   675
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "文到當日"
            Height          =   180
            Index           =   0
            Left            =   144
            TabIndex        =   25
            Top             =   180
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "文到次日"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   24
            Top             =   180
            Width           =   1095
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   972
         Left            =   1104
         TabIndex        =   71
         Top             =   1512
         Width           =   7812
         _ExtentX        =   13780
         _ExtentY        =   1715
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
      Begin VB.Label Label35 
         Caption         =   "對造號數 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   67
         Top             =   390
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "對造中文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   66
         Top             =   1515
         Width           =   1300
      End
      Begin VB.Label Label29 
         Caption         =   "對造英文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   65
         Top             =   1800
         Width           =   1300
      End
      Begin VB.Label Label28 
         Caption         =   "對造日文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   64
         Top             =   2040
         Width           =   1300
      End
      Begin VB.Label Label30 
         Caption         =   "對造案件名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   63
         Top             =   690
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "對造商品類別 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   62
         Top             =   2370
         Width           =   1305
      End
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73500
         TabIndex        =   61
         Top             =   1470
         Width           =   7095
         VariousPropertyBits=   679493659
         Size            =   "12509;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73500
         TabIndex        =   60
         Top             =   2040
         Width           =   7095
         VariousPropertyBits=   679493659
         Size            =   "12515;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37_1 
         Height          =   795
         Left            =   -73500
         TabIndex        =   59
         Top             =   660
         Width           =   7095
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "12509;1397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "本案期限 :"
         Height          =   255
         Left            =   60
         TabIndex        =   55
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "(N:不算)"
         Height          =   255
         Left            =   6510
         TabIndex        =   54
         Top             =   3210
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "是否算案件數 :"
         Height          =   255
         Left            =   4680
         TabIndex        =   53
         Top             =   3210
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不印)"
         Height          =   255
         Left            =   1980
         TabIndex        =   52
         Top             =   3210
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   255
         Left            =   60
         TabIndex        =   51
         Top             =   3210
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   60
         TabIndex        =   50
         Top             =   3525
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "機關文號 :"
         Height          =   255
         Left            =   60
         TabIndex        =   49
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "下一程序 :"
         Height          =   255
         Left            =   4680
         TabIndex        =   48
         Top             =   375
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "本所期限 :"
         Height          =   255
         Left            =   60
         TabIndex        =   47
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "承辦人 :"
         Height          =   255
         Left            =   60
         TabIndex        =   46
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   255
         Left            =   4680
         TabIndex        =   45
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label26 
         Caption         =   "承辦期限 :"
         Height          =   255
         Left            =   4680
         TabIndex        =   44
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "條款 :"
         Height          =   255
         Left            =   60
         TabIndex        =   43
         Top             =   2565
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "來函期限:"
         Height          =   255
         Left            =   60
         TabIndex        =   42
         Top             =   795
         Width           =   855
      End
      Begin MSForms.TextBox textCP64 
         Height          =   525
         Left            =   1140
         TabIndex        =   41
         Top             =   3525
         Width           =   7815
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13785;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   285
         Left            =   1950
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2865
         Width           =   1785
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "3149;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1170
      TabIndex        =   21
      Top             =   832
      Width           =   7710
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13600;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1170
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1154
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "優先權證明期限 :"
      Height          =   180
      Left            =   4215
      TabIndex        =   19
      Top             =   1850
      Width           =   1350
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
      Left            =   3750
      TabIndex        =   17
      Top             =   562
      Width           =   645
   End
   Begin VB.Label Label27 
      Caption         =   "申請案號 :"
      Height          =   255
      Left            =   4710
      TabIndex        =   16
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   14
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   90
      TabIndex        =   13
      Top             =   810
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   90
      TabIndex        =   12
      Top             =   1169
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   255
      Index           =   2
      Left            =   4710
      TabIndex        =   11
      Top             =   1491
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   255
      Index           =   3
      Left            =   4710
      TabIndex        =   10
      Top             =   1169
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   90
      TabIndex        =   9
      Top             =   1491
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   90
      TabIndex        =   8
      Top             =   1813
      Width           =   1215
   End
End
Attribute VB_Name = "frm03020403_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/20 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/09/13 改成Form2.0 ; cmbTM05、textTM23、textCP14_2、textCP64、grdList改字型=新細明體-ExtB
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
' 收文號
Dim m_CP09 As String
' 原案件性質
Dim m_CP10 As String
' 原業務區
Dim m_CP12 As String
' 原智權人員代號
Dim m_CP13 As String
' 國家代碼
Dim m_TM10 As String

Dim m_CurrSel As Integer
Dim strRvType As String 'Add By Sindy 2012/4/26
Dim strTextAddDate_NP08 As String 'Add By Sindy 2015/8/20 優先權證明期限的本所期限
'Added by Morgan 2017/5/4 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/5/4
Dim m_NewCP09 As String 'Added by Lydia 2022/02/10 新增C類收文號

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm03020403_03.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm03020403_03
   Unload frm03020403_02
   Unload frm03020403_01
   Unload Me
End Sub

Private Sub cmdok_Click()
Dim strFilePath As String 'Added by Lydia 2022/02/10 掃瞄檔的路徑

   If CheckDataValid = True Then
         'Added by Lydia 2022/02/10 FCT紙本公文來函，同時將公文函FCT_OA_SCAN匯入卷宗區
      If frm03020403_03.GetSelectResult() = "1" Then
        If m_DocNo = "" Then
            If PUB_FCTCheckPDF(m_TM01, m_TM02, m_TM03, m_TM04, "1201", m_CP09, strFilePath) = False Then
                 Exit Sub
            End If
        End If
      End If
      'end 2022/02/10
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      'edit by  nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
       'Added by Lydia 2022/02/10 FCT紙本公文來函，同時將公文函FCT_OA_SCAN匯入卷宗區
       'Move by Lydia 2022/02/23 從frm03020403_01.Show上方移過來
       If strFilePath <> "" Then
           If Pub_AutoSavePdf2_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_NewCP09, strRvType, strFilePath) = False Then
               Exit Sub
           End If
       End If
       'end 2022/02/10
       
      Unload Me
      Unload frm03020403_03
      Unload frm03020403_02
      'Modified by Morgan 2017/5/4 電子公文
      'frm03020403_01.Show
      If m_DocNo <> "" Then
         Unload frm03020403_01
         frm02010412.GoNext
      Else
         frm03020403_01.Show
      End If
      'end 2017/5/4
   End If
End Sub

'Added by Morgan 2022/1/11
Private Sub Form_Activate()
   Static bDone As Boolean
   
   If bDone = False Then
      '電子公文游標預設在下一程序--陳金蓮
      If m_DocWord <> "" And textCF15.Enabled Then
         textCF15.SetFocus
      End If
      bDone = True
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCF15_2.BackColor = &H8000000F
   textAddDate.BackColor = &H8000000F 'Add By Sindy 2015/7/29
   
   MoveFormToCenter Me
   'Add by Amy 2022/09/05
   'Modify by Amy 2022/09/26 GetSelectResult() = "1" 審查報表也改為關係案
   'If frm03020403_03.GetSelectResult() = "2" Then
        SSTab1.TabCaption(1) = "關係案"
        strExc(1) = "對方"
        Label35.Caption = strExc(1) & Mid(Label35.Caption, 3)
        Label30.Caption = strExc(1) & Mid(Label30.Caption, 3)
        Label31.Caption = strExc(1) & Mid(Label31.Caption, 3)
        Label29.Caption = strExc(1) & Mid(Label29.Caption, 3)
        Label28.Caption = strExc(1) & Mid(Label28.Caption, 3)
        Label27.Caption = strExc(1) & Mid(Label27.Caption, 3)
        Label2.Caption = strExc(1) & Mid(Label2.Caption, 3)
    'End If
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

' 讀取商標基本檔
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
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
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
      If IsNull(rsTmp.Fields("TM08")) = False Then
         If m_TM10 < "010" Then
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
         Else
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
         End If
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
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      '91.7.9 modify by sonia 根本不必帶此欄
      ' 機關文號
      'If IsNull(rsTmp.Fields("CP08")) = False Then
      '   textCP08 = rsTmp.Fields("CP08")
      'End If
      ' end 91.7.9
      
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 業務區
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      'Modified by Lydia 2021/08/03 改由PUB_GetFCTSalesNo帶出和產生的C類收文一致
      'If IsNull(rsTmp.Fields("CP13")) = False Then
      '   m_CP13 = rsTmp.Fields("CP13")
      'End If
      m_CP13 = Empty
      m_CP13 = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      'end 2021/08/03
      
      ' 下一程序
      textCF15 = GetNextProgress(m_TM01, m_TM10, m_CP10)
      If IsEmptyText(textCF15) = False Then
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
      End If
      '91.7.9 modify by sonia 根本不必帶此欄
      ' 本所期限
      'If IsNull(rsTmp.Fields("CP06")) = False Then
      '   If IsEmptyText(rsTmp.Fields("CP06")) = False Then
      '      textCP06 = TAIWANDATE(rsTmp.Fields("CP06"))
      '   End If
      'End If
      ' end 91.7.9
      '91.7.9 modify by sonia 根本不必帶此欄
      ' 法定期限
      'If IsNull(rsTmp.Fields("CP07")) = False Then
      '   If IsEmptyText(rsTmp.Fields("CP07")) = False Then
      '      textCP07 = TAIWANDATE(rsTmp.Fields("CP07"))
      '   End If
      'End If
      ' end 91.7.9
'      ' 承辦人預設為點選資料之智權人員
'      If IsNull(rsTmp.Fields("CP13")) = False Then
'         textCP14 = rsTmp.Fields("CP13")
'         textCP14_2 = GetStaffName(rsTmp.Fields("CP13"))
'      End If
        '預設承辦人
        Me.textCP14.Text = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
        Me.textCP14_2.Text = GetStaffName(Me.textCP14.Text)
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   ' 90.11.19 modify by sonia
   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   If textCP08 = "" Then
      textCP08 = "（" & strTmp & "）慧商字第號"
   End If

End Sub

Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strDay As String
      
   ' 來函收文日
   textCP05S = m_CP05
   ' 本所案號
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   
   ' 讀取商標基本檔
   QueryTradeMark
   
   ' 讀取案件進度檔
   QueryCaseProgress
   
   Call ChgType 'Add By Sindy 2012/4/17 讀取來函期限
   
   'Added by Morgan 2017/5/4 電子公文
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
   'end 2017/5/4
   
   ' 以案件性質"審查報告"或"核駁前先行通知"計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''   strDay = Empty
   Select Case frm03020403_03.GetSelectResult
      Case "1":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1201")
            'modify by sonia 2018/8/17 來函收文日+10個工作天自動產生承辦期限,但若大於本所期限時以本所期限減3工作天為承辦期限FCT-041983
            'textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1201", DBDATE(m_CP05), DBDATE(textCP06)))
            strDay = DBDATE(CompWorkDay(-3, DBDATE(textCP06), 1))
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1201", DBDATE(m_CP05), strDay))
      Case "2":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1202")
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1202", DBDATE(m_CP05), DBDATE(textCP06)))
   End Select
''''   If IsEmptyText(strDay) = False Then
''''      ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''      'textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''      textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''   End If
   
   'Add By Sindy 2015/7/29 當點選為申請案且是審查報告時,檢查是否有優先權證明期限
   textAddDate = ""
   strTextAddDate_NP08 = "" 'Add By Sindy 2015/8/20 優先權證明期限的本所期限
   If m_CP10 = "101" And frm03020403_03.GetSelectResult = "1" Then
      '有未收文的208補優先權證明期限
      strSql = "SELECT NP01,NP08,NP09 FROM NextProgress " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07='208' AND NP06 is null"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         textAddDate = Val(rsTmp.Fields("NP09")) - 19110000 '優先權證明期限=法定期限
         strTextAddDate_NP08 = Val(rsTmp.Fields("NP08")) - 19110000 'Add By Sindy 2015/8/20 優先權證明期限的本所期限
         If Val(textCP48) > Val(strTextAddDate_NP08) Then textCP48 = strTextAddDate_NP08   'add by sonia 2018/5/9
      Else
         rsTmp.Close
         '有未發文的208補優先權證明期限
         strSql = "SELECT CP09,CP06,CP07 FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP10='208' AND CP27 is null AND CP57 is null"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            textAddDate = Val(rsTmp.Fields("CP07")) - 19110000 '優先權證明期限=法定期限
            strTextAddDate_NP08 = Val(rsTmp.Fields("CP06")) - 19110000 'Add By Sindy 2015/8/20 優先權證明期限的本所期限
            If Val(textCP48) > Val(strTextAddDate_NP08) Then textCP48 = strTextAddDate_NP08   'add by sonia 2018/5/9
         End If
      End If
      rsTmp.Close
   End If
   '2015/7/29 END
   
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   strSql = "SELECT NP01,NP07,NP08,NP09,NP10,NP11,NP12,NP13,NP14,NP15,NP22 FROM NextProgress " & _
            "WHERE NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP06 IS NULL" & strNpSqlOfNoSalesDuty
   'Add By Sindy 2015/8/24
   If m_CP10 = "101" Or m_CP10 = "308" Then '申請或分割
      strSql = strSql & " UNION " & _
               "SELECT CP09,CP10,CP06,CP07,CP13,CP57,CP58,CP08,CP40,CP64,0 FROM CASEPROGRESS " & _
               "WHERE CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "'" & _
               " AND CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "'" & _
               " AND CP09<'C' and cp07>0 AND CP27 IS NULL AND CP57 IS NULL"
   End If
   '2015/8/24 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
'         ' 是否續辦欄位必須為空白
'         If IsNull(rsTmp.Fields("NP06")) = False Then
'            If IsEmptyText(rsTmp.Fields("NP06")) = False Then
'               GoTo NextRecord
'            End If
'         End If
         
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
         'If IsNull(rsTmp.Fields("NP22")) = False Then
         If Val((rsTmp.Fields("NP22"))) > 0 Then
            grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("NP22")
         'Add By Sindy 2015/8/24
         Else
            grdList.TextMatrix(grdList.row, 11) = "已收文"
         '2015/8/24 END
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
   Set rsTmp = Nothing
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 12 '11 'Modify By Sindy 2009/05/15
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
   grdList.ColAlignment(6) = flexAlignLeftCenter 'Add By Sindy 2015/8/20 儲存格內容中間靠左對齊
   grdList.col = 7
   grdList.Text = "收文號"
   grdList.ColWidth(7) = 0
   grdList.col = 8
   grdList.Text = "下一程序代號"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "序號"
   grdList.ColWidth(9) = 0
   'Add By Sindy 2009/05/15
   grdList.col = 10
   grdList.Text = "是否更新"
   grdList.ColWidth(10) = 0
   '2009/05/15 End
   'Add By Sindy 2015/8/24
   grdList.col = 11
   grdList.Text = "狀態"
   grdList.ColWidth(11) = 800
   '2015/8/24 End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Add By Cheng 2002/07/19
   Set frm03020403_04 = Nothing
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

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strSql As String
   Dim strCP09 As String
   'Dim strCP12 As String
   Dim strNP07 As String
   Dim strNP14 As String
   Dim strNP22 As String
   Dim nIndex As Integer
   Dim strNP15 As String
   'Dim bolHadNP As Boolean 'Add By Sindy 2015/8/24
   
   OnSaveData = True
   
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans
      
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '  新增資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   m_NewCP09 = strCP09 'Added by Lydia 2022/02/10 新增C類收文號
   ' 案件性質為審查報告或核駁前先行通知
   strRvType = "1002"
   Select Case frm03020403_03.GetSelectResult
      Case "1": strRvType = "1201"
      Case "2": strRvType = "1202"
   End Select
   ' 業務區別 91.8.26 MODIFY BY SONIA
   'strCP12 = GetStaffDepartment(m_CP13)
   ' 組成SQL語法
   ' 91.03.25 modify by louis (單引號)
    'Modify By Cheng 2003/04/07
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng 2003/09/05
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP43,CP49,CP64) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
'                    "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "') "
   'Modify by Amy 2022/09/05 +CP36~42及CP80
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP36,CP37,CP40,CP41,CP42,CP80,CP43,CP49,CP64) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                    "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & ChgSQL(textCP36) & "','" & ChgSQL(textCP37_1) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & ChgSQL(textCP42) & "','" & ChgSQL(textCP80) & "','" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "') "
   cnnConnection.Execute strSql
   
   'Modify By Sindy 2015/8/21
   '未輸入下一程序時, 更新208優先權證明期限
   If IsEmptyText(textCF15) = True Then
      ' 本所期限
      If IsEmptyText(strTextAddDate_NP08) = False Then
         strSql = "UPDATE CaseProgress SET CP06 = " & DBDATE(strTextAddDate_NP08) & " " & _
                  "WHERE CP09 = '" & strCP09 & "' "
         cnnConnection.Execute strSql
      End If
      ' 法定期限
      If IsEmptyText(textAddDate) = False Then
         strSql = "UPDATE CaseProgress SET CP07 = " & DBDATE(textAddDate) & " " & _
                  "WHERE CP09 = '" & strCP09 & "' "
         cnnConnection.Execute strSql
      End If
   Else
   '2015/8/21 END
      ' 有輸入本所期限時
      If IsEmptyText(textCP06) = False Then
         strSql = "UPDATE CaseProgress SET CP06 = " & DBDATE(textCP06) & " " & _
                  "WHERE CP09 = '" & strCP09 & "' "
         cnnConnection.Execute strSql
      End If
      ' 有輸入法定期限時
      If IsEmptyText(textCP07) = False Then
         strSql = "UPDATE CaseProgress SET CP07 = " & DBDATE(textCP07) & " " & _
                  "WHERE CP09 = '" & strCP09 & "' "
         cnnConnection.Execute strSql
      End If
   End If
   
   'Add By Sindy 2012/4/26 儲存官方發文日及官方期限月數
   If Trim(Text11) <> "" Then
      strSql = "UPDATE CaseProgress SET CP133=" & DBDATE(m_CP05) & ",CP134=" & Text11 & " " & _
               "WHERE CP09='" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   ' 有輸入承辦人時
   If IsEmptyText(textCP14) = False Then
      strSql = "UPDATE CaseProgress SET CP14 = '" & textCP14 & "' " & _
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
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 有輸入下一程序時, 新增資料到下一程序檔
   If IsEmptyText(textCF15) = False Then
      'Add By Sindy 2015/8/20 將備註內容串起來
      strNP15 = ""
      'bolHadNP = False 'Add By Sindy 2015/8/24
      For nIndex = 1 To grdList.Rows - 1
         ' 判斷該列是否有被選取
         If grdList.TextMatrix(nIndex, 0) = "V" Then
            'Modify By Sindy 2015/8/24
'            If Val(grdList.TextMatrix(nIndex, 9)) > 0 Then
'               bolHadNP = True
'            End If
            'If Trim(grdList.TextMatrix(nIndex, 6)) <> "" Then
            If Trim(grdList.TextMatrix(nIndex, 6)) <> "" And Val(grdList.TextMatrix(nIndex, 9)) > 0 Then
            '2015/8/24 END
               strNP15 = strNP15 & Trim(grdList.TextMatrix(nIndex, 6)) & ";"
            End If
         End If
      Next nIndex
      '2015/8/20 END
      
'      If bolHadNP = True Then 'Add By Sindy 2015/8/24 +if
         strNP14 = Empty
         strNP14 = GetRelatedPerson(m_CP09)
         ' 序號
         strNP22 = GetNextProgressNo()
         ' 組成SQL語法
         'Modify By Cheng 2002/09/25
   '      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
   '               "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
   '                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "','" & textCP64 & "'," & strNP22 & ")"
           'Modify By Cheng 2003/04/07
           '智權人員存最近收文A類接洽記錄單的智權人員
           'Modify By Cheng 2003/09/05
   '      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
   '               "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
   '                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "','" & textCP64 & "'," & strNP22 & ")"
         'Modified by Lydia 2024/07/02 +ChgSQL
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
                  "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                             DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "','" & ChgSQL(textCP64 & IIf(textCP64 <> "" And strNP15 <> "", ";", "") & strNP15) & "'," & strNP22 & ")"
         cnnConnection.Execute strSql
         ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
         Select Case textCF15
            Case "102", "105", "702", "708", "305", "998", "997":
            Case Else:
               'Modify By Cheng 2002/12/05
               '恢復列印接洽結案單
   '            'Modify By Cheng 2002/01/15
   '            '取消外商FCT列印接洽結案單
               ' 列印國內案件接洽及結案記錄單
   '            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
               'Add By Cheng 2003/06/23
               '新增列印接洽結案單資料
               pub_AddressListSN = pub_AddressListSN + 1
               PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
         End Select
'      End If
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新使用者所選取的本案期限資料
   For nIndex = 1 To grdList.Rows - 1
      ' 判斷該列是否有被選取
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         strNP07 = grdList.TextMatrix(nIndex, 8)
         strNP22 = grdList.TextMatrix(nIndex, 9)
         If Val(strNP22) > 0 Then
            strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
                     "WHERE NP02 = '" & m_TM01 & "' AND " & _
                           "NP03 = '" & m_TM02 & "' AND " & _
                           "NP04 = '" & m_TM03 & "' AND " & _
                           "NP05 = '" & m_TM04 & "' AND " & _
                           "NP07 = " & strNP07 & " AND " & _
                           "NP22 = " & strNP22 & " "
            cnnConnection.Execute strSql
         'Modify By Sindy 2015/8/24
         Else '申請或分割案的np22=0
            If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
               strSql = "UPDATE CaseProgress SET CP06 = " & DBDATE(textCP06) & ",CP07 = " & DBDATE(textCP07) & _
                        ",CP64 = '" & ChangeWStringToTDateString(DBDATE(textCP05S)) & "審查報告更新期限;'||CP64" & _
                        " WHERE CP01 = '" & m_TM01 & "' AND " & _
                              "CP02 = '" & m_TM02 & "' AND " & _
                              "CP03 = '" & m_TM03 & "' AND " & _
                              "CP04 = '" & m_TM04 & "' AND " & _
                              "CP09 = '" & grdList.TextMatrix(nIndex, 7) & "' "
               cnnConnection.Execute strSql
            End If
         End If
         '2015/8/24 END
      End If
   Next nIndex

   'Add By Sindy 2009/05/15
   'FCT之101.申請,102.延展,301.變更,501.移轉,502.授權等案件
   '判斷輸入的法定期限+6個月是否有大於催審的法定日期,否則更新之
   If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "301" Or _
      m_CP10 = "501" Or m_CP10 = "502" Then
      'Modified by Lydia 2016/05/09 點選之案件性質為申請101、延展102、變更301、移轉501、授權502時，在存檔時以法定期限+8個月更新點選進度之下一程序催審305期限，本所期限及法定期限都要更新。
'      For nIndex = 1 To grdList.Rows - 1
'         If Trim(grdList.TextMatrix(nIndex, 8)) = "305" And _
'            Trim(grdList.TextMatrix(nIndex, 10)) = "Y" Then '305.催審並且Y.要更新
'            strNP07 = grdList.TextMatrix(nIndex, 8)
'            strNP22 = grdList.TextMatrix(nIndex, 9)
'            strSql = "UPDATE NextProgress " & _
'                     "SET NP08 = " & ChangeTStringToWString(grdList.TextMatrix(nIndex, 2)) & ", " & _
'                         "NP09 = " & ChangeTStringToWString(grdList.TextMatrix(nIndex, 3)) & " " & _
'                     "WHERE NP02 = '" & m_TM01 & "' AND " & _
'                           "NP03 = '" & m_TM02 & "' AND " & _
'                           "NP04 = '" & m_TM03 & "' AND " & _
'                           "NP05 = '" & m_TM04 & "' AND " & _
'                           "NP07 = " & strNP07 & " AND " & _
'                           "NP22 = " & strNP22 & " "
'            cnnConnection.Execute strSql
'         End If
'      Next nIndex
      strExc(1) = CompDate(1, "8", ChangeTStringToWString(Trim(textCP07)))
      'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天 + PUB_GetWorkDay1()
      'Modified by Lydia 2023/11/13 若依原設定規則更新之催審期限小於原催審期限，則不更新原催審期限。=>AND NP09<
      strSql = "UPDATE NextProgress Set NP08=" & PUB_GetWorkDay1(strExc(1), True) & ", NP09=" & strExc(1) & _
               " WHERE NP01='" & m_CP09 & "' AND NP07='305' AND NP06 IS NULL AND NP09<" & strExc(1)
      cnnConnection.Execute strSql
   End If
   '2009/05/15 End
   
   'Added by Morgan 2017/5/4 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strRvType
   End If
   'end 2017/5/4
   
 '911107 nick transation
  cnnConnection.CommitTrans
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
    'edit by nick 2004/11/03
    OnSaveData = False
End Function

'Add by Amy 2022/09/05
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

'Add by Amy 2022/09/05 +關係案頁籤
Private Sub textCP36_GotFocus()
    InverseTextBox textCP36
End Sub

Private Sub textCP37_1_GotFocus()
    InverseTextBox textCP37_1
End Sub

'Add by Amy 2025/01/17
Private Sub textCP37_1_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label30, ":", "")
   Cancel = False
   If CheckLengthIsOK(textCP37_1, 160, True, strMsg) = False Then
      Cancel = True
      textCP37_1_GotFocus
   End If
End Sub

Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label31, ":", "")
   Cancel = False
   If CheckLengthIsOK(textCP40, 600, True, strMsg) = False Then
      Cancel = True
      textCP40_GotFocus
   End If
End Sub

Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label29, ":", "")
   Cancel = False
   If CheckLengthIsOK(textCP41, 600, True, strMsg) = False Then
      Cancel = True
      textCP41_GotFocus
   End If
End Sub

Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label28, ":", "")
   Cancel = False
   If CheckLengthIsOK(textCP42, 600, True, strMsg) = False Then
      Cancel = True
      textCP42_GotFocus
   End If
End Sub
'end 2025/01/17

Private Sub textCP40_GotFocus()
   InverseTextBox textCP40
End Sub

Private Sub textCP41_GotFocus()
   InverseTextBox textCP41
End Sub

Private Sub textCP42_GotFocus()
   InverseTextBox textCP42
End Sub

Private Sub textCP80_GotFocus()
   InverseTextBox textCP80
End Sub

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
'end 2022/09/05

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
      'Modify By Cheng 2002/07/22
      '條款每項可輸入1~3碼
'      If Len(strTemp) > 3 Then
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
      
      ' 檢查
      'Modify By Sindy 2012/7/5
'      strSql = "SELECT * FROM LAW " & _
'               "WHERE LW01 = '" & Left(strTemp, 3) & "' "
      strSql = "SELECT * FROM LAW " & _
               "WHERE LW01 = '" & Trim(strTemp) & "' "
      '2012/7/5 End
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04
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
      strTit = "資料檢核"
      strMsg = "進度備註資料內容長度太長"
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

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      ' 檢查日期的格式
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/09
      End If
      'Add By Cheng 2002/03/11
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Cancel = True
         textCP06_GotFocus
         GoTo EXITSUB
      End If
      
      If Val(textCP48) > Val(textCP06) Then textCP48 = textCP06   'add by sonia 2018/5/9
      
'2011/6/15 CANCEL BY SONIA
'按下確定時才檢查
'      ' 申請國家為台灣時需檢查來函記錄檔
'      If m_TM10 < "010" Then
'         strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05, "MR16")
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
Public Sub textCP07_Validate(Cancel As Boolean)
   Dim strDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      ' 檢查日期的格式
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
         GoTo EXITSUB
      End If
      
'2011/6/15 CANCEL BY SONIA
'按下確定時才檢查
'      ' 申請國家為台灣時需檢查來函記錄檔
'      If m_TM10 < "010" Then
'         strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05, "MR17")
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
     
      'Add By Sindy 2009/05/15
      'FCT之101.申請,102.延展,301.變更,501.移轉,502.授權等案件
      '判斷輸入的法定期限+6個月是否有大於催審的法定日期,否則更新之
      'Remove by Lydia 2016/05/09
'      If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "301" Or _
'         m_CP10 = "501" Or m_CP10 = "502" Then
'         '輸入的法定期限+6個月
'         strDate = CompDate(1, "6", ChangeTStringToWString(Trim(textCP07)))
'         For nIndex = 1 To grdList.Rows - 1
'            'Modify By Sindy 2015/8/24
'            'If Trim(grdList.TextMatrix(nIndex, 8)) = "305" Then '305.催審
'            If Trim(grdList.TextMatrix(nIndex, 8)) = "305" And Val(grdList.TextMatrix(nIndex, 9)) > 0 Then  '305.催審
'            '2015/8/24 END
'               If Val(ChangeTStringToWString(grdList.TextMatrix(nIndex, 3))) < Val(strDate) Then
'                  grdList.TextMatrix(nIndex, 2) = Val(strDate) - 19110000 '催審本所期限
'                  grdList.TextMatrix(nIndex, 3) = Val(strDate) - 19110000 '催審法定期限=輸入的法定期限+6個月
'                  grdList.TextMatrix(nIndex, 10) = "Y" '要更新
'               End If
'            End If
'         Next nIndex
'      End If
      '2009/05/15 End
   End If
EXITSUB:
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   ' 機關文號不可空白
   If IsEmptyText(textCP08) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入機關文號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP08.SetFocus
      GoTo EXITSUB
   End If
   
   'Add By Sindy 2015/7/29 有優先權證明期限時,不控管一定要輸入下一程序,來函期限,本所期限,法定期限
   If IsEmptyText(textAddDate) = True Then
   '2015/7/29 END
      ' 下一程序不可空白
      If IsEmptyText(textCF15) = True Then
         strTit = "資料檢核"
         strMsg = "請輸入下一程序"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2015/7/29 有優先權證明期限時,不控管一定要輸入下一程序,來函期限,本所期限,法定期限
   If IsEmptyText(textAddDate) = True Then
   '2015/7/29 END
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
      End If
   End If
   
   'Add By Sindy 2015/7/29 有優先權證明期限時,不控管一定要輸入下一程序,來函期限,本所期限,法定期限
   If IsEmptyText(textAddDate) = True Then
   '2015/7/29 END
      ' 本所期限不可空白
      If IsEmptyText(textCP06) = True Then
         strTit = "資料檢核"
         strMsg = "請輸入本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
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
   'Add By Sindy 2015/7/29 有優先權證明期限時,不控管一定要輸入下一程序,來函期限,本所期限,法定期限
   If IsEmptyText(textAddDate) = True Then
   '2015/7/29 END
      ' 法定期限不可空白
      If IsEmptyText(textCP07) = True Then
         strTit = "資料檢核"
         strMsg = "請輸入法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 本所期限的日期不可超過法定期限
   If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
      If Val(textCP06) > Val(textCP07) Then
         strTit = "資料檢核"
         strMsg = "本所期限的日期不可超過法定期限的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2015/7/29 有優先權證明期限時,檢查它的本所期限與承辦期限的狀況
   If IsEmptyText(textAddDate) = False Then
      '若優先權證明期限及承辦期限皆有輸入時, 承辦期限不可大於優先權證明期限的本所期限
      If Len(strTextAddDate_NP08) > 0 And Len(Me.textCP48.Text) > 0 Then
         If Val(strTextAddDate_NP08) < Val(Me.textCP48.Text) Then
            MsgBox "承辦期限不得大於優先權證明期限的本所期限（" & strTextAddDate_NP08 & "）!", vbExclamation + vbOKOnly
            textCP48.SetFocus
            GoTo EXITSUB
         End If
      End If
   Else
   '2015/7/29 END
      'Add By Cheng 2002/05/07
      '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
      If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
         If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
            MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
            textCP48.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   'add by nickc 2006/06/02
Dim Cancel As Boolean
   Cancel = False
   textCP07_Validate Cancel
   If Cancel = True Then GoTo EXITSUB
   textCF15_Validate Cancel
   If Cancel = True Then GoTo EXITSUB
   'Add by Amy 2025/01/17 避免輸完直接按Enter鍵不會檢查
   textCP37_1_Validate Cancel
   If Cancel = True Then GoTo EXITSUB
   textCP40_Validate Cancel
   If Cancel = True Then GoTo EXITSUB
   textCP41_Validate Cancel
   If Cancel = True Then GoTo EXITSUB
   textCP42_Validate Cancel
   If Cancel = True Then GoTo EXITSUB
   'end 2025/01/17
   '2011/6/15 ADD BY SONIA 自VALIDATE移過來並調整
   ' 申請國家為台灣時需檢查來函記錄檔
   If m_TM10 < "010" Then
      '本所期限
      strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05, "MR16")
      If IsEmptyText(strDate) = False Then
         If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
            strTit = "資料檢核"
            strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
            If nResponse = vbCancel Then
               Cancel = True
               textCP06_GotFocus
               GoTo EXITSUB
            End If
         End If
      Else
         '2011/6/15 MODIFY BY SONIA
'         strTit = "資料檢核"
'         strMsg = "來函記錄中無該筆記錄"
'         nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'         If nResponse = vbCancel Then
'            Cancel = True
'            textCP06_GotFocus
'            GoTo EXITSUB
'         End If
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
      '法定期限
      strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05, "MR17")
      If IsEmptyText(strDate) = False Then
         If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
            strTit = "資料檢核"
            strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
            If nResponse = vbCancel Then
               Cancel = True
               textCP07_GotFocus
               GoTo EXITSUB
            End If
         End If
      Else
         If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then  '2011/6/15 ADD BY SONIA
            'modify by sonia 2018/2/8 電子公文都不檢查來函記錄檔
            'If m_DocNo = "" Or textCP07 <> "" Then 'Added by Morgan 2017/5/4 電子公文
            If m_DocNo = "" And textCP07 <> "" Then 'Added by Morgan 2017/5/4 電子公文
               strTit = "資料檢核"
               strMsg = "來函記錄中無該筆記錄"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  Cancel = True
                  textCP07_GotFocus
                  GoTo EXITSUB
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
   '2011/6/15 END
   
   'Add By Sindy 2012/7/9 以防修改期限天數或月數,重新計算期限
   If Me.Text10.Enabled = True Then
      Cancel = False
      Text10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.Text11.Enabled = True Then
      Cancel = False
      Text11_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2012/7/9 End
   
    'Added by Lydia 2021/09/13 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         GoTo EXITSUB
    End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
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

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
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

'Add By Sindy 2012/4/17
Private Sub Option1_Click(Index As Integer)
   If Me.Option4(0).Value Then
      Text10_Validate False
   ElseIf Me.Option4(1).Value Then
      Text11_Validate False
   ElseIf Me.Option4(2).Value Then
      Text12_Validate False
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
               textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
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
   Dim strFromDate As String '期限起算日
   
   'Add By Sindy 2012/8/30
   If Option4(0).Value = False And Option4(1).Value = False Then Exit Sub
   '2012/8/30 End
   
   'strFromDate = DBDATE(textCP05)
   strFromDate = DBDATE(frm03020403_01.textCP05)
   
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
         textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      End If
   End If
End Sub

'讀取來函期限
Private Function ChgType() As Boolean
Dim strTempName As String, bolTmp As Boolean
Dim i As Integer
Dim strFromDate As String '期限起算日
   
   'strFromDate = DBDATE(textCP05)
   strFromDate = DBDATE(frm03020403_01.textCP05)
   
   ChgType = False
   If m_TM10 = 台灣國家代號 Then
      bolTmp = False
   Else
      bolTmp = True
   End If
   
   ' 案件性質為審查報告或核駁前先行通知
   strRvType = "1002"
   Select Case frm03020403_03.GetSelectResult
      Case "1": strRvType = "1201"
      Case "2": strRvType = "1202"
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
                     textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                  End If
               End If
            End With
         End If
      End If
      ChgType = True
   End If
End Function

