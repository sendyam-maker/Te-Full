VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040120 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人目標資料檔"
   ClientHeight    =   5820
   ClientLeft      =   570
   ClientTop       =   855
   ClientWidth     =   7515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7515
   Begin VB.TextBox textPE01 
      Height          =   270
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox textPE03_1 
      Height          =   270
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   2
      Top             =   948
      Width           =   492
   End
   Begin VB.TextBox textPE04 
      Height          =   270
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1296
      Width           =   855
   End
   Begin VB.TextBox textPE05 
      Height          =   270
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1644
      Width           =   855
   End
   Begin VB.TextBox textPE07 
      Height          =   270
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1992
      Width           =   855
   End
   Begin VB.TextBox textPE09 
      Height          =   270
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   9
      Top             =   2340
      Width           =   855
   End
   Begin VB.TextBox textPE11 
      Height          =   270
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   11
      Top             =   2688
      Width           =   855
   End
   Begin VB.TextBox textPE12 
      Height          =   270
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   12
      Top             =   3036
      Width           =   855
   End
   Begin VB.TextBox textPE14 
      Height          =   270
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   14
      Top             =   3384
      Width           =   855
   End
   Begin VB.TextBox textPE16 
      Height          =   270
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   16
      Top             =   3732
      Width           =   855
   End
   Begin VB.TextBox textPE18 
      Height          =   270
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   18
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox textPE29 
      Height          =   264
      Index           =   0
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   56
      Top             =   5436
      Width           =   492
   End
   Begin VB.TextBox textPE26 
      Height          =   264
      Index           =   0
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   44
      Top             =   5100
      Width           =   492
   End
   Begin VB.TextBox textPE23 
      Height          =   264
      Index           =   0
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   32
      Top             =   4764
      Width           =   492
   End
   Begin VB.TextBox textPE20 
      Height          =   264
      Index           =   0
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   20
      Top             =   4428
      Width           =   492
   End
   Begin VB.TextBox textPE24 
      Height          =   264
      Index           =   0
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   36
      Top             =   4764
      Width           =   492
   End
   Begin VB.TextBox textPE27 
      Height          =   264
      Index           =   0
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   48
      Top             =   5100
      Width           =   492
   End
   Begin VB.TextBox textPE28 
      Height          =   264
      Index           =   0
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   52
      Top             =   5100
      Width           =   492
   End
   Begin VB.TextBox textPE20 
      Height          =   264
      Index           =   1
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   21
      Top             =   4428
      Width           =   732
   End
   Begin VB.TextBox textPE20 
      Height          =   264
      Index           =   2
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   22
      Top             =   4428
      Width           =   252
   End
   Begin VB.TextBox textPE20 
      Height          =   264
      Index           =   3
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   23
      Top             =   4428
      Width           =   372
   End
   Begin VB.TextBox textPE25 
      Height          =   264
      Index           =   0
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   40
      Top             =   4764
      Width           =   492
   End
   Begin VB.TextBox textPE21 
      Height          =   264
      Index           =   0
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   24
      Top             =   4428
      Width           =   492
   End
   Begin VB.TextBox textPE21 
      Height          =   264
      Index           =   1
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   25
      Top             =   4428
      Width           =   732
   End
   Begin VB.TextBox textPE21 
      Height          =   264
      Index           =   2
      Left            =   4800
      MaxLength       =   1
      TabIndex        =   26
      Top             =   4428
      Width           =   252
   End
   Begin VB.TextBox textPE21 
      Height          =   264
      Index           =   3
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   27
      Top             =   4428
      Width           =   372
   End
   Begin VB.TextBox textPE22 
      Height          =   264
      Index           =   0
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   28
      Top             =   4428
      Width           =   492
   End
   Begin VB.TextBox textPE22 
      Height          =   264
      Index           =   1
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   29
      Top             =   4428
      Width           =   732
   End
   Begin VB.TextBox textPE22 
      Height          =   264
      Index           =   2
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   30
      Top             =   4428
      Width           =   252
   End
   Begin VB.TextBox textPE22 
      Height          =   264
      Index           =   3
      Left            =   6960
      MaxLength       =   2
      TabIndex        =   31
      Top             =   4428
      Width           =   372
   End
   Begin VB.TextBox textPE23 
      Height          =   264
      Index           =   1
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   33
      Top             =   4764
      Width           =   732
   End
   Begin VB.TextBox textPE23 
      Height          =   264
      Index           =   2
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   34
      Top             =   4764
      Width           =   252
   End
   Begin VB.TextBox textPE23 
      Height          =   264
      Index           =   3
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   35
      Top             =   4764
      Width           =   372
   End
   Begin VB.TextBox textPE24 
      Height          =   264
      Index           =   1
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   37
      Top             =   4764
      Width           =   732
   End
   Begin VB.TextBox textPE24 
      Height          =   264
      Index           =   2
      Left            =   4800
      MaxLength       =   1
      TabIndex        =   38
      Top             =   4764
      Width           =   252
   End
   Begin VB.TextBox textPE24 
      Height          =   264
      Index           =   3
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   39
      Top             =   4764
      Width           =   372
   End
   Begin VB.TextBox textPE25 
      Height          =   264
      Index           =   2
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   42
      Top             =   4764
      Width           =   252
   End
   Begin VB.TextBox textPE25 
      Height          =   264
      Index           =   3
      Left            =   6960
      MaxLength       =   2
      TabIndex        =   43
      Top             =   4764
      Width           =   372
   End
   Begin VB.TextBox textPE26 
      Height          =   264
      Index           =   1
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   45
      Top             =   5100
      Width           =   732
   End
   Begin VB.TextBox textPE26 
      Height          =   264
      Index           =   2
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   46
      Top             =   5100
      Width           =   252
   End
   Begin VB.TextBox textPE26 
      Height          =   264
      Index           =   3
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   47
      Top             =   5100
      Width           =   372
   End
   Begin VB.TextBox textPE27 
      Height          =   264
      Index           =   1
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   49
      Top             =   5100
      Width           =   732
   End
   Begin VB.TextBox textPE27 
      Height          =   264
      Index           =   2
      Left            =   4800
      MaxLength       =   1
      TabIndex        =   50
      Top             =   5100
      Width           =   252
   End
   Begin VB.TextBox textPE27 
      Height          =   264
      Index           =   3
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   51
      Top             =   5100
      Width           =   372
   End
   Begin VB.TextBox textPE28 
      Height          =   264
      Index           =   1
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   53
      Top             =   5100
      Width           =   732
   End
   Begin VB.TextBox textPE28 
      Height          =   264
      Index           =   2
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   54
      Top             =   5100
      Width           =   252
   End
   Begin VB.TextBox textPE28 
      Height          =   264
      Index           =   3
      Left            =   6960
      MaxLength       =   2
      TabIndex        =   55
      Top             =   5100
      Width           =   372
   End
   Begin VB.TextBox textPE29 
      Height          =   264
      Index           =   1
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   57
      Top             =   5436
      Width           =   732
   End
   Begin VB.TextBox textPE29 
      Height          =   264
      Index           =   2
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   58
      Top             =   5436
      Width           =   252
   End
   Begin VB.TextBox textPE29 
      Height          =   264
      Index           =   3
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   59
      Top             =   5436
      Width           =   372
   End
   Begin VB.TextBox textPE25 
      Height          =   264
      Index           =   1
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   41
      Top             =   4764
      Width           =   732
   End
   Begin VB.TextBox textPE13 
      Height          =   270
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   13
      Top             =   3036
      Width           =   855
   End
   Begin VB.TextBox textPE03_2 
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   3
      Top             =   948
      Width           =   372
   End
   Begin VB.TextBox textPE19 
      Height          =   270
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   19
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox textPE17 
      Height          =   270
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   17
      Top             =   3732
      Width           =   855
   End
   Begin VB.TextBox textPE15 
      Height          =   270
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   15
      Top             =   3384
      Width           =   855
   End
   Begin VB.TextBox textPE10 
      Height          =   270
      Left            =   5520
      MaxLength       =   5
      TabIndex        =   10
      Top             =   2340
      Width           =   855
   End
   Begin VB.TextBox textPE08 
      Height          =   270
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1992
      Width           =   855
   End
   Begin VB.TextBox textPE06 
      Height          =   270
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1644
      Width           =   855
   End
   Begin VB.TextBox textPE02 
      Height          =   270
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   1
      Top             =   600
      Width           =   492
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
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
            Picture         =   "frm12040120.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040120.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040120.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040120.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040120.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040120.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040120.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040120.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040120.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040120.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040120.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   7515
      _ExtentX        =   13256
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
   Begin MSForms.TextBox textPE01_2 
      Height          =   285
      Left            =   2580
      TabIndex        =   83
      Top             =   600
      Width           =   1215
      VariousPropertyBits=   671105055
      Size            =   "2143;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      Caption         =   "查名失誤案號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   82
      Top             =   4428
      Width           =   1212
   End
   Begin VB.Label Label7 
      Caption         =   "月"
      Height          =   252
      Left            =   3120
      TabIndex        =   81
      Top             =   948
      Width           =   252
   End
   Begin VB.Label Label6 
      Caption         =   "年"
      Height          =   252
      Left            =   2280
      TabIndex        =   80
      Top             =   948
      Width           =   252
   End
   Begin VB.Label Label5 
      Caption         =   "繪圖點數 : "
      Height          =   252
      Left            =   120
      TabIndex        =   79
      Top             =   2688
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "其他件數 :"
      Height          =   252
      Left            =   120
      TabIndex        =   78
      Top             =   1992
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "專業件數 :"
      Height          =   252
      Left            =   120
      TabIndex        =   77
      Top             =   1644
      Width           =   972
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標勝訴率 2 : "
      Height          =   180
      Index           =   15
      Left            =   3960
      TabIndex        =   75
      Top             =   4080
      Width           =   1176
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標勝訴率 1 : "
      Height          =   180
      Index           =   14
      Left            =   120
      TabIndex        =   74
      Top             =   4080
      Width           =   1176
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標預估準確率 : "
      Height          =   180
      Index           =   13
      Left            =   3960
      TabIndex        =   73
      Top             =   3732
      Width           =   1476
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標未輸入筆數 : "
      Height          =   180
      Index           =   12
      Left            =   120
      TabIndex        =   72
      Top             =   3732
      Width           =   1404
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標過期筆數 : "
      Height          =   180
      Index           =   11
      Left            =   3960
      TabIndex        =   71
      Top             =   3384
      Width           =   1224
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標圖形筆數 : "
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   70
      Top             =   3384
      Width           =   1224
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標英文筆數 : "
      Height          =   180
      Index           =   9
      Left            =   3960
      TabIndex        =   69
      Top             =   3036
      Width           =   1332
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標中文筆數 : "
      Height          =   180
      Index           =   8
      Left            =   120
      TabIndex        =   68
      Top             =   3036
      Width           =   1224
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "繪圖張數 :"
      Height          =   180
      Index           =   7
      Left            =   3960
      TabIndex        =   67
      Top             =   2340
      Width           =   936
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "繪圖件數 :"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   66
      Top             =   2340
      Width           =   816
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "其他點數 :"
      Height          =   180
      Index           =   5
      Left            =   3960
      TabIndex        =   65
      Top             =   1992
      Width           =   816
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專業點數 : "
      Height          =   180
      Index           =   4
      Left            =   3960
      TabIndex        =   64
      Top             =   1644
      Width           =   864
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務點數 :"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   63
      Top             =   1296
      Width           =   1056
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "目標年月 :"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   62
      Top             =   948
      Width           =   936
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別 :"
      Height          =   180
      Index           =   1
      Left            =   3960
      TabIndex        =   61
      Top             =   648
      Width           =   936
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號 :"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   60
      Top             =   648
      Width           =   816
   End
End
Attribute VB_Name = "frm12040120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/11/23 改成Form2.0 ; textPE01_2
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Const MAX_FIELD = 29

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

' 目前正在作用的資料項目索引
Dim m_FirstPE(3) As String
Dim m_CurrPE(3) As String
Dim m_LastPE(3) As String

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

'add by nickc 2006/11/10 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case 13:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
   End Select
End Sub

' Load Form
Private Sub Form_Load()
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040120", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040120", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040120", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040120", strFind, False)
   
   textPE01_2.BackColor = &H8000000F
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
   InitialField
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ClearFieldList
   'Add By Cheng 2002/07/18
   Set frm12040120 = Nothing
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE) AND " & _
                  "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE)) AND " & _
                  "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE) AND PE02 = (SELECT MIN(PE02) FROM PERFORMANCE WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE)))"
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_FirstPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_FirstPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_FirstPE(2) = rsTmp.Fields("PE03")
   End If
   rsTmp.Close

   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE) AND " & _
                  "PE02 = (SELECT MAX(PE02) FROM PERFORMANCE WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE)) AND " & _
                  "PE03 = (SELECT MAX(PE03) FROM PERFORMANCE WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE) AND PE02 = (SELECT MAX(PE02) FROM PERFORMANCE WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE)))"
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_LastPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_LastPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_LastPE(2) = rsTmp.Fields("PE03")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "PE" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
End Sub

Private Sub ClearFieldList()
   Erase m_FieldList
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
   SetFieldNewData "PE01", textPE01
   SetFieldNewData "PE02", textPE02
   SetFieldNewData "PE03", (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
   SetFieldNewData "PE04", textPE04
   SetFieldNewData "PE05", textPE05
   SetFieldNewData "PE06", textPE06
   SetFieldNewData "PE07", textPE07
   SetFieldNewData "PE08", textPE08
   SetFieldNewData "PE09", textPE09
   SetFieldNewData "PE10", textPE10
   SetFieldNewData "PE11", textPE11
   SetFieldNewData "PE12", textPE12
   SetFieldNewData "PE13", textPE13
   SetFieldNewData "PE14", textPE14
   SetFieldNewData "PE15", textPE15
   SetFieldNewData "PE16", textPE16
   SetFieldNewData "PE17", textPE17
   SetFieldNewData "PE18", textPE18
   SetFieldNewData "PE19", textPE19
   If IsEmptyText(textPE20(0)) = False Then
      SetFieldNewData "PE20", textPE20(0) & textPE20(1) & textPE20(2) & String(1 - Len(textPE20(2)), "0") & textPE20(3) & String(2 - Len(textPE20(3)), "0")
   Else
      SetFieldNewData "PE20", Empty
   End If
   If IsEmptyText(textPE21(0)) = False Then
      SetFieldNewData "PE21", textPE21(0) & textPE21(1) & textPE21(2) & String(1 - Len(textPE21(2)), "0") & textPE21(3) & String(2 - Len(textPE21(3)), "0")
   Else
      SetFieldNewData "PE21", Empty
   End If
   If IsEmptyText(textPE22(0)) = False Then
      SetFieldNewData "PE22", textPE22(0) & textPE22(1) & textPE22(2) & String(1 - Len(textPE22(2)), "0") & textPE22(3) & String(2 - Len(textPE22(3)), "0")
   Else
      SetFieldNewData "PE22", Empty
   End If
   If IsEmptyText(textPE23(0)) = False Then
      SetFieldNewData "PE23", textPE23(0) & textPE23(1) & textPE23(2) & String(1 - Len(textPE23(2)), "0") & textPE23(3) & String(2 - Len(textPE23(3)), "0")
   Else
      SetFieldNewData "PE23", Empty
   End If
   If IsEmptyText(textPE24(0)) = False Then
      SetFieldNewData "PE24", textPE24(0) & textPE24(1) & textPE24(2) & String(1 - Len(textPE24(2)), "0") & textPE24(3) & String(2 - Len(textPE24(3)), "0")
   Else
      SetFieldNewData "PE24", Empty
   End If
   If IsEmptyText(textPE25(0)) = False Then
      SetFieldNewData "PE25", textPE25(0) & textPE25(1) & textPE25(2) & String(1 - Len(textPE25(2)), "0") & textPE25(3) & String(2 - Len(textPE25(3)), "0")
   Else
      SetFieldNewData "PE25", Empty
   End If
   If IsEmptyText(textPE26(0)) = False Then
      SetFieldNewData "PE26", textPE26(0) & textPE26(1) & textPE26(2) & String(1 - Len(textPE26(2)), "0") & textPE26(3) & String(2 - Len(textPE26(3)), "0")
   Else
      SetFieldNewData "PE26", Empty
   End If
   If IsEmptyText(textPE27(0)) = False Then
      SetFieldNewData "PE27", textPE27(0) & textPE27(1) & textPE27(2) & String(1 - Len(textPE27(2)), "0") & textPE27(3) & String(2 - Len(textPE27(3)), "0")
   Else
      SetFieldNewData "PE27", Empty
   End If
   If IsEmptyText(textPE28(0)) = False Then
      SetFieldNewData "PE28", textPE28(0) & textPE28(1) & textPE28(2) & String(1 - Len(textPE28(2)), "0") & textPE28(3) & String(2 - Len(textPE28(3)), "0")
   Else
      SetFieldNewData "PE28", Empty
   End If
   If IsEmptyText(textPE29(0)) = False Then
      SetFieldNewData "PE29", textPE29(0) & textPE29(1) & textPE29(2) & String(1 - Len(textPE29(2)), "0") & textPE29(3) & String(2 - Len(textPE29(3)), "0")
   Else
      SetFieldNewData "PE29", Empty
   End If
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
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
   textPE01 = Empty
   textPE01_2 = Empty
   textPE02 = Empty
   textPE03_1 = Empty
   textPE03_2 = Empty
   textPE04 = Empty
   textPE05 = Empty
   textPE06 = Empty
   textPE07 = Empty
   textPE08 = Empty
   textPE09 = Empty
   textPE10 = Empty
   textPE11 = Empty
   textPE12 = Empty
   textPE13 = Empty
   textPE14 = Empty
   textPE15 = Empty
   textPE16 = Empty
   textPE17 = Empty
   textPE18 = Empty
   textPE19 = Empty
   textPE20(0) = Empty: textPE20(1) = Empty: textPE20(2) = Empty: textPE20(3) = Empty
   textPE21(0) = Empty: textPE21(1) = Empty: textPE21(2) = Empty: textPE21(3) = Empty
   textPE22(0) = Empty: textPE22(1) = Empty: textPE22(2) = Empty: textPE22(3) = Empty
   textPE23(0) = Empty: textPE23(1) = Empty: textPE23(2) = Empty: textPE23(3) = Empty
   textPE24(0) = Empty: textPE24(1) = Empty: textPE24(2) = Empty: textPE24(3) = Empty
   textPE25(0) = Empty: textPE25(1) = Empty: textPE25(2) = Empty: textPE25(3) = Empty
   textPE26(0) = Empty: textPE26(1) = Empty: textPE26(2) = Empty: textPE26(3) = Empty
   textPE27(0) = Empty: textPE27(1) = Empty: textPE27(2) = Empty: textPE27(3) = Empty
   textPE28(0) = Empty: textPE28(1) = Empty: textPE28(2) = Empty: textPE28(3) = Empty
   textPE29(0) = Empty: textPE29(1) = Empty: textPE29(2) = Empty: textPE29(3) = Empty
   
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textPE01.Locked = bEnable
   textPE02.Locked = bEnable
   textPE03_1.Locked = bEnable
   textPE03_2.Locked = bEnable
   textPE04.Locked = bEnable
   textPE05.Locked = bEnable
   textPE06.Locked = bEnable
   textPE07.Locked = bEnable
   textPE08.Locked = bEnable
   textPE09.Locked = bEnable
   textPE10.Locked = bEnable
   textPE11.Locked = bEnable
   textPE12.Locked = bEnable
   textPE13.Locked = bEnable
   textPE14.Locked = bEnable
   textPE15.Locked = bEnable
   textPE16.Locked = bEnable
   textPE17.Locked = bEnable
   textPE18.Locked = bEnable
   textPE19.Locked = bEnable
   textPE20(0).Locked = bEnable: textPE20(1).Locked = bEnable: textPE20(2).Locked = bEnable: textPE20(3).Locked = bEnable
   textPE21(0).Locked = bEnable: textPE21(1).Locked = bEnable: textPE21(2).Locked = bEnable: textPE21(3).Locked = bEnable
   textPE22(0).Locked = bEnable: textPE22(1).Locked = bEnable: textPE22(2).Locked = bEnable: textPE22(3).Locked = bEnable
   textPE23(0).Locked = bEnable: textPE23(1).Locked = bEnable: textPE23(2).Locked = bEnable: textPE23(3).Locked = bEnable
   textPE24(0).Locked = bEnable: textPE24(1).Locked = bEnable: textPE24(2).Locked = bEnable: textPE24(3).Locked = bEnable
   textPE25(0).Locked = bEnable: textPE25(1).Locked = bEnable: textPE25(2).Locked = bEnable: textPE25(3).Locked = bEnable
   textPE26(0).Locked = bEnable: textPE26(1).Locked = bEnable: textPE26(2).Locked = bEnable: textPE26(3).Locked = bEnable
   textPE27(0).Locked = bEnable: textPE27(1).Locked = bEnable: textPE27(2).Locked = bEnable: textPE27(3).Locked = bEnable
   textPE28(0).Locked = bEnable: textPE28(1).Locked = bEnable: textPE28(2).Locked = bEnable: textPE28(3).Locked = bEnable
   textPE29(0).Locked = bEnable: textPE29(1).Locked = bEnable: textPE29(2).Locked = bEnable: textPE29(3).Locked = bEnable
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textPE01.Locked = bEnable
   textPE02.Locked = bEnable
   textPE03_1.Locked = bEnable
   textPE03_2.Locked = bEnable
End Sub

Private Function ConvertAC(ByVal strData As String, ByRef strKey1 As String, ByRef StrKey2 As String, ByRef strKey3 As String, ByRef strKey4 As String) As Boolean
   ConvertAC = True
   Select Case Len(strData)
      Case 10:
         strKey1 = Mid(strData, 1, 1)
         StrKey2 = Mid(strData, 2, 6)
         strKey3 = Mid(strData, 8, 1)
         strKey4 = Mid(strData, 9, 2)
      Case 11:
         strKey1 = Mid(strData, 1, 2)
         StrKey2 = Mid(strData, 3, 6)
         strKey3 = Mid(strData, 9, 1)
         strKey4 = Mid(strData, 10, 2)
      Case 12:
         strKey1 = Mid(strData, 1, 3)
         StrKey2 = Mid(strData, 4, 6)
         strKey3 = Mid(strData, 10, 1)
         strKey4 = Mid(strData, 11, 2)
      Case Else:
         ConvertAC = False
         strKey1 = Empty
         StrKey2 = Empty
         strKey3 = Empty
         strKey4 = Empty
   End Select
End Function

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim bCnv As Boolean
   Dim strKey1 As String
   Dim StrKey2 As String
   Dim strKey3 As String
   Dim strKey4 As String
   
   strSql = "SELECT * FROM PERFORMANCE " & _
            "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                  "PE02 = '" & m_CurrPE(1) & "' AND " & _
                  "PE03 = '" & m_CurrPE(2) & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If Not IsNull(rsTmp.Fields("PE01")) Then: textPE01 = rsTmp.Fields("PE01")
      If Not IsNull(rsTmp.Fields("PE02")) Then: textPE02 = rsTmp.Fields("PE02")
      textPE01_2 = GetStaffName(textPE01, True) 'Added by Lydia 2021/11/23
      If Not IsNull(rsTmp.Fields("PE03")) Then
         '2005/4/25 MODIFY BY SONIA
         'textPE03_1 = Val(Mid(rsTmp.Fields("PE03"), 1, 4)) - 1911
         'textPE03_2 = Mid(rsTmp.Fields("PE03"), 5, 2)
         '2010/9/15 MODIFY BY SONIA
         'If rsTmp.Fields("PE03") < 1911 Then
         If Val(rsTmp.Fields("PE03")) < 1911 Then
            textPE03_2 = rsTmp.Fields("PE03")
         Else
            textPE03_1 = Val(Mid(rsTmp.Fields("PE03"), 1, 4)) - 1911
            textPE03_2 = Mid(rsTmp.Fields("PE03"), 5, 2)
         End If
         '2005/4/25 END
      End If
      If Not IsNull(rsTmp.Fields("PE04")) Then: textPE04 = rsTmp.Fields("PE04")
      If Not IsNull(rsTmp.Fields("PE05")) Then: textPE05 = rsTmp.Fields("PE05")
      If Not IsNull(rsTmp.Fields("PE06")) Then: textPE06 = rsTmp.Fields("PE06")
      If Not IsNull(rsTmp.Fields("PE07")) Then: textPE07 = rsTmp.Fields("PE07")
      If Not IsNull(rsTmp.Fields("PE08")) Then: textPE08 = rsTmp.Fields("PE08")
      If Not IsNull(rsTmp.Fields("PE09")) Then: textPE09 = rsTmp.Fields("PE09")
      If Not IsNull(rsTmp.Fields("PE10")) Then: textPE10 = rsTmp.Fields("PE10")
      If Not IsNull(rsTmp.Fields("PE11")) Then: textPE11 = rsTmp.Fields("PE11")
      If Not IsNull(rsTmp.Fields("PE12")) Then: textPE12 = rsTmp.Fields("PE12")
      If Not IsNull(rsTmp.Fields("PE13")) Then: textPE13 = rsTmp.Fields("PE13")
      If Not IsNull(rsTmp.Fields("PE14")) Then: textPE14 = rsTmp.Fields("PE14")
      If Not IsNull(rsTmp.Fields("PE15")) Then: textPE15 = rsTmp.Fields("PE15")
      If Not IsNull(rsTmp.Fields("PE16")) Then: textPE16 = rsTmp.Fields("PE16")
      If Not IsNull(rsTmp.Fields("PE17")) Then: textPE17 = rsTmp.Fields("PE17")
      If Not IsNull(rsTmp.Fields("PE18")) Then: textPE18 = rsTmp.Fields("PE18")
      If Not IsNull(rsTmp.Fields("PE19")) Then: textPE19 = rsTmp.Fields("PE19")
      ' 本所案號
      strKey1 = Empty: StrKey2 = Empty: strKey3 = Empty: strKey4 = Empty
      If Not IsNull(rsTmp.Fields("PE20")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE20"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE20(0) = strKey1: textPE20(1) = StrKey2: textPE20(2) = strKey3: textPE20(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE21")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE21"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE21(0) = strKey1: textPE21(1) = StrKey2: textPE21(2) = strKey3: textPE21(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE22")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE22"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE22(0) = strKey1: textPE22(1) = StrKey2: textPE22(2) = strKey3: textPE22(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE23")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE23"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE23(0) = strKey1: textPE23(1) = StrKey2: textPE23(2) = strKey3: textPE23(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE24")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE24"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE24(0) = strKey1: textPE24(1) = StrKey2: textPE24(2) = strKey3: textPE24(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE25")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE25"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE25(0) = strKey1: textPE25(1) = StrKey2: textPE25(2) = strKey3: textPE25(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE26")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE26"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE26(0) = strKey1: textPE26(1) = StrKey2: textPE26(2) = strKey3: textPE26(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE27")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE27"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE27(0) = strKey1: textPE27(1) = StrKey2: textPE27(2) = strKey3: textPE27(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE28")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE28"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE28(0) = strKey1: textPE28(1) = StrKey2: textPE28(2) = strKey3: textPE28(3) = strKey4
      End If
      If Not IsNull(rsTmp.Fields("PE29")) Then
         bCnv = ConvertAC(rsTmp.Fields("PE29"), strKey1, StrKey2, strKey3, strKey4)
         If bCnv = True Then: textPE29(0) = strKey1: textPE29(1) = StrKey2: textPE29(2) = strKey3: textPE29(3) = strKey4
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
EXITSUB:
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strPE01 As String, ByVal strPE02 As String, ByVal strPE03 As String)
   Dim strTemp As String
   Dim nIndex As Integer
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strPE01, strPE02, strPE03) = True Then
      m_CurrPE(0) = strPE01
      m_CurrPE(1) = strPE02
      m_CurrPE(2) = strPE03
   Else
      strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
               "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                     "PE02 = '" & m_CurrPE(1) & "' AND " & _
                     "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                             "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                   "PE02 = '" & m_CurrPE(1) & "' AND " & _
                                   "PE03 > '" & m_CurrPE(2) & "' ) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
         If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
         If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
   
      strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
               "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                     "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                             "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                   "PE02 > '" & m_CurrPE(1) & "') AND " & _
                     "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                             "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                   "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                                           "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                                 "PE02 > '" & m_CurrPE(1) & "'))"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
         If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
         If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      'Else
      '   ShowLastRecord
      '   GoTo ExitSub
      End If
      rsTmp.Close
      
      strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
               "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                             "WHERE PE01 > '" & m_CurrPE(0) & "') AND " & _
                     "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                             "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                           "WHERE PE01 > '" & m_CurrPE(0) & "')) AND " & _
                     "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                             "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                           "WHERE PE01 > '" & m_CurrPE(0) & "') AND " & _
                                   "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                                           "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                                         "WHERE PE01 > '" & m_CurrPE(0) & "')))"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
         If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
         If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrPE(0) = m_FirstPE(0)
   m_CurrPE(1) = m_FirstPE(1)
   m_CurrPE(2) = m_FirstPE(2)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrPE(0) = m_FirstPE(0) And m_CurrPE(1) = m_FirstPE(1) And m_CurrPE(2) = m_FirstPE(2) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If

   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                  "PE02 = '" & m_CurrPE(1) & "' AND " & _
                  "PE03 = (SELECT MAX(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 = '" & m_CurrPE(1) & "' AND " & _
                                "PE03 < '" & m_CurrPE(2) & "')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                  "PE02 = (SELECT MAX(PE02) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 < '" & m_CurrPE(1) & "') AND " & _
                  "PE03 = (SELECT MAX(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 = (SELECT MAX(PE02) FROM PERFORMANCE " & _
                                        "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                              "PE02 < '" & m_CurrPE(1) & "'))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE " & _
                          "WHERE PE01 < '" & m_CurrPE(0) & "') AND " & _
                  "PE02 = (SELECT MAX(PE02) FROM PERFORMANCE " & _
                          "WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE " & _
                                        "WHERE PE01 < '" & m_CurrPE(0) & "')) AND " & _
                  "PE03 = (SELECT MAX(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE " & _
                                        "WHERE PE01 < '" & m_CurrPE(0) & "') AND " & _
                                "PE02 = (SELECT MAX(PE02) FROM PERFORMANCE " & _
                                        "WHERE PE01 = (SELECT MAX(PE01) FROM PERFORMANCE " & _
                                                      "WHERE PE01 < '" & m_CurrPE(0) & "')))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
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
   
   If m_CurrPE(0) = m_LastPE(0) And m_CurrPE(1) = m_LastPE(1) And m_CurrPE(2) = m_LastPE(2) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                  "PE02 = '" & m_CurrPE(1) & "' AND " & _
                  "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 = '" & m_CurrPE(1) & "' AND " & _
                                "PE03 > '" & m_CurrPE(2) & "' ) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                  "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 > '" & m_CurrPE(1) & "') AND " & _
                  "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                                        "WHERE PE01 = '" & m_CurrPE(0) & "' AND " & _
                                              "PE02 > '" & m_CurrPE(1) & "'))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT PE01,PE02,PE03 FROM PERFORMANCE " & _
            "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                          "WHERE PE01 > '" & m_CurrPE(0) & "') AND " & _
                  "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                          "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                        "WHERE PE01 > '" & m_CurrPE(0) & "')) AND " & _
                  "PE03 = (SELECT MIN(PE03) FROM PERFORMANCE " & _
                          "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                        "WHERE PE01 > '" & m_CurrPE(0) & "') AND " & _
                                "PE02 = (SELECT MIN(PE02) FROM PERFORMANCE " & _
                                        "WHERE PE01 = (SELECT MIN(PE01) FROM PERFORMANCE " & _
                                                      "WHERE PE01 > '" & m_CurrPE(0) & "')))"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PE01")) = False Then: m_CurrPE(0) = rsTmp.Fields("PE01")
      If IsNull(rsTmp.Fields("PE02")) = False Then: m_CurrPE(1) = rsTmp.Fields("PE02")
      If IsNull(rsTmp.Fields("PE03")) = False Then: m_CurrPE(2) = rsTmp.Fields("PE03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close

   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrPE(0) = m_LastPE(0)
   m_CurrPE(1) = m_LastPE(1)
   m_CurrPE(2) = m_LastPE(2)
   
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
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
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
'edit by nickc 2006/11/10
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
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
         Unload Me
   End Select
EXITSUB:
End Sub

Private Sub textPE01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 員工代號
Private Sub textPE01_Validate(Cancel As Boolean)
   Dim strPE01 As String
   Dim strPE02 As String
   Dim strPE03 As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   textPE01_2 = Empty
   If IsEmptyText(textPE01) = False Then
      textPE01_2 = GetStaffName(textPE01)
      If IsEmptyText(textPE01_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "員工代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE01_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 檢查Key是否存在
   If m_EditMode = 1 Then
      If IsEmptyText(textPE01) = False And IsEmptyText(textPE02) = False And IsEmptyText(textPE03_1) = False And IsEmptyText(textPE03_2) = False Then
         strPE01 = textPE01
         strPE02 = textPE02
         strPE03 = (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
         ' 檢查記錄是否已存在
         If IsRecordExist(strPE01, strPE02, strPE03) = True Then
            strTit = "資料檢核"
            strMsg = "該筆記錄已存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE01_GotFocus
            GoTo EXITSUB
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub textPE02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 系統類別
Private Sub textPE02_Validate(Cancel As Boolean)
   Dim strPE01 As String
   Dim strPE02 As String
   Dim strPE03 As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False

   If IsEmptyText(textPE02) = False Then
      If m_EditMode = 1 Or m_EditMode = 4 Then
         If IsCorrectSysKind(textPE02) = False Then
            '2005/4/25 modify by sonia
            'Cancel = True
            'strTit = "資料檢核"
            'strMsg = "系統類別不正確"
            'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            'textPE02_GotFocus
            'GoTo EXITSUB
            If textPE02 = "TOT" And Pub_StrUserSt03 = "M51" Then
            Else
               Cancel = True
               strTit = "資料檢核"
               strMsg = "系統類別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE02_GotFocus
               GoTo EXITSUB
            End If
            '2005/4/25 end
         End If
      End If
      
      ' 檢查Key是否存在
      If m_EditMode = 1 Then
         If IsEmptyText(textPE01) = False And IsEmptyText(textPE02) = False And IsEmptyText(textPE03_1) = False And IsEmptyText(textPE03_2) = False Then
            strPE01 = textPE01
            strPE02 = textPE02
            strPE03 = (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
            ' 檢查記錄是否已存在
            If IsRecordExist(strPE01, strPE02, strPE03) = True Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "該筆記錄已存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE02_GotFocus
               GoTo EXITSUB
            End If
         End If
      End If
   End If
EXITSUB:
End Sub

' 年
Private Sub textPE03_1_Validate(Cancel As Boolean)
   Dim strPE01 As String
   Dim strPE02 As String
   Dim strPE03 As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPE03_1) = False Then
      ' 檢查Key是否存在
      If m_EditMode = 1 Then
         If IsEmptyText(textPE01) = False And IsEmptyText(textPE02) = False And IsEmptyText(textPE03_1) = False And IsEmptyText(textPE03_2) = False Then
            strPE01 = textPE01
            strPE02 = textPE02
            strPE03 = (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
            ' 檢查記錄是否已存在
            If IsRecordExist(strPE01, strPE02, strPE03) = True Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "該筆記錄已存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE03_1_GotFocus
               GoTo EXITSUB
            End If
         End If
      End If
   End If
EXITSUB:
End Sub

' 月
Private Sub textPE03_2_Validate(Cancel As Boolean)
   Dim strPE01 As String
   Dim strPE02 As String
   Dim strPE03 As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False

   If IsEmptyText(textPE03_2) = False Then
      If IsNumeric(textPE03_2) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "目標年月中的月份不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE03_2_GotFocus
         GoTo EXITSUB
      End If
      If Val(textPE03_2) < 1 Or Val(textPE03_2) > 12 Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "目標年月中的月份不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE03_2_GotFocus
         GoTo EXITSUB
      End If
      
      ' 檢查Key是否存在
      If m_EditMode = 1 Then
         If IsEmptyText(textPE01) = False And IsEmptyText(textPE02) = False And IsEmptyText(textPE03_1) = False And IsEmptyText(textPE03_2) = False Then
            strPE01 = textPE01
            strPE02 = textPE02
            strPE03 = (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
            ' 檢查記錄是否已存在
            If IsRecordExist(strPE01, strPE02, strPE03) = True Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "該筆記錄已存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE03_2_GotFocus
               GoTo EXITSUB
            End If
         End If
      End If
   End If
EXITSUB:
End Sub

' 業務點數
Private Sub textPE04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE04) = False Then
      If IsNumeric(textPE04) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "業務點數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE04_GotFocus
      End If
   End If
End Sub

' 專業件數
Private Sub textPE05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE05) = False Then
      If IsNumeric(textPE05) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "專業件數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE05_GotFocus
      End If
   End If
End Sub

' 專業點數
Private Sub textPE06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE06) = False Then
      If IsNumeric(textPE06) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "專業點數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE06_GotFocus
      End If
   End If
End Sub

' 其它件數
Private Sub textPE07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE07) = False Then
      If IsNumeric(textPE07) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "其它件數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE07_GotFocus
      End If
   End If
End Sub

' 其它點數
Private Sub textPE08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE08) = False Then
      If IsNumeric(textPE08) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "其它點數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE08_GotFocus
      End If
   End If
End Sub

' 繪圖件數
Private Sub textPE09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE09) = False Then
      If IsNumeric(textPE09) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "繪圖件數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE09_GotFocus
      End If
   End If
End Sub

' 繪圖張數
Private Sub textPE10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE10) = False Then
      If IsNumeric(textPE10) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "繪圖張數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE10_GotFocus
      End If
   End If
End Sub

' 繪圖點數
Private Sub textPE11_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE11) = False Then
      If IsNumeric(textPE11) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "繪圖點數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE11_GotFocus
      End If
   End If
End Sub

' 商標中文筆數
Private Sub textPE12_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE12) = False Then
      If IsNumeric(textPE12) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "商標中文筆數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE12_GotFocus
      End If
   End If
End Sub

' 商標英文筆數
Private Sub textPE13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE13) = False Then
      If IsNumeric(textPE13) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "商標英文筆數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE13_GotFocus
      End If
   End If
End Sub

' 商標圖形筆數
Private Sub textPE14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE14) = False Then
      If IsNumeric(textPE14) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "商標圖形筆數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE14_GotFocus
      End If
   End If
End Sub

' 商標過期筆數
Private Sub textPE15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE15) = False Then
      If IsNumeric(textPE15) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "商標過期筆數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE15_GotFocus
      End If
   End If
End Sub

' 商標未輸入筆數
Private Sub textPE16_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE16) = False Then
      If IsNumeric(textPE16) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "商標未輸入筆數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE16_GotFocus
      End If
   End If
End Sub

' 商標預估準確率
Private Sub textPE17_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE17) = False Then
      If IsNumeric(textPE17) = False Then
         strTit = "資料檢核"
         strMsg = "商標預估準確率請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE17_GotFocus
      End If
   End If
End Sub

' 商標勝訴率 1
Private Sub textPE18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE18) = False Then
      If IsNumeric(textPE18) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "商標勝訴率1請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE18_GotFocus
      End If
   End If
End Sub

' 商標勝訴率 2
Private Sub textPE19_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE19) = False Then
      If IsNumeric(textPE19) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "商標勝訴率2請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPE19_GotFocus
      End If
   End If
End Sub

Private Sub textPE20_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE21_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE22_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE23_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE24_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE25_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE26_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE27_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE28_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPE29_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號
Private Sub textPE20_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE20(Index)) = False Then
      Select Case Index
         Case 1:
            textPE20(Index) = String(6 - Len(textPE20(Index)), "0") & textPE20(Index)
         Case 2:
            textPE20(Index) = textPE20(Index) & String(1 - Len(textPE20(Index)), "0")
         Case 3:
            textPE20(Index) = textPE20(Index) & String(2 - Len(textPE20(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE21_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE21(Index)) = False Then
      Select Case Index
         Case 1:
            textPE21(Index) = String(6 - Len(textPE21(Index)), "0") & textPE21(Index)
         Case 2:
            textPE21(Index) = textPE21(Index) & String(1 - Len(textPE21(Index)), "0")
         Case 3:
            textPE21(Index) = textPE21(Index) & String(2 - Len(textPE21(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE22_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE22(Index)) = False Then
      Select Case Index
         Case 1:
            textPE22(Index) = String(6 - Len(textPE22(Index)), "0") & textPE22(Index)
         Case 2:
            textPE22(Index) = textPE22(Index) & String(1 - Len(textPE22(Index)), "0")
         Case 3:
            textPE22(Index) = textPE22(Index) & String(2 - Len(textPE22(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE23_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE23(Index)) = False Then
      Select Case Index
         Case 1:
            textPE23(Index) = String(6 - Len(textPE23(Index)), "0") & textPE23(Index)
         Case 2:
            textPE23(Index) = textPE23(Index) & String(1 - Len(textPE23(Index)), "0")
         Case 3:
            textPE23(Index) = textPE23(Index) & String(2 - Len(textPE23(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE24_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE24(Index)) = False Then
      Select Case Index
         Case 1:
            textPE24(Index) = String(6 - Len(textPE24(Index)), "0") & textPE24(Index)
         Case 2:
            textPE24(Index) = textPE24(Index) & String(1 - Len(textPE24(Index)), "0")
         Case 3:
            textPE24(Index) = textPE24(Index) & String(2 - Len(textPE24(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE25_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE25(Index)) = False Then
      Select Case Index
         Case 1:
            textPE25(Index) = String(6 - Len(textPE25(Index)), "0") & textPE25(Index)
         Case 2:
            textPE25(Index) = textPE25(Index) & String(1 - Len(textPE25(Index)), "0")
         Case 3:
            textPE25(Index) = textPE25(Index) & String(2 - Len(textPE25(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE26_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE26(Index)) = False Then
      Select Case Index
         Case 1:
            textPE26(Index) = String(6 - Len(textPE26(Index)), "0") & textPE26(Index)
         Case 2:
            textPE26(Index) = textPE26(Index) & String(1 - Len(textPE26(Index)), "0")
         Case 3:
            textPE26(Index) = textPE26(Index) & String(2 - Len(textPE26(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE27_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE27(Index)) = False Then
      Select Case Index
         Case 1:
            textPE27(Index) = String(6 - Len(textPE27(Index)), "0") & textPE27(Index)
         Case 2:
            textPE27(Index) = textPE27(Index) & String(1 - Len(textPE27(Index)), "0")
         Case 3:
            textPE27(Index) = textPE27(Index) & String(2 - Len(textPE27(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE28_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE28(Index)) = False Then
      Select Case Index
         Case 1:
            textPE28(Index) = String(6 - Len(textPE28(Index)), "0") & textPE28(Index)
         Case 2:
            textPE28(Index) = textPE28(Index) & String(1 - Len(textPE28(Index)), "0")
         Case 3:
            textPE28(Index) = textPE28(Index) & String(2 - Len(textPE28(Index)), "0")
      End Select
   End If
End Sub

Private Sub textPE29_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPE29(Index)) = False Then
      Select Case Index
         Case 1:
            textPE29(Index) = String(6 - Len(textPE29(Index)), "0") & textPE29(Index)
         Case 2:
            textPE29(Index) = textPE29(Index) & String(1 - Len(textPE29(Index)), "0")
         Case 3:
            textPE29(Index) = textPE29(Index) & String(2 - Len(textPE29(Index)), "0")
      End Select
   End If
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
Private Function IsRecordExist(ByVal strPE01 As String, ByVal strPE02 As String, ByVal strPE03 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   IsRecordExist = False
   strSql = "SELECT * FROM PERFORMANCE " & _
            "WHERE PE01 = '" & strPE01 & "' AND " & _
                  "PE02 = '" & strPE02 & "' AND " & _
                  "PE03 = '" & strPE03 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
EXITSUB:
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
   Dim strPE01, strPE02, strPE03 As String
   
   strPE01 = textPE01
   strPE02 = textPE02
   strPE03 = (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strPE01, strPE02, strPE03) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO PERFORMANCE ("
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
   strSql = strSql & ")"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      QueryDB
      ShowCurrRecord strPE01, strPE02, strPE03
   End If
   
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
   Dim strPE01, strPE02, strPE03 As String
   
   strPE01 = textPE01
   strPE02 = textPE02
   strPE03 = (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
   
   strSql = "UPDATE PERFORMANCE SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_FieldList(nIndex).fiName & " = '" & m_FieldList(nIndex).fiNewData & "'"
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
                  "WHERE PE01 = '" & strPE01 & "' AND " & _
                     "PE02 = '" & strPE02 & "' AND " & _
                     "PE03 = '" & strPE03 & "'"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      QueryDB
      ShowCurrRecord strPE01, strPE02, strPE03
   End If

End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strPE01, strPE02, strPE03 As String
   
   strPE01 = textPE01
   strPE02 = textPE02
   strPE03 = (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
   
   strSql = "DELETE FROM Performance " & _
            "WHERE PE01 = '" & strPE01 & "' AND " & _
                  "PE02 = '" & strPE02 & "' AND " & _
                  "PE03 = '" & strPE03 & "'"
                  
   cnnConnection.Execute strSql
   
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strPE01 = m_LastPE(0) And strPE02 = m_LastPE(1) And strPE03 = m_LastPE(2)) Or (strPE01 = m_FirstPE(0) And strPE02 = m_FirstPE(1) And strPE03 = m_FirstPE(2)) Then
      RefreshRange
   End If

   ShowCurrRecord strPE01, strPE02, strPE03
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strPE01 As String
   Dim strPE02 As String
   Dim strPE03 As String
   
   strPE01 = textPE01
   strPE02 = textPE02
   strPE03 = (CStr(Val(textPE03_1) + 1911)) & (String(2 - Len(textPE03_2), "0") & textPE03_2)
   
   QueryRecord = False
   
   If IsRecordExist(strPE01, strPE02, strPE03) = True Then
      m_CurrPE(0) = strPE01
      m_CurrPE(1) = strPE02
      m_CurrPE(2) = strPE03
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
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            AddRecord
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            ModRecord
         Else
            GoTo EXITSUB
         End If
      Case 3:
         DelRecord
      Case 4:
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textPE01.SetFocus
      Case 2: textPE04.SetFocus
      Case 4: textPE01.SetFocus
   End Select
End Sub

' 檢查本所案號是否存在
Private Function IsDataExist(ByVal strKey1 As String, ByVal StrKey2 As String, ByVal strKey3 As String, ByVal strKey4 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   IsDataExist = False
   
   If IsEmptyText(strKey3) = True Then: strKey3 = "0"
   If IsEmptyText(strKey4) = True Then: strKey4 = "00"
   
   Select Case strKey1
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         strSql = "SELECT * FROM TRADEMARK WHERE TM01 = '" & strKey1 & "' AND TM02 = '" & StrKey2 & "' AND TM03 = '" & strKey3 & "' AND TM04 = '" & strKey4 & "' "
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         strSql = "SELECT * FROM PATENT WHERE PA01 = '" & strKey1 & "' AND PA02 = '" & StrKey2 & "' AND PA03 = '" & strKey3 & "' AND PA04 = '" & strKey4 & "' "
      ' 讀取法務基本檔
      Case "L", "CFL", "FCL":
         strSql = "SELECT * FROM LAWCASE WHERE LC01 = '" & strKey1 & "' AND LC02 = '" & StrKey2 & "' AND LC03 = '" & strKey3 & "' AND LC04 = '" & strKey4 & "' "
      ' 讀取顧問案件基本檔
      Case "LA":
         strSql = "SELECT * FROM HIRECASE WHERE HC01 = '" & strKey1 & "' AND HC02 = '" & StrKey2 & "' AND HC03 = '" & strKey3 & "' AND HC04 = '" & strKey4 & "' "
      ' 讀取服務業務基本檔
      Case Else:
         strSql = "SELECT * FROM SERVICEPRACTICE WHERE SP01 = '" & strKey1 & "' AND SP02 = '" & StrKey2 & "' AND SP03 = '" & strKey3 & "' AND SP04 = '" & strKey4 & "' "
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsDataExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 檢查輸入的失誤案號是否有重覆
Private Function CheckPENoExist()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strPENo() As String
   Dim nCount As Integer
   Dim nX As Integer
   Dim nY As Integer
   Dim strDup As String
   
   CheckPENoExist = False
   nCount = 0
   If IsEmptyText(textPE20(0)) = False And IsEmptyText(textPE20(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE20(0) & textPE20(1) & textPE20(2) & String(1 - Len(textPE20(2)), "0") & textPE20(3) & String(2 - Len(textPE20(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE21(0)) = False And IsEmptyText(textPE21(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE21(0) & textPE21(1) & textPE21(2) & String(1 - Len(textPE21(2)), "0") & textPE21(3) & String(2 - Len(textPE21(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE22(0)) = False And IsEmptyText(textPE22(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE22(0) & textPE22(1) & textPE22(2) & String(1 - Len(textPE22(2)), "0") & textPE22(3) & String(2 - Len(textPE22(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE23(0)) = False And IsEmptyText(textPE23(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE23(0) & textPE23(1) & textPE23(2) & String(1 - Len(textPE23(2)), "0") & textPE23(3) & String(2 - Len(textPE23(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE24(0)) = False And IsEmptyText(textPE24(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE24(0) & textPE24(1) & textPE24(2) & String(1 - Len(textPE24(2)), "0") & textPE24(3) & String(2 - Len(textPE24(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE25(0)) = False And IsEmptyText(textPE25(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE25(0) & textPE25(1) & textPE25(2) & String(1 - Len(textPE25(2)), "0") & textPE25(3) & String(2 - Len(textPE25(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE26(0)) = False And IsEmptyText(textPE26(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE26(0) & textPE26(1) & textPE26(2) & String(1 - Len(textPE26(2)), "0") & textPE26(3) & String(2 - Len(textPE26(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE27(0)) = False And IsEmptyText(textPE27(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE27(0) & textPE27(1) & textPE27(2) & String(1 - Len(textPE27(2)), "0") & textPE27(3) & String(2 - Len(textPE27(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE28(0)) = False And IsEmptyText(textPE28(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE28(0) & textPE28(1) & textPE28(2) & String(1 - Len(textPE28(2)), "0") & textPE28(3) & String(2 - Len(textPE28(3)), "0")
      nCount = nCount + 1
   End If
   If IsEmptyText(textPE29(0)) = False And IsEmptyText(textPE29(1)) = False Then
      ReDim Preserve strPENo(nCount + 1)
      strPENo(nCount) = textPE29(0) & textPE29(1) & textPE29(2) & String(1 - Len(textPE29(2)), "0") & textPE29(3) & String(2 - Len(textPE29(3)), "0")
      nCount = nCount + 1
   End If
   
   strDup = Empty
   For nX = 0 To nCount - 1
      For nY = 0 To nCount - 1
         If nX <> nY Then
            If strPENo(nX) = strPENo(nY) Then
               strDup = strPENo(nX)
               CheckPENoExist = True
               Exit For
            End If
         End If
      Next nY
      If CheckPENoExist = True Then
         Exit For
      End If
   Next nX
   
   If CheckPENoExist = True Then
      strTit = "檢核資料"
      strMsg = "失誤案號<" & strDup & ">重覆"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
End Function

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 4:
         ' 員工代號
         If IsEmptyText(textPE01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入員工代號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE01.SetFocus
            GoTo EXITSUB
         End If
         ' 系統類別
         If IsEmptyText(textPE02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入系統類別"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE02.SetFocus
            GoTo EXITSUB
         End If
         ' 目標年月
         If IsEmptyText(textPE03_1) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入目標年月中的年份"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE03_1.SetFocus
            GoTo EXITSUB
         End If
         ' 目標年月
         If IsEmptyText(textPE03_2) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入目標年月中的月份"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE03_2.SetFocus
            GoTo EXITSUB
         End If
   End Select
   
   Select Case m_EditMode
      Case 1, 2:
         '''''''''''''''''''''''''''''''''''''''''''''
         ' 檢查輸入的失誤案號是否完整
         If (IsEmptyText(textPE20(0)) = False And IsEmptyText(textPE20(1)) = True) Or (IsEmptyText(textPE20(0)) = True And IsEmptyText(textPE20(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE20(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE21(0)) = False And IsEmptyText(textPE21(1)) = True) Or (IsEmptyText(textPE21(0)) = True And IsEmptyText(textPE21(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE21(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE22(0)) = False And IsEmptyText(textPE22(1)) = True) Or (IsEmptyText(textPE22(0)) = True And IsEmptyText(textPE22(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE22(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE23(0)) = False And IsEmptyText(textPE23(1)) = True) Or (IsEmptyText(textPE23(0)) = True And IsEmptyText(textPE23(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE23(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE24(0)) = False And IsEmptyText(textPE24(1)) = True) Or (IsEmptyText(textPE24(0)) = True And IsEmptyText(textPE24(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE24(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE25(0)) = False And IsEmptyText(textPE25(1)) = True) Or (IsEmptyText(textPE25(0)) = True And IsEmptyText(textPE25(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE25(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE26(0)) = False And IsEmptyText(textPE26(1)) = True) Or (IsEmptyText(textPE26(0)) = True And IsEmptyText(textPE26(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE26(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE27(0)) = False And IsEmptyText(textPE27(1)) = True) Or (IsEmptyText(textPE27(0)) = True And IsEmptyText(textPE27(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE27(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE28(0)) = False And IsEmptyText(textPE28(1)) = True) Or (IsEmptyText(textPE28(0)) = True And IsEmptyText(textPE28(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE28(0).SetFocus
            GoTo EXITSUB
         End If
         If (IsEmptyText(textPE29(0)) = False And IsEmptyText(textPE29(1)) = True) Or (IsEmptyText(textPE29(0)) = True And IsEmptyText(textPE29(1)) = False) Then
            strTit = "檢核資料"
            strMsg = "失誤案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPE29(0).SetFocus
            GoTo EXITSUB
         End If
      
         '''''''''''''''''''''''''''''''''''''''''''''
         ' 檢查輸入的失誤案號是否存在於檔案中
         If IsEmptyText(textPE20(0)) = False And IsEmptyText(textPE20(1)) = False Then
            If IsDataExist(textPE20(0), textPE20(1), textPE20(2), textPE20(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE20(0) & "-" & textPE20(1) & "-" & textPE20(2) & "-" & textPE20(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE20(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE21(0)) = False And IsEmptyText(textPE21(1)) = False Then
            If IsDataExist(textPE21(0), textPE21(1), textPE21(2), textPE21(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE21(0) & "-" & textPE21(1) & "-" & textPE21(2) & "-" & textPE21(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE21(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE22(0)) = False And IsEmptyText(textPE22(1)) = False Then
            If IsDataExist(textPE22(0), textPE22(1), textPE22(2), textPE22(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE22(0) & "-" & textPE22(1) & "-" & textPE22(2) & "-" & textPE22(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE22(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE23(0)) = False And IsEmptyText(textPE23(1)) = False Then
            If IsDataExist(textPE23(0), textPE23(1), textPE23(2), textPE23(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE23(0) & "-" & textPE23(1) & "-" & textPE23(2) & "-" & textPE23(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE23(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE24(0)) = False And IsEmptyText(textPE24(1)) = False Then
            If IsDataExist(textPE24(0), textPE24(1), textPE24(2), textPE24(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE24(0) & "-" & textPE24(1) & "-" & textPE24(2) & "-" & textPE24(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE24(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE25(0)) = False And IsEmptyText(textPE25(1)) = False Then
            If IsDataExist(textPE25(0), textPE25(1), textPE25(2), textPE25(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE25(0) & "-" & textPE25(1) & "-" & textPE25(2) & "-" & textPE25(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE25(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE26(0)) = False And IsEmptyText(textPE26(1)) = False Then
            If IsDataExist(textPE26(0), textPE26(1), textPE26(2), textPE26(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE26(0) & "-" & textPE26(1) & "-" & textPE26(2) & "-" & textPE26(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE26(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE27(0)) = False And IsEmptyText(textPE27(1)) = False Then
            If IsDataExist(textPE27(0), textPE27(1), textPE27(2), textPE27(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE27(0) & "-" & textPE27(1) & "-" & textPE27(2) & "-" & textPE27(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE27(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE28(0)) = False And IsEmptyText(textPE28(1)) = False Then
            If IsDataExist(textPE28(0), textPE28(1), textPE28(2), textPE28(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE28(0) & "-" & textPE28(1) & "-" & textPE28(2) & "-" & textPE28(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE28(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textPE29(0)) = False And IsEmptyText(textPE29(1)) = False Then
            If IsDataExist(textPE29(0), textPE29(1), textPE29(2), textPE29(3)) = False Then
               strTit = "檢核資料"
               strMsg = "本所案號<" & textPE29(0) & "-" & textPE29(1) & "-" & textPE29(2) & "-" & textPE29(3) & ">不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textPE29(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         '''''''''''''''''''''''''''''''''''''''''''''
         ' 檢查所輸入的失誤案號是否重覆
         If CheckPENoExist() = True Then
            GoTo EXITSUB
         End If
   End Select
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPE01_GotFocus()
   InverseTextBox textPE01
End Sub

Private Sub textPE02_GotFocus()
   InverseTextBox textPE02
End Sub

Private Sub textPE03_1_GotFocus()
   InverseTextBox textPE03_1
End Sub

Private Sub textPE03_2_GotFocus()
   InverseTextBox textPE03_2
End Sub

Private Sub textPE04_GotFocus()
   InverseTextBox textPE04
End Sub

Private Sub textPE05_GotFocus()
   InverseTextBox textPE05
End Sub

Private Sub textPE06_GotFocus()
   InverseTextBox textPE06
End Sub

Private Sub textPE07_GotFocus()
   InverseTextBox textPE07
End Sub

Private Sub textPE08_GotFocus()
   InverseTextBox textPE08
End Sub

Private Sub textPE09_GotFocus()
   InverseTextBox textPE09
End Sub

Private Sub textPE10_GotFocus()
   InverseTextBox textPE10
End Sub

Private Sub textPE11_GotFocus()
   InverseTextBox textPE11
End Sub

Private Sub textPE12_GotFocus()
   InverseTextBox textPE12
End Sub

Private Sub textPE13_GotFocus()
   InverseTextBox textPE13
End Sub

Private Sub textPE14_GotFocus()
   InverseTextBox textPE14
End Sub

Private Sub textPE15_GotFocus()
   InverseTextBox textPE15
End Sub

Private Sub textPE16_GotFocus()
   InverseTextBox textPE16
End Sub

Private Sub textPE17_GotFocus()
   InverseTextBox textPE17
End Sub

Private Sub textPE18_GotFocus()
   InverseTextBox textPE18
End Sub

Private Sub textPE19_GotFocus()
   InverseTextBox textPE19
End Sub

Private Sub textPE20_GotFocus(Index As Integer)
   InverseTextBox textPE20(Index)
End Sub

Private Sub textPE21_GotFocus(Index As Integer)
   InverseTextBox textPE21(Index)
End Sub

Private Sub textPE22_GotFocus(Index As Integer)
   InverseTextBox textPE22(Index)
End Sub

Private Sub textPE23_GotFocus(Index As Integer)
   InverseTextBox textPE23(Index)
End Sub

Private Sub textPE24_GotFocus(Index As Integer)
   InverseTextBox textPE24(Index)
End Sub

Private Sub textPE25_GotFocus(Index As Integer)
   InverseTextBox textPE25(Index)
End Sub

Private Sub textPE26_GotFocus(Index As Integer)
   InverseTextBox textPE26(Index)
End Sub

Private Sub textPE27_GotFocus(Index As Integer)
   InverseTextBox textPE27(Index)
End Sub

Private Sub textPE28_GotFocus(Index As Integer)
   InverseTextBox textPE28(Index)
End Sub

Private Sub textPE29_GotFocus(Index As Integer)
   InverseTextBox textPE29(Index)
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textPE01.Enabled = True Then
   Cancel = False
   textPE01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE02.Enabled = True Then
   Cancel = False
   textPE02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE03_1.Enabled = True Then
   Cancel = False
   textPE03_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE03_2.Enabled = True Then
   Cancel = False
   textPE03_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE04.Enabled = True Then
   Cancel = False
   textPE04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE05.Enabled = True Then
   Cancel = False
   textPE05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE06.Enabled = True Then
   Cancel = False
   textPE06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE07.Enabled = True Then
   Cancel = False
   textPE07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE08.Enabled = True Then
   Cancel = False
   textPE08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE09.Enabled = True Then
   Cancel = False
   textPE09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE10.Enabled = True Then
   Cancel = False
   textPE10_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE11.Enabled = True Then
   Cancel = False
   textPE11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE12.Enabled = True Then
   Cancel = False
   textPE12_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE13.Enabled = True Then
   Cancel = False
   textPE13_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE14.Enabled = True Then
   Cancel = False
   textPE14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE15.Enabled = True Then
   Cancel = False
   textPE15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE16.Enabled = True Then
   Cancel = False
   textPE16_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE17.Enabled = True Then
   Cancel = False
   textPE17_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE18.Enabled = True Then
   Cancel = False
   textPE18_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPE19.Enabled = True Then
   Cancel = False
   textPE19_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

For Each objTxt In Me.textPE20
   If objTxt.Enabled = True Then
      Cancel = False
      textPE20_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In textPE21
   If objTxt.Enabled = True Then
      Cancel = False
      textPE21_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In textPE22
   If objTxt.Enabled = True Then
      Cancel = False
      textPE22_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In textPE23
   If objTxt.Enabled = True Then
      Cancel = False
      textPE23_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In textPE24
   If objTxt.Enabled = True Then
      Cancel = False
      textPE24_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In textPE25
   If objTxt.Enabled = True Then
      Cancel = False
      textPE25_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In textPE26
   If objTxt.Enabled = True Then
      Cancel = False
      textPE26_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In textPE27
   If objTxt.Enabled = True Then
      Cancel = False
      textPE27_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In textPE28
   If objTxt.Enabled = True Then
      Cancel = False
      textPE28_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

For Each objTxt In textPE29
   If objTxt.Enabled = True Then
      Cancel = False
      textPE29_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function

