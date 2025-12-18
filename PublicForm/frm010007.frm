VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010007 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   6220
   ClientLeft      =   5690
   ClientTop       =   370
   ClientWidth     =   9010
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6220
   ScaleWidth      =   9010
   Begin VB.CommandButton cmdTSMap 
      BackColor       =   &H0000FF00&
      Caption         =   "查名代號"
      Height          =   400
      Left            =   4080
      Style           =   1  '圖片外觀
      TabIndex        =   39
      Top             =   -45
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCP64 
      Height          =   270
      Left            =   180
      TabIndex        =   84
      Top             =   45
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8220
      TabIndex        =   42
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6270
      TabIndex        =   40
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7095
      TabIndex        =   41
      Top             =   15
      Width           =   1100
   End
   Begin VB.Frame fraWindow1 
      BorderStyle     =   0  '沒有框線
      Height          =   5800
      Left            =   15
      TabIndex        =   43
      Top             =   384
      Width           =   8955
      Begin VB.TextBox txtRecieveCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1092
         TabIndex        =   107
         Top             =   10
         Width           =   1452
      End
      Begin VB.Frame FraTCN 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame1"
         Height          =   372
         Left            =   3984
         TabIndex        =   104
         Top             =   5472
         Width           =   4572
         Begin VB.TextBox txtTCN01 
            Height          =   300
            Left            =   1200
            MaxLength       =   9
            TabIndex        =   105
            Top             =   0
            Width           =   2700
         End
         Begin VB.Label Label38 
            Caption         =   "追蹤流水號："
            Height          =   280
            Left            =   24
            TabIndex        =   106
            Top             =   24
            Width           =   1188
         End
      End
      Begin VB.Frame Frame21 
         BorderStyle     =   0  '沒有框線
         Height          =   375
         Left            =   1300
         TabIndex        =   101
         Top             =   5460
         Visible         =   0   'False
         Width           =   6435
         Begin VB.TextBox textEP06 
            Height          =   270
            Left            =   1350
            MaxLength       =   1
            TabIndex        =   36
            Top             =   60
            Width           =   255
         End
         Begin VB.TextBox textEP34 
            Height          =   270
            Left            =   3480
            MaxLength       =   1
            TabIndex        =   37
            Top             =   60
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "文件是否齊備：       (Y/N)"
            Height          =   180
            Left            =   60
            TabIndex        =   103
            Top             =   105
            Width           =   1980
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "是否會稿：       (Y/N)"
            Height          =   180
            Left            =   2550
            TabIndex        =   102
            Top             =   105
            Visible         =   0   'False
            Width           =   1680
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "有★★的應收帳款簽核控管"
         Height          =   255
         Left            =   6330
         TabIndex        =   26
         Top             =   4500
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         Caption         =   "現金或支票"
         Height          =   200
         Left            =   4680
         TabIndex        =   33
         Top             =   5235
         Width           =   1215
      End
      Begin VB.Frame fraPromoter 
         BorderStyle     =   0  '沒有框線
         Height          =   320
         Left            =   4410
         TabIndex        =   81
         Top             =   4725
         Width           =   4245
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   20
            Left            =   705
            TabIndex        =   38
            Top             =   75
            Width           =   1110
            VariousPropertyBits=   679493659
            MaxLength       =   6
            Size            =   "1958;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPromoter 
            Height          =   255
            Left            =   1950
            TabIndex        =   83
            Top             =   45
            Width           =   1845
            VariousPropertyBits=   27
            Size            =   "3254;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "承辦人："
            Height          =   180
            Left            =   0
            TabIndex        =   82
            Top             =   135
            Width           =   720
         End
      End
      Begin VB.Frame fraWindow2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3345
         Left            =   60
         TabIndex        =   58
         Top             =   560
         Width           =   8865
         Begin VB.TextBox txtTS 
            Height          =   300
            Index           =   0
            Left            =   5850
            MaxLength       =   3
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   990
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtTS 
            Height          =   300
            Index           =   1
            Left            =   6330
            MaxLength       =   6
            TabIndex        =   32
            Top             =   990
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "聯絡人資料"
            Height          =   300
            Left            =   7600
            TabIndex        =   9
            Top             =   960
            Width           =   1155
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   3540
            MaxLength       =   2
            TabIndex        =   67
            Top             =   150
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   3120
            MaxLength       =   1
            TabIndex        =   66
            Top             =   150
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   288
            Index           =   0
            Left            =   1860
            MaxLength       =   6
            TabIndex        =   65
            Top             =   150
            Width           =   1212
         End
         Begin VB.TextBox txtSystem 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   59
            Top             =   150
            Width           =   732
         End
         Begin MSForms.ComboBox cboContact 
            Height          =   300
            Left            =   7020
            TabIndex        =   4
            Top             =   150
            Width           =   1770
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "3122;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   30
            Left            =   1410
            TabIndex        =   8
            Top             =   990
            Width           =   1035
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1826;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   29
            Left            =   5850
            TabIndex        =   7
            Top             =   720
            Width           =   1035
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1826;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   27
            Left            =   5670
            TabIndex        =   11
            Top             =   1290
            Width           =   2565
            VariousPropertyBits=   679493659
            MaxLength       =   50
            Size            =   "4524;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   26
            Left            =   1530
            TabIndex        =   17
            Top             =   2730
            Width           =   7280
            VariousPropertyBits=   -1466941413
            MaxLength       =   699
            ScrollBars      =   2
            Size            =   "12841;529"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   25
            Left            =   1530
            TabIndex        =   16
            Top             =   2430
            Width           =   7280
            VariousPropertyBits=   -1466941413
            MaxLength       =   395
            ScrollBars      =   2
            Size            =   "12841;529"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   830
            Index           =   24
            Left            =   1920
            TabIndex        =   12
            Top             =   1590
            Width           =   6890
            VariousPropertyBits=   -1466941413
            MaxLength       =   140
            ScrollBars      =   2
            Size            =   "12153;1464"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   21
            Left            =   990
            TabIndex        =   18
            Top             =   3020
            Width           =   7820
            VariousPropertyBits=   679493659
            MaxLength       =   2000
            Size            =   "13794;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   19
            Left            =   1410
            TabIndex        =   6
            Top             =   720
            Width           =   1035
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1826;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   18
            Left            =   5850
            TabIndex        =   5
            Top             =   450
            Width           =   1035
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1826;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   8
            Left            =   840
            TabIndex        =   10
            Top             =   1290
            Width           =   1095
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   7
            Left            =   1410
            TabIndex        =   3
            Top             =   450
            Width           =   1035
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1826;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   6
            Left            =   1920
            TabIndex        =   15
            Top             =   2130
            Width           =   6765
            VariousPropertyBits=   679493659
            MaxLength       =   160
            Size            =   "11933;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   5
            Left            =   1920
            TabIndex        =   14
            Top             =   1860
            Width           =   6765
            VariousPropertyBits=   679493659
            MaxLength       =   180
            Size            =   "11933;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   4
            Left            =   1920
            TabIndex        =   13
            Top             =   1590
            Width           =   6765
            VariousPropertyBits=   679493659
            MaxLength       =   160
            Size            =   "11933;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtOther 
            Height          =   300
            Index           =   3
            Left            =   5160
            TabIndex        =   2
            Top             =   150
            Width           =   615
            VariousPropertyBits=   679493659
            MaxLength       =   3
            Size            =   "1085;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label lblTS 
            Caption         =   "查名代號："
            Height          =   240
            Left            =   4920
            TabIndex        =   100
            Top             =   1020
            Visible         =   0   'False
            Width           =   975
         End
         Begin MSForms.Label lblPetition 
            Height          =   255
            Index           =   4
            Left            =   2460
            TabIndex        =   98
            Top             =   1050
            Width           =   1905
            VariousPropertyBits=   27
            Size            =   "3360;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label35 
            Caption         =   "申請人/當事人5："
            Height          =   255
            Left            =   30
            TabIndex        =   97
            Top             =   990
            Width           =   1605
         End
         Begin MSForms.Label lblPetition 
            Height          =   255
            Index           =   3
            Left            =   6900
            TabIndex        =   96
            Top             =   720
            Width           =   1905
            VariousPropertyBits=   27
            Size            =   "3360;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label34 
            Caption         =   "申請人/當事人4："
            Height          =   255
            Left            =   4470
            TabIndex        =   95
            Top             =   750
            Width           =   1605
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "接洽人："
            Height          =   180
            Left            =   6270
            TabIndex        =   94
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "代理人彼所案號："
            Height          =   180
            Left            =   4260
            TabIndex        =   92
            Top             =   1350
            Width           =   1440
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "商品組群  (699) ："
            Height          =   180
            Left            =   120
            TabIndex        =   91
            Top             =   2730
            Width           =   1425
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "商品類別  (395) ："
            Height          =   180
            Left            =   120
            TabIndex        =   90
            Top             =   2430
            Width           =   1425
         End
         Begin VB.Label Label29 
            Caption         =   "案件名稱（140）："
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   1620
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "案件備註："
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   3015
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "申請人/當事人3："
            Height          =   255
            Left            =   30
            TabIndex        =   80
            Top             =   750
            Width           =   1605
         End
         Begin MSForms.Label lblPetition 
            Height          =   255
            Index           =   2
            Left            =   2460
            TabIndex        =   79
            Top             =   720
            Width           =   1905
            VariousPropertyBits=   27
            Size            =   "3360;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label4 
            Caption         =   "申請人/當事人2："
            Height          =   255
            Left            =   4470
            TabIndex        =   78
            Top             =   480
            Width           =   1605
         End
         Begin MSForms.Label lblPetition 
            Height          =   255
            Index           =   1
            Left            =   6900
            TabIndex        =   77
            Top             =   480
            Width           =   1905
            VariousPropertyBits=   27
            Size            =   "3360;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label14 
            Caption         =   "案件名稱(日)（160）："
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   2160
            Width           =   1935
         End
         Begin VB.Label Label13 
            Caption         =   "案件名稱(英)（180）："
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1890
            Width           =   1935
         End
         Begin VB.Label Label11 
            Caption         =   "案件名稱(中)（160）："
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   1620
            Width           =   1900
         End
         Begin MSForms.Label lblAgent 
            Height          =   255
            Left            =   1950
            TabIndex        =   69
            Top             =   1350
            Width           =   2295
            VariousPropertyBits=   27
            Size            =   "4048;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label8 
            Caption         =   "代理人："
            Height          =   225
            Left            =   120
            TabIndex        =   68
            Top             =   1320
            Width           =   855
         End
         Begin MSForms.Label lblPetition 
            Height          =   255
            Index           =   0
            Left            =   2460
            TabIndex        =   64
            Top             =   450
            Width           =   1905
            VariousPropertyBits=   27
            Size            =   "3360;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label17 
            Caption         =   "申請人/當事人1："
            Height          =   255
            Left            =   30
            TabIndex        =   63
            Top             =   480
            Width           =   1605
         End
         Begin VB.Label Label9 
            Caption         =   "申請國家："
            Height          =   255
            Left            =   4200
            TabIndex        =   62
            Top             =   210
            Width           =   975
         End
         Begin VB.Label lblNation 
            Height          =   255
            Left            =   5850
            TabIndex        =   61
            Top             =   210
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "本所案號："
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   180
            Width           =   975
         End
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   0
         Left            =   5160
         TabIndex        =   108
         Top             =   10
         Width           =   1090
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   31
         Left            =   1050
         TabIndex        =   35
         Top             =   5460
         Width           =   7785
         VariousPropertyBits=   679493659
         MaxLength       =   2000
         Size            =   "13732;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   28
         Left            =   3450
         TabIndex        =   30
         Top             =   5145
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   23
         Left            =   5850
         TabIndex        =   23
         Top             =   4185
         Width           =   2715
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "4789;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   22
         Left            =   5850
         TabIndex        =   21
         Top             =   3885
         Width           =   2715
         VariousPropertyBits=   679493659
         MaxLength       =   30
         Size            =   "4789;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   15
         Left            =   4500
         TabIndex        =   25
         Top             =   4470
         Width           =   495
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "868;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   10
         Left            =   1050
         TabIndex        =   22
         Top             =   4185
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   6
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   9
         Left            =   3192
         TabIndex        =   20
         Top             =   3888
         Width           =   1092
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   13
         Left            =   1050
         TabIndex        =   29
         Top             =   5145
         Width           =   1095
         VariousPropertyBits=   679493659
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   12
         Left            =   1050
         TabIndex        =   27
         Top             =   4815
         Width           =   1095
         VariousPropertyBits=   679493659
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   14
         Left            =   1056
         TabIndex        =   19
         Top             =   3888
         Width           =   1092
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   17
         Left            =   7530
         TabIndex        =   34
         Top             =   5145
         Width           =   1095
         VariousPropertyBits=   679493659
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   16
         Left            =   3090
         TabIndex        =   28
         Top             =   4815
         Width           =   1095
         VariousPropertyBits=   679493659
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   11
         Left            =   1050
         TabIndex        =   24
         Top             =   4485
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   5
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   1
         Left            =   1090
         TabIndex        =   0
         Top             =   290
         Width           =   610
         VariousPropertyBits=   679493659
         MaxLength       =   4
         Size            =   "1080;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   2
         Left            =   5160
         TabIndex        =   1
         Top             =   290
         Width           =   370
         VariousPropertyBits=   679493659
         MaxLength       =   2
         Size            =   "656;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCaseSource 
         Height          =   200
         Left            =   5550
         TabIndex        =   46
         Top             =   350
         Width           =   2780
      End
      Begin VB.Label Label36 
         Caption         =   "主題："
         Height          =   255
         Left            =   90
         TabIndex        =   99
         Top             =   5460
         Width           =   975
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "預定收款日："
         Height          =   180
         Left            =   2340
         TabIndex        =   93
         Top             =   5205
         Width           =   1080
      End
      Begin VB.Label Label28 
         Caption         =   "分所案號："
         Height          =   255
         Left            =   4770
         TabIndex        =   88
         Top             =   4185
         Width           =   1005
      End
      Begin VB.Label Label15 
         Caption         =   "客戶案件案號："
         Height          =   255
         Left            =   4410
         TabIndex        =   87
         Top             =   3885
         Width           =   1395
      End
      Begin VB.Label Label19 
         Caption         =   "是否開電腦收據：           （N：不開)"
         Height          =   255
         Left            =   3030
         TabIndex        =   76
         Top             =   4500
         Width           =   2925
      End
      Begin VB.Label lblDepartment 
         Height          =   255
         Left            =   3930
         TabIndex        =   75
         Top             =   4185
         Width           =   750
      End
      Begin VB.Label Label18 
         Caption         =   "業務區："
         Height          =   255
         Left            =   3210
         TabIndex        =   74
         Top             =   4215
         Width           =   750
      End
      Begin VB.Label Label16 
         Caption         =   "法定期限："
         Height          =   252
         Left            =   90
         TabIndex        =   73
         Top             =   3912
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "收文號："
         Height          =   260
         Left            =   120
         TabIndex        =   57
         Top             =   70
         Width           =   720
      End
      Begin VB.Label Label7 
         Caption         =   "收文日："
         Height          =   260
         Left            =   4080
         TabIndex        =   56
         Top             =   70
         Width           =   860
      End
      Begin VB.Label Label20 
         Caption         =   "本所期限："
         Height          =   252
         Left            =   2256
         TabIndex        =   55
         Top             =   3912
         Width           =   972
      End
      Begin VB.Label Label21 
         Caption         =   "郵遞區號："
         Height          =   255
         Left            =   90
         TabIndex        =   54
         Top             =   4485
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "費用："
         Height          =   255
         Left            =   90
         TabIndex        =   53
         Top             =   4815
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "點數："
         Height          =   255
         Left            =   90
         TabIndex        =   52
         Top             =   5145
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "智權人員："
         Height          =   255
         Left            =   90
         TabIndex        =   51
         Top             =   4185
         Width           =   975
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "規費："
         Height          =   180
         Left            =   2550
         TabIndex        =   50
         Top             =   4860
         Width           =   540
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "後金："
         Height          =   180
         Left            =   6900
         TabIndex        =   49
         Top             =   5175
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "案件性質："
         Height          =   260
         Left            =   120
         TabIndex        =   48
         Top             =   350
         Width           =   980
      End
      Begin VB.Label Label5 
         Caption         =   "案件來源："
         Height          =   260
         Left            =   4080
         TabIndex        =   47
         Top             =   350
         Width           =   960
      End
      Begin VB.Label lblCaseProperty 
         Height          =   200
         Left            =   1710
         TabIndex        =   45
         Top             =   350
         Width           =   2180
      End
      Begin MSForms.Label lblSales 
         Height          =   255
         Left            =   2190
         TabIndex        =   44
         Top             =   4185
         Width           =   855
         VariousPropertyBits=   27
         Size            =   "1508;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "T及TS的商品類別請輸在""案件備註""欄內!!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1650
      TabIndex        =   86
      Top             =   75
      Visible         =   0   'False
      Width           =   3570
   End
End
Attribute VB_Name = "frm010007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/21 Form2.0已修改 txtOther()/lblpetition()/lblAgent/lblSales/lblPromoter/cboContact
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

'intCaseKind系統別
Public intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗
'LastData上一次存檔時，所輸入之收文日
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim bolLeave As Boolean, LastDate As String, intLeaveKind As Integer
Dim strNation As String, douStPrice As Double, douLowPrice As Double

'Add by Morgan 2004/4/15
'是否已觸發 Form Active 事件
Dim bolActive As Boolean
'add by nickc 2007/12/12
Dim IsSaveData As Boolean
Dim strAppNo1 As String '申請人1編號
Dim strDefCont1 As String '申請人1的預設接洽人

'Add By Sindy 2010/3/8 回傳值
Dim strSP30s As String, strSP75s As String
Dim bolCancel As Boolean
'2010/3/8 End
Dim dblAmt As Double, dblPFee As Double, dblTFee As Double, m_CP150 As String 'Add By Sindy 2012/11/06
Dim dblChkAmt As Double 'Add By Sindy 2012/12/10
'Added by Lydia 2020/02/03
Dim dblCu183 As Double '個人之應收帳款上限
Dim dblAmtR As Double, dblPFeeR As Double, dblTFeeR As Double '關係企業之應收帳款金額
'end 2020/02/03

Dim mFMPchk As Boolean, mCP31 As String 'Add by Lydia 2014/10/31 開放外專程序人員可進入專利處系統操作FMP寰華案件
'Added by Lydia 2016/04/25 新增查名單對應
Dim m_AttachPath As String
Public m_PrevForm As Form
Public Tmpfrm090130 As Form
Public TMQList As String
Dim bolOpen130 As Boolean '是否開啟過查名代號表單
'Added by Lydia 2019/02/14
Dim m_SalesST15 As String '畫面上智權人員的收文部門
Dim m_Tuser As String '創新業務部預設收文人員
'Added by Lydia 2019/09/16
Dim m_SalesST06 As String '智權人員的所別
'Added by Lydia 2020/05/20 法律所案源收文
Dim m_LOS01 As String '案源總收文號
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '案源總收文號之本所案號
Dim m_LOS02 As String '案源案件類型
Dim t_LOSkind As String '案件性質=>案源案件類型
Dim m_LOS15 As String '案源單號
Dim m_LOS04 As String  '介紹人
Dim m_LOS04_1 As String, m_LOS04_1st15 As String, m_LOS04_1st06 As String '介紹人(第一位)、收文部門、所別
Dim m_LOS05 As String  '介紹客戶
Dim m_LOS12 As String  '介紹日
Dim m_Los04_N1 As String, m_Los05_N As String  'Added by Lydia 2020/10/05 LA補案源之介紹人(第一位), 介紹客戶
'Mark by Lydia 2022/09/06 改抓特殊設定
'Private Const cnt應收帳款檢查排除 As String = "74018,70005" 'Added by Lydia 2022/06/15 應收帳款上限檢查排除特定人員: 如果人員有異動, 請一併修改接洽單frm090801和收文frm010004~frm010007
Dim m_EP06 As String 'Added by Lydia 2022/07/15 TC案之文件齊備日管控
'Add By Sindy 2022/8/17
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'Public m_PrevForm As Form '前一畫面
Public m_bMRecvBatch As Boolean '信件沖銷多案收文
Dim m_bolRecvOK As Boolean '是否收完文
Dim m_strMCR11 As String '多案收文時,第一筆的總收文號
'2022/8/17 END

'Added by Lydia 2022/09/14 櫃台收文模組化
Private Const 收文存檔模組化啟用日 = 20220928 '完成後先開始使用
Dim modCP() As String, modBase() As String ' 收文 和 基本檔
Dim m_strControl As String  '齊備日管制
Dim mType As String, mCaseNo As String  '特殊管制
Dim mChkStr As String   '其他操作結果
Dim bolisNP0809 As Boolean 'Added by Lydia 2023/06/08 是否從下一程序檔取回本所期限、法定期限
'Added by Lydia 2024/12/13 FG案輸入追蹤流水號TrackingNo
Dim bolMoveCheck As Boolean '存檔前先檢查TrackingNO檔案是否存在
Dim mSaveDir As String 'TrackingNO：暫存下載檔案的本機端資料夾
Dim bolMoveOK As Boolean ' TrackingNO是否已搬檔完成(True無問題)
'end 2024/12/13

'Added by Lydia 2022/09/14 設定陣列
Private Sub SetDBArray(ByVal bolReset As Boolean, ByVal pSNo As String, ByVal pCD01 As String, Optional ByVal pCD02 As String, Optional ByVal pCD03 As String, Optional ByVal pCD04 As String)
'pSNo: 現在的收文號 (1碼=新增)
'pCD01~pCD04: 本所案號
Dim intKind As Integer, intWhere As Integer
Dim strTmpA As String

   If bolReset = True Then
      If ClsPDGetSystemKind(pCD01, intKind) = True Then
        Select Case intKind
           Case 專利
              ReDim Preserve modBase(TF_PA) As String
           Case 商標
              ReDim Preserve modBase(TF_TM) As String
           Case 法務
              ReDim Preserve modBase(TF_LC) As String
           Case 顧問
              ReDim Preserve modBase(TF_HC) As String
           Case Else
              ReDim Preserve modBase(tf_SP) As String
        End Select
      End If
      ReDim Preserve modCP(TF_CP) As String
      modBase(1) = pCD01
      modBase(2) = pCD02
      'Added by Lydia 2022/11/11  debug: CFP-029190-0-40收文子案存成母案
      modBase(3) = pCD03
      modBase(4) = pCD04
      'end 2022/11/11
      If pCD01 <> "" And pCD02 <> "" And pCD02 <> "0" Then
         If modBase(3) = "" Then modBase(3) = "0"
         If modBase(4) = "" Then modBase(4) = "00"
         'Modified by Lydia 2023/05/12 + false
         If PUB_ReadCaseData(modBase, intKind, intWhere, False) = True Then
         End If
      End If
      modCP(1) = modBase(1)
      modCP(2) = modBase(2)
      modCP(3) = modBase(3)
      modCP(4) = modBase(4)
   Else
      If modBase(3) = "" Then
          modBase(3) = "0"
          modCP(3) = "0"
      End If
      If modBase(4) = "" Then
          modBase(4) = "00"
          modCP(4) = "00"
      End If
      '考慮多案收文,再設定一次
      modBase(1) = pCD01
      modBase(2) = pCD02
      modCP(1) = modBase(1)
      modCP(2) = modBase(2)
      modCP(3) = modBase(3)
      modCP(4) = modBase(4)
      '---------------
      
      '基本檔
      strExc(1) = "": strExc(2) = ""
      
      If ClsPDGetSystemKind(pCD01, intKind) = True Then
        Select Case intKind
           Case 法務
                modBase(5) = txtOther(4)  '案件名稱(中)
                modBase(6) = txtOther(5)  '案件名稱(英)
                modBase(7) = txtOther(6)  '案件名稱(日)
                '當事人1~5
                modBase(11) = ChangeCustomerL(txtOther(7))
                modBase(43) = ChangeCustomerL(txtOther(18))
                modBase(44) = ChangeCustomerL(txtOther(19))
                modBase(45) = ChangeCustomerL(txtOther(29))
                modBase(46) = ChangeCustomerL(txtOther(30))
                
                strExc(1) = "11"  '當事人1
                strExc(2) = "42" '聯絡人編號
                modBase(15) = txtOther(3) '申請國家
                modBase(16) = txtOther(23) '分所案號
                modBase(17) = txtOther(22) '客戶案件案號
                modBase(22) = ChangeCustomerL(txtOther(8))  'FC代理人
                modBase(23) = txtOther(27) '代理人彼所案號
                modBase(27) = txtOther(21) '案件備註
                
           Case 顧問
                If txtOther(24).Visible = True Then
                   modBase(6) = txtOther(24)
                Else
                   modBase(6) = txtOther(4)  '案件名稱(中)
                End If
                '當事人1~5
                modBase(5) = ChangeCustomerL(txtOther(7))
                modBase(24) = ChangeCustomerL(txtOther(18))
                modBase(25) = ChangeCustomerL(txtOther(19))
                modBase(26) = ChangeCustomerL(txtOther(29))
                modBase(27) = ChangeCustomerL(txtOther(30))
                
                modBase(7) = txtOther(23) '分所案號
                strExc(1) = "5"  '當事人1
                strExc(2) = "23" '聯絡人編號
                modBase(12) = txtOther(21) '案件備註
                
           Case Else  '服務
                If txtOther(24).Visible = True Then
                   modBase(5) = txtOther(24)
                Else
                   modBase(5) = txtOther(4)  '案件名稱(中)
                End If
                modBase(6) = txtOther(5)  '案件名稱(英)
                modBase(7) = txtOther(6)  '案件名稱(日)
                '申請人1~5
                modBase(8) = ChangeCustomerL(txtOther(7))
                modBase(58) = ChangeCustomerL(txtOther(18))
                modBase(59) = ChangeCustomerL(txtOther(19))
                modBase(65) = ChangeCustomerL(txtOther(29))
                modBase(66) = ChangeCustomerL(txtOther(30))
                modBase(28) = txtOther(23) '分所案號
                modBase(29) = txtOther(22) '客戶案件案號
                strExc(1) = "8"  '申請人1
                strExc(2) = "78" '聯絡人編號
                modBase(9) = txtOther(3) '申請國家
                modBase(26) = ChangeCustomerL(txtOther(8))  'FC代理人
                modBase(27) = txtOther(27) '代理人彼所案號
                modBase(18) = txtOther(21) '案件備註
                modBase(73) = txtOther(25) '商品類別
                modBase(74) = txtOther(26) '商品組群
                '聯絡人1~2 =>  frm010007_1.bolOK
                If bolCancel = True Then
                   modBase(30) = strSP30s
                   modBase(75) = strSP75s
                End If
        End Select
      End If
      
      '申請人聯絡人編號
      If cboContact.Locked = False Then
         If cboContact.ListIndex >= 0 Then
            modBase(Val(strExc(2))) = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
            If Val(modBase(Val(strExc(2)))) > 0 Then
            'Add by Morgan 2008/8/7 若個案接洽人與客戶檔的預設接洽人相同時不必設定
               PUB_GetContact modBase(Val(strExc(1))), strTmpA, True
               If modBase(Val(strExc(2))) = strTmpA Then
                  modBase(Val(strExc(2))) = ""
               End If
            '排除空白=00
            ElseIf modBase(Val(strExc(2))) = "00" And Trim(cboContact.Text) = "" Then
               modBase(Val(strExc(2))) = ""
            End If
         End If
      End If
      
      '收文CaseProgress
      modCP(9) = txtRecieveCode  '收文號
      modCP(5) = ChangeTStringToWString(txtOther(0)) '收文日
      modCP(6) = ChangeTStringToWString(txtOther(9)) '本所期限
      modCP(7) = ChangeTStringToWString(txtOther(14))  '法定期限
      modCP(10) = Trim(txtOther(1)) '案件性質
      modCP(11) = Trim(txtOther(2)) '案件來源
      modCP(12) = GetST15(txtOther(10))
      modCP(13) = Trim(txtOther(10))       '智權人員
      modCP(14) = Trim(txtOther(20))    '承辦人
      modCP(16) = txtOther(12)    '費用
      modCP(17) = txtOther(16)    '規費
      modCP(18) = txtOther(13)    '點數
      modCP(19) = txtOther(17)    '後金
      modCP(32) = txtOther(15) '是否開電腦收據
      modCP(33) = douStPrice '標準價
      modCP(34) = douLowPrice '底價
      modCP(64) = txtOther(31) '進度備註; 主題

      '有★★的應收帳款簽核控管
      If Check2.Visible = True Then
         modCP(150) = IIf(Check2.Value = 1, "Y", "")
      End If
      '特殊管制
      mType = "": mCaseNo = ""
      If m_LOS02 <> "" And m_LOS15 <> "" Then
          mType = "LOS案源收文"
          mCaseNo = m_LOS02 & "," & m_LOS15
      ElseIf txtSystem = "TS" And TMQList <> "" Then
          mType = "T查名單"
          mCaseNo = TMQList
      'Modify By Sindy 2025/8/18 發生了案源+信件沖銷 ex:FCP-057445/FCL-011034
      'ElseIf m_strIR01 <> "" Then
      End If
      If m_strIR01 <> "" Then
      '2025/8/18 END
          'Modify By Sindy 2023/5/31
          'mType = "外專信件沖銷"
          'mType = "信件沖銷"
          mType = mType & "-信件沖銷" 'Modify By Sindy 2025/8/18 + "-"
          '2023/5/31 END
          If m_bMRecvBatch = True Then mType = mType & "-多案收文"
          'Modified by Lydia 2022/10/06 debug
          'mCaseNo = m_strIR01 & m_strIR02 & m_strIR03 & m_strIR04
          'Modify By Sindy 2025/8/18 + IIf(mCaseNo <> "", mCaseNo & "-", "") &
          mCaseNo = IIf(mCaseNo <> "", mCaseNo & "-", "") & m_strIR01 & "," & m_strIR02 & "," & m_strIR03 & "," & m_strIR04
      End If
      
      '傳入其他操作結果
      mChkStr = ","
      If Trim(txtSystem) = "PS" And (mCP31 = "Y" Or frm010001.intSaveMode = 1) And m_LOS15 = "" And mFMPchk = True Then
          mChkStr = mChkStr & "寰華案件確認,"
      End If
      
      '齊備日  --m_strControl
      Call GetStrControl
   End If
   
   'Added by Lydia 2024/12/13 FG案輸入追蹤流水號
   If Trim(txtSystem) = "FG" And FraTCN.Visible = True And Trim(txtTCN01) <> "" Then
      mType = "追蹤流水號"
      mCaseNo = Trim(txtTCN01)
   End If
   'end 2024/12/13
End Sub

'Added by Lydia 2016/04/25
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      '2011/4/22 MODIFY BY SONIA 分所智權人員則多一天
      'txtOther(28) = PUB_GetWorkDayAfterSysDate(CDbl(txtOther(0)) + 19110000, 5)
      'Modified by Lydia 2019/09/16
      'If PUB_GetST06(txtOther(10)) <> "1" Then
      If m_SalesST06 <> "1" Then
      'end 2019/09/16
         txtOther(28) = PUB_GetWorkDayAfterSysDate(CDbl(txtOther(0)) + 19110000, 6)
      Else
         txtOther(28) = PUB_GetWorkDayAfterSysDate(CDbl(txtOther(0)) + 19110000, 5)
      End If
      '2011/4/22 END
      txtOther(28).Locked = True
   Else
      txtOther(28).Locked = False
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim varSaveCursor, strAuto1 As String, strAuto2 As String, i As Integer
Dim mBillNo As String, mMemo As String 'Added by Lydia 2019/05/13
Dim bolSaveOK As Boolean, mRetVal As String 'Added by Lydia 2022/09/14

'Dim tST15 As String 'Remove by Lydia 2019/02/14

'tST15 = Trim(PUB_GetStaffST15(txtOther(10), "1")) 'Remove Lydia 2019/02/14

If Index = 0 Then
   'Add By Sindy 2022/8/17 信件沖銷多案收文
   If m_strIR01 <> "" And m_bMRecvBatch = True Then
      '加入秀訊息
      If PUB_CheckFormExist("frmpic002") = False Then
         Load frmpic002
         frmpic002.Label1.Caption = "自動收文中...請稍候..."
         frmpic002.Show
      End If
      frmpic002.ZOrder 0
   End If
   '2022/8/17 END
   
   'Add by Amy 2021/12/21檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        Exit Sub
   End If
   'Added by Lydia 2017/07/31 預設和檢查-所有內部收文, 若有輸入本所期限或法定期限者
   'Modified by Lyddia 2023/11/08 傳入必需欄位
   'If PUB_CheckCP0607(0, txtOther(9), txtOther(14)) = False Then Exit Sub
   If PUB_CheckCP0607(0, txtOther(9), txtOther(14), IIf(frm010001.intModifyKind = 0, "Y", ""), txtOther(3), txtSystem, txtOther(1)) = False Then Exit Sub
   
   'Modified by Lydia 2019/09/16
   'm_SalesST15 = GetST15(txtOther(10)) 'Added by Lydia 2019/02/14
   m_SalesST15 = GetST15(txtOther(10), , , m_SalesST06)
   'Added by Lydia 2020/04/08 法務案(L、CFL)及顧問案LA之智權人員只能是法律所人員
   If PUB_ChkSalesL(txtSystem, txtOther(10).Text) = False Then
        txtOther(10).SetFocus
        txtOther_GotFocus 10
        Exit Sub
   End If
   'end 2020/04/08
   'Added by Lydia 2022/07/15 TC案之文件齊備日管控
   If Frame21.Visible = True Then
       If textEP06.Visible = True And Trim(textEP06) = "" Then
           MsgBox Left(Label41.Caption, 2) & "是否齊備不可空白!!!", vbExclamation + vbOKOnly
           Me.textEP06.SetFocus
           Exit Sub
       End If
       If textEP34.Visible = True And Trim(textEP34) = "" Then
           MsgBox "是否會稿不可空白!!!", vbExclamation + vbOKOnly
           Me.textEP34.SetFocus
           Exit Sub
       End If
   End If
   'end 2022/07/15
   
    'Added by Lydia 2020/06/05 法律所案源收文：(著作權)5/28台灣案之B1、B2及C收文時，增加"案源單號"欄位，B1、B2一定要輸入，C若未輸入則提醒'請確認接洽單沒有案源單號？'，案源單號更新至該筆收文的CP162。
    'Mark by Lydia 2020/06/10 重整判斷,以案源單的案源類型為準; 保留舊程式
'    If frm010001.intModifyKind = 0 And strSrvDate(1) >= 法律所案源收文啟用日 And txtSystem = "TC" Then
'         t_LOSkind = PUB_GetLOSkind(txtSystem, txtOther(1), txtOther(3))
'         If Left(t_LOSkind, 1) = "B" And m_LOS15 = "" Then
'             MsgBox "請先回前畫面輸入案源單號！", vbCritical
'             Exit Sub
'         End If
'         '判斷是否為補收文=>案源類別
'         strExc(1) = PUB_GetLOSplus(txtSystem, txtCode(0), txtCode(1), txtCode(2), txtOther(1), txtOther(3), t_LOSkind)
'         If m_LOS02 = "" And Left(strExc(1), 1) = "B" Then
'             If MsgBox("請確認接洽單沒有案源單號？" & vbCrLf & "選擇""是""會繼續作業，選擇""否""回前畫面輸入案源單號", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
'                 Exit Sub
'             End If
'         End If
'         If ((Left(t_LOSkind, 1) = "C" And txtCode(0) = "") Or (Left(strExc(1), 1) = "C" And txtCode(0) <> "")) And m_LOS15 = "" Then
'             If MsgBox("請確認接洽單沒有案源單號？" & vbCrLf & "選擇""是""會繼續作業，選擇""否""回前畫面輸入案源單號", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
'                 Exit Sub
'             End If
'         End If
'    End If
    'end 2020/06/05
     
     'Add By Sindy 2024/9/4 CF案申請國家不可為台灣
      If Left(txtSystem, 2) = "CF" And txtOther(3) = "000" Then
         MsgBox "CF案申請國家不可為台灣！", vbExclamation
         'Modify By Sindy 2025/4/8 +If txtOther(3).Enabled = True Then
         If txtOther(3).Enabled = True Then txtOther(3).SetFocus
         Call txtOther_GotFocus(3)
         Exit Sub
      End If
      '2024/9/4 END
     
     If frm010001.intModifyKind = 0 And strSrvDate(1) >= 法律所案源收文啟用日 And txtSystem = "TC" Then
           If txtOther(3) <> "000" Then '非台灣案=>清空資料
               m_LOS02 = ""
               m_LOS15 = ""
           Else
                t_LOSkind = PUB_GetLOSkind(txtSystem, txtOther(1), txtOther(3))
                If Left(t_LOSkind, 1) = "B" And m_LOS15 = "" Then
                    MsgBox "請先回前畫面輸入案源單號！", vbCritical, "檢核案源單號"
                    Exit Sub
                End If
                '判斷是否為補收文=>案源類別
                strExc(1) = PUB_GetLOSplus(txtSystem, txtCode(0), txtCode(1), txtCode(2), txtOther(1), txtOther(3), IIf(t_LOSkind = "", "C", t_LOSkind))
                If m_LOS02 = "" And strExc(1) <> "" And m_LOS15 = "" Then
                    If MsgBox("請確認接洽單左上角是否沒有案源單號？" & vbCrLf & "選擇""是""會繼續作業，選擇""否""回前畫面輸入案源單號", vbInformation + vbYesNo + vbDefaultButton1, "檢核案源單號") = vbNo Then
                        Exit Sub
                    End If
                End If
           End If
     End If
     'end 2020/06/10
     'Added by Lydia 2020/07/20 判斷法務案是否有案源 (非法律所的客戶)
     'Mark by Lydia 2022/11/03 法務案收文取消"是否林總同意本案由法律所自行收文？"的詢問，改為法律所自行收文智慧所人員客戶時，在判斷跨區收文發EMAIL給雙方智權人員時，加發林總。
'     If frm010001.intModifyKind = 0 And strSrvDate(1) >= 法律所案源收文啟用日 And InStr(txtSystem, "L") > 0 _
'          And m_LOS02 = "" And m_LOS15 = "" Then
'         'Added by Lydia 2020/08/31 排除台一
'         If txtOther(7) <> "" And InStr(txtOther(7), "X03072") > 0 Then
'         Else
'         'end 2020/08/31
'            'Modified by Lydia 2020/08/03 排除LA999999
'            'Modified by Lydia 2020/11/05 法務處P31
'            'If Trim(txtOther(7).Text) <> "" And Left(GetCuSales(ChangeCustomerL(txtOther(7))), 1) <> "L" And txtSystem & txtCode(0) <> "LA999999" Then
'            strExc(1) = GetCuSales(ChangeCustomerL(txtOther(7)))
'            If Trim(txtOther(7).Text) <> "" And Left(strExc(1), 1) <> "L" And strExc(1) <> "P31" And txtSystem & txtCode(0) <> "LA999999" Then
'            'end 2020/11/05
'                'Added by Lydia 2021/01/08 由法務人員直接收A2類案源的判斷(在存檔時會自動補上案源)
'                strExc(0) = ""
'                If txtCode(0) <> "" Then
'                    strSql = "Select * From Lawofficesource Where los02='A1' and los07||los08 is null and Los15 In " & _
'                                 "(select cp162 from caseprogress where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp162 is not null)  order by los12 desc "
'                    intI = 1
'                    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                    If intI = 1 Then
'                        strExc(0) = "A2"
'                    End If
'                End If
'                If strExc(0) = "" Then
'                'end 2021/01/08
'                  'Added by Morgan 2021/5/28 有林總同意法律所自行收文的例外 Ex:L-6396
'                  'Modified by Morgan 2022/5/19 國外部客戶除外 Ex:FCL-010967 --秀玲
'                  If Left(strExc(1), 1) <> "F" Then
'                     If MsgBox("是否林總同意本案由法律所自行收文？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
'                     'end 2021/5/28
'                       MsgBox "請回前畫面輸入案源單號！", vbCritical, "檢核案源單號"
'                       Exit Sub
'                     End If 'Added by Morgan 2021/5/28
'                  End If
'                  'end 2022/5/19
'                End If 'Added by Lydia 2021/01/08
'            End If
'         End If 'Added by Lydia 2020/08/31
'     End If
'     'end 2020/07/20
     'end 2022/11/03
     
   'Added by Lydia 2021/09/10 修正畫面所有含跳行符號的文字框; 9/10 FCT-47909收文申請,彼所案號中間有換行
   PUB_FilterFormText Me
      
   'Added by Lydia 2024/06/20 LA-999999收文有費用時，若收文日該月份已經有一筆有收費的進度，則顯示訊息不可收文。
   If frm010001.intModifyKind = 0 And txtSystem = "LA" And txtCode(0) = "999999" And Val(txtOther(12)) > 0 Then
      strSql = "select cp09,cp10 from caseprogress where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and cp04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' " & _
               "and cp159=0 and nvl(cp16,0) > 0 and cp05>=" & Mid(IIf(txtOther(0) = "", strSrvDate(1), DBDATE(txtOther(0))), 1, 6) & "01" & " and cp05<=" & Mid(IIf(txtOther(0) = "", strSrvDate(1), DBDATE(txtOther(0))), 1, 6) & "31"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         MsgBox "配合財務處報表，LA-999999每月只可有一筆有收費的進度，本月已收文過，不可再收費 !", vbCritical
         Exit Sub
      End If
   End If
   'end 2024/06/20
   
   'Added by Lydia 2024/12/13 FG案輸入追蹤流水號：檢查命名追蹤檔案
   bolMoveCheck = False
   If Trim(txtSystem) = "FG" And FraTCN.Visible = True And Trim(txtTCN01) <> "" Then
      If PUB_ChkTCNfileExist(txtTCN01.Text) = True Then
         bolMoveCheck = True
      Else
         strExc(1) = ""
         If Pub_StrUserSt03 = "M51" Or InStr(UCase(Forms(0).Caption), "M51") > 0 Then
            If MsgBox("測試新案立卷是否要搬檔案？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
               strExc(1) = "Y"
            End If
         Else
            strExc(1) = "Y"
         End If
         If strExc(1) = "Y" Then
            MsgBox "請先聯絡 " & IIf(lblSales.Caption <> "", lblSales.Caption, "外專承辦人員") & vbCrLf & "上傳TRACKING_NO檔案！", vbCritical + vbOKOnly, "TRACKING_NO檔案稽核"
            Screen.MousePointer = varSaveCursor
            GoTo ErrorHandler
         End If
      End If
   End If
   'end 2024/12/13
   
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   For i = 0 To 20
      '93.7.5 Add By Sonia
      If i = 9 Then
         If txtOther(9) <> "" Then
            If CheckIsTaiwanDate(txtOther(9).Text) Then
               If CheckReKey(txtOther(9)) Then
                  If Val(txtOther(9)) = Val(GetTaiwanTodayDate) Then
                     ShowMsg "此案件已屆本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
                  End If
                  If Val(txtOther(9)) < Val(GetTaiwanTodayDate) Then
                     ShowMsg "此案件已逾本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
                  End If
               Else
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
            End If
         End If
      End If
      '93.7.5 end
      'Add By Sindy 2010/12/31 費用檢查提到存檔前檢查
      If i = 12 Then '費用檢查
         '郭 請作單 X14843050 不管
         'Modify By Sindy 2011/1/18 增加當事人4,5檢查
         'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
         'modify by sonia 2014/9/11 取消X69514,已轉外專
         If Mid(txtOther(7), 1, 8) <> "X1484305" And Mid(txtOther(18), 1, 8) <> "X1484305" And Mid(txtOther(19), 1, 8) <> "X1484305" And Mid(txtOther(29), 1, 8) <> "X1484305" And Mid(txtOther(30), 1, 8) <> "X1484305" And _
            Mid(txtOther(7), 1, 8) <> "X3928904" And Mid(txtOther(18), 1, 8) <> "X3928904" And Mid(txtOther(19), 1, 8) <> "X3928904" And Mid(txtOther(29), 1, 8) <> "X3928904" And Mid(txtOther(30), 1, 8) <> "X3928904" Then
            If ClsPDGetCaseLowPrice(txtSystem, txtOther(3), txtOther(1), douStPrice, douLowPrice) = 1 Then
            End If
            
            'Added by Lydia 2020/05/20 法律所案源收文：(著作權)台灣案之B1及B2案件性質都不可收費用。
            If txtOther(3) = "000" And txtSystem = "TC" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
                If Val(txtOther(12)) > 0 Then
                    MsgBox "【B類】案源接洽單之費用、規費都必須為 0，點數必須為空白。", vbExclamation, "檢核案源單號"
                    Screen.MousePointer = varSaveCursor
                    Exit Sub
                End If
            Else
            'end 2020/06/05
                'Added by Lydia 2020/06/10 法律所案源收文
                If Val(txtOther(12)) > 0 And txtOther(3) = "000" And InStr(txtSystem, "L") > 0 And m_LOS02 <> "" And Left(m_LOS02, 1) = "C" Then
                     MsgBox "【C類】案源法務案之費用、規費都必須為 0，點數必須為空白。", vbExclamation, "檢核案源單號"
                     Screen.MousePointer = varSaveCursor
                     Exit Sub
                End If
                'end 2020/06/10
                'MODIFY BY SONIA 2014/7/17 +傳規費 CFP-027024
                If ClsPDGetCaseFee(txtSystem, txtOther(3), txtOther(1), Val(txtOther(12)), Val(txtOther(16))) = 0 Then
                   Screen.MousePointer = varSaveCursor
                   Exit Sub
                End If
            End If 'Added by Lydia 2020/06/05
         End If
      End If
      If i = 13 Then  '點數檢查
         'Add By Sindy 2010/12/31 點數檢查提到存檔前檢查
         '郭 請作單 X14843050 不管
         'Modify By Sindy 2011/1/18 增加當事人4,5檢查
         'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
         'modify by sonia 2014/9/11 取消X69514,已轉外專
         If Mid(txtOther(7), 1, 8) <> "X1484305" And Mid(txtOther(18), 1, 8) <> "X1484305" And Mid(txtOther(19), 1, 8) <> "X1484305" And Mid(txtOther(29), 1, 8) <> "X1484305" And Mid(txtOther(30), 1, 8) <> "X1484305" And _
            Mid(txtOther(7), 1, 8) <> "X3928904" And Mid(txtOther(18), 1, 8) <> "X3928904" And Mid(txtOther(19), 1, 8) <> "X3928904" And Mid(txtOther(29), 1, 8) <> "X3928904" And Mid(txtOther(30), 1, 8) <> "X3928904" Then
            If txtOther(13) = "" Then
               If txtOther(12) <> "" And txtOther(16) <> "" Then
                  ShowMsg MsgText(1035)
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
            ElseIf txtOther(12) <> "" Then
'               If Format((Val(txtOther(12)) - Val(txtOther(16))) / 1000, "0.0") <> Format(Val(txtOther(13)), "0.0") Then
'                  ShowMsg MsgText(1036)
'                  Screen.MousePointer = varSaveCursor
'                  Exit Sub
'               End If
            Else
               ShowMsg MsgText(1037)
               Screen.MousePointer = varSaveCursor
               Exit Sub
            End If
         End If
        'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
        'Modify By Sindy 2011/1/18 增加當事人4,5檢查
        'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
        'modify by sonia 2014/9/11 取消X69514,已轉外專
        If Mid(txtOther(7), 1, 8) <> "X1484305" And Mid(txtOther(18), 1, 8) <> "X1484305" And Mid(txtOther(19), 1, 8) <> "X1484305" And Mid(txtOther(29), 1, 8) <> "X1484305" And Mid(txtOther(30), 1, 8) <> "X1484305" And _
           Mid(txtOther(7), 1, 8) <> "X3928904" And Mid(txtOther(18), 1, 8) <> "X3928904" And Mid(txtOther(19), 1, 8) <> "X3928904" And Mid(txtOther(29), 1, 8) <> "X3928904" And Mid(txtOther(30), 1, 8) <> "X3928904" Then
            If Me.txtOther(12) <> "" Or Me.txtOther(16) <> "" Then
               If Format((Val(Me.txtOther(12)) - Val(Me.txtOther(16))) / 1000, "0.0") <> Format(Val(Me.txtOther(13)), "0.0") Then
                  ShowMsg MsgText(1036)
                  Screen.MousePointer = varSaveCursor
'                  txtOther(i).SetFocus
'                  txtOther_GotFocus (i)
                  Exit Sub
               End If
            End If
        End If
      End If
      'Add By Sindy 2010/12/31 規費檢查提到存檔前檢查
      If i = 16 Then '規費檢查
         If Val(txtOther(16)) > 0 Or txtOther(3) = "000" Then     'ADD by sonia 2014/7/17 加入未輸規費時不檢查此段,因為可能是依代理人帳單請款CFP-027024
            '郭 請作單 X14843050 不管
            'Modify By Sindy 2011/1/18 增加當事人4,5檢查
            'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
            'modify by sonia 2014/9/11 取消X69514,已轉外專
            If Mid(txtOther(7), 1, 8) <> "X1484305" And Mid(txtOther(18), 1, 8) <> "X1484305" And Mid(txtOther(19), 1, 8) <> "X1484305" And Mid(txtOther(29), 1, 8) <> "X1484305" And Mid(txtOther(30), 1, 8) <> "X1484305" And _
               Mid(txtOther(7), 1, 8) <> "X3928904" And Mid(txtOther(18), 1, 8) <> "X3928904" And Mid(txtOther(19), 1, 8) <> "X3928904" And Mid(txtOther(29), 1, 8) <> "X3928904" And Mid(txtOther(30), 1, 8) <> "X3928904" Then
                'Added by Lydia 2020/05/20 法律所案源收文：(著作權)台灣案之B1及B2案件性質都不可收費用。
                If Val(txtOther(16)) > 0 And txtOther(3) = "000" And txtSystem = "TC" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
                    MsgBox "【B類】案源接洽單之費用、規費都必須為 0，點數必須為空白。", vbExclamation, "檢核案源單號"
                    Screen.MousePointer = varSaveCursor
                    Exit Sub
                End If
                'end 2020/06/05
               'Added by Lydia 2020/06/10 法律所案源收文
               If Val(txtOther(16)) > 0 And txtOther(3) = "000" And InStr(txtSystem, "L") > 0 And m_LOS02 <> "" And Left(m_LOS02, 1) = "C" Then
                    MsgBox "【C類】案源法務案之費用、規費都必須為 0，點數必須為空白。", vbExclamation, "檢核案源單號"
                    Screen.MousePointer = varSaveCursor
                    Exit Sub
               End If
               'end 2020/06/10
               If GetCaseMoney(txtSystem, txtOther(3), txtOther(1), Val(txtOther(16))) = 0 Then
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
            End If
         End If 'ADD by sonia 2014/7/17 加入未輸規費時不檢查上面這段,因為可能是依代理人帳單請款CFP-027024
      End If
      
      'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
      'Modify By Sindy 2011/1/18 增加當事人4,5檢查
      'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
      'modify by sonia 2014/9/11 取消X69514,已轉外專
      If Mid(txtOther(7), 1, 8) <> "X1484305" And Mid(txtOther(18), 1, 8) <> "X1484305" And Mid(txtOther(19), 1, 8) <> "X1484305" And Mid(txtOther(29), 1, 8) <> "X1484305" And Mid(txtOther(30), 1, 8) <> "X1484305" And _
         Mid(txtOther(7), 1, 8) <> "X3928904" And Mid(txtOther(18), 1, 8) <> "X3928904" And Mid(txtOther(19), 1, 8) <> "X3928904" And Mid(txtOther(29), 1, 8) <> "X3928904" And Mid(txtOther(30), 1, 8) <> "X3928904" Then
        If Val(txtOther(12)) = 0 And Val(txtOther(16)) = 0 And Val(txtOther(13)) = 0 Then
           If i = 12 Or i = 16 Or i = 13 Then
              GoTo GoToNext
           End If
        End If
      End If
      'Add By Sindy 2010/5/26 檢查申請人/當事人的輸入順序
      'Modify By Sindy 2011/1/18 增加當事人4,5檢查
      If (Trim(txtOther(18)) <> "" And Trim(txtOther(7)) = "") Or _
         (Trim(txtOther(19)) <> "" And Trim(txtOther(18)) = "") Or _
         (Trim(txtOther(29)) <> "" And Trim(txtOther(19)) = "") Or _
         (Trim(txtOther(30)) <> "" And Trim(txtOther(29)) = "") Then
         ShowMsg "請依序輸入申請人/當事人!"
         If Trim(txtOther(18)) <> "" And Trim(txtOther(7)) = "" Then txtOther(18).SetFocus: Call txtOther_GotFocus(18): Exit For
         If Trim(txtOther(19)) <> "" And Trim(txtOther(18)) = "" Then txtOther(19).SetFocus: Call txtOther_GotFocus(19): Exit For
         If Trim(txtOther(29)) <> "" And Trim(txtOther(19)) = "" Then txtOther(29).SetFocus: Call txtOther_GotFocus(29): Exit For
         If Trim(txtOther(30)) <> "" And Trim(txtOther(29)) = "" Then txtOther(30).SetFocus: Call txtOther_GotFocus(30): Exit For
      End If
      '2010/5/26 End
      'Modify By Cheng 2001/12/27
      If i = 7 Then
         If txtOther(7) = "" And txtOther(8) = "" Then
            ShowMsg "申請人或代理人不可同時空白!"
            txtOther(7).SetFocus
            txtOther_GotFocus (7)
            Exit For
        End If
      End If
      'modify by sonia 2017/1/23 +第二~五申請人
      If i = 7 Or i = 18 Or i = 19 Or i = 29 Or i = 30 Then
         If Len(Trim(Me.txtOther(i).Text)) > 0 Then
            If CheckKeyIn(i) <> 1 Then
               If txtOther(i).Enabled Then
                  txtOther(i).SetFocus
                  txtOther_GotFocus (i)
               End If
               Exit For
            End If
         End If
      ElseIf i = 8 Then
         If Len(Trim(Me.txtOther(i).Text)) > 0 Then
            If CheckKeyIn(i) <> 1 Then
               If txtOther(i).Enabled Then
                  txtOther(i).SetFocus
                  txtOther_GotFocus (i)
               End If
               Exit For
            End If
         End If
      ElseIf txtOther(i).Enabled And txtOther(i).Visible Then
         If CheckKeyIn(i) <> 1 Then
            If txtOther(i).Enabled Then
               txtOther(i).SetFocus
               txtOther_GotFocus (i)
            End If
            Exit For
         End If
      End If
GoToNext:
   Next
   If i = 21 Then
      'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
      'Modify By Sindy 2011/1/18 增加當事人4,5檢查
      'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
      'modify by sonia 2014/9/11 取消X69514,已轉外專
      If Mid(txtOther(7), 1, 8) <> "X1484305" And Mid(txtOther(18), 1, 8) <> "X1484305" And Mid(txtOther(19), 1, 8) <> "X1484305" And Mid(txtOther(29), 1, 8) <> "X1484305" And Mid(txtOther(30), 1, 8) <> "X1484305" And _
         Mid(txtOther(7), 1, 8) <> "X3928904" And Mid(txtOther(18), 1, 8) <> "X3928904" And Mid(txtOther(19), 1, 8) <> "X3928904" And Mid(txtOther(29), 1, 8) <> "X3928904" And Mid(txtOther(30), 1, 8) <> "X3928904" Then
         
         If Val(txtOther(16)) > 0 Or txtOther(3) = "000" Then     'ADD by sonia 2014/7/17 加入未輸規費時不檢查此段,因為可能是依代理人帳單請款CFP-027024
            'Add By Cheng 2003/11/20
            '檢查規費欄位
            'edit by nickc 2006/12/05 change call basquery
            'If objPublicData.GetCaseMoney(txtSystem, txtOther(3), txtOther(1), Val(txtOther(16))) <> 1 Then
            If GetCaseMoney(txtSystem, txtOther(3), txtOther(1), Val(txtOther(16))) <> 1 Then
                Screen.MousePointer = varSaveCursor
                Me.txtOther(16).SetFocus
                txtOther_GotFocus (16)
                Exit Sub
            End If
         End If 'ADD by sonia 2014/7/17 加入未輸規費時不檢查上面這段,因為可能是依代理人帳單請款CFP-027024
         'Add By Cheng 2003/08/28
         '檢查點數是否低於底價
         If ChkPointValue(Me.txtSystem.Text, Me.txtOther(3).Text, Me.txtOther(1).Text, Me.txtOther(13).Text, Me.txtOther(10).Text) = False Then
             Screen.MousePointer = varSaveCursor
             Me.txtOther(13).SetFocus
             txtOther_GotFocus (13)
             Exit Sub
         End If
         
         'Added by Lydia 2020/06/01 法律所案源收文：若案源案件類型為B2，則要檢查案源總收文號LOS01之案件性質之費用及規費是否符合CASEFEE的設定
         'Modified by Lydia 2020/06/05 排除著作權TC =>And InStr(txtSystem, "L") > 0
         'Modified by Lydia 2022/10/06 判斷案源若為FCP或FCT時，因為向國外請款是以FCP或FCT案號請款，所以FCL案號就不檢查費用和規費; ex.FCP-065025的AB1040290和FCL-109735的AB1040329
         'If frm010001.intModifyKind = 0 And strSrvDate(1) >= 法律所案源收文啟用日 And InStr(txtSystem, "L") > 0 And m_LOS15 <> "" And Left(m_LOS02, 2) = "B2" Then
         If frm010001.intModifyKind = 0 And strSrvDate(1) >= 法律所案源收文啟用日 And InStr(txtSystem, "L") > 0 And m_LOS15 <> "" _
             And Left(m_LOS02, 2) = "B2" And m_LOS01cp01 <> "FCP" And m_LOS01cp01 <> "FCT" Then
             'Modify By Sindy 2025/4/14 改成共用函數
             'Call GetTotalFeeLOS01(m_LOS01, strExc(1), strExc(2), strExc(3))
             Call PUB_GetCaseFee_000(m_LOS01, strExc(1), strExc(2), strExc(3))
             '2025/4/14 END
             'Modify By Sindy 2025/10/27 mark;已有Flow簽核不需檢查
'             If Val(strExc(1)) > 0 And Val(txtOther(12)) < Val(strExc(1)) Then
'                  MsgBox "費用不可少於" & strExc(1) & "(P/T案費用)！", vbExclamation
'                  Screen.MousePointer = varSaveCursor
'                  Me.txtOther(13).SetFocus
'                  txtOther_GotFocus (13)
'                  Exit Sub
'             End If
             '2025/10/27 END
             If Val(strExc(3)) > 0 And Val(txtOther(16)) < Val(strExc(3)) Then
                  MsgBox "規費不可少於" & strExc(3) & "(P/T案規費)！", vbExclamation
                  Screen.MousePointer = varSaveCursor
                  Me.txtOther(16).SetFocus
                  txtOther_GotFocus (16)
                  Exit Sub
             End If
         End If
         'end 2020/06/01
      End If
      
      'add by nickc 2006/11/30 查名時，類別一定要輸
      If txtOther(25).Visible = True And txtOther(25).Enabled = True And txtOther(1) = "001" Then
          If Trim(txtOther(25)) = "" Then
            'Modify By Sindy 2017/3/28 S案在櫃台收文時會控管「類別」欄必須輸入，
            '但有些案件無法指定類別, 故請取消控管, 並在收文及分案時改以提醒方式
            If MsgBox("查名，是否有商品類別要輸入？", vbExclamation + vbYesNo + vbDefaultButton1, "注意！") = vbYes Then
               Screen.MousePointer = varSaveCursor
               Me.txtOther(25).SetFocus
               txtOther_GotFocus (25)
               Exit Sub
            End If
            '2017/3/28 END
'              Screen.MousePointer = varSaveCursor
'              MsgBox "查名必須要有商品類別！", , "注意！"
'              Me.txtOther(25).SetFocus
'              txtOther_GotFocus (25)
'              Exit Sub
          End If
      End If
      
      'Add By Sindy 2011/7/26 L或LA新案時, 收文程式檢查
      If (txtSystem = "L" Or txtSystem = "LA") And txtCode(0) = "" Then
         '申請人1為X65299謝律師, 未輸入申請人2時
         If Left(Trim(txtOther(7)), 6) = "X65299" And Trim(txtOther(18)) = "" Then
            Screen.MousePointer = varSaveCursor
            MsgBox "請輸入申請人/當事人2, 若接洽單未填寫請智權人員補填實際客戶資料！", vbExclamation + vbOKOnly
            txtOther(18).SetFocus
            Call txtOther_GotFocus(18)
            Exit Sub
         End If
         '申請人2~5為X65299謝律師時
         If Left(Trim(txtOther(18)), 6) = "X65299" Or Left(Trim(txtOther(19)), 6) = "X65299" Or _
            Left(Trim(txtOther(29)), 6) = "X65299" Or Left(Trim(txtOther(30)), 6) = "X65299" Then
            Screen.MousePointer = varSaveCursor
            MsgBox "與謝律師合作案件請於申請人1輸入X65299謝智硯律師事務所, 申請人2欄填實際客戶資料！", vbExclamation + vbOKOnly
            If Left(Trim(txtOther(18)), 6) = "X65299" Then txtOther(18).SetFocus: Call txtOther_GotFocus(18)
            If Left(Trim(txtOther(19)), 6) = "X65299" Then txtOther(19).SetFocus: Call txtOther_GotFocus(19)
            If Left(Trim(txtOther(29)), 6) = "X65299" Then txtOther(29).SetFocus: Call txtOther_GotFocus(29)
            If Left(Trim(txtOther(30)), 6) = "X65299" Then txtOther(30).SetFocus: Call txtOther_GotFocus(30)
            Exit Sub
         End If
      End If
      '2011/7/26 End
      
      strAuto1 = txtRecieveCode
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
      
      'add by nickc 2007/11/12 加入檢查特殊客戶
      Dim IsSpecCu As Boolean
      IsSpecCu = False
      If IsSpecCu = False And txtOther(7) <> "" Then
          strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtOther(7)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtOther(7)), 9, 1) & "' "
          CheckOC3
          AdoRecordSet3.CursorLocation = adUseClient
          AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          If AdoRecordSet3.RecordCount <> 0 Then
              If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                  IsSpecCu = True
              End If
          End If
      End If
      If IsSpecCu = False And txtOther(18) <> "" Then
          strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtOther(18)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtOther(18)), 9, 1) & "' "
          CheckOC3
          AdoRecordSet3.CursorLocation = adUseClient
          AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          If AdoRecordSet3.RecordCount <> 0 Then
              If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                  IsSpecCu = True
              End If
          End If
      End If
      If IsSpecCu = False And txtOther(19) <> "" Then
          strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtOther(19)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtOther(19)), 9, 1) & "' "
          CheckOC3
          AdoRecordSet3.CursorLocation = adUseClient
          AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          If AdoRecordSet3.RecordCount <> 0 Then
              If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                  IsSpecCu = True
              End If
          End If
      End If
      'Add By Sindy 2011/1/18 增加當事人4,5檢查
      If IsSpecCu = False And txtOther(29) <> "" Then
          strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtOther(29)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtOther(29)), 9, 1) & "' "
          CheckOC3
          AdoRecordSet3.CursorLocation = adUseClient
          AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          If AdoRecordSet3.RecordCount <> 0 Then
              If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                  IsSpecCu = True
              End If
          End If
      End If
      If IsSpecCu = False And txtOther(30) <> "" Then
          strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtOther(30)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtOther(30)), 9, 1) & "' "
          CheckOC3
          AdoRecordSet3.CursorLocation = adUseClient
          AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          If AdoRecordSet3.RecordCount <> 0 Then
              If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                  IsSpecCu = True
              End If
          End If
      End If
      '2011/1/18 End
      If IsSpecCu Then
         'Modify By Sindy 2023/1/30 排除有輸入案源編號者,已有Flow簽核不需檢查
         If m_LOS15 = "" Then
         '2023/1/30 END
            If MsgBox("請確認此客戶接洽單主管是否核示??", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                Exit Sub
            End If
         End If
      End If
      
      '2011/3/25 add by sonia
      'Modified by Lydia 2019/02/14
      'strSql = "select st15 from staff where st01='" & txtOther(10) & "'"
      'intI = 1
      'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      'If intI = 1 Then
      '   '國外部收文台灣案必須收FG或S案號
      '   If Left(Trim(RsTemp.Fields("st15")), 1) = "F" And txtOther(3) = "000" And txtSystem <> "FG" And txtSystem <> "S" And txtSystem <> "FCL" And txtSystem <> "LIN" Then
          If Left(m_SalesST15, 1) = "F" And txtOther(3) = "000" And txtSystem <> "FG" And txtSystem <> "S" And txtSystem <> "FCL" And txtSystem <> "LIN" Then
            If txtSystem = "PS" Or txtSystem = "CPS" Or txtSystem = "TS" Or txtSystem = "L" Then    '2011/12/23 add by sonia 陳經理收TD-132無法收文
               '2015/4/14 MODIFY BY SONIA 林總指示開放投資法務人員可收文L案及P案
               'MsgBox "國外部台灣案必須收 FG,S,FCL,LIN 案號 !!!", vbExclamation + vbOKOnly
               'Screen.MousePointer = varSaveCursor
               'Exit Sub
               If MsgBox("國外部台灣案必須收 FG,S,FCL,LIN 案號, 是否修改系統類別？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
            End If       '2011/12/23 add by sonia
         End If
      'End If 'Mark by Lydia 2019/02/14
      '2011/3/25 end
         
      'add by nickc 2007/03/27 非台灣要詢問
      '2009/11/25 MODIFY BY SONIA 新案才要詢問
      'If GetPrjNationNumber1(ChangeCustomerL(txtOther(7))) > "010" Then
      '2010/10/20 modify by sonia 非智權部收文才要問 CFP-023621
      'If GetPrjNationNumber1(ChangeCustomerL(txtOther(7))) > "010" And txtCode(0) = "" Then
      'Modified by Lydia 2019/02/14
      'If GetPrjNationNumber1(ChangeCustomerL(txtOther(7))) > "010" And txtCode(0) = "" And Left(tST15, 1) <> "S" Then
      If GetPrjNationNumber1(ChangeCustomerL(txtOther(7))) > "010" And txtCode(0) = "" And Left(m_SalesST15, 1) <> "S" Then
            If txtOther(8) = "" Then
                If MsgBox("請確定  無代理人   !!", vbYesNo, "警告！") = vbNo Then
                    Screen.MousePointer = varSaveCursor
                    txtOther(8).SetFocus
                    txtOther_GotFocus (8)
                    Exit Sub
                End If
            'Modify by Amy 2017/01/03 從下面搬上來,上面訊息若選擇"是",就不要再詢問下列訊息-秀玲
            ElseIf txtOther(27) = "" Then
                If MsgBox("請確定  無代理人彼所案號  !!", vbYesNo, "警告！") = vbNo Then
                    Screen.MousePointer = varSaveCursor
                    txtOther(27).SetFocus
                    txtOther_GotFocus (27)
                    Exit Sub
                End If
            End If
      End If
      'Add By Sindy 2010/3/19
      If Left(Trim(GetStaffDepartment(txtOther(10).Text)), 2) = "F2" And _
         frm010001.intSaveMode = "1" And _
         strSP30s = "" And strSP75s = "" Then
         If MsgBox("是否輸入國外聯絡人資料?", vbExclamation + vbOKCancel) = vbOK Then
            Screen.MousePointer = vbDefault
            Call Command1_Click
            Exit Sub
         End If
      End If
      '2010/3/19 End
      
      '2011/4/21 add by sonia
      Dim strSP78 As String, strContact As String
      If cboContact.Locked = False Then
         strContact = ""
         If cboContact.ListCount > 2 Then
            'Modify by Amy 2021/12/21 改成Form 2.0
            'strSP78 = Format(cboContact.ItemData(cboContact.ListIndex), "00")
            strSP78 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
            PUB_GetContact strAppNo1, strContact, True
            If strSP78 = strContact Or strSP78 = "00" Then
               If MsgBox("請確定接洽人欄是否有為★, 是否要選擇其他接洽人!!", vbYesNo, "警告！") = vbYes Then
                   Screen.MousePointer = varSaveCursor
                   cboContact.SetFocus
                   Exit Sub
               End If
            End If
         End If
      End If
      '2011/4/21 end
      
      'Add By Sindy 2011/6/3
      'Modified by Lydia 2018/07/09 +L
      'If txtSystem = "LA" And txtOther(31).Enabled = True And Trim(txtOther(31)) = "" Then
      If (txtSystem = "LA" Or txtSystem = "L") And txtOther(31).Enabled = True And Trim(txtOther(31)) = "" Then
         If MsgBox("請輸入接洽記錄單上之主題，是否需要輸入？", vbYesNo, "警告！") = vbYes Then
            Screen.MousePointer = varSaveCursor
            txtOther(31).SetFocus
            Exit Sub
         End If
      End If
      '2011/6/3 End
      
      'Added by Lydia 2019/05/13 改模組(一併取得)
      If Left(m_SalesST15, 1) <> "F" And txtOther(7).Text <> "" And Val(txtOther(12)) > 0 Then
          'Modified by Lydia 2022/06/13 傳入收文之本所案號,案件性質(可用,串接)
          'Call PUB_GetBillDataAll("3", txtOther(7), dblAmt, dblPFee, dblTFee, , , TransDate(txtOther(0), 2), mBillNo, mMemo)
          'Modified by Lydia 2022/06/15 傳入收文之智權人員
          Call PUB_GetBillDataAll("3", txtOther(7), txtSystem & IIf(txtCode(0) <> "", txtCode(0) & Left(txtCode(1) & "0", 1) & Left(txtCode(2) & "00", 2), ""), txtOther(1), Trim(txtOther(10)), dblAmt, dblPFee, dblTFee, , , TransDate(txtOther(0), 2), mBillNo, mMemo)
      End If
      
      'Add By Sindy 2012/11/06 非T*案件(TF要含)若已送件之應收款超過15萬以上,智權人員非國外部且有費用者須做下列控管
      'Modified by Lydia 2017/06/19 +判斷有申請人編號
      'If (Left(Trim(txtSystem), 1) <> "T" Or Trim(txtSystem) = "TF") And _
         Left(tST15, 1) <> "F" And _
         Val(txtOther(12)) > 0 And _
         Check2.Value = 0 Then
      'Modified by Lydia 2019/02/14
      'If (Left(Trim(txtSystem), 1) <> "T" Or Trim(txtSystem) = "TF") And Left(tST15, 1) <> "F" And Val(txtOther(12)) > 0 And Check2.Value = 0 And Trim(txtOther(7)) <> "" Then
      If (Left(Trim(txtSystem), 1) <> "T" Or Trim(txtSystem) = "TF") And Left(m_SalesST15, 1) <> "F" And Val(txtOther(12)) > 0 And Check2.Value = 0 And Trim(txtOther(7)) <> "" Then
      'end 2017/06/19
         'Mark by Lydia 2019/05/13 改模組(一併取得)
         'GetBillData txtOther(7), dblAmt, dblPFee, dblTFee
         
         'Add By Sindy 2012/12/10 取得客戶應收帳款收文檢查上限
         'Modified by Lydia 2020/02/03 應收帳款上限分開管制為個人"應收帳款上限"和"集團應收帳款上限"
         'dblChkAmt = PUB_GetCustRecAmtLmt(txtOther(7))
         ''2012/12/10 End
         dblCu183 = PUB_GetCustRecAmtLmt(txtOther(7), dblChkAmt)
         'Added by Lydia 2020/02/03 判斷是否有集團上限
         If dblChkAmt = 0 Then
             dblAmtR = 0: dblPFeeR = 0: dblTFeeR = 0
         Else   '有集團上限才抓關係企業的應收帳款金額
             GetBillData txtOther(7), dblAmtR, dblPFeeR, dblTFeeR
         End If
         'end 2020/02/03
            
         '已送件之應收款超過30萬以上(不含T*案件應收款),提醒
         'Modify By Sindy 2012/12/10 檢查的30萬改不要固定金額,抓CustRecAmtLmt
         'If dblAmt >= 300000 Then
         'Modified by Lydia 2020/02/03 應收帳款上限分開管制為個人"應收帳款上限"和"集團應收帳款上限"
         'If dblAmt >= dblChkAmt Then
         ''2012/12/10 End
         'Modified by Lydia 2022/06/15 排除特定人員
         'Modified by Lydia 2022/09/06 改抓特殊設定
         'If InStr(cnt應收帳款檢查排除, Trim(txtOther(10))) = 0 And dblAmt >= dblCu183 Or (dblAmtR >= dblChkAmt And dblChkAmt > 0) Then
         'Modified by Lydia 2022/09/21 案源要判斷介紹人是否在應收帳款上限檢查排除名單內
         'If InStr(Pub_GetSpecMan("應收帳款上限檢查排除"), Trim(txtOther(10))) = 0 And dblAmt >= dblCu183 Or (dblAmtR >= dblChkAmt And dblChkAmt > 0) Then
         strExc(7) = InStr(Pub_GetSpecMan("應收帳款上限檢查排除"), IIf(m_LOS04_1 <> "", m_LOS04_1, Trim(txtOther(10))))
         If Val(strExc(7)) = 0 And dblAmt >= dblCu183 Or (dblAmtR >= dblChkAmt And dblChkAmt > 0) Then
         'end 2022/09/21
            'Modified by Lydia 2018/09/20 預設按鈕改成"否" vbDefaultButton1=>vbDefaultButton2
            'Modify By Sindy 2023/1/30 排除有輸入案源編號者,已有Flow簽核不需檢查
            If m_LOS15 = "" Then
            '2023/1/30 END
               If MsgBox("請注意接洽單上是否有註明應收帳款超額，需主管簽核才可收文！是否可收文？" & vbCrLf & _
                         "（接洽單上若有★★的應收帳款簽核控管，是否已勾選畫面上的註記欄位了？）", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
            End If
'         '已送件之應收款超過15萬以上(不含T*案件應收款),提醒
'         ElseIf dblAmt >= 150000 Then
'            If MsgBox("請注意接洽單上是否有註明應收帳款超額，需主管簽核才可收文！是否可收文？" & vbCrLf & _
'                      "（接洽單上若有★★的應收帳款簽核控管，是否已勾選畫面上的註記欄位了？）", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
'               Screen.MousePointer = varSaveCursor
'               Exit Sub
'            End If
         End If
      End If
      '2012/11/06 End
      
      'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
      'Modified by Lydia 2019/04/08 智權人員非國外部
      'If txtOther(7).Text <> "" And Val(txtOther(12)) > 0 Then
      If Left(m_SalesST15, 1) <> "F" And txtOther(7).Text <> "" And Val(txtOther(12)) > 0 Then
         'Modified by Lydia 2019/05/13 改模組(一併取得)
         'If GetBillDate(txtOther(7), TransDate(txtOther(0), 2), strExc(1), strExc(2)) = True Then
         If mMemo <> "" Then
            'Modified by Lydia 2018/10/29 改訊息
            'If MsgBox("請注意接洽單上是否有註明" & vbCrLf & strExc(2) & vbCrLf & "，請交主管簽核並且有主管簽核。" & vbCrLf & "是否可收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
            'Modify By Sindy 2023/1/30 排除有輸入案源編號者,已有Flow簽核不需檢查
            If m_LOS15 = "" Then
            '2023/1/30 END
               If MsgBox("請注意接洽單上是否有註明" & vbCrLf & mMemo & "，請交主管簽核。" & vbCrLf & "並且有主管簽核，是否可收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
            End If
         End If
      End If
      'end 2018/08/22
            
      'Add by Lydia 2014/10/31 開放外專程序人員可進入專利處系統操作FMP寰華案件=>收文時詢問CP44是否設為Y53374000
       mFMPchk = False
       If Trim(txtSystem) = "PS" And (mCP31 = "Y" Or frm010001.intSaveMode = 1) Then '新增-無案號,修改-讀CP31
          'Modified by Lydia 2019/02/14
          'If Left(tST15, 1) = "F" Then
          If Left(m_SalesST15, 1) = "F" Then
             If MsgBox("請確認是否為寰華案件？", vbOKCancel) = 1 Then mFMPchk = True
           End If
       End If
       'end 2014/10/31
       
       'Added by Lydia 2016/04/25 提示輸入查名代號
       'Modified by Lydia 2016/04/27 改成直接在畫面輸入查名代號
       'If cmdTSMap.Visible = True And TMQList = "" And bolOpen130 = False Then
       'Modified by Lydia 2016/05/09 +台灣案
       If txtOther(3) = "000" And lblTS.Visible = True And txtTS(0).Visible = True And txtTS(1).Enabled = True Then
          'If MsgBox("查名代號應輸入?", vbInformation + vbYesNo, "輸入查名代號") = vbYes Then
          '   Screen.MousePointer = varSaveCursor
          '   Call cmdTSMap_Click
            If Len(txtTS(0) & txtTS(1)) <> 9 Then
                If MsgBox("查名代號應輸入?", vbInformation + vbYesNo, "輸入查名代號") = vbYes Then
                   Screen.MousePointer = varSaveCursor
                   txtTS(1).SetFocus
                   Exit Sub
                End If
            Else
                'Modified dy Lydia 2019/08/23 +智權人員的部門
                'If PUB_TQCtoTMQ(txtOther(10).Text, txtTS(0).Text & txtTS(1).Text, TMQList) = False Then
                'Modified by Lydia 2024/03/14 +True
                If PUB_TQCtoTMQ(True, m_SalesST15, txtOther(10).Text, txtTS(0).Text & txtTS(1).Text, TMQList) = False Then
                    Screen.MousePointer = varSaveCursor
                    txtTS(1).SetFocus
                    Exit Sub
                End If
                TMQList = txtTS(0).Text & txtTS(1).Text & "|" & TMQList
            End If
            'end 2016/04/27
       End If
       'end 2016/04/25
      
      'Add By Sindy 2022/8/17
      If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
         If PUB_ChkFileOpening2(m_PrevForm.m_strFullFileName, "後續才能一併歸卷！") = True Then
            Screen.MousePointer = varSaveCursor
            GoTo ErrorHandler
         End If
      End If
      '2022/8/17 END
      
      'Modified by Lydia 2022/09/14 判斷啟用日
      'If SaveDatabase(strAuto1, strAuto2) Then
      bolSaveOK = False
      If strSrvDate(1) < 收文存檔模組化啟用日 Then
           bolSaveOK = SaveDatabase(strAuto1, strAuto2)
           'Added by Lydia 2016/04/25 查名單對應存檔
           If TMQList <> "" Then
               strExc(1) = Mid(TMQList, 1, InStr(TMQList, "|") - 1)
               strExc(2) = Mid(TMQList, InStr(TMQList, "|") + 1)
               'Modified by Lydia 2024/03/14 +False
               If PUB_TMQtoCP(False, m_AttachPath, strAuto1, strExc(2), strExc(1)) = False Then
               End If
           End If
           'end 2016/04/25
      Else
           Call SetDBArray(False, txtRecieveCode, txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
           '(已包含)查名單對應存檔
           bolSaveOK = PUB_SaveFrm010007(Me.Name, frm010001.intSaveMode, frm010001.intModifyKind, frm010001.intCaseKind, frm010001.intChoose, _
                              modBase, modCP, txtOther(11), mChkStr, m_strControl, IsSaveData, mType, mCaseNo, mRetVal)
                              
           If frm010001.intModifyKind = 0 And bolSaveOK = True Then
               txtCode(0) = modBase(2)
               strAuto1 = modCP(9)
               strAuto2 = modBase(2)
           End If
           If bolSaveOK = True Then
              '外專信件沖銷: 收完文
              If InStr("," & mRetVal, "m_bolRecvOK = True") > 0 Then
                 m_bolRecvOK = True
              End If
              '多案收文的總收文號要傳入第一筆總收文號
              If InStr("," & mRetVal, "MCR11:") > 0 Then
                  m_strMCR11 = Mid(mRetVal, InStr(mRetVal, "MCR11:") + 6, 9)
              End If
              'Added by Lydia 2024/12/13 FG案輸入追蹤流水號：搬移命名追蹤檔案
              If frm010001.intModifyKind = 0 And Trim(txtTCN01.Text) <> "" And bolMoveCheck = True Then
                  Call PUB_UpdTCNfile(txtTCN01, txtSystem & txtCode(0) & IIf(txtCode(1) = "", "0", txtCode(1)) & IIf(txtCode(2) = "", "00", txtCode(2)), strAuto1, DBDATE(txtOther(0)), mSaveDir, bolMoveOK)
              End If
              'end 2024/12/13
           End If
      End If
'-----------------------------------------------
      If bolSaveOK = True Then
      'end 2022/09/14
         'Add By Sindy 2022/8/17 信件沖銷多案收文
         If m_strIR01 <> "" And m_bMRecvBatch = True Then
            '多案收文未收文完畢,繼續...
            strSql = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                     " and mcr11 is null" & _
                     " order by decode(mcr02||mcr03||mcr04||mcr05,mcr07||mcr08||mcr09||mcr10,1,2) asc,mcr02,mcr03,mcr04,mcr05 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               txtSystem = RsTemp.Fields("mcr02")
               txtCode(0) = RsTemp.Fields("mcr03")
               txtCode(1) = RsTemp.Fields("mcr04")
               txtCode(2) = RsTemp.Fields("mcr05")
               txtOther(1) = RsTemp.Fields("mcr06")
               IsSaveData = False
               ReadOtherDatabaseR '重新查詢
               DoEvents
               cmdOK(0).Value = True
               Exit Sub
            End If
         End If
         '2022/8/17 END
         
         PUB_SendMailCache 'Add by Sindy 2022/9/29
         
         frm010001.ClearForm strAuto1, strAuto2
         bolLeave = True
         intLeaveKind = 1
         If frm010001.intModifyKind = 0 Then LastDate = txtOther(0).Text
         
         'Modify By Sindy 2022/8/17 信件內部收文執行完畢後,關閉視窗
         If m_strIR01 <> "" Then
            If Pub_StrUserSt03 = "F23" And m_bolRecvOK = True Then
               '多案收文的總收文號要傳入第一筆總收文號
               'Modify By Sindy 2025/8/19 + ,iif(m_LOS02 <> "" And m_LOS15 <> "",m_LOS15,"")
               Call PUB_RecvOutLookF23(m_strIR01, m_strIR02, m_strIR03, m_strIR04, IIf(m_strMCR11 <> "", "2", "1"), IIf(m_strMCR11 <> "", m_strMCR11, strAuto1), IIf(m_LOS02 <> "" And m_LOS15 <> "", m_LOS15, ""))
            End If
            If Not m_PrevForm Is Nothing Then
               Call m_PrevForm.GoNext
            End If
            Unload Me
            Unload frm010001
         Else
         '2022/8/17 END
            Unload Me
         End If
      End If
   End If
   Screen.MousePointer = varSaveCursor
Else
   If Index = 2 Then
      intLeaveKind = 0
   Else
      intLeaveKind = 1
   End If
   Unload Me
End If

Exit Sub
   
'Add By Sindy 2022/8/17
ErrorHandler:
   If PUB_CheckFormExist("frmpic002") = True Then Unload frmpic002
   
   Screen.MousePointer = varSaveCursor 'Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Function SaveDatabase(ByRef strRecieveAuto As String, ByRef strCaseAuto As String) As Boolean
Dim adoquery As New ADODB.Recordset
Dim strSP78 As String, strContact As String
'Dim tST15 As String 'Remove by Lydia 2019/02/14
   
   'Modified by Lydia 2019/02/14
   'tST15 = Trim(PUB_GetStaffST15(txtOther(10).Text, "1"))
   'Modified by Lydia 2019/09/16
   'm_SalesST15 = GetST15(txtOther(10).Text)
   m_SalesST15 = GetST15(txtOther(10).Text, , , m_SalesST06)
   
   'Add by Morgan 2008/8/5
   If cboContact.Locked = False Then
      If cboContact.ListIndex >= 0 Then
         'Modify by Amy 2021/12/21 改成Form 2.0
         'If Val(cboContact.ItemData(cboContact.ListIndex)) > 0 Then
         '   strSP78 = Format(cboContact.ItemData(cboContact.ListIndex), "00")
         strSP78 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
         If Val(strSP78) > 0 Then
            'Add by Morgan 2008/8/7 若個案接洽人與客戶檔的預設接洽人相同時不必設定
            PUB_GetContact strAppNo1, strContact, True
            If strSP78 = strContact Then
               strSP78 = ""
            End If
         'Added by Lydia 2022/09/16 排除空白=00
         ElseIf strSP78 = "00" And Trim(cboContact.Text) = "" Then
             strSP78 = ""
         'end 2022/09/16
         End If
      End If
   Else
      strSP78 = "SP78"
   End If

   If frm010001.intModifyKind = 0 Then
      If strSP78 = "SP78" Then strSP78 = "" 'Add by Morgan 2008/8/7
      SaveDatabase = InsertOtherDatabase(frm010001.intSaveMode, frm010001.intCaseKind, txtSystem, txtCode(0), _
              IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtOther(3), IIf(Me.txtSystem.Text = "S" Or Me.txtSystem.Text = "TS" Or Me.txtSystem.Text = "T", Me.txtOther(24).Text, txtOther(4)), txtOther(5), txtOther(6), txtOther(7), _
              txtOther(18), txtOther(19), txtOther(8), txtOther(21), txtOther(0), txtOther(9), txtOther(14), txtOther(1), txtOther(2), txtOther(10), txtOther(12), _
              txtOther(16), txtOther(13), txtOther(17), txtOther(15), txtOther(11), txtOther(20), strRecieveAuto, strCaseAuto, douStPrice, douLowPrice, txtOther(31), txtOther(25), txtOther(26), txtOther(27), strSP78, txtOther(29), txtOther(30))
   Else
      SaveDatabase = UpdateOtherDatabase(frm010001.intCaseKind, txtSystem, txtCode(0), _
              IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtOther(3), IIf(Me.txtSystem.Text = "S" Or Me.txtSystem.Text = "TS" Or Me.txtSystem.Text = "T", Me.txtOther(24).Text, txtOther(4)), txtOther(5), txtOther(6), txtOther(7), _
              txtOther(18), txtOther(19), txtOther(8), Me.txtOther(21), txtRecieveCode, txtOther(0), txtOther(9), txtOther(14), txtOther(1), txtOther(2), txtOther(10), txtOther(12), _
              txtOther(16), txtOther(13), txtOther(17), txtOther(15), txtOther(11), txtOther(20), douStPrice, douLowPrice, txtOther(31), txtOther(25), txtOther(26), txtOther(27), strSP78, txtOther(29), txtOther(30))
   End If
   'add by nickc 2007/11/09 測試解決mail 發不到的時候會存兩筆的錯誤
   On Error GoTo 0    '歸零
   'add by nickc 2005/09/05
   If frm010001.intModifyKind = 0 Then
      'add by nick 2004/10/15  當收文業務區與客戶檔業務區不同時發 mail  及提示
      Dim oStrCuSales1 As String
      Dim oStrCuSales2 As String
      Dim oStrCuSales3 As String
      Dim oStrCuSales4 As String
      Dim oStrCuSales5 As String
      Dim oContext As String
      Dim oMailCount As String
      'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
      Dim IsMail As Boolean
      IsMail = True
      'add by nickc 2007/05/16 加入若是本所期限小於等於當天，要發mail  通知
      Dim oContext2 As String
      oContext2 = ""
      
      oStrCuSales1 = ""
      oStrCuSales2 = ""
      oStrCuSales3 = ""
      oStrCuSales4 = ""
      oStrCuSales5 = ""
      '2009/6/29 modify by sonia
      'oContext = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtOther(24) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtOther(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
      If Me.txtSystem.Text = "S" Or Me.txtSystem.Text = "TS" Or Me.txtSystem.Text = "T" Then
         oContext = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtOther(24) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtOther(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
      Else
         oContext = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtOther(4) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtOther(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
      End If
      ''2009/6/29 end
      'add by nickc 2007/05/16 加入若是本所期限小於等於當天，要發mail  通知
      'edit by nickc 2008/04/23 加入國家
      'oContext2 = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtOther(24) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtOther(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
      '2009/6/29 modify by sonia
      'oContext2 = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtOther(24) + vbCrLf + "申請國家：" + txtOther(3) + lblNation.Caption + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtOther(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
      If Me.txtSystem.Text = "S" Or Me.txtSystem.Text = "TS" Or Me.txtSystem.Text = "T" Then
         oContext2 = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtOther(24) + vbCrLf + "申請國家：" + txtOther(3) + lblNation.Caption + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtOther(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
      Else
         oContext2 = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtOther(4) + vbCrLf + "申請國家：" + txtOther(3) + lblNation.Caption + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtOther(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
      End If
      '2009/6/29 end
      
      'Added by Lydia 2020/10/05 (9/30) 若該收文號點數>0但無案源(自行收文者)時，若案件的客戶為非法律所的客戶時不論新舊案，系統自動新增TT-999999案進度(B類收文)及法律所案源資料。若為新案業務區不同的Email照舊通知。
      If m_Los05_N <> "" Then  '因為櫃台無法處理,所以只發email
          m_LOS05 = m_Los05_N
          m_LOS04_1 = m_Los04_N1
          m_LOS04_1st15 = GetST15(m_LOS04_1)
      End If
      'end 2020/10/05
            
      oMailCount = ""
      'Added by Lydia 2020/04/08 智慧所更名日起取消智權人員與客戶檔智權人員的控制
      'Modified by Lydia 2020/06/05 +著作權TC
      'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
      'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") Then
      If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 Then
            'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
            If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(7)) <> "" Then
                If ChkSameCuArea(Trim(txtOther(7)), m_LOS04_1) = False Then
                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(7).Text), oStrCuSales1)), 1) = "F" Then
                        '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
                    Else
                        oMailCount = oMailCount & oStrCuSales1 & ";"
                        oContext = oContext & vbCrLf + "申請人／當事人1： " + GetCustomerName(ChangeCustomerL(txtOther(7).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales1)
                    End If
                Else
                       IsMail = False
                End If
            ElseIf Trim(txtOther(7)) <> "" Then
            'end 2020/05/20
                GoTo JumpToChk01 'Added by Lydia 2022/10/18 無案源介紹人以畫面輸入判斷; ex.L-006576(桂英的客戶，法律所自行收文),因為之前無案源所以也不會補上案源資料
                IsMail = False
            End If 'Added by Lydia 2020/05/20
      Else
      'end 2020/04/08
                'PUB_SendMail strUserNum, Trim(txtother(12).Text) & ";" & GetCuSales(ChangeCustomerL(txtother(9).Text)), "", "案件收文通知--此案收文非原智權人員(區)！", vbCrLf + "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtother(6) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtother(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf + "申請人： " + GetCustomerName(ChangeCustomerL(txtother(9).Text)) + "原智權人員： " + oStrCuSales + vbCrLf + "收文智權人員： " + lblSales.Caption + vbCrLf + vbCrLf + "智權人員(區)不同！", ""
            'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
            'If tST15 <> GetCuSales(ChangeCustomerL(txtOther(7).Text), oStrCuSales1) And Trim(txtOther(10).Text) <> "" And Trim(txtOther(7).Text) <> "" Then
            'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
JumpToChk01: 'Added by Lydia 2022/10/18
            If ChkSameCuArea(Trim(txtOther(7)), Trim(txtOther(10)), , , , , Trim(txtOther(8))) = False And Trim(txtOther(10).Text) <> "" And Trim(txtOther(7).Text) <> "" Then
               'Add By Sindy 2009/10/19
               'Modified by Lydia 2019/02/14 tST15=>m_SalesST15
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(7).Text), oStrCuSales1)), 1) = "F" Then
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales1 & ";"
                  'edit by nickc 2005/08/16
                  'oContext = oContext & vbCrLf + "申請人／當事人1： " + GetCustomerName(ChangeCustomerL(txtOther(7).Text)) + "原智權人員： " + oStrCuSales1
                  oContext = oContext & vbCrLf + "申請人／當事人1： " + GetCustomerName(ChangeCustomerL(txtOther(7).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales1)
               End If
             'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
            Else
                   If Trim(txtOther(10).Text) <> "" And Trim(txtOther(7).Text) <> "" Then
                       IsMail = False
                   End If
            End If
      End If 'Added by Lydia 2020/04/08
      'Added by Lydia 2020/05/20 法律所案源收文：檢查是否為待活化客戶,並且更新DB
      If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(7)) <> "" Then
            If PUB_ChkOldCustomer(True, txtOther(7), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                IsMail = False
            End If
      Else
      'end 2020/05/20
            'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtOther(7)) <> "" And Trim(txtOther(10)) <> "" Then
                If PUB_ChkOldCustomer(True, txtOther(7), Trim(txtOther(10)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
                End If
            End If
      End If 'Added by Lydia 2020/05/20
      
      'Added by Lydia 2020/04/08 智慧所更名日起取消智權人員與客戶檔智權人員的控制
      'Modified by Lydia 2020/06/05 +著作權TC
      'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
      'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") Then
      If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 Then
            'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
            If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(18)) <> "" Then
                If ChkSameCuArea(Trim(txtOther(18)), m_LOS04_1) = False Then
                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(18).Text), oStrCuSales2)), 1) = "F" Then
                        '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
                    Else
                        oMailCount = oMailCount & oStrCuSales2 & ";"
                        oContext = oContext & vbCrLf + "申請人／當事人2： " + GetCustomerName(ChangeCustomerL(txtOther(18).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales2)
                    End If
                Else
                       IsMail = False
                End If
            ElseIf Trim(txtOther(18)) <> "" Then
            'end 2020/05/20
                GoTo JumpToChk02 'Added by Lydia 2022/10/18 無案源介紹人以畫面輸入判斷
                IsMail = False
            End If 'Added by Lydia 2020/05/20
      Else
      'end 2020/04/08
            'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
            'If tST15 <> GetCuSales(ChangeCustomerL(txtOther(18).Text), oStrCuSales2) And Trim(txtOther(10).Text) <> "" And Trim(txtOther(18).Text) <> "" Then
            'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
JumpToChk02: 'Added by Lydia 2022/10/18
            If ChkSameCuArea(Trim(txtOther(18)), Trim(txtOther(10)), , , , , Trim(txtOther(8))) = False And Trim(txtOther(10).Text) <> "" And Trim(txtOther(18).Text) <> "" Then
               'Add By Sindy 2009/10/19
                'Modified by Lydia 2019/02/14 tST15=>m_SalesST15
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(18).Text), oStrCuSales2)), 1) = "F" Then
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales2 & ";"
                  'edit by nickc 2005/08/16
                  'oContext = oContext & vbCrLf + "申請人／當事人2： " + GetCustomerName(ChangeCustomerL(txtOther(18).Text)) + "原智權人員： " + oStrCuSales2
                  oContext = oContext & vbCrLf + "申請人／當事人2： " + GetCustomerName(ChangeCustomerL(txtOther(18).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales2)
               End If
             'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
             Else
                   If Trim(txtOther(10).Text) <> "" And Trim(txtOther(18).Text) <> "" Then
                       IsMail = False
                   End If
            End If
      End If 'Added by Lydia 2020/04/08
      'Added by Lydia 2020/05/20 法律所案源收文：檢查是否為待活化客戶,並且更新DB
      If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(18)) <> "" Then
            If PUB_ChkOldCustomer(True, txtOther(18), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                IsMail = False
            End If
      Else
      'end 2020/05/20
            'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtOther(18)) <> "" And Trim(txtOther(10)) <> "" Then
                If PUB_ChkOldCustomer(True, txtOther(18), Trim(txtOther(10)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
                End If
            End If
      End If 'Added by Lydia 2020/05/20
      
      'Added by Lydia 2020/04/08 智慧所更名日起取消智權人員與客戶檔智權人員的控制
      'Modified by Lydia 2020/06/05 +著作權TC
      'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
      'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") Then
      If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 Then
            'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
            If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(19)) <> "" Then
                If ChkSameCuArea(Trim(txtOther(19)), m_LOS04_1) = False Then
                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(19).Text), oStrCuSales3)), 1) = "F" Then
                        '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
                    Else
                        oMailCount = oMailCount & oStrCuSales3 & ";"
                        oContext = oContext & vbCrLf + "申請人／當事人3： " + GetCustomerName(ChangeCustomerL(txtOther(19).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales3)
                    End If
                Else
                       IsMail = False
                End If
            ElseIf Trim(txtOther(19)) <> "" Then
            'end 2020/05/20
                GoTo JumpToChk03 'Added by Lydia 2022/10/18 無案源介紹人以畫面輸入判斷
                IsMail = False
            End If 'Added by Lydia 2020/05/20
      Else
      'end 2020/04/08
            'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
            'If tST15 <> GetCuSales(ChangeCustomerL(txtOther(19).Text), oStrCuSales3) And Trim(txtOther(10).Text) <> "" And Trim(txtOther(19).Text) <> "" Then
            'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
JumpToChk03: 'Added by Lydia 2022/10/18
            If ChkSameCuArea(Trim(txtOther(19)), Trim(txtOther(10)), , , , , Trim(txtOther(8))) = False And Trim(txtOther(10).Text) <> "" And Trim(txtOther(19).Text) <> "" Then
               'Add By Sindy 2009/10/19
               'Modified by Lydia 2019/02/14 tST15=>m_SalesST15
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(19).Text), oStrCuSales3)), 1) = "F" Then
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales3 & ";"
                  'edit by nickc 2005/08/16
                  'oContext = oContext & vbCrLf + "申請人／當事人3： " + GetCustomerName(ChangeCustomerL(txtOther(19).Text)) + "原智權人員： " + oStrCuSales3
                  oContext = oContext & vbCrLf + "申請人／當事人3： " + GetCustomerName(ChangeCustomerL(txtOther(19).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales3)
               End If
             'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
             Else
                   If Trim(txtOther(10).Text) <> "" And Trim(txtOther(19).Text) <> "" Then
                       IsMail = False
                   End If
            End If
      End If 'Added by Lydia 2020/04/08
      'Added by Lydia 2020/05/20 法律所案源收文：檢查是否為待活化客戶,並且更新DB
      If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(19)) <> "" Then
            If PUB_ChkOldCustomer(True, txtOther(19), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                IsMail = False
            End If
      Else
      'end 2020/05/20
            'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtOther(19)) <> "" And Trim(txtOther(10)) <> "" Then
                If PUB_ChkOldCustomer(True, txtOther(19), Trim(txtOther(10)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
                End If
            End If
      End If 'Added by Lydia 2020/05/20
      
      'Added by Lydia 2020/04/08 智慧所更名日起取消智權人員與客戶檔智權人員的控制
      'Modified by Lydia 2020/06/05 +著作權TC
      'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
      'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") Then
      If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 Then
            'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
            If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(29)) <> "" Then
                If ChkSameCuArea(Trim(txtOther(29)), m_LOS04_1) = False Then
                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(29).Text), oStrCuSales4)), 1) = "F" Then
                        '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
                    Else
                        oMailCount = oMailCount & oStrCuSales4 & ";"
                        oContext = oContext & vbCrLf + "申請人／當事人4： " + GetCustomerName(ChangeCustomerL(txtOther(29).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales4)
                    End If
                Else
                       IsMail = False
                End If
            ElseIf Trim(txtOther(29)) <> "" Then
            'end 2020/05/20
                GoTo JumpToChk04 'Added by Lydia 2022/10/18 無案源介紹人以畫面輸入判斷
                IsMail = False
            End If 'Added by Lydia 2020/05/20
      Else
      'end 2020/04/08
            'Add By Sindy 2011/1/18
            'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
            'If tST15 <> GetCuSales(ChangeCustomerL(txtOther(29).Text), oStrCuSales4) And Trim(txtOther(10).Text) <> "" And Trim(txtOther(29).Text) <> "" Then
            'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
JumpToChk04: 'Added by Lydia 2022/10/18
            If ChkSameCuArea(Trim(txtOther(29)), Trim(txtOther(10)), , , , , Trim(txtOther(8))) = False And Trim(txtOther(10).Text) <> "" And Trim(txtOther(29).Text) <> "" Then
               'Add By Sindy 2009/10/19
               'Modified by Lydia 2019/02/14 tST15=>m_SalesST15
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(29).Text), oStrCuSales4)), 1) = "F" Then
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales4 & ";"
                  oContext = oContext & vbCrLf + "申請人／當事人4： " + GetCustomerName(ChangeCustomerL(txtOther(29).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales4)
               End If
             Else
                   If Trim(txtOther(10).Text) <> "" And Trim(txtOther(29).Text) <> "" Then
                       IsMail = False
                   End If
            End If
      End If 'Added by Lydia 2020/04/08
      'Added by Lydia 2020/05/20 法律所案源收文：檢查是否為待活化客戶,並且更新DB
      If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(29)) <> "" Then
            If PUB_ChkOldCustomer(True, txtOther(29), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                IsMail = False
            End If
      Else
      'end 2020/05/20
            'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtOther(29)) <> "" And Trim(txtOther(10)) <> "" Then
                If PUB_ChkOldCustomer(True, txtOther(29), Trim(txtOther(10)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
               End If
            End If
      End If 'Added by Lydia 2020/05/20
      
      'Added by Lydia 2020/04/08 智慧所更名日起取消智權人員與客戶檔智權人員的控制
      'Modified by Lydia 2020/06/05 +著作權TC
      'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
      'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") Then
      If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 Then
            'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
            If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(30)) <> "" Then
                If ChkSameCuArea(Trim(txtOther(30)), m_LOS04_1) = False Then
                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(30).Text), oStrCuSales5)), 1) = "F" Then
                        '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
                    Else
                        oMailCount = oMailCount & oStrCuSales5 & ";"
                        oContext = oContext & vbCrLf + "申請人／當事人5： " + GetCustomerName(ChangeCustomerL(txtOther(30).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales5)
                    End If
                Else
                       IsMail = False
                End If
            ElseIf Trim(txtOther(30)) <> "" Then
            'end 2020/05/20
                GoTo JumpToChk05 'Added by Lydia 2022/10/18 無案源介紹人以畫面輸入判斷
                IsMail = False
            End If 'Added by Lydia 2020/05/20
      Else
      'end 2020/04/08
            'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
            'If tST15 <> GetCuSales(ChangeCustomerL(txtOther(30).Text), oStrCuSales5) And Trim(txtOther(10).Text) <> "" And Trim(txtOther(30).Text) <> "" Then
            'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
JumpToChk05: 'Added by Lydia 2022/10/18
            If ChkSameCuArea(Trim(txtOther(30)), Trim(txtOther(10)), , , , , Trim(txtOther(8))) = False And Trim(txtOther(10).Text) <> "" And Trim(txtOther(30).Text) <> "" Then
               'Add By Sindy 2009/10/19
               'Modified by Lydia 2019/02/14 tST15=>m_SalesST15
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(30).Text), oStrCuSales5)), 1) = "F" Then
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales5 & ";"
                  oContext = oContext & vbCrLf + "申請人／當事人5： " + GetCustomerName(ChangeCustomerL(txtOther(30).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales5)
               End If
             Else
                   If Trim(txtOther(10).Text) <> "" And Trim(txtOther(30).Text) <> "" Then
                       IsMail = False
                   End If
            End If
            '2011/1/18 End
      End If 'Added by Lydia 2020/04/08
      'Added by Lydia 2020/05/20 法律所案源收文：檢查是否為待活化客戶,並且更新DB
      If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(30)) <> "" Then
            If PUB_ChkOldCustomer(True, txtOther(30), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                IsMail = False
            End If
      Else
      'end 2020/05/20
            'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtOther(30)) <> "" And Trim(txtOther(10)) <> "" Then
                If PUB_ChkOldCustomer(True, txtOther(30), Trim(txtOther(10)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
                End If
            End If
      End If 'Added by Lydia 2020/05/20
      
'Remove by Morgan 2009/8/20
'      '2008/12/3 ADD BY SONIA 客戶檔之智權人員為國外部者不發mail
'      If IsMail = True Then
'         IsMail = PUB_CHKcusales(Trim(ChangeCustomerL(txtOther(7))) & "," & Trim(ChangeCustomerL(txtOther(18))) & "," & Trim(ChangeCustomerL(txtOther(19))))
'      End If
'      '2008/12/3 END
      
      'edit by nickc 2007/08/21 若申請人全空白，不發
      'If IsMail = False  Then
      'Modify By Sindy 2011/1/18
      If IsMail = False Or (Trim(txtOther(7)) = "" And Trim(txtOther(18)) = "" And Trim(txtOther(19)) = "" And Trim(txtOther(29)) = "" And Trim(txtOther(30)) = "") Then
           oMailCount = ""
      End If
      
      '2006/8/2 MODIFY BY SONIA TXTSYSTEM只判斷1碼,因為FG
      'If UCase(Mid(txtSystem, 1, 2)) <> "FC" And oMailCount <> "" Then
      'Modified by Lydia 2020/05/20 法律所案源收文：加上FCL
      'If UCase(Mid(txtSystem, 1, 1)) <> "F" And oMailCount <> "" Then
      If (UCase(Mid(txtSystem, 1, 1)) <> "F" Or (UCase(txtSystem) = "FCL" And m_LOS05 <> "")) And oMailCount <> "" Then
         'edit by nickc 2005/08/10
         'MsgBox "收文智權人員與客戶智權人員不同業務區，準備發 mail ，請定時刪除郵件備份！", , "注意！"
         'Modify By Sindy 2010/11/26 申請人為 X65299 或 X03072 的所有關係企業都不檢查業務區
         'Modify By Sindy 2011/1/18
         If Left(Trim(txtOther(7)), 6) <> "X65299" And Left(Trim(txtOther(7)), 6) <> "X03072" And _
            Left(Trim(txtOther(18)), 6) <> "X65299" And Left(Trim(txtOther(18)), 6) <> "X03072" And _
            Left(Trim(txtOther(19)), 6) <> "X65299" And Left(Trim(txtOther(19)), 6) <> "X03072" And _
            Left(Trim(txtOther(29)), 6) <> "X65299" And Left(Trim(txtOther(29)), 6) <> "X03072" And _
            Left(Trim(txtOther(30)), 6) <> "X65299" And Left(Trim(txtOther(30)), 6) <> "X03072" Then
            'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
            'Modified by Lydia 2020/06/05 +著作權TC
            'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
            'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") And m_LOS05 <> "" And m_LOS04_1 <> "" Then
            If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 And m_LOS05 <> "" And m_LOS04_1 <> "" Then
                  MsgBox "案源介紹人員與客戶智權人員不同業務區！", , "注意！"
            Else
            'end 2020/05/20
                  MsgBox "收文智權人員與客戶智權人員不同業務區，準備發 mail ！", , "注意！"
            End If 'Added by Lydia 2020/05/20
            'edit by nickc 2005/08/10 加發秀玲
            'oMailCount = oMailCount & Trim(txtOther(10).Text)
            'Added by Lydia 2022/07/15 通知法律所的智權人員沒有意義，應該要改為案源介紹人員. ex.L-006547
            'Modified by Lydia 2022/10/24 debug: 著作權TC用畫面的智權人員判斷
            'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") And m_LOS05 <> "" And m_LOS04_1 <> "" Then
            If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 And m_LOS05 <> "" And m_LOS04_1 <> "" Then
                 oMailCount = oMailCount & m_LOS04_1 & ";83002"
            Else
            'end 2022/07/15
                 oMailCount = oMailCount & Trim(txtOther(10).Text) & ";83002"
            End If 'Added by Lydia 2022/07/15
            
            'Added by Lydia 2022/11/03 法務案收文取消"是否林總同意本案由法律所自行收文？"的詢問，改為法律所自行收文智慧所人員客戶時，在判斷跨區收文發EMAIL給雙方智權人員時，加發林總。
            If strSrvDate(1) >= 法律所案源收文啟用日 And InStr(txtSystem, "L") > 0 And m_LOS02 = "" And m_LOS15 = "" Then
                oMailCount = oMailCount & PUB_ChkForLawMan(Trim(txtOther(7)), txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
            End If
            'end 2022/11/03
            
            'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
            'Modified by Lydia 2022/10/24 debug: 著作權TC用畫面的智權人員判斷
            'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") And m_LOS05 <> "" And m_LOS04_1 <> "" Then
            If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 And m_LOS05 <> "" And m_LOS04_1 <> "" Then
                oContext = oContext & vbCrLf + "案源介紹人員： " + GetStaffName(m_LOS04_1) + vbCrLf + vbCrLf + "智權人員(區)不同！"
            Else
            'end 2020/05/20
                oContext = oContext & vbCrLf + "收文智權人員： " + lblSales.Caption + vbCrLf + vbCrLf + "智權人員(區)不同！"
            End If 'Added by Lydia 2020/05/20
            PUB_SendMail strUserNum, oMailCount, "", "案件收文通知--此案收文非原智權人員(區)！", oContext
         End If
      End If
      
        'add by nickc 2007/05/16 加入若是本所期限小於等於當天，要發mail  通知
        oMailCount = ""
        If txtSystem = "P" Or txtSystem = "PS" Then
            oMailCount = Pub_GetSpecMan("A")
        ElseIf txtSystem = "CFP" Or txtSystem = "CPS" Then
             'edit by nickc 2007/10/16 修改到table
             'oMailCount = "PATENT"
             oMailCount = Pub_GetSpecMan("B")
        ElseIf txtSystem = "FCP" Or txtSystem = "FG" Then
             'edit by nickc 2007/10/16 修改到table
             'oMailCount = "73023;79012"
             oMailCount = Pub_GetSpecMan("C")
        'edit by nickc 2008/04/23 S 發 68005、72012
        'ElseIf txtSystem = "CFT" Or txtSystem = "S" Or txtSystem = "CFC" Then
        ElseIf txtSystem = "CFT" Or txtSystem = "CFC" Then
             'edit by nickc 2007/10/16 修改到table
             'oMailCount = "68005;72012"
             'edit by nickc 2008/04/23
             'oMailCount = Pub_GetSpecMan("D")
             'Modified by Lydia 2021/07/30 商標及商標服務業務收文-因外商陳經理退休而修改程式控制
             'oMailCount = Pub_GetSpecMan("L")
              oMailCount = GetCFTSt16Manager(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
        'edit by nickc 2008/04/23 S 發 68005、72012
        'ElseIf txtSystem = "FCT" Then
        ElseIf txtSystem = "FCT" Or txtSystem = "S" Then
            'edit by nickc 2007/10/16 修改到table
            'oMailCount = "68005;72012"
            'Added by Lydia 2021/07/30 商標及商標服務業務收文-因外商陳經理退休而修改程式控制
            If txtSystem = "S" Then
                 If txtOther(3) = "000" Then
                     'S台灣案:以本所案號呼叫PUB_GetFCTSalesNo抓出負責的人，再抓該員除ST55之外的最高主管NVL(NVL(ST54,ST53),ST52)
                     strExc(1) = PUB_GetFCTSalesNo(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
                     If strExc(1) = "" Then
                         oMailCount = Pub_GetSpecMan("D")
                     Else
                         oMailCount = PUB_GetSTManLimit(strExc(1), "4") '模組化
                        'Added by Lydia 2022/01/28 發給系統特殊設定「D」之人員和主管
                        strExc(2) = Pub_GetSpecMan("D")
                        oMailCount = oMailCount & ";" & strExc(2)
                        'end 2022/01/28
                     End If
                 Else
                     'S非台灣案:以本所案號呼叫GetCFTSt16Manager抓主管
                     oMailCount = GetCFTSt16Manager(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
                 End If
            Else
            'end 2021/07/30
                 oMailCount = Pub_GetSpecMan("D")
            End If 'Added by Lydia 2021/07/30
            'add by nickc 2007/06/23 加入FCT 爭議案通知內商商爭  84027;69008     案件性質 202 除外，還是送外商，阿蓮說自己判斷，若為內商案件，他會再轉過來
            If txtSystem = "FCT" Then    'add by nickc 2007/08/01 不判斷的話  CFT 也會進入
'edit by nickc 2007/08/10  內外商協議 202 都發
                If txtOther(1) = "202" Then
                    'edit by nickc 2007/10/16 修改到table
                    'oMailCount = "68005;72012;84027;69008"
                    oMailCount = Pub_GetSpecMan("F")
                End If
'                Dim tmp960623 As New ADODB.Recordset
'                Set tmp960623 = New ADODB.Recordset
'                If tmp960623.State = 1 Then tmp960623.Close
'                tmp960623.CursorLocation = adUseClient
'                tmp960623.Open "select * from staff_group where sg01='C1' and sg02='FCT' and sg03='" & txtOther(1) & "' and sg03<>'202'  ", cnnConnection, adOpenStatic, adLockReadOnly
'                If tmp960623.RecordCount <> 0 Then
'                     If Trim(txtOther(9)) <> "" And Trim(txtOther(14)) <> "" Then
'                         oMailCount = "84027;69008"
'                     End If
'                End If
'                tmp960623.Close
'                Set tmp960623 = Nothing
            End If
        ElseIf Mid(txtSystem, 1, 1) = "T" Then
             'edit by nickc 2007/10/16 修改到table
             'oMailCount = "84027;69008"
             oMailCount = Pub_GetSpecMan("E")
        End If
        If DBDATE(txtOther(9).Text) < strSrvDate(1) And Trim(txtOther(9).Text) <> "" And Trim(oMailCount) <> "" Then
           '2007/8/13 MODIFY BY SONIA 加智權人員
           'Modify By Sindy 2010/12/16 加業務區,費用
           PUB_SendMail strUserNum, oMailCount, "", "案件收文 緊急 通知--此案已逾本所期限，請儘速辦理！", oContext2 & vbCrLf & "本所期限：" & ChangeWStringToTDateString(DBDATE(txtOther(9).Text)) & vbCrLf & "法定期限：" & ChangeWStringToTDateString(DBDATE(txtOther(14).Text)) & vbCrLf & "智權人員　：" & lblSales & vbCrLf & "業務區　：" & lblDepartment & vbCrLf & "費用　　：" & Format(txtOther(12), "##,##0")
        End If
        If DBDATE(txtOther(9).Text) = strSrvDate(1) And Trim(txtOther(9).Text) <> "" And Trim(oMailCount) <> "" Then
           '2007/8/13 MODIFY BY SONIA 加智權人員
           'Modify By Sindy 2010/12/16 加業務區,費用
           PUB_SendMail strUserNum, oMailCount, "", "案件收文 緊急 通知--此案已屆本所期限，請儘速辦理！", oContext2 & vbCrLf & "本所期限：" & ChangeWStringToTDateString(DBDATE(txtOther(9).Text)) & vbCrLf & "法定期限：" & ChangeWStringToTDateString(DBDATE(txtOther(14).Text)) & vbCrLf & "智權人員　：" & lblSales & vbCrLf & "業務區　：" & lblDepartment & vbCrLf & "費用　　：" & Format(txtOther(12), "##,##0")
        End If
      
   End If

End Function

Private Sub ReadOtherDatabaseR()
Dim ot01 As String, ot02 As String, ot03 As String, ot04 As String, ot05 As String, _
              ot06 As String, ot07 As String, ot08 As String, ot09 As String, ot10 As String, _
              ot11 As String, ot12 As String, ot18 As String, _
              cp05 As String, cp06 As String, cp07 As String, CP10 As String, cp11 As String, _
              cp13 As String, cp14 As String, cp16 As String, cp17 As String, _
              cp18 As String, cp19 As String, cp32 As String, cu30 As String, _
              CP64 As String, rt As Boolean, ot13 As String, ot14 As String
'add by nickc 2006/11/30
Dim SP73 As String, SP74 As String
'add by nickc 2007/03/27
Dim SP27 As String
Dim CP150 As String 'Add By Sindy 2012/11/08

CP10 = txtOther(1)
'edit by nickc 2007/03/27 加入彼所案號
'rt = ReadOtherDatabase(frm010001.intModifyKind, frm010001.intCaseKind, txtSystem, txtCode(0), _
       IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), ot05, _
       ot06, ot07, ot08, ot09, ot10, ot11, ot12, ot18, txtRecieveCode, cp05, cp06, cp07, CP10, cp11, cp13, cp16, _
       cp17, cp18, cp19, cp32, cu30, cp14, CP64, SP73, SP74)
rt = ReadOtherDatabase(frm010001.intModifyKind, frm010001.intCaseKind, txtSystem, txtCode(0), _
       IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), ot05, _
       ot06, ot07, ot08, ot09, ot10, ot11, ot12, ot18, txtRecieveCode, cp05, cp06, cp07, CP10, cp11, cp13, cp16, _
       cp17, cp18, cp19, cp32, cu30, cp14, CP64, SP73, SP74, SP27, ot13, ot14, CP150)
'If rt Then
''NICK 900803 **********************
'txtCP64 = CP64
''**********************
   If frm010001.intModifyKind <> 0 Then
      txtOther(0) = cp05
      txtOther(1) = CP10
      txtOther(2) = cp11
      txtOther(10) = cp13
      txtOther(11) = cu30
      txtOther(12) = cp16
      txtOther(13) = cp18
      txtOther(15) = cp32
      txtOther(16) = cp17
      txtOther(17) = cp19
      txtOther(20) = cp14
      txtOther(31) = CP64 'Add By Sindy 2011/6/3 進度備註 : LA主題
      
      CheckKeyIn 1
      CheckKeyIn 2
      CheckKeyIn 10
      CheckKeyIn 20
      
      'Add By Sindy 2012/11/08
      If CP150 = "Y" Then
         Me.Check2.Value = 1
      End If
      '2012/11/08 End
   End If
   'add by nickc 2007/03/27
   txtOther(27) = SP27
   
   txtOther(9) = cp06
   txtOther(14) = cp07
   txtOther(3) = ot05
    'Modify By Cheng 2004/02/25
    Select Case Me.txtSystem.Text
    Case "TS", "S"
        txtOther(24) = ot06
    Case Else
        txtOther(4) = ot06
    End Select
    'End
   txtOther(5) = ot07
   txtOther(6) = ot08
   txtOther(7) = ot09 '當事人1
   txtOther(18) = ot10 '當事人2
   txtOther(19) = ot11 '當事人3
   txtOther(8) = ot12
   'Add By Sindy 2011/1/18
   txtOther(29) = ot13 '當事人4
   txtOther(30) = ot14 '當事人5
   '2011/1/18 End
    '案件備註
   txtOther(21) = ot18
   'Add By Cheng 2001/12/17
   '顯示智權人員代號
   'txtOther(10) = cp13  '2011/5/11 cancel by sonia 偶而改智權人員收文會忘記打所以不自動帶
   'Modify By Cheng 2002/01/03
   If Len("" & txtOther(7).Text) > 0 Then CheckKeyIn 7
   
   CheckKeyIn 8
   CheckKeyIn 18
   CheckKeyIn 19
   'Add By Sindy 2011/1/18
   CheckKeyIn 29
   CheckKeyIn 30
   '2011/1/18 End
   'Add By Cheng 2001/12/17
   If txtOther(10) <> "" Then CheckKeyIn 10
   
'Else
'   If frm010001.intModifyKind <> 0 Then
'      MsgBox "讀取資料時發生錯誤!!", vbCritical
'      bolLeave = True
'      Unload Me
'   Else
'      txtOther(9) = cp06
'      txtOther(14) = cp07
'      txtOther(3) = ot05
'      txtOther(4) = ot06
'      txtOther(5) = ot07
'      txtOther(6) = ot08
'      txtOther(7) = ot09
'      txtOther(18) = ot10
'      txtOther(19) = ot11
'      txtOther(8) = ot12
'      CheckKeyIn 7
'      CheckKeyIn 8
'      CheckKeyIn 18
'      CheckKeyIn 19
'   End If
'End If
'NICK 900803 **********************
If frm010001.intChoose = 1 Then
   txtOther(2) = "90"
   CheckKeyIn (2)
End If
' **********************
' 92.11.3 ADD BY SONIA
OnUpdateFee
'92.11.5 add by sonia
'Modify By Sindy 2009/07/24 增加LIN系統類別
'modify by sonia 2019/7/29 +ACS系統類別
Select Case txtSystem
   'Added by Lydia 2022/06/27 LawCase欄位名稱從40放到160
   Case "L", "CFL", "FCL", "LIN", "ACS"
      Label11.Caption = "案件名稱(中)（160）："
      txtOther(4).MaxLength = 40
      Label14.Caption = "案件名稱(日)（160）："
      txtOther(6).MaxLength = 40
   'end 2022/06/27
   'Modified by Lydia 2022/06/27 只保留HireCase=> LA
   'Case "L", "CFL", "FCL", "LA", "LIN", "ACS"
   Case "LA"
      Label11.Caption = "案件名稱(中)（40）："
      txtOther(4).MaxLength = 40
      Label14.Caption = "案件名稱(日)（40）："
      txtOther(6).MaxLength = 40
End Select
'92.11.5 end
'Add By Sindy 2011/6/3
'Modified by Lydia 2018/07/09 +L
'If txtSystem = "LA" Then
If txtSystem = "LA" Or txtSystem = "L" Then
   txtOther(31).Visible = True '主題
   Label36.Visible = True
Else
   txtOther(31).Visible = False
   Label36.Visible = False
End If
End Sub

'Add By Sindy 2010/3/8
Private Sub Command1_Click()
'開啟聯絡人視窗
frm010007_1.ReadServicePractice
frm010007_1.txt1(6) = strSP30s
frm010007_1.txt1(7) = strSP75s
frm010007_1.Show vbModal
bolCancel = frm010007_1.bolOK
If bolCancel = True Then
   strSP30s = frm010007_1.strSP30s
   strSP75s = frm010007_1.strSP75s
End If
Unload frm010007_1
Set frm010007_1 = Nothing
txtOther(8).SetFocus 'Add By Sindy 2010/3/19
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   If frm010001.intChoose = 1 Then
      fraPromoter.Visible = True
      txtOther(15) = "N"
   Else
      fraPromoter.Visible = False
   End If
   'add by nickc 2007/12/12
   IsSaveData = False
   'Add by Morgan 2008/8/5
   If frm010001.m_blnNewCase = True Then
      cboContact.Locked = False
   Else
      cboContact.Locked = True
   End If
   'end 2008/8/5
   
   'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
   Label32.Visible = False
   txtOther(28).Visible = False
   Check1.Visible = False
   
   'Add By Sindy 2022/8/17
   m_strIR01 = frm010001.m_strIR01
   m_strIR02 = frm010001.m_strIR02
   m_strIR03 = frm010001.m_strIR03
   m_strIR04 = frm010001.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2022/8/17 END
   
   'Added by Lydia 2024/12/13 FG案輸入追蹤流水號
   If frm010001.mRole = "" And frm010001.intModifyKind = 0 And frm010001.txtCode(0) = "" And frm010001.txtSystem = "FG" Then '排除外專後續案收文
      Call Pub_ChkExcelPath(App.path & "\" & strUserNum)  '先檢查個人資料夾
      mSaveDir = App.path & "\" & strUserNum & "\暫存區"
      If Dir(mSaveDir, vbDirectory) = "" Then
           MkDir mSaveDir
      End If
   End If
   'end 2024/12/13
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If frm010001.intModifyKind = 0 Or frm010001.intModifyKind = 1 Then
      If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Cancel = 1
      End If
   End If
End If
End Sub

Private Sub Form_Activate()
   'Add by Morgan 2004/4/15
   If bolActive Then
      Exit Sub
   Else
      bolActive = True
   End If
   
Dim strPKindName As String, strDate1 As String, StrDate2 As String, strCode(5) As String, i As Integer
Dim bolAdd As Boolean 'Added by Lydia 2016/04/27

Me.Refresh

'Add By Sindy 2012/11/12
If Left(Trim(txtSystem), 1) <> "T" Or Trim(txtSystem) = "TF" Then
   Check2.Visible = True
Else
   Check2.Visible = False
End If
'2012/11/12 End

'根據intModifyMode來調整fraWindow1 , fraWindow2
Select Case frm010001.intModifyKind
             Case 0
                        '新增：所有欄位皆可輸入
                        fraWindow1.Enabled = True
                        Select Case frm010001.intSaveMode
                                     Case 0
                                                fraWindow2.Enabled = False
                                     Case 1
                                                fraWindow2.Enabled = True
                                                Dim intWhere As Integer
                                                'edit by nickc 2007/02/02 不用 dll 了
                                                'If objPublicData.GetSystemKind(txtSystem.Text, , , intWhere) Then
                                                If ClsPDGetSystemKind(txtSystem.Text, , , intWhere) Then
                                                   If intWhere <> 國外_CF Then
                                                      txtOther(3) = 台灣國家代號
                                                      CheckKeyIn 3
                                                   End If
                                                End If
                        End Select
                        If LastDate = "" Then
                           txtOther(0).Text = GetTaiwanTodayDate
                        Else
                           txtOther(0).Text = LastDate
                        End If
                        txtOther_GotFocus 0
                        If txtSystem = 內商著作權 Or txtSystem = "L" Or txtSystem = "CFC" Then
                           txtOther(18).Enabled = True
                           txtOther(19).Enabled = True
                           'Modify By Sindy 2011/1/18
                           txtOther(29).Enabled = True
                           txtOther(30).Enabled = True
                           '2011/1/18 End
                        Else
                           txtOther(18).Enabled = False
                           txtOther(19).Enabled = False
                           'Modify By Sindy 2011/1/18
                           txtOther(29).Enabled = False
                           txtOther(30).Enabled = False
                           '2011/1/18 End
                        End If
                        
             Case 1
                        '修改：中間欄位不可輸入
                        fraWindow1.Enabled = True
                        Dim bolNew As Boolean
                        'edit by nickc 2007/02/06 不用 dll 了
                        'If obj001.IsNewCase(txtRecieveCode, bolNew) Then
                        If Cls001IsNewCase(txtRecieveCode, bolNew) Then
                           If bolNew Then
                              fraWindow2.Enabled = True
                              If txtSystem = 內商著作權 Or txtSystem = "L" Or txtSystem = "CFC" Then
                                 txtOther(18).Enabled = True
                                 txtOther(19).Enabled = True
                                 'Modify By Sindy 2011/1/18
                                 txtOther(29).Enabled = True
                                 txtOther(30).Enabled = True
                                 '2011/1/18 End
                              Else
                                 txtOther(18).Enabled = False
                                 txtOther(19).Enabled = False
                                 'Modify By Sindy 2011/1/18
                                 txtOther(29).Enabled = False
                                 txtOther(30).Enabled = False
                                 '2011/1/18 End
                              End If
                           Else
                              fraWindow2.Enabled = False
                           End If
                        Else
                           bolLeave = True
                           Unload Me
                           Exit Sub
                        End If
             Case 2
                        '刪除：所有欄位皆不可輸入
                        cmdOK(0).Visible = False
                        fraWindow1.Enabled = False
                        fraWindow2.Enabled = False
End Select

'Add By Sindy 2010/3/8 預設值
strSP30s = "": strSP75s = ""
bolCancel = False
'2010/3/8 End

If frm010001.intModifyKind <> 0 Or frm010001.intSaveMode <> 1 Then
   ReadOtherDatabaseR
End If

'Added by Lydia 2021/02/19 改欄寬設定
If txtSystem = "L" Or txtSystem = "CFL" Or txtSystem = "FCL" Or txtSystem = "LA" Or _
    txtSystem = "LIN" Or txtSystem = "ACS" Then
    'Table: LawCase, HireCase
    'Added by Lydia 2022/06/27 LawCase欄位名稱從40放到160
    If txtSystem <> "LA" Then   '排除HireCase
       Label11.Caption = "案件名稱(中)（160）："
       Label13.Caption = "案件名稱(英)（160）："
       Label14.Caption = "案件名稱(日)（160）："
       txtOther(4).MaxLength = 160
       txtOther(5).MaxLength = 160
       txtOther(6).MaxLength = 160
    Else
    'end 2022/06/27
       Label11.Caption = "案件名稱(中)（40）："
       Label13.Caption = "案件名稱(英)（60）："
       Label14.Caption = "案件名稱(日)（40）："
       txtOther(4).MaxLength = 40
       txtOther(5).MaxLength = 60
       txtOther(6).MaxLength = 40
    End If 'Added by Lydia 2022/06/27
ElseIf txtSystem <> "T" And txtSystem <> "TF" Then
    'Tabel: ServicePractice
    Label11.Caption = "案件名稱(中)（160）："
    Label13.Caption = "案件名稱(英)（180）："
    Label14.Caption = "案件名稱(日)（160）："
    txtOther(4).MaxLength = 160
    txtOther(5).MaxLength = 180
    txtOther(6).MaxLength = 160
Else
    'Table: TradeMark
    '商標案：只顯示Label29, txtOther(24)
End If
'end 2021/02/21

'Add By Cheng 2004/02/25
Select Case Me.txtSystem.Text
Case "TS", "S"
   Me.Label29.Visible = True
   Me.txtOther(24).Visible = True
   Me.txtOther(24).Enabled = True
   Me.Label11.Visible = False
   Me.txtOther(4).Visible = False
   Me.txtOther(4).Enabled = False
   Me.Label13.Visible = False
   Me.txtOther(5).Visible = False
   Me.txtOther(5).Enabled = False
   Me.Label14.Visible = False
   Me.txtOther(6).Visible = False
   Me.txtOther(6).Enabled = False
   'add by nickc 2006/11/30
   Me.txtOther(25).Visible = True
   Me.txtOther(25).Enabled = True
   Me.txtOther(26).Visible = True
   Me.txtOther(26).Enabled = True
    
   'Added by Lydia 2016/04/25 +查名單對應
   TMQList = ""
   bolOpen130 = False
   If strSrvDate(1) >= TMQ電子化啟用日 And TypeName(Tmpfrm090130) <> "" And Me.txtOther(1) = "001" And Me.txtSystem.Text = "TS" Then
        m_AttachPath = App.path & "\" & strUserNum
        If Dir(m_AttachPath, vbDirectory) = "" Then
           MkDir m_AttachPath
        End If
        'Modified by Lydia 2016/04/27 改成直接在畫面輸入查名代號
        'cmdTSMap.Visible = True
        lblTS.Visible = True: txtTS(0).Visible = True: txtTS(1).Visible = True
        txtTS(0) = Mid(strSrvDate(2), 1, 3)
        If frm010001.intModifyKind <> "0" Then
           strExc(1) = PUB_GetTMQCaseMapNo(txtRecieveCode, "", bolAdd)
           If bolAdd = False Then
              Call SetTxtTS(strExc(1))
              txtTS(0).Enabled = False: txtTS(1).Enabled = False
           ElseIf frm010001.intModifyKind = "1" Then
                txtTS(0).Enabled = True: txtTS(1).Enabled = True
           End If
        End If
   Else
        'Modified by Lydia 2016/04/27
        'cmdTSMap.Visible = False
        lblTS.Visible = False: txtTS(0).Visible = False: txtTS(1).Visible = False
        'Added by Lydia 2017/12/05 非查名，接洽人回到原位置
        Label37.Top = lblTS.Top: Label37.Left = lblTS.Left
        cboContact.Top = txtTS(0).Top: cboContact.Left = txtTS(0).Left
        'end 2017/12/05
   End If
   'end 2016/04/25
Case Else
   Me.Label29.Visible = False
   Me.txtOther(24).Visible = False
   Me.txtOther(24).Enabled = False
   Me.Label11.Visible = True
   Me.txtOther(4).Visible = True
   Me.txtOther(4).Enabled = True
   Me.Label13.Visible = True
   Me.txtOther(5).Visible = True
   Me.txtOther(5).Enabled = True
   Me.Label14.Visible = True
   Me.txtOther(6).Visible = True
   Me.txtOther(6).Enabled = True
   'add by nickc 2006/11/30
   Me.txtOther(25).Visible = False
   Me.txtOther(25).Enabled = False
   Me.txtOther(26).Visible = False
   Me.txtOther(26).Enabled = False
   'Add By Sindy 2011/6/3
   'Modified by Lydia 2018/07/09  +L
   'If Me.txtSystem.Text = "LA" Then
   If Me.txtSystem.Text = "LA" Or Me.txtSystem.Text = "L" Then
      txtOther(31).Visible = True '主題
      Label36.Visible = True
   Else
      txtOther(31).Visible = False
      Label36.Visible = False
   End If
   Call ReadLOS 'Added by Lydia 2020/05/20 法律所案源收文：讀取法務案源檔
End Select
'End
'Add by Amy 2021/12/21 改form2.0 TopIndex 有問題,因判斷複雜,故先寫死
If UCase(App.EXEName) = "TEWRITER" Or UCase(App.EXEName) = "WRITER" Then
    'Modify by Amy 2022/01/06 +txtOther(2).Enabled = True,避免txtOther(2).Enabled = false 會出現引號異常就跳離開
    If txtOther(0) <> MsgText(601) And txtOther(1) <> MsgText(601) And txtOther(2).Enabled = True Then
        txtOther(2).SetFocus
    End If
End If

'Added by Lydia 2022/07/15 TC案之文件齊備日管控
Frame21.Left = 30
If txtSystem = "TC" Then
     Frame21.Visible = True
End If

   'Added by Lydia 2022/09/14
   If strSrvDate(1) >= 收文存檔模組化啟用日 Then
       Call SetDBArray(True, txtRecieveCode, txtSystem, txtCode(0), txtCode(1), txtCode(2))
   End If
   
   'Added by Lydia 2022/12/06 FCL案件收文，若有案源資料且LOS02>'B'時，進入frm010007時彈訊息"此類案源FCL案不用輸入費用資料！"，並鎖住畫面之費用、規費、點數不可輸入。
   If frm010001.intModifyKind < 2 And txtSystem = "FCL" And m_LOS02 > "B" Then
       MsgBox "此類案源FCL案不用輸入費用資料！", vbInformation
      txtOther(12).Locked = True
      txtOther(16).Locked = True
      txtOther(13).Locked = True
   End If
   
   'Added by Lydia 2024/12/13  FG案輸入追蹤流水號TrackingNo
   FraTCN.BackColor = &H8000000F
   FraTCN.Visible = False
   FraTCN.Left = 72
   'end 2024/12/13
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Where01ToGo intLeaveKind
   intLeaveKind = 0
   PUB_SendMailCache 'Added by Lydia 2020/05/20
   
   'Added by Lydia 2024/12/13 FG案輸入追蹤流水號TrackingNo
   If strSrvDate(1) >= XY特殊權限啟用日by檔案 And mSaveDir <> "" Then
       If bolMoveOK = True Then  'TrackingNO是否已搬檔完成(True無問題)，若有問題則TrackingNO和本機端的資料夾不刪除
           Call PUB_KillAnyFile(mSaveDir)
           RmDir mSaveDir  '移除資料夾
       End If
   End If
   'end 2024/12/13
   
   'Add By Cheng 2002/07/18
   'Modify by Amy 2021/12/20 改Form2.0後,存檔按Enter會當掉,改在呼叫時清除記憶體變數
   'Set frm010007 = Nothing
   'Added by Lydia 2016/04/25
   'Modify By Sindy 2022/8/17 + And TypeName(Tmpfrm090130) <> "Nothing"
   If TypeName(m_PrevForm) <> "Nothing" And TypeName(Tmpfrm090130) <> "Nothing" Then
      Set m_PrevForm.Tmpfrm090130 = Tmpfrm090130
   End If
   
   'Add By Sindy 2022/8/17
   If m_strIR01 <> "" Then
      If Not m_PrevForm Is Nothing Then
         Set m_PrevForm = Nothing
      End If
   End If
   If PUB_CheckFormExist("frmpic002") = True Then Unload frmpic002
   '2022/8/17 END
   
   stChkForm = Me.Name 'Add by Amy 2021/12/21
End Sub

Private Sub lblPetition_Change(Index As Integer)
'2005/11/17 MODIFY BY SONIA
'If Me.txtSystem.Text = "TB" Or Me.txtSystem.Text = "CFL" Or Me.txtSystem.Text = "FCL" Or Me.txtSystem.Text = "L" Then
If Me.txtSystem.Text = "TB" Then
'2005/11/17 END
   Select Case Index
   Case 0 '申請人名稱1
      If frm010001.intModifyKind = 0 Then '新增狀態
         Me.txtOther(4).Text = Me.lblPetition(Index).Caption
      End If
   Case Else
   '無
   End Select
End If
End Sub

Private Sub txtOther_Change(Index As Integer)
   Select Case Index
      Case 2
                 lblCaseSource.Caption = ""
      Case 3
                 lblNation.Caption = ""
      Case 7      '申請人/當事人1
                 lblPetition(0).Caption = ""
                 txtOther(11).Text = ""
      Case 18, 19 '申請人/當事人2,3
                 lblPetition(Index - 17).Caption = ""
      Case 8
                 lblAgent.Caption = ""
      Case 10
                 lblSales.Caption = ""
                 lblDepartment = ""
                 m_SalesST15 = "" 'Added by Lydia 2019/02/14
                 m_SalesST06 = "" 'Addded by Lydia 2019/09/16
      Case 20
                 lblPromoter = ""
      'Add By Sindy 2011/1/18
      Case 29, 30 '申請人/當事人4,5
                 lblPetition(Index - 26).Caption = ""
   End Select
End Sub

Private Sub txtOther_Validate(Index As Integer, Cancel As Boolean)
'add by nickc 2006/11/30
Dim ii As Integer
Dim arrTM09

Select Case Index
            Case 10
                    'add by nick 2005/01/04
                    If txtOther(Index).Text <> "" And txtOther(Index) < "63001" Then
                         MsgBox "智權人員不可小於 63001！", , "注意！"
                         Cancel = True
                         Exit Sub
                    End If
                'add by nick 2004/12/08 因為之前的 智權人員並沒有抓
                    Dim strTemp As String, strTemp1 As String
                    'edit by nickc 2007/02/02 不用 dll 了
                    'If Not objPublicData.GetStaff(txtOther(10).Text, strTemp, strTemp1) Then
                    If Not ClsPDGetStaff(txtOther(10).Text, strTemp, strTemp1) Then
                        Cancel = True
                        Exit Sub
                    End If
                    'add by nickc 2006/11/02
                    'Modified by Lydia 2019/02/14
                    'PUB_GetStaffST15 txtOther(10).Text, strTemp1
                    'Modified by Lydia 2019/09/16
                    'm_SalesST15 = GetST15(txtOther(10).Text, strTemp1)
                    m_SalesST15 = GetST15(txtOther(10).Text, strTemp1, , m_SalesST06)
                    
                    'Added by Lydia 2020/04/08 法務案(L、CFL)及顧問案LA之智權人員只能是法律所人員
                    If PUB_ChkSalesL(txtSystem, txtOther(10).Text) = False Then
                    End If
                    'end 2020/04/08
                    'Added by Lydia 2024/12/13 FG案輸入追蹤流水號TrackingNo
                    If txtSystem = "FG" And txtCode(0) = "" And Left(m_SalesST15, 1) = "F" Then
                       FraTCN.Visible = True
                    Else
                       FraTCN.Visible = False
                    End If
                    'end 2024/12/13
                    
                    lblSales.Caption = strTemp
                    lblDepartment = strTemp1
                    'Added by Lydia 2019/02/14 創新業務部人員收文控管
                    If PUB_ChkIsT10T20("2", txtOther(10).Text, m_Tuser, strTemp) = True Then
                        txtOther(10) = m_Tuser
                        lblSales.Caption = strTemp
                        txtOther(10).SetFocus
                        Call txtOther_GotFocus(10)
                        Cancel = True
                        Exit Sub
                    End If
                    'end 2019/02/14
                    
                'add by nick 2004/10/15  當收文業務區與客戶檔業務區不同時發 mail  及提示
                Dim oStrCuSales1 As String
                Dim oStrCuSales2 As String
                Dim oStrCuSales3 As String
                Dim oStrCuSales4 As String
                Dim oStrCuSales5 As String
                Dim oMailCount As String
                'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                Dim IsMail As Boolean
                IsMail = True
                oStrCuSales1 = ""
                oStrCuSales2 = ""
                oStrCuSales3 = ""
                oStrCuSales4 = ""
                oStrCuSales5 = ""
                oMailCount = ""
                'Remove by Lydia 2019/02/14
                'Dim tST15 As String
                'tST15 = Trim(PUB_GetStaffST15(txtOther(10).Text, "1"))
                'end 2019/02/14
   
                'Added by Lydia 2020/04/08 智慧所更名日起取消智權人員與客戶檔智權人員的控制
                'Modified by Lydia 2020/06/05 +著作權TC
                'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
                'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") Then
                If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 Then
                    'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
                    If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(7)) <> "" Then
                        If ChkSameCuArea(Trim(txtOther(7)), m_LOS04_1) = False Then
                        Else
                               IsMail = False
                        End If
                    ElseIf Trim(txtOther(7)) <> "" Then
                    'end 2020/05/20
                        IsMail = False
                    End If 'Added by Lydia 2020/05/20
                Else
                'end 2020/04/08
                        'PUB_SendMail strUserNum, Trim(txtother(12).Text) & ";" & GetCuSales(ChangeCustomerL(txtother(9).Text)), "", "案件收文通知--此案收文非原智權人員(區)！", vbCrLf + "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtother(6) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtother(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf + "申請人： " + GetCustomerName(ChangeCustomerL(txtother(9).Text)) + "原智權人員： " + oStrCuSales + vbCrLf + "收文智權人員： " + lblSales.Caption + vbCrLf + vbCrLf + "智權人員(區)不同！", ""
                    'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
                    'If tST15 <> GetCuSales(ChangeCustomerL(txtOther(7).Text), oStrCuSales1) And Trim(txtOther(10).Text) <> "" And Trim(txtOther(7).Text) <> "" Then
                    'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
                    If ChkSameCuArea(Trim(txtOther(7)), Trim(txtOther(10)), , , , , Trim(txtOther(8))) = False And Trim(txtOther(10).Text) <> "" And Trim(txtOther(7).Text) <> "" Then
                        'oMailCount = oMailCount & oStrCuSales1 & ";"
                    'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    Else
                          If Trim(txtOther(10).Text) <> "" And Trim(txtOther(7).Text) <> "" Then
                              IsMail = False
                          End If
                    End If
                End If 'Added by Lydia 2020/04/08
                'Added by Lydia 2020/05/20 法律所案源收文：檢查是否為待活化客戶
                If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(7)) <> "" Then
                       If PUB_ChkOldCustomer(False, txtOther(7), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                          IsMail = False
                       End If
                Else
                'end 2020/05/20
                    'Added by Lydia 2019/09/16 檢查是否為待活化客戶
                    If m_SalesST06 <> "" And Trim(txtOther(7)) <> "" And Trim(txtOther(10)) <> "" Then
                        If PUB_ChkOldCustomer(False, txtOther(7), Trim(txtOther(10)), m_SalesST15, m_SalesST06) = True Then
                           IsMail = False
                        End If
                    End If
                End If 'end 2020/05/20
                
                'Added by Lydia 2020/04/08 智慧所更名日起取消智權人員與客戶檔智權人員的控制
                'Modified by Lydia 2020/06/05 +著作權TC
                'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
                'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") Then
                If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 Then
                    'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
                    If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(18)) <> "" Then
                        If ChkSameCuArea(Trim(txtOther(18)), m_LOS04_1) = False Then
                        Else
                               IsMail = False
                        End If
                    ElseIf Trim(txtOther(18)) <> "" Then
                    'end 2020/05/20
                        IsMail = False
                    End If 'Added by Lydia 2020/05/20
                Else
                'end 2020/04/08
                    'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
                    'If tST15 <> GetCuSales(ChangeCustomerL(txtOther(18).Text), oStrCuSales2) And Trim(txtOther(10).Text) <> "" And Trim(txtOther(18).Text) <> "" Then
                    'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
                    If ChkSameCuArea(Trim(txtOther(18)), Trim(txtOther(10)), , , , , Trim(txtOther(8))) = False And Trim(txtOther(10).Text) <> "" And Trim(txtOther(18).Text) <> "" Then
                        'oMailCount = oMailCount & oStrCuSales2 & ";"
                    'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    Else
                          If Trim(txtOther(10).Text) <> "" And Trim(txtOther(18).Text) <> "" Then
                              IsMail = False
                          End If
                    End If
                End If 'Added by Lydia 2020/04/08
                'Added by Lydia 2020/05/20 法律所案源收文：檢查是否為待活化客戶
                If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(18)) <> "" Then
                       If PUB_ChkOldCustomer(False, txtOther(18), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                          IsMail = False
                       End If
                Else
                'end 2020/05/20
                        'Added by Lydia 2019/09/16 檢查是否為待活化客戶
                        If m_SalesST06 <> "" And Trim(txtOther(18)) <> "" And Trim(txtOther(10)) <> "" Then
                            If PUB_ChkOldCustomer(False, txtOther(18), Trim(txtOther(10)), m_SalesST15, m_SalesST06) = True Then
                               IsMail = False
                            End If
                        End If
                End If 'Added by Lydia 2020/05/20
                
                'Added by Lydia 2020/04/08 智慧所更名日起取消智權人員與客戶檔智權人員的控制
                'Modified by Lydia 2020/06/05 +著作權TC
                'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
                'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") Then
                If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 Then
                    'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
                    If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(19)) <> "" Then
                        If ChkSameCuArea(Trim(txtOther(19)), m_LOS04_1) = False Then
                        Else
                               IsMail = False
                        End If
                    ElseIf Trim(txtOther(19)) <> "" Then
                    'end 2020/05/20
                        IsMail = False
                    End If 'Added by Lydia 2020/05/20
                Else
                'end 2020/04/08
                    'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
                    'If tST15 <> GetCuSales(ChangeCustomerL(txtOther(19).Text), oStrCuSales3) And Trim(txtOther(10).Text) <> "" And Trim(txtOther(19).Text) <> "" Then
                    'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
                    If ChkSameCuArea(Trim(txtOther(19)), Trim(txtOther(10)), , , , , Trim(txtOther(8))) = False And Trim(txtOther(10).Text) <> "" And Trim(txtOther(19).Text) <> "" Then
                    'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    Else
                          If Trim(txtOther(10).Text) <> "" And Trim(txtOther(19).Text) <> "" Then
                              IsMail = False
                          End If
                    End If
                End If 'Added by Lydia 2020/04/08
                'Added by Lydia 2020/05/20 法律所案源收文：檢查是否為待活化客戶
                If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(19)) <> "" Then
                       If PUB_ChkOldCustomer(False, txtOther(19), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                          IsMail = False
                       End If
                Else
                'end 2020/05/20
                        'Added by Lydia 2019/09/16 檢查是否為待活化客戶
                        If m_SalesST06 <> "" And Trim(txtOther(19)) <> "" And Trim(txtOther(10)) <> "" Then
                            If PUB_ChkOldCustomer(False, txtOther(19), Trim(txtOther(10)), m_SalesST15, m_SalesST06) = True Then
                               IsMail = False
                            End If
                        End If
                End If 'Added by Lydia 2020/05/20
                
                'Added by Lydia 2020/04/08 智慧所更名日起取消智權人員與客戶檔智權人員的控制
                'Modified by Lydia 2020/06/05 +著作權TC
                'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
                'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") Then
                If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 Then
                    'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
                    If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(29)) <> "" Then
                        If ChkSameCuArea(Trim(txtOther(29)), m_LOS04_1) = False Then
                        Else
                               IsMail = False
                        End If
                    ElseIf Trim(txtOther(29)) <> "" Then
                    'end 2020/05/20
                        IsMail = False
                    End If 'Added by Lydia 2020/05/20
                Else
                'end 2020/04/08
                    'Add By Sindy 2011/1/18
                    'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
                    'If tST15 <> GetCuSales(ChangeCustomerL(txtOther(29).Text), oStrCuSales4) And Trim(txtOther(10).Text) <> "" And Trim(txtOther(29).Text) <> "" Then
                    'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
                    If ChkSameCuArea(Trim(txtOther(29)), Trim(txtOther(10)), , , , , Trim(txtOther(8))) = False And Trim(txtOther(10).Text) <> "" And Trim(txtOther(29).Text) <> "" Then
                    'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    Else
                          If Trim(txtOther(10).Text) <> "" And Trim(txtOther(29).Text) <> "" Then
                              IsMail = False
                          End If
                    End If
                End If 'Added by Lydia 2020/04/08
                'Added by Lydia 2020/05/20 法律所案源收文：檢查是否為待活化客戶
                If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(29)) <> "" Then
                       If PUB_ChkOldCustomer(False, txtOther(29), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                          IsMail = False
                       End If
                Else
                'end 2020/05/20
                        'Added by Lydia 2019/09/16 檢查是否為待活化客戶
                        If m_SalesST06 <> "" And Trim(txtOther(29)) <> "" And Trim(txtOther(10)) <> "" Then
                            If PUB_ChkOldCustomer(False, txtOther(29), Trim(txtOther(10)), m_SalesST15, m_SalesST06) = True Then
                               IsMail = False
                            End If
                        End If
                End If 'Added by Lydia 2020/05/20
                
                'Added by Lydia 2020/04/08 智慧所更名日起取消智權人員與客戶檔智權人員的控制
                'Modified by Lydia 2020/06/05 +著作權TC
                'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") Then
                If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 Then
                    'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
                    If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(30)) <> "" Then
                        If ChkSameCuArea(Trim(txtOther(30)), m_LOS04_1) = False Then
                        Else
                               IsMail = False
                        End If
                    ElseIf Trim(txtOther(30)) <> "" Then
                    'end 2020/05/20
                        IsMail = False
                    End If 'Added by Lydia 2020/05/20
                Else
                'end 2020/04/08
                    'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
                    'If tST15 <> GetCuSales(ChangeCustomerL(txtOther(30).Text), oStrCuSales5) And Trim(txtOther(10).Text) <> "" And Trim(txtOther(30).Text) <> "" Then
                    'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
                    If ChkSameCuArea(Trim(txtOther(30)), Trim(txtOther(10)), , , , , Trim(txtOther(8))) = False And Trim(txtOther(10).Text) <> "" And Trim(txtOther(30).Text) <> "" Then
                    'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    Else
                          If Trim(txtOther(10).Text) <> "" And Trim(txtOther(30).Text) <> "" Then
                              IsMail = False
                          End If
                    End If
                    '2011/1/18 End
                End If 'Added by Lydia 2020/04/08
                'Added by Lydia 2020/05/20 法律所案源收文：檢查是否為待活化客戶
                If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(txtOther(30)) <> "" Then
                       If PUB_ChkOldCustomer(False, txtOther(30), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                          IsMail = False
                       End If
                Else
                'end 2020/05/20
                        'Added by Lydia 2019/09/16 檢查是否為待活化客戶
                        If m_SalesST06 <> "" And Trim(txtOther(30)) <> "" And Trim(txtOther(10)) <> "" Then
                            If PUB_ChkOldCustomer(False, txtOther(30), Trim(txtOther(10)), m_SalesST15, m_SalesST06) = True Then
                               IsMail = False
                            End If
                        End If
                End If 'Added by Lydia 2020/05/20
                
'Remove by Morgan 2009/8/20
'                  '2008/12/3 ADD BY SONIA 客戶檔之智權人員為國外部者不發mail
'                  If IsMail = True Then
'                     IsMail = PUB_CHKcusales(Trim(ChangeCustomerL(txtOther(7))) & "," & Trim(ChangeCustomerL(txtOther(18))) & "," & Trim(ChangeCustomerL(txtOther(19))))
'                  End If
'                  '2008/12/3 END
                  
                '2006/8/2 MODIFY BY SONIA TXTSYSTEM只判斷1碼,因為FG
                'If UCase(Mid(txtSystem, 1, 2)) <> "FC" And oMailCount <> "" Then
                'If UCase(Mid(txtSystem, 1, 1)) <> "F" And oMailCount <> "" Then
                'edit by nickc 2007/05/10 秀玲說 , 其中一個符合就不發了
                'If UCase(Mid(txtSystem, 1, 1)) <> "F" And oMailCount <> "" Then
                'edit by nickc 2008/03/26 若是申請人空白，則不管
                'If UCase(Mid(txtSystem, 1, 1)) <> "F" And IsMail = True Then
                'Modify By Sindy 2011/1/18
                'Modified by Lydia 2020/05/20 法律所案源收文：加上FCL
                'If UCase(Mid(txtSystem, 1, 1)) <> "F" And IsMail = True And (txtOther(7) <> "" Or txtOther(18) <> "" Or txtOther(19) <> "" Or txtOther(29) <> "" Or txtOther(30) <> "") Then
                If (UCase(Mid(txtSystem, 1, 1)) <> "F" Or (UCase(txtSystem) = "FCL" And m_LOS05 <> "")) And IsMail = True And (txtOther(7) <> "" Or txtOther(18) <> "" Or txtOther(19) <> "" Or txtOther(29) <> "" Or txtOther(30) <> "") Then
                     'Add By Sindy 2009/10/19
                     '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail，不顯示訊息
                     oMailCount = ""
                     'Added by Lydia 2020/04/08 智慧所更名日起取消智權人員與客戶檔智權人員的控制
                     'Modified by Lydia 2020/06/05 +著作權TC
                     'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
                     'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") Then
                     If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 Then
                            'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
                            If m_LOS05 <> "" And m_LOS04_1 <> "" Then
                                If txtOther(7) <> "" Then
                                   If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(7).Text), oStrCuSales1)), 1) = "F" Then
                                   Else
                                      oMailCount = "Y"
                                   End If
                                End If
                                If txtOther(18) <> "" Then
                                   If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(18).Text), oStrCuSales1)), 1) = "F" Then
                                   Else
                                      oMailCount = "Y"
                                   End If
                                End If
                                If txtOther(19) <> "" Then
                                   If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(19).Text), oStrCuSales1)), 1) = "F" Then
                                   Else
                                      oMailCount = "Y"
                                   End If
                                End If
                                If txtOther(29) <> "" Then
                                   If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(29).Text), oStrCuSales1)), 1) = "F" Then
                                   Else
                                      oMailCount = "Y"
                                   End If
                                End If
                                If txtOther(30) <> "" Then
                                   If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(30).Text), oStrCuSales1)), 1) = "F" Then
                                   Else
                                      oMailCount = "Y"
                                   End If
                                End If
                            End If
                            'end 2020/05/20
                     Else
                     'end 2020/04/08
                            If txtOther(7) <> "" Then
                               'Modified by Lydia 2019/02/14 tST15=>m_SalesST15
                               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(7).Text), oStrCuSales1)), 1) = "F" Then
                               Else
                                  oMailCount = "Y"
                               End If
                            End If
                            If txtOther(18) <> "" Then
                               'Modified by Lydia 2019/02/14 tST15=>m_SalesST15
                               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(18).Text), oStrCuSales1)), 1) = "F" Then
                               Else
                                  oMailCount = "Y"
                               End If
                            End If
                            If txtOther(19) <> "" Then
                               'Modified by Lydia 2019/02/14 tST15=>m_SalesST15
                               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(19).Text), oStrCuSales1)), 1) = "F" Then
                               Else
                                  oMailCount = "Y"
                               End If
                            End If
                            'Modify By Sindy 2011/1/18
                            If txtOther(29) <> "" Then
                               'Modified by Lydia 2019/02/14 tST15=>m_SalesST15
                               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(29).Text), oStrCuSales1)), 1) = "F" Then
                               Else
                                  oMailCount = "Y"
                               End If
                            End If
                            If txtOther(30) <> "" Then
                               'Modified by Lydia 2019/02/14 tST15=>m_SalesST15
                               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtOther(30).Text), oStrCuSales1)), 1) = "F" Then
                               Else
                                  oMailCount = "Y"
                               End If
                            End If
                            '2011/1/18 End
                     End If 'Added by Lydia 2020/04/08
                     If Trim(oMailCount) <> "" Then
                     '2009/10/19 End
                        'Modify By Sindy 2010/11/26 申請人為 X65299 或 X03072 的所有關係企業都不檢查業務區
                        'Modify By Sindy 2011/1/18
                        If Left(Trim(txtOther(7)), 6) <> "X65299" And Left(Trim(txtOther(7)), 6) <> "X03072" And _
                           Left(Trim(txtOther(18)), 6) <> "X65299" And Left(Trim(txtOther(18)), 6) <> "X03072" And _
                           Left(Trim(txtOther(19)), 6) <> "X65299" And Left(Trim(txtOther(19)), 6) <> "X03072" And _
                           Left(Trim(txtOther(29)), 6) <> "X65299" And Left(Trim(txtOther(29)), 6) <> "X03072" And _
                           Left(Trim(txtOther(30)), 6) <> "X65299" And Left(Trim(txtOther(30)), 6) <> "X03072" Then
                           'Added by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
                           'Modified by Lydia 2020/06/05 +著作權TC
                           'Modified by Lydia 2022/10/18 debug: 著作權TC用畫面的智權人員判斷
                           'If strSrvDate(1) >= 智慧所更名日 And (InStr(txtSystem, "L") > 0 Or txtSystem = "TC") And m_LOS05 <> "" And m_LOS04_1 <> "" Then
                           If strSrvDate(1) >= 智慧所更名日 And InStr(txtSystem, "L") > 0 And m_LOS05 <> "" And m_LOS04_1 <> "" Then
                                 MsgBox "案源介紹人員與客戶智權人員不同業務區！", , "注意！"
                           Else
                           'end 2020/05/20
                                 MsgBox "收文智權人員與客戶智權人員不同業務區！", , "注意！"
                           End If 'Added by Lydia 2020/05/20
                        End If
                     End If
                End If
                            
            Case 3
                  '2005/12/6 add by sonia
                  'Modify By Sindy 2009/07/24 增加LIN系統類別
                  'modify by sonia 2019/7/30 +ACS系統類別
                  If (txtSystem = "FCL" Or txtSystem = "LIN" Or txtSystem = "ACS") And txtOther(Index).Text <> 台灣國家代號 Then
                     ShowMsg MsgText(9219)
                     Cancel = True
                     Exit Sub
                  End If
                  '2005/12/6 END
                  'Added by Lydia 2020/12/16 外商臺灣案收文
                  If Left(frm010001.mRole, 2) = "F1" And txtSystem = "S" And txtOther(Index).Text <> 台灣國家代號 Then
                       MsgBox "外商臺灣案收文之申請國家必須為 000 !", vbCritical + vbOKOnly, MsgText(9001)
                       Cancel = True
                       Exit Sub
                  End If
                  'end 2020/12/16
                 
                  If CheckKeyIn(Index) <> -1 Then
                     CheckKeyIn 1
                     If Me.txtOther(12).Text <> "" Then CheckKeyIn 12
                     CheckKeyIn 16
                  Else
                     Cancel = True
                  End If
                  ' 92.11.3 by SONIA 更新費用與規費
                  If Cancel = False Then
                     OnUpdateFee  '變更國籍
                  End If
                  '92.11.3 END
            'Add By Cheng 2001/12/27
            Case 7 '申請人
               '若申請人有輸入才做Check動作
               If Len(Trim(Me.txtOther(Index).Text)) > 0 Then
                  If CheckKeyIn(Index) = -1 Then
                     Cancel = True
                  End If
               End If
            Case 8 '代理人
               '若代理人有輸入才做Check動作
               If Len(Trim(Me.txtOther(Index).Text)) > 0 Then
                  If CheckKeyIn(Index) = -1 Then
                     Cancel = True
                  End If
               End If
               '若申請人與代理人同時空白時
               If Len(Trim(Me.txtOther(7).Text)) <= 0 And Len(Trim(Me.txtOther(8).Text)) <= 0 Then
                  MsgBox "申請人與代理人必須至少輸入一項!!!", vbExclamation
'                  Cancel = True
               End If
            Case 25
                If Me.txtOther(Index).Text <> "" Then
                    If CheckKeyIn(Index) = -1 Then
                       Cancel = True
                       GoTo EXITSUB
                    End If
                    arrTM09 = Split(Me.txtOther(Index).Text, ",")
                    For ii = LBound(arrTM09) To UBound(arrTM09)
                        If Len(arrTM09(ii)) < 2 Or Len(arrTM09(ii)) > 3 Then
                            MsgBox "商品類別 <" & arrTM09(ii) & "> 不可小於二碼且不可大於三碼!!!", vbExclamation + vbOKOnly
                            Cancel = True
                            Exit For
                        End If
                    Next ii
                End If
                txtOther(Index).Text = Replace(txtOther(Index).Text, " ", "")
            Case 26
                If Me.txtOther(Index).Text <> "" Then
                    If CheckKeyIn(Index) = -1 Then
                       Cancel = True
                       GoTo EXITSUB
                    End If
                    'Modify By Sindy 2024/4/18 商品組群欄人員貼上資料後將全形或半形的「；」分號，轉為半形的逗號存入TM32。
                    Me.txtOther(Index).Text = Replace(Replace(Me.txtOther(Index).Text, ";", ","), "；", ",")
                    '2024/4/18 END
                    arrTM09 = Split(Me.txtOther(Index).Text, ",")
                    For ii = LBound(arrTM09) To UBound(arrTM09)
                        If Len(arrTM09(ii)) < 4 Or Len(arrTM09(ii)) > 6 Then
                            MsgBox "商品組群 <" & arrTM09(ii) & "> 不可小於四碼且不可大於六碼!!!", vbExclamation + vbOKOnly
                            Cancel = True
                            Exit For
                        End If
                    Next ii
                End If
                txtOther(Index).Text = Replace(txtOther(Index).Text, " ", "")
            'add by sonia 2019/8/6
            Case 9   '本所期限
               If CheckKeyIn(Index) = -1 Then
                  Cancel = True
               Else
                  'ACS 若本所期限非工作天則直接調整至最近的工作天
                  'Modified by Lydia 2020/07/07 本所期限檢查：所有系統類別的本所期限都要控制是工作日
                  'If txtOther(Index) <> "" And (txtSystem = "ACS") Then
                  If txtOther(Index) <> "" Then
                     txtOther(Index).Text = TransDate(PUB_GetWorkDay1(txtOther(Index).Text, True), 1)
                  End If
               End If
            'end 2019/8/6
            'Added by Lydia 2023/06/08
            Case 14 '法定期限
               '依林總指示，FCP、FG案件之本所期限由系統依輸入之法定期限計算，本所期限=法定期限-2工作天；收文畫面已調整本所期限與法定期限欄的先後順序。
               '不論新案或舊案收文，只要不是系統自動帶出期限的案件，收文人員輸入法定期限後，系統自動計算本所期限，並且限制不可修改；
               If txtSystem = "FG" And bolisNP0809 = False And (frm010001.intModifyKind = 0 Or frm010001.intModifyKind = 1) Then
                  If txtOther(Index) <> "" Then
                     txtOther(9).Locked = True '限制本所期限不可輸入
                     If txtOther(Index).Text <> txtOther(Index).Tag Then
                        strExc(1) = PUB_GetFCPOurDeadline(DBDATE(txtOther(Index)))
                        If strExc(1) < strSrvDate(1) Then strExc(1) = strSrvDate(1) '小於系統日=改用系統日
                        txtOther(9) = TransDate(strExc(1), 1)
                     End If
                  Else
                     txtOther(9).Locked = False
                  End If
               End If
               If CheckKeyIn(Index) <> 1 Then
                  Cancel = True
               End If
            'end 2023/06/08
            Case Else
                     If CheckKeyIn(Index) = -1 Then
                        Cancel = True
                     End If
End Select
EXITSUB:
If Cancel Then txtOther_GotFocus (Index)
End Sub

' 92.11.3 ADD BY SONIA
' 更新畫面中費用及規費的欄位內容
Private Sub OnUpdateFee()
    'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
    'Modify By Sindy 2011/1/18
    'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
    'modify by sonia 2014/9/11 取消X69514,已轉外專
    If Mid(txtOther(7), 1, 8) <> "X1484305" And Mid(txtOther(18), 1, 8) <> "X1484305" And Mid(txtOther(19), 1, 8) <> "X1484305" And Mid(txtOther(29), 1, 8) <> "X1484305" And Mid(txtOther(30), 1, 8) <> "X1484305" And _
       Mid(txtOther(7), 1, 8) <> "X3928904" And Mid(txtOther(18), 1, 8) <> "X3928904" And Mid(txtOther(19), 1, 8) <> "X3928904" And Mid(txtOther(29), 1, 8) <> "X3928904" And Mid(txtOther(30), 1, 8) <> "X3928904" Then
        If txtSystem = "FG" Then
           ' 規費
           txtOther(16) = GetPatentOfficialFee(txtSystem, txtOther(1), txtOther(14), "", txtOther(3), "")
           ' 費用
           txtOther(12) = Val(GetFCPFee(txtSystem, txtOther(1))) + Val(txtOther(16))
           '點數
           txtOther(13) = (Val(txtOther(12)) - Val(txtOther(16))) / 1000
        End If
    End If
End Sub
'92.11.3 END

Private Function CheckKeyIn(ByRef intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, bolIsChina As Boolean, strCusTemp As String
Static strLastCus As String
'Add By Cheng 2003/06/12
Dim nCount As Integer
Dim nIndex As Integer
Dim strTit As String
Dim strMsg As String
'Dim strTemp As String
Dim nResponse

CheckKeyIn = -1
Select Case intIndex
             'add by nickc 2006/11/30
             Case 25
                       If CheckLengthIsOK(txtOther(intIndex), 395) Then
                          CheckKeyIn = 1
                       End If
             Case 26
                       If CheckLengthIsOK(txtOther(intIndex), 349) Then
                          CheckKeyIn = 1
                       End If
             'Modified by Lydia 2021/02/19 改成依欄寬設定
             'Case 4, 6
             '          '92.11.5 MODIFY BY SONIA
             '          'Modify By Sindy 2009/07/24 增加LIN系統類別
             '          'modify by sonia 2019/7/29 +ACS系統類別
             '          If txtSystem = "L" Or txtSystem = "CFL" Or txtSystem = "FCL" Or txtSystem = "LA" Or _
             '              txtSystem = "LIN" Or txtSystem = "ACS" Then
             '              If CheckLengthIsOK(txtOther(intIndex), 40) Then
             '                 CheckKeyIn = 1
             '              End If
             '          Else
             '              If CheckLengthIsOK(txtOther(intIndex), 60) Then
             '                 CheckKeyIn = 1
             '              End If
             '          End If
             '          '92.11.5 END
             '          If intIndex = 6 Then
             '              If txtOther(4) = "" And txtOther(5) = "" And txtOther(6) = "" Then
             '                 ShowMsg MsgText(1031)
             '                 intIndex = 4
             '                 CheckKeyIn = 0
             '              ElseIf CheckLengthIsOK(txtOther(intIndex), 60) Then
             '                 CheckKeyIn = 1
             '              End If
             '          End If
             'Case 5
             '          If CheckLengthIsOK(txtOther(intIndex), 60) Then
             '             CheckKeyIn = 1
             '          End If
             Case 4, 5, 6 '案件名稱（中／英／日）
                       If CheckLengthIsOK(txtOther(intIndex), txtOther(intIndex).MaxLength) Then
                          CheckKeyIn = 1
                       End If
                       If intIndex = 6 Then
                           If txtOther(4) = "" And txtOther(5) = "" And txtOther(6) = "" Then
                              ShowMsg MsgText(1031)
                              intIndex = 4
                              CheckKeyIn = 0
                           End If
                       End If
             'end 2021/02/19
             Case 0
                        If CheckIsTaiwanDate(txtOther(intIndex).Text) Then
                            CheckKeyIn = 1
                        End If
             Case 1
                       If txtOther(3) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
                       'edit by nickc 2007/02/02 不用 dll 了
                       'If objPublicData.GetCaseProperty(txtSystem, txtOther(intIndex), strTemp, bolIsChina) Then
                       If ClsPDGetCaseProperty(txtSystem, txtOther(intIndex), strTemp, bolIsChina) Then
                           lblCaseProperty.Caption = strTemp
                           CheckKeyIn = 1
                        End If
                        Call setFrame21 'Added by Lydia 2022/07/15 TC案之文件齊備日管制
             Case 2
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCaseSource(txtOther(intIndex).Text, strTemp) Then
                        If ClsPDGetCaseSource(txtOther(intIndex).Text, strTemp) Then
                           lblCaseSource.Caption = strTemp
                           CheckKeyIn = 1
                        End If
             Case 3
                        If frm010001.intCaseKind <> 顧問 Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetNation(txtOther(intIndex).Text, strTemp) Then
                           If ClsPDGetNation(txtOther(intIndex).Text, strTemp) Then
                              lblNation.Caption = strTemp
                              CheckKeyIn = 1
                           End If
                           '92.11.6 add by sonia
                           If txtSystem = "FG" And txtOther(intIndex).Text <> 台灣國家代號 Then
                              ShowMsg MsgText(9219)
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           '92.11.6 END
                        Else
                           CheckKeyIn = 1
                        End If
                        If Val(txtOther(intIndex)) >= 1 And Val(txtOther(intIndex)) <= 8 Then
                           ShowMsg MsgText(38)
                           CheckKeyIn = -1
                        End If
             Case 7 '申請人/當事人1
                        If txtOther(intIndex) = "" Then
                           CheckKeyIn = 1
                           Exit Function
                        End If
                        If intIndex = 7 Then
                           If txtOther(intIndex) = txtOther(18) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(19) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           'Add By Sindy 2011/1/18
                           If txtOther(intIndex) = txtOther(29) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(30) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           '2011/1/18 End
                        End If
                        
                        'Added by Lydia 2020/05/20 法律所案源收文：檢查介紹客戶是否為申請人1
                        If m_LOS05 <> "" And txtOther(7) <> "" And ChangeCustomerS(m_LOS05) <> ChangeCustomerS(txtOther(7)) Then
                            MsgBox "申請人1請輸入 " & ChangeCustomerS(m_LOS05), vbExclamation, "檢查介紹客戶"
                            txtOther(7).SetFocus
                            txtOther_GotFocus 7
                            CheckKeyIn = -1
                            Exit Function
                        End If
                        'end 2020/05/20
                        
                        strCusTemp = txtOther(intIndex)
                        'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
                        'If objPublicData.GetCustomer(strCusTemp, strTemp, strTemp1) Then
                        'Modify By Sindy 2015/8/27 +txtSystem
                        'Modified by Lydia 2023/03/06 傳入本所案號 , , , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
                        If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, , , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) Then
                           txtOther(intIndex) = strCusTemp
                           lblPetition(0).Caption = strTemp
                           If strLastCus <> strCusTemp Or txtOther(11).Text = "" Then
                              txtOther(11).Text = strTemp1
                              strLastCus = strCusTemp
                           End If
                           CheckKeyIn = 1
                           'Add by Morgan 2008/8/5
                           If ChangeCustomerL(strCusTemp) <> strAppNo1 Then
                              strAppNo1 = ChangeCustomerL(strCusTemp)
                              'Modify by Amy 2021/12/21 改成Form 2.0
                              'PUB_AddContact strAppNo1, cboContact, , True
                              strExc(10) = cboContact.Tag
                              'Added by Lydia 2022/11/25 區分有無輸入接洽人; ex.P-130652接洽人不是客戶預設接洽人
                              If cboContact.Text <> "" Then
                                 strExc(9) = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
                                 PUB_AddContact strAppNo1, cboContact, strExc(9), True, True, strExc(10)
                              Else
                              'end 2022/22/25
                                  PUB_AddContact strAppNo1, cboContact, , True, True, strExc(10)
                              End If  'Added by Lydia 2022/11/25
                              cboContact.Tag = strExc(10)
                           End If
                        End If
                        'Modify By Cheng 2001/12/27
                        '若上項客戶代號檢查無誤, 則繼續檢查客戶之國籍
                        If CheckKeyIn = 1 Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetCustomerNation(strCusTemp, strNation) Then
                           If ClsPDGetCustomerNation(strCusTemp, strNation) Then
                           'If strNation >= "010" Then
                           '   txtOther(15) = "N"
                           'Else
                           '   txtOther(15) = ""
                           'End If
                           End If
                        End If
                        'Add By Cheng 2003/09/08
                        If CheckKeyIn = 1 Then
                            '2010/9/30 modify by sonia 新增時才要檢查
                            'If frm010001.m_blnNewCase = True Then
                            If frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 Then
                                '若輸入9碼且最後一碼不為"0"
                                If Len(Me.txtOther(intIndex).Text) = 9 And Right(Me.txtOther(intIndex).Text, 1) <> "0" Then
                                    MsgBox "此申請人/當事人已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
                                    CheckKeyIn = -1
                                End If
                            End If
                        End If
             Case 8 '代理人
                        strCusTemp = txtOther(intIndex)
                        If txtOther(intIndex) <> "" Then
                           'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
                           'If objPublicData.GetAgent(strCusTemp, strTemp) Then
                           'Modify By Sindy 2015/8/27 +txtSystem
                           If GetAgentAndState(strCusTemp, strTemp, , , , txtSystem) Then
                              txtOther(intIndex) = strCusTemp
                              lblAgent.Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
                        'Add By Cheng 2003/09/08
                        If CheckKeyIn = 1 Then
                            If frm010001.m_blnNewCase = True Then
                                '若輸入9碼且最後一碼不為"0"
                                If Len(Me.txtOther(intIndex).Text) = 9 And Right(Me.txtOther(intIndex).Text, 1) <> "0" Then
                                    MsgBox "此代理人已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
                                    CheckKeyIn = -1
                                End If
                            End If
                        End If
             Case 9
                        If txtOther(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           If CheckIsTaiwanDate(txtOther(intIndex).Text) Then
                               If CheckReKey(txtOther(intIndex)) Then
                                  '93.7.5 cancel by sonia
                                  'If Val(txtOther(intIndex)) = GetTaiwanTodayDate Then
                                  '   ShowMsg "此案件已屆本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
                                  'End If
                                  'If Val(txtOther(intIndex)) < GetTaiwanTodayDate Then
                                  '   ShowMsg "此案件已逾本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
                                  'End If
                                  '93.7.5 end
                                  CheckKeyIn = 1
                               Else
                                  CheckKeyIn = 0
                               End If
                            End If
                        End If
             Case 14
                        If txtOther(intIndex) <> "" Then
                           If CheckIsTaiwanDate(txtOther(intIndex).Text) Then
                              If Val(txtOther(9)) <= Val(txtOther(14)) Then
                                 If CheckReKey(txtOther(intIndex)) Then
                                    CheckKeyIn = 1
                                 Else
                                    CheckKeyIn = 0
                                 End If
                              Else
                                 ShowMsg MsgText(1033)
                              End If
                           End If
                        ElseIf txtOther(9) <> "" Then
                           ShowMsg MsgText(1033)
                        Else
                           CheckKeyIn = 1
                        End If
             Case 10
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetStaff(txtOther(intIndex).Text, strTemp, strTemp1) Then
                        If ClsPDGetStaff(txtOther(intIndex).Text, strTemp, strTemp1) Then
                           CheckKeyIn = 1
                        End If
                        lblSales.Caption = strTemp
                        
                        'Modified by Lydia 2019/02/14
                        'strTemp = GetST15(txtOther(intIndex).Text, strTemp1)
                        'Modifeid by Lydia 2019/09/16
                        'm_SalesST15 = GetST15(txtOther(intIndex).Text, strTemp1)
                        m_SalesST15 = GetST15(txtOther(intIndex).Text, strTemp1, , m_SalesST06)
                        
                        lblDepartment = strTemp1
'             Case 12
'                'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'                If Mid(txtOther(7), 1, 8) <> "X1484305" And Mid(txtOther(18), 1, 8) <> "X1484305" And Mid(txtOther(19), 1, 8) <> "X1484305" Then
''                        If objPublicData.GetCaseLowPrice(txtSystem, txtOther(4), txtOther(1), douStPrice, douLowPrice) = 1 Then
'                        'edit by nickc 2007/02/02 不用 dll 了
'                        'If objPublicData.GetCaseLowPrice(txtSystem, txtOther(3), txtOther(1), douStPrice, douLowPrice) = 1 Then
'                        If ClsPDGetCaseLowPrice(txtSystem, txtOther(3), txtOther(1), douStPrice, douLowPrice) = 1 Then
'                        End If
'                        'If txtOther(intIndex) <> "" Then
''                           If objPublicData.GetCaseFee(txtSystem, txtOther(4), txtOther(1), Val(txtOther(intIndex))) = 1 Then
'                           'edit by nickc 2007/02/02 不用 dll 了
'                           'If objPublicData.GetCaseFee(txtSystem, txtOther(3), txtOther(1), Val(txtOther(intIndex))) = 1 Then
'                           If ClsPDGetCaseFee(txtSystem, txtOther(3), txtOther(1), Val(txtOther(intIndex))) = 1 Then
'                              CheckKeyIn = 1
'                           End If
'                        'ElseIf txtOther(16) <> "" Then
'                        '   ShowMsg MsgText(1034)
'                        '   CheckKeyIn = 0
'                        'Else
'                        '   CheckKeyIn = 1
'                        'End If
'                    Else
'                        CheckKeyIn = 1
'                    End If
'             Case 13
'                'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'                If Mid(txtOther(7), 1, 8) <> "X1484305" And Mid(txtOther(18), 1, 8) <> "X1484305" And Mid(txtOther(19), 1, 8) <> "X1484305" Then
'                        If txtOther(intIndex) = "" Then
'                           If txtOther(12) <> "" And txtOther(16) <> "" Then
'                              ShowMsg MsgText(1035)
'                              CheckKeyIn = 0
'                           Else
'                              CheckKeyIn = 1
'                           End If
'                        'MODIFY BY SONIA 90.10.12
'                        'ElseIf txtOther(12) <> "" And txtOther(16) <> "" Then
'                        ElseIf txtOther(12) <> "" Then
'
'                        'If txtOther(12) <> "" Then
'                           If Format((Val(txtOther(12)) - Val(txtOther(16))) / 1000, "0.0") <> Format(Val(txtOther(13)), "0.0") Then
''                              ShowMsg MsgText(1036)
'                              CheckKeyIn = 0
'                           Else
'                              CheckKeyIn = 1
'                           End If
'                        Else
'                           ShowMsg MsgText(1037)
'                           CheckKeyIn = -1
'                        End If
'                Else
'                    CheckKeyIn = 1
'                End If
             Case 15
                        'If strNation >= "010" Then
                        '   If txtOther(15) <> "N" Then
                        '      ShowMsg "申請人國籍非台灣時, 是否開電腦收據必須為 N"
                        '      CheckKeyIn = -1
                        '      Exit Function
                        '   End If
                        'End If
                        If txtOther(intIndex) = "" Or txtOther(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
'             Case 16 '規費
'                    'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'                    If Mid(txtOther(7), 1, 8) <> "X1484305" And Mid(txtOther(18), 1, 8) <> "X1484305" And Mid(txtOther(19), 1, 8) <> "X1484305" Then
'                        '2009/6/12 modify by sonia取消有輸入才檢查的限制
'                        'If txtOther(intIndex) <> "" Then              '2009/6/12 cancel by sonia
'                           '91.11.21 CANCEL BY SONIA
'                           'If txtOther(12) = "" Then
'                           '   ShowMsg MsgText(1039)
'                           'ElseIf objPublicData.GetCaseMoney(txtSystem, txtOther(3), txtOther(1), Val(txtOther(intIndex))) = 1 Then
'                           'edit by nickc 2006/12/05 change call basquery
'                           'If objPublicData.GetCaseMoney(txtSystem, txtOther(3), txtOther(1), Val(txtOther(intIndex))) = 1 Then
'                           If GetCaseMoney(txtSystem, txtOther(3), txtOther(1), Val(txtOther(intIndex))) = 1 Then
'                              CheckKeyIn = 1
'                           End If
'                        'Else                                           '2009/6/12 cancel by sonia
''                           If txtOther(12) <> "" Then
''                              ShowMsg MsgText(1040)
''                              CheckKeyIn = 0
''                           Else
'                        '      CheckKeyIn = 1                           '2009/6/12 cancel by sonia
''                           End If
'                        'End If                                         '2009/6/12 cancel by sonia
'                    Else
'                        CheckKeyIn = 1
'                    End If
             Case 18, 19, 29, 30 '申請人/當事人2,3,4,5
                        If txtOther(intIndex) = "" Then
                           CheckKeyIn = 1
                           Exit Function
                        End If
                        If intIndex = 18 Then
                           If txtOther(intIndex) = txtOther(7) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(19) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           'Add By Sindy 2011/1/18
                           If txtOther(intIndex) = txtOther(29) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(30) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           '2011/1/18 End
                        End If
                        If intIndex = 19 Then
                           If txtOther(intIndex) = txtOther(7) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(18) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           'Add By Sindy 2011/1/18
                           If txtOther(intIndex) = txtOther(29) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(30) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           '2011/1/18 End
                        End If
                        'Add By Sindy 2011/1/18
                        If intIndex = 29 Then
                           If txtOther(intIndex) = txtOther(7) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(18) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(19) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(30) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        '2011/1/18 End
                        'Add By Sindy 2011/1/18
                        If intIndex = 30 Then
                           If txtOther(intIndex) = txtOther(7) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(18) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(19) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtOther(intIndex) = txtOther(29) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        '2011/1/18 End
                        If txtOther(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           strCusTemp = txtOther(intIndex)
                           'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
                           'If objPublicData.GetCustomer(strCusTemp, strTemp, strTemp1) Then
                           'Modify By Sindy 2015/8/27 +txtSystem
                           'Modified by Lydia 2023/03/06 傳入本所案號 , , , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
                           If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, , , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) Then
                              txtOther(intIndex) = strCusTemp
                              Select Case intIndex
                                 Case 18, 19
                                    lblPetition(intIndex - 17).Caption = strTemp
                                 Case 29, 30
                                    lblPetition(intIndex - 26).Caption = strTemp
                              End Select
                              CheckKeyIn = 1
                           End If
                        End If
                        'Add By Cheng 2003/09/08
                        If CheckKeyIn = 1 Then
                            '2010/9/30 modify by sonia 新增時才要檢查
                            'If frm010001.m_blnNewCase = True Then
                            If frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 Then
                                '若輸入9碼且最後一碼不為"0"
                                If Len(Me.txtOther(intIndex).Text) = 9 And Right(Me.txtOther(intIndex).Text, 1) <> "0" Then
                                    MsgBox "此申請人/當事人已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
                                    CheckKeyIn = -1
                                End If
                            End If
                        End If
             Case 20
                        If txtOther(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetStaff(txtOther(intIndex), strTemp) Then
                        If ClsPDGetStaff(txtOther(intIndex), strTemp) Then
                           lblPromoter = strTemp
                           CheckKeyIn = 1
                        End If
                        End If
             'add by nickc 2005/10/06 加長分所號
             Case 23
                        If CheckLengthIsOK(txtOther(intIndex), 50) Then
                            CheckKeyIn = 1
                        End If
             'add by nickc 2008/05/02 加預定收款日
             Case 28
                        If txtOther(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           If CheckIsTaiwanDate(txtOther(intIndex).Text) Then
                                'edit by nickc 2008/05/21 原先判斷系統日，改成判斷收文日，因為分所會去補分所號
                                'If DBDATE(txtOther(intIndex).Text) >= strSrvDate(1) Then
                                If DBDATE(txtOther(intIndex).Text) >= DBDATE(txtOther(0).Text) Then
                                   CheckKeyIn = 1
                                Else
                                    'edit by nickc 2008/05/21 原先判斷系統日，改成判斷收文日，因為分所會去補分所號
                                    'MsgBox "預定收款日必須>= 系統日", vbOKOnly + vbCritical, "輸入錯誤！"
                                    MsgBox "預定收款日必須>= 收文日", vbOKOnly + vbCritical, "輸入錯誤！"
                                End If
                           End If
                        End If
             '2014/3/12 add by sonia
             Case 24
                  If txtOther(intIndex) = "" Then
                     ShowMsg "案件名稱不可空白"
                     CheckKeyIn = 0
                  ElseIf CheckLengthIsOK(txtOther(intIndex), 140) Then
                     CheckKeyIn = 1
                  End If
             '2014/3/12 end
             Case Else
                        CheckKeyIn = 1
End Select
End Function

'Modify by Amy 2021/12/20 原:Integer
Private Sub txtOther_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
             'Modify By Sindy 2011/1/18 +29,30
             Case 7, 8, 10, 15, 18, 19, 20, 29, 30
                       KeyAscii = UpperCase(KeyAscii)
             Case 11
                       'Modify by Amy 2021/12/20 +txtOther(Index)
                       KeyAscii = ChangeZIP(KeyAscii, txtOther(Index))
End Select
End Sub

Private Sub txtOther_GotFocus(Index As Integer)
txtOther(Index).SelStart = 0
txtOther(Index).SelLength = Len(txtOther(Index).Text)
'儲存未修改前之值至Tag中,供再確認時使用
txtOther(Index).Tag = txtOther(Index)
'切換輸入法
Select Case Index
'             Case 4
             Case 4, 24, 31 '案件中文名稱, 案件名稱, 主題
                        'edit by nickc 2007/06/06
                        'txtOther(Index).IMEMode = 1
                        OpenIme
             Case Else
                        'edit by nickc 2007/06/06
                        'txtOther(Index).IMEMode = 2
                        CloseIme
End Select
End Sub

Private Sub txtOther_LostFocus(Index As Integer)
'關閉輸入法
'edit by nickc 2007/06/06
'txtOther(Index).IMEMode = 2
'CloseIme 'Removed by Morgan 2016/10/20 會造成 Win7 的切換錯誤
'Add By Cheng 2001/12/27
If Index = 8 And Len(Trim(Me.txtOther(7).Text)) <= 0 And Len(Trim(Me.txtOther(8).Text)) <= 0 Then
   '若申請人與代理人皆未輸入, 則將游標設定在申請人欄位
   Me.txtOther(7).SetFocus
End If
End Sub

'新增至Other資料庫
Private Function InsertOtherDatabase(ByRef intSaveMode As Integer, ByRef intCaseKind As Integer, ByRef ot01 As String, _
             ByRef ot02 As String, ByRef ot03 As String, ByRef ot04 As String, ByRef ot05 As String, _
             ByRef ot06 As String, ByRef ot07 As String, ByRef ot08 As String, ByRef ot09 As String, ByRef ot10 As String, _
             ByRef ot11 As String, ByRef ot12 As String, ByRef ot18 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, ByRef cp11 As String, _
             ByRef cp13 As String, ByRef cp16 As String, ByRef cp17 As String, ByRef cp18 As String, ByRef cp19 As String, _
             ByRef cp32 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef CP09 As String, ByRef cp02 As String, ByRef cp33 As Double, ByRef cp34 As Double, ByRef CP64 As String, ByRef SP73 As String, ByRef SP74 As String, ByRef SP27 As String, ByRef SP78 As String, ByRef ot13 As String, ByRef ot14 As String) As Boolean

Dim strSql As String, sp34 As String, cp31 As String, strAutoNumber As String
Dim np13 As String, np14 As String, bolRt As Boolean, isp34 As Integer
Dim bolError As Boolean
Dim adoquery As New ADODB.Recordset
Dim cp48 As String 'Add by Morgan 2008/8/23
Dim cp20 As String 'Add by Morgan 2010/4/29
Dim cp12 As String 'add by sonia  2015/5/12
Dim strCusReceipt As String 'Add by Amy 2018/10/11 收據公司別
Dim iRound As Integer 'Added by Lydia 2020/05/20

'add by nickc 2007/12/12
If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand
   '傳入0為重複之本所案號(新增舊案)，1為正確之本所案號(新增新案)
   cp05 = ChangeTStringToWString(cp05)
   cp06 = ChangeTStringToWString(cp06)
   cp07 = ChangeTStringToWString(cp07)
   ot09 = ChangeCustomerL(ot09) '當事人1
   ot10 = ChangeCustomerL(ot10) '當事人2
   ot11 = ChangeCustomerL(ot11) '當事人3
   ot12 = ChangeCustomerL(ot12)
   'Add By Sindy 2011/1/18
   ot13 = ChangeCustomerL(ot13) '當事人4
   ot14 = ChangeCustomerL(ot14) '當事人5
   '2011/1/18 End
   'edit by nickc 2007/02/06 不用 dll 了
   'Dim objPublicData As Object
   'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
   cnnConnection.BeginTrans
   If intSaveMode = 1 Then
      If ot02 = "" Then
         'edit by nickc 2007/02/06 不用 dll 了
         'If objPublicData.GetAutoNumber(ot01, strAutoNumber, True, False) Then
         If ClsPDGetAutoNumber(ot01, strAutoNumber, True, False) Then
            ot02 = strAutoNumber
         Else
            bolError = True
         End If
      End If
      If bolError = False Then
         cp02 = ot02
         'Add by Amy 2018/10/11 收據公司別
         'Modified by Lydia 2021/07/12 排除法務案; ex.L-006408已將LC48=J公司拿掉
         'If intCaseKind <> 顧問 Then
         If txtSystem <> "L" And txtSystem <> "LA" And txtSystem <> "LIN" And txtSystem <> "FCL" And txtSystem <> "CFL" Then
            strCusReceipt = GetReceiptCmp(Mid(ChangeCustomerL(txtOther(7)), 1, 8), Mid(ChangeCustomerL(txtOther(7)), 9, 1), txtSystem, txtOther(3))
         End If
         'end 2018/10/11
         Select Case intCaseKind
                      Case 法務
                           'Modify by Morgan 2008/8/5 +LC42
                           'Modify By Sindy 2011/1/18 +lc43,lc44,lc45,lc46
                           'Modify by Amy 2018/10/11 +lc48收據公司別
                           strSql = "insert into lawcase (lc01,lc02,lc03,lc04,lc05,lc06,lc07,lc11,lc15,lc22,lc23,lc42,lc43,lc44,lc45,lc46,lc48) " + _
                               "values (" + CNULL(ot01) + "," + CNULL(ot02) + "," + CNULL(ot03) + "," + CNULL(ot04) + "," + CNULL(ChgSQL(ot06)) + "," + _
                               CNULL(ChgSQL(ot07)) + "," + CNULL(ChgSQL(ot08)) + "," + CNULL(ot09) + "," + CNULL(ot05) + "," + CNULL(ot12) + "," + _
                               CNULL(SP27) + "," + CNULL(SP78) + "," + CNULL(ot10) + "," + CNULL(ot11) + "," + CNULL(ot13) + "," + CNULL(ot14) + "," + CNULL(strCusReceipt) + ")"
                           cnnConnection.Execute strSql
                      Case 顧問
                           'Modify by Morgan 2008/8/5 +hC23
                           'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
                           strSql = "insert into hirecase (hc01,hc02,hc03,hc04,hc05,hc06,hc23,hc24,hc25,hc26,hc27) values (" + _
                               CNULL(ot01) + "," + CNULL(ot02) + "," + CNULL(ot03) + "," + CNULL(ot04) + "," + _
                               CNULL(ot09) + "," + CNULL(ChgSQL(ot06)) + "," + CNULL(ChgSQL(SP78)) + "," + _
                               CNULL(ot10) + "," + CNULL(ot11) + "," + CNULL(ot13) + "," + CNULL(ot14) + ")"
                           cnnConnection.Execute strSql
                      Case Else
                           'edit by nickc 2007/02/06 不用 dll 了
                           'If objPublicData.GetSystemKind(ot01, , , isp34) Then
                           If ClsPDGetSystemKind(ot01, , , isp34) Then
                              sp34 = IIf(isp34 = 2, 2, 1)
                              'edit by nickc 2007/03/27 加入彼所案號
                              'strSQL = "insert into servicepractice (sp01,sp02,sp03,sp04,sp05,sp06,sp07,sp08,sp58,sp59,sp09,sp26, sp18,sp73,sp74) " + _
                                  "values (" + CNULL(ot01) + "," + CNULL(ot02) + "," + CNULL(ot03) + "," + CNULL(ot04) + "," + _
                                  CNULL(ChgSQL(ot06)) + "," + CNULL(ChgSQL(ot07)) + "," + CNULL(ChgSQL(ot08)) + "," + CNULL(ot09) + "," + CNULL(ot10) + _
                                  "," + CNULL(ot11) + "," + CNULL(ot05) + "," + CNULL(ot12) + "," + CNULL(ChgSQL(ot18)) + "," + CNULL(SP73) + "," + CNULL(SP74) + ")"
                              'Modify by Morgan 2008/8/5 +SP78
                              'Modify By Sindy 2010/3/8 增加聯絡人sp30,sp75欄位
                              If bolCancel = False Then
                                 strSP30s = "": strSP75s = ""
                              End If
   '                           strSql = "insert into servicepractice (sp01,sp02,sp03,sp04,sp05,sp06,sp07,sp08,sp58,sp59,sp09,sp26, sp18,sp73,sp74,sp27,SP78) " + _
   '                               "values (" + CNULL(ot01) + "," + CNULL(ot02) + "," + CNULL(ot03) + "," + CNULL(ot04) + "," + _
   '                               CNULL(ChgSQL(ot06)) + "," + CNULL(ChgSQL(ot07)) + "," + CNULL(ChgSQL(ot08)) + "," + CNULL(ot09) + "," + CNULL(ot10) + _
   '                               "," + CNULL(ot11) + "," + CNULL(ot05) + "," + CNULL(ot12) + "," + CNULL(ChgSQL(ot18)) + "," + CNULL(SP73) + "," + CNULL(SP74) + "," + CNULL(SP27) + "," + CNULL(SP78) + ")"
                              'Modify By Sindy 2011/1/18 +sp65,sp66
                              'Moidfy by Amy 2018/10/11 +sp85收據公司別
                              strSql = "insert into servicepractice (sp01,sp02,sp03,sp04,sp05,sp06,sp07,sp08,sp58,sp59,sp65,sp66,sp09,sp26, sp18,sp73,sp74,sp27,SP78,SP30,SP75,SP85) " + _
                                  "values (" + CNULL(ot01) + "," + CNULL(ot02) + "," + CNULL(ot03) + "," + CNULL(ot04) + "," + _
                                  CNULL(ChgSQL(ot06)) + "," + CNULL(ChgSQL(ot07)) + "," + CNULL(ChgSQL(ot08)) + "," + CNULL(ot09) + "," + CNULL(ot10) + _
                                  "," + CNULL(ot11) + "," + CNULL(ot13) + "," + CNULL(ot14) + "," + CNULL(ot05) + "," + CNULL(ot12) + "," + CNULL(ChgSQL(ot18)) + "," + CNULL(SP73) + "," + CNULL(SP74) + "," + CNULL(SP27) + "," + _
                                  CNULL(SP78) + "," + CNULL(ChgSQL(strSP30s)) + "," + CNULL(ChgSQL(strSP75s)) + "," + CNULL(ChgSQL(strCusReceipt)) + ")"
                              cnnConnection.Execute strSql
                           Else
                              bolError = True
                           End If
         End Select
         cp31 = "Y"
      Else
         bolError = True
      End If
   End If
   If bolError = False Then
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetAutoNumber(Left(CP09, 1), strAutoNumber, True, True) Then
      If ClsPDGetAutoNumber(Left(CP09, 1), strAutoNumber, True, True) Then
         CP09 = CP09 + strAutoNumber
         'edit by nickc 2007/02/06 不用 dll 了
         'bolRt = obj001.GetNextProgressData(ot01, ot02, ot03, ot04, CP10, np13, np14)
         bolRt = Cls001GetNextProgressData(ot01, ot02, ot03, ot04, CP10, np13, np14)
         'add by nick 2005/01/07 從上面搬下來
         Select Case intCaseKind
         Case 法務
                  If txtOther(23).Text & txtOther(22).Text <> "" Then
                       strSql = "Update Lawcase Set LC16='" & ChgSQL(Me.txtOther(23).Text) & "', LC17='" & ChgSQL(Me.txtOther(22).Text) & "' Where lc01=" + CNULL(ot01) + " and lc02=" + CNULL(ot02) + " and lc03=" + CNULL(ot03) + " and lc04=" + CNULL(ot04)
                       cnnConnection.Execute strSql
                  End If
         Case 顧問
                  If txtOther(23).Text <> "" Then
                       strSql = "Update Hirecase Set HC07='" & ChgSQL(Me.txtOther(23).Text) & "' Where HC01=" + CNULL(ot01) + " and HC02=" + CNULL(ot02) + " and HC03=" + CNULL(ot03) + " and HC04=" + CNULL(ot04)
                       cnnConnection.Execute strSql
                  End If
         Case Else
                 If txtOther(23).Text & txtOther(22).Text <> "" Then
                       strSql = "Update Servicepractice Set SP28='" & ChgSQL(Me.txtOther(23).Text) & "', SP29='" & ChgSQL(Me.txtOther(22).Text) & "' Where SP01=" + CNULL(ot01) + " and SP02=" + CNULL(ot02) + " and SP03=" + CNULL(ot03) + " and SP04=" + CNULL(ot04)
                       cnnConnection.Execute strSql
                 End If
         End Select
         'add end by nick 2005/01/07
         
         'Add by Morgan 2008/9/23
         If Me.txtSystem.Text = "FG" Then
            cp48 = Pub_GetHandleDay("FG", "000", CP10, , cp06)
            cp20 = PUB_GetCP20(txtSystem, CP10) 'Add by Morgan 2010/4/29
         End If
         
         'Add By Sindy 2012/11/06 有★★的應收帳款簽核控管
         m_CP150 = ""
         If Check2.Value = 1 Then m_CP150 = "Y"
         '2012/11/06 End
         
         'Modify By Sindy 2012/11/06 +CP150
         If bolRt Then
            strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," + _
              "cp16,cp17,cp18,cp19,cp31,cp32,cp40,cp33,cp34,CP64,cp48,cp20,CP150) values (" + CNULL(ot01) + "," + CNULL(ot02) + "," + CNULL(ot03) + "," + CNULL(ot04) + "," + CNULL(cp05) + "," + _
              CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(np13) + "," + CNULL(CP09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
              CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(ChgSQL(np14)) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(CP64)) + "," & CNULL(cp48, True) & ",'" & cp20 & "'," + CNULL(m_CP150) + ")"
         Else
            strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp13,cp14," + _
              "cp16,cp17,cp18,cp19,cp31,cp32,cp33,cp34,CP64,cp48,cp20,CP150) values (" + CNULL(ot01) + "," + CNULL(ot02) + "," + CNULL(ot03) + "," + CNULL(ot04) + "," + CNULL(cp05) + "," + _
              CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(CP09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
              CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(CP64)) + "," & CNULL(cp48, True) & ",'" & cp20 & "'," + CNULL(m_CP150) + ")"
         End If
         cnnConnection.Execute strSql, intI
         'MODIFY BY SONIA 2015/5/12 國外部收文之FCL,CFL,LIN案,承辦人都預設系統特殊人員U2(桂所長)
         'strSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(cp13) + ") where cp09=" + CNULL(CP09)
         cp12 = PUB_GetStaffST15(cp13, 1)
         strSql = "update caseprogress set cp12=" & CNULL(cp12) & " where cp09=" + CNULL(CP09)
         cnnConnection.Execute strSql
         If (ot01 = "FCL" Or ot01 = "CFL" Or ot01 = "LIN") And Left(cp12, 1) = "F" Then
            strSql = "update caseprogress set cp14='" & Pub_GetSpecMan("U2") & "' where cp09=" + CNULL(CP09)
            cnnConnection.Execute strSql
         End If
'         '2015/5/12 END
         
         'Added by Lydia 2020/05/20 法律所案源收文：(著作權)台灣案B1、B2及C收文時，增加"案源單號"欄位一定要輸入，並將案源單號更新至該筆收文的CP162。
         If frm010001.intModifyKind = 0 And txtOther(3) = "000" And txtSystem = "TC" And m_LOS02 <> "" And m_LOS15 <> "" Then
              If Left(m_LOS02, 1) = "B" Or Left(m_LOS02, 1) = "C" Then
                  strSql = "update caseprogress set CP162='" & m_LOS15 & "' where cp09='" & CP09 & "' "
                  cnnConnection.Execute strSql
              End If
         End If
         'end 2020/05/20
         
         'Added by Lydia 2020/05/20 法律所案源收文：存檔時案源單號存CP162、案源總收文號(LOS01)存CP64欄"案源：本所案號(總收文號)
         If strSrvDate(1) >= 法律所案源收文啟用日 And InStr(txtSystem, "L") > 0 And m_LOS15 <> "" Then
             strSql = " "
             If m_LOS01 <> "" And m_LOS01cp01 <> "" Then
                 strSql = ",cp64=" & IIf(CP64 <> "", "cp64||';'||", "") & CNULL("案源：" & m_LOS01cp01 & "-" & m_LOS01cp02 & IIf(m_LOS01cp03 <> "0", "-" & m_LOS01cp03, "") & IIf(m_LOS01cp04 <> "00", "-" & m_LOS01cp04, "") & "(" & m_LOS01 & ");")
             End If
             strSql = "update caseprogress set CP162=" & CNULL(m_LOS15) & strSql & " where cp09=" & CNULL(CP09)
             cnnConnection.Execute strSql
            
             'Added by Lydia 2020/06/24 法務補收款78輸入案源單號若為B1類表示為A4轉B1，回寫收文號至案源檔的法律所總收文號欄2。
             If m_LOS02 = "B1" And txtOther(1) = "78" Then
                strSql = "update LawOfficeSource set los21='" & CP09 & "' where los21 is null and los15=" & CNULL(frm010001.txtLOS15)
                cnnConnection.Execute strSql, intI
             Else
             'end 2020/06/24
             '並回寫收文號至案源檔的法律所總收文號欄。
             '5/26 若輸入之案源單號已有法律所總收文號且為同案號同日收文者，則為同一接洽單之其他性質。
                strSql = "update LawOfficeSource set los06='" & CP09 & "' where los06 is null and los15=" & CNULL(frm010001.txtLOS15)
                cnnConnection.Execute strSql, intI
             End If 'Added by Lydia 2020/06/24
             
             'Added by Lydia 2020/06/10 更新卷宗區客戶文件(CPP01="LOS"+LOS15)的總收文號為法律案接洽單號(LOS17)，檔名也要一併更正。
             If m_LOS02 <> "" Then
                strSql = ""
                Select Case Left(m_LOS02, 1)
                    Case "A"
                        strSql = ", cpp02=replace(cpp02,'TT999999.735.','" & ot01 & ot02 & IIf(ot03 <> "0", ot03, "") & IIf(ot04 <> "00", ot04, "") & "." & CP10 & ".') "
                    Case "B"
                        strSql = ", cpp02=replace(cpp02,'TT999999.736.','" & ot01 & ot02 & IIf(ot03 <> "0", ot03, "") & IIf(ot04 <> "00", ot04, "") & "." & CP10 & ".') "
                    Case "C"
                        strSql = ", cpp02=replace(cpp02,'TT999999.','" & ot01 & ot02 & IIf(ot03 <> "0", ot03, "") & IIf(ot04 <> "00", ot04, "") & "." & CP10 & ".') "
                End Select
                If strSql <> "" Then
                    strSql = "Update CasePaperPdf set CPP01=" & CNULL(CP09) & strSql & " Where CPP01=" & CNULL("LOS" & m_LOS15)
                    cnnConnection.Execute strSql, intI
                End If
             End If
             'end 2020/06/10
             
             '若為新案時法務基本檔的案件屬性LC47依案源系統別(LOS01)存專利或商標或著作權
             If intCaseKind = 法務 And intSaveMode = 1 And m_LOS01 <> "" Then
                  'Modified by Lydia 2021/09/10 抓接洽單主檔的案件屬性
                  'strSql = "select cp01,sk04 from caseprogress,systemkind where cp09=" & CNULL(m_LOS01) & " and cp01=sk01(+) "
                  strSql = "select cp01,sk04,crl84 from caseprogress,systemkind, Consultrecordlist " & _
                              "where cp09=" & CNULL(m_LOS01) & " and cp01=sk01(+) and cp140=crl01(+) "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                      strSql = ""
                      If InStr("" & RsTemp.Fields("cp01"), "P") > 0 Then
                          strSql = " LC47='專利' "
                      'Modified by Lydia 2020/07/08
                      'ElseIf InStr("TC,CFC", "" & RsTemp.Fields("cp01")) > 0 Then
                      ElseIf "" & RsTemp.Fields("CP01") = "TC" Or "" & RsTemp.Fields("CP01") = "CFC" Then
                          strSql = " LC47='著作權' "
                      ElseIf InStr("" & RsTemp.Fields("cp01"), "T") > 0 And "" & RsTemp.Fields("cp01") <> "TT" Then
                          strSql = " LC47='商標' "
                      'Added by Lydia 2021/09/10 案源為TT時，預設法務基本檔的案件屬性LC47為TT接洽單之CRL84法務案件屬性。ex.L-006435剛收文的案件屬性為空白
                      ElseIf "" & RsTemp.Fields("cp01") = "TT" Then
                          strSql = " LC47='" & ChgSQL("" & RsTemp.Fields("crl84")) & "' "
                      'end 2021/09/10
                      End If
                      If strSql <> "" Then
                            strSql = "Update Lawcase Set " & strSql & " Where lc01=" + CNULL(ot01) + " and lc02=" + CNULL(ot02) + " and lc03=" + CNULL(ot03) + " and lc04=" + CNULL(ot04)
                            cnnConnection.Execute strSql
                      End If
                  End If
                  'Added by Lydia 2020/07/14  楊世安：法律所接案後自動新增相關卷號資料(雙向案號)；
                  If Left(m_LOS02, 1) = "B" Or Left(m_LOS02, 1) = "C" Then 'Added by Lydia 2020/07/29 排除TT案
                      strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(ot01) & ", " & CNULL(ot02) & ", " & CNULL(ot03) & ", " & CNULL(ot04) & ", " & CNULL(m_LOS01cp01) & ", " & CNULL(m_LOS01cp02) & ", " & CNULL(m_LOS01cp03) & ", " & CNULL(m_LOS01cp04) & " ) "
                      cnnConnection.Execute strSql
                      strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_LOS01cp01) & ", " & CNULL(m_LOS01cp02) & ", " & CNULL(m_LOS01cp03) & ", " & CNULL(m_LOS01cp04) & ", " & CNULL(ot01) & ", " & CNULL(ot02) & ", " & CNULL(ot03) & ", " & CNULL(ot04) & " ) "
                      cnnConnection.Execute strSql
                  End If 'Added by Lydia 2020/07/29 排除TT案
                  'end 2020/07/14
             End If
             '收文新案且不同審級的C類案源時，檢查相同LC02的案號若已有B類案源時Email通知秀玲要調整該法務案的審級順序。
             If intCaseKind = 法務 And intSaveMode = 1 And Val(ot03) >= 1 And Left(m_LOS02, 1) = "C" Then
                strSql = "select los06 from LawOfficeSource where los06 in (select cp09 from caseprogress where cp01='" & ot01 & "' and cp02='" & ot02 & "' and cp03<='" & Val(ot03) - 1 & "' and cp04='" & ot04 & "' and cp159=0) and los02 like 'B%' "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI = 1 Then
                    strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                             " values ('" & strUserNum & "','83002'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss')" & _
                             ",'收文" & ot01 & "-" & ot02 & "-" & ot03 & "-" & ot04 & "有不同審級的B類案源，請檢查該法務案的審級順序！','同摘要')"
                    cnnConnection.Execute strSql, intI
                End If
             End If
             
             '若案源資料的介紹客戶LOS05為空時表示新客戶要回寫並更新(收文時輸入的)客戶智權人員(CU12CU13)為介紹人(LOS04第一人)
             If m_LOS05 = "" And Trim(txtOther(7) & txtOther(18) & txtOther(19) & txtOther(29) & txtOther(30)) <> "" Then
                 '並且回寫案源介紹客戶編號LOS05
                 strSql = "update LawOfficeSource set los05='" & ChangeCustomerL(txtOther(7)) & "' where los05 is null and los15=" & CNULL(m_LOS15)
                 cnnConnection.Execute strSql, intI
                 If intI > 0 Then
                    strExc(1) = "7": strExc(2) = "18": strExc(3) = "19": strExc(4) = "29": strExc(5) = "30"
                    For intI = 1 To 5
                        If Trim(txtOther(Val(strExc(intI)))) <> "" Then
                             strSql = "update customer set cu12='" & m_LOS04_1st15 & "',cu13='" & m_LOS04_1 & "' where cu01='" & Left(ChangeCustomerL(txtOther(Val(strExc(intI)))), 8) & "' and cu02='" & Right(ChangeCustomerL(txtOther(Val(strExc(intI)))), 1) & "'"
                             Pub_SeekTbLog strSql
                             cnnConnection.Execute strSql
                        End If
                    Next intI
                    m_Los05_N = ChangeCustomerL(txtOther(7))    'Added by Lydia 2022/10/19 客戶編號後建=m_LOS05=空白
                 End If
             End If
             '最後才做-->客戶編號回寫後，案源案件類型A，若無點數則保留類型A，若有點數則判斷同一客戶編號介紹日前若有A1則此筆設為A2，若無則設為A1。
                                 '計算案源之費用及點數，更新回案源總收文號LOS01之費用及點數，以利智慧所開立收據。
                                 '案源為TT-999999時同時上發文日CP27為系統日(為無發文日者才更新)。
                                 '5/6跟楊監察人確認國外部介紹案源以相同分潤方式計算，不管國外代理人仍以客戶為介紹基準。
             'Modified by Lydia 2020/06/23 性質屬於B1歸屬於A類(A4)
             'If m_LOS02 = "A" And Val(txtOther(13)) > 0 Then
             If (m_LOS02 = "A" Or m_LOS02 = "A4") And Val(txtOther(13)) > 0 Then
                If m_LOS02 = "A" Then
             'end 2020/06/23
                    
                    'Modified by Morgan 2020/9/24 全部以第一次分潤A1，例外才算第二次A2，是指後續由法律所直接收文(不是介紹人再走案源流程)時則以A2計算
                    'strSql = "select los02 from LawOfficeSource where los12<'" & m_LOS12 & "' and los02='A1' and los05='" & IIf(m_LOS05 <> "", m_LOS05, ChangeCustomerL(txtOther(7))) & "' "
                    'intI = 1
                    'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                    'If intI = 1 Then
                    '    strSql = "update LawOfficeSource set los02='A2' where los15='" & m_LOS15 & "' "
                    '    cnnConnection.Execute strSql
                    'Else
                        strSql = "update LawOfficeSource set los02='A1' where los15='" & m_LOS15 & "' "
                        cnnConnection.Execute strSql
                    'End If
                    'end 2020/9/24
                    
                End If 'Added by Lydia 2020/06/23
                
                '案源為TT-999999時同時上發文日CP27為系統日(為無發文日者才更新)。
                If m_LOS01cp01 & m_LOS01cp02 = "TT999999" Then
                    strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_LOS01 & "' and nvl(cp27,0)=0 "
                    cnnConnection.Execute strSql
                End If
             End If
             
             '計算案源之費用及點數，更新回案源總收文號LOS01之費用及點數，以利智慧所開立收據。
             PUB_UpdateTTFee m_LOS15 'Added by Morgan 2020/9/29 同案源單號的每個收文性質都要(費用加總)
         End If
         'end 2020/05/20
         
         'Added by Lydia 2020/05/20 法律所案源收文：若該收文號點數>0但無案源(自行收文者)時，若案件的客戶為非法律所的客戶時則仍算A類案源(另寫函數參照作帳規則設定為A1~A4)。
                                            '系統自動新增TT-999999案進度(B類收文)及法律所案源資料(同最後一筆案源的資料)。
         'Modified by Lydia 2020/10/05 +舊案txtCode(0) <> ""
         'Modified by Lydia 2021/01/08 拿掉台灣案的限制 And txtOther(3) = "000"
         If strSrvDate(1) >= 法律所案源收文啟用日 And InStr(txtSystem, "L") > 0 And m_LOS15 = "" And Val(txtOther(13)) > 0 And txtCode(0) <> "" Then
             'Modified by Lydia 2020/10/05 + st01
             strSql = "select cu01,cu02,st15,st01 from customer,staff where cu01='" & Mid(ChangeCustomerL(txtOther(7)), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(txtOther(7)), 9, 1) & "'  and cu13=st01(+) "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strSql)
             If intI = 1 Then
                 strExc(1) = Left("" & RsTemp.Fields("st15"), 1)
                 If strExc(1) <> "L" Then
                    '非法律所的舊客戶時則仍算A類案源(另寫函數參照作帳規則設定為A1~A4)
                    'Modified by Lydia 2020/10/05 (9/30) 如為舊案並且曾有A1類案源時，則為A2類案源
                    'strSql = "select * from LawOfficeSource where los12<'" & strSrvDate(1) & "' and los02 like 'A%' and los05='" & ChangeCustomerL(txtOther(7)) & "' " & _
                                "order by los12 desc, los13 desc "
                    'Modified by Lydia 2021/01/06 以最後一道案源為準
                    'strSql = "Select * From Lawofficesource Where los02='A1' and los07||los08 is null and Los15 In " & _
                                 "(select max(cp162) from caseprogress where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp162 is not null) "
                    'Modified by Lydia 2022/11/09 增加LA案的A3案源
                    'strSql = "Select * From Lawofficesource Where los02='A1' and los07||los08 is null and Los15 In " & _
                                 "(select cp162 from caseprogress where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp162 is not null) order by los12 desc "
                    strSql = "Select * From Lawofficesource Where los02 in ('A1','A3') and los07||los08 is null and Los15 In " & _
                                 "(select cp162 from caseprogress where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp162 is not null) order by los12 desc "
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                    If intI = 1 Then
                       RsTemp.MoveFirst
                       '案源類別
                       'Modified by Lydia 2020/10/05 固定為A2類案源
                        'If "" & RsTemp.Fields("los02") = "A" Then
                        '    strExc(2) = "A1"
                        'Else
                        '    strExc(2) = "A2"
                        'End If
                        'Added by Lydia 2022/11/09 增加LA案的A3案源
                        If txtSystem = "LA" Then
                           strExc(2) = "A3"
                        Else
                        'end 2022/11/09
                           strExc(2) = "A2"
                           'end 2020/10/05
                        End If 'Added by Lydia 2022/11/09
                        'TT新增B類收文
                        'Modified by Lydia 2020/10/05
                        'strExc(3) = txtOther(10)
                        strExc(3) = ""
                        If "" & RsTemp.Fields("LOS04") <> "" Then '抓介紹人1
                           'Modified by Lydia 2020/10/05 用於案源之接洽人取得在職員工編號和介紹人第一人
                           'If InStr("" & RsTemp.Fields("LOS04"), ",") = 0 Then
                           '    strExc(3) = "" & RsTemp.Fields("LOS04")
                           'Else
                           '    strExc(3) = Mid(RsTemp.Fields("LOS04"), 1, InStr("" & RsTemp.Fields("LOS04"), ",") - 1)
                           'End If
                           strExc(5) = PUB_GetNowStaff("" & RsTemp.Fields("los04"), strExc(3))
                           'end 2020/10/05
                        End If
                        'Added by Lydia 2020/10/05
                        m_Los04_N1 = strExc(3)
                        If strExc(3) <> "" Then
                            strExc(1) = AutoNo("B", 6) 'TT收文號 'Move  by Lydia 2020/10/07 從RsTemp.MoveFirst下方移過來
                            'end 2020/10/05
                            'Modified by Morgan 2021/1/8 +CP20,CP27,CP32
                            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp11,cp12,cp13,cp20,cp27,cp32,CP162)" & _
                               " values('TT','999999','0','00'," & strSrvDate(1) & "," & CNULL(TransDate(txtOther(9), 2), True) & ",'" & strExc(1) & "'" & _
                               ",'735','07','" & GetST15(strExc(3)) & "','" & strExc(3) & "','N'," & strSrvDate(1) & ",'N',null)"
                            cnnConnection.Execute strSql
                            '法律所案源資料(同最後一筆案源的資料), 案源單號=TT總收文號
                            strExc(4) = AutoNo("LOS", 5, , True)
                            'Modified by Lydia 2020/10/05
                            'strSql = "insert into LawOfficeSource(LOS01,LOS02,LOS03,LOS04,LOS05,LOS06,LOS10,LOS11,LOS12,LOS13,LOS15)" & _
                               " values ('" & strExc(1) & "','" & strExc(2) & "' ,'" & RsTemp.Fields("los03") & "'" & _
                               ",'" & RsTemp.Fields("los04") & "','" & RsTemp.Fields("los05") & "','" & CP09 & "','" & strExc(1) & "'" & _
                               ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss'),'" & strExc(4) & "')"
                            strSql = "insert into LawOfficeSource(LOS01,LOS02,LOS03,LOS04,LOS05,LOS06,LOS10,LOS11,LOS12,LOS13,LOS15)" & _
                               " values ('" & strExc(1) & "','" & strExc(2) & "' ,'" & txtOther(10) & "'" & _
                               ",'" & strExc(5) & "','" & ChangeCustomerL(txtOther(7)) & "','" & CP09 & "','" & strExc(1) & "'" & _
                               ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss'),'" & strExc(4) & "')"
                            'Modified by Lydia 2022/09/12 debug
                            'm_Los05_N = strExc(4)
                            m_Los05_N = ChangeCustomerL(txtOther(7))
                            'end 2020/10/05
                            cnnConnection.Execute strSql
                            'Added by Lydia 2020/10/05 收文之進度加註案源
                            'Modified by Lydia 2021/01/08 補上案源單號CP162
                            strSql = "Update CaseProgress Set cp64=" & CNULL("案源：TT-999999(" & strExc(1) & ");") & "||cp64,CP162=" & CNULL(strExc(4)) & " where cp09=" & CNULL(CP09)
                            cnnConnection.Execute strSql
                        
                            '計算案源之費用及點數，更新回案源總收文號LOS01之費用及點數，以利智慧所開立收據。
                            PUB_UpdateTTFee strExc(4) 'Added by Morgan 2020/9/29
                        End If 'Added by Lydia 2020/10/05
                    End If
                 End If
             End If
         End If
         'end 2020/05/20
         
         '92.3.8 ADD BY SONIA FCL之請款收文直接發文日
         '2008/8/19 cancel by sonia 蘇月星說自行上發文日
         'If ot01 = "FCL" And CP10 = "78" Then
         '   strSQL = "update caseprogress set cp27=cp05 where cp09=" + CNULL(CP09)
         '   cnnConnection.Execute strSQL
         'End If
         '2008/8/19 end
         '92.3.8 END
           
           'add by nickc 2008/01/04 加入回代時，承辦期限為本所收文日(當天不算)之第二個工作天
           If CP10 = "720" Then
               'Modify by Morgan 2008/9/23 FG 改在上面設定
               If Me.txtSystem.Text <> "FG" Then
                  strSql = "update caseprogress set cp48=" + CNULL(CompWorkDay(3, DBDATE(cp05), 0)) + " where cp09=" + CNULL(CP09)
                  cnnConnection.Execute strSql
               End If
           End If
   
           '若為接洽記錄單(櫃台收文)
           'Modify by Morgan 2007/10/26 費用可改時才做，否則已收款資料會被還原
           'If frm010001.intChoose = 0 Then
           If frm010001.intChoose = 0 And txtOther(12).Enabled = True Then
           'end 2007/10/26
               '未收金額 = 費用
               strSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(CP09)
               cnnConnection.Execute strSql
           End If
                    
         'Add By Cheng 2002/05/10
         '若為內部收文作業時, 案件進度檔的是否向客戶收款設定為"N"
         If frm010001.intChoose = 1 Then
            strSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(CP09)
            cnnConnection.Execute strSql
         End If
   
         strSql = "update customer set cu30=" + CNULL(cu30) + " where cu01=" + CNULL(Mid(ot09, 1, 8)) + " and cu02=" + CNULL(Mid(ot09, 9, 1))
         cnnConnection.Execute strSql
         
           'Add by Lydia 2014/10/31 開放外專程序人員可進入專利處系統操作FMP寰華案件
           If mFMPchk = True Then
               strSql = "update caseprogress set cp44='Y53374000' where cp09='" & CP09 & "' "
               cnnConnection.Execute strSql
           End If
           'end. 'Add by Lydia 2014/10/31
                 
         '92.2.19 MODIFY BY SONIA先抓NP06=NULL, 沒有資料才抓NU06<>'Y', 只有一筆才更新
         'If bolRt Then
         '  'Move By Cheng 2002/12/18
         '  adoquery.CursorLocation = adUseClient
         '  adoquery.Open "select np01 from nextprogress where np02 = '" & ot01 & "' and np03 = '" & ot02 & "' and np04 = '" & ot03 & "' and np05 = '" & ot04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
         '  'Modify By Cheng 2002/05/10
         '  '若在下一程序檔只抓到一筆資料時, 才要抓下一程序檔的總收文號更新案件進度檔的相關總收文號
         '  'If adoquery.RecordCount <> 0 Then
         '  If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
         '     If IsNull(adoquery.Fields(0).Value) = False Then
         '        cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & cp09 & "'"
         '     End If
         '  End If
         '  adoquery.Close
         '
         '   strSQL = "update nextprogress set np06='Y' where np02=" + CNULL(ot01) + " and np03=" + _
         '       CNULL(ot02) + " and np04=" + CNULL(ot03) + " and np05=" + CNULL(ot04) + _
         '       " and np07=" + CNULL(CP10) + " and (np06<>'Y' or np06 is null)"
         '   cnnConnection.Execute strSQL
         'End If
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select np01 from nextprogress where np02 = '" & ot01 & "' and np03 = '" & ot02 & "' and np04 = '" & ot03 & "' and np05 = '" & ot04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount > 0 Then
            If adoquery.RecordCount = 1 Then
               If IsNull(adoquery.Fields(0).Value) = False Then
                  '2011/6/17 add by sonia 異議答辯、評定答辯、廢止答辯要一並更新對造資料
                  If (CP10 = "602" Or CP10 = "604" Or CP10 = "606") Then
                     cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & CP09 & "'", intI
                  Else
                  '2011/6/17 end
                     cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & CP09 & "'"
                  End If  '2011/6/17 add by sonia
               End If
               'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
               strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" + CNULL(ot01) + " and np03=" + _
                  CNULL(ot02) + " and np04=" + CNULL(ot03) + " and np05=" + CNULL(ot04) + _
                  " and np07=" + CNULL(CP10) + " and np06 is null"
               cnnConnection.Execute strSql
            End If
         Else
            adoquery.Close
            adoquery.CursorLocation = adUseClient
            adoquery.Open "select np01 from nextprogress where np02 = '" & ot01 & "' and np03 = '" & ot02 & "' and np04 = '" & ot03 & "' and np05 = '" & ot04 & "' and np06 <>'Y' and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount > 0 Then
               If adoquery.RecordCount = 1 Then
                  If IsNull(adoquery.Fields(0).Value) = False Then
                     '2011/6/17 add by sonia 異議答辯、評定答辯、廢止答辯要一並更新對造資料
                     If (CP10 = "602" Or CP10 = "604" Or CP10 = "606") Then
                        cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & CP09 & "'", intI
                     Else
                     '2011/6/17 end
                        cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & CP09 & "'"
                     End If  '2011/6/17 add by sonia
                  End If
                  'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
                  strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" + CNULL(ot01) + " and np03=" + _
                     CNULL(ot02) + " and np04=" + CNULL(ot03) + " and np05=" + CNULL(ot04) + _
                     " and np07=" + CNULL(CP10) + " and np06 <> 'Y'"
                  cnnConnection.Execute strSql
               End If
            End If
         End If
         adoquery.Close

         '92.2.19 END
         If ot05 = "" Then ot05 = 台灣國家代號
         'edit by nickc 2007/02/06 不用 dll 了
         'If obj001.SetCaseProgressFee(ot01, ot05, CP10, CP09) = False Then bolError = True
         If Cls001SetCaseProgressFee(ot01, ot05, CP10, CP09) = False Then bolError = True
      Else
         bolError = True
      End If
   End If
   'Modify By Cheng 2002/12/18
   'adoquery.CursorLocation = adUseClient
   ''adoquery.Open "select np01 from nextprogress where np02 = '" & ot01 & "' and np03 = '" & ot02 & "' and np04 = '" & ot03 & "' and np05 = '" & ot04 & "' and np07 = '" & cp10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
   'adoquery.Open "select np01 from nextprogress where np02 = '" & ot01 & "' and np03 = '" & ot02 & "' and np04 = '" & ot03 & "' and np05 = '" & ot04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
   ''Modify By Cheng 2002/05/10
   ''若在下一程序檔只抓到一筆資料時, 才要抓下一程序檔的總收文號更新案件進度檔的相關總收文號
   ''If adoquery.RecordCount <> 0 Then
   'If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
   '   If IsNull(adoquery.Fields(0).Value) = False Then
   '      cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & cp09 & "'"
   '   End If
   'End If
   'adoquery.Close
   
   'add by nickc 2008/05/02 儲存預定收款日
   'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'   If bolError = False Then
'       Dim rtCnt As Integer
'       'Modify by Morgan 2010/12/9
'       'If txtOther(28) <> "" Then
'       '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtOther(28)) & " ", rtCnt
'       If txtOther(28) <> "" And txtOther(28) <> txtOther(28).Tag Then
'           cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
'       'end 2010/12/9
'           If rtCnt = 0 Then
'               cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from dual "
'           End If
'       End If
'   End If
   'end 2018/08/22
   
   'Add by Sindy 2022/8/17
   If m_strIR01 <> "" Then
      m_bolRecvOK = True
      m_strMCR11 = ""
      If m_bMRecvBatch = True Then '多案收文
         '更新總收文號
         strSql = "update multiCaseRecv set mcr11='" & CP09 & "'" & _
                  " where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                  " and mcr02='" & ot01 & "' and mcr03='" & ot02 & "' and mcr04='" & ot03 & "' and mcr05='" & ot04 & "'" & _
                  " and mcr06='" & CP10 & "'"
                  cnnConnection.Execute strSql
                  
         'Modify By Sindy 2022/8/26
         '下載信件檔,上傳卷宗區
         Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, CP09)
         
         '檢查多案收文狀況
         strExc(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                     " and mcr02||mcr03||mcr04||mcr05<>'" & ot01 & ot02 & ot03 & ot04 & "'" & _
                     " and mcr11 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_bolRecvOK = False '尚有未收文
            
            'Modify By Sindy 2022/8/26 此處Mark,程式往上移
'            '下載信件檔,上傳卷宗區
'            Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, CP09)
         Else
            m_bolRecvOK = True '全部收完文
            '抓第一筆的總收文號
            strExc(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                        " and mcr02||mcr03||mcr04||mcr05=mcr07||mcr08||mcr09||mcr10 and mcr11 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_strMCR11 = RsTemp.Fields("mcr11")
            Else
               MsgBox "多案收文，無讀取到第一筆案件的總收文號，請洽電腦中心!!", vbExclamation '此狀況應不會發生, 以防外一
               GoTo ErrHand
            End If
         End If
      End If
      If m_bolRecvOK = True Then '全部收完文
         '多案收文的總收文號要傳入第一筆總收文號
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, _
               IIf(m_strMCR11 <> "", "多案收文", "frm010001"), _
               IIf(Left(Pub_StrUserSt03, 2) = "F2", IIf(m_strMCR11 <> "", m_strMCR11, CP09), "")
      End If
   End If
   '2022/8/17 END
   
   If bolError Then
      cnnConnection.RollbackTrans
      ShowMsg MsgText(9004)
      'add by nickc 2007/12/12
   IsSaveData = False
   Else
      'Modified by Lydia 2022/09/08 改成共用模組
      'Call SaveFrame21(CP09) 'Added by Lydia 2022/07/15 TC案之文件齊備日管控
      Call GetStrControl
      'Modified by Lydia 2022/09/29 傳入系統別,國家,案件性質 => ot01, txtOther(3), CP10
      Call PUB_SaveByControl(CP09, m_strControl, ot01, txtOther(3), CP10)
      'end 2022/09/08
      
      cnnConnection.CommitTrans
      InsertOtherDatabase = True
      'add by nickc 2006/03/27
      txtCode(0) = ot02
   End If
   'add by nickc 2005/08/12
   txtCode(0) = ot02
   Exit Function
ErrHand:
   cnnConnection.RollbackTrans
   'edit by nickc 2006/03/07 解決 cp02=null 的問題
   'add by nickc 2005/08/25
   ' txtCode(0) = ""
   ShowMsg MsgText(9004)
   'add by nickc 2007/12/12
   IsSaveData = False
End Function


'修改案件進度檔資料庫
Private Function UpdateOtherDatabase(ByRef intCaseKind As Integer, ByRef ot01 As String, _
             ByRef ot02 As String, ByRef ot03 As String, ByRef ot04 As String, ByRef ot05 As String, _
             ByRef ot06 As String, ByRef ot07 As String, ByRef ot08 As String, ByRef ot09 As String, ByRef ot10 As String, _
             ByRef ot11 As String, ByRef ot12 As String, ByRef ot18 As String, ByRef CP09 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, ByRef cp11 As String, _
             ByRef cp13 As String, ByRef cp16 As String, ByRef cp17 As String, ByRef cp18 As String, ByRef cp19 As String, _
             ByRef cp32 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef cp33 As Double, ByRef cp34 As Double, _
             ByRef CP64 As String, ByRef SP73 As String, ByRef SP74 As String, ByRef SP27 As String, ByRef SP78 As String, _
             ByRef ot13 As String, ByRef ot14 As String) As Boolean

Dim strSql As String
Dim adoquery As New ADODB.Recordset
Dim cp48 As String, stUpdate As String 'Add by Morgan 2008/8/23

'add by nickc 2007/12/12
If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand
cnnConnection.BeginTrans
cp05 = ChangeTStringToWString(cp05)
cp06 = ChangeTStringToWString(cp06)
cp07 = ChangeTStringToWString(cp07)
ot09 = ChangeCustomerL(ot09) '當事人1
ot10 = ChangeCustomerL(ot10) '當事人2
ot11 = ChangeCustomerL(ot11) '當事人3
ot12 = ChangeCustomerL(ot12)
'Add By Sindy 2011/1/18
ot13 = ChangeCustomerL(ot13) '當事人4
ot14 = ChangeCustomerL(ot14) '當事人5
'2011/1/18 End

'Add by Lydia 2014/10/31 開放外專程序人員可進入專利處系統操作FMP寰華案件=>寫入代理人
If mFMPchk = True Then
    strSql = "update caseprogress set cp44='Y53374000' where cp09='" & txtRecieveCode & "' "
Else
    strSql = "update caseprogress set cp44='' where cp09='" & txtRecieveCode & "' "
End If
cnnConnection.Execute strSql
'end. 'Add by Lydia 2014/10/31
        
Select Case intCaseKind
             Case 法務
                        'edit by nickc 2007/03/27 加入彼所案號
                        'strSQL = "update lawcase set lc05=" + CNULL(ChgSQL(ot06)) + ",lc06=" + CNULL(ChgSQL(ot07)) + _
                            ",lc07=" + CNULL(ChgSQL(ot08)) + ",lc11=" + CNULL(ot09) + ",lc15=" + CNULL(ot05) + _
                            ",lc22=" + CNULL(ot12) + " where lc01=" + CNULL(ot01) + " and lc02=" + _
                            CNULL(ot02) + " and lc03=" + CNULL(ot03) + " and lc04=" + CNULL(ot04)
                        'Modify By Sindy 2011/1/18 +lc43,lc44,lc45,lc46
                        strSql = "update lawcase set lc05=" + CNULL(ChgSQL(ot06)) + ",lc06=" + CNULL(ChgSQL(ot07)) + _
                            ",lc07=" + CNULL(ChgSQL(ot08)) + ",lc11=" + CNULL(ot09) + ",lc15=" + CNULL(ot05) + _
                            ",lc22=" + CNULL(ot12) + ",lc23=" + CNULL(SP27) + ",lc43=" + CNULL(ot10) + ",lc44=" + CNULL(ot11) + ",lc45=" + CNULL(ot13) + ",lc46=" + CNULL(ot14)
                        'Add by Morgan 2008/8/5 +LC42
                        If UCase(SP78) <> "SP78" Then
                           strSql = strSql + ",LC42=" + CNULL(SP78)
                        End If
                        strSql = strSql + " where lc01=" + CNULL(ot01) + " and lc02=" + _
                            CNULL(ot02) + " and lc03=" + CNULL(ot03) + " and lc04=" + CNULL(ot04)
                        cnnConnection.Execute strSql
                        strSql = "Update Lawcase Set LC16='" & ChgSQL(Me.txtOther(23).Text) & "', LC17='" & ChgSQL(Me.txtOther(22).Text) & "' Where lc01=" + CNULL(ot01) + " and lc02=" + CNULL(ot02) + " and lc03=" + CNULL(ot03) + " and lc04=" + CNULL(ot04)
                        cnnConnection.Execute strSql
             Case 顧問
                        'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
                        strSql = "update hirecase set hc05=" + CNULL(ot09) + ",hc06=" + CNULL(ChgSQL(ot06)) + ",hc24=" + CNULL(ot10) + ",hc25=" + CNULL(ot11) + ",hc26=" + CNULL(ot13) + ",hc27=" + CNULL(ot14)
                        'Add by Morgan 2008/8/5 +HC23
                        If UCase(SP78) <> "SP78" Then
                           strSql = strSql + ",HC23=" + CNULL(SP78)
                        End If
                        strSql = strSql & " where hc01=" + CNULL(ot01) + " and hc02=" + CNULL(ot02) + " and hc03=" + CNULL(ot03) + " and hc04=" + CNULL(ot04)
                        cnnConnection.Execute strSql
                        strSql = "Update Hirecase Set HC07='" & ChgSQL(Me.txtOther(23).Text) & "' Where HC01=" + CNULL(ot01) + " and HC02=" + CNULL(ot02) + " and HC03=" + CNULL(ot03) + " and HC04=" + CNULL(ot04)
                        cnnConnection.Execute strSql
             Case Else
                        'Modify By Sindy 2011/1/18 +sp65,sp66
                        strSql = "update servicepractice set sp05=" + CNULL(ChgSQL(ot06)) + ",sp06=" + CNULL(ChgSQL(ot07)) + _
                           ",sp07=" + CNULL(ChgSQL(ot08)) + ",sp08=" + CNULL(ot09) + ",sp58=" + CNULL(ot10) + _
                           ",sp59=" + CNULL(ot11) + ",sp65=" + CNULL(ot13) + ",sp66=" + CNULL(ot14) + ",sp09=" + CNULL(ot05) + ",sp26=" + CNULL(ot12) + ",sp18=" + CNULL(ChgSQL(ot18)) + ",sp73=" + CNULL(SP73) + ",sp74=" + CNULL(SP74) + ",sp27=" + CNULL(SP27)
                        'Add by Morgan 2008/8/5 +HC23
                        If UCase(SP78) <> "SP78" Then
                           strSql = strSql + ",SP78=" + CNULL(SP78)
                        End If
                        'Add By Sindy 2010/3/8
                        If bolCancel = True Then
                           strSql = strSql + ",SP30=" + CNULL(ChgSQL(strSP30s)) + ",SP75=" + CNULL(ChgSQL(strSP75s))
                        End If
                        '2010/3/8 End
                        strSql = strSql & " where sp01=" + CNULL(ot01) + " and sp02=" + CNULL(ot02) + " and sp03=" + CNULL(ot03) + " and sp04=" + CNULL(ot04)
                        cnnConnection.Execute strSql
                        strSql = "Update Servicepractice Set SP28='" & ChgSQL(Me.txtOther(23).Text) & "', SP29='" & ChgSQL(Me.txtOther(22).Text) & "' Where SP01=" + CNULL(ot01) + " and SP02=" + CNULL(ot02) + " and SP03=" + CNULL(ot03) + " and SP04=" + CNULL(ot04)
                        cnnConnection.Execute strSql
End Select

'Add by Morgan 2008/9/23
If Me.txtSystem.Text = "FG" Then
   cp48 = Pub_GetHandleDay("FG", "000", CP10, , cp06)
   stUpdate = ",cp48=" & CNULL(cp48, True)
End If

'Add By Sindy 2012/11/06 有★★的應收帳款簽核控管
m_CP150 = ""
If Check2.Value = 1 Then m_CP150 = "Y"
'2012/11/06 End

'Modify By Sindy 2012/11/06 +CP150
strSql = "update caseprogress set cp05=" + CNULL(cp05) + ",cp06=" + CNULL(cp06) + ",cp07=" + CNULL(cp07) + ",cp10=" + CNULL(CP10) + _
         ",cp11=" + CNULL(cp11) + ",cp13=" + CNULL(cp13) + ",cp14=" + CNULL(cp14) + ",cp16=" + CNULL(cp16) + ",cp17=" + CNULL(cp17) + _
         ",cp18=" + CNULL(cp18) + ",cp19=" + CNULL(cp19) + ",cp32=" + CNULL(cp32) + ",cp33=" & cp33 & ",cp34=" & cp34 & ",CP64=" + CNULL(ChgSQL(CP64)) & stUpdate & ",cp150=" & CNULL(m_CP150) & " where cp09='" + CP09 + "'"
cnnConnection.Execute strSql
strSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(cp13) + ") where cp09=" + CNULL(CP09)
cnnConnection.Execute strSql
'add by nickc 2008/01/04 加入回代時，承辦期限為本所收文日(當天不算)之第二個工作天
If CP10 = "720" Then
     strSql = "update caseprogress set cp48=" + CNULL(CompWorkDay(3, DBDATE(cp05), 0)) + " where cp09=" + CNULL(CP09)
     cnnConnection.Execute strSql
End If
        'Add By nickc 2007/08/21
        '若為接洽記錄單(櫃台收文)
        'Modify by Morgan 2007/10/26 費用可改時才做，否則已收款資料會被還原
        'If frm010001.intChoose = 0 Then
        If frm010001.intChoose = 0 And txtOther(12).Enabled = True Then
        'end 2007/10/26
            '未收金額 = 費用
            strSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(CP09)
            cnnConnection.Execute strSql
        End If
        
'Added by Lydia 2022/11/29 非內部收文並且有費用，先統一設定CP20=Null ;
If frm010001.intChoose = 0 And Val(cp16) > 0 Then
    stUpdate = ""
    If Me.txtSystem.Text = "FG" Then
       stUpdate = PUB_GetCP20(txtSystem, CP10)
    End If
    If stUpdate = "" Then
       strSql = "update caseprogress set cp20=null where cp09=" + CNULL(CP09)
       cnnConnection.Execute strSql
    End If
End If
'end 2022/11/29

'Add By Cheng 2002/05/10
'若為內部收文作業時, 案件進度檔的是否向客戶收款設定為"N"
If frm010001.intChoose = 1 Then
   strSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(CP09)
   cnnConnection.Execute strSql
End If

strSql = "update customer set cu30=" + CNULL(cu30) + " where cu01=" + CNULL(Mid(ot09, 1, 8)) + " and cu02=" + CNULL(Mid(ot09, 9, 1))
cnnConnection.Execute strSql
adoquery.CursorLocation = adUseClient
'adoquery.Open "select np01 from nextprogress where np02 = '" & ot01 & "' and np03 = '" & ot02 & "' and np04 = '" & ot03 & "' and np05 = '" & ot04 & "' and np07 = '" & cp10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
adoquery.Open "select np01 from nextprogress where np02 = '" & ot01 & "' and np03 = '" & ot02 & "' and np04 = '" & ot03 & "' and np05 = '" & ot04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
'Modify By Cheng 2002/05/10
'若在下一程序檔只抓到一筆資料時, 才要抓下一程序檔的總收文號更新案件進度檔的相關總收文號
'If adoquery.RecordCount <> 0 Then
If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
   If IsNull(adoquery.Fields(0).Value) = False Then
      '2011/6/17 add by sonia 異議答辯、評定答辯、廢止答辯要一並更新對造資料
      If (CP10 = "602" Or CP10 = "604" Or CP10 = "606") Then
         cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & CP09 & "'", intI
      Else
      '2011/6/17 end
         cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & CP09 & "'"
      End If  '2011/6/17 add by sonia
   End If
End If
adoquery.Close
'add by nickc 2008/05/02 儲存預定收款日
'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'Dim rtCnt As Integer
''Modify by Morgan 2010/12/9
''If txtOther(28) <> "" Then
''    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtOther(28)) & " ", rtCnt
'If txtOther(28) <> "" And txtOther(28) <> txtOther(28).Tag Then
'    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
''end 2010/12/9
'    If rtCnt = 0 Then
'        cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from dual "
'    End If
'End If
'end 2018/08/22

If m_LOS15 <> "" Then PUB_UpdateTTFee m_LOS15 'Added by Morgan 2022/4/14

'Modified by Lydia 2022/09/08 改成共用模組
'Call SaveFrame21(CP09) 'Added by Lydia 2022/07/15 TC案之文件齊備日管控
Call GetStrControl
'Modified by Lydia 2022/09/29 傳入系統別,國家,案件性質 => ot01, txtOther(3), CP10
Call PUB_SaveByControl(CP09, m_strControl, ot01, txtOther(3), CP10)
'end 2022/09/08
      
cnnConnection.CommitTrans
UpdateOtherDatabase = True
Exit Function
ErrHand:
cnnConnection.RollbackTrans
ShowMsg MsgText(9004)
'add by nickc 2007/12/12
IsSaveData = False
End Function

'讀取Other資料庫
Private Function ReadOtherDatabase(ByRef intModifyKind As Integer, ByRef intCaseKind As Integer, ByRef ot01 As String, _
             ByRef ot02 As String, ByRef ot03 As String, ByRef ot04 As String, ByRef ot05 As String, _
             ByRef ot06 As String, ByRef ot07 As String, ByRef ot08 As String, ByRef ot09 As String, ByRef ot10 As String, _
             ByRef ot11 As String, ByRef ot12 As String, ByRef ot18 As String, ByRef CP09 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, ByRef cp11 As String, _
             ByRef cp13 As String, ByRef cp16 As String, ByRef cp17 As String, ByRef cp18 As String, _
             ByRef cp19 As String, ByRef cp32 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef CP64 As String, _
             ByRef SP73 As String, ByRef SP74 As String, ByRef SP27 As String, ByRef ot13 As String, ByRef ot14 As String, _
             ByRef CP150 As String) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset, strTemp As String
   'Add by Morgan 2004/4/15
   '收據號碼
   Dim stCP60 As String
   
On Error GoTo ErrHand

bolisNP0809 = False 'Added by Lydia 2023/06/08

If intModifyKind <> 0 Then
   'Add by Morgan 2004/4/15
   '收據號碼
   'strSQL = "select cp05,cp06,cp07,cp10,cp11,cp13,cp16,cp17,cp18,cp19,cp32,cp14 from caseprogress where cp09='" + cp09 + "'"
   'Modify by Morgan 2005/12/13 加cp33,cp34
   'Modify by Sindy 2011/6/3 +cp64
   'Add by Lydia 2014/10/31 開放外專程序 =>讀cp31
   strSql = "select cp05,cp06,cp07,cp10,cp11,cp13,cp16,cp17,cp18,cp19,cp32,cp14,cp60,cp33,cp34,cp64,cp150,cp31 from caseprogress where cp09='" + CP09 + "'"
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      mCP31 = "" & rsRecordset.Fields("CP31") 'Add by Lydia 2014/10/31 開放外專程序 =>讀cp31
      cp05 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
      If cp05 <> "" Then cp05 = ChangeWStringToTString(cp05)
      cp06 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
      cp07 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
      CP10 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
      cp11 = IIf(IsNull(rsRecordset.Fields(4)), "", rsRecordset.Fields(4))
      cp13 = IIf(IsNull(rsRecordset.Fields(5)), "", rsRecordset.Fields(5))
      cp16 = IIf(IsNull(rsRecordset.Fields(6)), "", rsRecordset.Fields(6))
      cp17 = IIf(IsNull(rsRecordset.Fields(7)), "", rsRecordset.Fields(7))
      cp18 = IIf(IsNull(rsRecordset.Fields(8)), "", rsRecordset.Fields(8))
      cp19 = IIf(IsNull(rsRecordset.Fields(9)), "", rsRecordset.Fields(9))
      cp32 = IIf(IsNull(rsRecordset.Fields(10)), "", rsRecordset.Fields(10))
      cp14 = IIf(IsNull(rsRecordset.Fields(11)), "", rsRecordset.Fields(11))
      CP64 = IIf(IsNull(rsRecordset.Fields("cp64")), "", rsRecordset.Fields("cp64")) 'Add by Sindy 2011/6/3
      
      'Add by Morgan 2005/12/13
      douStPrice = Val("" & rsRecordset("CP33"))
      douLowPrice = Val("" & rsRecordset("CP34"))
      '2005/12/13 end
         
      'Add by Morgan 2004/4/15
      stCP60 = "" & rsRecordset.Fields("cp60")
      If stCP60 <> "" Then
         txtOther(12).Enabled = False: txtOther(13).Enabled = False: txtOther(16).Enabled = False
         'add by nickc 2006/12/25 加鎖智權人員
         txtOther(10).Enabled = False
      End If
      CP150 = "" & rsRecordset.Fields("cp150") 'Add By Sindy 2012/11/08
   Else
      ShowMsg MsgText(1502)
      rsRecordset.Close
      Exit Function
   End If
   rsRecordset.Close
Else
'NICK 900803 ********************** 邱小姐說要改成不管系統別都要讀出本所期限法定期限
   'If intCaseKind <> 顧問 Then
'      If GetNextProgressDate(ot01, ot02, ot03, ot04, cp10, cp06, cp07, CP64) = False Then
      If GetNextProgressDate(ot01, ot02, ot03, ot04, CP10, cp06, cp07, CP64, cp13) = False Then
         Exit Function
      End If
   'End If
' **********************
End If
If cp06 <> "" Then cp06 = ChangeWStringToTString(cp06)
If cp07 <> "" Then cp07 = ChangeWStringToTString(cp07)
Select Case intCaseKind
             Case 顧問
                        'Modify By Cheng 2003/06/13
'                        strSQL = "select '',hc06,'','',hc05,'','','' from hirecase where hc01=" + CNULL(ot01) + " and hc02=" + CNULL(ot02) + " and hc03=" + CNULL(ot03) + " and hc04=" + CNULL(ot04)
                        'Modify By Cheng 2003/08/28
'                        strSQL = "select '', hc06, '', '', hc05, '', '', '', '' from hirecase where hc01=" + CNULL(ot01) + " and hc02=" + CNULL(ot02) + " and hc03=" + CNULL(ot03) + " and hc04=" + CNULL(ot04)
                        'edit by nickc 2007/03/27 加入彼所案號
                        'strSQL = "select '', hc06, '', '', hc05, '', '', '', '', HC07, '' from hirecase where hc01=" + CNULL(ot01) + " and hc02=" + CNULL(ot02) + " and hc03=" + CNULL(ot03) + " and hc04=" + CNULL(ot04)
                        'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
                        strSql = "select '', hc06, '', '', hc05, '', hc24, hc25, '', HC07, '','' as sp27,HC23 as sp78,HC26 as sp65,HC27 as sp66 from hirecase where hc01=" + CNULL(ot01) + " and hc02=" + CNULL(ot02) + " and hc03=" + CNULL(ot03) + " and hc04=" + CNULL(ot04)
             Case 法務
                        'Modify By Cheng 2003/06/13
'                        strSQL = "select lc15,lc05,lc06,lc07,lc11,lc22,'','' from lawcase where lc01=" + CNULL(ot01) + " and lc02=" + CNULL(ot02) + " and lc03=" + CNULL(ot03) + " and lc04=" + CNULL(ot04)
                        'Modify By Cheng 2003/08/28
'                        strSQL = "select lc15, lc05, lc06, lc07, lc11, lc22, '', '', '' from lawcase where lc01=" + CNULL(ot01) + " and lc02=" + CNULL(ot02) + " and lc03=" + CNULL(ot03) + " and lc04=" + CNULL(ot04)
                        'edit by nickc 2007/03/27 加入彼所案號
                        'strSQL = "select lc15, lc05, lc06, lc07, lc11, lc22, '', '', '', LC16, LC17 from lawcase where lc01=" + CNULL(ot01) + " and lc02=" + CNULL(ot02) + " and lc03=" + CNULL(ot03) + " and lc04=" + CNULL(ot04)
                        'Modify By Sindy 2011/1/18 +lc43,lc44,lc45,lc46
                        strSql = "select lc15, lc05, lc06, lc07, lc11, lc22, lc43, lc44, '', LC16, LC17,lc23 as sp27,lc42 as sp78,lc45 as sp65,lc46 as sp66 from lawcase where lc01=" + CNULL(ot01) + " and lc02=" + CNULL(ot02) + " and lc03=" + CNULL(ot03) + " and lc04=" + CNULL(ot04)
             Case Else
                        'Modify By Cheng 2003/06/13
'                        strSQL = "select sp09,sp05,sp06,sp07,sp08,sp26,sp58,sp59 from servicepractice where sp01=" + CNULL(ot01) + " and sp02=" + CNULL(ot02) + " and sp03=" + CNULL(ot03) + " and sp04=" + CNULL(ot04)
                        'Modify By Cheng 2003/08/28
'                        strSQL = "select sp09, sp05, sp06, sp07, sp08, sp26, sp58, sp59, sp18 from servicepractice where sp01=" + CNULL(ot01) + " and sp02=" + CNULL(ot02) + " and sp03=" + CNULL(ot03) + " and sp04=" + CNULL(ot04)
                        'edit by nickc 2006/11/30
                        'strSQL = "select sp09, sp05, sp06, sp07, sp08, sp26, sp58, sp59, sp18, SP28, SP29 from servicepractice where sp01=" + CNULL(ot01) + " and sp02=" + CNULL(ot02) + " and sp03=" + CNULL(ot03) + " and sp04=" + CNULL(ot04)
                        'edit by nickc 2007/03/27 加入彼所案號
                        'strSQL = "select sp09, sp05, sp06, sp07, sp08, sp26, sp58, sp59, sp18, SP28, SP29,sp73,sp74 from servicepractice where sp01=" + CNULL(ot01) + " and sp02=" + CNULL(ot02) + " and sp03=" + CNULL(ot03) + " and sp04=" + CNULL(ot04)
                        'Modify By Sindy 2010/3/8 增加SP30,SP75
                        'Modify By Sindy 2011/1/18 +sp65,sp66
                        strSql = "select sp09, sp05, sp06, sp07, sp08, sp26, sp58, sp59, sp18, SP28, SP29,sp73,sp74,sp27,sp78,SP30,SP75,sp65,sp66 from servicepractice where sp01=" + CNULL(ot01) + " and sp02=" + CNULL(ot02) + " and sp03=" + CNULL(ot03) + " and sp04=" + CNULL(ot04)
End Select
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSql, cnnConnection
If rsRecordset.RecordCount > 0 Then
   ot05 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
   ot06 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
   ot07 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
   ot08 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
   ot09 = IIf(IsNull(rsRecordset.Fields(4)), "", rsRecordset.Fields(4)) '當事人1
   ot12 = IIf(IsNull(rsRecordset.Fields(5)), "", rsRecordset.Fields(5))
   ot10 = IIf(IsNull(rsRecordset.Fields(6)), "", rsRecordset.Fields(6)) '當事人2
   ot11 = IIf(IsNull(rsRecordset.Fields(7)), "", rsRecordset.Fields(7)) '當事人3
   'Add By Sindy 2011/1/18
   ot13 = IIf(IsNull(rsRecordset.Fields("sp65")), "", rsRecordset.Fields("sp65")) '當事人4
   ot14 = IIf(IsNull(rsRecordset.Fields("sp66")), "", rsRecordset.Fields("sp66")) '當事人5
   '2011/1/18 End
   'Add By Cheng 2003/06/13
   ot18 = IIf(IsNull(rsRecordset.Fields(8)), "", rsRecordset.Fields(8))
   'add by nickc 2007/03/27
   SP27 = IIf(IsNull(rsRecordset.Fields("sp27")), "", rsRecordset.Fields("sp27"))
   
    'Add By Cheng 2003/08/28
    If intCaseKind = 顧問 Then
        Me.Label15.Visible = False
        Me.txtOther(22).Visible = False
        Me.txtOther(22).Enabled = False
        Me.txtOther(23).Text = "" & rsRecordset.Fields(9).Value
    Else
        Me.txtOther(22).Text = "" & rsRecordset.Fields(10).Value
        Me.txtOther(23).Text = "" & rsRecordset.Fields(9).Value
    End If
    
    'add by nickc 2006/11/30
    If intCaseKind <> 顧問 And intCaseKind <> 法務 Then
        Me.txtOther(25).Text = "" & rsRecordset.Fields("SP73").Value
        Me.txtOther(26).Text = "" & rsRecordset.Fields("SP74").Value
        'Add By Sindy 2010/3/8
        strSP30s = IIf(IsNull(rsRecordset.Fields("SP30")), "", rsRecordset.Fields("SP30"))
        strSP75s = IIf(IsNull(rsRecordset.Fields("SP75")), "", rsRecordset.Fields("SP75"))
        '2010/3/8 End
    End If
   'Add by Morgan 2008/8/5
   strAppNo1 = "" & rsRecordset(4)
   'Modify by Amy 2021/12/21 改成Form 2.0
   'PUB_AddContact strAppNo1, cboContact, "" & rsRecordset("sp78"), True
   strExc(10) = cboContact.Tag
   PUB_AddContact strAppNo1, cboContact, "" & rsRecordset("sp78"), True, True, strExc(10)
   cboContact.Tag = strExc(10)
   'end 2008/8/5
   
   'Modify By Cheng 2002/01/03
   '若有申請人/當事人
   If Len("" & ot09) > 0 Then
      rsRecordset.Close
      strSql = "select cu30 from customer where cu01=" + CNULL(Mid(ot09, 1, 8)) + " AND cu02=" + CNULL(Mid(ot09, 9, 1))
      rsRecordset.CursorLocation = adUseClient
      rsRecordset.Open strSql, cnnConnection
      If rsRecordset.RecordCount > 0 Then
         cu30 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
         ot09 = ChangeCustomerS(ot09)
         ot10 = ChangeCustomerS(ot10)
         ot11 = ChangeCustomerS(ot11)
         ot12 = ChangeCustomerS(ot12)
         'Add By Sindy 2011/1/18
         ot13 = ChangeCustomerS(ot13)
         ot14 = ChangeCustomerS(ot14)
         '2011/1/18 End
         ReadOtherDatabase = True
      Else
         ShowMsg MsgText(1503)
      End If
   Else
      ReadOtherDatabase = True
   End If
Else
   If intModifyKind <> 0 Then
      ShowMsg "找不到此本所案號在基本檔之資料"
   End If
End If
'rsRecordset.Close

'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'If rsRecordset.State <> adStateClosed Then rsRecordset.Close
''add by nickc 2008/05/02 抓預定收款日
'strSql = "select rd05 from ReceivablesDay where (rd01,rd02,rd03) in (select rd01,rd02,max(rd03) from ReceivablesDay where (rd02) in (select max(rd02) from ReceivablesDay where rd01='" & CP09 & "' ) and rd01='" & CP09 & "' group by rd01,rd02) "
'rsRecordset.CursorLocation = adUseClient
'rsRecordset.Open strSql, cnnConnection
'If rsRecordset.RecordCount > 0 Then
'   txtOther(28) = IIf(IsNull(rsRecordset.Fields(0)), "", TAIWANDATE(rsRecordset.Fields(0)))
'Else
'   txtOther(28) = ""
'End If
'txtOther(28).Tag = txtOther(28) 'Add by Morgan 2010/12/9
'end 2018/08/22

rsRecordset.Close

Set rsRecordset = Nothing
Exit Function
ErrHand:
   ShowMsg "資料讀取失敗,請洽系統管理者!"  '2010/8/18 add by sonia
End Function

'從下一程序檔取回本所期限、法定期限
Private Function GetNextProgressDate(ByVal np02 As String, ByVal np03 As String, ByVal np04 As String, _
       ByVal np05 As String, ByVal NP07 As String, ByRef strDate1 As String, ByRef StrDate2 As String, ByRef strNP15 As String, _
       ByRef strNP10 As String) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset

On Error GoTo ErrHand
'NICK 900803 **********************
'strSQL = "select np08,np09,NP15 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
'          CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
'          " and np07=" + CNULL(np07) + " and (np06<>'Y' or np06 is null)"
'911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
'strSQL = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
          CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
          " and np07=" + CNULL(NP07) + " and (np06<>'Y' or np06 is null)"
'Added by Lydia 2023/03/22 法律案(L,FCL)不同審級收文期限沖銷：案件不同審級會收文-1、-2…案號，但實際是同一案件。
If np02 = "L" Or np02 = "FCL" Then
   strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
             CNULL(np03) + " and np07=" + CNULL(NP07) + " and  np06 is null "
Else
'end 2023/03/22
   strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
             CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
             " and np07=" + CNULL(NP07) + " and  np06 is null "
End If 'Added by Lydia 2023/03/22
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSql, cnnConnection, adOpenStatic
If rsRecordset.RecordCount > 0 Then
   rsRecordset.MoveLast
   rsRecordset.MoveFirst
   If rsRecordset.RecordCount = 1 Then
      strDate1 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
      StrDate2 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
      strNP15 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
      'Add By Cheng 2001/12/17
      strNP10 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
      'Added by Lydia 2017/12/21 檢查是否有不續辦相同性質且未到期的期限，若有則提醒操作人員注意要輸入接洽單上填寫的期限
      If frm010001.mRole = "" Then 'Added by Lydia 2024/10/18 排除外專/外商自行收文
         strExc(1) = Pub_GetNPDoubleMsg(DBDATE(txtOther(0).Text), np02, np03, np04, np05, NP07)
         If strExc(1) <> "" Then MsgBox strExc(1), vbExclamation + vbOKOnly
      End If 'Added by Lydia 2024/10/18
      'end 2017/12/21
      'Added by Lydia 2023/06/08
      If strDate1 <> "" Or StrDate2 <> "" Then
         bolisNP0809 = True
      End If
      'end 2023/06/08
   End If
Else
    '911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
    'Added by Lydia 2023/03/22 法律案(L,FCL)不同審級收文期限沖銷：案件不同審級會收文-1、-2…案號，但實際是同一案件。
    If np02 = "L" Or np02 = "FCL" Then
        strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
                  CNULL(np03) + " and np07=" + CNULL(NP07) + " and np06 <>'Y' "
    Else
    'end 2023/03/22
        strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
                  CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
                 " and np07=" + CNULL(NP07) + " and np06 <>'Y' "
    End If 'Added by Lydia 2023/03/22
    Set rsRecordset = New ADODB.Recordset
    rsRecordset.CursorLocation = adUseClient
    rsRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsRecordset.RecordCount = 1 Then
        strDate1 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
        StrDate2 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
        strNP15 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
        '取得下一程序資料檔之智權人員代號
        strNP10 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
        'Added by Lydia 2023/06/08
        If strDate1 <> "" Or StrDate2 <> "" Then
           bolisNP0809 = True
        End If
        'end 2023/06/08
    End If
End If
' **********************
GetNextProgressDate = True
rsRecordset.Close
Exit Function
ErrHand:
ShowMsg MsgText(1515)

End Function

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

'Add by Amy 2013/07/19 lblCaseProperty顯示（無）不可以存檔
   If lblCaseProperty = "（無）" Then
      MsgBox "案件性質錯誤!!", vbExclamation
      Exit Function
   End If
'end 2013/07/19

For Each objTxt In Me.txtOther
   If objTxt.Enabled = True Then
      Cancel = False
      txtOther_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Added by Lydia 2022/09/21
If Trim(txtOther(24)) & Trim(txtOther(4)) & Trim(txtOther(5)) & Trim(txtOther(6)) = "" Then
    MsgBox "案件名稱不可空白！", vbExclamation
    Exit Function
End If
'end 2022/09/21

   'Added by Lydia 2024/12/13 FG案輸入追蹤流水號TrackingNo
   If frm010001.intModifyKind = 0 And txtSystem = "FG" And txtCode(0) = "" And Left(m_SalesST15, 2) = "F2" Then
      FraTCN.Visible = True
      If Trim(txtTCN01) = "" Then
          MsgBox "追蹤流水號不可為空!!", vbExclamation
          txtTCN01.SetFocus
          Exit Function
      End If
      Cancel = False
      txtTCN01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   Else
      FraTCN.Visible = False
   End If
   'end 2024/12/13
   
TxtValidate = True
End Function

''Add By Cheng 2003/08/28
''比較點數與底價
'Private Function ChkPointValue(strCF01 As String, strCF02 As String, strCF03 As String, strPointValue As String, strClerk As String) As Boolean
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'Dim dblFreePoV As Double 'Add By Sindy 2010/3/23
'
'ChkPointValue = True
'If strCF02 = "" Then strCF02 = "000"
'StrSQLa = "Select * From CaseFee Where CF01='" & strCF01 & "' And CF02='" & strCF02 & "' And CF03='" & strCF03 & "' "
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'    If Val("" & rsA("CF14").Value) > 0 Then
'        'Add By Sindy 2010/3/23
'        dblFreePoV = 0
'        StrSQLa = "SELECT * FROM staff WHERE ST01='" & Trim(strClerk) & "' "
'        intI = 1
'        Set RsTemp = ClsLawReadRstMsg(intI, StrSQLa)
'        If intI = 1 Then
'           If Not IsNull(RsTemp("ST20")) Then
'              If Left(Trim(RsTemp("ST20")), 1) = "4" Then dblFreePoV = 1
'              If Left(Trim(RsTemp("ST20")), 1) = "3" Then dblFreePoV = 3
'           End If
'        End If
'        '2010/3/23 End
'        'Modify By Sindy 2010/3/23
'        'If Val("" & rsA("CF14").Value) > Val(strPointValue) Then
'        If Val("" & rsA("CF14").Value) > (Val(strPointValue) + dblFreePoV) Then
'        '2010/3/23 End
'            If MsgBox("您輸入的點數 (" & Val(strPointValue) & ") 低於底價 (" & Val("" & rsA("CF14").Value) & ")，請確認此客戶接洽單主管是否核示???", vbExclamation + vbYesNo) = vbNo Then
'                ChkPointValue = False
'            End If
'        End If
'    End If
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'End Function

'Added by Lydia 2016/04/25 新增查名單對應
'Memo by Lydia 2016/04/27 改成直接在畫面輸入查名代號
Private Sub cmdTSMap_Click()
  '要先輸入智權人員(預設查詢)
  If Me.txtOther(10).Text = "" Or lblSales.Caption = "" Then
     MsgBox "請先輸入智權人員!", vbInformation
     Exit Sub
  End If
  
  bolOpen130 = True
  
  If Len(txtRecieveCode) <> 9 Then
     ProcTSMap "0"
  Else
     ProcTSMap txtRecieveCode
  End If
End Sub
Private Sub ProcTSMap(ByRef tCP09 As String)
' iStiu '0:新增收文, 1:修改,  2:查詢
    Tmpfrm090130.SetParent Me
    Tmpfrm090130.iStiu = frm010001.intModifyKind
    Tmpfrm090130.mbolCall = False
    
    If tCP09 = "0" Then '未存檔完成
       Tmpfrm090130.m_CP09 = ""
    Else
       Tmpfrm090130.m_CP09 = tCP09
    End If
    Tmpfrm090130.m_CP13 = Me.txtOther(10).Text '智權人員(預設查詢)
    Tmpfrm090130.Show
    Tmpfrm090130.Caption = cmdTSMap.Caption
    Me.Hide
End Sub
'end 2016/04/25
'Added by Lydia 2016/04/27
Private Sub KillTemp()
On Error GoTo ErrHnd
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
   Exit Sub
   
ErrHnd:
   Resume Next
End Sub

Private Sub txtTS_GotFocus(Index As Integer)
'Memo by Lydia 2016/04/29 將TabIndex 設定在預定收款日的後面
    TextInverse txtTS(Index)
End Sub

Private Sub txtTS_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtTS_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        If txtTS(Index).Text > txtTS(Index).Tag Or Val(txtTS(Index).Text) - Val(txtTS(Index).Tag) < -1 Or Val(txtTS(Index).Text) - Val(txtTS(Index).Tag) > 1 Then
           MsgBox "年份輸入錯誤!!"
           Cancel = True
           txtTS(Index).SetFocus
        End If
    'Added by Lydia 2016/05/11 跨年度自動改民國年
    ElseIf Index = 1 Then
        If Val(Right(strSrvDate(1), 4)) < 200 And Val(txtTS(1)) > 600 Then
           txtTS(0).Text = Trim(Val(txtTS(0)) - 1)
           txtTS(0).Tag = txtTS(0)
        End If
    End If
End Sub
Private Sub SetTxtTS(ByVal RCode As String)
    txtTS(0) = Mid(RCode, 1, txtTS(0).MaxLength)
    txtTS(1) = Mid(RCode, txtTS(0).MaxLength + 1, txtTS(1).MaxLength)
End Sub
'end 2016/04/27

'Added by Lydia 2020/05/20 法律所案源收文：讀取法務案源檔
Private Sub ReadLOS()
Dim strR As String, intR As Integer
Dim rsRD As New ADODB.Recordset
     
    m_LOS15 = ""
    m_Los05_N = "": m_Los04_N1 = ""   'Added by Lydia 2020/10/05
    If frm010001.intModifyKind = 0 And strSrvDate(1) >= 法律所案源收文啟用日 Then
        If frm010001.txtLOS15 <> "" Then
            'Modified by Morgan 2022/11/10 B2類案源法務案收文預設接洽單期限--秀玲
            'strR = "select X.*,cp01,cp02,cp03,cp04 from LawOfficeSource X,caseprogress where los15=" & CNULL(frm010001.txtLOS15) & " and los01=cp09(+) "
            strR = "select X.*,cp01,cp02,cp03,cp04,CRL12,CRL13 from LawOfficeSource X,caseprogress,ConsultRecordList where los15=" & CNULL(frm010001.txtLOS15) & " and los01=cp09(+) and CRL01(+)=LOS17"
            'end 2022/11/10
            intR = 1
            Set rsRD = ClsLawReadRstMsg(intR, strR)
            If intR = 1 Then
                'Added by Morgan 2022/11/10 B2類案源法務案收文預設接洽單期限--秀玲
                If rsRD.Fields("LOS02") = "B2" Then
                  If Not IsNull(rsRD.Fields("CRL12")) Then txtOther(9) = TransDate(rsRD.Fields("CRL12"), 1)
                  If Not IsNull(rsRD.Fields("CRL13")) Then txtOther(14) = TransDate(rsRD.Fields("CRL13"), 1)
                End If
                'end2022/11/10
                '案源總收文號
                m_LOS01 = "" & rsRD.Fields("LOS01")
                '案源總收文號之本所案號
                m_LOS01cp01 = "" & rsRD.Fields("cp01")
                m_LOS01cp02 = "" & rsRD.Fields("cp02")
                m_LOS01cp03 = "" & rsRD.Fields("cp03")
                m_LOS01cp04 = "" & rsRD.Fields("cp04")
                '(原)案源案件類型
                m_LOS02 = "" & rsRD.Fields("LOS02")
                t_LOSkind = m_LOS02
                '案源單號
                m_LOS15 = "" & rsRD.Fields("LOS15")
                '介紹人, 介紹人(第一位)
                m_LOS04 = "" & rsRD.Fields("LOS04")
                If m_LOS04 <> "" And InStr(m_LOS04, ",") > 0 Then
                    m_LOS04_1 = Mid(m_LOS04, 1, InStr(m_LOS04, ",") - 1)
                Else
                    m_LOS04_1 = m_LOS04
                End If
                If m_LOS04_1 <> "" Then
                    m_LOS04_1st15 = GetST15(m_LOS04_1, , , m_LOS04_1st06)
                End If
                
                '(原)介紹客戶:
                m_LOS05 = "" & rsRD.Fields("LOS05")
                '介紹日
                m_LOS12 = "" & rsRD.Fields("LOS12")
            End If
        Else
            '案件性質=>案源案件類型 ; 前畫面未輸入案源編號也要另外判斷
            t_LOSkind = PUB_GetLOSkind(txtSystem, txtOther(1), txtOther(3))
        End If
    'Added by Morgan 2022/4/14
    ElseIf frm010001.intModifyKind = 1 Then
      If InStr(txtSystem, "L") > 0 Then
         'Modified by Lydia 2022/09/14
         'strR = "select cp162 from caseprogress where cp09='" & txtRecieveCode & "' and cp162 is not null"
         'Modified by Lydia 2022/09/21 +los04
         strR = "select los02,cp162,los04 from caseprogress, LawOfficeSource where cp09='" & txtRecieveCode & "' and cp162 is not null and cp162=los15(+) "
         intR = 1
         Set rsRD = ClsLawReadRstMsg(intR, strR)
         If intR = 1 Then
            m_LOS02 = "" & rsRD.Fields("los02")
            m_LOS15 = "" & rsRD.Fields("cp162")
            'Added by Lydia 2022/09/21 介紹人, 介紹人(第一位)
            m_LOS04 = "" & rsRD.Fields("LOS04")
            If m_LOS04 <> "" And InStr(m_LOS04, ",") > 0 Then
                m_LOS04_1 = Mid(m_LOS04, 1, InStr(m_LOS04, ",") - 1)
            Else
                m_LOS04_1 = m_LOS04
            End If
            If m_LOS04_1 <> "" Then
                m_LOS04_1st15 = GetST15(m_LOS04_1, , , m_LOS04_1st06)
            End If
            'end 2022/09/21
         End If
      End If
    'end 2022/4/14
    End If
    Set rsRD = Nothing
End Sub

''Added by Lydia 2020/06/01 法律所案源收文：抓P/T案收文的費用和規費(總計)
'Private Sub GetTotalFeeLOS01(ByVal pKeyNo As String, ByRef pFee1 As String, ByRef pFee2 As String, ByRef pOfficalFee As String)
''pKeyNo : 案源單號
''pFee1,  pFee2 : 費用起迄
''pOfficalFee: 規費
'Dim strQ1 As String, intQ As Integer
'Dim RsQ As New ADODB.Recordset
'
'    pFee1 = "0": pFee2 = "0": pOfficalFee = "0"
'
'    strQ1 = "select sum(nvl(cf06,0)) cnt1, sum(nvl(cf07,0)) cnt2, sum(nvl(cf08,0)) cnt3 from casefee where cf02='000' " & _
'                "and (cf01,cf03) in (select cp01,cp10 from caseprogress where cp09='" & pKeyNo & "' ) "
'    intQ = 1
'    Set RsQ = ClsLawReadRstMsg(intQ, strQ1)
'    If intQ = 1 Then
'        pFee1 = "" & RsQ.Fields("cnt1")
'        pFee2 = "" & RsQ.Fields("cnt2")
'        pOfficalFee = "" & RsQ.Fields("cnt3")
'    End If
'    Set RsQ = Nothing
'End Sub

'Added by Lydia 2022/07/15
Private Sub textEP06_GotFocus()
   TextInverse textEP06
End Sub
Private Sub textEP06_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub textEP34_GotFocus()
   TextInverse textEP34
End Sub
Private Sub textEP34_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2022/07/15 TC案之文件齊備日管控: 顯示
Private Sub setFrame21()
   If txtOther(3).Text <> txtOther(3).Tag Then   '判斷有修改才重設
        Frame21.Visible = False
        m_EP06 = "": textEP34.Enabled = True
        Label39.Visible = False: textEP34.Visible = False
        If txtSystem = "TC" Then
           Frame21.Visible = True
           '台灣TC案不會稿(不顯示),但大陸TC案要會稿
           If txtOther(3) = "000" Then
               Label39.Visible = False: textEP34.Visible = False
           ElseIf txtOther(3) = "020" Then
               Label39.Visible = True: textEP34.Visible = True
           End If
           If frm010001.intModifyKind = 0 Then
              '新增
           Else
              '讀取資料:
              strSql = "SELECT ep06,ep34 FROM engineerprogress WHERE ep02='" & txtRecieveCode & "' "
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strSql)
              If intI = 1 Then
                 If Not IsNull(RsTemp.Fields("ep06")) Then
                    If RsTemp.Fields("ep06") > 0 Then
                       textEP06.Text = "Y"
                    Else
                       textEP06.Text = "N"
                    End If
                 End If
                 If Not IsNull(RsTemp.Fields("ep34")) Then
                    textEP34.Text = RsTemp.Fields("ep34")
                 End If
              End If
           End If
           m_EP06 = textEP06
        End If

        txtOther(3).Tag = txtOther(3).Text
   End If

End Sub

'Added by Lydia 2022/07/15 TC案之文件齊備日管控: 存檔
Private Sub SaveFrame21(strCP09 As String)
   If Frame21.Visible = True Then
      
      If textEP06.Visible = True Then
          '文件是否齊備(101申請)、資料是否齊備
          If Trim(textEP06) = "Y" Then
             strSql = "update engineerprogress set ep06=" & strSrvDate(1) & ",ep36=" & strSrvDate(1) & " where ep02='" & strCP09 & "'"
             cnnConnection.Execute strSql
          ElseIf Trim(textEP06) = "N" Then
             strSql = "update engineerprogress set ep06=0,ep36=0 where ep02='" & strCP09 & "'"
             cnnConnection.Execute strSql
          ElseIf Trim(textEP06) = "" Then
             strSql = "update engineerprogress set ep06=null,ep36=null where ep02='" & strCP09 & "'"
             cnnConnection.Execute strSql
          End If
        '資料齊備
          If (m_EP06 = "" Or m_EP06 = "N") And Trim(textEP06) = "Y" Then
             strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                      " values('" & strCP09 & "','1','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & strSrvDate(1) & ",'收文')"
             cnnConnection.Execute strSql
          '收文取消齊備
          ElseIf m_EP06 = "Y" And (Trim(textEP06) = "N" Or Trim(textEP06) = "") Then
             strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                      " values('" & strCP09 & "','1','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",null,'收文取消齊備')"
             cnnConnection.Execute strSql
          End If
      End If
      '是否會稿
      If textEP34.Visible = True Then
            strSql = "update engineerprogress set ep34='" & textEP34 & "' where ep02='" & strCP09 & "'"
            cnnConnection.Execute strSql
      End If
   End If
End Sub

'Added by Lydia 2022/09/14 處理齊備日相關欄位的變數
Private Sub GetStrControl()
      m_strControl = ""
      If Frame21.Visible = True Then
         '資料是否齊備
         If textEP06.Visible = True Then
             m_strControl = m_strControl & ",EP06|" & Trim(textEP06) & "|" & m_EP06
         End If
         '是否會稿
         If textEP34.Visible = True Then
             m_strControl = m_strControl & ",EP34|" & Trim(textEP34)
         End If
         If m_strControl <> "" Then m_strControl = Mid(m_strControl, 2)
      End If
End Sub

'Added by Lydia 2024/12/13
Private Sub txtTCN01_Validate(Cancel As Boolean)
Dim ii As Integer
    If frm010001.intModifyKind = 0 And Left(m_SalesST15, 2) = "F2" And txtSystem = "FG" Then
        If Len(Trim(txtTCN01)) > 0 Then
            If Pub_ChkTCN01Status(Trim(txtTCN01), Trim(txtOther(10))) = False Then
               Cancel = True
               txtTCN01.SetFocus
               txtTCN01_GotFocus
               Exit Sub
            End If
        End If
    End If

End Sub

Private Sub txtTCN01_GotFocus()
   TextInverse txtTCN01
End Sub

