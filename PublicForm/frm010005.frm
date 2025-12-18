VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010005 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   6580
   ClientLeft      =   840
   ClientTop       =   1030
   ClientWidth     =   8640
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6580
   ScaleWidth      =   8640
   Begin VB.CommandButton cmdInventor 
      Caption         =   "發明人資料"
      Height          =   300
      Left            =   7020
      TabIndex        =   112
      Top             =   2520
      Width           =   1305
   End
   Begin VB.TextBox txtCP64 
      Height          =   360
      Left            =   6720
      TabIndex        =   98
      Top             =   6765
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   2
      Left            =   7704
      TabIndex        =   61
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   5748
      TabIndex        =   57
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   6576
      TabIndex        =   59
      Top             =   10
      Width           =   1100
   End
   Begin VB.Frame fraWindow1 
      BorderStyle     =   0  '沒有框線
      Height          =   6255
      Left            =   60
      TabIndex        =   45
      Top             =   330
      Width           =   8532
      Begin VB.TextBox txtRecieveCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1092
         TabIndex        =   46
         Top             =   0
         Width           =   1452
      End
      Begin VB.CheckBox Check2 
         Caption         =   "有★★的應收帳款簽核控管"
         Height          =   285
         Left            =   6150
         TabIndex        =   36
         Top             =   4980
         Width           =   2505
      End
      Begin VB.CheckBox chkWebApp 
         Caption         =   "電子送件"
         Height          =   255
         Left            =   2160
         TabIndex        =   118
         Top             =   4110
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         Caption         =   "聯絡人資料"
         Height          =   300
         Left            =   7290
         TabIndex        =   23
         Top             =   4050
         Width           =   1155
      End
      Begin VB.Frame fraTCT 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame1"
         Height          =   1425
         Left            =   1920
         TabIndex        =   119
         Top             =   4860
         Visible         =   0   'False
         Width           =   8715
         Begin VB.TextBox txtData 
            Enabled         =   0   'False
            Height          =   270
            Index           =   2
            Left            =   5040
            MaxLength       =   1
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtData 
            Height          =   270
            Index           =   3
            Left            =   960
            MaxLength       =   1
            TabIndex        =   47
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtData 
            Height          =   270
            Index           =   0
            Left            =   1320
            MaxLength       =   7
            TabIndex        =   43
            Top             =   0
            Width           =   855
         End
         Begin VB.TextBox txtData 
            Height          =   270
            Index           =   1
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   44
            Top             =   0
            Width           =   615
         End
         Begin VB.CheckBox ChkExpDate 
            Caption         =   "急件，請於　　　　　　　　　前譯畢名稱"
            Height          =   255
            Left            =   0
            TabIndex        =   124
            Top             =   0
            Width           =   3975
         End
         Begin MSForms.CheckBox ChkAdd968 
            Height          =   255
            Left            =   5880
            TabIndex        =   138
            Top             =   840
            Visible         =   0   'False
            Width           =   2055
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3625;450"
            Value           =   "0"
            Caption         =   "回復說明書校閱"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox ChkAdd435 
            Height          =   255
            Left            =   6960
            TabIndex        =   137
            Top             =   315
            Visible         =   0   'False
            Width           =   1575
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2778;450"
            Value           =   "0"
            Caption         =   "續行母案再審"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox ChkAdd228 
            Height          =   255
            Left            =   5880
            TabIndex        =   136
            Top             =   615
            Visible         =   0   'False
            Width           =   2055
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3625;450"
            Value           =   "0"
            Caption         =   "呈國際階段修正內容"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox ChkAdd106 
            Height          =   255
            Left            =   2880
            TabIndex        =   135
            Top             =   615
            Visible         =   0   'False
            Width           =   1695
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2990;450"
            Value           =   "0"
            Caption         =   "主張國際優先權"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox ChkAdd939 
            Height          =   255
            Left            =   1320
            TabIndex        =   134
            Top             =   620
            Visible         =   0   'False
            Width           =   975
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1720;450"
            Value           =   "0"
            Caption         =   "超項費"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox ChkAdd938 
            Height          =   255
            Left            =   0
            TabIndex        =   133
            Top             =   620
            Visible         =   0   'False
            Width           =   975
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1720;450"
            Value           =   "0"
            Caption         =   "超頁費"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox ChkAdd414 
            Height          =   255
            Left            =   5880
            TabIndex        =   132
            Top             =   315
            Visible         =   0   'False
            Width           =   1095
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1931;450"
            Value           =   "0"
            Caption         =   "恢復權利"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox ChkAdd902 
            Height          =   255
            Left            =   2880
            TabIndex        =   128
            Top             =   315
            Visible         =   0   'False
            Width           =   975
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1720;450"
            Value           =   "0"
            Caption         =   "回代"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox ChkAdd203 
            Height          =   255
            Left            =   1320
            TabIndex        =   130
            Top             =   315
            Visible         =   0   'False
            Width           =   1260
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2222;450"
            Value           =   "0"
            Caption         =   "主動修正"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.CheckBox ChkAdd416 
            Height          =   255
            Left            =   0
            TabIndex        =   127
            Top             =   315
            Visible         =   0   'False
            Width           =   975
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1720;450"
            Value           =   "0"
            Caption         =   "實審"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "5.核對中說235    6.檢視PCT公開本與FCP相異處)　"
            Height          =   180
            Left            =   1560
            TabIndex        =   126
            Top             =   1200
            Visible         =   0   'False
            Width           =   3915
         End
         Begin MSForms.CheckBox ChkAdd924 
            Height          =   255
            Left            =   4560
            TabIndex        =   131
            Top             =   315
            Visible         =   0   'False
            Width           =   975
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1720;450"
            Value           =   "0"
            Caption         =   "會稿"
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "(1.翻譯中說201　2.檢視中說209　3.製作中說210　4.製作中說210＆外文提申本242　"
            Height          =   180
            Left            =   1515
            TabIndex        =   123
            Top             =   1005
            Visible         =   0   'False
            Width           =   6630
         End
         Begin VB.Label Label36 
            Caption         =   "中說類型："
            Height          =   255
            Left            =   0
            TabIndex        =   122
            Top             =   975
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "(1. 電子電機　2.化學　3.日文　4.機械設計　B.退程序）"
            Height          =   180
            Left            =   5595
            TabIndex        =   121
            Top             =   105
            Visible         =   0   'False
            Width           =   4410
         End
         Begin VB.Label Label34 
            Caption         =   "分案組別："
            Height          =   255
            Left            =   4080
            TabIndex        =   120
            Top             =   75
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.TextBox txtTCN01 
         Height          =   300
         Left            =   1050
         MaxLength       =   9
         TabIndex        =   32
         Top             =   4980
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.TextBox txtCopy 
         Height          =   300
         Left            =   7965
         MaxLength       =   2
         TabIndex        =   27
         Top             =   4380
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "現金或支票"
         Height          =   285
         Left            =   4770
         TabIndex        =   30
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox textYear 
         Height          =   270
         Left            =   4350
         MaxLength       =   5
         TabIndex        =   34
         Top             =   4980
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   5460
         MaxLength       =   5
         TabIndex        =   35
         Top             =   4980
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CheckBox chkEnglish 
         Caption         =   "同時申請三國(含)以上之美日德"
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   4350
         Width           =   2800
      End
      Begin VB.Frame fraPromoter 
         BorderStyle     =   0  '沒有框線
         Height          =   320
         Left            =   90
         TabIndex        =   95
         Top             =   4950
         Width           =   3015
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   24
            Left            =   960
            TabIndex        =   33
            Top             =   0
            Width           =   1095
            VariousPropertyBits=   679493659
            MaxLength       =   6
            Size            =   "1931;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPromoter 
            Height          =   255
            Left            =   2100
            TabIndex        =   96
            Top             =   0
            Width           =   825
            VariousPropertyBits=   27
            Size            =   "1455;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label25 
            Caption         =   "承辦人："
            Height          =   255
            Left            =   0
            TabIndex        =   97
            Top             =   30
            Width           =   975
         End
      End
      Begin VB.Frame fraPatition 
         BorderStyle     =   0  '沒有框線
         Height          =   1005
         Left            =   60
         TabIndex        =   93
         Top             =   5280
         Visible         =   0   'False
         Width           =   8355
         Begin MSForms.TextBox txtPetitionx 
            Height          =   300
            Index           =   5
            Left            =   1125
            TabIndex        =   42
            Top             =   630
            Width           =   1125
            VariousPropertyBits=   16411
            MaxLength       =   9
            Size            =   "1984;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPetitionx 
            Height          =   300
            Index           =   4
            Left            =   5550
            TabIndex        =   41
            Top             =   330
            Width           =   1125
            VariousPropertyBits=   16411
            MaxLength       =   9
            Size            =   "1984;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPetitionx 
            Height          =   300
            Index           =   3
            Left            =   1125
            TabIndex        =   40
            Top             =   330
            Width           =   1125
            VariousPropertyBits=   16411
            MaxLength       =   9
            Size            =   "1984;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPetitionx 
            Height          =   300
            Index           =   2
            Left            =   5550
            TabIndex        =   39
            Top             =   30
            Width           =   1125
            VariousPropertyBits=   16411
            MaxLength       =   9
            Size            =   "1984;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   23
            Left            =   1125
            TabIndex        =   38
            Top             =   30
            Width           =   1125
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1984;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "讓與申請人1："
            Height          =   180
            Left            =   0
            TabIndex        =   109
            Top             =   60
            Width           =   1170
         End
         Begin MSForms.Label lblPetitionNamex 
            Height          =   255
            Index           =   5
            Left            =   2265
            TabIndex        =   108
            Top             =   630
            Width           =   2010
            VariousPropertyBits=   27
            Size            =   "3545;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "讓與申請人4："
            Height          =   180
            Index           =   4
            Left            =   4365
            TabIndex        =   107
            Top             =   330
            Width           =   1170
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "讓與申請人5："
            Height          =   180
            Index           =   5
            Left            =   0
            TabIndex        =   106
            Top             =   630
            Width           =   1170
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "讓與申請人3："
            Height          =   180
            Index           =   3
            Left            =   0
            TabIndex        =   105
            Top             =   330
            Width           =   1170
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "讓與申請人2："
            Height          =   180
            Index           =   2
            Left            =   4365
            TabIndex        =   104
            Top             =   60
            Width           =   1170
         End
         Begin MSForms.Label lblPetitionNamex 
            Height          =   255
            Index           =   4
            Left            =   6700
            TabIndex        =   103
            Top             =   330
            Width           =   1600
            VariousPropertyBits=   27
            Size            =   "2822;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetitionNamex 
            Height          =   255
            Index           =   3
            Left            =   2265
            TabIndex        =   102
            Top             =   330
            Width           =   2010
            VariousPropertyBits=   27
            Size            =   "3545;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetitionNamex 
            Height          =   255
            Index           =   2
            Left            =   6700
            TabIndex        =   101
            Top             =   0
            Width           =   1600
            VariousPropertyBits=   27
            Size            =   "2822;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetitionName 
            Height          =   255
            Left            =   2265
            TabIndex        =   94
            Top             =   30
            Width           =   2010
            VariousPropertyBits=   27
            Size            =   "3545;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
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
         Height          =   2865
         Left            =   96
         TabIndex        =   64
         Top             =   600
         Width           =   8292
         Begin VB.ComboBox Combo3 
            Height          =   300
            ItemData        =   "frm010005.frx":0000
            Left            =   5430
            List            =   "frm010005.frx":000D
            TabIndex        =   3
            Top             =   150
            Width           =   1785
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   288
            Index           =   2
            Left            =   3720
            MaxLength       =   2
            TabIndex        =   79
            Top             =   150
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   288
            Index           =   1
            Left            =   3240
            MaxLength       =   1
            TabIndex        =   78
            Top             =   150
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   288
            Index           =   0
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   77
            Top             =   150
            Width           =   1212
         End
         Begin VB.TextBox txtSystem 
            Enabled         =   0   'False
            Height          =   288
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   65
            Top             =   150
            Width           =   732
         End
         Begin MSForms.ComboBox cboContact 
            Height          =   300
            Left            =   5040
            TabIndex        =   139
            Top             =   1650
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
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   27
            Left            =   5640
            TabIndex        =   15
            Top             =   2520
            Width           =   2565
            VariousPropertyBits=   679493659
            MaxLength       =   50
            Size            =   "4524;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   4
            Left            =   4920
            TabIndex        =   5
            Top             =   435
            Width           =   612
            VariousPropertyBits=   679493659
            MaxLength       =   3
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   11
            Left            =   930
            TabIndex        =   10
            Top             =   1950
            Width           =   1092
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   12
            Left            =   930
            TabIndex        =   12
            Top             =   2235
            Width           =   1092
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   13
            Left            =   930
            TabIndex        =   14
            Top             =   2520
            Width           =   1092
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   10
            Left            =   5025
            TabIndex        =   13
            Top             =   2235
            Width           =   1092
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   3
            Left            =   1080
            TabIndex        =   4
            Top             =   435
            Width           =   372
            VariousPropertyBits=   679493659
            MaxLength       =   1
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   9
            Left            =   5025
            TabIndex        =   11
            Top             =   1950
            Width           =   1092
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   8
            Left            =   930
            TabIndex        =   9
            Top             =   1650
            Width           =   1092
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   7
            Left            =   1920
            TabIndex        =   8
            Top             =   1335
            Width           =   6252
            VariousPropertyBits=   679493659
            Size            =   "11028;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   6
            Left            =   1920
            TabIndex        =   7
            Top             =   1050
            Width           =   6252
            VariousPropertyBits=   679493659
            Size            =   "11028;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtPatent 
            Height          =   300
            Index           =   5
            Left            =   1920
            TabIndex        =   6
            Top             =   750
            Width           =   6252
            VariousPropertyBits=   679493659
            Size            =   "11028;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "案件屬性："
            Height          =   180
            Index           =   168
            Left            =   4470
            TabIndex        =   116
            Top             =   150
            Width           =   900
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "接洽人："
            Height          =   180
            Left            =   4215
            TabIndex        =   111
            Top             =   1710
            Width           =   720
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "代理人彼所案號："
            Height          =   180
            Left            =   4215
            TabIndex        =   110
            Top             =   2580
            Width           =   1440
         End
         Begin VB.Label lblPetition 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   2130
            TabIndex        =   88
            Top             =   2250
            Width           =   1965
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "申請人5："
            Height          =   180
            Left            =   4215
            TabIndex        =   87
            Top             =   2280
            Width           =   810
         End
         Begin VB.Label lblTrademarkKind 
            Height          =   195
            Left            =   1560
            TabIndex        =   86
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label10 
            Caption         =   "專利種類："
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   465
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "專利名稱(外)（160）："
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   1380
            Width           =   1935
         End
         Begin VB.Label Label13 
            Caption         =   "專利名稱(英)（250）："
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label11 
            Caption         =   "專利名稱(中)（160）："
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   82
            Top             =   780
            Width           =   1935
         End
         Begin MSForms.Label lblAgent 
            Height          =   255
            Left            =   2130
            TabIndex        =   81
            Top             =   2550
            Width           =   1965
            VariousPropertyBits=   27
            Size            =   "3466;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label8 
            Caption         =   "代理人："
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   2550
            Width           =   855
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "申請人3："
            Height          =   180
            Left            =   4215
            TabIndex        =   76
            Top             =   1995
            Width           =   810
         End
         Begin VB.Label lblPetition 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   6225
            TabIndex        =   75
            Top             =   2250
            Width           =   1965
         End
         Begin VB.Label lblPetition 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   2130
            TabIndex        =   74
            Top             =   1965
            Width           =   1965
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "申請人2："
            Height          =   180
            Left            =   120
            TabIndex        =   73
            Top             =   2010
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "申請人4："
            Height          =   180
            Left            =   120
            TabIndex        =   72
            Top             =   2295
            Width           =   810
         End
         Begin VB.Label lblPetition 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   2130
            TabIndex        =   71
            Top             =   1665
            Width           =   1965
         End
         Begin VB.Label lblPetition 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   6225
            TabIndex        =   70
            Top             =   1965
            Width           =   1965
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "申請人1："
            Height          =   180
            Left            =   120
            TabIndex        =   69
            Top             =   1680
            Width           =   810
         End
         Begin VB.Label Label9 
            Caption         =   "申請國家："
            Height          =   255
            Left            =   3960
            TabIndex        =   68
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblNation 
            Height          =   255
            Left            =   5640
            TabIndex        =   67
            Top             =   510
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "本所案號："
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   180
            Width           =   975
         End
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   1
         Left            =   1095
         TabIndex        =   1
         Top             =   300
         Width           =   600
         VariousPropertyBits=   679493659
         MaxLength       =   4
         Size            =   "1058;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   0
         Left            =   5040
         TabIndex        =   0
         Top             =   0
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   2
         Left            =   5040
         TabIndex        =   2
         Top             =   300
         Width           =   375
         VariousPropertyBits=   679493659
         MaxLength       =   2
         Size            =   "656;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCaseSource 
         Height          =   255
         Left            =   5490
         TabIndex        =   50
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblCaseProperty 
         Height          =   255
         Left            =   1770
         TabIndex        =   49
         Top             =   390
         Width           =   2175
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   29
         Left            =   1050
         TabIndex        =   37
         Top             =   5280
         Width           =   315
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "556;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label28 
         Caption         =   "(1.電子 2.紙本)"
         Height          =   315
         Index           =   1
         Left            =   1410
         TabIndex        =   141
         Top             =   5340
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "證書形式："
         Height          =   285
         Index           =   141
         Left            =   90
         TabIndex        =   140
         Top             =   5340
         Width           =   945
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   22
         Left            =   6780
         TabIndex        =   31
         Top             =   4650
         Width           =   1092
         VariousPropertyBits=   679493659
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   18
         Left            =   1050
         TabIndex        =   28
         Top             =   4620
         Width           =   1092
         VariousPropertyBits=   679493659
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   28
         Left            =   3330
         TabIndex        =   29
         Top             =   4620
         Width           =   1092
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   26
         Left            =   6420
         TabIndex        =   20
         Top             =   3765
         Width           =   1995
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   25
         Left            =   5670
         TabIndex        =   18
         Top             =   3480
         Width           =   2745
         VariousPropertyBits=   679493659
         Size            =   "4842;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   20
         Left            =   5940
         TabIndex        =   22
         Top             =   4080
         Width           =   435
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "767;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   19
         Left            =   1050
         TabIndex        =   16
         Top             =   3480
         Width           =   1092
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   14
         Left            =   3168
         TabIndex        =   17
         Top             =   3480
         Width           =   1092
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   17
         Left            =   1050
         TabIndex        =   24
         Top             =   4335
         Width           =   1092
         VariousPropertyBits=   679493659
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   15
         Left            =   1050
         TabIndex        =   19
         Top             =   3765
         Width           =   1092
         VariousPropertyBits=   679493659
         MaxLength       =   6
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   21
         Left            =   5670
         TabIndex        =   26
         Top             =   4380
         Width           =   1092
         VariousPropertyBits=   679493659
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPatent 
         Height          =   300
         Index           =   16
         Left            =   1050
         TabIndex        =   21
         Top             =   4050
         Width           =   1092
         VariousPropertyBits=   679493659
         MaxLength       =   5
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LbTracking 
         AutoSize        =   -1  'True
         Caption         =   "追蹤流水號："
         Height          =   180
         Left            =   60
         TabIndex        =   125
         Top             =   5000
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label21 
         Caption         =   "郵遞區號："
         Height          =   255
         Left            =   90
         TabIndex        =   58
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label lblCopy 
         AutoSize        =   -1  'True
         Caption         =   "優先權份數："
         Height          =   180
         Left            =   6885
         TabIndex        =   117
         Top             =   4410
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "繳費年度：第            年至第            年"
         Height          =   180
         Index           =   1
         Left            =   3210
         TabIndex        =   115
         Top             =   5040
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.Label Label27 
         Caption         =   "後金："
         Height          =   255
         Left            =   6180
         TabIndex        =   114
         Top             =   4710
         Width           =   615
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "預定收款日："
         Height          =   180
         Left            =   2220
         TabIndex        =   113
         Top             =   4680
         Width           =   1080
      End
      Begin VB.Label Label30 
         Caption         =   "分所案號："
         Height          =   255
         Left            =   5400
         TabIndex        =   100
         Top             =   3780
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "客戶案件案號："
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   99
         Top             =   3480
         Width           =   1365
      End
      Begin VB.Label Label19 
         Caption         =   "是否開電腦收據：         （N：不開)"
         Height          =   255
         Left            =   4500
         TabIndex        =   92
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Label lblDepartment 
         Height          =   255
         Left            =   4320
         TabIndex        =   91
         Top             =   3780
         Width           =   945
      End
      Begin VB.Label Label18 
         Caption         =   "業務區："
         Height          =   255
         Left            =   3360
         TabIndex        =   90
         Top             =   3780
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "法定期限："
         Height          =   252
         Left            =   84
         TabIndex        =   89
         Top             =   3504
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "收文號："
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   36
         Width           =   720
      End
      Begin VB.Label Label7 
         Caption         =   "收文日："
         Height          =   252
         Left            =   4080
         TabIndex        =   62
         Top             =   36
         Width           =   852
      End
      Begin VB.Label Label20 
         Caption         =   "本所期限："
         Height          =   252
         Left            =   2232
         TabIndex        =   60
         Top             =   3504
         Width           =   900
      End
      Begin VB.Label Label22 
         Caption         =   "費用："
         Height          =   255
         Left            =   90
         TabIndex        =   56
         Top             =   4350
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "點數："
         Height          =   255
         Left            =   90
         TabIndex        =   55
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "智權人員："
         Height          =   255
         Left            =   90
         TabIndex        =   54
         Top             =   3780
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "規費："
         Height          =   255
         Left            =   5070
         TabIndex        =   53
         Top             =   4380
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "案件性質："
         Height          =   252
         Left            =   120
         TabIndex        =   52
         Top             =   336
         Width           =   972
      End
      Begin VB.Label Label5 
         Caption         =   "案件來源："
         Height          =   252
         Index           =   0
         Left            =   4080
         TabIndex        =   51
         Top             =   336
         Width           =   960
      End
      Begin MSForms.Label lblSales 
         Height          =   255
         Left            =   2250
         TabIndex        =   48
         Top             =   3780
         Width           =   945
         VariousPropertyBits=   27
         Size            =   "1667;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
End
Attribute VB_Name = "frm010005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/16 Form2.0已修改 txtPatent/txtPetitionx/lblPetitionName/lblPetitionNamex/lblSales/lblAgent/lblPromoter/cboContact
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

'bolLeave判斷離開時，是否要彈出詢問視窗
'LastData上一次存檔時，所輸入之收文日
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim bolLeave As Boolean, LastDate As String, intLeaveKind As Integer
Dim strNation As String, douStPrice As Double, douLowPrice As Double
'20080822 add by Toni
'strAddDeadline存放發明人資料
Dim strInventorNo As String
Dim strInventorNo_Old As String 'Add By Sindy 2014/11/6 記錄原發明人
Dim strInventor(100) As String     'add by toni 2008/8/26 寫發明人data use
'Dim strIntor As String
Dim varInventorNo As Variant
Dim strPetition  As String

' 91.09.11 modify by louis
' 目前准駁
Dim m_PA16 As String
Dim m_PA14 As String           '2010/8/17 ADD BY SONIA
Dim m_PA91 As String 'Added by Lydiad 2021/04/15 案件備註

'Add by Morgan 2004/4/15
'是否已觸發 Form Active 事件
Dim bolActive As Boolean
'add by nickc 2007/12/12
Dim IsSaveData As Boolean
Dim strAppNo1 As String '申請人1編號
'Add By Sindy 2009/07/08
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_CP10 As String

'Add By Sindy 2010/3/8 回傳值
Dim strPA51s As String, strPA52s As String, strPA53s As String
Dim strPA54s As String, strPA55s As String, strPA56s As String
Dim bolCancel As Boolean
'2010/3/8 End
Dim dblAmt As Double, dblPFee As Double, dblTFee As Double, m_CP150 As String 'Add By Sindy 2012/11/06
Dim dblChkAmt As Double 'Add By Sindy 2012/12/10
'Added by Lydia 2020/02/03
Dim dblCu183 As Double '個人之應收帳款上限
Dim dblAmtR As Double, dblPFeeR As Double, dblTFeeR As Double '關係企業之應收帳款金額
'end 2020/02/03

Dim ii As Integer 'Add by Amy 2013/06/26
Dim mFMPchk As Boolean, mCP31 As String 'Add by Lydia 2014/10/31 開放外專程序人員可進入專利處系統操作FMP寰華案件
'Added by Lydia 2017/11/14 FCP案件命名電子化
Dim bolExistTCT As Boolean '是否存在中說輸入相關設定
Dim m_TCT01 As String '案件名稱資料檔PK
Dim m_TCT04 As String '分案主管
Dim mSaveDir As String 'Added by Lydia 2018/03/01 TrackingNO：暫存下載檔案的本機端資料夾
Dim bolMoveOK As Boolean 'Added by Lydia 2020/02/13 TrackingNO是否已搬檔完成(True無問題)，若有問題則TrackingNO和本機端的資料夾不刪除
Dim bolMoveCheck As Boolean 'Added by Lydia 2020/04/29 存檔前先下載TrackingNO檔案是否完成
Dim m_SalesST15 As String 'Addded by Lydia 2018/09/06 畫面上智權人員的收文部門
Dim m_Tuser As String 'Added by Lydia 2019/02/14 創新業務部預設收文人員
'Added by Lydia 2019/09/16
Dim m_SalesST06 As String '智權人員的所別
Public m_bBatch As Boolean, m_CP11 As String, m_CP13 As String 'Added by Morgan 2020/4/8 整批收文
'Added by Lydia 2020/05/20 法律所案源收文
Dim m_LOS02 As String '案源案件類型
Dim t_LOSkind As String '案件性質=>案源案件類型
Dim m_LOS15 As String '案源單號
Dim m_CaseNa239() As String   ''Added by Lydia 2020/11/19 CFP英國脫歐案管制：歐盟案案號
Dim strXState(1 To 50) As String, strYState As String 'Add By Sindy 2021/2/1 回傳客戶狀態
Dim m_strCPM34 As String, m_strCP06 As String 'Add By Sindy 2021/4/29
Dim strPA176 As String 'Added by Morgan 2021/7/21
Dim m_bolFMP As Boolean 'Add By Sindy 2022/3/22
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/04/13 (舊案)是否為寰華案
'Mark by Lydia 2022/09/06 改抓特殊設定
'Private Const cnt應收帳款檢查排除 As String = "74018,70005" 'Added by Lydia 2022/06/15 應收帳款上限檢查排除特定人員: 如果人員有異動, 請一併修改接洽單frm090801和收文frm010004~frm010007
'Add By Sindy 2022/6/29
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_PrevForm As Form '前一畫面
Public m_bMRecvBatch As Boolean '信件沖銷多案收文
'2022/6/29 END
Dim m_bolRecvOK As Boolean 'Add By Sindy 2022/7/8 是否收完文
Dim m_strMCR11 As String 'Add By Sindy 2022/7/8 多案收文時,第一筆的總收文號

'Added by Lydia 2022/09/05 櫃台收文模組化
Private Const 收文存檔模組化啟用日 = 20220916 '完成後先開始使用
Dim modCP() As String, modBase() As String ' 收文 和 基本檔
Dim mChkStr As String   '其他操作結果
Dim mType As String, mCaseNo As String  '特殊管制
Dim m_NowCP16 As String, m_NowCP17 As String, m_NowCP18 As String '(國外部)計算後的費用、規費、點數
Dim mTCTVal As String '傳入畫面有關命名作業的資料
Dim mTCTList As String  '回傳命名作業一併產生之收文號
Dim bolisNP0809 As Boolean 'Added by Lydia 2023/06/08 是否從下一程序檔取回本所期限、法定期限

'Added by Lydia 2022/09/05 設定陣列
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
      If pCD01 <> "" And pCD02 <> "" Then
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
      modBase(5) = txtPatent(5)  '案件名稱(中)
      modBase(6) = txtPatent(6)  '案件名稱(英)
      modBase(7) = txtPatent(7)  '案件名稱(日)
      modBase(8) = txtPatent(3) '專利種類
      modBase(9) = Trim(txtPatent(4)) '申請國家
      '申請人1~5
      modBase(26) = ChangeCustomerL(txtPatent(8))
      modBase(27) = ChangeCustomerL(txtPatent(11))
      modBase(28) = ChangeCustomerL(txtPatent(9))
      modBase(29) = ChangeCustomerL(txtPatent(12))
      modBase(30) = ChangeCustomerL(txtPatent(10))
      modBase(48) = Trim(txtPatent(25)) '客戶案件案號
      modBase(47) = Trim(txtPatent(26)) '分所案號
      '代理人
      modBase(75) = ChangeCustomerL(txtPatent(13))
      modBase(77) = Trim(txtPatent(27)) '彼所案號
      '聯絡人1~2 =>  frm010007_1.bolOK
      If bolCancel = True Then
          modBase(51) = strPA51s
          modBase(52) = strPA52s
          modBase(53) = strPA53s
          modBase(54) = strPA54s
          modBase(55) = strPA55s
          modBase(56) = strPA56s
      End If
      modBase(158) = Left(Combo3, 1) '案件屬性
      'FCP工程師組別
      If fraTCT.Visible = True And txtData(2).Text <> "" And txtData(2) <> "B" Then
         modBase(150) = txtData(2)
      End If
      
      '大陸發明生醫案是否新藥專利設定
      If strPA176 <> "" Then
          modBase(176) = strPA176
      End If

      '申請人聯絡人編號
      If cboContact.Locked = False Then
         If cboContact.ListIndex >= 0 Then
            modBase(149) = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
            If Val(modBase(149)) > 0 Then
            'Add by Morgan 2008/8/7 若個案接洽人與客戶檔的預設接洽人相同時不必設定
               PUB_GetContact modBase(26), strTmpA, True
               If modBase(149) = strTmpA Then
                  modBase(149) = ""
               End If
            '排除空白=00
            ElseIf modBase(149) = "00" And Trim(cboContact.Text) = "" Then
               modBase(149) = ""
            End If
         End If
      End If
      'Add By Sindy 2022/12/7
      If txtPatent(29).Visible = True Then
         modBase(178) = Trim(txtPatent(29)) '證書形式
      End If
      '2022/12/7
      
      modCP(9) = txtRecieveCode  '收文號
      modCP(5) = ChangeTStringToWString(txtPatent(0)) '收文日
      modCP(6) = ChangeTStringToWString(txtPatent(14)) '本所期限
      modCP(7) = ChangeTStringToWString(txtPatent(19))  '法定期限
      modCP(10) = Trim(txtPatent(1)) '案件性質
      modCP(11) = Trim(txtPatent(2)) '案件來源
      modCP(12) = GetST15(txtPatent(15))
      modCP(13) = Trim(txtPatent(15))       '智權人員
      modCP(14) = Trim(txtPatent(24))    '承辦人
      'Added by Lydia 2023/04/13 FMP案收文(601)領證或(605)繳年費時，非寰華案請預設承辦人
      'Modified by Lydia 2023/04/25 若為新案或已閉卷可以不輸入年度，但不預設承辦人+ And Val(textYear) > 0 And Val(Text1(0)) > 0
      If modCP(14) = "" And m_bolFMP = True And m_bolFMP2 = False And Val(modBase(58)) = 0 And _
           txtSystem = "P" And modCP(2) <> "" And InStr("601,605,", modCP(10)) > 0 And Val(textYear) > 0 And Val(Text1(0)) > 0 Then
           '參考接洽單自動收文P案601、605之承辦人預設規則設定，但新案號或是已閉卷案件不預設。
           'Added by Morgan 2025/2/12
            If strSrvDate(1) >= P業務區劃分啟用日 Then
               modCP(14) = PUB_GetPHandler(modCP(1) & modCP(2) & modCP(3) & modCP(4))
            Else
            'end 2025/2/12
               modCP(14) = Pub_GetSpecMan("A111") 'P非臺灣案領證、繳年費
            End If
      End If
      'end 2023/04/13
      
      modCP(16) = txtPatent(17)    '費用
      modCP(17) = txtPatent(21)    '規費
      modCP(18) = txtPatent(18)    '點數
      modCP(19) = txtPatent(22)    '後金
      modCP(32) = txtPatent(20) '是否開電腦收據
      modCP(33) = douStPrice '標準價
      modCP(34) = douLowPrice '底價
      modCP(64) = txtCP64
      '讓與人1-5,受讓人1-5
      modCP(56) = ChangeCustomerL(txtPatent(23))
      modCP(89) = ChangeCustomerL(txtPetitionx(2))
      modCP(90) = ChangeCustomerL(txtPetitionx(3))
      modCP(91) = ChangeCustomerL(txtPetitionx(4))
      modCP(92) = ChangeCustomerL(txtPetitionx(5))
      '有★★的應收帳款簽核控管
      If Check2.Visible = True Then
         modCP(150) = IIf(Check2.Value = 1, "Y", "")
      End If
      '優先權份數
      If txtCopy.Visible = True Then
          modCP(71) = Val(txtCopy)
      End If
      '電子送件
      'Modified by Lydia 2022/09/20 因為P案有Trigger會自動設定電子送件CP118 = Y, 所以改成兩個判斷
      'If chkWebApp.Visible = True And chkWebApp.Value = 1 Then
          'modCP(118) = "Y"
      If chkWebApp.Visible = True Then
         If chkWebApp.Value = 1 Then
             modCP(118) = "YY"
         Else
             modCP(118) = "YN"
         End If
      'end 2022/09/20
      End If
      '年費期間
      If textYear.Visible = True Then
          modCP(53) = Val(textYear)
          modCP(54) = Val(Text1(0))
      End If
      
      '特殊管制
      mType = "": mCaseNo = ""
      If txtSystem = "CFP" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" Then
          mType = "CFP英國脫歐案"
          mCaseNo = m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4)
      ElseIf m_LOS02 <> "" And m_LOS15 <> "" Then
          mType = "LOS案源收文"
          mCaseNo = m_LOS02 & "," & m_LOS15
      'Modify By Sindy 2025/8/18 發生了案源+信件沖銷 ex:FCP-057445/FCL-011034
      'ElseIf m_strIR01 <> "" Then
      End If
      If m_strIR01 <> "" Then
      '2025/8/18 END
          'Modify By Sindy 2023/5/31
          'mType = "外專信件沖銷"
'          mType = "信件沖銷"
          mType = mType & "-信件沖銷" 'Modify By Sindy 2025/8/18 + "-"
          '2023/5/31 END
          If m_bMRecvBatch = True Then mType = mType & "-多案收文"
          'Modify By Sindy 2025/8/18 + IIf(mCaseNo <> "", mCaseNo & "-", "") &
          mCaseNo = IIf(mCaseNo <> "", mCaseNo & "-", "") & m_strIR01 & "," & m_strIR02 & "," & m_strIR03 & "," & m_strIR04
      End If
                  
      '傳入其他操作結果
      mChkStr = ","
      If Trim(txtSystem) = "P" And (mCP31 = "Y" Or frm010001.intSaveMode = 1) And m_LOS15 = "" And mFMPchk = True Then
          mChkStr = mChkStr & "寰華案件確認,"
      End If
      
      'Add By Sindy 2022/3/22 是否為FMP案件
      If PUB_ChkIsFMP(txtSystem, txtCode(0), txtCode(1), txtCode(2)) = True Or _
        (txtSystem = "P" And (frm010001.txtFMP = "Y" And frm010001.txtNA01 = "020") And InStr(NewCasePtyList, txtPatent(1)) > 0) Then
        mChkStr = mChkStr & "m_bolFMP,"
      End If
      'Added by Lydia 2018/09/06 櫃臺P案(FMP)中間程序和中間接進來案件的收文，系統自動發e-mail通知。
      If frm010001.intModifyKind = 0 And txtSystem = "P" And Left(m_SalesST15, 2) = "F2" And frm010001.mRole = "" And InStr(FcpAddTct, txtPatent(1)) = 0 Then
         mChkStr = mChkStr & "FMP中間程序EMAIL通知,"
      End If
      
      'Add By Sindy 2021/4/29 不是主管機關期限
      If (txtSystem = "FCP" Or txtSystem = "FG") And m_strCPM34 = "N" And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
         If Val(m_strCP06) > 0 And DBDATE(m_strCP06) <> DBDATE(modCP(6)) Then
             mChkStr = mChkStr & "原本所期限為" & ChangeWStringToTDateString(DBDATE(m_strCP06)) & ","
         End If
      End If
      
      '傳入畫面有關命名作業的資料
      mTCTVal = "": mTCTList = ""
      If fraTCT.Visible = True And fraTCT.Enabled = True Then
           '中說類型
           mTCTVal = Trim(txtData(3)) & "|"
           '勾選的收文性質
           If ChkAdd416.Visible = True And ChkAdd416.Value = True Then mTCTVal = mTCTVal & "416,"
           If ChkAdd203.Visible = True And ChkAdd203.Value = True Then mTCTVal = mTCTVal & "203,"
           If ChkAdd902.Visible = True And ChkAdd902.Value = True Then mTCTVal = mTCTVal & "902,"
           If ChkAdd924.Visible = True And ChkAdd924.Value = True Then mTCTVal = mTCTVal & "924,"
           If ChkAdd968.Visible = True And ChkAdd968.Value = True Then mTCTVal = mTCTVal & "968,"
           If ChkAdd414.Visible = True And ChkAdd414.Value = True Then mTCTVal = mTCTVal & "414,"
           If ChkAdd938.Visible = True And ChkAdd938.Value = True Then mTCTVal = mTCTVal & "938,"
           If ChkAdd939.Visible = True And ChkAdd939.Value = True Then mTCTVal = mTCTVal & "939,"
           If ChkAdd106.Visible = True And ChkAdd106.Value = True Then mTCTVal = mTCTVal & "106,"
           If ChkAdd228.Visible = True And ChkAdd228.Value = True Then mTCTVal = mTCTVal & "228,"
           If ChkAdd435.Visible = True And ChkAdd435.Value = True Then mTCTVal = mTCTVal & "435,"
           mTCTVal = mTCTVal & "|"
           '命名追蹤流水號Tracking No
           mTCTVal = mTCTVal & Trim(txtTCN01.Text) & "|"
           '工程師組別(隱藏欄位): 預設B退程序
           mTCTVal = mTCTVal & IIf(Trim(txtData(2)) <> "", txtData(2), "B") & "|"
           '譯畢期限
           mTCTVal = mTCTVal & Trim(txtData(0)) & Trim(txtData(1))
      'Added by Lydia 2023/03/09 輸入追蹤流水號,不走命名作業
      ElseIf txtTCN01.Visible = True And txtTCN01.Enabled = True Then
            mTCTVal = Trim(txtTCN01.Text)
      'end 2023/03/09
      End If
   End If
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      '2011/4/22 MODIFY BY SONIA 分所智權人員則多一天
      'txtPatent(28) = PUB_GetWorkDayAfterSysDate(CDbl(txtPatent(0)) + 19110000, 5)
      'Modified by Lydia 2019/09/16
      'If PUB_GetST06(txtPatent(15)) <> "1" Then
      If m_SalesST06 <> "1" Then
         txtPatent(28) = PUB_GetWorkDayAfterSysDate(CDbl(txtPatent(0)) + 19110000, 6)
      Else
         txtPatent(28) = PUB_GetWorkDayAfterSysDate(CDbl(txtPatent(0)) + 19110000, 5)
      End If
      '2011/4/22 END
      txtPatent(28).Locked = True
   Else
      txtPatent(28).Locked = False
   End If
End Sub

Private Sub chkWebApp_Click()
   'Modified by Lydia 2022/08/25改成共用模組
   'OnUpdateFee 'Added by Lydia 2017/12/08 勾選後,重新計算費用和規費
   If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), txtPatent(1), txtPatent(15), _
        IIf(chkWebApp.Visible = True And chkWebApp.Value = 1, "Y", ""), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
          txtPatent(17) = m_NowCP16
          txtPatent(21) = m_NowCP17
          txtPatent(18) = m_NowCP18
   End If
   'end 2022/08/25
End Sub

Private Sub cmdInventor_Click()
Dim i As Integer, strPetition As String, varInventorNo As Variant
   
   For i = 0 To 3
      If txtPatent(i + 8) <> "" Then
          strPetition = strPetition + txtPatent(i + 8) + ","
      End If
   Next
   If Right(strPetition, 1) = "," Then strPetition = Left(strPetition, Len(strPetition) - 1)
   'Modify By Sindy 2010/10/26
   'strPetition = strPetition + txtPatent(i + 8)
   If Trim(txtPatent(i + 8)) <> "" Then
      strPetition = strPetition + "," + txtPatent(i + 8)
   End If
   '2010/10/26 End
   ModifyInventor strPetition, strInventorNo
   If strInventorNo <> "" Then
      varInventorNo = Split(strInventorNo, ",")
   End If

   txtPatent(13).SetFocus
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim varSaveCursor, strAuto1 As String, strAuto2 As String, i As Integer
Dim strTemp As String
'Add By Sindy 2009/07/06
Dim strYear As String '抓下次繳費年度
Dim m_Nexttimes As String '抓下次繳費次數
Dim Cancel As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strFee
'Added by Lydia 2017/12/05
Dim strPath1 As String, strPath2 As String
Dim mBillNo As String, mMemo As String 'Added by Lydia 2019/05/13
Dim strDCase(1 To 4) As String 'Added by Lydia 2022/07/04
Dim bolSaveOK As Boolean, mRetVal As String  'Added by Lydia 2022/09/05

On Error GoTo ErrorHandler
   
'Modify By Sindy 2022/7/11
'註:這函數要離開時,原使用 Exit Sub, 全改為 GoTo ErrorHandler
   'Modified by Lydia 2019/09/16
   'm_SalesST15 = GetST15(txtPatent(15).Text) 'Added by Lydia 2018/09/06
   m_SalesST15 = GetST15(txtPatent(15).Text, , , m_SalesST06)
   
   If Index = 0 Then
      'Add By Sindy 2022/7/11 信件沖銷多案收文
      If m_strIR01 <> "" And m_bMRecvBatch = True Then
         '加入秀訊息
         If PUB_CheckFormExist("frmpic002") = False Then
            Load frmpic002
            frmpic002.Label1.Caption = "自動收文中...請稍候..."
            frmpic002.Show
         End If
         frmpic002.ZOrder 0
      End If
      '2022/7/11 END
      
      'Add by Amy 2021/12/16檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me, , True) = False Then
         GoTo ErrorHandler
      End If
  
      'Added by Lydia 2017/07/31 預設和檢查-所有內部收文, 若有輸入本所期限或法定期限者
      'Modified by Lyddia 2023/11/08 傳入必需欄位
      'If PUB_CheckCP0607(0, txtPatent(14), txtPatent(19)) = False Then GoTo ErrorHandler
      If PUB_CheckCP0607(0, txtPatent(14), txtPatent(19), IIf(frm010001.intModifyKind = 0, "Y", ""), txtPatent(4), txtSystem, txtPatent(1)) = False Then GoTo ErrorHandler
      
      'Added by Lydia 2018/10/12 檢查FMP輸入是否正確
      'If frm010001.intModifyKind = 0 And frm010001.txtFMP.Text <> "Y" And txtSystem = "P" And Left(m_SalesST15, 2) = "F2" And frm010001.mRole = "" And InStr(AddTrackingNo, txtPatent(1)) > 0 Then
      'Modified by Lydia 2021/06/16 排除國內接洽單自動轉收文
      If frm010001.intModifyKind = 0 And txtSystem = "P" And frm010001.mRole = "" And InStr(AddTrackingNo, txtPatent(1)) > 0 Then
          If frm010001.txtFMP.Text <> "Y" And Left(m_SalesST15, 2) = "F2" Then
              MsgBox "國外部收 FMP 案，必須在前一畫面輸入FMP案=Y ", vbCritical
              GoTo ErrorHandler
          End If
          If frm010001.txtFMP.Text = "Y" And Left(m_SalesST15, 2) <> "F2" Then
              MsgBox "內專收 P 案，前一畫面不可輸入為FMP案", vbCritical
              GoTo ErrorHandler
          End If
      End If
      'end 2018/10/12

      'Added by Lydia 2020/05/20 法律所案源收文：5/28台灣案之B1、B2及C收文時，增加"案源單號"欄位，B1、B2一定要輸入，C若未輸入則提醒'請確認接洽單沒有案源單號？'，案源單號更新至該筆收文的CP162。
      'Mark by Lydia 2020/06/10 重整判斷,以案源單的案源類型為準; 保留舊程式
'      If frm010001.intModifyKind = 0 And Trim(frm010001.txtFMP) <> "Y" And strSrvDate(1) >= 法律所案源收文啟用日 And (txtSystem = "FCP" Or txtSystem = "P") Then
'           t_LOSkind = PUB_GetLOSkind(txtSystem, txtPatent(1), txtPatent(4))
'           If Left(t_LOSkind, 1) = "B" And m_LOS15 = "" Then
'               MsgBox "請先回前畫面輸入案源單號！", vbCritical, "檢核案源單號"
'               GoTo ErrorHandler
'           End If
'           'Added by Lydia 2020/06/04 法律所案源收文：判斷是否為補收文=>案源類別
'           strExc(1) = PUB_GetLOSplus(txtSystem, txtCode(0), txtCode(1), txtCode(2), txtPatent(1), txtPatent(4), t_LOSkind)
'           If m_LOS02 = "" And Left(strExc(1), 1) = "B" Then
'               If MsgBox("請確認接洽單沒有案源單號？" & vbCrLf & "選擇""是""會繼續作業，選擇""否""回前畫面輸入案源單號", vbInformation + vbYesNo + vbDefaultButton1, "檢核案源單號") = vbNo Then
'                   GoTo ErrorHandler
'               End If
'           End If
'           'end 2020/06/04
'           'Modified by Lydia 2020/06/04
'           'If Left(t_LOSkind, 1) = "C" And m_LOS15 = "" Then
'           If ((Left(t_LOSkind, 1) = "C" And txtCode(0) = "") Or (Left(strExc(1), 1) = "C" And txtCode(0) <> "")) And m_LOS15 = "" Then
'               If MsgBox("請確認接洽單沒有案源單號？" & vbCrLf & "選擇""是""會繼續作業，選擇""否""回前畫面輸入案源單號", vbInformation + vbYesNo + vbDefaultButton1, "檢核案源單號") = vbNo Then
'                   GoTo ErrorHandler
'               End If
'           End If
'      End If
'      'end 2020/05/20
      
      'Add By Sindy 2024/9/4 CF案申請國家不可為台灣
      If Left(txtSystem, 2) = "CF" And txtPatent(4) = "000" Then
         MsgBox "CF案申請國家不可為台灣！", vbExclamation
         txtPatent(4).SetFocus
         Call txtPatent_GotFocus(4)
         Exit Sub
      End If
      '2024/9/4 END
      
      If frm010001.intModifyKind = 0 And strSrvDate(1) >= 法律所案源收文啟用日 And (txtSystem = "FCP" Or txtSystem = "P") Then
           If txtPatent(4) <> "000" Then '非台灣案=>清空資料
               m_LOS02 = ""
               m_LOS15 = ""
           Else
                t_LOSkind = PUB_GetLOSkind(txtSystem, txtPatent(1), txtPatent(4))
                If Left(t_LOSkind, 1) = "B" And m_LOS15 = "" Then
                    MsgBox "請先回前畫面輸入案源單號！", vbCritical, "檢核案源單號"
                    GoTo ErrorHandler
                End If
                '判斷是否為補收文=>案源類別
                strExc(1) = PUB_GetLOSplus(txtSystem, txtCode(0), txtCode(1), txtCode(2), txtPatent(1), txtPatent(4), IIf(t_LOSkind = "", "C", t_LOSkind))
                If m_LOS02 = "" And strExc(1) <> "" And m_LOS15 = "" Then
                    'Modified by Lydia 2020/07/20 預設"否"要輸入案源單號 vbDefaultButton2 (原本預設vbDefaultButton1)
                    If MsgBox("請確認接洽單左上角是否沒有案源單號？" & vbCrLf & "選擇""是""會繼續作業，選擇""否""回前畫面輸入案源單號", vbInformation + vbYesNo + vbDefaultButton2, "檢核案源單號") = vbNo Then
                        GoTo ErrorHandler
                    End If
                End If
           End If
      End If
      'end 2020/06/10
      
      'Added by Lydia 2021/09/10 修正畫面所有含跳行符號的文字框; 9/10 FCT-47909收文申請,彼所案號中間有換行
      PUB_FilterFormText Me
      
      bolMoveCheck = False 'Added by Lydia 2020/04/29

      varSaveCursor = Screen.MousePointer
      Screen.MousePointer = vbHourglass
      For i = 0 To 24
         'Add By Cheng 2001/12/12
         If i = 14 Then
            If txtPatent(14) <> "" Then
               If CheckIsTaiwanDate(txtPatent(14).Text) Then
                  If CheckReKey(txtPatent(14)) Then
                     If Val(txtPatent(14)) = Val(GetTaiwanTodayDate) Then
                        ShowMsg "此案件已屆本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
                     End If
                     If Val(txtPatent(14)) < Val(GetTaiwanTodayDate) Then
                        ShowMsg "此案件已逾本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
                     End If
                  Else
                     Screen.MousePointer = varSaveCursor
                     GoTo ErrorHandler
                  End If
               End If
            End If
         End If
         
         '2005/5/24 ADD BY SONIA
         If i = 3 Then
            If txtPatent(3) = "2" And txtPatent(1) = "107" And txtPatent(4) = 台灣國家代號 Then
               MsgBox "台灣新型案不可收 再審 程序 !!!", vbExclamation + vbOKOnly
               Screen.MousePointer = varSaveCursor
               'txtPatent(i).SetFocus
               txtPatent_GotFocus (i)
               GoTo ErrorHandler
            End If
            '2005/6/28 ADD BY SONIA
            'Modify by Morgan 2007/8/29 加807
            'If txtPatent(3) <> "2" And txtPatent(1) = "421" And txtPatent(4) = 台灣國家代號 Then
            If txtPatent(3) <> "2" And (txtPatent(1) = "421" Or txtPatent(1) = "807") And txtPatent(4) = 台灣國家代號 Then
               MsgBox "台灣案只有新型才可收【" & lblCaseProperty & "】程序 !!!", vbExclamation + vbOKOnly
               Screen.MousePointer = varSaveCursor
               txtPatent_GotFocus (i)
               GoTo ErrorHandler
            End If
            '2005/6/28 END
         End If
         '2005/5/24 END
         
         'Add By Sindy 2010/12/31 費用檢查提到存檔前檢查
         If i = 17 Then
            '郭 請作單 X14843050 不管
            'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
            'modify by sonia 2014/9/11 取消X69514,已轉外專
            'Modified by Lydia 2020/01/13 X38120030寧遠縣碩寧電子有限公司,其大陸年費605不要控制金額
            'If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" And _
            '   Mid(txtPatent(8), 1, 8) <> "X3928904" And Mid(txtPatent(9), 1, 8) <> "X3928904" And Mid(txtPatent(10), 1, 8) <> "X3928904" And Mid(txtPatent(11), 1, 8) <> "X3928904" And Mid(txtPatent(12), 1, 8) <> "X3928904" Then
            'Modified by Lydia 2022/08/19 改共用模組
            'If CheckExcept = False Then
            'end 2020/01/13
            If PUB_CheckExceptFrm010005(txtSystem, txtPatent(4), txtPatent(1), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10))) = False Then
               If ClsPDGetCaseLowPrice(txtSystem, txtPatent(4), txtPatent(1), douStPrice, douLowPrice) = 1 Then
               End If
               If txtPatent(17) <> "" Then
                  If ClsPDGetStaffArea(txtPatent(15), strTemp) Then
                  End If
                  'Added by Lydia 2020/05/20 法律所案源收文：台灣案之B1及B2案件性質都不可收費用。
                  'Modified by Lydia 2020/07/03 不確定國外部是否收費, 先排除
                  'If txtPatent(4) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" And Val(txtPatent(17)) > 0 Then
                  'Modified by Lydia 2021/11/09 改判斷;ex.FCP-093322案源收文503行政訴訟的金額=0
                  'If txtSystem <> "FCP" And txtPatent(4) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" And Val(txtPatent(17)) > 0 Then
                  If txtSystem <> "FCP" And txtPatent(4) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
                      If Val(txtPatent(17)) > 0 Then
                  'end 2021/11/09
                          MsgBox "【B類】案源接洽單之費用、規費、點數都必須為 0", vbExclamation, "檢核案源單號"
                          Screen.MousePointer = varSaveCursor
                          GoTo ErrorHandler
                      End If 'Added by Lydia 2021/11/09
                  'Modified by Lydia 2021/11/09 改判斷
                  'End If
                  ''end 2020/05/20
                  'If strTemp = "F23" And (txtSystem = "P" Or txtSystem = "CFP") Then
                  ElseIf strTemp = "F23" And (txtSystem = "P" Or txtSystem = "CFP") Then
                  'end 2021/11/09
                     '外專收文之P,CFP案不檢查費用
                  Else
                     'MODIFY BY SONIA 2014/7/17 +傳規費 CFP-027024
                     If ClsPDGetCaseFee(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(17)), Val(txtPatent(21))) = 0 Then
                        Screen.MousePointer = varSaveCursor
                        GoTo ErrorHandler
                     End If
                  End If
               End If
            Else
                If ClsPDGetCaseLowPrice(txtSystem, txtPatent(4), txtPatent(1), douStPrice, douLowPrice) = 1 Then
                End If
            End If
         End If
         'Add By Cheng 2001/12/12
         If i = 18 Then
            'Add By Sindy 2010/12/31 點數檢查提到存檔前檢查
            '郭 請作單 X14843050 不管
            'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
            'modify by sonia 2014/9/11 取消X69514,已轉外專
            'Modified by Lydia 2020/01/13 X38120030寧遠縣碩寧電子有限公司,其大陸年費605不要控制金額
            'If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" And _
            '   Mid(txtPatent(8), 1, 8) <> "X3928904" And Mid(txtPatent(9), 1, 8) <> "X3928904" And Mid(txtPatent(10), 1, 8) <> "X3928904" And Mid(txtPatent(11), 1, 8) <> "X3928904" And Mid(txtPatent(12), 1, 8) <> "X3928904" Then
            'Modified by Lydia 2022/08/19 改共用模組
            'If CheckExcept = False Then
            'end 2020/01/13
            If PUB_CheckExceptFrm010005(txtSystem, txtPatent(4), txtPatent(1), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10))) = False Then
                If txtPatent(18) = "" Then
                  If txtPatent(17) <> "" Or txtPatent(21) <> "" Then
                     ShowMsg MsgText(1035)
                     Screen.MousePointer = varSaveCursor
                     GoTo ErrorHandler
                  End If
               'Modified by Lydia 2020/04/27
               'ElseIf txtPatent(17) <> "" Or txtPatent(21) <> "" Then
               'Else
               Else
               'end 2020/04/27
                  'Added by Lydia 2020/05/20 法律所案源收文：台灣案之B1及B2案件性質都不可收費用。
                  'Modified by Lydia 2020/07/03 不確定國外部是否收費, 先排除
                  'If txtPatent(4) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" And Val(txtPatent(18)) > 0 Then
                  If txtSystem <> "FCP" And txtPatent(4) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" And Val(txtPatent(18)) > 0 Then
                      MsgBox "【B類】案源接洽單之費用、規費、點數都必須為 0", vbExclamation, "檢核案源單號"
                      Screen.MousePointer = varSaveCursor
                      GoTo ErrorHandler
                  End If
                  'end 2020/05/20
                  
                  If txtPatent(17) <> "" Or txtPatent(21) <> "" Then
                  Else
                        ShowMsg MsgText(1037)
                        Screen.MousePointer = varSaveCursor
                        GoTo ErrorHandler
                  End If 'Added by Lydia 2020/04/27
               End If
            End If
            'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
            'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
            'modify by sonia 2014/9/11 取消X69514,已轉外專
            'Modified by Lydia 2020/01/13 X38120030寧遠縣碩寧電子有限公司,其大陸年費605不要控制金額
            'If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" And _
            '   Mid(txtPatent(8), 1, 8) <> "X3928904" And Mid(txtPatent(9), 1, 8) <> "X3928904" And Mid(txtPatent(10), 1, 8) <> "X3928904" And Mid(txtPatent(11), 1, 8) <> "X3928904" And Mid(txtPatent(12), 1, 8) <> "X3928904" Then
            'Modified by Lydia 2022/08/19 改共用模組
            'If CheckExcept = False Then
            'end 2020/01/13
            If PUB_CheckExceptFrm010005(txtSystem, txtPatent(4), txtPatent(1), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10))) = False Then
               If txtPatent(17) <> "" Or txtPatent(21) <> "" Then
                  If Format((Val(txtPatent(17)) - Val(txtPatent(21))) / 1000, "0.0") <> Format(Val(txtPatent(18)), "0.0") Then
                     ShowMsg MsgText(1036)
                     Screen.MousePointer = varSaveCursor
'                     txtPatent(i).SetFocus
'                     txtPatent_GotFocus (i)
                     GoTo ErrorHandler
                  End If
               End If
            End If
            'Added by Lydia 2016/05/27 CFP控制年費.延展費及維持費智權同仁可加的點數
            'Modified by Lydia 2020/11/19 CFP英國脫歐案管制：因為定稿是否存在所以排除英國脫歐案，於後面另外判斷
            'If txtSystem = "CFP" And ((txtPatent(1) = "605" And Val(txtPatent(i)) > CFP_dg605) Or (txtPatent(1) = "606" And Val(txtPatent(i)) > CFP_dg606) Or (txtPatent(1) = "607" And Val(txtPatent(i)) > CFP_dg607)) Then
            If txtSystem = "CFP" And m_CaseNa239(1) = "" And m_CaseNa239(2) = "" And ((txtPatent(1) = "605" And Val(txtPatent(i)) > CFP_dg605) Or (txtPatent(1) = "606" And Val(txtPatent(i)) > CFP_dg606) Or (txtPatent(1) = "607" And Val(txtPatent(i)) > CFP_dg607)) Then
                'Modified by Lydia 2018/06/04 +lcv06
                'If PUB_GetLCV04LCV10(txtSystem, Trim(txtCode(0)), Trim(txtCode(1)), Trim(txtCode(2)), txtPatent(1), strExc(2), strExc(3), strExc(4), strExc(5)) Then
                If PUB_GetLCV04LCV10(txtSystem, Trim(txtCode(0)), Trim(txtCode(1)), Trim(txtCode(2)), txtPatent(1), strExc(2), strExc(3), strExc(4), strExc(5), strExc(6)) Then
                     Select Case txtPatent(1)
                         Case "605": strExc(1) = lblCaseProperty & "(605)超過" & CFP_dg605 & "點"
                         Case "606": strExc(1) = lblCaseProperty & "(606)超過" & CFP_dg606 & "點"
                         Case "607": strExc(1) = lblCaseProperty & "(607)超過" & CFP_dg607 & "點"
                     End Select
                   '不需主管簽核或主管不同意調整
                   If strExc(4) <> "Y" Or strExc(3) = "0" Then
                       'Modified by Lydia 2018/06/04  (ex.CFP-23599已進入國家階段的年費,各國調整不超過5點不需主管簽核,改用智權人員報價來比對)
                       'If Val(txtPatent(i)) > Val(strExc(2)) Then
                       If Val(txtPatent(i)) > IIf(strExc(4) = "" And Val(strExc(6)) > 0, Val(strExc(6)), Val(strExc(2))) Then
                          'Modified by Lydia 2018/09/20 預設按鈕改成"否"
                          If MsgBox(strExc(1) & "，請確認主管是否已簽核？", vbCritical + vbYesNo + vbDefaultButton2, "控制點數") = vbNo Then
                             Screen.MousePointer = varSaveCursor
                             GoTo ErrorHandler
                          End If
                       End If
                   '有主管簽核
                   ElseIf strExc(4) = "Y" Then
                       If Val(txtPatent(i)) > Val(strExc(3)) Then '輸入點數>主管簽核點數
                          'Modified by Lydia 2018/09/20 預設按鈕改成"否"
                          If MsgBox(strExc(1) & "，請確認主管是否已簽核？", vbCritical + vbYesNo + vbDefaultButton2, "控制點數") = vbNo Then
                             Screen.MousePointer = varSaveCursor
                             GoTo ErrorHandler
                          End If
                       End If
                   End If
                End If
            End If
         End If
         'Added by Lydia 2020/11/19 CFP英國脫歐案管制：另外判斷同仁可加的點數
         If i = 18 And txtSystem = "CFP" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" And (txtPatent(1) = "607" And Val(txtPatent(i)) > CFP_dg607) Then
            If MsgBox(lblCaseProperty & "(607)超過" & CFP_dg607 & "點" & "，請確認主管是否已簽核？", vbCritical + vbYesNo + vbDefaultButton2, "控制點數") = vbNo Then
               Screen.MousePointer = varSaveCursor
               GoTo ErrorHandler
            End If
         End If
         'end 2020/11/19
         
         'Add By Sindy 2010/12/31 規費檢查提到存檔前檢查
         If i = 21 Then
            '郭 請作單 X14843050 不管
            'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
            'modify by sonia 2014/9/11 取消X69514,已轉外專
            'Modified by Lydia 2020/01/13 X38120030寧遠縣碩寧電子有限公司,其大陸年費605不要控制金額
            'If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" And _
            '   Mid(txtPatent(8), 1, 8) <> "X3928904" And Mid(txtPatent(9), 1, 8) <> "X3928904" And Mid(txtPatent(10), 1, 8) <> "X3928904" And Mid(txtPatent(11), 1, 8) <> "X3928904" And Mid(txtPatent(12), 1, 8) <> "X3928904" Then
            'Modified by Lydia 2022/08/19 改共用模組
            'If CheckExcept = False Then
            'end 2020/01/13
            If PUB_CheckExceptFrm010005(txtSystem, txtPatent(4), txtPatent(1), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10))) = False Then
               ' 主張優先權無規費, 但申請優先權證明一定有規費
               If txtPatent(1) = 主張優先權 And Val(txtPatent(21)) <> 0 And txtPatent(4) = "000" And txtSystem = "P" Then
                  MsgBox "案件性質為主張優先權時, 一定沒有規費!!!", vbExclamation + vbOKOnly
                  Screen.MousePointer = varSaveCursor
                  GoTo ErrorHandler
               End If
               If txtPatent(1) = 申請優先權證明 And Val(txtPatent(21)) = 0 And txtPatent(4) = "000" And txtSystem = "P" Then
                  MsgBox "案件性質為申請優先權證明時, 一定有規費!!!", vbExclamation + vbOKOnly
                  Screen.MousePointer = varSaveCursor
                  GoTo ErrorHandler
               End If
               ' P之讓與及專利權讓與規費不同
               If txtPatent(1) = 讓與 And txtPatent(21) <> "2000" And txtPatent(4) = "000" And txtSystem = "P" Then
                  MsgBox "台灣讓與案規費必須為2000 !!!", vbExclamation + vbOKOnly
                  Screen.MousePointer = varSaveCursor
                  GoTo ErrorHandler
               End If
               If txtPatent(1) = 專利權讓與 And txtPatent(21) <> "2000" And txtPatent(4) = "000" And txtSystem = "P" Then
                  MsgBox "台灣專利權讓與案規費必須為2000 !!!", vbExclamation + vbOKOnly
                  Screen.MousePointer = varSaveCursor
                  GoTo ErrorHandler
               End If
               ' 閱卷無規費, 但請求閱卷一定有規費
               If txtPatent(1) = 閱卷 And Val(txtPatent(21)) <> 0 And txtPatent(4) = "000" And txtSystem = "P" Then
                  MsgBox "案件性質為閱卷時, 一定沒有規費!!!", vbExclamation + vbOKOnly
                  Screen.MousePointer = varSaveCursor
                  GoTo ErrorHandler
               End If
            End If
            
            If Val(txtPatent(21)) > 0 Or txtPatent(4) = "000" Then     'ADD by sonia 2014/7/17 加入未輸規費時不檢查此段,因為可能是依代理人帳單請款CFP-027024
               '郭 請作單 X14843050 不管
               'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
               'modify by sonia 2014/9/11 取消X69514,已轉外專
               'Modified by Lydia 2020/01/13 X38120030寧遠縣碩寧電子有限公司,其大陸年費605不要控制金額
               'If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" And _
               '   Mid(txtPatent(8), 1, 8) <> "X3928904" And Mid(txtPatent(9), 1, 8) <> "X3928904" And Mid(txtPatent(10), 1, 8) <> "X3928904" And Mid(txtPatent(11), 1, 8) <> "X3928904" And Mid(txtPatent(12), 1, 8) <> "X3928904" Then
               'Modified by Lydia 2022/08/19 改共用模組
               'If CheckExcept = False Then
               'end 2020/01/13
               If PUB_CheckExceptFrm010005(txtSystem, txtPatent(4), txtPatent(1), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10))) = False Then
                  'Modified by Morgan 2015/1/19 程式重整
                  If ClsPDGetStaffArea(txtPatent(15), strTemp) Then
                  End If
                  'Added by Lydia 2020/05/20 法律所案源收文：台灣案之B1及B2案件性質都不可收費用。
                  'Modified by Lydia 2020/07/03 不確定國外部是否收費, 先排除
                  'If txtPatent(4) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" And Val(txtPatent(21)) > 0 Then
                  If txtSystem <> "FCP" And txtPatent(4) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" And Val(txtPatent(21)) > 0 Then
                      MsgBox "【B類】案源接洽單之費用、規費、點數都必須為 0", vbExclamation, "檢核案源單號"
                      Screen.MousePointer = varSaveCursor
                      GoTo ErrorHandler
                  End If
                  'end 2020/05/20
                  
                  If strTemp = "F23" And (txtSystem = "P" Or txtSystem = "CFP") Then
                     '外專收文之P,CFP案不檢查規費
                  'Added by Lydia 2020/05/20 法律所案源收文：B類不檢查規費
                  ElseIf txtPatent(4) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
                  'end 2020/05/20
                  Else
                  
                     If txtPatent(4) = "000" Then
                        '台灣發明實審調整規費
                        If txtCode(1) = "" Then txtCode(1).Text = "0"
                        If txtCode(2) = "" Then txtCode(2).Text = "00"
                        'Memo by Lydia 2018/02/07 不傳CP118,電子送件後面才算
                        strFee = GetPatentOfficialFee(txtSystem, txtPatent(1), txtPatent(19), txtPatent(3), txtPatent(4), m_PA16, m_PA14, txtCode(0), txtCode(1), txtCode(2))
                     End If
                     
                     '台灣電子送件,有英文摘要減免
                     If txtPatent(4) = "000" And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "125" Or Left(txtPatent(1), 1) = "3") And (chkEnglish.Value = 1 Or chkWebApp.Value = 1) Then
                        If Val(strFee) = 0 Then
                           strExc(0) = "SELECT cf08 FROM CASEFEE WHERE CF01 = '" & txtSystem & "' AND CF02 = '" & txtPatent(4) & "' AND CF03 = '" & txtPatent(1) & "' and cf08>0"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              strFee = RsTemp(0)
                           End If
                        End If
                        
                        If Val(strFee) > 0 Then
                           strExc(2) = lblNation & lblCaseProperty
                           If txtPatent(1) = "101" Or txtPatent(1) = "307" Then
                              'Modified by Lydia 2018/02/07 有顯示才算
                              'If chkEnglish.Value = 1 Then
                              If chkEnglish.Value = 1 And chkEnglish.Visible = True Then
                                 strFee = Val(strFee) - 800
                                 strExc(2) = strExc(2) & "若有英文摘要"
                              End If
                           End If
                           If chkWebApp.Value = 1 Then
                              strFee = Val(strFee) - 600
                               'Modified by Lydia 2018/02/07 有顯示才算
                              'If chkEnglish.Value = 1 Then
                              If chkEnglish.Value = 1 And chkEnglish.Visible = True Then
                                 strExc(2) = strExc(2) & "且為電子送件"
                              Else
                                 strExc(2) = strExc(2) & "若為電子送件"
                              End If
                           End If
                        End If
                     
                     '優先權份數
                     ElseIf txtCopy.Visible Then
                        strExc(0) = "SELECT cf08 FROM CASEFEE WHERE CF01 = '" & txtSystem & "' AND CF02 = '" & txtPatent(4) & "' AND CF03 = '" & txtPatent(1) & "' and cf08>0"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           strFee = RsTemp(0)
                        End If
                        strFee = Val(txtCopy) * Val(strFee)
                        
                     End If
                     
                     If Val(strFee) > 0 Then
                        If Val(txtPatent(21)) <> Val(strFee) Then
                           strTit = "檢核資料"
                           strMsg = "規費數值應為<" & strFee & ">"
                           nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
                           Screen.MousePointer = varSaveCursor
                           GoTo ErrorHandler
                        End If
                     'Modified by Lydia 2018/02/05 有顯示才算
                     'ElseIf GetCaseMoney(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(21)), False, chkEnglish.Value) = 0 Then
                     'Modified by Lydia 2023/02/10 傳入專利種類
                     'ElseIf GetCaseMoney(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(21)), False, chkEnglish.Value And chkEnglish.Visible) = 0 Then
                     ElseIf GetCaseMoney(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(21)), False, chkEnglish.Value And chkEnglish.Visible, IIf(txtSystem = "P", txtPatent(3), "")) = 0 Then
                        Screen.MousePointer = varSaveCursor
                        GoTo ErrorHandler
                     End If
                     
                  End If
                  'end 2015/1/19
               End If
            End If 'ADD by sonia 2014/7/17 加入未輸規費時不檢查上面這段,因為可能是依代理人帳單請款CFP-027024
         End If
         
         'Add by Sindy 2015/4/1 收”延期”時一定要輸本所期限及法定期限
         If (txtSystem = "P" Or txtSystem = "CFP") And txtPatent(1).Text = "404" Then
            If Val(txtPatent(14).Text) = 0 Or Val(txtPatent(19).Text) = 0 Then
               MsgBox "延期一定要有期限，請退回智權人員補填期限！", vbExclamation + vbOKOnly
               Screen.MousePointer = varSaveCursor
               GoTo ErrorHandler
            End If
         End If
         '2015/4/1 END
         
         'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
         'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
         'modify by sonia 2014/9/11 取消X69514,已轉外專
         'Modified by Lydia 2020/01/13 X38120030寧遠縣碩寧電子有限公司,其大陸年費605不要控制金額
         'If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" And _
         '   Mid(txtPatent(8), 1, 8) <> "X3928904" And Mid(txtPatent(9), 1, 8) <> "X3928904" And Mid(txtPatent(10), 1, 8) <> "X3928904" And Mid(txtPatent(11), 1, 8) <> "X3928904" And Mid(txtPatent(12), 1, 8) <> "X3928904" Then
         'Modified by Lydia 2022/08/19 改共用模組
         'If CheckExcept = False Then
         'end 2020/01/13
         If PUB_CheckExceptFrm010005(txtSystem, txtPatent(4), txtPatent(1), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10))) = False Then
           If Val(txtPatent(17)) = 0 And Val(txtPatent(18)) = 0 And Val(txtPatent(21)) = 0 Then
              If i = 17 Or i = 18 Or i = 21 Then
                 GoTo GoToNext
              End If
           End If
         End If
         'Add By Sindy 2010/5/26 檢查申請人及讓與申請人的輸入順序
         If (Trim(txtPatent(11)) <> "" And Trim(txtPatent(8)) = "") Or _
            (Trim(txtPatent(9)) <> "" And Trim(txtPatent(11)) = "") Or _
            (Trim(txtPatent(12)) <> "" And Trim(txtPatent(9)) = "") Or _
            (Trim(txtPatent(10)) <> "" And Trim(txtPatent(12)) = "") Then
            ShowMsg "請依序輸入申請人!"
            If Trim(txtPatent(10)) <> "" Then txtPatent(10).SetFocus: Call txtPatent_GotFocus(10): Exit For
            If Trim(txtPatent(12)) <> "" Then txtPatent(12).SetFocus: Call txtPatent_GotFocus(12): Exit For
            If Trim(txtPatent(9)) <> "" Then txtPatent(9).SetFocus: Call txtPatent_GotFocus(9): Exit For
            If Trim(txtPatent(11)) <> "" Then txtPatent(11).SetFocus: Call txtPatent_GotFocus(11): Exit For
         End If
         If (Trim(txtPetitionx(2)) <> "" And Trim(txtPatent(23)) = "") Or _
            (Trim(txtPetitionx(3)) <> "" And Trim(txtPetitionx(2)) = "") Or _
            (Trim(txtPetitionx(4)) <> "" And Trim(txtPetitionx(3)) = "") Or _
            (Trim(txtPetitionx(5)) <> "" And Trim(txtPetitionx(4)) = "") Then
            ShowMsg "請依序輸入讓與申請人!"
            If Trim(txtPetitionx(5)) <> "" Then txtPetitionx(5).SetFocus: Call txtPetitionx_GotFocus(5): Exit For
            If Trim(txtPetitionx(4)) <> "" Then txtPetitionx(4).SetFocus: Call txtPetitionx_GotFocus(4): Exit For
            If Trim(txtPetitionx(3)) <> "" Then txtPetitionx(3).SetFocus: Call txtPetitionx_GotFocus(3): Exit For
            If Trim(txtPetitionx(2)) <> "" Then txtPetitionx(2).SetFocus: Call txtPetitionx_GotFocus(2): Exit For
         End If
         '2010/5/26 End
         'Modify By Cheng 2001/12/27
         If i = 8 Then
            If txtPatent(8) = "" And txtPatent(13) = "" Then
               ShowMsg "申請人或代理人不可同時空白!"
               txtPatent(8).SetFocus
               txtPatent_GotFocus (8)
               Exit For
           End If
         End If
         'modify by sonia 2017/1/23 +第二~五申請人
         If i = 8 Or i = 9 Or i = 10 Or i = 11 Or i = 12 Then
            If Len(Trim(Me.txtPatent(i).Text)) > 0 Then
               If CheckKeyIn(i) <> 1 Then
                  If txtPatent(i).Enabled = True Then
                     txtPatent(i).SetFocus
                     txtPatent_GotFocus (i)
                  End If
                  Exit For
               End If
            End If
               'add by nick 2004/12/08  加條件，如果是 x13175040 則他的客戶案件案號一定要輸
               If UCase(txtPatent(8)) = "X1317504" Then
                   If Trim(txtPatent(25)) = "" Then
                       ShowMsg "申請人為 X1317504 ，則他的客戶案件案號不可空白!"
                       txtPatent(25).SetFocus
                       txtPatent_GotFocus (25)
                       Exit For
                   End If
               End If
         ElseIf i = 13 Then
            If Len(Trim(Me.txtPatent(i).Text)) > 0 Then
               If CheckKeyIn(i) <> 1 Then
                  If txtPatent(i).Enabled = True Then
                     txtPatent(i).SetFocus
                     txtPatent_GotFocus (i)
                  End If
                  Exit For
               End If
            End If
         ElseIf txtPatent(i).Enabled And txtPatent(i).Visible Then
            If CheckKeyIn(i) <> 1 Then
               If i = 3 Then
                  txtPatent(4).SetFocus
                  txtPatent(4).SetFocus
               Else
                  txtPatent(i).SetFocus
                  txtPatent_GotFocus (i)
               End If
               Exit For
            End If
         End If
         
GoToNext:
      Next
      If i = 25 Then
        'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
        'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
        'modify by sonia 2014/9/11 取消X69514,已轉外專
        'Modified by Lydia 2020/01/13 X38120030寧遠縣碩寧電子有限公司,其大陸年費605不要控制金額
        'If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" And _
        '   Mid(txtPatent(8), 1, 8) <> "X3928904" And Mid(txtPatent(9), 1, 8) <> "X3928904" And Mid(txtPatent(10), 1, 8) <> "X3928904" And Mid(txtPatent(11), 1, 8) <> "X3928904" And Mid(txtPatent(12), 1, 8) <> "X3928904" Then
        'Modified by Lydia 2022/08/19 改共用模組
        'If CheckExcept = False Then
        'end 2020/01/13
        If PUB_CheckExceptFrm010005(txtSystem, txtPatent(4), txtPatent(1), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10))) = False Then
               'Modify by Morgan 2004/6/23   加判斷欄位是否可修改
               If txtPatent(21).Enabled = True Then
                  '2007/11/5 add by sonia 外專收文之P,CFP案不檢查規費
                  If ClsPDGetStaffArea(txtPatent(15), strTemp) Then
                     If strTemp = "F23" And (txtSystem = "P" Or txtSystem = "CFP") Then
                        GoTo CheckTag1
                     End If
                  End If
                  '2007/11/5 END
                  
'Removed by Morgan 2012/6/20 上面迴圈內已做檢查此處無需再重複
'                  '檢查規費欄位
'                  'Add by Morgan 2008/8/5
'                  '台灣發明申請若有英文摘要時規費應為2700
'                  If txtPatent(4) = "000" And txtPatent(1) = "101" And chkEnglish.Value = 1 Then
'                     If Val(txtPatent(21)) <> 2700 Then
'                        MsgBox "台灣發明申請若有英文摘要時規費應為2700！"
'                        Screen.MousePointer = varSaveCursor
'                        txtPatent(21).SetFocus
'                        txtPatent_GotFocus (21)
'                        GoTo ErrorHandler
'                     End If
'                  Else
'                  'end 2008/8/5
'                     '2009/2/10 modify by sonia 是否為同時申請三國(含)以上之美日德可多5點
'                     'If GetCaseMoney(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(21))) <> 1 Then
'                     If GetCaseMoney(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(21)), False, chkEnglish.Value) <> 1 Then
'                     '2009/2/10 end
'                         Screen.MousePointer = varSaveCursor
'                         txtPatent(21).SetFocus
'                         txtPatent_GotFocus (21)
'                         GoTo ErrorHandler
'                     End If
'                  End If
'end 2012/6/20

               End If
        End If
        
        'Added by Lydia 2017/04/10 檢查客戶案件案號
        If CheckKeyIn(25) <> 1 Then
            txtPatent(25).SetFocus
            Call txtPatent_GotFocus(25)
            Screen.MousePointer = vbDefault
            GoTo ErrorHandler
        End If
        'end 2017/04/10
               
        'add by Toni   2008/8/26
         '若為新案 檢查P or CFP 輸入發明人
         'Modified by Morgan 2018/3/7 案件屬性不該只在新增控制,修改也要才對
         'If (Trim(txtSystem.Text) = "P" Or Trim(txtSystem.Text) = "CFP") And txtCode(0) = "" Then
         If (Trim(txtSystem.Text) = "P" Or Trim(txtSystem.Text) = "CFP") And (txtCode(0) = "" Or mCP31 = "Y") Then
            If strInventorNo = "" Then
               If MsgBox("是否輸入發明人資料?", vbExclamation + vbOKCancel) = vbOK Then
                  cmdInventor.SetFocus
                  Screen.MousePointer = vbDefault
                 GoTo ErrorHandler
               End If
            End If
            'Add By Sindy 2010/10/28 智權人員ST15<>F時且為發明或新型, 不可空白
            If Left(m_SalesST15, 1) <> "F" Then
               If (txtPatent(3) = "1" Or txtPatent(3) = "2") And Combo3.Text = "" Then
                  MsgBox "案件屬性不可空白！"
                  Screen.MousePointer = varSaveCursor
                  Combo3.SetFocus
                  GoTo ErrorHandler
               'Added by Morgan 2018/3/7 要和接洽單的控制一致
               ElseIf txtPatent(4) <> "000" And txtPatent(3) = "3" And Combo3.Text <> "" Then
                  MsgBox "非台灣設計案不可輸入案件屬性！", vbExclamation
                  Screen.MousePointer = varSaveCursor
                  Combo3.SetFocus
                  GoTo ErrorHandler
               'end 2018/3/7
               End If
            End If
         End If

         'Add By Sindy 2010/10/28
         If Combo3.Enabled = True And Combo3.Text <> "" Then
            Cancel = False
            Combo3_Validate Cancel
            If Cancel = True Then
               Screen.MousePointer = varSaveCursor
               GoTo ErrorHandler
            End If
         End If
         
CheckTag1:
       'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
       'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
       'modify by sonia 2014/9/11 取消X69514,已轉外專
        'Modified by Lydia 2020/01/13 X38120030寧遠縣碩寧電子有限公司,其大陸年費605不要控制金額
        'If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" And _
        '   Mid(txtPatent(8), 1, 8) <> "X3928904" And Mid(txtPatent(9), 1, 8) <> "X3928904" And Mid(txtPatent(10), 1, 8) <> "X3928904" And Mid(txtPatent(11), 1, 8) <> "X3928904" And Mid(txtPatent(12), 1, 8) <> "X3928904" Then
        'Modified by Lydia 2022/08/19 改共用模組
        'If CheckExcept = False Then
        'end 2020/01/13
        If PUB_CheckExceptFrm010005(txtSystem, txtPatent(4), txtPatent(1), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10))) = False Then
           'Modify by Morgan 2004/6/23   加判斷欄位是否可修改
           If txtPatent(18).Enabled = True Then
              '2007/11/5 add by sonia 外專收文之P,CFP案不檢查規費
              If ClsPDGetStaffArea(txtPatent(15), strTemp) Then
                 If strTemp = "F23" And (txtSystem = "P" Or txtSystem = "CFP") Then
                    GoTo CheckTag2
                 End If
              End If
              '2007/11/5 END
              
              'Added by Lydia 2020/05/20 法律所案源收文：B類不檢查規費
              If txtPatent(4) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
                  GoTo CheckTag2
              End If
              'end 2020/05/20
                  
             'Add By Cheng 2003/08/28
             '檢查點數是否低於底價
             If ChkPointValue(Me.txtSystem.Text, Me.txtPatent(4).Text, Me.txtPatent(1).Text, Me.txtPatent(18).Text, Me.txtPatent(15).Text) = False Then
                 Screen.MousePointer = varSaveCursor
                 txtPatent(18).SetFocus
                 txtPatent_GotFocus (18)
                 GoTo ErrorHandler
             End If
           End If
       End If
CheckTag2:
         
         'Added by Lydia 2023/04/25 外專後續案收文：FMP(非寰華案)領證和年費收文，若為舊案未閉卷，控管一定要輸入年度
         If m_bolFMP = True And m_bolFMP2 = False And txtCode(0) <> "" And InStr("601,605,", Trim(txtPatent(1))) > 0 Then
            If Val(modBase(58)) = 0 And textYear.Visible = True And Text1(0).Visible = True Then
                If Val(textYear) = 0 Or Val(Text1(0)) = 0 Then
                    MsgBox "請輸入繳費年度！", vbExclamation + vbOKOnly
                    If Val(textYear) = 0 Then
                        textYear.SetFocus
                    Else
                        Text1(0).SetFocus
                    End If
                    GoTo ErrorHandler
                End If
            End If
         End If
         'Add By Sindy 2009/07/06
         If txtCode(0) <> "" Then 'Modify by Sindy 2010/8/4 舊案才檢查
            If textYear.Visible = True And Text1(0).Visible = True Then
               ' 檢查繳費年度/次數是否不正確
               If IsEmptyText(textYear) = False And m_CP10 <> "601" Then
                  If txtCode(1) = "" Then txtCode(1).Text = "0"
                  If txtCode(2) = "" Then txtCode(2).Text = "00"
                  'Modified by Morgan 2022/6/15 +m_CP10
                  m_Nexttimes = PUB_Getnexttimes(txtSystem, txtCode(0), txtCode(1), txtCode(2), strYear, , m_CP10)
                  If m_Nexttimes <> "" Then
                     If m_CP10 = "601" Or m_CP10 = "605" Then '繳費年度
                        '2010/12/29 add by sonia P-097385
                        If m_Nexttimes = "1" And m_CP10 = "605" And txtSystem = "P" Then
                           MsgBox "此案件無繳費記錄, 請不要輸入繳費年度起迄資料！"
                           textYear.SetFocus
                           GoTo ErrorHandler
                        '2010/12/29 end
                        ElseIf Val(textYear) <> Val(strYear) Then
                           MsgBox "繳費(起)年度有誤，應為" & strYear & "！"
                           Screen.MousePointer = vbDefault
                           textYear.SetFocus
                           GoTo ErrorHandler
                        End If
                     Else '繳費次數
                        If Val(textYear) <> Val(m_Nexttimes) Then
                           MsgBox "繳費(起)次數有誤，應為" & m_Nexttimes & "！"
                           Screen.MousePointer = vbDefault
                           textYear.SetFocus
                           GoTo ErrorHandler
                        End If
                     End If
                  Else
                     If m_CP10 = "601" Or m_CP10 = "605" Then '繳費年度
                        MsgBox "無下次繳費年度！"
                     Else '繳費次數
                        MsgBox "無下次繳費次數！"
                     End If
                     Screen.MousePointer = vbDefault
                     If textYear.Enabled = False Then
                        Text1(0).SetFocus
                     Else
                        textYear.SetFocus
                     End If
                     GoTo ErrorHandler
                  End If
               End If
               If IsEmptyText(textYear) = False Or IsEmptyText(Text1(0)) = False Then
                  If textYear = "" Or textYear = "0" Then
                     If m_CP10 = "601" Or m_CP10 = "605" Then '繳費年度
                        MsgBox "無繳費(起)年度，請清空起迄年度！"
                     Else '繳費次數
                        MsgBox "無繳費(起)次數，請清空起迄次數！"
                     End If
                     Screen.MousePointer = vbDefault
                     If textYear = "0" Then
                        textYear.SetFocus
                     Else
                        Text1(0).SetFocus
                     End If
                     GoTo ErrorHandler
                  End If
                  If m_CP10 = "601" And Text1(0) = "" Then
                     '不跑else段程式
                  Else
                     If Text1(0) = "" Then Text1(0) = "0"
                     If Val(textYear) > Val(Text1(0)) Then
                        If m_CP10 = "601" Or m_CP10 = "605" Then '繳費年度
                           MsgBox "繳費(迄)年度不可小於(起)年度！"
                        Else '繳費次數
                           MsgBox "繳費(迄)次數不可小於(起)次數！"
                        End If
                        Screen.MousePointer = vbDefault
                        Text1(0).SetFocus
                        GoTo ErrorHandler
                     End If
                  End If
               End If
            End If
         End If
         '2009/07/06 End
         
         'Add By Sindy 2022/12/7 證書形式
         If txtPatent(29).Visible = True Then
            If Len(txtPatent(29)) = 0 And strSrvDate(1) >= "20230101" Then
               MsgBox "證書形式不可空白！", vbExclamation, "檢核資料"
               txtPatent(29).SetFocus
               Screen.MousePointer = varSaveCursor
               GoTo ErrorHandler
            End If
         End If
         
         'Add By Sindy 2011/01/05
         'Modified by Lydia 2018/09/06 改判斷
         'strSql = "select st15 from staff where st01='" & txtPatent(15) & "'"
         'intI = 1
         'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         'If intI = 1 Then
         '   If Not IsNull(RsTemp.Fields("st15")) Then
               '國外部收文台灣案必須收FCP案號
               'If Left(Trim(RsTemp.Fields("st15")), 1) = "F" And txtPatent(4) = "000" And txtSystem <> "FCP" Then
               If Left(m_SalesST15, 1) = "F" And txtPatent(4) = "000" And txtSystem <> "FCP" Then
                  '2015/4/14 MODIFY BY SONIA 林總指示開放投資法務人員可收文L案及P案
                  'MsgBox "國外部台灣案必須收 FCP 案號!!!", vbExclamation + vbOKOnly
                  'Screen.MousePointer = varSaveCursor
                  'GoTo ErrorHandler
                  If MsgBox("國外部台灣案必須收 FCP 案號, 是否修改系統類別？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                     Screen.MousePointer = varSaveCursor
                     GoTo ErrorHandler
                  End If
               End If
            'End If
         'End If
     
         strAuto1 = txtRecieveCode
         'Add By Cheng 2002/05/23txtPatent(17)
         '重新檢查欄位有效性
         If TxtValidate = False Then Screen.MousePointer = vbDefault: GoTo ErrorHandler
         
         'add by nickc 2007/11/12 加入檢查特殊客戶
         Dim IsSpecCu As Boolean
         IsSpecCu = False
         If fraPatition.Visible = True Then
               If IsSpecCu = False And txtPatent(23) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtPatent(23)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtPatent(23)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtPetitionx(2) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtPetitionx(2)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtPetitionx(2)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtPetitionx(3) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtPetitionx(3)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtPetitionx(3)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtPetitionx(4) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtPetitionx(4)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtPetitionx(4)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtPetitionx(5) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtPetitionx(5)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtPetitionx(5)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
         Else
               If IsSpecCu = False And txtPatent(8) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtPatent(8)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtPatent(8)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtPatent(11) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtPatent(11)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtPatent(11)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtPatent(9) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtPatent(9)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtPatent(9)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtPatent(12) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtPatent(12)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtPatent(12)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtPatent(10) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtPatent(10)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtPatent(10)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
         End If
         If IsSpecCu Then
               'Modify by Sindy 2010/8/19
               If MsgBox("特殊客戶，請確認此客戶接洽單主管是否核示??", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                   Screen.MousePointer = vbDefault 'Add by Sindy 2010/8/19
                   GoTo ErrorHandler
               End If
         End If
         
         'add by nickc 2007/03/27 非台灣要詢問
         '2009/11/25 MODIFY BY SONIA 新案才要詢問
         'If GetPrjNationNumber1(ChangeCustomerL(txtPatent(8))) > "010" Then
         '2010/10/20 modify by sonia 非智權部收文才要問 CFP-023621
         'If GetPrjNationNumber1(ChangeCustomerL(txtPatent(8))) > "010" And txtCode(0) = "" Then
         If GetPrjNationNumber1(ChangeCustomerL(txtPatent(8))) > "010" And txtCode(0) = "" And Left(Trim(m_SalesST15), 1) <> "S" Then
               If txtPatent(13) = "" Then
                   If MsgBox("請確定  無代理人   !!", vbYesNo, "警告！") = vbNo Then
                       Screen.MousePointer = varSaveCursor
                       txtPatent(13).SetFocus
                       txtPatent_GotFocus (13)
                       GoTo ErrorHandler
                   End If
               'Modify by Amy 2017/01/03 從下面搬上來,上面訊息若選擇"是",就不要再詢問下列訊息-秀玲
               ElseIf txtPatent(27) = "" Then
                   If MsgBox("請確定  無代理人彼所案號  !!", vbYesNo, "警告！") = vbNo Then
                       Screen.MousePointer = varSaveCursor
                       txtPatent(27).SetFocus
                       txtPatent_GotFocus (27)
                       GoTo ErrorHandler
                   End If
               End If
         End If
         
         'Add By Sindy 2010/3/8
         If Left(Trim(GetStaffDepartment(txtPatent(15).Text)), 2) = "F2" And _
            frm010001.intSaveMode = "1" And _
            strPA51s = "" And strPA52s = "" And strPA53s = "" And _
            strPA54s = "" And strPA55s = "" And strPA56s = "" Then
            If MsgBox("是否輸入國外聯絡人資料?", vbExclamation + vbOKCancel) = vbOK Then
               Screen.MousePointer = vbDefault
               Call Command1_Click
               GoTo ErrorHandler
            End If
         End If
         '2010/3/8 End
         
      '2011/4/21 add by sonia
Dim strPA149 As String, strContact As String

      If cboContact.Locked = False Then
         strContact = ""
         If cboContact.ListCount > 2 Then
            'Modify by Amy 2022/11/10 改成Form 2.0
            'strPA149 = Format(cboContact.ItemData(cboContact.ListIndex), "00")
            strPA149 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
            PUB_GetContact strAppNo1, strContact, True
            If strPA149 = strContact Or strPA149 = "00" Then
               If MsgBox("請確定接洽人欄是否有為★, 是否要選擇其他接洽人!!", vbYesNo, "警告！") = vbYes Then
                   Screen.MousePointer = varSaveCursor
                   cboContact.SetFocus
                   GoTo ErrorHandler
               End If
            End If
         End If
      End If
      '2011/4/21 end
      
      'Add By Sindy 2015/8/26
      Dim dblAmt As Double, dblPFee As Double
      'Modified by Lydia 2017/06/19 新案不檢查
      'If (txtSystem = "P" Or txtSystem = "CFP") And (txtPatent(1) = "605" Or txtPatent(1) = "606" Or txtPatent(1) = "607") And _
         GetBillData(txtPatent(8), dblAmt, dblPFee, , , txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2)) = True Then
      If (txtSystem = "P" Or txtSystem = "CFP") And (txtPatent(1) = "605" Or txtPatent(1) = "606" Or txtPatent(1) = "607") And Trim(txtCode(0)) <> "" Then
         'Modified by Lydia 2019/05/14
         'If GetBillData(txtPatent(8), dblAmt, dblPFee, , , txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2)) = True Then
         'Modified by Lydia 2022/06/13 傳入收文之本所案號,案件性質(可用,串接)
         'If PUB_GetBillDataAll("1", txtPatent(8), dblAmt, dblPFee, , , txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2)) = True Then
          'Modified by Lydia 2022/06/15 傳入收文之智權人員
         If PUB_GetBillDataAll("1", txtPatent(8), txtSystem & IIf(txtCode(0) <> "", txtCode(0) & Left(txtCode(1) & "0", 1) & Left(txtCode(2) & "00", 2), ""), txtPatent(1), Trim(txtPatent(15)), dblAmt, dblPFee, , , txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2)) = True Then
         'end 2017/06/19
            'Modify By Sindy 2022/5/18 排除智權人員的部門是F1X或F2X者，因為他們的通知是在晚在批次做的，收文程式是控制其他部門的人。
            If Left(PUB_GetStaffST15(Trim(txtPatent(15)), "1"), 2) <> "F1" And Left(PUB_GetStaffST15(Trim(txtPatent(15)), "1"), 2) <> "F2" Then
            '2022/5/18 END
               If MsgBox("此案有前款未收，請確認主管是否核示??", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                   Screen.MousePointer = vbDefault
                   GoTo ErrorHandler
               End If
            End If
         End If 'End 2017/06/19
      End If
      '2015/8/26 END
      
      'Added by Lydia 2019/05/13 改模組(一併取得)
      If Left(m_SalesST15, 1) <> "F" And txtPatent(8).Text <> "" And Val(txtPatent(17)) > 0 Then
          'Modified by Lydia 2022/06/13 傳入收文之本所案號,案件性質(可用,串接)
          'Call PUB_GetBillDataAll("3", txtPatent(8), dblAmt, dblPFee, dblTFee, , , TransDate(txtPatent(0), 2), mBillNo, mMemo)
          'Modified by Lydia 2022/06/15 傳入收文之智權人員
          Call PUB_GetBillDataAll("3", txtPatent(8), txtSystem & IIf(txtCode(0) <> "", txtCode(0) & Left(txtCode(1) & "0", 1) & Left(txtCode(2) & "00", 2), ""), txtPatent(1), Trim(txtPatent(15)), dblAmt, dblPFee, dblTFee, , , TransDate(txtPatent(0), 2), mBillNo, mMemo)
      End If
      
      'Add By Sindy 2012/11/06 非T*案件(TF要含)若已送件之應收款超過15萬以上,智權人員非國外部且有費用者須做下列控管
'      If Left(PUB_GetStaffST15(Trim(txtPatent(15)), "1"), 1) <> "F" And _
'         Val(txtPatent(17)) > 0 And _
'         Check2.Value = 0 Then
       'Modified by Morgan 2014/11/18 +判斷有申請人編號
      If Left(m_SalesST15, 1) <> "F" And _
         Val(txtPatent(17)) > 0 And _
         Check2.Value = 0 And txtPatent(8) <> "" Then
      'end 2014/11/18
         'Mark by Lydia 2019/05/13 改模組(一併取得)
         'GetBillData txtPatent(8), dblAmt, dblPFee, dblTFee
         
         'Add By Sindy 2012/12/10 取得客戶應收帳款收文檢查上限
         'Modified by Lydia 2020/02/03 應收帳款上限分開管制為個人"應收帳款上限"和"集團應收帳款上限"
         'dblChkAmt = PUB_GetCustRecAmtLmt(txtPatent(8))
         ''2012/12/10 End
         dblCu183 = PUB_GetCustRecAmtLmt(txtPatent(8), dblChkAmt)
         'Added by Lydia 2020/02/03 判斷是否有集團上限
         If dblChkAmt = 0 Then
             dblAmtR = 0: dblPFeeR = 0: dblTFeeR = 0
         Else   '有集團上限才抓關係企業的應收帳款金額
             GetBillData txtPatent(8), dblAmtR, dblPFeeR, dblTFeeR
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
         'If InStr(cnt應收帳款檢查排除, Trim(txtPatent(15))) = 0 And dblAmt >= dblCu183 Or (dblAmtR >= dblChkAmt And dblChkAmt > 0) Then
         If InStr(Pub_GetSpecMan("應收帳款上限檢查排除"), Trim(txtPatent(15))) = 0 And dblAmt >= dblCu183 Or (dblAmtR >= dblChkAmt And dblChkAmt > 0) Then
            'Modified by Lydia 2018/09/20 預設按鈕改成"否" vbDefaultButton1=>vbDefaultButton2
            If MsgBox("請注意接洽單上是否有註明應收帳款超額，需主管簽核才可收文！是否可收文？" & vbCrLf & _
                      "（接洽單上若有★★的應收帳款簽核控管，是否已勾選畫面上的註記欄位了？）", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               Screen.MousePointer = varSaveCursor
               GoTo ErrorHandler
            End If
'         '已送件之應收款超過15萬以上(不含T*案件應收款),提醒
'         ElseIf dblAmt >= 150000 Then
'            If MsgBox("請注意接洽單上是否有註明應收帳款超額，需主管簽核才可收文！是否可收文？" & vbCrLf & _
'                      "（接洽單上若有★★的應收帳款簽核控管，是否已勾選畫面上的註記欄位了？）", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
'               Screen.MousePointer = varSaveCursor
'               GoTo ErrorHandler
'            End If
         End If
      End If
      '2012/11/06 End
      
      'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
      'Moddified by Lydia 2019/01/19 排除國外部
      'If txtPatent(8).Text <> "" And Val(txtPatent(17)) > 0 Then
      If Left(m_SalesST15, 1) <> "F" And txtPatent(8).Text <> "" And Val(txtPatent(17)) > 0 Then
         'Modified by Lydia 2019/05/13 改模組(一併取得)
         'If GetBillDate(txtPatent(8), TransDate(txtPatent(0), 2), strExc(1), strExc(2)) = True Then
         If mMemo <> "" Then
             'Modified by Lydia 2018/10/29 改訊息
             'If MsgBox("請注意接洽單上是否有註明" & vbCrLf & strExc(2) & vbCrLf & "，請交主管簽核並且有主管簽核。" & vbCrLf & "是否可收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
             If MsgBox("請注意接洽單上是否有註明" & vbCrLf & mMemo & "，請交主管簽核。" & vbCrLf & "並且有主管簽核，是否可收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                Screen.MousePointer = varSaveCursor
                GoTo ErrorHandler
             End If
         End If
      End If
      'end 2018/08/22
      
      'Added by Lydia 2017/11/14 FCP案件命名電子化：中說輸入相關設定-檢查
     'Modified by Lydia 2019/06/11 判斷走命名流程才檢查; FCP-62285收307分割,修改時增加申請人2-X76639會彈未輸入分案組別
      'If fraTCT.Visible = True And fraTCT.Enabled = True Then
      If fraTCT.Visible = True And fraTCT.Enabled = True And InStr(FcpAddTct, txtPatent(1)) > 0 Then
         If Trim(txtData(0) & txtData(1)) <> "" Or ChkExpDate.Value = 1 Then
             ChkExpDate.Value = 1
             If Trim(txtData(0)) = "" Or Trim(txtData(1)) = "" Then
                MsgBox "急件請輸入譯畢期限!", vbExclamation
                If Trim(txtData(0)) = "" Then
                   txtData(0).SetFocus
                   Txtdata_GotFocus 0
                Else
                   txtData(1).SetFocus
                   Txtdata_GotFocus 1
                End If
                Screen.MousePointer = varSaveCursor
                GoTo ErrorHandler
             End If
         End If

         If bolExistTCT = False And ChkExpDate.Value = 1 Then
            strExc(1) = strSrvDate(2) & Mid(Format(ServerTime, "000000"), 1, 4)
            If Trim(txtData(0).Text) & Trim(txtData(1)) < strExc(1) Then
               MsgBox "譯畢期限不可早於系統日期和時間!! ", vbExclamation
               Screen.MousePointer = varSaveCursor
               GoTo ErrorHandler
            End If
         End If
         
         If txtData(2) = "" Then
            MsgBox "請輸入分案組別!! ", vbExclamation
            txtData(2).SetFocus
            Txtdata_GotFocus 2
            Screen.MousePointer = varSaveCursor
            GoTo ErrorHandler
         End If
         'Modified by Lydia 2018/05/23 香港案標準專利記錄請求(110),可以不輸入中說
         'If txtData(3).Visible = True And txtData(3).Text = "" Then
         'Modified by Lydia 2019/07/11 澳門案不用輸入中說, 因為都是拿台灣案的中說去申請
         'If txtData(3).Visible = True And txtData(3).Text = "" And txtPatent(1) <> "110" Then
         If txtData(3).Visible = True And txtData(3).Text = "" And txtPatent(1) <> "110" And txtPatent(4) <> "044" Then
            MsgBox "請輸入中說類型!! ", vbExclamation
            txtData(3).SetFocus
            Txtdata_GotFocus 3
            Screen.MousePointer = varSaveCursor
            GoTo ErrorHandler
         End If
         For intI = 0 To IIf(txtData(3).Visible = True, 3, 2)
            Txtdata_Validate intI, Cancel
            If Cancel = True Then
               Screen.MousePointer = varSaveCursor
               GoTo ErrorHandler
            End If
         Next
      End If
      'end 2017/11/14
      
         'Add by Lydia 2014/10/31 開放外專程序人員可進入專利處系統操作FMP寰華案件=>收文時詢問CP44是否設為Y53374000
         mFMPchk = False
         If Trim(txtSystem) = "P" And (mCP31 = "Y" Or frm010001.intSaveMode = 1) Then  '新增-無案號,修改-讀CP31
            If Left(m_SalesST15, 1) = "F" Then
                If MsgBox("請確認是否為寰華案件？", vbOKCancel) = 1 Then mFMPchk = True
            End If
         End If
         'end. 'Add by Lydia 2014/10/31
         
      'Added by Lydia 2020/04/29 因為109/03/30以後較常發生FCP新案立卷搬檔不全，所以在存檔前先下載檔案並且比對檔案，下載有問題彈相關訊息，直到完整下載方可存檔。
      If frm010001.intModifyKind = 0 And Trim(txtTCN01.Text) <> "" Then
         '改用原始檔區存放
         If PUB_ChkTCNfileExist(txtTCN01.Text) = True Then
            bolMoveCheck = True
         Else
            'Added by Lydia 2024/06/18 測試新案立卷是否要搬檔案
            strExc(1) = ""
            If Pub_StrUserSt03 = "M51" Or InStr(UCase(Forms(0).Caption), "M51") > 0 Then
               If MsgBox("測試新案立卷是否要搬檔案？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                  strExc(1) = "Y"
               End If
            Else
               strExc(1) = "Y"
            End If
            If strExc(1) = "Y" Then
            'end 2024/06/18
               MsgBox "請先聯絡 " & IIf(lblSales.Caption <> "", lblSales.Caption, "外專承辦人員") & vbCrLf & "上傳TRACKING_NO檔案！", vbCritical + vbOKOnly, "TRACKING_NO檔案稽核"
               Screen.MousePointer = varSaveCursor
               GoTo ErrorHandler
            End If
         End If
         'end 2023/10/12
      End If
      'end 2020/04/29
      
      'Add By Sindy 2021/6/23 外專人員時檢查...
      If (Me.txtSystem.Text = "FCP" Or Me.txtSystem.Text = "FG") And strSrvDate(1) >= 外專台灣案約定期限啟用日 _
         And Left(Pub_StrUserSt03, 2) = "F2" Then
         If m_strCPM34 = "Y" And Val(txtPatent(19).Text) = 0 Then
            If MsgBox("此案件性質屬有主管機關期限，確定沒有法定期限嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
               txtPatent(19).SetFocus
               Screen.MousePointer = varSaveCursor
               GoTo ErrorHandler
            End If
         ElseIf m_strCPM34 = "N" And Val(txtPatent(19).Text) > 0 Then
            If MsgBox("此案件性質屬非主管機關期限，確定有法定期限嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
               txtPatent(19).SetFocus
               Screen.MousePointer = varSaveCursor
               GoTo ErrorHandler
            End If
         End If
      End If
      '2021/6/23 END
      
      'Added by Morgan 2021/7/21 大陸發明生醫案是否新藥專利設定
      strPA176 = ""
      If Trim(txtSystem) = "P" And txtPatent(4) = "020" And txtPatent(3) = "1" And Left(Combo3, 1) = "3" And (mCP31 = "Y" Or frm010001.intSaveMode = 1) Then
         intI = MsgBox("是否新藥專利？" & vbCrLf & vbCrLf & "請看接洽單案件說明事項處理情形!!", vbYesNoCancel + vbDefaultButton3 + vbQuestion, "大陸發明生醫案是否新藥專利確認")
         If intI = vbYes Then
            strPA176 = "Y"
         ElseIf intI = vbNo Then
            strPA176 = "N"
         Else
            Screen.MousePointer = varSaveCursor
            GoTo ErrorHandler
         End If
      End If
      'end 2021/7/21

      'Added by Lydia 2022/07/04 一案兩請僅其中一案收文701~709、401變更時，於確認收文時彈提醒：為一案兩請，請確認發明案及新型案是否一併收文。
      If frm010001.intModifyKind = 0 And (txtSystem = "FCP" Or txtSystem = "P") And txtCode(0) <> "" And (Left(txtPatent(1), 2) = "70" Or txtPatent(1) = "401") Then
          strExc(0) = ""
          If txtSystem = "P" Then
              If PUB_ChkIsFMP(txtSystem, txtCode(0), Left(txtCode(1) & "0", 1), Left(txtCode(2) & "00", 2)) = False Then
                  strExc(0) = "N"
              End If
          End If
          If strExc(0) <> "N" Then
              strExc(1) = txtSystem:   strExc(2) = txtCode(0):   strExc(3) = txtCode(1):  strExc(4) = txtCode(2)
              If PUB_IsDualApply(strExc, strDCase, , , , , , True) = True Then
                 If MsgBox(txtSystem & "-" & txtCode(0) & IIf(txtCode(1) & txtCode(2) <> "000", "-" & txtCode(1) & "-" & txtCode(2), "") & "為一案兩請，請確認發明案及新型案是否一併收文，" & vbCrLf & _
                       "另一案件：" & strDCase(1) & "-" & strDCase(2) & IIf(strDCase(3) & strDCase(4) <> "000", "-" & strDCase(3) & "-" & strDCase(4), "") & vbCrLf & vbCrLf & _
                        "選擇""是""會繼續作業，選擇""否""會中斷作業。", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
                       GoTo ErrorHandler
                 End If
              End If
          End If
      End If
      'end 2022/07/04
      
         'Add By Sindy 2022/7/1
         If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
            If PUB_ChkFileOpening2(m_PrevForm.m_strFullFileName, "後續才能一併歸卷！") = True Then
               Screen.MousePointer = varSaveCursor
               GoTo ErrorHandler
            End If
         End If
         '2022/7/1 END
         'Modified by Lydia 2022/09/05 判斷啟用日
         'If SaveDatabase(strAuto1, strAuto2) Then
         bolSaveOK = False
         
         'Removed by Morgan 2024/11/18 收文存檔模組已啟用,舊程式標記為註解,後續無需再修改
         'If strSrvDate(1) < 收文存檔模組化啟用日 Then
         '    bolSaveOK = SaveDatabase(strAuto1, strAuto2)
         'Else
         'end 2024/11/18
         
             Call SetDBArray(False, txtRecieveCode, txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
             bolSaveOK = PUB_SaveFrm010005(Me.Name, frm010001.intSaveMode, frm010001.intModifyKind, frm010001.intChoose, modBase, modCP, txtPatent(16), strInventorNo, mChkStr, IsSaveData, mType, mCaseNo, mRetVal, mTCTVal, mTCTList)
                
             If frm010001.intModifyKind = 0 And bolSaveOK = True Then
                '注意若是外專要輸入命名追蹤流水號的新案，要在存檔之前的檢查先完成下載檔案(PUB_GetTCNfile)，存檔完成後開始從本機端上傳到原始檔區和卷宗區(PUB_MoveTCNfile)
                 txtCode(0) = modBase(2)
                 strAuto1 = modCP(9)
                 strAuto2 = modBase(2)
             End If
             If bolSaveOK = True Then
                 'Added by Lydia 2022/09/16 外專命名之收文號
                 If mTCTList <> "" Then
                    frm010001.lblTCT.Caption = "中說或其他收文號："
                    frm010001.lblTCTNO.Caption = mTCTList
                 End If
                 'end 2022/09/16
                 '外專信件沖銷: 收完文
                 If InStr("," & mRetVal, "m_bolRecvOK = True") > 0 Then
                    m_bolRecvOK = True
                 End If
                 '多案收文的總收文號要傳入第一筆總收文號
                 If InStr("," & mRetVal, "MCR11:") > 0 Then
                     m_strMCR11 = Mid(mRetVal, InStr(mRetVal, "MCR11:") + 6, 9)
                 End If
             End If
             
         'End If 'Removed by Morgan 2024/11/18
         
'-----------------------------------------------
         If bolSaveOK = True Then
         'end 2022/09/05
             'add by nick 2004/07/27 提示要同時收文 211及 212
             'If Trim(txtSystem.Text) = "P" And txtPatent(4).Text = "000" And txtPatent(1).Text = "503" And Val(txtPatent(17).Text) > 45000 Then
             'Modify By Sindy 2009/06/12
             If Trim(txtSystem.Text) = "P" And txtPatent(4).Text = "000" And txtPatent(1).Text = "503" And Val(txtPatent(17).Text) > 20000 Then
                   MsgBox "請 P 國內行政訴訟案，請同時收文 211 準備程序及 212 言詞辯論！", , "請注意！"
             End If
                              
            'Added by Lydia 2017/12/05 FCP案件命名電子化：搬移命名追蹤檔案
            'Modified by Lydia 2020/04/29 +判斷已完成本機端下載bolMoveCheck
            If frm010001.intModifyKind = 0 And strSrvDate(1) >= FCP案件命名啟用日 And Trim(txtTCN01.Text) <> "" And bolMoveCheck = True Then
                'Added by Lydia 2020/04/29 FCP新案立卷: 先從Typing2\English_Vers下載檔案到本機端
                'Move by Lydia 2022/09/01 從frm010005搬來，並且改成共用PUB_GetTCNfile,PUB_MoveTCNfile
                '因為109/3/30以後較常發生FCP新案立卷搬檔不全發email通知(Ex.4/29的FCP063149(13556),FCP063150(13565))，推測是原本TrackingNo資料夾的所有檔案下載到本機端不完整，然後未經過檢查直接進行上傳作業。
                '修改程式為二階段作業：
                '1.按「確定」時先下載檔案並且比對檔案的上傳規則，若經過檢查檔案有問題則彈相關訊息「檔案xxx正在使用中/檔案名稱不符合規則/檔案名稱過長，請先聯絡智權人員XXX將檔案關閉/修改檔名」，並且中斷存檔作業。
                '2.直到下載檔案完成才能繼續存檔作業從本機端搬到原始檔區和卷宗區，並且刪除TRACKING_NO資料夾。
                '2023/08/21改用原始檔區存放---2024/12/13 以上Code刪除

                 Call PUB_UpdTCNfile(txtTCN01, txtSystem & txtCode(0) & IIf(txtCode(1) = "", "0", txtCode(1)) & IIf(txtCode(2) = "", "00", txtCode(2)), strAuto1, DBDATE(txtPatent(0)), mSaveDir, bolMoveOK)

            End If
            'end 2017/12/05
            
JumpToNext: 'Added by Lydia 2019/12/17

            'Added by Lydia 2018/09/06 櫃臺P案(FMP)中間程序和中間接進來案件的收文，系統自動發e-mail通知。
            'Modified by Lydia 2022/09/05 判斷存檔模組化啟用日：現在直接併入存檔模組
            If strSrvDate(1) < 收文存檔模組化啟用日 And frm010001.intModifyKind = 0 And txtSystem = "P" And Left(m_SalesST15, 2) = "F2" And frm010001.mRole = "" And InStr(FcpAddTct, txtPatent(1)) = 0 Then
                 PUB_SendMail strUserNum, txtPatent(15).Text, "", txtSystem & "-" & txtCode(0) & IIf(Val(txtCode(1) & txtCode(2)) > 0, "-" & txtCode(1) & "-" & txtCode(2), "") & " 已收文" & lblCaseProperty.Caption & " , 請進卷宗區移檔!", "同主旨"
            End If
            'end 2018/09/06
            
            PUB_SendMailCache 'Added by Morgan 2013/8/1 'Move by Lydia 2018/03/29 從FCP案件命名電子化：搬移命名追蹤檔案的上方移下來
            
            'Add By Sindy 2022/7/11 信件沖銷多案收文
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
                  txtPatent(1) = RsTemp.Fields("mcr06")
                  IsSaveData = False
                  ReadPatentDatabaseR '重新查詢
                  Call SetDBArray(True, txtRecieveCode, txtSystem, txtCode(0), txtCode(1), txtCode(2)) 'Added by Morgan 2025/3/3 變數也要重設否則會抓到前一案資料
                  DoEvents
                  cmdok(0).Value = True
                  Exit Sub
               End If
            End If
            '2022/7/11 END
            
            frm010001.ClearForm strAuto1, strAuto2
            bolLeave = True
            intLeaveKind = 1
            If frm010001.intModifyKind = 0 Then LastDate = txtPatent(0).Text
            
            'Modify By Sindy 2022/6/29 信件內部收文執行完畢後,關閉視窗
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
               Unload Me
            End If
            '2022/6/29 END
         End If
      End If
      Screen.MousePointer = vbDefault
   Else
      If Index = 2 Then
         intLeaveKind = 0
      Else
         intLeaveKind = 1
      End If
      Unload Me
   End If
   
   Exit Sub
   
'Add By Sindy 2012/3/8
ErrorHandler:
   If PUB_CheckFormExist("frmpic002") = True Then Unload frmpic002 'Add By Sindy 2022/7/11
   
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Removed by Morgan 2024/11/18 收文存檔模組已啟用,舊程式標記為註解,後續無需再修改
'Private Function SaveDatabase(ByRef strRecieveAuto As String, ByRef strCaseAuto As String) As Boolean
'Dim adoquery As New ADODB.Recordset
'Dim strPA149 As String, strContact As String
'Dim strAddNo As String 'Added by Lydia 2017/11/14
'Dim strFromPath As String, strToPath As String  'Added by Lydia 2017/12/04
'
'   'Add by Morgan 2008/8/5
'   If cboContact.Locked = False Then
'      If cboContact.ListIndex >= 0 Then
'         If Val(cboContact.ItemData(cboContact.ListIndex)) > 0 Then
'            'Modify by Lydia2022/09/16 改成Form 2.0
'            'strPA149 = Format(cboContact.ItemData(cboContact.ListIndex), "00")
'            ''Add by Morgan 2008/8/7 若個案接洽人與客戶檔的預設接洽人相同時不必設定
'            'PUB_GetContact strAppNo1, strContact, True
'            'If strPA149 = strContact Then
'            '   strPA149 = ""
'            'End If
'            strPA149 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
'            If Val(strPA149) > 0 Then
'                PUB_GetContact strAppNo1, strContact, True
'                If strPA149 = strContact Then
'                   strPA149 = ""
'                End If
'            '排除空白=00
'            ElseIf strPA149 = "00" And Trim(cboContact.Text) = "" Then
'                strPA149 = ""
'            End If
'            'end 2022/09/16
'         End If
'      End If
'   Else
'      strPA149 = "PA149"
'   End If
'
'   'Modified by Lydia 2019/09/16
'   'm_SalesST15 = GetST15(txtPatent(15).Text) 'Added by Lydia 2018/09/06
'   m_SalesST15 = GetST15(txtPatent(15).Text, , , m_SalesST06)
'
'   '若為新增
'   If frm010001.intModifyKind = 0 Then
'      If strPA149 = "PA149" Then strPA149 = "" 'Add by Morgan 2008/8/7
'      'edit by nickc 2007/03/27 加入彼所案號
'      'Modified by Morgan 2021/7/21 +PA176
'      SaveDatabase = InsertPatentDatabase(frm010001.intSaveMode, txtSystem, txtCode(0), _
'              IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(5), txtPatent(6), txtPatent(7), txtPatent(3), txtPatent(4), _
'              txtPatent(8), txtPatent(11), txtPatent(9), txtPatent(12), txtPatent(10), txtPatent(13), _
'              txtPatent(0), txtPatent(14), txtPatent(19), txtPatent(1), txtPatent(2), txtPatent(15), _
'              txtPatent(17), txtPatent(21), txtPatent(18), txtPatent(22), txtPatent(20), txtPatent(23), _
'              txtPatent(16), txtPatent(24), strRecieveAuto, strCaseAuto, douStPrice, douLowPrice, txtCP64, _
'              txtPetitionx(2), txtPetitionx(3), txtPetitionx(4), txtPetitionx(5), txtPatent(27), strPA149, strPA176)
'   '若為修改
'   Else
'      'edit by nickc 2007/03/27 加入彼所案號
'      'Modified by Morgan 2021/7/21 +PA176
'      SaveDatabase = UpdatePatentDatabase(frm010001.intSaveMode, txtSystem, txtCode(0), _
'              IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(5), txtPatent(6), txtPatent(7), txtPatent(3), txtPatent(4), _
'              txtPatent(8), txtPatent(11), txtPatent(9), txtPatent(12), txtPatent(10), txtPatent(13), _
'              txtRecieveCode, txtPatent(0), txtPatent(14), txtPatent(19), txtPatent(1), txtPatent(2), txtPatent(15), _
'              txtPatent(17), txtPatent(21), txtPatent(18), txtPatent(22), txtPatent(20), txtPatent(23), txtPatent(16), txtPatent(24), douStPrice, douLowPrice, txtCP64, _
'              txtPetitionx(2), txtPetitionx(3), txtPetitionx(4), txtPetitionx(5), txtPatent(27), strPA149, strPA176)
'   End If
''add by nickc 2007/11/09 測試解決mail 發不到的時候會存兩筆的錯誤
''on error GoTo 0    '歸零
'
'   'Added by Lydia 2021/01/08 CFP英國脫歐案：複製代表圖
'   If txtSystem = "CFP" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" Then
'      strExc(9) = ""
'      If GetImgByteFile_Case(m_CaseNa239(1), m_CaseNa239(2), m_CaseNa239(3), m_CaseNa239(4), strExc(9), 0, strExc(5), strExc(6)) = True Then
'          Call SaveImgByteFile(strExc(9), txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strExc(5), strExc(6))
'      End If
'   End If
'   'end 2021/01/08
'
'   'add by nickc 2005/09/05
'   If frm010001.intModifyKind = 0 Then
'      Dim oContext As String
'      Dim oMailCount As String
'      Dim strTemp As String
'      'Add By Sindy 2021/2/1 不得代理的後續舊案收文控管，通知收文人員（CP13）
'      If InStr(strYState, "不得代理") > 0 Then
'         oContext = oContext & vbCrLf + "代理人： " + ChangeCustomerL(txtPatent(13).Text) + " " + lblAgent.Caption + vbCrLf
'         strTemp = strTemp & "," & ChangeCustomerL(txtPatent(13).Text)
'      End If
'      If InStr(strXState(8), "不得代理") > 0 Then
'         oContext = oContext & vbCrLf + "申請人1： " + ChangeCustomerL(txtPatent(8).Text) + " " + lblPetition(0).Caption + vbCrLf
'         strTemp = strTemp & "," & ChangeCustomerL(txtPatent(8).Text)
'      End If
'      If InStr(strXState(11), "不得代理") > 0 Then
'         oContext = oContext & vbCrLf + "申請人2： " + ChangeCustomerL(txtPatent(11).Text) + " " + lblPetition(3).Caption + vbCrLf
'         strTemp = strTemp & "," & ChangeCustomerL(txtPatent(11).Text)
'      End If
'      If InStr(strXState(9), "不得代理") > 0 Then
'         oContext = oContext & vbCrLf + "申請人3： " + ChangeCustomerL(txtPatent(9).Text) + " " + lblPetition(1).Caption + vbCrLf
'         strTemp = strTemp & "," & ChangeCustomerL(txtPatent(9).Text)
'      End If
'      If InStr(strXState(12), "不得代理") > 0 Then
'         oContext = oContext & vbCrLf + "申請人4： " + ChangeCustomerL(txtPatent(12).Text) + " " + lblPetition(4).Caption + vbCrLf
'         strTemp = strTemp & "," & ChangeCustomerL(txtPatent(12).Text)
'      End If
'      If InStr(strXState(10), "不得代理") > 0 Then
'         oContext = oContext & vbCrLf + "申請人5： " + ChangeCustomerL(txtPatent(10).Text) + " " + lblPetition(2).Caption + vbCrLf
'         strTemp = strTemp & "," & ChangeCustomerL(txtPatent(10).Text)
'      End If
'      If InStr(strXState(23), "不得代理") > 0 Then
'         oContext = oContext & vbCrLf + "讓與申請人1： " + ChangeCustomerL(txtPatent(23).Text) + " " + lblPetitionName.Caption + vbCrLf
'         strTemp = strTemp & "," & ChangeCustomerL(txtPatent(23).Text)
'      End If
'      If InStr(strXState(2), "不得代理") > 0 Then
'         oContext = oContext & vbCrLf + "讓與申請人2： " + ChangeCustomerL(txtPetitionx(2).Text) + " " + lblPetitionNamex(2).Caption + vbCrLf
'         strTemp = strTemp & "," & ChangeCustomerL(txtPetitionx(2).Text)
'      End If
'      If InStr(strXState(3), "不得代理") > 0 Then
'         oContext = oContext & vbCrLf + "讓與申請人3： " + ChangeCustomerL(txtPetitionx(3).Text) + " " + lblPetitionNamex(3).Caption + vbCrLf
'         strTemp = strTemp & "," & ChangeCustomerL(txtPetitionx(3).Text)
'      End If
'      If InStr(strXState(4), "不得代理") > 0 Then
'         oContext = oContext & vbCrLf + "讓與申請人4： " + ChangeCustomerL(txtPetitionx(4).Text) + " " + lblPetitionNamex(4).Caption + vbCrLf
'         strTemp = strTemp & "," & ChangeCustomerL(txtPetitionx(4).Text)
'      End If
'      If InStr(strXState(5), "不得代理") > 0 Then
'         oContext = oContext & vbCrLf + "讓與申請人5： " + ChangeCustomerL(txtPetitionx(5).Text) + " " + lblPetitionNamex(5).Caption + vbCrLf
'         strTemp = strTemp & "," & ChangeCustomerL(txtPetitionx(5).Text)
'      End If
'      If oContext <> "" Then
'         strTemp = Mid(strTemp, 2)
'         oContext = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + _
'                    "案件名稱： " + txtPatent(5) + vbCrLf + _
'                    "收文日： " + ChangeTStringToTDateString(txtPatent(0)) + vbCrLf + _
'                    "案件性質： " + lblCaseProperty.Caption + vbCrLf + vbCrLf + _
'                    "【不得代理】" + vbCrLf + _
'                    oContext
'         oMailCount = Trim(txtPatent(15).Text) & ";" & PUB_GetFCPProSup(Trim(txtPatent(15).Text))
'         PUB_SendMail strUserNum, oMailCount, "", IIf("-" + txtCode(1) + "-" + txtCode(2) = "-0-00", txtSystem + "-" + txtCode(0), txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2)) & _
'            " 已確認續行收文，請注意該" & strTemp & "編號已設為不得代理。", oContext
'      End If
'      '2021/2/1 END
'
'      'add by nick 2004/10/15  當收文業務區與客戶檔業務區不同時發 mail  及提示
'      Dim oStrCuSales1 As String
'      Dim oStrCuSales2 As String
'      Dim oStrCuSales3 As String
'      Dim oStrCuSales4 As String
'      Dim oStrCuSales5 As String
'      'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
'      Dim IsMail  As Boolean
'      IsMail = True
'      'add by nickc 2007/05/16 加入若是本所期限小於等於當天，要發mail  通知
'      Dim oContext2 As String
'      oContext = "": oContext2 = ""
'
'      oStrCuSales1 = ""
'      oStrCuSales2 = ""
'      oStrCuSales3 = ""
'      oStrCuSales4 = ""
'      oStrCuSales5 = ""
'      oContext = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtPatent(5) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtPatent(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
'      'add by nickc 2007/05/16 加入若是本所期限小於等於當天，要發mail  通知
'      'edit by nickc 2008/04/23  加入國家
'      'oContext2 = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtPatent(5) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtPatent(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
'      oContext2 = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtPatent(5) + vbCrLf + "申請國家：" + txtPatent(4) + " " + lblNation.Caption + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtPatent(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
'
'      oMailCount = ""
'      'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
'      'If m_salesst15 <> GetCuSales(ChangeCustomerL(txtPatent(8).Text), oStrCuSales1) And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(8).Text) <> "" Then
'      'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
'      If ChkSameCuArea(Trim(txtPatent(8)), Trim(txtPatent(15)), , , , , Trim(txtPatent(13))) = False And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(8).Text) <> "" Then
'         'Add By Sindy 2009/10/19
'         If Left(Trim(m_SalesST15), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtPatent(8).Text), oStrCuSales1)), 1) = "F" Then
'            '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
'         Else
'            oMailCount = oMailCount & oStrCuSales1 & ";"
'            'edit by nickc 2005/08/16
'            'oContext = oContext & vbCrLf + "申請人1： " + GetCustomerName(ChangeCustomerL(txtPatent(8).Text)) + "原智權人員： " + oStrCuSales1
'            oContext = oContext & vbCrLf + "申請人1： " + GetCustomerName(ChangeCustomerL(txtPatent(8).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales1)
'         End If
'      'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
'      Else
'           If Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(8).Text) <> "" Then
'               IsMail = False
'           End If
'      End If
'      'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
'      If m_SalesST06 <> "" And Trim(txtPatent(8)) <> "" And Trim(txtPatent(15)) <> "" Then
'          If PUB_ChkOldCustomer(True, txtPatent(8), Trim(txtPatent(15)), m_SalesST15, m_SalesST06) = True Then
'             IsMail = False
'         End If
'      End If
'
'      'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
'      'If m_salesst15 <> GetCuSales(ChangeCustomerL(txtPatent(11).Text), oStrCuSales2) And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(11).Text) <> "" Then
'      'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
'      If ChkSameCuArea(Trim(txtPatent(11)), Trim(txtPatent(15)), , , , , Trim(txtPatent(13))) = False And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(11).Text) <> "" Then
'         'Add By Sindy 2009/10/19
'         If Left(Trim(m_SalesST15), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtPatent(11).Text), oStrCuSales2)), 1) = "F" Then
'            '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
'         Else
'            oMailCount = oMailCount & oStrCuSales2 & ";"
'            'edit by nickc 2005/08/16
'            'oContext = oContext & vbCrLf + "申請人4： " + GetCustomerName(ChangeCustomerL(txtPatent(11).Text)) + "原智權人員： " + oStrCuSales4
'            oContext = oContext & vbCrLf + "申請人2： " + GetCustomerName(ChangeCustomerL(txtPatent(11).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales2)
'         End If
'      'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
'      Else
'           If Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(11).Text) <> "" Then
'               IsMail = False
'           End If
'      End If
'      'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
'      If m_SalesST06 <> "" And Trim(txtPatent(11)) <> "" And Trim(txtPatent(15)) <> "" Then
'          If PUB_ChkOldCustomer(True, txtPatent(11), Trim(txtPatent(15)), m_SalesST15, m_SalesST06) = True Then
'             IsMail = False
'         End If
'      End If
'
'      'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
'      'If m_salesst15 <> GetCuSales(ChangeCustomerL(txtPatent(9).Text), oStrCuSales3) And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(9).Text) <> "" Then
'      'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
'      If ChkSameCuArea(Trim(txtPatent(9)), Trim(txtPatent(15)), , , , , Trim(txtPatent(13))) = False And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(9).Text) <> "" Then
'         'Add By Sindy 2009/10/19
'         If Left(Trim(m_SalesST15), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtPatent(9).Text), oStrCuSales3)), 1) = "F" Then
'            '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
'         Else
'            oMailCount = oMailCount & oStrCuSales3 & ";"
'            'edit by nickc 2005/08/16
'            'oContext = oContext & vbCrLf + "申請人2： " + GetCustomerName(ChangeCustomerL(txtPatent(9).Text)) + "原智權人員： " + oStrCuSales2
'            oContext = oContext & vbCrLf + "申請人3： " + GetCustomerName(ChangeCustomerL(txtPatent(9).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales3)
'         End If
'      'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
'      Else
'           If Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(9).Text) <> "" Then
'               IsMail = False
'           End If
'      End If
'      'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
'      If m_SalesST06 <> "" And Trim(txtPatent(9)) <> "" And Trim(txtPatent(15)) <> "" Then
'          If PUB_ChkOldCustomer(True, txtPatent(9), Trim(txtPatent(15)), m_SalesST15, m_SalesST06) = True Then
'             IsMail = False
'         End If
'      End If
'
'      'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
'      'If m_salesst15 <> GetCuSales(ChangeCustomerL(txtPatent(12).Text), oStrCuSales4) And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(12).Text) <> "" Then
'      'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
'      If ChkSameCuArea(Trim(txtPatent(12)), Trim(txtPatent(15)), , , , , Trim(txtPatent(13))) = False And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(12).Text) <> "" Then
'         'Add By Sindy 2009/10/19
'         If Left(Trim(m_SalesST15), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtPatent(12).Text), oStrCuSales4)), 1) = "F" Then
'            '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
'         Else
'            oMailCount = oMailCount & oStrCuSales4 & ";"
'            'edit by nickc 2005/08/16
'            'oContext = oContext & vbCrLf + "申請人5： " + GetCustomerName(ChangeCustomerL(txtPatent(12).Text)) + "原智權人員： " + oStrCuSales5
'            oContext = oContext & vbCrLf + "申請人4： " + GetCustomerName(ChangeCustomerL(txtPatent(12).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales4)
'         End If
'      'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
'      Else
'           If Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(12).Text) <> "" Then
'               IsMail = False
'           End If
'      End If
'      'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
'      If m_SalesST06 <> "" And Trim(txtPatent(12)) <> "" And Trim(txtPatent(15)) <> "" Then
'          If PUB_ChkOldCustomer(True, txtPatent(12), Trim(txtPatent(15)), m_SalesST15, m_SalesST06) = True Then
'             IsMail = False
'         End If
'      End If
'
'      'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
'      'If m_salesst15 <> GetCuSales(ChangeCustomerL(txtPatent(10).Text), oStrCuSales5) And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(10).Text) <> "" Then
'      'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
'      If ChkSameCuArea(Trim(txtPatent(10)), Trim(txtPatent(15)), , , , , Trim(txtPatent(13))) = False And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(10).Text) <> "" Then
'         'Add By Sindy 2009/10/19
'         If Left(Trim(m_SalesST15), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtPatent(10).Text), oStrCuSales5)), 1) = "F" Then
'            '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
'         Else
'            oMailCount = oMailCount & oStrCuSales5 & ";"
'            'edit by nickc 2005/08/16
'            'oContext = oContext & vbCrLf + "申請人3： " + GetCustomerName(ChangeCustomerL(txtPatent(10).Text)) + "原智權人員： " + oStrCuSales3
'            oContext = oContext & vbCrLf + "申請人5： " + GetCustomerName(ChangeCustomerL(txtPatent(10).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales5)
'         End If
'      'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
'      Else
'           If Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(10).Text) <> "" Then
'               IsMail = False
'           End If
'      End If
'      'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
'      If m_SalesST06 <> "" And Trim(txtPatent(10)) <> "" And Trim(txtPatent(15)) <> "" Then
'          If PUB_ChkOldCustomer(True, txtPatent(10), Trim(txtPatent(15)), m_SalesST15, m_SalesST06) = True Then
'             IsMail = False
'         End If
'      End If
'
'   'Remove by Morgan 2009/8/20 國外部智權人員改可收所內信件
'   '   '2008/12/3 ADD BY SONIA 客戶檔之智權人員為國外部者不發mail
'   '   If IsMail = True Then
'   '      IsMail = PUB_CHKcusales(Trim(ChangeCustomerL(txtPatent(8))) & "," & Trim(ChangeCustomerL(txtPatent(9))) & "," & Trim(ChangeCustomerL(txtPatent(10))) & "," & Trim(ChangeCustomerL(txtPatent(11))) & "," & Trim(ChangeCustomerL(txtPatent(12))))
'   '   End If
'   '   '2008/12/3 END
'
'      'edit by nickc 2007/08/21 若申請人全空白，不發
'      'If IsMail = False Then
'      If IsMail = False Or (Trim(txtPatent(8)) = "" And Trim(txtPatent(9)) = "" And Trim(txtPatent(10)) = "" And Trim(txtPatent(11)) = "" And Trim(txtPatent(12)) = "") Then
'           oMailCount = ""
'      End If
'
'      '2006/8/2 MODIFY BY SONIA TXTSYSTEM只判斷1碼,因為FG
'      'If UCase(Mid(txtSystem, 1, 2)) <> "FC" And oMailCount <> "" Then
'      If UCase(Mid(txtSystem, 1, 1)) <> "F" And oMailCount <> "" Then
'         'edit by nickc 2005/08/10
'         'MsgBox "收文智權人員與客戶智權人員不同業務區，準備發 mail ，請定時刪除郵件備份！", , "注意！"
'         'Modify By Sindy 2010/11/26 申請人1~5為 X65299 或 X03072 的所有關係企業都不檢查業務區
'         If Left(Trim(txtPatent(8)), 6) <> "X65299" And Left(Trim(txtPatent(8)), 6) <> "X03072" And _
'            Left(Trim(txtPatent(11)), 6) <> "X65299" And Left(Trim(txtPatent(11)), 6) <> "X03072" And _
'            Left(Trim(txtPatent(9)), 6) <> "X65299" And Left(Trim(txtPatent(9)), 6) <> "X03072" And _
'            Left(Trim(txtPatent(12)), 6) <> "X65299" And Left(Trim(txtPatent(12)), 6) <> "X03072" And _
'            Left(Trim(txtPatent(10)), 6) <> "X65299" And Left(Trim(txtPatent(10)), 6) <> "X03072" Then
'            MsgBox "收文智權人員與客戶智權人員不同業務區，準備發 mail ！", , "注意！"
'            'edit by nickc 2005/08/10 加發秀玲
'            'oMailCount = oMailCount & Trim(txtPatent(15).Text)
'            oMailCount = oMailCount & Trim(txtPatent(15).Text) & ";83002"
'            oContext = oContext & vbCrLf + "收文智權人員： " + lblSales.Caption + vbCrLf + vbCrLf + "智權人員(區)不同！"
'            'Modify by Morgan 2006/6/26 收文號或內容一定要有,否則不會寄
'            'PUB_SendMail strUserNum, oMailCount, "", oContext, ""
'            PUB_SendMail strUserNum, oMailCount, "", "案件收文通知--此案收文非原智權人員(區)！", oContext
'         End If
'      End If
'
'      'add by nickc 2007/05/16 加入若是本所期限小於等於當天，要發mail  通知
'      oMailCount = ""
'      'Added by Lydia 2015/12/30 FMP寰華案發給FCP人員
'      'If txtSystem = "P" Or txtSystem = "PS" Then
'      If PUB_FMPtoCheck(1, 2, "", txtSystem, txtCode(0), txtCode(1), txtCode(2)) = True Or mFMPchk = True Then
'           oMailCount = Pub_GetSpecMan("C")
'      ElseIf txtSystem = "P" Or txtSystem = "PS" Then
'      'end 2015/12/30
'           oMailCount = Pub_GetSpecMan("A")
'      ElseIf txtSystem = "CFP" Or txtSystem = "CPS" Then
'           'edit by nickc 2007/10/16 修改到table
'           'oMailCount = "PATENT"
'           oMailCount = Pub_GetSpecMan("B")
'      ElseIf txtSystem = "FCP" Or txtSystem = "FG" Then
'           'edit by nickc 2007/10/16 修改到table
'           'oMailCount = "73023;79012"
'           oMailCount = Pub_GetSpecMan("C")
'      End If
'      If DBDATE(txtPatent(14).Text) < strSrvDate(1) And Trim(txtPatent(14).Text) <> "" And Trim(oMailCount) <> "" Then
'         '2007/8/13 MODIFY BY SONIA 加智權人員
'         'Modify By Sindy 2010/12/16 加業務區,費用
'         PUB_SendMail strUserNum, oMailCount, "", "案件收文 緊急 通知--此案已逾本所期限，請儘速辦理！", oContext2 & vbCrLf & "本所期限：" & ChangeWStringToTDateString(DBDATE(txtPatent(14).Text)) & vbCrLf & "法定期限：" & ChangeWStringToTDateString(DBDATE(txtPatent(19).Text)) & vbCrLf & "智權人員　：" & lblSales & vbCrLf & "業務區　：" & lblDepartment & vbCrLf & "費用　　：" & Format(txtPatent(17), "##,##0")
'      End If
'      If DBDATE(txtPatent(14).Text) = strSrvDate(1) And Trim(txtPatent(14).Text) <> "" And Trim(oMailCount) <> "" Then
'         '2007/8/13 MODIFY BY SONIA 加智權人員
'         'Modify By Sindy 2010/12/16 加業務區,費用
'         PUB_SendMail strUserNum, oMailCount, "", "案件收文 緊急 通知--此案已屆本所期限，請儘速辦理！", oContext2 & vbCrLf & "本所期限：" & ChangeWStringToTDateString(DBDATE(txtPatent(14).Text)) & vbCrLf & "法定期限：" & ChangeWStringToTDateString(DBDATE(txtPatent(19).Text)) & vbCrLf & "智權人員　：" & lblSales & vbCrLf & "業務區　：" & lblDepartment & vbCrLf & "費用　　：" & Format(txtPatent(17), "##,##0")
'      End If
'
'      '2010/2/5 ADD BY SONIA 測試追蹤用
'      If txtSystem = "FCP" And txtPatent(1) = "305" Then
'         PUB_SendMail strUserNum, "83002", "", "FCP收文改請聯合，請追蹤申請案號變化情形！CP30為何會與新申請案號相同?", oContext2 & vbCrLf
'      End If
'      '2010/2/5 END
'   End If
'
'End Function

Private Sub ReadPatentDatabaseR()
Dim pa01 As String, pa02 As String, pa03 As String, pa04 As String, pa05 As String, _
              pa06 As String, pa07 As String, PA08 As String, PA09 As String, _
              pa21 As String, pa24 As String, pa25 As String, pa26 As String, _
              pa27 As String, pa28 As String, pa29 As String, pa30 As String, pa75 As String, _
              cp05 As String, cp06 As String, cp07 As String, CP10 As String, cp11 As String, _
              cp12 As String, cp13 As String, cp14 As String, cp16 As String, cp17 As String, cp18 As String, _
              cp19 As String, cp32 As String, cp56 As String, cu30 As String, CP64 As String, CP89 As String, _
              CP90 As String, CP91 As String, CP92 As String, i As Integer, rt As Boolean
'add by nickc 2007/03/27
Dim PA77 As String
Dim CP150 As String 'Add By Sindy 2012/11/08
Dim PA150 As String 'Added by Lydia 2017/11/14
Dim nPA01 As String, nPA02 As String, nPA03 As String, nPA04 As String 'Added by Lydia 2020/11/19 要讀取之案號
Dim PA91 As String 'Added by Lydia 2021/04/15 +PA91案件備註
Dim PA178 As String 'Add by Sindy 2022/12/7 +PA178證書形式
   
   m_strCP06 = "" 'Add By Sindy 2021/4/29
   'Added by Lydia 2020/11/19  CFP英國脫歐案管制：要讀取之案號
   nPA01 = IIf(m_CaseNa239(1) <> "", m_CaseNa239(1), txtSystem.Text)
   nPA02 = IIf(m_CaseNa239(2) <> "", m_CaseNa239(2), txtCode(0))
   nPA03 = IIf(m_CaseNa239(3) <> "", m_CaseNa239(3), IIf(txtCode(1) = "", "0", txtCode(1)))
   nPA04 = IIf(m_CaseNa239(4) <> "", m_CaseNa239(4), IIf(txtCode(2) = "", "00", txtCode(2)))
   'end 2020/11/19
   
   CP10 = txtPatent(1)
   m_CP10 = txtPatent(1) 'Add By Sindy 2009/09/01
   'edit by nickc 2007/03/27  加入彼所案號
   'rt = ReadPatentDatabase(frm010001.intModifyKind, txtSystem, txtCode(0), _
          IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), pa05, _
          pa06, pa07, PA08, PA09, pa26, pa27, pa28, pa29, pa30, pa75, txtRecieveCode, _
          cp05, cp06, cp07, CP10, cp11, cp13, cp16, cp17, cp18, cp19, cp32, cp56, cu30, cp14, CP64, CP89, CP90, CP91, CP92)
   'Modifie by Lydia 2017/11/14 +pa150
   'Modified by Lydia 2020/11/19 改變數
   'rt = ReadPatentDatabase(frm010001.intModifyKind, txtSystem, txtCode(0), _
          IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), pa05, _
          pa06, pa07, PA08, PA09, pa26, pa27, pa28, pa29, pa30, pa75, txtRecieveCode, _
          cp05, cp06, cp07, CP10, cp11, cp13, cp16, cp17, cp18, cp19, cp32, cp56, cu30, cp14, CP64, CP89, CP90, CP91, CP92, PA77, CP150, PA150)
   'Modified by Lydia 2021/04/15 +PA91案件備註
   'Modified by Sindy 2022/12/7 +PA178證書形式
   rt = ReadPatentDatabase(frm010001.intModifyKind, nPA01, nPA02, _
          nPA03, nPA04, pa05, _
          pa06, pa07, PA08, PA09, pa26, pa27, pa28, pa29, pa30, pa75, txtRecieveCode, _
          cp05, cp06, cp07, CP10, cp11, cp13, cp16, cp17, cp18, cp19, cp32, cp56, cu30, cp14, CP64, CP89, CP90, CP91, CP92, PA77, CP150, PA150, PA91, PA178)
          
   'NICK 900803 **********************
   txtCP64 = CP64
   '**********************
   If rt Then
      If frm010001.intModifyKind <> 0 Then
         txtPatent(0) = cp05
         txtPatent(1) = CP10
         txtPatent(2) = cp11
         txtPatent(15) = cp13
         txtPatent(16) = cu30
         txtPatent(17) = cp16
         txtPatent(18) = cp18
         txtPatent(20) = cp32
         txtPatent(21) = cp17
         txtPatent(22) = cp19
         txtPatent(24) = cp14
         CheckKeyIn 1
         CheckKeyIn 2
         CheckKeyIn 15
         If txtPatent(24).Visible Then
            CheckKeyIn 24
         End If
         If txtPatent(1) = 讓與 Or txtPatent(1).Text = 專利權讓與 Or txtPatent(1).Text = 合併 Or txtPatent(1).Text = 繼承 Then
            fraPatition.Visible = True
            txtPatent(23) = cp56
            CheckKeyIn 23
            txtPetitionx(2) = CP89
            txtPetitionx_Validate 2, False
            txtPetitionx(3) = CP90
            txtPetitionx_Validate 3, False
            txtPetitionx(4) = CP91
            txtPetitionx_Validate 4, False
            txtPetitionx(5) = CP92
            txtPetitionx_Validate 5, False
            If txtPatent(1).Text = 合併 Then
               Label29.Caption = "合併申請人1："
               For intI = 2 To 5
                  Label31(intI).Caption = "合併申請人" & intI & "："
               Next
            ElseIf txtPatent(1).Text = 繼承 Then
               Label29.Caption = "繼承申請人1："
               For intI = 2 To 5
                  Label31(intI).Caption = "繼承申請人" & intI & "："
               Next
            End If
            
         Else
            fraPatition.Visible = False
         End If
         'Add By Sindy 2012/11/08
         If CP150 = "Y" Then
            Me.Check2.Value = 1
         End If
         '2012/11/08 End
      End If
      txtPatent(14) = cp06
      '2013/2/7 add by sonia 20130208~20130318收文之美國發明維持費法定期限20130209-201309918之間若本限晚於3/11則改成3/11
      If frm010001.intModifyKind = 0 Then '新增
         If txtPatent(1) = "606" Then
            If Val(DBDATE(txtPatent(14))) >= 20130311 And Val(DBDATE(txtPatent(14))) <= 20130918 And Val(strSrvDate(1)) >= 20130208 And Val(strSrvDate(1)) <= 20130318 Then
               txtPatent(14) = 1020311
            End If
         End If
      End If
      '2013/2/7 end
      txtPatent(19) = cp07
      txtPatent(3) = PA08
      'Added by Lydia 2020/11/19 CFP英國脫歐案管制
      If nPA01 = "CFP" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" Then
           txtPatent(4) = "201"  '預設:英國
           m_CaseNa239(5) = pa05
           m_CaseNa239(6) = pa06
           m_CaseNa239(7) = pa07
      Else
      'end 2020/11/19
           txtPatent(4) = PA09
      End If 'Added by Lydia 2020/11/19
      txtPatent(5) = pa05
      txtPatent(6) = pa06
      txtPatent(7) = pa07
      txtPatent(8) = pa26
      txtPatent(9) = pa28
      txtPatent(10) = pa30
      txtPatent(11) = pa27
      txtPatent(12) = pa29
      txtPatent(13) = pa75
      'add by nickc 2007/03/27
      txtPatent(27) = PA77
        
      txtData(2).Text = PA150 'Added by Lydia 2017/11/14
      m_PA91 = PA91 'Added by Lydia 2021/04/15 +PA91案件備註
      txtPatent(29) = PA178 'Add by Sindy 2022/12/7 +PA178證書形式
      
      'Add By Cheng 2001/12/17
      '顯示智權人員代號
      'txtPatent(15) = cp13   '2011/5/11 cancel by sonia 偶而改智權人員收文會忘記打所以不自動帶
      'Modify By Cheng 2002/01/03
      If Len("" & txtPatent(8).Text) > 0 Then CheckKeyIn 8
      
      CheckKeyIn 9
      CheckKeyIn 10
      CheckKeyIn 11
      CheckKeyIn 12
      CheckKeyIn 13
      CheckKeyIn 3
      CheckKeyIn 4
      'Add By Cheng 2001/12/17
      '顯示智權人員姓名
      If txtPatent(15) <> "" Then CheckKeyIn 15

   Else
      If frm010001.intModifyKind <> 0 Then
         MsgBox "讀取資料時發生錯誤!!", vbCritical
         bolLeave = True
         Unload Me
      Else
         txtPatent(14) = cp06
         txtPatent(19) = cp07
         txtPatent(3) = PA08
         txtPatent(4) = PA09
         txtPatent(5) = pa05
         txtPatent(6) = pa06
         'Added by Lydia 2018/01/04
         txtPatent(5).Tag = pa05
         txtPatent(6).Tag = pa06
         'end 2018/01/04
         txtPatent(7) = pa07
         txtPatent(8) = pa26
         txtPatent(9) = pa28
         txtPatent(10) = pa30
         txtPatent(11) = pa27
         txtPatent(12) = pa29
         txtPatent(13) = pa75
         'add by nickc 2007/03/27
         txtPatent(27) = PA77
         CheckKeyIn 9
         CheckKeyIn 10
         CheckKeyIn 11
         CheckKeyIn 12
         CheckKeyIn 13
         If frm010001.intSaveMode <> 1 Then
            CheckKeyIn 8
            CheckKeyIn 3
            CheckKeyIn 4
         End If
      End If
   End If
   m_strCP06 = txtPatent(14) 'Add By Sindy 2021/4/29
   
   'NICK 900803 **********************
   If frm010001.intChoose = 1 Then
      txtPatent(2) = "90"
      CheckKeyIn (2)
   End If
   ' **********************
   
'Add by Amy 2013/06/26 FCP/P 新申請案101/102/103及衍生設計125 則抓取案件命名追蹤流水號
'Modified by Lydia 2018/09/06 改成變數
'If strSrvDate(2) >= Val("1020723") And (txtSystem = "FCP" Or txtSystem = "P") And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "125") Then
'Modified by Lydia 2020/11/19 改變數
'If (txtSystem = "FCP" Or txtSystem = "P") And InStr(AddTrackingNo, txtPatent(1)) > 0 Then
If (nPA01 = "FCP" Or nPA01 = "P") And InStr(AddTrackingNo, txtPatent(1)) > 0 Then
   'Modified by Lydia 2024/12/13
   'txtTCN01 = GetTCN01
   txtTCN01 = Pub_GetTCN01(txtRecieveCode)
End If
'end 2013/06/26

   ' 91.09.11 modify by louis
   'Modified by Lydia 2022/08/25改成共用模組
   'OnUpdateFee
   If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), txtPatent(1), txtPatent(15), _
        IIf(chkWebApp.Visible = True And chkWebApp.Value = 1, "Y", ""), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
          txtPatent(17) = m_NowCP16
          txtPatent(21) = m_NowCP17
          txtPatent(18) = m_NowCP18
   End If
   'end 2022/08/25
End Sub

'Add By Sindy 2010/10/28
Private Sub Combo3_Validate(Cancel As Boolean)
   If Combo3 <> "" Then
      'Modified by Lydia 2018/11/22 區分專利種類
      'Combo3 = Left(Combo3, 1) + "." + PUB_GetCaseAttributeName(Left(Combo3, 1))
      Combo3 = Left(Combo3, 1) + "." + PUB_GetCaseAttributeName(Left(Combo3, 1), txtPatent(3))
      If Combo3 = Left(Combo3, 1) + "." Then
         Combo3 = Left(Combo3, 1)
         Cancel = True
         Combo3.SetFocus
      End If
   End If
End Sub
'2010/10/28 End

'Add By Sindy 2010/3/8
Private Sub Command1_Click()
   '開啟聯絡人視窗
   frm010007_1.ReadPatent
   frm010007_1.txt1(0) = strPA51s
   frm010007_1.txt1(1) = strPA52s
   frm010007_1.txt1(2) = strPA53s
   frm010007_1.txt1(3) = strPA54s
   frm010007_1.txt1(4) = strPA55s
   frm010007_1.txt1(5) = strPA56s
   
   'Added by Lydia 2018/09/07 Elaine表示：為了方便核對資料，開放可查看；若要修改，則到維護作業
   If Left(frm010001.mRole, 1) = "F" Then
        frm010007_1.cmdok(0).Visible = False
        frm010007_1.txt1(0).Locked = True
        frm010007_1.txt1(1).Locked = True
        frm010007_1.txt1(2).Locked = True
        frm010007_1.txt1(3).Locked = True
        frm010007_1.txt1(4).Locked = True
        frm010007_1.txt1(5).Locked = True
   Else
        frm010007_1.cmdok(0).Visible = True
        frm010007_1.txt1(1).Locked = False
        frm010007_1.txt1(2).Locked = False
        frm010007_1.txt1(3).Locked = False
        frm010007_1.txt1(4).Locked = False
        frm010007_1.txt1(5).Locked = False
   End If
   'end 2018/09/07
   frm010007_1.Show vbModal
   bolCancel = frm010007_1.bolOK
   If bolCancel = True Then
      strPA51s = frm010007_1.strPA51s
      strPA52s = frm010007_1.strPA52s
      strPA53s = frm010007_1.strPA53s
      strPA54s = frm010007_1.strPA54s
      strPA55s = frm010007_1.strPA55s
      strPA56s = frm010007_1.strPA56s
   End If
   Unload frm010007_1
   Set frm010007_1 = Nothing
   'Modified by Lydia 2017/11/14
   'txtPatent(17).SetFocus 'Add By Sindy 2010/3/19
   'Modify by Amy 2021/12/16 +txtPatent(17).Enabled=True 修正本就有的bug 收文->接洽單修改->進入此畫面->按「聯絡人資料」->再回來此頁 會error
   If txtPatent(17).Visible = True And txtPatent(17).Enabled = True Then txtPatent(17).SetFocus
End Sub

Private Sub Form_Load()
Dim oLbl As LABEL
Dim oTxt As TextBox 'Added by Lydia 2017/11/14

   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   If frm010001.intChoose = 1 Then
      txtPatent(20) = "N"
      fraPromoter.Visible = True
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
   For Each oLbl In lblPetition
      oLbl.BackColor = &H8000000F
   Next
   'end 2008/8/5
   
   'Added by Lydia 2017/11/14 FCP案件命名電子化
   For Each oTxt In txtData
      oTxt.Text = ""
      oTxt.Tag = ""
   Next
   ChkExpDate.Value = 0
   'end 2017/11/14

   If frm010001.mRole = "" Then 'Added by Lydia 2018/09/05 排除外專後續案收文
        'Added by Lydia 2018/03/01 暫存下載的檔案
        'Added by Lydia 2021/07/13 櫃台人員居家上班發生找不到路徑問題；調整當收文新案為FCP案或FMP案時，才需要建立TrackingNO暫存區。
        If frm010001.intModifyKind = 0 And frm010001.txtCode(0) = "" And ((frm010001.txtSystem = "FCP" And InStr(FcpAddTct & ",307", frm010001.txtCaseProperty) > 0) _
                                Or (frm010001.txtSystem = "P" And frm010001.txtFMP = "Y" And InStr(FcpAddTct & ",307", frm010001.txtCaseProperty) > 0)) Then
              Call Pub_ChkExcelPath(App.path & "\" & strUserNum)  '先檢查個人資料夾
        'end 2021/07/13
             'Modified by Lydia 2021/06/18 配合遠端桌面,改到個人
             'mSaveDir = App.path & "\暫存區"
             mSaveDir = App.path & "\" & strUserNum & "\暫存區"
             'end 2018/10/12
             If Dir(mSaveDir, vbDirectory) = "" Then
                  MkDir mSaveDir
             End If
             'end 2018/03/01
        End If 'Added by Lydia 2021/07/13
   End If
    
   'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
   Label32.Visible = False
   txtPatent(28).Visible = False
   Check1.Visible = False
   
   'Add By Sindy 2022/6/29
   m_strIR01 = frm010001.m_strIR01
   m_strIR02 = frm010001.m_strIR02
   m_strIR03 = frm010001.m_strIR03
   m_strIR04 = frm010001.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2022/6/29 END
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
   If bolActive Then Exit Sub 'Add by Morgan 2004/4/15
   
Dim strPKindName As String, strDate1 As String, strDate2 As String, strCode(5) As String, i As Integer

   Me.Refresh
   
   '根據intModifyMode來調整fraWindow1 , fraWindow2
   Select Case frm010001.intModifyKind
             Case 0
                        '新增：所有欄位皆可輸入
                        fraWindow1.Enabled = True
                        Select Case frm010001.intSaveMode
                                     Case 0
                                                fraWindow2.Enabled = False
                                                cmdInventor.Enabled = False
                                     Case 1
                                                fraWindow2.Enabled = True
                                                Dim intWhere As Integer
                                                
                                                'Added by Lydia 2021/11/10 改在前畫面輸入申請國家; 寰華和FMP大陸案衍生的香港、澳門案不走命名之流程
                                                'Modified by Lydia 2024/02/20 有關" 香港013專利開放收文集體設計申請105"，請比照" 香港013專利收文設計申請103"
                                                'If (txtSystem = "P" Or txtSystem = "FCP") And txtCode(0) = "" And frm010001.txtFMP = "Y" And frm010001.txtNA01 <> "" Then
                                                If (txtSystem = "P" Or txtSystem = "FCP") And frm010001.m_blnNewCase = True And frm010001.txtFMP = "Y" And frm010001.txtNA01 <> "" Then
                                                    txtPatent(4) = frm010001.txtNA01
                                                    CheckKeyIn 4
                                                Else
                                                'end 2021/11/10
                                                    'edit by nickc 2007/02/02 不用 dll 了
                                                    'If objPublicData.GetSystemKind(txtSystem.Text, , , intWhere) Then
                                                    If ClsPDGetSystemKind(txtSystem.Text, , , intWhere) Then
                                                       If intWhere <> 國外_CF Then
                                                          txtPatent(4) = 台灣國家代號
                                                          CheckKeyIn 4
                                                       End If
                                                    End If
                                                End If 'Added by Lydia 2021/11/10
                                                Dim strTemp As String
                                                'edit by nickc 2007/02/06 不用 dll 了
                                                'obj001.SetPatentProperty txtPatent(1), strTemp
                                                Cls001SetPatentProperty txtPatent(1), strTemp
                                                If strTemp <> "" Then
                                                   txtPatent(3) = strTemp
                                                   CheckKeyIn 3
                                                End If
                        End Select
                        If LastDate = "" Then
                           txtPatent(0).Text = GetTaiwanTodayDate
                        Else
                           txtPatent(0).Text = LastDate
                        End If
                        txtPatent_GotFocus 0
                        If txtSystem = "FCP" And frm010001.intChoose = 1 Then CheckKeyIn 24
             Case 1
                        '修改：中間欄位不可輸入
                        fraWindow1.Enabled = True
                        Dim bolNew As Boolean
                        'edit by nickc 2007/02/06 不用 dll 了
                        'If obj001.IsNewCase(txtRecieveCode, bolNew) Then
                        If Cls001IsNewCase(txtRecieveCode, bolNew) Then
                           If bolNew Then
                              fraWindow2.Enabled = True
                              cmdInventor.Enabled = True
                           Else
                              fraWindow2.Enabled = False
                              cmdInventor.Enabled = False
                           End If
                        Else
                           bolLeave = True
                           Unload Me
                           Exit Sub
                        End If
                        'Add by Amy 2013/06/26修改時追蹤流水號不可修改
                        txtTCN01.Enabled = False
             Case 2
                        '刪除：所有欄位皆不可輸入
                        cmdok(0).Visible = False
                        fraWindow1.Enabled = False
                        fraWindow2.Enabled = False
   End Select
   
   'Add By Sindy 2010/3/8 預設值
   strPA51s = "": strPA52s = "": strPA53s = ""
   strPA54s = "": strPA55s = "": strPA56s = ""
   bolCancel = False
   '2010/3/8 End
   
   ReDim m_CaseNa239(1 To TF_PA) 'Move by Lydia 2020/11/29 從CFP英國脫歐案管制內移上來
   
   If frm010001.intModifyKind <> 0 Or frm010001.intSaveMode <> 1 Then
      ReadPatentDatabaseR
   End If
 
   'Added by Lydia 2020/11/19 CFP英國脫歐案管制：新增英國脫歐案,先讀取歐盟案
   'Modified by Lydia 2020/12/01 判斷新案
   'If txtSystem = "CFP" And frm010001.txtCaseNa239 <> "" Then
   If txtSystem = "CFP" And frm010001.txtCode(0) & frm010001.txtCode(1) & frm010001.txtCode(2) = "" And frm010001.txtCaseNa239 <> "" Then
       Call ChgCaseNo(frm010001.txtCaseNa239.Text, m_CaseNa239)
       If m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" Then
           ReadPatentDatabaseR
       End If
   End If
   'end 2020/11/19
   
   'Add By Sindy 2021/4/29 主管機關期限
   CheckOC3
   m_strCPM34 = ""
   strSql = "select cpm34 from casepropertymap where cpm01='" & txtSystem & "' and cpm02='" & txtPatent(1) & "'"
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount > 0 Then
      m_strCPM34 = "" & AdoRecordSet3.Fields(0)
   End If
   '2021/4/29 END
   'Add By Sindy 2022/3/22 是否為FMP案件
   If PUB_ChkIsFMP(txtSystem, txtCode(0), txtCode(1), txtCode(2)) = True Or _
      (txtSystem = "P" And (frm010001.txtFMP = "Y" And frm010001.txtNA01 = "020") And InStr(NewCasePtyList, txtPatent(1)) > 0) Then
      m_bolFMP = True
   Else
      m_bolFMP = False
   End If
   '2022/3/22 END
   
   'Added by Lydia 2023/04/13 (舊案)是否為寰華案
   If m_bolFMP = True And txtSystem = "P" And txtCode(0) <> "" Then
      If PUB_FMPtoCheck(1, 2, Pub_strUserST05, txtSystem, txtCode(0), txtCode(1), txtCode(2)) = True Then
          m_bolFMP2 = True
      End If
   End If
   'end 2023/04/13
   
   'Added by Lydia 2018/11/22 設定案件屬性
   'Add by Lydia 2018/12/19 若為新案,預設為發明案屬性 (ex.收P案變更401,在輸入專案種類後,案件屬性被清空)
   If txtPatent(3) = "" Then
      Call SetCombo3("1")
   Else
   'end 2018/12/19
      Call SetCombo3(txtPatent(3))
   End If 'end 2018/12/19
   
   '2010/12/31 END
   'Add by Morgan 2004/5/28
   '專利修法：減免退費、退費時，費用、規費及點數不可輸入
   If (txtSystem = "P" Or txtSystem = "FCP") And (txtPatent(4) = "000") And (txtPatent(1) = "908" Or txtPatent(1) = "919") Then
      '2010/3/23 MODIFY BY SONIA P有預繳99年年費者退費可收500(0.5)
      'txtPatent(17).Enabled = False: txtPatent(18).Enabled = False: txtPatent(21).Enabled = False
      If Not (txtSystem = "P" And txtPatent(4) = "000" And txtPatent(1) = "908") Then
         txtPatent(17).Enabled = False: txtPatent(18).Enabled = False: txtPatent(21).Enabled = False
      Else
         'Modify by Morgan 2011/6/17 改呼叫共用函數判斷
         'strSQL = "select T14 from T99 where t01='" & txtSystem & "' and t02='" & txtCode(0) & "' and t03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and t04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
         'CheckOC3
         'AdoRecordSet3.CursorLocation = adUseClient
         'AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
         'If AdoRecordSet3.RecordCount <> 0 Then
         '   If Not IsNull(AdoRecordSet3.Fields("T14")) Then
         '      txtPatent(17).Enabled = False: txtPatent(18).Enabled = False: txtPatent(21).Enabled = False
         '   End If
         'Else
         '   txtPatent(17).Enabled = False: txtPatent(18).Enabled = False: txtPatent(21).Enabled = False
         'End If
         'CheckOC3
'2012/6/25 modify by sonia P-094915 郭雅娟說只要智權人員收得到P的退費就可收費,但不可收規費
'         strExc(1) = txtSystem
'         strExc(2) = txtCode(0)
'         strExc(3) = IIf(txtCode(1) = "", "0", txtCode(1))
'         strExc(4) = IIf(txtCode(2) = "", "00", txtCode(2))
'         If PUB_ChkRefund(strExc()) = False Then
'            txtPatent(17).Enabled = False: txtPatent(18).Enabled = False: txtPatent(21).Enabled = False
'         End If
         txtPatent(21).Enabled = False
'2012/6/25 END
         'end 2011/6/17
      End If
      '2010/3/23 END
   End If
   
   
   'Added by Morgan 2012/4/25
   If txtSystem = "P" And txtPatent(4) = "000" And txtPatent(1) = "405" Then
      lblCopy.Visible = True
      txtCopy.Visible = True
      If txtPatent(21) = "" Then txtCopy = "1" 'Added by Morgan 2012/6/20
   Else
      lblCopy.Visible = False
      txtCopy.Visible = False
   End If
   'end 2012/4/25
   
   'Added by Morgan 2012/6/20
   '台灣新申請案電子送件可減免 600
   'Modified by Morgan 2014/1/20 +307,125可電子送件
   'Modified by Morgan 2014/11/10 +分割及改請電子送件也可減免--郭
   If txtPatent(4) = "000" And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "307" Or txtPatent(1) = "125" Or Left(txtPatent(1), 1) = "3") Then
      chkWebApp.Visible = True
   Else
      chkWebApp.Visible = False
   End If
   'end 2012/6/20
   
   'Add by  Morgan 2008/8/5
   '2013/8/19 MODIFY BY SONIA 加入分割307
   If (txtPatent(1) = "101" Or txtPatent(1) = "307") And txtPatent(4) = "000" Then
      chkEnglish.Visible = True
      'Modified by Morgan 2012/6/20 可同時有電子送件減免
      'If txtPatent(21) = "2700" Then
      If txtPatent(21) = "2700" Or txtPatent(21) = "2100" Then
         chkEnglish.Value = 1
         chkEnglish.Caption = "附英文摘要"  '2009/2/10 add by sonia 同時申請三國(含)以上之美日德可多5點
      End If
   '2009/2/10 add by sonia 同時申請三國(含)以上之美日德可多5點
   ElseIf (txtPatent(4) = "011" Or txtPatent(4) = "101" Or txtPatent(4) = "231") And InStr(CaseMapIn, txtPatent(1)) > 0 Then
      chkEnglish.Visible = True
      chkEnglish.Caption = "同時申請三國(含)以上之美日德"
   '2009/2/10 end
   Else
      chkEnglish.Visible = False
      chkEnglish.Caption = ""
   End If
   'end 2008/8/5
   
   'Add by Amy 2013/06/26
   'Modified by Lydia 2018/05/15 P案非FMP案才顯示
   'If strSrvDate(2) >= Val("1020723") And (txtSystem = "FCP" Or txtSystem = "P") And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "125") Then
   'Modified by Lydia 2018/05/23 香港案標準專利記錄請求(110)不走命名流程, 仍要輸入Tracking_no
   'If (txtSystem = "FCP" And InStr(FcpAddTct, txtPatent(1)) > 0) Or (txtSystem = "P" And InStr(FcpAddTct, txtPatent(1)) > 0 And (txtPatent(4) <> "000" Or frm010001.txtFMP.Text = "Y")) Then
   'Modified by Lydia 2018/09/06 改成變數
   'If (txtSystem = "FCP" And InStr(FcpAddTct, txtPatent(1)) > 0) Or (txtSystem = "P" And InStr(FcpAddTct, txtPatent(1)) > 0 And (txtPatent(4) <> "000" Or frm010001.txtFMP.Text = "Y")) _
       Or (txtSystem = "P" And txtPatent(1) = "110" And txtCode(0) = "") Then
    '只有FCP及P的新申請案及衍生設計 才出現案件命名流水號
    'Modified by Lydia 2023/05/12 +限新案+ txtCode(0) = "" and ()
   'Modified by Lydia 2024/02/20 調整新案的判斷: 有關" 香港013專利開放收文集體設計申請105"，請比照" 香港013專利收文設計申請103"
   'If txtCode(0) = "" And ((txtSystem = "FCP" And InStr(AddTrackingNo, txtPatent(1)) > 0) Or (txtSystem = "P" And InStr(AddTrackingNo, txtPatent(1)) > 0 And (txtPatent(4) <> "000" Or frm010001.txtFMP.Text = "Y"))) Then
   If (txtCode(0) = "" And txtSystem = "FCP" And InStr(AddTrackingNo, txtPatent(1)) > 0) Or _
      (txtSystem = "P" And InStr(AddTrackingNo, txtPatent(1)) > 0 And frm010001.m_blnNewCase = True And frm010001.txtNA01 <> "000" And frm010001.txtFMP.Text = "Y") Then
        LbTracking.Visible = True: DoEvents
        txtTCN01.Visible = True: DoEvents
   End If
   'end 2013/06/26
   
   bolMoveOK = False 'Added by Lydia 2020/02/13
   
   'Added by Lydia 2017/11/14 FCP案件命名電子化：中說輸入相關設定
   bolExistTCT = False
   ChkExpDate.Value = 0: ChkAdd924.Value = 0: ChkAdd416.Value = 0
   ChkAdd902.Value = 0 'Added by Lydia 2018/04/13
   ChkAdd203.Value = 0 'Added by Lydia 2018/04/12
   'Added by Lydia 2018/05/07 FMP案-收文
   ChkAdd414.Value = 0
   ChkAdd938.Value = 0
   ChkAdd939.Value = 0
   ChkAdd106.Value = 0
   ChkAdd228.Value = 0
   'end 2018/05/07
   'Added by Lydia 2018/09/06 分割案-收文
   ChkAdd435.Value = 0
   ChkAdd435.Left = ChkAdd414.Left
   'Added by Lydia 2021/08/27 回復說明書校閱968
   ChkAdd968.Value = 0
   ChkAdd968.Top = ChkAdd414.Top
   
   '新申請案才顯示設定
   'Modified by Lydia 2018/03/05 先上101~103
   'If strSrvDate(1) >= FCP案件命名啟用日 And txtSystem = "FCP" And InStr(NewCasePtyList, txtPatent(1)) > 0 Then
   'Modified by Lydia 2018/04/17 +125 衍生設計案
   'Modified by Lydia 2018/05/07 +FMP案-收文
   'If strSrvDate(1) >= FCP案件命名啟用日 And txtSystem = "FCP" And InStr("101,102,103,125", txtPatent(1)) > 0 Then
   'Modified by Lydia 2018/09/06 +分割案307
   'Modified by Lydia 2019/06/18 P-122972(AA8026783) 要修改急件時間
   'If ((txtSystem = "FCP" And InStr(FcpAddTct & ",307", txtPatent(1)) > 0) _
                                Or (txtSystem = "P" And frm010001.txtFMP = "Y" And InStr(FcpAddTct & ",307", txtPatent(1)) > 0)) Then
  'Modified by Lydia 2021/11/10 判斷申請國家; 寰華和FMP大陸案衍生的香港、澳門案不走命名之流程
  'If ((txtSystem = "FCP" And InStr(FcpAddTct & ",307", txtPatent(1)) > 0) _
                                Or (txtSystem = "P" And (frm010001.txtFMP = "Y" Or Me.txtTCN01.Text <> "") And InStr(FcpAddTct & ",307", txtPatent(1)) > 0)) Then
  'Modified by Lydia 2024/02/01 (P-133037) Phoebe詢問工程師後結論是：FMP案之香港專利申請案101,102,103（非衍申案），需要跑命名流程，故之後凡是此類案件皆需跑命名流程
  'If ((txtSystem = "FCP" And InStr(FcpAddTct & ",307", txtPatent(1)) > 0) _
                                Or (txtSystem = "P" And (frm010001.txtFMP = "Y" And frm010001.txtNA01 = "020") And InStr(FcpAddTct & ",307", txtPatent(1)) > 0)) Then
  If ((txtSystem = "FCP" And InStr(FcpAddTct & ",307", txtPatent(1)) > 0) _
                                Or (txtSystem = "P" And (frm010001.txtFMP = "Y" And frm010001.txtNA01 = "020") And InStr(FcpAddTct & ",307", txtPatent(1)) > 0) _
                                Or (txtSystem = "P" And (frm010001.txtFMP = "Y" And frm010001.txtNA01 = "013") And InStr(FcpAddTct, txtPatent(1)) > 0)) Then
      '隱藏智權人員以下,不用之欄位
      Label30.Visible = False: txtPatent(26).Visible = False '分所案號
      Label21.Visible = False: txtPatent(16).Visible = False '郵遞區號
      Label22.Visible = False: txtPatent(17).Visible = False '費用
      Label23.Visible = False: txtPatent(18).Visible = False '點數
      Label19.Visible = False: txtPatent(20).Visible = False '是否開電腦收據
      Label26.Visible = False: txtPatent(21).Visible = False '規費
      chkEnglish.Visible = False '附英文摘要 / 同時申請三國(含)以上之美日德
      Label32.Visible = False: txtPatent(28).Visible = False '預定收款日
      Check1.Visible = False '現金或支票
      Label27.Visible = False: txtPatent(22).Visible = False '後金
      Check2.Visible = False '有★★的應收帳款簽核控管
      
      '移動位置和順序
      fraPromoter.Visible = False
      fraPatition.Visible = False
      'Remove by Lydia 2018/05/07 改到下面
      'LbTracking.Top = 5300
      'txtTCN01.Top = 5300
      'txtTCN01.TabIndex = 50
      'end 2018/05/07
      chkWebApp.Left = Label18.Left + 300
      chkWebApp.Left = Label21.Top
      chkWebApp.TabIndex = 45
      ChkAdd416.TabIndex = 46
      ChkAdd203.TabIndex = 47
      ChkAdd902.TabIndex = 48
      ChkAdd924.TabIndex = 49
      Command1.Top = txtPatent(15).Top
            
      'Added by Lydia 2018/05/07 FMP案-收文
      ChkAdd414.TabIndex = 50
      ChkAdd938.TabIndex = 51
      ChkAdd939.TabIndex = 52
      ChkAdd106.TabIndex = 53
      ChkAdd228.TabIndex = 54
      ChkAdd435.TabIndex = 55 'Added by Lydia 2018/09/06
      txtData(3).TabIndex = 56
      LbTracking.Top = 5600
      txtTCN01.Top = 5600
      txtTCN01.TabIndex = 57
      'end 2018/05/07
      '新申請案才可設定
      If frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 Then
         ChkAdd924.Visible = True
         ChkAdd416.Visible = True
         ChkAdd203.Visible = True 'Added by Lydia 2018/04/12
         ChkAdd902.Visible = True 'Added by Lydia 2018/04/13
         'Added by Lydia 2021/08/27 回復說明書校閱968
         If txtSystem = "FCP" Then ChkAdd968.Visible = True
         
         'Added by Lydia 2018/05/07 FMP案-收文
         'Modified by Lydia 2018/09/06 排除分割案
         'If txtSystem = "P" Then
         'Modified by Lydia 2018/09/26 Elaine: P分割案和一般發明選項相同
         'If txtSystem = "P" And txtPatent(1) <> "307" Then
         If txtSystem = "P" Then
            'Modified by Lydia 2024/02/01 P案名稱不同: 製作中說210>>撰稿210
            Label38.Caption = "(1.翻譯中說201　2.檢視中說209　3.撰稿210　" 'Added by Lydia 2022/10/07 P案無外文提申本242
            Label39.Caption = "5.核對中說235    6.檢視PCT公開本與FCP相異處942)"
            chkWebApp.Visible = False
            ChkAdd414.Visible = True
            ChkAdd938.Visible = True
            ChkAdd939.Visible = True
            ChkAdd106.Visible = True
            ChkAdd228.Visible = True
         Else
            Label38.Caption = "(1.翻譯中說201　2.檢視中說209　3.製作中說210　4.製作中說210＆外文提申本242" 'Added by Lydia 2022/10/07
            Label39.Caption = "5.核對中說235)"
            ChkAdd414.Visible = False
            ChkAdd938.Visible = False
            ChkAdd939.Visible = False
            ChkAdd106.Visible = False
            ChkAdd228.Visible = False
            'Added by Lydia 2018/09/06
            If txtSystem = "P" Then chkWebApp.Visible = False
         End If
         'end 2018/05/07
         
         'Added by Lydia 2018/09/06 分割案不顯示其他項目
         If txtPatent(1) = "307" Then
                Label36.Visible = False: Label38.Visible = False: Label39.Visible = False
                txtData(3).Visible = False
                txtData(2).Text = "B"
                If txtSystem = "FCP" Then ChkAdd435.Visible = True
         Else
         'end 2018/09/06
                Label36.Visible = True: Label38.Visible = True: Label39.Visible = True
                txtData(3).Visible = True
                txtData(2).Text = "B" 'Added by Lydia 2018/03/31 命名作業人工流程變更(3/27),所有新案在櫃台收文時一律輸入"退程序",退Gill後做分案作業,再退程序人員輸入指定送件期限後,再到新案建檔設定工程師組別。
                txtPatent(5).Text = "待命名" 'Added by Lydia 2017/12/20 預設中文名稱
         End If 'end 2018/09/06
         txtPatent(0).SetFocus
      End If
      
      fraTCT.Left = Label21.Left
      fraTCT.Top = Label21.Top
      fraTCT.BackColor = &H8000000F
      fraTCT.Visible = True
      Call SetTCTdata
      If frm010001.intModifyKind = 2 Then fraTCT.Enabled = False
      'Added by Lydia 2018/09/06 不顯示急件
      If InStr(FcpAddTct, txtPatent(1)) = 0 Then
          ChkExpDate.Visible = False: txtData(0).Visible = False: txtData(1).Visible = False
      End If
      'end 2018/09/06
   
   End If
   'end 2017/11/14
   
    'Added by Lydia 2018/05/07FMP寰華案件後續收文時,提醒收文人員, 接洽單交外專(從frm010001移來)
    'Modified by Lydia 2022/06/21 排除外專後續案收文
    'If txtCode(0) <> "" Then
    If txtCode(0) <> "" And Left(frm010001.mRole, 2) <> "F2" Then
       If PUB_FMPtoCheck(1, 2, "", txtSystem, txtCode(0), txtCode(1), txtCode(2)) = True Then
          MsgBox "此案號為FMP寰華案件, 收文後請將文件交 外專！", , "注意！"
       End If
    End If
    'end 2018/05/08
   
   'Added by Lydia 2020/05/20 法律所案源收文：讀取法務案源檔
   If strSrvDate(1) >= 法律所案源收文啟用日 And frm010001.intModifyKind = 0 And txtPatent(4) = "000" And (txtSystem = "FCP" Or txtSystem = "P") Then
      Call ReadLOS
   End If
   'end 2020/05/20
                        
   'Add by Amy 2021/12/20 改form2.0 TopIndex 有問題,因判斷複雜,故先寫死
   If UCase(App.EXEName) = "TEWRITER" Or UCase(App.EXEName) = "WRITER" Then
       'Modify by Amy 2022/04/15 +txtPatent(2).Enabled = True
       If txtPatent(0) <> MsgText(601) And txtPatent(1) <> MsgText(601) And txtPatent(2).Enabled = True Then
           txtPatent(2).SetFocus
       End If
   End If
      
   'Added by Morgan 2020/4/9 '批次收文
   If m_bBatch = True Then
      txtPatent(2) = m_CP11
      txtPatent(15) = m_CP13
      txtPatent_Validate 15, False
      cmdok(0).Value = True
   End If
   'end 2020/4/9
   
   'Added by Lydia 2022/09/05
   If strSrvDate(1) >= 收文存檔模組化啟用日 Then
       Call SetDBArray(True, txtRecieveCode, txtSystem, txtCode(0), txtCode(1), txtCode(2))
   End If
   
   'Added by Lydia 2022/12/14 鑑於FCP案已無紙本送件，且櫃台人員收文時會漏勾電子送件，故請將FCP案櫃台收文之電子送件欄位預設打勾
   If txtSystem = "FCP" And frm010001.mRole = "" And m_strIR01 = "" And chkWebApp.Visible = True Then
      If frm010001.m_blnNewCase = True Then
         chkWebApp.Value = 1
         chkWebApp.Enabled = False
      End If
   End If
   'end 2022/12/14
   
   bolActive = True 'Add by Morgan 2004/4/15
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Where01ToGo intLeaveKind

   intLeaveKind = 0
   PUB_SendMailCache 'Added by Lydia 2018/03/01
   
   'Added by Lydia 2020/02/13 English_Vers檔案：判斷啟用日
On Error Resume Next

   'Modified by Lydia 2021/07/13 +判斷有建資料夾
   If strSrvDate(1) >= XY特殊權限啟用日by檔案 And mSaveDir <> "" Then
       If bolMoveOK = True Then  'TrackingNO是否已搬檔完成(True無問題)，若有問題則TrackingNO和本機端的資料夾不刪除
           Call PUB_KillAnyFile(mSaveDir)
           RmDir mSaveDir  '移除資料夾
       End If
   End If
   'end 2020/02/13
   
   'Add By Sindy 2022/6/29
   If m_strIR01 <> "" Then
      If Not m_PrevForm Is Nothing Then
         Set m_PrevForm = Nothing
      End If
   End If
   If PUB_CheckFormExist("frmpic002") = True Then Unload frmpic002 'Add By Sindy 2022/7/11
   '2022/6/29 END
   
   'Add By Cheng 2002/07/18
   'Modify by Amy 2021/12/17 改Form2.0後,存檔按Enter會當掉,改在呼叫時清除記憶體變數
   'Set frm010005 = Nothing
   stChkForm = Me.Name 'Add by Amy 2021/12/21
End Sub

'Add By Sindy 2009/07/06
Private Sub textYear_GotFocus()
   InverseTextBox textYear
End Sub
Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub
Private Sub textYear_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub
'2009/07/06 End

Private Sub txtCopy_GotFocus()
   TextInverse txtCopy
   CloseIme
End Sub

Private Sub txtCopy_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And (KeyAscii > 57 Or KeyAscii < 48) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtPatent_Change(Index As Integer)
Dim strTemp As String, bolIsChina As Boolean  '2010/3/25 add by sonia
   
   Select Case Index
             Case 2
                        lblCaseSource.Caption = ""
             Case 3
                        lblTrademarkKind = ""
             Case 4
                        lblNation.Caption = ""
                        'Add by Morgan 2008/8/5
                        '2013/8/19 MODIFY BY SONIA 加入分割307
                        If (txtPatent(1) = "101" Or txtPatent(1) = "307") And txtPatent(4) = "000" Then
                           chkEnglish.Visible = True
                        '2009/2/10 add by sonia 同時申請三國(含)以上之美日德可多5點
                           chkEnglish.Caption = "附英文摘要"
                        ElseIf (txtPatent(4) = "011" Or txtPatent(4) = "101" Or txtPatent(4) = "231") And InStr(CaseMapIn, txtPatent(1)) > 0 Then
                           chkEnglish.Visible = True
                           chkEnglish.Caption = "同時申請三國(含)以上之美日德"
                        '2009/2/10 end
                        Else
                           chkEnglish.Visible = False
                           chkEnglish.Caption = ""
                        End If
                        'end 2008/8/5
                        '2010/3/25 add by sonia 因修改時不會依國家帶案件性質名稱故加入此段
                        'If txtPatent(Index) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
                        If txtPatent(Index) <> 台灣國家代號 Then bolIsChina = True Else bolIsChina = False
                        If ClsPDGetCaseProperty(txtSystem, txtPatent(1), strTemp, bolIsChina) Then
                           lblCaseProperty = strTemp
                        End If
                        '2010/3/25 end
                        
                        
                        'Added by Morgan 2012/6/20
                        '台灣新申請案電子送件可減免 600
                        'Modified by Morgan 2014/1/20 +307,125可電子送件
                        'Modified by Morgan 2014/11/10 +分割及改請電子送件也可減免--郭
                        If txtPatent(4) = "000" And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "307" Or txtPatent(1) = "125" Or Left(txtPatent(1), 1) = "3") Then
                           chkWebApp.Visible = True
                        Else
                           chkWebApp.Visible = False
                           chkWebApp.Value = 0
                        End If
                        'end 2012/6/20
                        
                        'Added by Morgan 2012/6/20
                        If txtSystem = "P" And txtPatent(4) = "000" And txtPatent(1) = "405" Then
                           lblCopy.Visible = True
                           txtCopy.Visible = True
                           If txtCopy = "" Then txtCopy = "1"
                        Else
                           lblCopy.Visible = False
                           txtCopy.Visible = False
                        End If
                        'end 2012/6/20

             Case 8, 9, 10, 11, 12
                        lblPetition(Index - 8).Caption = ""
                        If Index = 8 Then txtPatent(16).Text = ""
             Case 13
                        lblAgent.Caption = ""
             Case 15
                        lblSales.Caption = ""
                        lblDepartment = ""
                        m_SalesST15 = "" 'Added by Lydia 2018/09/06
                        m_SalesST06 = "" 'Added by Lydia 2019/09/16
             Case 23
                        lblPetitionName = ""
             Case 24
                        lblPromoter = ""
   End Select
End Sub
Private Sub txtPatent_Validate(Index As Integer, Cancel As Boolean)

   Select Case Index
            'add by nick 2004/12/08 當收文業務區與客戶檔業務區不同時發 mail  及提示
            Case 15
                    'add by nick 2005/01/04
                    If txtPatent(Index).Text <> "" And txtPatent(Index) < "63001" Then
                         MsgBox "智權人員不可小於 63001！", , "注意！"
                         Cancel = True
                         Exit Sub
                    End If
                    'add by nick 2004/12/08 因為之前的 智權人員並沒有抓
                    Dim strTemp As String, strTemp1 As String
                    'edit by nickc 2007/02/02 不用 dll 了
                    'If Not objPublicData.GetStaff(txtPatent(15).Text, strTemp, strTemp1) Then
                    If Not ClsPDGetStaff(txtPatent(15).Text, strTemp, strTemp1) Then
                        Cancel = True
                        Exit Sub
                    End If
                    'add by nickc 2006/11/02
                    'Modified by Lydia 2019/02/14
                    'GetST15 txtPatent(15).Text, strTemp1
                    'Modified by Lydia 2019/09/16
                    'm_SalesST15 = GetST15(txtPatent(15).Text, strTemp1)
                    m_SalesST15 = GetST15(txtPatent(15).Text, strTemp1, , m_SalesST06)
                    
                    lblSales.Caption = strTemp
                    lblDepartment = strTemp1
                    'm_SalesST15 = GetST15(txtPatent(15).Text) 'Added by Lydia 2018/09/06 'Mark by Lydia 2019/02/14
                    'Added by Lydia 2019/02/14 創新業務部人員收文控管
                    If PUB_ChkIsT10T20("2", txtPatent(15).Text, m_Tuser, strTemp) = True Then
                        txtPatent(15) = m_Tuser
                        lblSales.Caption = strTemp
                        txtPatent(15).SetFocus
                        Call txtPatent_GotFocus(15)
                        Cancel = True
                        Exit Sub
                    End If
                    'end 2019/02/14
                    Dim oStrCuSales1 As String
                    Dim oStrCuSales2 As String
                    Dim oStrCuSales3 As String
                    Dim oStrCuSales4 As String
                    Dim oStrCuSales5 As String
                    Dim oMailCount As String
                    'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    Dim IsMail  As Boolean
                    IsMail = True
                    oStrCuSales1 = ""
                    oStrCuSales2 = ""
                    oStrCuSales3 = ""
                    oStrCuSales4 = ""
                    oStrCuSales5 = ""
                    oMailCount = ""
                    ''Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
                    'If GetST15(txtPatent(15).Text) <> GetCuSales(ChangeCustomerL(txtPatent(8).Text), oStrCuSales1) And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(8).Text) <> "" Then
                    'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuAre
                    If ChkSameCuArea(Trim(txtPatent(8)), Trim(txtPatent(15)), , , , , Trim(txtPatent(13))) = False And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(8).Text) <> "" Then
                    'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    Else
                         If Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(8).Text) <> "" Then
                             IsMail = False
                         End If
                    End If
                    'Added by Lydia 2019/09/16 檢查是否為待活化客戶
                    If m_SalesST06 <> "" And Trim(txtPatent(8)) <> "" And Trim(txtPatent(15)) <> "" Then
                        If PUB_ChkOldCustomer(False, txtPatent(8), Trim(txtPatent(15)), m_SalesST15, m_SalesST06) = True Then
                            IsMail = False
                        End If
                    End If
          
                    'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
                    'If GetST15(txtPatent(15).Text) <> GetCuSales(ChangeCustomerL(txtPatent(9).Text), oStrCuSales2) And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(9).Text) <> "" Then
                    'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuAre
                    If ChkSameCuArea(Trim(txtPatent(9)), Trim(txtPatent(15)), , , , , Trim(txtPatent(13))) = False And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(9).Text) <> "" Then
                    'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    Else
                         If Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(9).Text) <> "" Then
                             IsMail = False
                         End If
                    End If
                    'Added by Lydia 2019/09/16 檢查是否為待活化客戶
                    If m_SalesST06 <> "" And Trim(txtPatent(9)) <> "" And Trim(txtPatent(15)) <> "" Then
                        If PUB_ChkOldCustomer(False, txtPatent(9), Trim(txtPatent(15)), m_SalesST15, m_SalesST06) = True Then
                            IsMail = False
                        End If
                    End If
                    
                    'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
                    'If GetST15(txtPatent(15).Text) <> GetCuSales(ChangeCustomerL(txtPatent(10).Text), oStrCuSales3) And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(10).Text) <> "" Then
                    'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuAre
                    If ChkSameCuArea(Trim(txtPatent(10)), Trim(txtPatent(15)), , , , , Trim(txtPatent(13))) = False And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(10).Text) <> "" Then
                    'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    Else
                         If Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(10).Text) <> "" Then
                             IsMail = False
                         End If
                    End If
                    'Added by Lydia 2019/09/16 檢查是否為待活化客戶
                    If m_SalesST06 <> "" And Trim(txtPatent(10)) <> "" And Trim(txtPatent(15)) <> "" Then
                        If PUB_ChkOldCustomer(False, txtPatent(10), Trim(txtPatent(15)), m_SalesST15, m_SalesST06) = True Then
                            IsMail = False
                        End If
                    End If
                    
                    'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
                    'If GetST15(txtPatent(15).Text) <> GetCuSales(ChangeCustomerL(txtPatent(11).Text), oStrCuSales4) And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(11).Text) <> "" Then
                    'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuAre
                    If ChkSameCuArea(Trim(txtPatent(11)), Trim(txtPatent(15)), , , , , Trim(txtPatent(13))) = False And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(11).Text) <> "" Then
                    'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    Else
                         If Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(11).Text) <> "" Then
                             IsMail = False
                         End If
                    End If
                    'Added by Lydia 2019/09/16 檢查是否為待活化客戶
                    If m_SalesST06 <> "" And Trim(txtPatent(11)) <> "" And Trim(txtPatent(15)) <> "" Then
                        If PUB_ChkOldCustomer(False, txtPatent(11), Trim(txtPatent(15)), m_SalesST15, m_SalesST06) = True Then
                            IsMail = False
                        End If
                    End If
                    
                    'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
                    'If GetST15(txtPatent(15).Text) <> GetCuSales(ChangeCustomerL(txtPatent(12).Text), oStrCuSales5) And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(12).Text) <> "" Then
                    'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuAre
                    If ChkSameCuArea(Trim(txtPatent(12)), Trim(txtPatent(15)), , , , , Trim(txtPatent(13))) = False And Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(12).Text) <> "" Then
                    'add by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    Else
                         If Trim(txtPatent(15).Text) <> "" And Trim(txtPatent(12).Text) <> "" Then
                             IsMail = False
                         End If
                    End If
                    'Added by Lydia 2019/09/16 檢查是否為待活化客戶
                    If m_SalesST06 <> "" And Trim(txtPatent(12)) <> "" And Trim(txtPatent(15)) <> "" Then
                        If PUB_ChkOldCustomer(False, txtPatent(12), Trim(txtPatent(15)), m_SalesST15, m_SalesST06) = True Then
                            IsMail = False
                        End If
                    End If
                    
'Remove by Morgan 2009/8/20 國外部智權人員改可收所內信件
'                     '2008/12/3 ADD BY SONIA 客戶檔之智權人員為國外部者不發mail
'                     If IsMail = True Then
'                        IsMail = PUB_CHKcusales(Trim(ChangeCustomerL(txtPatent(8))) & "," & Trim(ChangeCustomerL(txtPatent(9))) & "," & Trim(ChangeCustomerL(txtPatent(10))) & "," & Trim(ChangeCustomerL(txtPatent(11))) & "," & Trim(ChangeCustomerL(txtPatent(12))))
'                     End If
'                     '2008/12/3 END
   
                    '2006/8/2 MODIFY BY SONIA TXTSYSTEM只判斷1碼,因為FG
                    'If UCase(Mid(txtSystem, 1, 2)) <> "FC" And oMailCount <> "" Then
                    'edit by nickc 2007/05/10 秀玲說，其中一個符合就不發了
                    'If UCase(Mid(txtSystem, 1, 1)) <> "F" And oMailCount <> "" Then
                    'edit by nickc 2008/03/26 若是申請人全空白，就不管
                    'If UCase(Mid(txtSystem, 1, 1)) <> "F" And IsMail = True Then
                    If UCase(Mid(txtSystem, 1, 1)) <> "F" And IsMail = True And (txtPatent(8) <> "" Or txtPatent(9) <> "" Or txtPatent(10) <> "" Or txtPatent(11) <> "" Or txtPatent(12) <> "") Then
                        'Add By Sindy 2009/10/19
                        '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail，不顯示訊息
                        oMailCount = ""
                        If txtPatent(8) <> "" Then
                           If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtPatent(8).Text), oStrCuSales1)), 1) = "F" Then
                           Else
                              oMailCount = "Y"
                           End If
                        End If
                        If txtPatent(9) <> "" Then
                           If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtPatent(9).Text), oStrCuSales1)), 1) = "F" Then
                           Else
                              oMailCount = "Y"
                           End If
                        End If
                        If txtPatent(10) <> "" Then
                           If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtPatent(10).Text), oStrCuSales1)), 1) = "F" Then
                           Else
                              oMailCount = "Y"
                           End If
                        End If
                        If txtPatent(11) <> "" Then
                           If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtPatent(11).Text), oStrCuSales1)), 1) = "F" Then
                           Else
                              oMailCount = "Y"
                           End If
                        End If
                        If txtPatent(12) <> "" Then
                           If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtPatent(12).Text), oStrCuSales1)), 1) = "F" Then
                           Else
                              oMailCount = "Y"
                           End If
                        End If
                        If Trim(oMailCount) <> "" Then
                        '2009/10/19 End
                           'Modify By Sindy 2010/11/26 申請人1~5為 X65299 或 X03072 的所有關係企業都不檢查業務區
                           If Left(Trim(txtPatent(8)), 6) <> "X65299" And Left(Trim(txtPatent(8)), 6) <> "X03072" And _
                              Left(Trim(txtPatent(11)), 6) <> "X65299" And Left(Trim(txtPatent(11)), 6) <> "X03072" And _
                              Left(Trim(txtPatent(9)), 6) <> "X65299" And Left(Trim(txtPatent(9)), 6) <> "X03072" And _
                              Left(Trim(txtPatent(12)), 6) <> "X65299" And Left(Trim(txtPatent(12)), 6) <> "X03072" And _
                              Left(Trim(txtPatent(10)), 6) <> "X65299" And Left(Trim(txtPatent(10)), 6) <> "X03072" Then
                              MsgBox "收文智權人員與客戶智權人員不同業務區", , "注意！"
                           End If
                        End If
                    End If
                    
                    'Added by Morgan 2013/2/5
                    'FCP國外部收文才要預設費用規費點數
                    If Cancel = False Then
                       'Modified by Lydia 2022/08/25改成共用模組
                       'OnUpdateFee
                       If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), txtPatent(1), txtPatent(15), _
                               IIf(chkWebApp.Visible = True And chkWebApp.Value = 1, "Y", ""), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
                                 txtPatent(17) = m_NowCP16
                                 txtPatent(21) = m_NowCP17
                                 txtPatent(18) = m_NowCP18
                       End If
                       'end 2022/08/25
                    End If
                    'end 2013/2/5
                    
            'Add by Morgan 2003/11/26
            Case 3
               If (txtPatent(1).Text = "109" And txtPatent(3) <> "1") Then
                  MsgBox "'PCT申請'(109)，專利種類只可為 '1' 發明！", vbCritical
                  Cancel = True
               '92.12.21 ADD BY SONIA
               ElseIf (txtPatent(1).Text = 記錄請求_標準專利 And txtPatent(3) <> "1") Then
                  MsgBox "記錄請求_標準專利，專利種類只可為 '1' 發明！", vbCritical
                  Cancel = True
               ElseIf (txtPatent(1).Text = 短期專利申請 And txtPatent(3) <> "2") Then
                  MsgBox "短期專利申請，專利種類只可為 '2' 新型！", vbCritical
                  Cancel = True
               '92.12.21 END
               ElseIf CheckKeyIn(Index) <> 1 Then
                  Cancel = True
               End If
               Call SetCombo3(txtPatent(3)) 'Added by Lydia 2018/11/22 設定案件屬性
            '---end
            
            Case 4
                              
               'Add By Morgan 2003/11/26
               If (txtPatent(1).Text = "109" And txtPatent(4).Text <> "056") Then
                  MsgBox "'PCT申請'(109)，申請國家只可為 'PCT'(056)！", vbCritical
                  Cancel = True
               'Add by Morgan 2005/11/25 只要限制PCT不可收101...敏惠，玲玲
               ElseIf (txtPatent(1).Text = "101" And txtPatent(4).Text = "056") Then
                  MsgBox "申請國家為 'PCT'(056) 時不可收 '發明申請'(101)，請改收 'PCT申請'(109)！", vbCritical
                  Cancel = True
               'Modified by Lydia 2024/02/20 有關" 香港013專利開放收文集體設計申請105"，請比照" 香港013專利收文設計申請103"
               'ElseIf (txtPatent(4).Text = "013" And (txtPatent(1).Text = "101" Or txtPatent(1).Text = "102" Or txtPatent(1).Text = "104" Or txtPatent(1).Text = "105")) Then
               '   MsgBox "申請國家為 '香港'(013)時，案件性質不可為 '101','102','104','105','125'！", vbCritical
               ElseIf (txtPatent(4).Text = "013" And (txtPatent(1).Text = "101" Or txtPatent(1).Text = "102" Or txtPatent(1).Text = "104")) Then
                  MsgBox "申請國家為 '香港'(013)時，案件性質不可為 '101','102','104','125'！", vbCritical
               'end 2024/02/20
                  Cancel = True
               '2007/11/6 ADD BY SONIA
               ElseIf (txtPatent(1).Text = "110" Or txtPatent(1).Text = "112") And txtPatent(4).Text <> "013" Then
                  MsgBox "'案件性質為標準專利記錄請求或短期專利申請，申請國家只可為 '香港'(013)！", vbCritical
                  Cancel = True
               '2007/11/6 EN
               Else
               '--- END
               '93.6.25 ADD BY SONIA
               '2004/3/8 MODIFY BY SONIA 加入澳門044
               'If txtSystem = "P" And (txtPatent(4) <> "000" And txtPatent(4) <> "020" And txtPatent(4) <> "013" And txtPatent(4) <> "056") Then
               If txtSystem = "P" And (txtPatent(4) <> "000" And txtPatent(4) <> "020" And txtPatent(4) <> "013" And txtPatent(4) <> "056" And txtPatent(4) <> "044") Then
                  MsgBox "系統類別為'P'時, 申請國家只可為 台灣,香港, 大陸, 澳門, PCT ！", vbCritical
                  Cancel = True
               End If
               '93.6.25 END
               'Added by Lydia 2018/05/09 新增FMP案
               'Modified by Lydia 2021/11/10 FMP新案之申請國家為013香港或044澳門時，不走命名流程。
               'If txtSystem = "P" And frm010001.txtFMP.Text = "Y" And fraTCT.Visible = True _
                        And txtPatent(4) <> "020" And txtPatent(4) <> "013" And txtPatent(4) <> "056" And txtPatent(4) <> "044" Then
               If txtSystem = "P" And frm010001.txtFMP.Text = "Y" And txtPatent(4) <> "020" And txtPatent(4) <> "013" And txtPatent(4) <> "056" And txtPatent(4) <> "044" Then
                       MsgBox "FMP案的申請國家只可為香港, 大陸, 澳門, PCT ！", vbCritical
                       Cancel = True
               End If
               'end 2018/05/09
               'Added by Lydia 2020/11/19 英國脫歐案管制
               If txtSystem = "CFP" And frm010001.txtCaseNa239 <> "" And txtPatent(4) <> "201" Then
                    MsgBox "英國脫歐案的申請國家只可為英國！", vbCritical
                    Cancel = True
               End If
               'end 2020/11/19
               
                  If CheckKeyIn(Index) <> -1 Then
                     CheckKeyIn 1
                     If CheckKeyIn(3) = 0 Then
                        Cancel = True
                        txtPatent(4).SetFocus
                        Exit Sub
                     End If
                     '更新費用與規費
                     If Cancel = False Then
                        'Modified by Lydia 2022/08/25改成共用模組
                        'OnUpdateFee
                        If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), txtPatent(1), txtPatent(15), _
                                IIf(chkWebApp.Visible = True And chkWebApp.Value = 1, "Y", ""), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
                                   txtPatent(17) = m_NowCP16
                                   txtPatent(21) = m_NowCP17
                                   txtPatent(18) = m_NowCP18
                        End If
                        'end 2022/08/25
                     End If
                  Else
                     Cancel = True
                  End If
                  
               End If
               
               
               '2008/6/6 ADD BY SONIA 荷蘭2008/6/5起取消新型專利,因仍有舊案未發證故不能改國家檔
               If txtPatent(4) = "207" And txtPatent(3) = "2" Then
                  ShowMsg MsgText(9156)
                  Cancel = True
                  txtPatent(4).SetFocus
                  Exit Sub
               End If
               '2008/6/6 END
            'Add By Cheng 2001/12/27
            Case 8 '申請人
               '若申請人有輸入才做Check動作
               If Len(Trim(Me.txtPatent(Index).Text)) > 0 Then
                  If CheckKeyIn(Index) = -1 Then
                     Cancel = True
                  End If
               End If
            Case 13 '代理人
               '若代理人有輸入才做Check動作
               If Len(Trim(Me.txtPatent(Index).Text)) > 0 Then
                  If CheckKeyIn(Index) = -1 Then
                     Cancel = True
                  End If
               End If
               '若申請人與代理人同時空白時
               If Len(Trim(Me.txtPatent(8).Text)) <= 0 And Len(Trim(Me.txtPatent(13).Text)) <= 0 Then
                  MsgBox "申請人與代理人必須至少輸入一項!!!", vbExclamation
'                  Cancel = True
               End If
            '92.2.16 add by sonia
            Case 14 '本所期限
               If CheckKeyIn(Index) <> 1 Then
                  Cancel = True
               'Added by Morgan 2012/2/13
               Else
                  'P,CFP 若本所期限非工作天則直接調整至最近的工作天
                  'Modified by Lydia 2020/07/07 本所期限檢查：所有系統類別的本所期限都要控制是工作日
                  'If txtPatent(Index) <> "" And (txtSystem = "P" Or txtSystem = "CFP") Then
                  If txtPatent(Index) <> "" Then
                     txtPatent(Index).Text = TransDate(PUB_GetWorkDay1(txtPatent(Index).Text, True), 1)
                  End If
               'End 2012/2/13
               End If
                
               ' 主動修正一般無期限, 但修正一定有期限
               If Cancel = False And txtPatent(1) = 修正 And txtPatent(14) = "" And txtPatent(4) = "000" And txtSystem = "P" Then
                  MsgBox "案件性質為修正時, 一定有期限!!, 若該案無期限可能是 主動修正 !!!", vbExclamation + vbOKOnly
                  Cancel = True
               End If
            '92.2.16 end
               '92.2.18 ADD BY SONIA
               ' 請求面詢或閱卷無期限, 但面詢或閱卷一定有期限
               If Cancel = False And (txtPatent(1) = 面詢 Or txtPatent(1) = 閱卷) And txtPatent(14) = "" And txtPatent(4) = "000" And txtSystem = "P" Then
                  MsgBox "案件性質為面詢或閱卷時, 一定有期限!!, 若該案無期限可能是 請求面詢 或 請求閱卷 !!!", vbExclamation + vbOKOnly
                  Cancel = True
               End If
               If Cancel = False And (txtPatent(1) = 請求面詢 Or txtPatent(1) = 請求閱卷) And txtPatent(14) <> "" And txtPatent(4) = "000" And txtSystem = "P" Then
                  MsgBox "案件性質為請求面詢或請求閱卷時, 一定沒有期限!!!", vbExclamation + vbOKOnly
                  Cancel = True
               End If
               '92.2.18 END
            ' 91.09.11 modify by louis
            Case 19
               'Added by Lydia 2023/06/08 依林總指示，FCP、FG案件之本所期限由系統依輸入之法定期限計算，本所期限=法定期限-2工作天；收文畫面已調整本所期限與法定期限欄的先後順序。
               '不論新案或舊案收文，只要不是系統自動帶出期限的案件，收文人員輸入法定期限後，系統自動計算本所期限，並且限制不可修改；
               If txtSystem = "FCP" And bolisNP0809 = False And (frm010001.intModifyKind = 0 Or frm010001.intModifyKind = 1) Then
                  If txtPatent(Index) <> "" Then
                     txtPatent(14).Locked = True '限制本所期限不可輸入
                     If txtPatent(Index).Text <> txtPatent(Index).Tag Then
                        strExc(1) = PUB_GetFCPOurDeadline(DBDATE(txtPatent(Index)))
                        If strExc(1) < strSrvDate(1) Then strExc(1) = strSrvDate(1) '小於系統日=改用系統日
                        txtPatent(14) = TransDate(strExc(1), 1)
                     End If
                  Else
                     txtPatent(14).Locked = False
                  End If
               End If
               'end 2023/06/08
               
               If CheckKeyIn(Index) <> 1 Then
                  Cancel = True
               End If
               
               ' 91.09.11 modify by louis 更新費用與規費
               If Cancel = False Then
                  'Modified by Lydia 2022/08/25改成共用模組
                  'OnUpdateFee
                  If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), txtPatent(1), txtPatent(15), _
                       IIf(chkWebApp.Visible = True And chkWebApp.Value = 1, "Y", ""), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
                         txtPatent(17) = m_NowCP16
                         txtPatent(21) = m_NowCP17
                         txtPatent(18) = m_NowCP18
                  End If
                  'end 2022/08/25
               End If
            
            '92.2.16 add by sonia
'            Case 21 '規費
'                'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'                If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" Then
'                       If CheckKeyIn(Index) <> 1 Then
'                          Cancel = True
'                       End If
'
'                        ' 主張優先權無規費, 但申請優先權證明一定有規費
'                        If Cancel = False And txtPatent(1) = 主張優先權 And Val(txtPatent(21)) <> 0 And txtPatent(4) = "000" And txtSystem = "P" Then
'                           MsgBox "案件性質為主張優先權時, 一定沒有規費!!!", vbExclamation + vbOKOnly
'                           Cancel = True
'                        End If
'                        If Cancel = False And txtPatent(1) = 申請優先權證明 And Val(txtPatent(21)) = 0 And txtPatent(4) = "000" And txtSystem = "P" Then
'                           MsgBox "案件性質為申請優先權證明時, 一定有規費!!!", vbExclamation + vbOKOnly
'                           Cancel = True
'                        End If
'                        ' P之讓與及專利權讓與規費不同
'                        If Cancel = False And txtPatent(1) = 讓與 And txtPatent(21) <> "2000" And txtPatent(4) = "000" And txtSystem = "P" Then
'                           MsgBox "台灣讓與案規費必須為2000 !!!", vbExclamation + vbOKOnly
'                           Cancel = True
'                        End If
'                        If Cancel = False And txtPatent(1) = 專利權讓與 And txtPatent(21) <> "2000" And txtPatent(4) = "000" And txtSystem = "P" Then
'                           MsgBox "台灣專利權讓與案規費必須為2000 !!!", vbExclamation + vbOKOnly
'                           Cancel = True
'                        End If
'                        '92.2.16 end
'                        '92.2.18 ADD BY SONIA
'                        ' 面詢無規費, 但請求面詢一定有規費
'                        '2007/7/26 cancel by sonia 玲玲說請求面詢已取消控制, 此處也可取消 P-081733
'                        'If Cancel = False And txtPatent(1) = 面詢 And Val(txtPatent(21)) <> 0 And txtPatent(4) = "000" And txtSystem = "P" Then
'                        '   MsgBox "案件性質為面詢時, 一定沒有規費!!!", vbExclamation + vbOKOnly
'                        '   Cancel = True
'                        'End If
'                        '2007/7/26 end
'                        '2005/6/13 CANCEL BY SONIA 專業部說現在可於面詢再繳規費
'                        'If Cancel = False And txtPatent(1) = 請求面詢 And Val(txtPatent(21)) = 0 And txtPatent(4) = "000" And txtSystem = "P" Then
'                        '   MsgBox "案件性質為請求面詢時, 一定有規費!!!", vbExclamation + vbOKOnly
'                        '   Cancel = True
'                        'End If
'                        '2005/6/13 END
'                        ' 閱卷無規費, 但請求閱卷一定有規費
'                        If Cancel = False And txtPatent(1) = 閱卷 And Val(txtPatent(21)) <> 0 And txtPatent(4) = "000" And txtSystem = "P" Then
'                           MsgBox "案件性質為閱卷時, 一定沒有規費!!!", vbExclamation + vbOKOnly
'                           Cancel = True
'                        End If
'                        '93.7.15 cancel by sonia
'                        'If Cancel = False And txtPatent(1) = 請求閱卷 And Val(txtPatent(21)) = 0 And txtPatent(4) = "000" And txtSystem = "P" Then
'                        '   MsgBox "案件性質為請求閱卷時, 一定有規費!!!", vbExclamation + vbOKOnly
'                        '   Cancel = True
'                        'End If
'                        '93.7.15
'                        '92.2.18 END
'
'                        'Cancel by Morgan 2004/6/9
'        '                'Add by Morgan 2004/6/9
'        '                If Cancel = False And txtPatent(1) = 601 And txtPatent(21) <> "" And txtPatent(4) = "000" And txtSystem = "P" And Val(txtPatent(21)) > 5000 Then
'        '                   MsgBox "案件性質為領證時, 規費不可大於 5000 !!!" & Chr(13) & "若要繳第二年以後年費，請分開收文。", vbExclamation + vbOKOnly
'        '                   Cancel = True
'        '                End If
'                        'end
'                End If
            Case Else
'                Select Case Index
'                   Case 8, 9, 10, 11, 12, 13
'                      If Len(txtPatent(Index).Text) = 6 Then
'                         txtPatent(Index).Text = txtPatent(Index).Text & "000"
'                      Else
'                         If Len(txtPatent(Index).Text) = 8 Then
'                            txtPatent(Index).Text = txtPatent(Index).Text & "0"
'                         End If
'                      End If
'                End Select
               If CheckKeyIn(Index) <> 1 Then
                  Cancel = True
               End If
   End Select
   If Cancel Then txtPatent_GotFocus (Index)
End Sub
Private Function CheckKeyIn(ByRef intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, bolIsChina As Boolean, strCusTemp As String
Static strLastCus As String
Dim intCounter As Integer
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strFee

   CheckKeyIn = -1
   Select Case intIndex
             Case 5
                        If CheckLengthIsOK(txtPatent(intIndex), 160) Then
                            CheckKeyIn = 1
                        End If
             Case 6
                        'Modified by Lydia 2021/07/19 專利-英文名稱從180放大到250
                        If CheckLengthIsOK(txtPatent(intIndex), 250) Then
                            CheckKeyIn = 1
                        End If
             Case 0
                        If CheckIsTaiwanDate(txtPatent(intIndex).Text) Then
                            CheckKeyIn = 1
                        End If
             Case 1
                       If txtPatent(4) <> 台灣國家代號 Then bolIsChina = True Else bolIsChina = False
                       Call SetPA178 'Add By Sindy 2022/12/7 證書形式
                       'edit by nickc 2007/02/02 不用 dll 了
                       'If objPublicData.GetCaseProperty(txtSystem, txtPatent(intIndex), strTemp, bolIsChina) Then
                       If ClsPDGetCaseProperty(txtSystem, txtPatent(intIndex), strTemp, bolIsChina) Then
                           lblCaseProperty.Caption = strTemp
                           CheckKeyIn = 1
                       End If
             Case 2
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCaseSource(txtPatent(intIndex).Text, strTemp) Then
                        If ClsPDGetCaseSource(txtPatent(intIndex).Text, strTemp) Then
                           lblCaseSource.Caption = strTemp
                           CheckKeyIn = 1
                        End If
             Case 3
                        Select Case txtPatent(1)
                           Case "101"
                              If txtPatent(intIndex) <> "1" Then
                                 ShowMsg "專利種類輸入錯誤"
                                 CheckKeyIn = -1
                                 Exit Function
                              End If
                           Case "102"
                              If txtPatent(intIndex) <> "2" Then
                                 ShowMsg "專利種類輸入錯誤"
                                 CheckKeyIn = -1
                                 Exit Function
                              End If
                           Case "103"
                              If txtPatent(intIndex) <> "3" Then
                                 ShowMsg "專利種類輸入錯誤"
                                 CheckKeyIn = -1
                                 Exit Function
                              End If
                           Case "104"
                              If txtPatent(intIndex) <> "1" And txtPatent(intIndex) <> "2" Then
                                 ShowMsg "專利種類輸入錯誤"
                                 CheckKeyIn = -1
                                 Exit Function
                              End If
                           'Modified by Morgan 2012/12/19 +125
                           Case "105", "125"
                              If txtPatent(intIndex) <> "3" Then
                                 ShowMsg "專利種類輸入錯誤"
                                 CheckKeyIn = -1
                                 Exit Function
                              End If
                        End Select
                        If txtPatent(4) <> 台灣國家代號 Then bolIsChina = True Else bolIsChina = False
                        'edit by nickc 2007/02/02 不用 dll 了
                        'CheckKeyIn = objPublicData.GetPatentTrademarkKind(專利, txtPatent(intIndex).Text, strTemp, bolIsChina, IIf(txtPatent(4) = "", 台灣國家代號, txtPatent(4)))
                        CheckKeyIn = ClsPDGetPatentTrademarkKind(專利, txtPatent(intIndex).Text, strTemp, bolIsChina, IIf(txtPatent(4) = "", 台灣國家代號, txtPatent(4)))
                        If CheckKeyIn = 1 Then
                           lblTrademarkKind = strTemp
                        End If
             Case 4
                        Call SetPA178 'Add By Sindy 2022/12/7 證書形式
                        '91.10.25 add by sonia
                        If txtSystem = "FCP" And txtPatent(intIndex).Text <> 台灣國家代號 Then
                           ShowMsg MsgText(9219)
                           CheckKeyIn = -1
                           Exit Function
                        End If
                        '91.10.25 END
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetNation(txtPatent(intIndex).Text, strTemp) Then
                        If ClsPDGetNation(txtPatent(intIndex).Text, strTemp) Then
                           lblNation.Caption = strTemp
                           CheckKeyIn = 1
                        End If
                        If Val(txtPatent(intIndex)) >= 1 And Val(txtPatent(intIndex)) <= 8 Then
                           ShowMsg MsgText(38)
                           CheckKeyIn = -1
                        End If
             Case 7
                        If txtPatent(5) = "" And txtPatent(6) = "" And txtPatent(7) = "" Then
                           ShowMsg MsgText(1031)
                           intIndex = 5
                           CheckKeyIn = 0
                        ElseIf CheckLengthIsOK(txtPatent(intIndex), 160) Then
                            CheckKeyIn = 1
                        End If
             Case 8 '申請人
                        If txtPatent(intIndex) = "" Then
                           ShowMsg MsgText(1015)
                           CheckKeyIn = -1
                           Exit Function
                        End If
                        For intCounter = 9 To 12
                           If txtPatent(intIndex) = txtPatent(intCounter) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        Next intCounter
                        strCusTemp = txtPatent(intIndex)
                        'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
                        'If objPublicData.GetCustomer(strCusTemp, strTemp, strTemp1) Then
                        'Modify By Sindy 2015/8/27 +txtSystem
                        'Modify By Sindy 2021/2/1 + , strXState(8), IIf(frm010001.intSaveMode = 0, True, False)
                        'Modified by Lydia 2023/03/06 傳入本所案號 , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
                        If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, strXState(8), IIf(frm010001.intSaveMode = 0, True, False), , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) Then
                           txtPatent(intIndex) = strCusTemp
                           lblPetition(intIndex - 8).Caption = strTemp
                           If strLastCus <> strCusTemp Or txtPatent(16).Text = "" Then
                              txtPatent(16).Text = strTemp1
                              strLastCus = strCusTemp
                           End If
                           CheckKeyIn = 1
                           'Add by Morgan 2008/8/5
                           '2015/8/13 modif by sonia 取消不同才做,否則已輸入編號後才去新增的接洽人不會帶出來
                           'If ChangeCustomerL(strCusTemp) <> strAppNo1 Then
                              strAppNo1 = ChangeCustomerL(strCusTemp)
                              'Modify By Sindy 2015/9/11 要傳入畫面上欄位值,不然使用者點選了資料,程式再Run到這又不見了,變成loop
                              'Modify by Amy 2022/11/10 改成Form 2.0
'                              If cboContact.Text = "" Then
'                                 PUB_AddContact strAppNo1, cboContact, , True
'                              Else
'                                 PUB_AddContact strAppNo1, cboContact, Format(cboContact.ItemData(cboContact.ListIndex), "00"), True
'                              End If
'                              '2015/9/11 END
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
                           'End If
                        End If
                        If CheckKeyIn = 1 Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetCustomerNation(strCusTemp, strNation) Then
                           If ClsPDGetCustomerNation(strCusTemp, strNation) Then
                              'If strNation >= "010" Then
                              '   txtPatent(20) = "N"
                              'Else
                              '   txtPatent(20) = ""
                              'End If
                           End If
                        End If
                        'Add By Cheng 2003/09/08
                        If CheckKeyIn = 1 Then
                            'edit by nick 2004/11/10  新增時才有作用
                            'If frm010001.m_blnNewCase = True Then
                            If frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 Then
                                '若輸入9碼且最後一碼不為"0"
                                If Len(Me.txtPatent(intIndex).Text) = 9 And Right(Me.txtPatent(intIndex).Text, 1) <> "0" Then
                                   'Added by Lydia 2024/02/16 專利案件FCP、P、CFP的分割307時，改為彈訊息並可選擇是或否？
                                   If (txtSystem = "FCP" Or txtSystem = "CFP" Or txtSystem = "P") And txtPatent(1) = "307" Then
                                      If MsgBox("申請人為舊名稱編號，分割案確定是要用舊名稱編號收文嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                                         CheckKeyIn = -1
                                      End If
                                   Else
                                   'end 2024/02/16
                                      MsgBox "此申請人已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
                                      CheckKeyIn = -1
                                   End If
                                End If
                            End If
                        End If
                        '2007/7/4 add by sonia 檢查是否同意重新委任
                        If ChkAgree928(txtPatent(intIndex)) = False Then
                           CheckKeyIn = -1
                        End If
                        '2007/7/4 end
             Case 9, 10, 11, 12 '申請人
                        If txtPatent(intIndex) = "" Then
                           CheckKeyIn = 1
                           Exit Function
                        End If
                        If intIndex = 9 Then
                           If txtPatent(intIndex) = txtPatent(8) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(10) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(11) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(12) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           '2007/7/4 add by sonia 檢查是否同意重新委任
                           If ChkAgree928(txtPatent(intIndex)) = False Then
                              CheckKeyIn = -1
                           End If
                           '2007/7/4 end
                        End If
                        If intIndex = 10 Then
                           If txtPatent(intIndex) = txtPatent(8) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(9) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(11) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(12) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           '2007/7/4 add by sonia 檢查是否同意重新委任
                           If ChkAgree928(txtPatent(intIndex)) = False Then
                              CheckKeyIn = -1
                           End If
                           '2007/7/4 end
                        End If
                        If intIndex = 11 Then
                           If txtPatent(intIndex) = txtPatent(8) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(9) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(10) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(12) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           '2007/7/4 add by sonia 檢查是否同意重新委任
                           If ChkAgree928(txtPatent(intIndex)) = False Then
                              CheckKeyIn = -1
                           End If
                           '2007/7/4 end
                        End If
                        If intIndex = 12 Then
                           If txtPatent(intIndex) = txtPatent(8) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(9) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(10) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtPatent(intIndex) = txtPatent(11) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           '2007/7/4 add by sonia 檢查是否同意重新委任
                           If ChkAgree928(txtPatent(intIndex)) = False Then
                              CheckKeyIn = -1
                           End If
                           '2007/7/4 end
                        End If
                        If txtPatent(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           strCusTemp = txtPatent(intIndex)
                           'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
                           'If objPublicData.GetCustomer(strCusTemp, strTemp, strTemp1) Then
                           'Modify By Sindy 2015/8/27 +txtSystem
                           'Modify By Sindy 2021/2/1 + , strXState(intIndex), IIf(frm010001.intSaveMode = 0, True, False)
                           'Modified by Lydia 2023/03/06 傳入本所案號 , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
                           If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, strXState(intIndex), IIf(frm010001.intSaveMode = 0, True, False), , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) Then
                              txtPatent(intIndex) = strCusTemp
                              lblPetition(intIndex - 8).Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        End If
                        'Add By Cheng 2003/09/08
                        If CheckKeyIn = 1 Then
                            '2010/9/30 modify by sonia 新增時才要檢查
                            'If frm010001.m_blnNewCase = True Then
                            If frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 Then
                                '若輸入9碼且最後一碼不為"0"
                                If Len(Me.txtPatent(intIndex).Text) = 9 And Right(Me.txtPatent(intIndex).Text, 1) <> "0" Then
                                   'Added by Lydia 2024/02/16 專利案件FCP、P、CFP的分割307時，改為彈訊息並可選擇是或否？
                                   If (txtSystem = "FCP" Or txtSystem = "CFP" Or txtSystem = "P") And txtPatent(1) = "307" Then
                                      If MsgBox("申請人為舊名稱編號，分割案確定是要用舊名稱編號收文嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                                         CheckKeyIn = -1
                                      End If
                                   Else
                                   'end 2024/02/16
                                      MsgBox "此申請人已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
                                      CheckKeyIn = -1
                                   End If
                                End If
                            End If
                        End If
             Case 13 '代理人
                        strCusTemp = txtPatent(intIndex)
                        If txtPatent(intIndex) <> "" Then
                           'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
                           'If objPublicData.GetAgent(strCusTemp, strTemp) Then
                           'Modify By Sindy 2015/8/27 +txtSystem
                           'Modify By Sindy 2021/2/1 + , strYState, IIf(frm010001.intSaveMode = 0, True, False)
                           If GetAgentAndState(strCusTemp, strTemp, , , , txtSystem, strYState, IIf(frm010001.intSaveMode = 0, True, False)) Then
                              txtPatent(intIndex) = strCusTemp
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
                                If Len(Me.txtPatent(intIndex).Text) = 9 And Right(Me.txtPatent(intIndex).Text, 1) <> "0" Then
                                    MsgBox "此代理人已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
                                    CheckKeyIn = -1
                                End If
                            End If
                        End If
             Case 14
                        If txtPatent(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           If CheckIsTaiwanDate(txtPatent(intIndex).Text) Then
                              If CheckReKey(txtPatent(intIndex)) Then
'                                 If txtPatent(intIndex) = GetTaiwanTodayDate Then
'                                    ShowMsg "此案件已屆本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
'                                 End If
'                                 If txtPatent(intIndex) < GetTaiwanTodayDate Then
'                                    ShowMsg "此案件已逾本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
'                                 End If
                                 CheckKeyIn = 1
                              Else
                                 CheckKeyIn = 0
                              End If
                           End If
                        End If
             Case 19
                        If txtPatent(intIndex) <> "" Then
                           If CheckIsTaiwanDate(txtPatent(intIndex).Text) Then
                              If Val(txtPatent(14)) <= Val(txtPatent(19)) Then
                                 If CheckReKey(txtPatent(intIndex)) Then
                                    CheckKeyIn = 1
                                 Else
                                    CheckKeyIn = 0
                                 End If
                              Else
                                 ShowMsg MsgText(1033)
                              End If
                           End If
                        'Modify By Sindy 2021/4/19 FCP,FG開放可以只輸入本所期限,無法限(不是真正智慧局的期限)
                        ElseIf txtPatent(14) <> "" Then
                           If strSrvDate(1) >= 外專台灣案約定期限啟用日 And _
                              (Me.txtSystem.Text = "FCP" Or Me.txtSystem.Text = "FG") Then
                              CheckKeyIn = 1
                           Else
                        '2021/4/19 END
                              ShowMsg MsgText(1033)
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
             Case 15
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetStaff(txtPatent(intIndex).Text, strTemp, strTemp1) Then
                        If ClsPDGetStaff(txtPatent(intIndex).Text, strTemp, strTemp1) Then
                           CheckKeyIn = 1
                        End If
                        lblSales.Caption = strTemp
                        
                        'Modified by Lydia 2019/02/14
                        'strTemp = GetST15(txtPatent(intIndex).Text, strTemp1)
                        'Modified by Lydia 2019/09/16
                        'm_SalesST15 = GetST15(txtPatent(intIndex).Text, strTemp1)
                        m_SalesST15 = GetST15(txtPatent(intIndex).Text, strTemp1, , m_SalesST06)
                        
                        lblDepartment = strTemp1
'             Case 17
'                    'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'                    If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" Then
'                        'edit by nickc 2007/02/02 不用 dll 了
'                        'If objPublicData.GetCaseLowPrice(txtSystem, txtPatent(4), txtPatent(1), douStPrice, douLowPrice) = 1 Then
'                        If ClsPDGetCaseLowPrice(txtSystem, txtPatent(4), txtPatent(1), douStPrice, douLowPrice) = 1 Then
'                           '94.1.21 ADD BY SONIA 再審依專利種類重新設定標準價底價
'                           'edit by nickc 2006/05/03 沒作用的功能，移除
'                           'If txtSystem = "P" And txtPatent(4) = "000" And txtPatent(1) = "107" Then
'                           'End If
'                           '94.1.21 END
'                        End If
'                        '900803 nick '***************
'                        If txtPatent(intIndex) <> "" Then
'                           '2007/11/5 add by sonia 外專收文之P,CFP案不檢查費用
'                           If ClsPDGetStaffArea(txtPatent(15), strTemp) Then
'                              If strTemp = "F23" And (txtSystem = "P" Or txtSystem = "CFP") Then
'                                 CheckKeyIn = 1
'                                 GoTo EXITSUB
'                              End If
'                           End If
'                           '2007/11/5 END
'                           'edit by nickc 2007/02/02 不用 dll 了
'                           'If objPublicData.GetCaseFee(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(intIndex))) = 1 Then
'                           If ClsPDGetCaseFee(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(intIndex))) = 1 Then
'                              CheckKeyIn = 1
'                           Else
'                              CheckKeyIn = 0
'                           End If
'                        Else
'                           CheckKeyIn = 1
'                        End If
'                        '*************************
'                    Else
'                        '2010/12/31 ADD BY SONIA 也要抓CP33,CP34(P-097399實審)
'                        If ClsPDGetCaseLowPrice(txtSystem, txtPatent(4), txtPatent(1), douStPrice, douLowPrice) = 1 Then
'                        End If
'                        '2010/12/31 END
'                        CheckKeyIn = 1
'                    End If
'             Case 18
'                    'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'                    If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" Then
'                        If txtPatent(intIndex) = "" Then
'                           If txtPatent(17) <> "" Or txtPatent(21) <> "" Then
'                              ShowMsg MsgText(1035)
'                              CheckKeyIn = 0
'                           Else
'                              CheckKeyIn = 1
'                           End If
'                        ElseIf txtPatent(17) <> "" Or txtPatent(21) <> "" Then
''                           If (Val(txtPatent(17)) - Val(txtPatent(21))) / 1000 <> Val(txtPatent(18)) Then
''                              ShowMsg MsgText(1036)
'                              CheckKeyIn = 0
''                           Else
'                              CheckKeyIn = 1
''                           End If
'                        Else
'                           ShowMsg MsgText(1037)
'                        End If
'                    Else
'                        CheckKeyIn = 1
'                    End If
             Case 20
                        'If strNation >= "010" Then
                           'If txtPatent(20) <> "N" Then
                           '   ShowMsg "申請人國籍非台灣時, 是否開電腦收據必須為 N"
                           '   CheckKeyIn = -1
                           '   Exit Function
                           'End If
                        'End If
                        If txtPatent(intIndex) = "" Or txtPatent(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
'             Case 21 '規費
'                    'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'                    If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" Then
'                        ' 91.10.24 modify by sonia
'                        ' 91.09.11 modify by louis
'                        'strFee = GetPatentOfficialFee(txtSystem, txtPatent(1), txtPatent(19), txtPatent(3), txtPatent(4), m_PA16)
'                        If txtPatent(4) = "000" Then
'                           '2009/12/31 modify by sonia 台灣發明實審調整規費
'                           'strFee = GetPatentOfficialFee(txtSystem, txtPatent(1), txtPatent(19), txtPatent(3), txtPatent(4), m_PA16)
'                           If txtCode(1) = "" Then txtCode(1).Text = "0"
'                           If txtCode(2) = "" Then txtCode(2).Text = "00"
'                           '2010/8/17 MODIFY BY SONIA
'                           'strFee = GetPatentOfficialFee(txtSystem, txtPatent(1), txtPatent(19), txtPatent(3), txtPatent(4), m_PA16, , txtCode(0), txtCode(1), txtCode(2))
'                           strFee = GetPatentOfficialFee(txtSystem, txtPatent(1), txtPatent(19), txtPatent(3), txtPatent(4), m_PA16, m_PA14, txtCode(0), txtCode(1), txtCode(2))
'                        End If
'                        ' 91.10.24 end
'                        ' 有抓到規費時才去檢查
'                        If Val(strFee) > 0 Then
'                           If Val(txtPatent(intIndex)) <> Val(strFee) Then
'                              strTit = "檢核資料"
'                              strMsg = "規費數值應為<" & strFee & ">"
'                              nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
'                              GoTo EXITSUB
'                           End If
'                        End If
'
'                        '若有輸入規費
'                        '2009/6/12 modify by sonia取消有輸入才檢查的限制
'                        'If txtPatent(intIndex) <> "" Then   '2009/6/12 cancel by sonia
'                           '2007/11/5 add by sonia 外專收文之P,CFP案不檢查規費
'                           If ClsPDGetStaffArea(txtPatent(15), strTemp) Then
'                              If strTemp = "F23" And (txtSystem = "P" Or txtSystem = "CFP") Then
'                                 CheckKeyIn = 1
'                                 GoTo EXITSUB
'                              End If
'                           End If
'                           '2007/11/5 END
'                           '91.11.21 CANCEL BY SONIA
'                           '若沒輸入費用
'                           'If txtPatent(17) = "" Then
'                           '   ShowMsg MsgText(1039)
'                           '若有輸入費用
'                           'ElseIf objPublicData.GetCaseMoney(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(intIndex))) = 1 Then
'                           'edit by nickc 2006/12/05 搬到 basQuery
'                           'If objPublicData.GetCaseMoney(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(intIndex))) = 1 Then
'                           'Add by Morgan 2008/8/5
'                           '台灣發明申請若有英文摘要時規費應為2700
'                           If txtPatent(4) = "000" And txtPatent(1) = "101" And chkEnglish.Value = 1 Then
'                               If Val(txtPatent(intIndex)) = 2700 Then
'                                 CheckKeyIn = 1
'                               Else
'                                 MsgBox "台灣發明申請若有英文摘要時規費應為2700！"
'                               End If
'                           Else
'                           'end 2008/8/5
'                              '2009/2/10 modify by sonia 是否為同時申請三國(含)以上之美日德可多5點
'                              'If GetCaseMoney(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(intIndex))) = 1 Then
'                              If GetCaseMoney(txtSystem, txtPatent(4), txtPatent(1), Val(txtPatent(intIndex)), False, chkEnglish.Value) = 1 Then
'                              '2009/2/10 end
'                              '91.11.21 END
'                                 CheckKeyIn = 1
'                              End If
'                           End If
'                        'Else                      '2009/6/12 cancel by sonia
''                           If txtPatent(17) <> "" Then
''                              ShowMsg MsgText(1040)
''                              CheckKeyIn = 0
''                           Else
'                        '      CheckKeyIn = 1      '2009/6/12 cancel by sonia
''                           End If
'                        'End If                    '2009/6/12 cancel by sonia
'                    Else
'                        CheckKeyIn = 1
'                    End If
             Case 23
                        If txtPatent(23) <> "" Then
                           strCusTemp = txtPatent(intIndex)
                           'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
                           'If objPublicData.GetCustomer(strCusTemp, strTemp, strTemp1) Then
                           'Modify By Sindy 2015/8/27 +txtSystem
                           'Modify By Sindy 2021/2/1 + , strXState(23), IIf(frm010001.intSaveMode = 0, True, False)
                           'Modified by Lydia 2023/03/06 傳入本所案號 , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
                           If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, strXState(23), IIf(frm010001.intSaveMode = 0, True, False), , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) Then
                              txtPatent(intIndex) = strCusTemp
                              lblPetitionName = strTemp
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
             Case 24
                        If txtPatent(24) <> "" Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetStaff(txtPatent(intIndex), strTemp) Then
                           If ClsPDGetStaff(txtPatent(intIndex), strTemp) Then
                              lblPromoter = strTemp
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
             'Added by Lydia 2017/04/10 客戶案件案件放寬到100
             Case 25
                        'Modified by Lydia 2017/06/14 改常數
                        'If CheckLengthIsOK(txtPatent(intIndex), 100) Then
                        If CheckLengthIsOK(txtPatent(intIndex), 專利客戶案號max) Then
                            CheckKeyIn = 1
                        End If
             'end 2017/04/10
             'add by nickc 2005/10/06 加長分所號
             Case 26
                        If CheckLengthIsOK(txtPatent(intIndex), 50) Then
                            CheckKeyIn = 1
                        End If
             'add by nickc 2008/05/02 加預定收款日
             Case 28
                        If txtPatent(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           If CheckIsTaiwanDate(txtPatent(intIndex).Text) Then
                                'edit by nickc 2008/05/21 原先判斷系統日，改成判斷收文日，因為分所會去補分所號
                                'If DBDATE(txtPatent(intIndex).Text) >= strSrvDate(1) Then
                                If DBDATE(txtPatent(intIndex).Text) >= DBDATE(txtPatent(0).Text) Then
                                   CheckKeyIn = 1
                                Else
                                    'edit by nickc 2008/05/21 原先判斷系統日，改成判斷收文日，因為分所會去補分所號
                                    'MsgBox "預定收款日必須>= 系統日", vbOKOnly + vbCritical, "輸入錯誤！"
                                    MsgBox "預定收款日必須>= 收文日", vbOKOnly + vbCritical, "輸入錯誤！"
                                End If
                           End If
                        End If
             Case Else
                        CheckKeyIn = 1
   End Select
EXITSUB:
End Function

'Add By Sindy 2022/12/7 證書形式
Private Sub SetPA178()
   
   Label1(141).Visible = False
   txtPatent(29).Visible = False
   Label28(1).Visible = False
   '台灣P案領證收文,下列欄位必須顯示出來
   If txtPatent(4) = "000" Then
      If txtSystem = "P" And Trim(txtPatent(1)) = 領證及繳年費 Then
         Label1(141).Visible = True
         txtPatent(29).Visible = True
         Label28(1).Visible = True
      End If
   End If
End Sub

'Modify by Amy 2021/12/16 原:Integer
Private Sub txtPatent_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
             Case 4, 8, 9, 10, 11, 12, 13, 15, 20, 23, 24
                       KeyAscii = UpperCase(KeyAscii)
             Case 16
                       'Modify by Amy 2021/12/16 +txtPatent(Index)
                       KeyAscii = ChangeZIP(KeyAscii, txtPatent(Index))
             'Modify By Sindy 2022/12/7 證書形式
             Case 29
                  If (KeyAscii < 49 Or KeyAscii > 50) And KeyAscii <> 8 Then
                     KeyAscii = 0
                     Beep
                  End If
                  '2022/12/7 END
   End Select
End Sub
Private Sub txtPatent_GotFocus(Index As Integer)
   If Index = 8 Then
      If txtPatent(5) = "" And txtPatent(6) = "" And txtPatent(7) = "" Then
         txtPatent(5).SetFocus
         Exit Sub
      End If
   End If
   txtPatent(Index).SelStart = 0
   txtPatent(Index).SelLength = Len(txtPatent(Index).Text)
   '儲存未修改前之值至Tag中,供再確認時使用
   txtPatent(Index).Tag = txtPatent(Index)
   '切換輸入法
   Select Case Index
             Case 5
                        'edit by nickc 2007/06/06 切換輸入法改用API
                        'txtPatent(Index).IMEMode = 1
                        OpenIme
             Case Else
                        'edit by nickc 2007/06/06 切換輸入法改用API
                        'txtPatent(Index).IMEMode = 2
                        CloseIme
   End Select
End Sub
Private Sub txtPatent_LostFocus(Index As Integer)
   '關閉輸入法
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtPatent(Index).IMEMode = 2
   'CloseIme 'Removed by Morgan 2016/10/20 會造成 Win7 的切換錯誤
   
   'Add By Cheng 2001/12/27
   If Index = 13 And Len(Trim(Me.txtPatent(8).Text)) <= 0 And Len(Trim(Me.txtPatent(13).Text)) <= 0 Then
      '若申請人與代理人皆未輸入, 則將游標設定在申請人欄位
      Me.txtPatent(8).SetFocus
   End If
End Sub

'新增Patent至資料庫
'edit by nickc 2007/03/27 加入彼所案號
'Modified by Morgan 2021/7/21 +PA167
'Removed by Morgan 2024/11/18 收文存檔模組已啟用,舊程式標記為註解,後續無需再修改
'Private Function InsertPatentDatabase(ByRef intSaveMode As Integer, ByRef pa01 As String, _
'             ByRef pa02 As String, ByRef pa03 As String, ByRef pa04 As String, ByRef pa05 As String, _
'             ByRef pa06 As String, ByRef pa07 As String, ByRef PA08 As String, ByRef PA09 As String, _
'             ByRef pa26 As String, ByRef pa27 As String, ByRef pa28 As String, _
'             ByRef pa29 As String, ByRef pa30 As String, ByRef pa75 As String, _
'             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, _
'             ByRef cp11 As String, ByRef cp13 As String, ByRef cp16 As String, _
'             ByRef cp17 As String, ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, _
'             ByRef cp56 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef CP09 As String, _
'             ByRef cp02 As String, ByRef cp33 As Double, ByRef cp34 As Double, ByRef CP64 As String, _
'             ByRef CP89 As String, ByRef CP90 As String, ByRef CP91 As String, ByRef CP92 As String, ByRef PA77 As String, ByRef PA149 As String, ByRef PA167 As String) As Boolean
'Dim strSql As String, pa23 As String, PA46 As String, cp31 As String, strAutoNumber As String
'Dim np13 As String, np14 As String, bolRt As Boolean, cp55 As String, cp26 As String, cp20 As String
'Dim cp93 As String, cp94 As String, cp95 As String, cp96 As String 'Add by Morgan 2006/6/23
'Dim cp48 As String 'Add by Morgan 2008/8/19
'Dim strCustomer(4) As String, i As Integer, pa85 As String, bolError As Boolean, ipa85 As Integer
'Dim CP43 As String 'Added by Morgan 2012/8/9
'Dim bolNoAutoCP14 As Boolean 'Added by Morgan 2012/8/24
''edit by nickc 2007/02/06 不用 dll 了
''Dim objPublicData As Object
'Dim adoquery As New ADODB.Recordset
'Dim strPA161 As String 'Add by Amy 2018/10/11 收據公司別
'
'   'add by nickc 2007/12/12
'   If IsSaveData = True Then
'      Exit Function
'   End If
'   IsSaveData = True
'
'On Error GoTo ErrHand
'   '傳入0為重複之本所案號(新增舊案)，1為正確之本所案號(新增新案)
'   cp05 = ChangeTStringToWString(cp05)
'   cp06 = ChangeTStringToWString(cp06)
'   cp07 = ChangeTStringToWString(cp07)
'   pa26 = ChangeCustomerL(pa26)
'   pa27 = ChangeCustomerL(pa27)
'   pa28 = ChangeCustomerL(pa28)
'   pa29 = ChangeCustomerL(pa29)
'   pa30 = ChangeCustomerL(pa30)
'   pa75 = ChangeCustomerL(pa75)
'   'edit by nickc 2007/02/06 不用 dll 了
'   'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
'   cnnConnection.BeginTrans
'   If intSaveMode = 1 Then
'      'edit by nickc 2007/02/06 不用 dll 了
'      'obj001.SetPAFileProperty CP10, pa23, pa46
'      'Modified by Lydia 2022/08/11 拿掉PA46
'      'Cls001SetPAFileProperty CP10, pa23, PA46
'      Cls001SetPAFileProperty CP10, pa23
'      If pa02 = "" Then
'         'edit by nickc 2007/02/06 不用 dll 了
'         'If objPublicData.GetAutoNumber(PA01, strAutoNumber, True, False) Then
'         If ClsPDGetAutoNumber(pa01, strAutoNumber, True, False) Then
'            pa02 = strAutoNumber
'         Else
'            bolError = True
'         End If
'      End If
'      If bolError = False Then
'         'edit by nickc 2007/02/06 不用 dll 了
'         'If objPublicData.GetSystemKind(PA01, , , ipa85) Then
'         If ClsPDGetSystemKind(pa01, , , ipa85) Then
'            pa85 = IIf(ipa85 = 2, 2, 1)
'            cp02 = pa02
'            '91.12.6 modify by sonia pa17預設null
'            'strSQL = "insert into patent (pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa08,pa09,pa23,pa26," & _
'            '   "pa27,pa28,pa29,pa30,pa46,pa75,pa17) values (" & CNULL(PA01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & CNULL(ChgSQL(pa05)) & "," & _
'            '   CNULL(Replace(pa06, "'", "''")) & "," & CNULL(ChgSQL(pa07)) & "," & CNULL(PA08) & "," & CNULL(PA09) & "," & CNULL(pa23) & "," & CNULL(pa26) & "," & CNULL(pa27) & "," & _
'            '   CNULL(pa28) & "," & CNULL(pa29) & "," & CNULL(pa30) & "," & CNULL(pa46) & "," & CNULL(pa75) & ", 'N')"
'            'edit by nickc 2007/03/27 加入彼所案號
'            'strSQL = "insert into patent (pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa08,pa09,pa23,pa26," & _
'               "pa27,pa28,pa29,pa30,pa46,pa75,pa17) values (" & CNULL(PA01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & CNULL(ChgSQL(pa05)) & "," & _
'               CNULL(Replace(pa06, "'", "''")) & "," & CNULL(ChgSQL(pa07)) & "," & CNULL(PA08) & "," & CNULL(PA09) & "," & CNULL(pa23) & "," & CNULL(pa26) & "," & CNULL(pa27) & "," & _
'               CNULL(pa28) & "," & CNULL(pa29) & "," & CNULL(pa30) & "," & CNULL(pa46) & "," & CNULL(pa75) & ", '')"
'            'Modify by Morgan 2008/8/5 &PA149
'   '         strSQL = "insert into patent (pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa08,pa09,pa23,pa26," & _
'   '            "pa27,pa28,pa29,pa30,pa46,pa75,pa17,pa77,pa149) values (" & CNULL(PA01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & CNULL(ChgSQL(pa05)) & "," & _
'   '            CNULL(Replace(pa06, "'", "''")) & "," & CNULL(ChgSQL(pa07)) & "," & CNULL(PA08) & "," & CNULL(PA09) & "," & CNULL(pa23) & "," & CNULL(pa26) & "," & CNULL(pa27) & "," & _
'   '            CNULL(pa28) & "," & CNULL(pa29) & "," & CNULL(pa30) & "," & CNULL(pa46) & "," & CNULL(pa75) & ", ''," & CNULL(PA77) & "," & CNULL(PA149) & ") "
'
'            'Modify by Toni   2008/8/26   為發明人
'            varInventorNo = Split(strInventorNo, ",")
'            For i = 0 To UBound(varInventorNo)
'               strInventor(i) = varInventorNo(i)
'            Next
'            For i = i + 1 To 99 '9
'               strInventor(i) = ""
'            Next
'            'Add By Sindy 2014/11/6 更新專利發明人檔
'            For i = 0 To 99
'               If strInventor(i) <> "" Then
'                  strSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
'                           CNULL(pa01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & i + 1 & ",'" & strInventor(i) & "')"
'                  Pub_SeekTbLog strSql 'Add By Sindy 2017/8/23
'                  cnnConnection.Execute strSql
'               Else
'                  Exit For
'               End If
'            Next i
'            '2014/11/6 END
'
'            'Modify By Sindy 2010/3/8 增加聯絡人pa51~pa56欄位
'            If bolCancel = False Then
'               strPA51s = "": strPA52s = "": strPA53s = ""
'               strPA54s = "": strPA55s = "": strPA56s = ""
'            End If
'   '         strSql = "insert into patent (pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa08,pa09,pa23,pa26," & _
'   '            "pa27,pa28,pa29,pa30,pa46,pa75,pa17,pa77,pa149," & strIntor & ") values (" & CNULL(PA01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & CNULL(ChgSQL(pa05)) & "," & _
'   '            CNULL(Replace(pa06, "'", "''")) & "," & CNULL(ChgSQL(pa07)) & "," & CNULL(PA08) & "," & CNULL(PA09) & "," & CNULL(pa23) & "," & CNULL(pa26) & "," & CNULL(pa27) & "," & _
'   '            CNULL(pa28) & "," & CNULL(pa29) & "," & CNULL(pa30) & "," & CNULL(PA46) & "," & CNULL(pa75) & ", ''," & CNULL(PA77) & "," & CNULL(PA149) & "," & CNULL(strInventor(0)) & "," & _
'   '            CNULL(strInventor(1)) & "," & CNULL(strInventor(2)) & "," & CNULL(strInventor(3)) & "," & CNULL(strInventor(4)) & "," & CNULL(strInventor(5)) & "," & CNULL(strInventor(6)) & " ," & _
'   '            CNULL(strInventor(7)) & "," & CNULL(strInventor(8)) & "," & CNULL(strInventor(9)) & ")"
'            'Modify By Sindy 2010/10/28 增加pa158
'            'Modify By Sindy 2014/11/6
''            strSql = "insert into patent (pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa08,pa09,pa23,pa26," & _
''               "pa27,pa28,pa29,pa30,pa46,pa75,pa17,pa77,pa149," & strIntor & ",pa51,pa52,pa53,pa54,pa55,pa56,pa158) " & _
''               "values (" & CNULL(pa01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & CNULL(ChgSQL(pa05)) & "," & _
''               CNULL(Replace(pa06, "'", "''")) & "," & CNULL(ChgSQL(pa07)) & "," & CNULL(PA08) & "," & CNULL(PA09) & "," & CNULL(pa23) & "," & CNULL(pa26) & "," & CNULL(pa27) & "," & _
''               CNULL(pa28) & "," & CNULL(pa29) & "," & CNULL(pa30) & "," & CNULL(PA46) & "," & CNULL(pa75) & ", ''," & CNULL(PA77) & "," & CNULL(PA149) & "," & CNULL(strInventor(0)) & "," & _
''               CNULL(strInventor(1)) & "," & CNULL(strInventor(2)) & "," & CNULL(strInventor(3)) & "," & CNULL(strInventor(4)) & "," & CNULL(strInventor(5)) & "," & CNULL(strInventor(6)) & " ," & _
''               CNULL(strInventor(7)) & "," & CNULL(strInventor(8)) & "," & CNULL(strInventor(9)) & "," & _
''               CNULL(ChgSQL(strPA51s)) & "," & CNULL(ChgSQL(strPA52s)) & "," & CNULL(ChgSQL(strPA53s)) & "," & CNULL(ChgSQL(strPA54s)) & "," & CNULL(ChgSQL(strPA55s)) & "," & CNULL(ChgSQL(strPA56s)) & "," & CNULL(Left(Combo3, 1)) & ")"
'            strPA161 = GetReceiptCmp(Left(pa26, 8), Mid(pa26, 9, 1), pa01, PA09) 'Added by Amy 2018/10/11 ++收據公司別pa161
'            'Added by Lydia 2020/11/19 CFP英國脫歐案管制：新增英國案時同時把歐盟案相關欄位帶過來(參考PUB_SaveCountry)
'            If pa01 = "CFP" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" Then
'                If PUB_ReadPatentData(m_CaseNa239(), m_CaseNa239(1), m_CaseNa239(2), m_CaseNa239(3), m_CaseNa239(4)) Then
'                   strExc(0) = "": strExc(1) = ""
'                   For i = 5 To TF_PA
'                       Select Case i
'                          Case 92, 93, 94, 95, 96, 97, 108, 136, 137, 138 'Create + Update, 銷卷日
'                          Case 9  '申請國家
'                              strExc(0) = strExc(0) & "PA" & Format(i, "00") & "," 'Insert
'                              strExc(1) = strExc(1) & " '201' as PA09, " 'Select
'                          Case 22 '專利號數: 專利號以歐盟設計專利號(拿掉-符號)前加上9
'                              strExc(0) = strExc(0) & "PA" & Format(i, "00") & ","
'                              strExc(1) = strExc(1) & " '9'||REPLACE(PA22,'-','') AS PA" & Format(i, "00") & ", "
'                          Case 91  '案件備註: 加註歐盟案案號
'                              strExc(0) = strExc(0) & "PA" & Format(i, "00") & ","
'                              strExc(1) = strExc(1) & CNULL("歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";") & "||PA" & Format(i, "00") & " AS PA" & Format(i, "00") & ","
'                          Case 161
'                              strExc(0) = strExc(0) & "PA" & Format(i, "00") & ","
'                              strExc(1) = strExc(1) & " '" & strPA161 & "' as PA161, "
'                          Case Else
'                              strExc(0) = strExc(0) & "PA" & Format(i, "00") & ","
'                              strExc(1) = strExc(1) & "PA" & Format(i, "00") & ","
'                       End Select
'                   Next
'                   strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
'                   strExc(1) = Left(strExc(1), Len(strExc(1)) - 1)
'                   strSql = "INSERT INTO PATENT (PA01,PA02,PA03,PA04," & strExc(0) & ") " & _
'                               "SELECT '" & pa01 & "' as PA01,'" & pa02 & "' as PA02,'" & pa03 & "' as PA03,'" & pa04 & "' as pa04, " & strExc(1) & _
'                               " FROM PATENT WHERE pa01='" & m_CaseNa239(1) & "' and pa02='" & m_CaseNa239(2) & "' and pa03='" & m_CaseNa239(3) & "' and pa04='" & m_CaseNa239(4) & "' "
'                   cnnConnection.Execute strSql
'
'                   If cp07 <> "" Then PUB_UpdUkPayYr cp07, pa01, pa02, pa03, pa04 'Added by Morgan 2020/12/8 更新英國案繳費紀錄
'                End If
'            Else
'            'end 2020/11/19
'                'Modify by Amy 2018/10/11 +收據公司別pa161
'                strSql = "insert into patent (pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa08,pa09,pa23,pa26," & _
'                   "pa27,pa28,pa29,pa30,pa46,pa75,pa17,pa77,pa149,pa51,pa52,pa53,pa54,pa55,pa56,pa158,pa150,pa161,pa176) " & _
'                   "values (" & CNULL(pa01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & CNULL(ChgSQL(pa05)) & "," & _
'                   CNULL(Replace(pa06, "'", "''")) & "," & CNULL(ChgSQL(pa07)) & "," & CNULL(PA08) & "," & CNULL(PA09) & "," & CNULL(pa23) & "," & CNULL(pa26) & "," & CNULL(pa27) & "," & _
'                   CNULL(pa28) & "," & CNULL(pa29) & "," & CNULL(pa30) & "," & CNULL(PA46) & "," & CNULL(pa75) & ", ''," & CNULL(PA77) & "," & CNULL(PA149) & "," & _
'                   CNULL(ChgSQL(strPA51s)) & "," & CNULL(ChgSQL(strPA52s)) & "," & CNULL(ChgSQL(strPA53s)) & "," & CNULL(ChgSQL(strPA54s)) & "," & CNULL(ChgSQL(strPA55s)) & "," & CNULL(ChgSQL(strPA56s)) & "," & CNULL(Left(Combo3, 1)) _
'                   & "," & CNULL(IIf(fraTCT.Visible = True And txtData(2).Text <> "" And txtData(2).Text <> "B", txtData(2), "")) & "," & CNULL(ChgSQL(strPA161)) & ",'" & PA167 & "')"
'                cnnConnection.Execute strSql
'                '2014/11/6 END
'                '91.12.6 end
'                strCustomer(0) = pa26
'                strCustomer(1) = pa27
'                strCustomer(2) = pa28
'                strCustomer(3) = pa29
'                strCustomer(4) = pa30
'                'Memo by Lydia 2020/11/19 CFP英國脫歐案管制：新增英國案時同時把歐盟案相關欄位帶過來，所以不要變更資料
'                For i = 0 To 4
'                       strSql = "update patent set pa" + Format(31 + i) + "=(select cu23 from customer where cu01=" + CNULL(Mid(strCustomer(i), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(i), 9, 1)) + _
'                          "),pa" + Format(36 + i) + "=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(strCustomer(i), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(i), 9, 1)) + _
'                          "),pa" + Format(41 + i) + "=(select cu29 from customer where cu01=" + CNULL(Mid(strCustomer(i), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(i), 9, 1)) + ") where pa01=" + CNULL(pa01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)
'                       cnnConnection.Execute strSql
'                Next
'            End If 'Added by Lydia 2020/11/19
'
'           'Add By Cheng 2003/08/28
'           'Begin
'   'edit by nick 2005/01/07 搬到下面
'   '        strSQL = "Update Patent Set PA47='" & ChgSQL(Me.txtPatent(26).Text) & "', PA48='" & ChgSQL(Me.txtPatent(25).Text) & "' Where pa01=" + CNULL(PA01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)
'   '        cnnConnection.Execute strSQL
'           'End
'            cp31 = "Y"
'         Else
'            bolError = True
'         End If
'      Else
'         bolError = True
'      End If
'   End If
'   If bolError = False Then
'
'      'Modify By Cheng 2002/01/09
'      'Modified by Lydia 2018/05/09 改成模組
''      If Me.txtSystem.Text = "P" Then
''         If CP10 = 主張優先權 Or CP10 = 補文件 Or CP10 = 催審 Or CP10 = 請求公告 Or CP10 = 加註追加 Or CP10 = 加註聯合 _
''            Or CP10 = 合併 Or CP10 = 終止授權 Or CP10 = 繼承 Or CP10 = 設定質權 Or CP10 = 終止設定質權 Or CP10 = 退費 _
''            Or CP10 = 後金 Or CP10 = 補收款 Then
''            cp26 = "N"
''         Else
''            'edit by nickc 2007/02/06 不用 dll 了
''            'obj001.SetPAIsCase CP10, cp26
''            Cls001SetPAIsCase CP10, cp26
''         End If
''      Else
''         '91.7.3 modify by sonia
''         'obj001.SetPAIsCase CP10, cp26
''         'edit by nickc 2007/02/06 不用 dll 了
''         'If Me.txtSystem.Text = "CFP" Then obj001.SetPAIsCase CP10, cp26
''         If Me.txtSystem.Text = "CFP" Then Cls001SetPAIsCase CP10, cp26
''      End If
'      Pub_SetPAIsCase Me.txtSystem.Text, CP10, cp26
'      'end 2018/05/09
'      cp20 = ""
'      '92.11.3 ADD BY SONIA
'      '2010/3/11 MODIFY BY SONIA
'      'If Me.txtSystem.Text = "FCP" Then
'      'Add By Sindy 2022/3/31
'      If m_bolFMP = True Then
'         cp48 = Pub_GetHandleDay("FCP", PA09, CP10, , cp06)
'      '2022/3/31 END
'      ElseIf Me.txtSystem.Text = "FCP" Or Me.txtSystem.Text = "FG" Then
''Modify by Morgan 2010/4/29 改抓案件性質檔設定
''         'add by nick 2004/10/18
''         'If CP10 = 告知代理人 Or CP10 = 回覆代理人 Then
''         '2007/6/21 MODIFY BY SONIA 加928重新委任
''         '2009/11/9 modify by sonia 加935案件轉至本所
''         '2010/3/11 modify by sonia 加924會稿
''         If CP10 = 告知代理人 Or CP10 = 回覆代理人 Or CP10 = 退費 Or CP10 = "928" Or CP10 = "935" Or CP10 = "924" Then
''            cp20 = "N"
''         End If
'         'cp20 = PUB_GetCP20(txtSystem, CP10)      '2010/11/26 CANCEL BY SONIA 移到下一句
''end 2010/4/29
'         'Add by Morgan 2008/8/28 非例外的案件性質要預設承辦期限
'         'Modify by Morgan 2008/10/23
'         'Ciba Y45697的年費承辦期限掛15個工作天
'         If pa75 = "Y45697000" And CP10 = "605" Then
'            cp48 = CompWorkDay(15, strSrvDate(1))
'            If Val(cp06) > 0 And Val(cp48) > Val(cp06) Then
'               cp48 = cp06
'            End If
'         'end 2008/10/23
'
'         'Added by Morgan 2012/7/13
'         '加速審查要判斷已輸入通知實審日才掛承辦期限
'         ElseIf CP10 = "422" Then
'            strExc(1) = pa01
'            strExc(2) = pa02
'            strExc(3) = pa03
'            strExc(4) = pa04
'            If PUB_ChkCPExist(strExc(), "1204") Then
'               cp48 = Pub_GetHandleDay("FCP", "000", CP10, , cp06)
'            End If
'         'end 2012/7/13
'
'         'Add By Sindy 2021/6/24 968回復說明書校閱
'         ElseIf CP10 = "968" Then
'            cp48 = Pub_GetHandleDay("FCP", "000", CP10, , cp06, , , pa01 & "-" & pa02 & "-" & pa03 & "-" & pa04)
'
'         ElseIf InStr(SkipCasePtyList, CP10) = 0 Then
'            cp48 = Pub_GetHandleDay("FCP", "000", CP10, , cp06)
'            'Y54732000 & X30299000組合之回代承辦期限下面另有更新
'         End If
'
'         'Add By Sindy 2021/4/29 不是主管機關期限
'         If m_strCPM34 = "N" And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
'            '(2)收文時無設本所期限，以承辦期限＋5個工作天為本所期限
'            If Val(cp06) = 0 Then
'               cp06 = PUB_GetFCPOurDeadline(DBDATE(cp48), , , , "N")
'            '(1)收文時有設本所期限，自動備註:本所期限為yyy/mm/dd(本所期限)
'            Else
'               CP64 = "本所期限為" & ChangeWStringToTDateString(cp06) & ";" & CP64
'            End If
'         End If
'      End If
'
'      '2010/11/26 add by sonia P的942預設不請款
'      'Modified by Morgan 2019/8/16 +國外部收文條件(目前案件性質對照表設定只有外專用)
'      'If Me.txtSystem.Text = "FCP" Or Me.txtSystem.Text = "FG" Or Me.txtSystem.Text = "P" Then
'      'Modified by Lydia 2022/08/29 debug
'      'If Left(m_SalesST15, 1) = "F" And Me.txtSystem.Text = "FCP" Or Me.txtSystem.Text = "FG" Or Me.txtSystem.Text = "P" Then
'      If Left(m_SalesST15, 1) = "F" And (Me.txtSystem.Text = "FCP" Or Me.txtSystem.Text = "FG" Or Me.txtSystem.Text = "P") Then
'         cp20 = PUB_GetCP20(txtSystem, CP10)
'      End If
'      '2010/11/26 END
'
'      'Added by Lydia 2022/05/03  FCP-062174審定前不收費控制:補上是否向客戶收款=N
'      If m_PA16 = "" And InStr("FCP062174000", Me.txtSystem & Me.txtCode(0) & IIf(Me.txtCode(1) = "", "0", Me.txtCode(1)) & IIf(Me.txtCode(2) = "", "00", Me.txtCode(2))) > 0 Then
'          cp20 = "N"
'      End If
'      'end 2022/05/03
'      'Added by Lydia 2022/05/03 FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
'      If m_PA16 <> "1" And InStr("FCP067004000", Me.txtSystem & Me.txtCode(0) & IIf(Me.txtCode(1) = "", "0", Me.txtCode(1)) & IIf(Me.txtCode(2) = "", "00", Me.txtCode(2))) > 0 Then
'          cp20 = "N"
'      End If
'      'end 2022/05/03
'      '92.11.3 END
'      'edit by nickc 2007/02/06 不用 dll 了
'      'If objPublicData.GetAutoNumber(Left(CP09, 1), strAutoNumber, True, True) Then
'      If ClsPDGetAutoNumber(Left(CP09, 1), strAutoNumber, True, True) Then
'         cp56 = ChangeCustomerL(cp56)
'         If cp56 <> "" Then
'            cp55 = pa26
'            'Add by Morgan 2006/6/23
'            '讓與人2-5,受讓人2-5
'            CP89 = ChangeCustomerL(CP89)
'            CP90 = ChangeCustomerL(CP90)
'            CP91 = ChangeCustomerL(CP91)
'            CP92 = ChangeCustomerL(CP92)
'            cp93 = pa27
'            cp94 = pa28
'            cp95 = pa29
'            cp96 = pa30
'            'end 2006/6/23
'         End If
'         CP09 = CP09 + strAutoNumber
'         'edit by nickc 2007/02/06 不用 dll 了
'         'bolRt = obj001.GetNextProgressData(PA01, pa02, pa03, pa04, CP10, np13, np14)
'         bolRt = Cls001GetNextProgressData(pa01, pa02, pa03, pa04, CP10, np13, np14)
'         'add by nick 2005/01/07 從上面搬下來
'         If Me.txtPatent(26).Text & Me.txtPatent(25).Text <> "" Then
'              strSql = "Update Patent Set PA47='" & ChgSQL(Me.txtPatent(26).Text) & "', PA48='" & ChgSQL(Me.txtPatent(25).Text) & "' Where pa01=" + CNULL(pa01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)
'              cnnConnection.Execute strSql
'         End If
'
'         'Add By Sindy 2012/11/06 有★★的應收帳款簽核控管
'         m_CP150 = ""
'         If Check2.Value = 1 Then m_CP150 = "Y"
'         '2012/11/06 End
'
'         'Modify By Sindy 2012/11/06 +CP150
'         If pa23 <> "1" And cp31 = "Y" Then
'            'Modify by Morgan 2006/6/23 加cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96
'            If bolRt Then
'               strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," & _
'                 "cp16,cp17,cp18,cp19,cp20,cp26,cp31,cp32,cp33,cp34,cp37,cp38,cp39,cp40,CP48,cp55,cp56,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150) values (" & CNULL(pa01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & CNULL(cp05) & "," & _
'                 CNULL(cp06) & "," & CNULL(cp07) & "," & CNULL(np13) & "," & CNULL(CP09) & "," & CNULL(CP10) & "," & CNULL(cp11) & "," & CNULL(cp13) & "," & CNULL(cp14) & "," & CNULL(cp16) & "," & _
'                 CNULL(cp17) & "," & CNULL(cp18) & "," & CNULL(cp19) & "," & CNULL(cp20) & "," & CNULL(cp26) & "," & CNULL(cp31) & "," & CNULL(cp32) & ", " & cp33 & ", " & cp34 & ", " & CNULL(ChgSQL(pa05)) & ", " & CNULL(ChgSQL(pa06)) & ", " & CNULL(ChgSQL(pa07)) & "," & CNULL(ChgSQL(np14)) & "," & CNULL(cp48, True) & "," & CNULL(cp55) & "," & CNULL(cp56) & "," & CNULL(ChgSQL(CP64)) & "," & _
'                 CNULL(CP89) & "," & CNULL(CP90) & "," & CNULL(CP91) & "," & CNULL(CP92) & "," & CNULL(cp93) & "," & CNULL(cp94) & "," & CNULL(cp95) & "," & CNULL(cp96) & "," + CNULL(m_CP150) + ")"
'            Else
'               strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp13,cp14," & _
'                 "cp16,cp17,cp18,cp19,cp20,cp26,cp31,cp32,cp33,cp34,cp37,cp38,cp39,CP48,cp55,cp56,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150) values (" & CNULL(pa01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & CNULL(cp05) & "," & _
'                 CNULL(cp06) & "," & CNULL(cp07) & "," & CNULL(CP09) & "," & CNULL(CP10) & "," & CNULL(cp11) & "," & CNULL(cp13) & "," & CNULL(cp14) & "," & CNULL(cp16) & "," & _
'                 CNULL(cp17) & "," & CNULL(cp18) & "," & CNULL(cp19) & "," & CNULL(cp20) & "," & CNULL(cp26) & "," & CNULL(cp31) & "," & CNULL(cp32) & ", " & cp33 & ", " & cp34 & ", " & CNULL(ChgSQL(pa05)) & ", " & CNULL(ChgSQL(pa06)) & ", " & CNULL(ChgSQL(pa07)) & "," & CNULL(cp48, True) & "," & CNULL(cp55) & "," & CNULL(cp56) & "," & CNULL(ChgSQL(CP64)) & "," & _
'                 CNULL(CP89) & "," & CNULL(CP90) & "," & CNULL(CP91) & "," & CNULL(CP92) & "," & CNULL(cp93) & "," & CNULL(cp94) & "," & CNULL(cp95) & "," & CNULL(cp96) & "," + CNULL(m_CP150) + ")"
'            End If
'         Else
'            If bolRt Then
'               strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," & _
'                 "cp16,cp17,cp18,cp19,cp20,cp26,cp31,cp32,cp33,cp34,cp40,CP48,cp55,cp56,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150) values (" & CNULL(pa01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & CNULL(cp05) & "," & _
'                 CNULL(cp06) & "," & CNULL(cp07) & "," & CNULL(np13) & "," & CNULL(CP09) & "," & CNULL(CP10) & "," & CNULL(cp11) & "," & CNULL(cp13) & "," & CNULL(cp14) & "," & CNULL(cp16) & "," & _
'                 CNULL(cp17) & "," & CNULL(cp18) & "," & CNULL(cp19) & "," & CNULL(cp20) & "," & CNULL(cp26) & "," & CNULL(cp31) & "," & CNULL(cp32) & ", " & cp33 & ", " & cp34 & "," & CNULL(ChgSQL(np14)) & "," & CNULL(cp48, True) & "," & CNULL(cp55) & "," & CNULL(cp56) & "," & CNULL(ChgSQL(CP64)) & "," & _
'                 CNULL(CP89) & "," & CNULL(CP90) & "," & CNULL(CP91) & "," & CNULL(CP92) & "," & CNULL(cp93) & "," & CNULL(cp94) & "," & CNULL(cp95) & "," & CNULL(cp96) & "," + CNULL(m_CP150) + ")"
'            Else
'               strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp13,cp14," & _
'                 "cp16,cp17,cp18,cp19,cp20,cp26,cp31,cp32,cp33,cp34,CP48,cp55,cp56,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150) values (" & CNULL(pa01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & CNULL(cp05) & "," & _
'                 CNULL(cp06) & "," & CNULL(cp07) & "," & CNULL(CP09) & "," & CNULL(CP10) & "," & CNULL(cp11) & "," & CNULL(cp13) & "," & CNULL(cp14) & "," & CNULL(cp16) & "," & _
'                 CNULL(cp17) & "," & CNULL(cp18) & "," & CNULL(cp19) & "," & CNULL(cp20) & "," & CNULL(cp26) & "," & CNULL(cp31) & "," & CNULL(cp32) & ", " & cp33 & ", " & cp34 & "," & CNULL(cp48, True) & "," & CNULL(cp55) & "," & CNULL(cp56) & "," & CNULL(ChgSQL(CP64)) & "," & _
'                 CNULL(CP89) & "," & CNULL(CP90) & "," & CNULL(CP91) & "," & CNULL(CP92) & "," & CNULL(cp93) & "," & CNULL(cp94) & "," & CNULL(cp95) & "," & CNULL(cp96) & "," + CNULL(m_CP150) + ")"
'            End If
'         End If
'         cnnConnection.Execute strSql
'
'         'add by sonia 2019/7/31 Y54732000 & X30299000組合,且會稿924發文後,新案翻譯201發文前收文之回代902,設回代相關收文號掛會稿,承辦期限掛新案翻譯的本所期限
'         If pa75 = "Y54732000" And Left(pa26, 8) = "X3029900" And CP10 = "902" Then
'            strExc(0) = "select c2.cp06,c1.cp09 from caseprogress c1,caseprogress c2 where c1.cp01='" & pa01 & "' and c1.cp02='" & pa02 & "' and c1.cp03='" & pa03 & "' and c1.cp04='" & pa04 & "' and c1.cp10='924' and c1.cp27>0 " & _
'                        "   and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and '201'=c2.cp10(+) and c2.cp158=0"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strSql = "update caseprogress set cp43='" & "" & RsTemp(1) & "',cp48=" & "" & RsTemp(0) & " where cp09=" & CNULL(CP09)
'               cnnConnection.Execute strSql
'            End If
'         End If
'         'end 2019/7/31
'
'         'Modified by Morgan 2012/4/25 +cp71(優先權份數)
'         'Modified by Morgan 2012/6/20 +cp118(電子送件)
'         'strSql = "update caseprogress set cp12=(select st15 from staff where st01=" & CNULL(cp13) & ") where cp09=" & CNULL(CP09)
'         strSql = "update caseprogress set cp12=(select st15 from staff where st01=" & CNULL(cp13) & ")" & IIf(txtCopy.Visible, ",cp71=" & Val(txtCopy), "") & IIf(chkWebApp.Visible, ",cp118='" & IIf(chkWebApp.Value = 1, "Y", "") & "'", "") & " where cp09=" & CNULL(CP09)
'         cnnConnection.Execute strSql
'
'         'Added by Lydia 2020/05/20 法律所案源收文：台灣案B1、B2及C收文時，增加"案源單號"欄位一定要輸入，並將案源單號更新至該筆收文的CP162。
'         If frm010001.intModifyKind = 0 And txtPatent(4) = "000" And (txtSystem = "FCP" Or txtSystem = "P") And m_LOS02 <> "" And m_LOS15 <> "" Then
'              If Left(m_LOS02, 1) = "B" Or Left(m_LOS02, 1) = "C" Then
'                  strSql = "update caseprogress set CP162='" & m_LOS15 & "' where cp09='" & CP09 & "' "
'                  cnnConnection.Execute strSql
'              End If
'         End If
'         'end 2020/05/20
'
'         'Add By Sindy 2009/07/06
'         If textYear.Visible = True Then
'            If m_CP10 = "601" Then
'               If Val(Text1(0)) > 0 Then
'                  strSql = "update caseprogress set cp53=" & CNULL(textYear, True) & ",cp54=" & CNULL(Text1(0), True) & " where cp09=" & CNULL(CP09)
'               End If
'            Else
'               strSql = "update caseprogress set cp53=" & CNULL(textYear, True) & ",cp54=" & CNULL(Text1(0), True) & " where cp09=" & CNULL(CP09)
'            End If
'            cnnConnection.Execute strSql
'         End If
'         '2009/07/06 End
'
'          '若為接洽記錄單(櫃台收文)
'          'Modify by Morgan 2007/10/26 費用可改時才做，否則已收款資料會被還原
'          'If frm010001.intChoose = 0 Then
'          If frm010001.intChoose = 0 And txtPatent(17).Enabled = True Then
'          'end 2007/10/26
'              '未收金額 = 費用
'              strSql = "update caseprogress set cp79=cp16 where cp09=" & CNULL(CP09)
'              cnnConnection.Execute strSql
'          End If
'         'Add By Cheng 2002/05/10
'         '若為內部收文作業時, 案件進度檔的是否向客戶收款設定為"N"
'         If frm010001.intChoose = 1 Then
'            strSql = "Update CaseProgress Set CP20='N' Where cp09=" & CNULL(CP09)
'            cnnConnection.Execute strSql
'         End If
'
'         strSql = "update customer set cu30=" & CNULL(cu30) & " where cu01=" & CNULL(Mid(pa26, 1, 8)) & " and cu02=" & CNULL(Mid(pa26, 9, 1))
'         cnnConnection.Execute strSql
'        'Add by Lydia 2014/10/31 開放外專程序人員可進入專利處系統操作FMP寰華案件=>寫入代理人
'        If mFMPchk = True Then
'            strSql = "update caseprogress set cp44='Y53374000' where cp09='" & CP09 & "' "
'            cnnConnection.Execute strSql
'        End If
'        'end. 'Add by Lydia 2014/10/31
'        'Add by Morgan 2010/8/10
'        '收文美專正式申請案(原暫時申請案號-1)時,沖暫時申請案的其他期限
'        If pa01 = "CFP" And pa03 = "1" And PA09 = "101" And CP10 = "101" Then
'           strExc(0) = "select np01,np08,np09,np22 from nextprogress where np02='" & pa01 & "' and np03='" & pa02 & "' and np04='0' and np05='" & pa04 & "' and np06 is null and np07='910' "
'           intI = 1
'           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'           If intI = 1 Then
'              strSql = "update caseprogress set cp06=" & RsTemp("NP08") & ",cp07=" & RsTemp("NP09") & " where cp09='" & CP09 & "'"
'              cnnConnection.Execute strSql, intI
'              'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
'              strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np01='" & RsTemp("NP01") & "' and np22=" & RsTemp("NP22")
'              cnnConnection.Execute strSql, intI
'           End If
'        End If
'        'Added by Lydia 2020/11/19 CFP英國脫歐案管制
'        If pa01 = "CFP" And cp31 = "Y" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" Then
'             strExc(0) = "select cp09,cp30 from caseprogress where cp01='" & m_CaseNa239(1) & "' and cp02='" & m_CaseNa239(2) & "' and cp03='" & m_CaseNa239(3) & "' and cp04='" & m_CaseNa239(4) & "' " & _
'                              "and substr(cp09,1,1) ='C' and cp10='1608' and cp159=0 order by cp05 desc "
'             intI = 1
'             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'             If intI = 1 Then
'                'A. 歐盟案若有「通知英國再註冊」的C類來函1608之CP30存至新英國案之專利號數PA22
'                If "" & RsTemp.Fields("CP30") <> "" Then
'                    strSql = "update patent set pa22='" & ChgSQL(RsTemp.Fields("cp30")) & "' where pa01='" & pa01 & "' and pa02='" & pa02 & "' and pa03='" & pa03 & "' and pa04='" & pa04 & "' "
'                    cnnConnection.Execute strSql
'                End If
'                'B. 歐盟案若有「通知英國再註冊」的C類來函也轉至新英國案號
'                If "" & RsTemp.Fields("cp09") <> "" Then
'                     strSql = "update caseprogress set cp01='" & pa01 & "', cp02='" & pa02 & "', cp03='" & pa03 & "', cp04='" & pa04 & "' where cp09='" & RsTemp.Fields("cp09") & "' "
'                     cnnConnection.Execute strSql
'                End If
'             End If
'             'Added by Lydia 2020/12/01
'             If CP10 = "444" Then
'                  '委任代理人上續辦；下一程序備註加註「英國案案號」
'                  strSql = "update nextprogress set np06='Y', np24='" & CP09 & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "英國案案號：" & pa01 & pa02 & pa03 & pa04 & ";'||np15 " & _
'                               "where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='444' and np06 is null "
'                  cnnConnection.Execute strSql
'                  'E. 若收文「委任代理人(CFP.444)」時歐盟案下一程序之/「延展費(英國)613」期限轉至新案號並改案件性質為「延展費607」；下一程序備註加註「歐盟案案號」
'                  'Modified by Lydia 2020/12/16 將NP01改為英國案收文號; 否則分案作業會錯誤(本所案號不同)
'                  'strSql = "update nextprogress set np02='" & pa01 & "', np03='" & pa02 & "', np04='" & pa03 & "', np05='" & pa04 & "', np07='607', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";'||np15 " & _
'                              "where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='613' and np06 is null "
'                  strSql = "update nextprogress set np01='" & CP09 & "', np02='" & pa01 & "', np03='" & pa02 & "', np04='" & pa03 & "', np05='" & pa04 & "', np07='607', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";'||np15 " & _
'                              "where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='613' and np06 is null "
'                  cnnConnection.Execute strSql
'             Else
'             'end 2020/12/01
'                  'C. 歐盟案下一程序之「延展(英國)」(CFP.613)期限上續辦NP06，下一單據編號NP24記錄新英國案之總收文號；下一程序備註加註「英國案案號」
'                  strSql = "update nextprogress set np06='Y', np24='" & CP09 & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "英國案案號：" & pa01 & pa02 & pa03 & pa04 & ";'||np15 where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='613' and np06 is null "
'                  cnnConnection.Execute strSql
'                  'Added by Lydia 2020/12/04 將歐盟案「委任代理人」期限轉至新英國案；下一程序備註加註「歐盟案案號」
'                  'Modified by Lydia 2020/12/16 將NP01改為英國案收文號; 否則分案作業會錯誤(本所案號不同)
'                  'strSql = "update nextprogress set np02='" & pa01 & "', np03='" & pa02 & "', np04='" & pa03 & "', np05='" & pa04 & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";'||np15 " & _
'                              "where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='444' and np06 is null "
'                  strSql = "update nextprogress set np01='" & CP09 & "', np02='" & pa01 & "', np03='" & pa02 & "', np04='" & pa03 & "', np05='" & pa04 & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";'||np15 " & _
'                              "where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='444' and np06 is null "
'                  cnnConnection.Execute strSql
'                  'end 2020/12/04
'             End If 'Added by Lydia 2020/12/01
'
'             'D. 建立歐盟案及英國案之關聯(相關卷號、國內外案、多國案)
'             '----相關卷號
'             strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(pa01) & ", " & CNULL(pa02) & ", " & CNULL(pa03) & ", " & CNULL(pa04) & ", " & CNULL(m_CaseNa239(1)) & ", " & CNULL(m_CaseNa239(2)) & ", " & CNULL(m_CaseNa239(3)) & ", " & CNULL(m_CaseNa239(4)) & " ) "
'             cnnConnection.Execute strSql
'             strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_CaseNa239(1)) & ", " & CNULL(m_CaseNa239(2)) & ", " & CNULL(m_CaseNa239(3)) & ", " & CNULL(m_CaseNa239(4)) & ", " & CNULL(pa01) & ", " & CNULL(pa02) & ", " & CNULL(pa03) & ", " & CNULL(pa04) & " ) "
'             cnnConnection.Execute strSql
'             '----國內外案
'             strSql = "insert into CaseMap(CM01,CM02,CM03,CM04,CM05,CM06,CM07,CM08,CM09,CM10,CM11) select '" & pa01 & "', '" & pa02 & "', '" & pa03 & "','" & pa04 & "',CM05,CM06,CM07,CM08,CM09,CM10,CM11 " & _
'                          "from CaseMap where CM01='" & m_CaseNa239(1) & "' and CM02='" & m_CaseNa239(2) & "' and CM03='" & m_CaseNa239(3) & "' and CM04='" & m_CaseNa239(4) & "' and cm10 in ('0','3','4','5','6') "
'             cnnConnection.Execute strSql
'             strSql = "insert into CaseMap(CM01,CM02,CM03,CM04,CM05,CM06,CM07,CM08,CM09,CM10,CM11) select CM01,CM02,CM03,CM04,'" & pa01 & "', '" & pa02 & "', '" & pa03 & "','" & pa04 & "',CM09,CM10,CM11 " & _
'                          "from CaseMap where CM05='" & m_CaseNa239(1) & "' and CM06='" & m_CaseNa239(2) & "' and CM07='" & m_CaseNa239(3) & "' and CM08='" & m_CaseNa239(4) & "' and cm10 in ('0','3','4','5','6') "
'             cnnConnection.Execute strSql
'             '----多國案號(先增加歐盟案原本關聯，再增加歐盟案和英國案之關聯)
'             strSql = "insert into caserelation(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) select '" & pa01 & "', '" & pa02 & "', '" & pa03 & "','" & pa04 & "',cr05,cr06,cr07,cr08 " & _
'                          "from caserelation where cr01='" & m_CaseNa239(1) & "' and cr02='" & m_CaseNa239(2) & "' and cr03='" & m_CaseNa239(3) & "' and cr04='" & m_CaseNa239(4) & "' "
'             cnnConnection.Execute strSql
'             strSql = "insert into caserelation(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) select cr01,cr02,cr03,cr04,'" & pa01 & "', '" & pa02 & "', '" & pa03 & "','" & pa04 & "' " & _
'                          "from caserelation where cr05='" & m_CaseNa239(1) & "' and cr06='" & m_CaseNa239(2) & "' and cr07='" & m_CaseNa239(3) & "' and cr08='" & m_CaseNa239(4) & "' "
'             cnnConnection.Execute strSql
'             strSql = "insert into caserelation(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(pa01) & ", " & CNULL(pa02) & ", " & CNULL(pa03) & ", " & CNULL(pa04) & ", " & CNULL(m_CaseNa239(1)) & ", " & CNULL(m_CaseNa239(2)) & ", " & CNULL(m_CaseNa239(3)) & ", " & CNULL(m_CaseNa239(4)) & " ) "
'             cnnConnection.Execute strSql
'             strSql = "insert into caserelation(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_CaseNa239(1)) & ", " & CNULL(m_CaseNa239(2)) & ", " & CNULL(m_CaseNa239(3)) & ", " & CNULL(m_CaseNa239(4)) & ", " & CNULL(pa01) & ", " & CNULL(pa02) & ", " & CNULL(pa03) & ", " & CNULL(pa04) & " ) "
'             cnnConnection.Execute strSql
'             'Added by Lydia 2020/12/04 歐盟案案件備註加註「英國案案號」；新英國案之新案收文的進度備註加註「歐盟案案號」
'             strSql = "Update Patent set PA91=" & CNULL("英國案案號：" & pa01 & pa02 & pa03 & pa04 & ";") & "||PA91 where PA01='" & m_CaseNa239(1) & "' and PA02='" & m_CaseNa239(2) & "' and PA03='" & m_CaseNa239(3) & "' and PA04='" & m_CaseNa239(4) & "' "
'             cnnConnection.Execute strSql
'             strSql = "Update CaseProgress set CP64=" & CNULL("歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";") & "||CP64 where CP09='" & CP09 & "' "
'             cnnConnection.Execute strSql
'             'end 2020/12/04
'             'Added by Lydia 2021/01/11 複製優先權資料
'             strSql = "insert into pridate (pd01,pd02,pd03,pd04,pd05,pd06,pd07,pd08,pd09,pd10) " & _
'                          "select " & CNULL(pa01) & ", " & CNULL(pa02) & ", " & CNULL(pa03) & ", " & CNULL(pa04) & ", pd05,pd06,pd07,pd08,pd09,pd10 from pridate where pd01='" & m_CaseNa239(1) & "' and pd02='" & m_CaseNa239(2) & "' and pd03='" & m_CaseNa239(3) & "' and pd04='" & m_CaseNa239(4) & "' "
'             cnnConnection.Execute strSql
'             'end 2021/01/11
'             'Added by Lydia 2021/04/15 CFP英國脫歐委任代理之後續處理：收文英國延展及委任代理人新案，同時將代理人存入CP44。
'             strExc(0) = "select np01,np15 from nextprogress where np07='444' and np15 like '%脫歐英國案代理人：%' " & _
'                              "and ((np02='" & pa01 & "' and np03='" & pa02 & "' and np04='" & pa03 & "' and np05='" & pa04 & "') or (np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "'  and np04='" & m_CaseNa239(3) & "'  and np05='" & m_CaseNa239(4) & "')) "
'             intI = 1
'             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'             If intI = 1 Then
'                 strExc(1) = Mid("" & RsTemp.Fields("np15"), InStr(RsTemp.Fields("np15"), "脫歐英國案代理人：") + 9, 9)
'                 If Left(strExc(1), 1) = "Y" Then
'                     strSql = "Update CaseProgress set cp44=" & CNULL(strExc(1)) & " where cp09=" & CNULL(CP09)
'                     cnnConnection.Execute strSql
'                 End If
'             End If
'             'end 2021/04/15
'             PUB_EUtoUK pa01, pa02, pa03, pa04, m_CaseNa239(1), m_CaseNa239(2), m_CaseNa239(3), m_CaseNa239(4), CP09, CP10  'Added by Morgan 2020/12/21 回覆單歸卷
'        End If
'        'end 2020/11/19
'
'        'Added by Lydia 2021/04/15 CFP英國脫歐委任代理之後續處理：收文英國延展及委任代理人新案，同時將代理人存入CP44。
'                                                 '同一天接洽單之後收文的處理
'        If txtSystem = "CFP" And cp31 <> "Y" And txtPatent(4) = "201" And (txtPatent(1) = "444" Or txtPatent(1) = "607") And m_PA91 <> "" And InStr(m_PA91, "歐盟案案號：") > 0 Then
'             strExc(0) = "select cp44 from caseprogress where cp01='" & pa01 & "' and cp02='" & pa02 & "' and cp03='" & pa03 & "' and cp04='" & pa04 & "' and cp05=" & DBDATE(txtPatent(0)) & " and cp10 in ('607','444') and cp159=0 and cp31='Y' "
'             intI = 1
'             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'             If intI = 1 Then
'                 If "" & RsTemp.Fields("cp44") <> "" Then
'                     strSql = "Update CaseProgress set cp44=" & CNULL(RsTemp.Fields("cp44")) & " where cp09=" & CNULL(CP09)
'                     cnnConnection.Execute strSql
'                 End If
'             End If
'        End If
'        'end 2021/04/15
'
'        'Modify by Morgan 2006/5/4
'        'FCP的補文件202不要做
'        If Not (pa01 = "FCP" And CP10 = "202") Then
'           strSql = "select np01 from nextprogress where np02 = '" & pa01 & "' and np03 = '" & pa02 & "' and np04 = '" & pa03 & "' and np05 = '" & pa04 & "' and np06 is null and np07 = '" & CP10 & "'"
'           'Add by Morgan 2007/1/12 台灣專利的申復或修正時下一程序兩個都要抓
'           If PA09 = "000" And (CP10 = "205" Or CP10 = "204") Then
'              strSql = "select np01 from nextprogress where np02 = '" & pa01 & "' and np03 = '" & pa02 & "' and np04 = '" & pa03 & "' and np05 = '" & pa04 & "' and np06 is null and np07 IN ('204','205')"
'           End If
'           'end 2007/1/12
'           adoquery.CursorLocation = adUseClient
'           adoquery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'           If adoquery.RecordCount > 0 Then
'              If adoquery.RecordCount = 1 Then
'                 CP43 = adoquery.Fields(0) 'Added by Morgan 2012/8/9
'                 If IsNull(adoquery.Fields(0).Value) = False Then
'                    'Add by Morgan 2010/6/30 異議答辯、舉發答辯要一並更新對造資料
'                    If (CP10 = "802" Or CP10 = "804") Then
'                       cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & CP09 & "'", intI
'                    Else
'                    'End 2010/6/30
'                       cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & CP09 & "'"
'                    End If
'                 End If
'                 'add by nick 2004/09/08
'                 If txtPatent(1).Text <> "411" Then
'                    'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
'                    strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" & CNULL(pa01) & " and np03=" & _
'                            CNULL(pa02) & " and np04=" & CNULL(pa03) & " and np05=" & CNULL(pa04) & _
'                            " and np07=" & CNULL(CP10) & " and np06 is null"
'
'                    'Add by Morgan 2007/1/12 台灣專利的申復或修正時下一程序兩個都要抓
'                    If PA09 = "000" And (CP10 = "205" Or CP10 = "204") Then
'                       'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
'                       strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" & CNULL(pa01) & " and np03=" & CNULL(pa02) & " and np04=" & CNULL(pa03) & " and np05=" & CNULL(pa04) & " and np07 in ('204','205') and np06 is null"
'                    End If
'                    'end 2007/1/12
'                    cnnConnection.Execute strSql
'                 End If
'              End If
'           Else
'              adoquery.Close
'              strSql = "select np01 from nextprogress where np02 = '" & pa01 & "' and np03 = '" & pa02 & "' and np04 = '" & pa03 & "' and np05 = '" & pa04 & "' and np06 <>'Y' and np07 = '" & CP10 & "'"
'              'Add by Morgan 2007/1/12 台灣專利的申復或修正時下一程序兩個都要抓
'              If PA09 = "000" And (CP10 = "205" Or CP10 = "204") Then
'                 strSql = "select np01 from nextprogress where np02 = '" & pa01 & "' and np03 = '" & pa02 & "' and np04 = '" & pa03 & "' and np05 = '" & pa04 & "' and np06<>'Y' and np07 IN ('204','205')"
'              End If
'              'end 2007/1/12
'              adoquery.CursorLocation = adUseClient
'              adoquery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'              If adoquery.RecordCount > 0 Then
'                 If adoquery.RecordCount = 1 Then
'                    CP43 = adoquery.Fields(0) 'Added by Morgan 2012/8/9
'                    If IsNull(adoquery.Fields(0).Value) = False Then
'                       'Add by Morgan 2010/6/30 異議答辯、舉發答辯要一並更新對造資料
'                       If (CP10 = "802" Or CP10 = "804") Then
'                          cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & CP09 & "'", intI
'                       Else
'                       'End 2010/6/30
'                          cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & CP09 & "'"
'                       End If
'                    End If
'                    'add by nick 2004/09/08
'                    If txtPatent(1).Text <> "411" Then
'                         'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
'                         strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" & CNULL(pa01) & " and np03=" & _
'                            CNULL(pa02) & " and np04=" & CNULL(pa03) & " and np05=" & CNULL(pa04) & _
'                            " and np07=" & CNULL(CP10) & " and np06 <> 'Y'"
'                         'Add by Morgan 2007/1/12 台灣專利的申復或修正時下一程序兩個都要抓
'                          If PA09 = "000" And (CP10 = "205" Or CP10 = "204") Then
'                             'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
'                             strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" & CNULL(pa01) & " and np03=" & CNULL(pa02) & " and np04=" & CNULL(pa03) & " and np05=" & CNULL(pa04) & " and np07 in ('204','205') and np06<>'Y'"
'                          End If
'                          'end 2007/1/12
'                         cnnConnection.Execute strSql
'                    End If
'                 End If
'              End If
'           End If
'           adoquery.Close
'        End If
'        '2006/5/4 end
'
'         '92.2.19 END
'         'edit by nickc 2007/02/06 不用 dll 了
'         'If obj001.SetCaseProgressFee(PA01, PA09, CP10, CP09) = False Then bolError = True
'         If Cls001SetCaseProgressFee(pa01, PA09, CP10, CP09) = False Then bolError = True
'      Else
'         bolError = True
'      End If
'   End If
'   'Modify By Cheng 2002/12/18
'   'adoquery.CursorLocation = adUseClient
'   ''adoquery.Open "select np01 from nextprogress where np02 = '" & pa01 & "' and np03 = '" & pa02 & "' and np04 = '" & pa03 & "' and np05 = '" & pa04 & "' and np07 = '" & cp10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
'   'adoquery.Open "select np01 from nextprogress where np02 = '" & PA01 & "' and np03 = '" & pa02 & "' and np04 = '" & pa03 & "' and np05 = '" & pa04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
'   ''Modify By Cheng 2002/05/10
'   ''若在下一程序檔只抓到一筆資料時, 才要抓下一程序檔的總收文號更新案件進度檔的相關總收文號
'   ''If adoquery.RecordCount <> 0 Then
'   'If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
'   '   If IsNull(adoquery.Fields(0).Value) = False Then
'   '      cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & cp09 & "'"
'   '   End If
'   'End If
'   'adoquery.Close
'   'add by nickc 2008/05/02 儲存預定收款日
'   If bolError = False Then
'       'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
''       Dim rtCnt As Integer
''       'Modify by Morgan 2010/12/9
''       'If txtPatent(28) <> "" Then
''       '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0) + 1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " ", rtCnt
''       If txtPatent(28) <> "" And txtPatent(28) <> txtPatent(28).Tag Then
''           cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0) + 1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
''       'end 2010/12/9
''           If rtCnt = 0 Then
''               cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from dual "
''           End If
''       End If
'       'end 2018/08/22
'
'      'Added by Morgan 2012/8/9
'      If pa01 = "FCP" Then
'         'Modified by Morgan 2012/9/13 +603
'         'Modified by Morgan 2012/10/24 +935
'         'Modified by Morgan 2012/12/19 +125
'         If InStr("101,102,103,105,125,401,404,416,601,603,605,701,702,908,929,935", CP10) > 0 Then
'            'Added by Morgan 2012/8/24
'            '分割案的實審不要預設
'            If CP10 = "416" Then
'               strExc(1) = pa01
'               strExc(2) = pa02
'               strExc(3) = pa03
'               strExc(4) = pa04
'               If PUB_ChkCPExist(strExc, "307") Then
'                  bolNoAutoCP14 = True
'               End If
'            End If
'            'end 2012/8/24
'            If bolNoAutoCP14 = False Then
'               strExc(1) = PUB_GetFCPHandler(pa01, pa02, pa03, pa04, CP10)
'               If strExc(1) <> "" Then
'                  strSql = "update caseprogress set cp14='" & strExc(1) & "' where cp09='" & CP09 & "'"
'                  cnnConnection.Execute strSql, intI
'               End If
'            End If
'         ElseIf CP43 > "C" Then
'            strSql = "update caseprogress a set cp14=(select nvl(max(b.cp14),a.cp14) from caseprogress b,staff where b.cp09=a.cp43 and st01(+)=cp14 and st03 in ('F21','F81') and st04='1') where cp09='" & CP09 & "'"
'            cnnConnection.Execute strSql, intI
'         End If
'      End If
'   End If
'
'   'Add by Amy 2013/06/26 FCP/P 新申請案101/102/103及衍生申請125 則追蹤流水號不可為空
'   '2015/4/14 MODIFY BY SONIA 林總指示開放投資法務人員可收文L案及P案,故加Trim(txtTCN01) <> ""
'   'If strSrvDate(2) >= Val("1020723") And (txtSystem = "FCP" Or txtSystem = "P") And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "125") Then
'   'Modified by Lydia 2018/05/25 改成有輸入就更新(P-120302香港案標準專利記錄的Tracking_no.9662沒有回寫, 程式上線前人工更新)
'   'If Trim(txtTCN01) <> "" And strSrvDate(2) >= Val("1020723") And (txtSystem = "FCP" Or txtSystem = "P") And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "125") Then
'   If Trim(txtTCN01) <> "" And txtTCN01.Visible = True Then
'        UpdateTCN01 CP09
'   End If
'   'end 2013/06/26
'
'   'Modify By Sindy 2016/4/13 改移到專利處的分案作業執行
''   'Added by Morgan 2013/8/1
''   'P或CFP案收文主動修正203或修正204時若有相關新案已齊備未發文則清除完稿日及會稿日並EMail通知承辦人
''   If (pa01 = "CFP" Or pa01 = "P") And (CP10 = "203" Or CP10 = "204") Then
''      strExc(0) = "select cp09,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo,cp14,ep09,ep07" & _
''         " from (select cm01,cm02,cm03,cm04 from casemap where cm10='0' and cm05='" & pa01 & "' and cm06='" & pa02 & "' and cm07='" & pa03 & "' and cm08='" & pa04 & "'" & _
''         " union select cm05,cm06,cm07,cm08 from casemap where cm10='0' and cm01='" & pa01 & "' and cm02='" & pa02 & "' and cm03='" & pa03 & "' and cm04='" & pa04 & "'" & _
''         " union select cr01,cr02,cr03,cr04 from caserelation where cr05='" & pa01 & "' and cr06='" & pa02 & "' and cr07='" & pa03 & "' and cr08='" & pa04 & "'" & _
''         "),caseprogress,engineerprogress where cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03 and cp04(+)=cm04 and instr('" & NewCasePtyList & "',cp10)>0" & _
''         " and cp27||cp57 is null and ep02(+)=cp09 and ep06>0"
''      intI = 1
''      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''      If intI = 1 Then
''         Do While Not RsTemp.EOF
''            If RsTemp("ep09") > 0 Then
''               'Modify By Sindy 2014/1/10 +EP12
''               'strSql = "update engineerprogress set ep09=null,ep07=null where ep02='" & RsTemp("cp09") & "'"
''               strSql = "update engineerprogress set ep09=null,ep07=null,ep12='" & ChangeTStringToTDateString(strSrvDate(2)) & "因相關案" & pa01 & "-" & pa02 & IIf(pa03 & pa04 = "000", "", "-" & pa03 & "-" & pa04) & "已收文" & lblCaseProperty & "固清除相關日期,原完稿日：'||ep09||'原會稿日：'||ep07||';'||ep12 where ep02='" & RsTemp("cp09") & "'"
''               cnnConnection.Execute strSql, intI
''            End If
''            strExc(1) = RsTemp("Cno") & " 的相關案 " & pa01 & "-" & pa02 & IIf(pa03 & pa04 = "000", "", "-" & pa03 & "-" & pa04) & " 已收文" & lblCaseProperty & "，請於2日內確認預定修正的內容..."
''            If IsNull(RsTemp("ep09")) Then
''               strExc(2) = "無"
''            Else
''               strExc(2) = TranslateKeyWord(incCNV_CHINESE_MINKO, RsTemp("ep09"), "")
''            End If
''            If IsNull(RsTemp("ep07")) Then
''               strExc(3) = "無"
''            Else
''               strExc(3) = TranslateKeyWord(incCNV_CHINESE_MINKO, RsTemp("ep07"), "")
''            End If
''            strExc(4) = RsTemp("Cno") & " 的相關案 " & pa01 & "-" & pa02 & IIf(pa03 & pa04 = "000", "", "-" & pa03 & "-" & pa04) & " 已收文" & lblCaseProperty & "，請於2日內確認預定修正的內容，" & _
''               "若修正內容已實質改變原齊備內容，請修改齊備日；若修正內容未實質改變原齊備內容，請將原完稿日及原會稿日(若有)填入系統。" & _
''               vbCrLf & "原完稿日：" & strExc(2) & _
''               vbCrLf & "原會稿日：" & strExc(3)
''            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08) values ('" & strUserNum & "','" & RsTemp("cp14") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & strExc(1) & "','" & strExc(4) & "')"
''            cnnConnection.Execute strSql, intI
''            RsTemp.MoveNext
''         Loop
''      End If
''   End If
''   'end 2013/8/1
'
'   'Added by Lydia 2015/12/31 收文FCP案領證601且五個申請人有一個為X47794(三星鑽石)時，檢查該案號的行事曆期限，若事由有"可收文領證"時，於收文存檔時同時解除該筆行事曆期限。
'   If CP10 = "601" And InStr(pa26 & "," & pa27 & "," & pa28 & "," & pa29 & "," & pa30, "X47794") > 0 Then
'       strExc(0) = "select * from staff_calendar where sc05='" & pa01 & "' and sc06='" & pa02 & "' and sc07='" & pa03 & "' and sc08='" & pa04 & "' " & _
'                   "and sc04 like '%FCP%(三星鑽石)%可收文領證%' and sc18 is null "
'       intI = 1
'       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'       If intI = 1 Then
'          If RsTemp.Fields("SC10") > 1 Then
'             If PUB_AddFCPStaffCalendar(RsTemp.Fields("SC01"), RsTemp.Fields("SC10"), RsTemp.Fields("SC03"), RsTemp.Fields("SC04"), RsTemp.Fields("SC09"), RsTemp.Fields("SC10"), pa01, pa02, pa03, pa04, , , RsTemp.Fields("SC11")) Then
'             End If
'          End If
'          strSql = "UPDATE staff_calendar SET sc17='" & strUserNum & "',sc18=" & strSrvDate(1) & ",sc19=" & CNULL(Mid(Right("000000" & ServerTime, 6), 1, 4), True) & _
'                   " where sc01=" & RsTemp.Fields("SC01") & " and sc02=" & RsTemp.Fields("SC02")
'          cnnConnection.Execute strSql, intI
'       End If
'   End If
'   'end 2015/12/31
'
'   'Added by Lydia 2020/10/14 Murgitroyd呈送期限設定: 中間程序報告(申復205、再審107): 收到代理人指示7日內完成並請款報告。
'   'Modified by Lydia 2020/10/19 增加判斷系統別和代理人編號; ex: P-109161的申復AA9043040進度備註有誤
'   'If CP10 = "205" Or CP10 = "107" Then
'   'Modified by Lydia 2021/01/06 +非新案收文 txtCode(0) <> ""
'   If (pa01 = "P" Or pa01 = "FCP") And (CP10 = "205" Or CP10 = "107") And pa75 <> "" And txtCode(0) <> "" Then
'       strExc(0) = Pub_GetSpecMan("外專MURGITROYD設定")
'       If strExc(0) <> "" And InStr(strExc(0), ChangeCustomerL(pa75)) > 0 Then
'           strExc(1) = CompWorkDay(1, CompDate(2, 7, DBDATE(txtPatent(0))))
'           'Added by Lydia 2022/03/17 若指定送件的日大於本所期限，請以本所期限為準
'           If txtPatent(14) <> "" Then
'              If strExc(1) > TransDate(txtPatent(14), 2) Then
'                  strExc(1) = TransDate(txtPatent(14), 2)
'              End If
'           End If
'           'end 2022/03/17
'           '自動帶至收文那道進度備註：為Murgitroyd案需xx月xx日（收文日＋7個日曆天，若為假日則抓下一個工作天）完成送件並報告
'           strExc(4) = "為Murgitroyd案需" & ChangeWStringToTDateString(strExc(1)) & "完成送件並報告"
'           strExc(3) = PUB_GetFCPHandler(pa01, pa02, pa03, pa04)
'           'Modified by Lydia 2021/05/20 在行事曆事由增加[解除管制不通知]，排除" 解除人員非建立行事曆人員會發email通知建立人員"行事曆已被解除管制"
'           If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(3), strExc(4) & "[解除管制不通知]", strExc(3), "1", pa01, pa02, pa03, pa04) = True Then
'               'Modified by Lydia 2020/10/26 承辦期限：設收文日+7個日曆天，若為假日則抓下一個工作天。
'               'strSql = "Update CaseProgress set cp64=" & CNULL(strExc(4)) & "||';'||cp64 where cp09=" & CNULL(CP09)
'               strSql = "Update CaseProgress set cp64=" & CNULL(strExc(4)) & "||';'||cp64, cp48=" & strExc(1) & " where cp09=" & CNULL(CP09)
'               cnnConnection.Execute strSql
'           End If
'       End If
'   End If
'   'end 2020/10/14
'   'Added by Lydia 2021/01/06 同一天收文提申後之實體審查+主動修正之案件，比照中間程序之內部控管。
'   If (pa01 = "P" Or pa01 = "FCP") And (CP10 = "416" Or CP10 = "203") And pa75 <> "" And txtCode(0) <> "" Then
'       strExc(0) = Pub_GetSpecMan("外專MURGITROYD設定")
'       If strExc(0) <> "" And InStr(strExc(0), ChangeCustomerL(pa75)) > 0 Then
'           strExc(1) = "select pa10,cp09 from patent,caseprogress where pa01='" & pa01 & "' and pa02='" & pa02 & "' and pa03='" & pa03 & "' and pa04='" & pa04 & "' " & _
'                            "and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp05(+)=" & CNULL(DBDATE(txtPatent(0))) & " and cp10(+)=" & CNULL(IIf(CP10 = "416", "203", "416")) & " and cp158(+)=0 and cp159(+)=0"
'           intI = 1
'           Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
'           If intI = 1 Then
'               strExc(9) = "" & RsTemp.Fields("cp09")
'               If "" & RsTemp.Fields("pa10") <> "" And strExc(9) <> "" Then
'                    '於最後收文之實體審查 or主動修正時，才產生行事曆並且一併更新進度備註和承辦期限。
'                    strExc(1) = CompWorkDay(1, CompDate(2, 7, DBDATE(txtPatent(0))))
'                    'Added by Lydia 2022/03/17 若指定送件的日大於本所期限，請以本所期限為準
'                    If txtPatent(14) <> "" Then
'                       If strExc(1) > TransDate(txtPatent(14), 2) Then
'                           strExc(1) = TransDate(txtPatent(14), 2)
'                       End If
'                    End If
'                    'end 2022/03/17
'                    strExc(4) = "為Murgitroyd案需" & ChangeWStringToTDateString(strExc(1)) & "完成送件並報告"
'                    strExc(3) = PUB_GetFCPHandler(pa01, pa02, pa03, pa04)
'                    'Modified by Lydia 2021/06/21 在行事曆事由增加[解除管制不通知]，排除" 解除人員非建立行事曆人員會發email通知建立人員"行事曆已被解除管制"
'                    If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(3), strExc(4) & "[解除管制不通知]", strExc(3), "1", pa01, pa02, pa03, pa04) = True Then
'                        strSql = "Update CaseProgress set cp64=" & CNULL(strExc(4)) & "||';'||cp64, cp48=" & strExc(1) & " where cp09=" & CNULL(CP09)
'                        cnnConnection.Execute strSql
'                        strSql = "Update CaseProgress set cp64=" & CNULL(strExc(4)) & "||';'||cp64, cp48=" & strExc(1) & " where cp09=" & CNULL(strExc(9))
'                        cnnConnection.Execute strSql
'                    End If
'               End If
'           End If
'       End If
'   End If
'   'end 2021/01/06
'
'    'Added by Lydia 2020/03/30 FCP案和FMP案: 因為分割案也有中說檔, 所以改在最前面新增收文D類English_Vers和專利
'    If strSrvDate(1) >= XY特殊權限啟用日by檔案 And frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 _
'        And (pa01 = "FCP" Or (pa01 = "P" And Left(m_SalesST15, 1) = "F")) Then
'        strExc(1) = pa01: strExc(2) = pa02: strExc(3) = pa03: strExc(4) = pa04
'        strExc(5) = "": strExc(6) = "": strExc(7) = ""
'        If PUB_ChkCPExist(strExc, cntEnglish_Vers, , , , "D") = False Then
'              strExc(0) = AutoNo("D", 6)
'              strExc(6) = PUB_GetFCPSalesNo(pa01, pa02, pa03, pa04)   'FCP承辦
'              strExc(5) = GetSalesArea(strExc(6))
'              strExc(7) = PUB_GetFCPHandler(pa01, pa02, pa03, pa04) 'FCP程序
'
'              strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
'                 ",cp12,cp13,cp14,cp20,cp26,cp27,cp32 ) values ('" & pa01 & "','" & pa02 & "','" & pa03 & "','" & pa04 & "',19221111,'" & strExc(0) & "','" & cntEnglish_Vers & "' " & _
'                  ",'" & strExc(5) & "','" & strExc(6) & "','" & strExc(7) & "','N','N',19221111,'N')"
'              cnnConnection.Execute strSql
'        End If
'        If PUB_ChkCPExist(strExc, cnt專利案件, , , , "D") = False Then
'              strExc(0) = AutoNo("D", 6)
'              If strExc(5) = "" Or strExc(6) = "" Or strExc(7) = "" Then
'                strExc(6) = PUB_GetFCPSalesNo(pa01, pa02, pa03, pa04)   'FCP承辦
'                strExc(5) = GetSalesArea(strExc(6))
'                strExc(7) = PUB_GetFCPHandler(pa01, pa02, pa03, pa04) 'FCP程序
'              End If
'              strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
'                 ",cp12,cp13,cp14,cp20,cp26,cp27,cp32 ) values ('" & pa01 & "','" & pa02 & "','" & pa03 & "','" & pa04 & "',19221111,'" & strExc(0) & "','" & cnt專利案件 & "' " & _
'                  ",'" & strExc(5) & "','" & strExc(6) & "','" & strExc(7) & "','N','N',19221111,'N')"
'              cnnConnection.Execute strSql
'        End If
'    End If
'    'end 2020/03/30
'   'Added by Lydia 2017/11/14 FCP案件命名電子化：中說輸入相關設定-存檔
'   'Modified by Lydia 2019/06/11 判斷走命名流程才檢查;
'   'If fraTCT.Visible = True And fraTCT.Enabled = True Then
'    'Modified by Lydia 2019/07/04 分割案不走命名流程，但是要能勾選其他收文
'   'If fraTCT.Visible = True And fraTCT.Enabled = True And InStr(FcpAddTct, txtPatent(1)) > 0 Then
'   If fraTCT.Visible = True And fraTCT.Enabled = True And InStr(AddTrackingNo, txtPatent(1)) > 0 Then
'      'Added by Lydia 2018/03/07 在櫃檯收文新案時若是跑工程師命名流程的新案(101~103)，將新案(101~103)分案自動上"Y"
'      If frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 And txtData(2).Text <> "" And txtData(2).Text <> "B" Then
'          'Modified by Lydia 2018/03/09 只要上已分案,Trigger會更新分案日期
'          'strSql = "update caseprogress set cp149=" & strSrvDate(1) & "  where cp09='" & CP09 & "' "
'          strSql = "update caseprogress set cp122='Y'  where cp09='" & CP09 & "' "
'          cnnConnection.Execute strSql, intI
'      End If
'      'end 2018/03/07
'      'Modified by Lydia 2022/09/01 改成共用模組
'      'Call UpdateTCTrecord(pa01, pa02, pa03, pa04, IIf(m_TCT01 <> "", m_TCT01, CP09))
'      '勾選的收文性質
'      strExc(1) = "": strExc(2) = ""
'      If ChkAdd416.Visible = True And ChkAdd416.Value = True Then strExc(1) = strExc(1) & "416,"
'      If ChkAdd203.Visible = True And ChkAdd203.Value = True Then strExc(1) = strExc(1) & "203,"
'      If ChkAdd902.Visible = True And ChkAdd902.Value = True Then strExc(1) = strExc(1) & "902,"
'      If ChkAdd924.Visible = True And ChkAdd924.Value = True Then strExc(1) = strExc(1) & "924,"
'      If ChkAdd968.Visible = True And ChkAdd968.Value = True Then strExc(1) = strExc(1) & "968,"
'      If ChkAdd414.Visible = True And ChkAdd414.Value = True Then strExc(1) = strExc(1) & "414,"
'      If ChkAdd938.Visible = True And ChkAdd938.Value = True Then strExc(1) = strExc(1) & "938,"
'      If ChkAdd939.Visible = True And ChkAdd939.Value = True Then strExc(1) = strExc(1) & "939,"
'      If ChkAdd106.Visible = True And ChkAdd106.Value = True Then strExc(1) = strExc(1) & "106,"
'      If ChkAdd228.Visible = True And ChkAdd228.Value = True Then strExc(1) = strExc(1) & "228,"
'      If ChkAdd435.Visible = True And ChkAdd435.Value = True Then strExc(1) = strExc(1) & "435,"
'      Call PUB_UpdTCTrecord(Trim(txtData(3)), strExc(1), Trim(txtTCN01.Text), strExc(2), pa01, pa02, pa03, pa04, pa05, pa06, _
'               CP09, CP10, cp06, cp07, cp13, PA08, PA09, m_PA16, m_PA14, ChangeCustomerL(pa26) & ChangeCustomerL(pa27) & ChangeCustomerL(pa28) & ChangeCustomerL(pa29) & ChangeCustomerL(pa30), _
'               pa75, IIf(fraTCT.Visible = True And txtData(2).Text <> "" And txtData(2).Text <> "B", txtData(2), ""), IIf(Trim(txtData(2)) <> "", txtData(2), "B"), Trim(txtData(0)) & Trim(txtData(1)))
'      '移到外層變更
'      frm010001.lblTCT.Caption = "中說或其他收文號："
'      frm010001.lblTCTNO.Caption = strExc(2)
'      'end 2022/09/01
'   'Added by Lydia 2018/06/28 後補:急件翻譯將新案翻譯收文號回寫到翻譯費用檔和命名記錄檔(TCN14)
'   ElseIf (txtSystem = "P" Or txtSystem = "FCP") And txtCode(0) <> "" And txtPatent(1) = "201" And frm010001.intModifyKind = 0 Then
'       'Modified by Lydia 2020/01/06 不限制急件翻譯
'       'strExc(0) = "select cp09,tcn01,tcn14 from caseprogress,trackingcasename where cp01='" & PA01 & "' and cp02='" & PA02 & "' and cp03='" & PA03 & "' and cp04='" & PA04 & "' " & _
'                         "and cp31='Y' and cp09=tcn05(+) and nvl(tcn14,'N')<>'N' "
'       strExc(0) = "select cp09,tcn01,tcn14 from caseprogress,trackingcasename where cp01='" & pa01 & "' and cp02='" & pa02 & "' and cp03='" & pa03 & "' and cp04='" & pa04 & "' " & _
'                         "and cp31='Y' and cp09=tcn05(+) "
'       intI = 1
'       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'       If intI = 1 Then
'            If "" & RsTemp.Fields("tcn01") <> "" Then '有命名追蹤
'                'Added by Lydia 2020/01/06 非急件翻譯(未提申先翻譯自動勾選)
'                If "" & RsTemp.Fields("tcn14") = "" Then
'                     strSql = "Insert into TransFee(TF01,TF31) values(" & CNULL(CP09) & ", 'Y' )"
'                     cnnConnection.Execute strSql
'                Else  '急件翻譯
'                'end 2020/01/06
'                    'Added by Lydia 2019/12/23 FMP案收文，新案建檔未提申先翻譯自動勾選
'                    If txtSystem = "P" Then
'                         strSql = "update TransFee set TF01=" & CNULL(CP09) & ",TF31='Y' where TF01=" & CNULL("" & RsTemp.Fields("tcn14"))
'                    Else
'                    'end 2019/12/23
'                         strSql = "update TransFee set TF01=" & CNULL(CP09) & " where TF01=" & CNULL("" & RsTemp.Fields("tcn14"))
'                    End If 'end 2019/12/23
'                    cnnConnection.Execute strSql, intI
'                    strSql = "update TrackingCaseName set TCN14=" & CNULL(CP09) & " where TCN01=" & CNULL("" & RsTemp.Fields("tcn01"))
'                    cnnConnection.Execute strSql, intI
'                End If 'end 2020/01/06
'                'Added by Lydia 2020/08/24 FMP案預設發"未提申先翻譯"email ; 參考frm060102
'                strExc(0) = Pub_GetSpecMan("M")
'                If strExc(0) <> "" Then
'                    strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'                       " values( '" & strUserNum & "','" & strExc(0) & "',to_char(sysdate,'yyyymmdd')" & _
'                       ",to_char(sysdate,'hh24miss'),'" & pa01 & pa02 & IIf(pa03 & pa04 <> "000", pa03 & pa04, "") & " 未提申先翻譯" & "','同主旨',null)"
'                    cnnConnection.Execute strSql
'                End If
'                'end 2020/08/24
'            End If
'       End If
'   'end 2018/06/28
'   End If
'   'end 2017/11/14
'
'   'Add by Sindy 2022/6/29
'   If m_strIR01 <> "" Then
'      m_bolRecvOK = True
'      m_strMCR11 = ""
'      If m_bMRecvBatch = True Then '多案收文
'         '更新總收文號
'         strSql = "update multiCaseRecv set mcr11='" & CP09 & "'" & _
'                  " where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
'                  " and mcr02='" & pa01 & "' and mcr03='" & pa02 & "' and mcr04='" & pa03 & "' and mcr05='" & pa04 & "'" & _
'                  " and mcr06='" & CP10 & "'"
'                  cnnConnection.Execute strSql
'
'         'Modify By Sindy 2022/8/26
'         '下載信件檔,上傳卷宗區
'         Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, CP09)
'
'         '檢查多案收文狀況
'         strExc(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
'                     " and mcr02||mcr03||mcr04||mcr05<>'" & pa01 & pa02 & pa03 & pa04 & "'" & _
'                     " and mcr11 is null"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            m_bolRecvOK = False '尚有未收文
'
'            'Modify By Sindy 2022/8/26 此處Mark,程式往上移
''            '下載信件檔,上傳卷宗區
''            Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, CP09)
'         Else
'            m_bolRecvOK = True '全部收完文
'            '抓第一筆的總收文號
'            strExc(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
'                        " and mcr02||mcr03||mcr04||mcr05=mcr07||mcr08||mcr09||mcr10 and mcr11 is not null"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               m_strMCR11 = RsTemp.Fields("mcr11")
'            Else
'               MsgBox "多案收文，無讀取到第一筆案件的總收文號，請洽電腦中心!!", vbExclamation '此狀況應不會發生, 以防外一
'               GoTo ErrHand
'            End If
'         End If
'      End If
'      If m_bolRecvOK = True Then '全部收完文
'         '多案收文的總收文號要傳入第一筆總收文號
'         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, _
'               IIf(m_strMCR11 <> "", "多案收文", "frm010001"), _
'               IIf(Left(Pub_StrUserSt03, 2) = "F2", IIf(m_strMCR11 <> "", m_strMCR11, CP09), "")
'      End If
'   End If
'   '2022/6/29 END
'
'   If bolError Then
'      cnnConnection.RollbackTrans
'      ShowMsg MsgText(9004)
'      'add by nickc 2007/12/12
'      IsSaveData = False
'   Else
'      cnnConnection.CommitTrans
'      InsertPatentDatabase = True
'      'add by nickc 2006/03/27
'      txtCode(0) = pa02
'   End If
'   'add by nickc 2005/08/12
'   txtCode(0) = pa02
'   'edit by nickc 2007/02/06 不用 dll 了
'   'Set objPublicData = Nothing
'
'   Exit Function
'
'ErrHand:
'   If PUB_CheckFormExist("frmpic002") = True Then Unload frmpic002 'Add By Sindy 2022/7/11
'
'   'edit by nickc 2007/02/06 不用 dll 了
'   'Set objPublicData = Nothing
'   cnnConnection.RollbackTrans
'   'edit by nickc 2006/03/07 解決 cp02=null 的問題
'   'add by nickc 2005/08/25
'   'txtCode(0) = ""
'   ShowMsg MsgText(9004)
'   'add by nickc 2007/12/12
'   IsSaveData = False
'End Function

'讀取Patent資料庫
'Modify by Morgan 2006/6/23 加cp89,cp90,cp91,cp92
'edit by nickc 2007/03/27 加入彼所案號
'Private Function ReadPatentDatabase(ByRef intModifyKind As Integer, ByRef PA01 As String, _
             ByRef pa02 As String, ByRef pa03 As String, ByRef pa04 As String, ByRef pa05 As String, _
             ByRef pa06 As String, ByRef pa07 As String, ByRef PA08 As String, ByRef PA09 As String, _
             ByRef pa26 As String, ByRef pa27 As String, ByRef pa28 As String, ByRef pa29 As String, _
             ByRef pa30 As String, ByRef pa75 As String, ByRef CP09 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, ByRef cp11 As String, _
             ByRef cp13 As String, ByRef cp16 As String, ByRef cp17 As String, _
             ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, ByRef cp56 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef CP64 As String, ByRef CP89 As String, ByRef CP90 As String, ByRef CP91 As String, ByRef CP92 As String) As Boolean
'Modified by Lydia 2017/11/14 +pa150
'Modified by Lydia 2021/04/15 +PA91案件備註
'Modified by Sindy 2022/12/7 +PA178證書形式
Private Function ReadPatentDatabase(ByRef intModifyKind As Integer, ByRef pa01 As String, _
             ByRef pa02 As String, ByRef pa03 As String, ByRef pa04 As String, ByRef pa05 As String, _
             ByRef pa06 As String, ByRef pa07 As String, ByRef PA08 As String, ByRef PA09 As String, _
             ByRef pa26 As String, ByRef pa27 As String, ByRef pa28 As String, ByRef pa29 As String, _
             ByRef pa30 As String, ByRef pa75 As String, ByRef CP09 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, ByRef cp11 As String, _
             ByRef cp13 As String, ByRef cp16 As String, ByRef cp17 As String, _
             ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, ByRef cp56 As String, ByRef cu30 As String, _
             ByRef cp14 As String, ByRef CP64 As String, ByRef CP89 As String, ByRef CP90 As String, ByRef CP91 As String, _
             ByRef CP92 As String, ByRef PA77 As String, ByRef CP150 As String, ByRef PA150 As String, ByRef PA91 As String, ByRef PA178 As String) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset, strTemp As String

'Add by Morgan 2004/4/15
'收據號碼
Dim stCP60 As String

On Error GoTo ErrHand
   If intModifyKind <> 0 Then
      'Modify by Morgan 2004/4/15
      '加收據號碼
      'strSQL = "select cp05,cp06,cp07,cp10,cp11,cp13,cp16,cp17,cp18,cp19,cp32,cp56,cp14 from caseprogress where cp09='" + cp09 + "'"strSQL = "select cp05,cp06,cp07,cp10,cp11,cp13,cp16,cp17,cp18,cp19,cp32,cp56,cp14 from caseprogress where cp09='" + cp09 + "'"
      'Modify by Morgan 2005/12/13 加cp33,cp34
      'Modify by Morgan 2006/6/23 加cp89,cp90,cp91,cp92
      'Modify By Sindy 2009/07/06
      'strSQL = "select cp05,cp06,cp07,cp10,cp11,cp13,cp16,cp17,cp18,cp19,cp32,cp56,cp14, cp60,cp33,cp34,cp89,cp90,cp91,cp92 from caseprogress where cp09='" + CP09 + "'"
      'Modified by Morgan 2012/4/25 +cp71
      'Modified by Morgan 2012/6/20 +cp118
      'Add by Lydia 2014/10/31 開放外專程序 =>讀cp31
      strSql = "select cp05,cp06,cp07,cp10,cp11,cp13,cp16,cp17,cp18,cp19,cp32,cp56,cp14, cp60,cp33,cp34,cp89,cp90,cp91,cp92,cp53,cp54,cp01,cp02,cp03,cp04,cp71,cp118,cp31,cp150 from caseprogress where cp09='" + CP09 + "'"
      rsRecordset.CursorLocation = adUseClient
      rsRecordset.Open strSql, cnnConnection
      If rsRecordset.RecordCount > 0 Then
         If rsRecordset.Fields("cp118") = "Y" Then
            chkWebApp.Value = 1
         End If
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
         cp56 = IIf(IsNull(rsRecordset.Fields(11)), "", rsRecordset.Fields(11))
         cp14 = IIf(IsNull(rsRecordset.Fields(12)), "", rsRecordset.Fields(12))
         'Add by Morgan 2005/12/13
         douStPrice = Val("" & rsRecordset("CP33"))
         douLowPrice = Val("" & rsRecordset("CP34"))
         '2005/12/13 end
         
         'Add By Sindy 2009/07/06
         m_CP01 = IIf(IsNull(rsRecordset.Fields("CP01")), "", rsRecordset.Fields("CP01"))
         m_CP02 = IIf(IsNull(rsRecordset.Fields("CP02")), "", rsRecordset.Fields("CP02"))
         m_CP03 = IIf(IsNull(rsRecordset.Fields("CP03")), "", rsRecordset.Fields("CP03"))
         m_CP04 = IIf(IsNull(rsRecordset.Fields("CP04")), "", rsRecordset.Fields("CP04"))
         m_CP10 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
         If (m_CP01 = "P" And m_CP10 = "601") Or _
            ((m_CP01 = "P" Or m_CP01 = "CFP") And (m_CP10 = "605" Or m_CP10 = "606" Or m_CP10 = "607")) Then
            textYear.Visible = True
            Text1(0).Visible = True
            Label11(1).Visible = True
            If m_CP10 = "605" Or m_CP10 = "601" Then
               Label11(1).Caption = "繳費年度：第            年至第            年"
            Else
               Label11(1).Caption = "繳費次數：第            次至第            次"
            End If
            textYear.Text = "" & rsRecordset("CP53")
            Text1(0).Text = "" & rsRecordset("CP54")
         End If
         
         'Add by Morgan 2004/4/15
         stCP60 = "" & rsRecordset.Fields("cp60")
         If stCP60 <> "" Then
            txtPatent(17).Enabled = False: txtPatent(18).Enabled = False: txtPatent(21).Enabled = False
            txtPatent(15).Enabled = False  '2010/10/19 ADD BY SONIA 加鎖智權人員
         End If
         'Add by Morgan 2006/6/23
         CP89 = "" & rsRecordset.Fields("cp89")
         CP90 = "" & rsRecordset.Fields("cp90")
         CP91 = "" & rsRecordset.Fields("cp91")
         CP92 = "" & rsRecordset.Fields("cp92")
         txtCopy = Val("" & rsRecordset.Fields("cp71")) 'Added by Morgan 2012/4/25
         CP150 = "" & rsRecordset.Fields("cp150") 'Add By Sindy 2012/11/08
      Else
         ShowMsg MsgText(1502)
         rsRecordset.Close
         Exit Function
      End If
      rsRecordset.Close
'Remove by Morgan 2007/1/25 搬到下面這樣才會有PA09
'   Else
'      'Modify by Morgan 2006/5/4
'      'FCP的補文件202不要做
'      If Not (PA01 = "FCP" And CP10 = "202") Then
'         If GetNextProgressDate(PA01, pa02, pa03, pa04, CP10, cp06, cp07, CP64, cp13) = False Then
'            Exit Function
'         End If
'      End If
'      'end 2006/5/4
   End If
   
   
   ' 91.09.11 modify by louis 增加准駁欄位
   'strSQL = "select pa05,pa06,pa07,pa08,pa09,pa26,pa27,pa28,pa29,pa30,pa75 " + _
   '       "from patent where pa01=" + CNULL(PA01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)
    'Modify By Cheng 2003/08/28
'   strSQL = "select pa05,pa06,pa07,pa08,pa09,pa26,pa27,pa28,pa29,pa30,pa75,pa16 " + _
'          "from patent where pa01=" + CNULL(PA01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)
   'edit by nickc 2007/03/27 加入彼所案號
   'strSQL = "select pa05,pa06,pa07,pa08,pa09,pa26,pa27,pa28,pa29,pa30,pa75,pa16, PA48, PA47 " + _
          "from patent where pa01=" + CNULL(PA01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)
'   strSQL = "select pa05,pa06,pa07,pa08,pa09,pa26,pa27,pa28,pa29,pa30,pa75,pa16, PA48, PA47,pa77,pa149 " + _
'          "from patent where pa01=" + CNULL(PA01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)

   'Modify By Sindy 2010/3/8 增加PA51~PA56
   '2010/8/18 modify by sonia 加pa14
   'Modify By Sindy 2010/10/28 增加pa158
   'Modified by Lydia 2017/11/14 +pa150
   'Modified by Lydia 2021/04/15 +PA91案件備註
   'Modified by Sindy 2022/12/7 +PA178證書形式
   strSql = "select pa05,pa06,pa07,pa08,pa09,pa26,pa27,pa28,pa29,pa30,pa75,pa16, PA48, PA47,pa77,pa149,PA51,PA52,PA53,PA54,PA55,PA56,pa14,pa158,pa150,pa91,PA178 " + _
          "from patent where pa01=" + CNULL(pa01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      pa05 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
      pa06 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
      pa07 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
      PA08 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
      PA09 = IIf(IsNull(rsRecordset.Fields(4)), "", rsRecordset.Fields(4))
      pa26 = IIf(IsNull(rsRecordset.Fields(5)), "", rsRecordset.Fields(5))
      pa27 = IIf(IsNull(rsRecordset.Fields(6)), "", rsRecordset.Fields(6))
      pa28 = IIf(IsNull(rsRecordset.Fields(7)), "", rsRecordset.Fields(7))
      pa29 = IIf(IsNull(rsRecordset.Fields(8)), "", rsRecordset.Fields(8))
      pa30 = IIf(IsNull(rsRecordset.Fields(9)), "", rsRecordset.Fields(9))
      pa75 = IIf(IsNull(rsRecordset.Fields(10)), "", rsRecordset.Fields(10))
      'add by nickc 2007/03/27 加入彼所案號
      PA77 = IIf(IsNull(rsRecordset.Fields("pa77")), "", rsRecordset.Fields("pa77"))
      
      'Add By Sindy 2010/3/8
      strPA51s = IIf(IsNull(rsRecordset.Fields("PA51")), "", rsRecordset.Fields("PA51"))
      strPA52s = IIf(IsNull(rsRecordset.Fields("PA52")), "", rsRecordset.Fields("PA52"))
      strPA53s = IIf(IsNull(rsRecordset.Fields("PA53")), "", rsRecordset.Fields("PA53"))
      strPA54s = IIf(IsNull(rsRecordset.Fields("PA54")), "", rsRecordset.Fields("PA54"))
      strPA55s = IIf(IsNull(rsRecordset.Fields("PA55")), "", rsRecordset.Fields("PA55"))
      strPA56s = IIf(IsNull(rsRecordset.Fields("PA56")), "", rsRecordset.Fields("PA56"))
      '2010/3/8 End
      
      'Add By Sindy 2010/10/28
      If IsNull(rsRecordset.Fields("PA158")) Then
         Combo3 = ""
      Else
        'Modified by Lydia 2019/05/14  區分專利種類
         'Combo3 = rsRecordset.Fields("PA158") + "." + PUB_GetCaseAttributeName(rsRecordset.Fields("PA158"))
         Combo3 = rsRecordset.Fields("PA158") + "." + PUB_GetCaseAttributeName(rsRecordset.Fields("PA158"), PA08)
      End If
      '2010/10/28 End
      Combo3.Tag = PA08 'Added by Lydia 2018/11/22
      
      PA150 = "" & rsRecordset.Fields("PA150") 'Added by Lydia 2017/11/14
      PA91 = "" & rsRecordset.Fields("PA91")  'Added by Lydia 2021/04/15 +PA91案件備註
      PA178 = "" & rsRecordset.Fields("PA178") 'Add by Sindy 2022/12/7 +PA178證書形式
'      'add by Toni 2008/08/27 加入發明人
      'for Toni
      'Modify By Sindy 2014/11/6 Mark
'      Dim i As Integer
'      strInventorNo = ""
'      For i = 16 To 25
'         If rsRecordset.Fields(i) <> "" Then
'            strInventorNo = strInventorNo & rsRecordset.Fields(i) & ","
'         End If
'      Next
'      If Right(strInventorNo, 1) = "," Then strInventorNo = Left(strInventorNo, Len(strInventorNo) - 1)
      
      'Add by Morgan 2008/8/5
      strAppNo1 = "" & rsRecordset("pa26")
      'Modify by Amy 2022/11/10 改成Form 2.0
      'PUB_AddContact strAppNo1, cboContact, "" & rsRecordset("pa149"), True
      'end 2008/8/5
      strExc(10) = cboContact.Tag
      PUB_AddContact strAppNo1, cboContact, "" & rsRecordset("pa149"), True, True, strExc(10)
      cboContact.Tag = strExc(10)
      'end 2022/11/10
   
      m_PA16 = "": m_PA14 = "" '2010/8/17 ADD BY SONIA
      
      ' 91.09.11 modify by louis
      If Not IsNull(rsRecordset.Fields("PA16")) Then
         m_PA16 = rsRecordset.Fields("PA16")
      End If
      '2010/8/17 ADD BY SONIA
      If Not IsNull(rsRecordset.Fields("PA14")) Then
         m_PA14 = rsRecordset.Fields("PA14")
      End If
      '2010/8/17 END
        'Add By Cheng 2003/08/28
        Me.txtPatent(25).Text = "" & rsRecordset("PA48").Value
        Me.txtPatent(26).Text = "" & rsRecordset("PA47").Value
      '若有申請人
      If Len("" & pa26) > 0 Then
         rsRecordset.Close
         strSql = "select cu30 from customer where cu01=" + CNULL(Mid(pa26, 1, 8)) + " AND cu02=" + CNULL(Mid(pa26, 9, 1))
         rsRecordset.CursorLocation = adUseClient
         rsRecordset.Open strSql, cnnConnection
         If rsRecordset.RecordCount > 0 Then
            cu30 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
            cp56 = ChangeCustomerS(cp56)
            pa26 = ChangeCustomerS(pa26)
            pa27 = ChangeCustomerS(pa27)
            pa28 = ChangeCustomerS(pa28)
            pa29 = ChangeCustomerS(pa29)
            pa30 = ChangeCustomerS(pa30)
            pa75 = ChangeCustomerS(pa75)
            ReadPatentDatabase = True
         Else
            ShowMsg MsgText(1503)
            Exit Function
         End If
      Else
         ReadPatentDatabase = True
      End If
            
      'Modify by Morgan 2007/1/25 從上面移下來這樣才有PA09
      If intModifyKind = 0 Then
         'FCP的補文件202不要做
         If Not (pa01 = "FCP" And CP10 = "202") Then
            'Added by Lydia 2020/11/19 CFP英國脫歐案管制：改抓歐盟案之下一程序性質
            If pa01 = "CFP" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" Then
                If GetNextProgressDate(pa01, pa02, pa03, pa04, "613", cp06, cp07, CP64, cp13, PA09) = False Then
                   Exit Function
                End If
            Else
            'end 2020/11/19
                If GetNextProgressDate(pa01, pa02, pa03, pa04, CP10, cp06, cp07, CP64, cp13, PA09) = False Then
                   Exit Function
                End If
            End If 'Added by Lydia 2020/11/19
         End If
      End If
      'end 2007/1/25
      
   Else
      If intModifyKind <> 0 Then
         ShowMsg "找不到此本所案號在專利基本檔之資料"
         Exit Function
      End If
   End If
   If cp06 <> "" Then cp06 = ChangeWStringToTString(cp06)
   If cp07 <> "" Then cp07 = ChangeWStringToTString(cp07)
   rsRecordset.Close
   
   'Add By Sindy 2014/11/6 讀取發明人資料
   'Modified by Lydia 2022/08/22 改用模組
   'strSql = "select * from PatentInventor where pi01=" + CNULL(pa01) + " and pi02=" + CNULL(pa02) + " and pi03=" + CNULL(pa03) + " and pi04=" + CNULL(pa04) & _
            " order by pi05 asc"
   'rsRecordset.CursorLocation = adUseClient
   'rsRecordset.Open strSql, cnnConnection
   'strInventorNo = ""
   'strInventorNo_Old = ""
   'If rsRecordset.RecordCount > 0 Then
   '   rsRecordset.MoveFirst
   '   Do While Not rsRecordset.EOF
   '      strInventorNo = strInventorNo & rsRecordset.Fields("pi06") & ","
   '      rsRecordset.MoveNext
   '   Loop
   '   If Right(strInventorNo, 1) = "," Then strInventorNo = Left(strInventorNo, Len(strInventorNo) - 1)
   '   strInventorNo_Old = strInventorNo
   'End If
   'rsRecordset.Close
   ''2014/11/6 END
   strInventorNo = PUB_GetPatentInventorList(pa01, pa02, pa03, pa04)
   strInventorNo_Old = strInventorNo
   'end 2022/08/22
   
   'add by nickc 2008/05/02 抓預定收款日
   'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'   strSql = "select rd05 from ReceivablesDay where (rd01,rd02,rd03) in (select rd01,rd02,max(rd03) from ReceivablesDay where (rd02) in (select max(rd02) from ReceivablesDay where rd01='" & CP09 & "' ) and rd01='" & CP09 & "' group by rd01,rd02) "
'   rsRecordset.CursorLocation = adUseClient
'   rsRecordset.Open strSql, cnnConnection
'   If rsRecordset.RecordCount > 0 Then
'      txtPatent(28) = IIf(IsNull(rsRecordset.Fields(0)), "", TAIWANDATE(rsRecordset.Fields(0)))
'   Else
'      txtPatent(28) = ""
'   End If
'   txtPatent(28).Tag = txtPatent(28) 'Add by Morgan 2010/12/9
'   rsRecordset.Close
   'end 2018/08/22
   
'Add by Amy 2013/06/26 FCP/P 新申請案及衍生設計101/102/103/125 則抓取案件命名追蹤流水號
'Modified by Lydia 2018/09/06 改成變數
'If strSrvDate(2) >= Val("1020723") And (txtSystem = "FCP" Or txtSystem = "P") And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "125") Then
If txtTCN01.Visible = True And (txtSystem = "FCP" Or txtSystem = "P") And InStr(AddTrackingNo, txtPatent(1)) > 0 Then
   'Modified by Lydia 2024/12/13
   'txtTCN01 = GetTCN01
   txtTCN01 = Pub_GetTCN01(txtRecieveCode)
End If
'end 2013/06/26


   Exit Function
ErrHand:
   ShowMsg "資料讀取失敗,請洽系統管理者!"  '2010/8/18 add by sonia
End Function

'修改專利資料庫
'edit by nickc 2007/03/27 加入彼所案號
'Modified by Morgan 2021/7/21 +PA176
'Removed by Morgan 2024/11/18 收文存檔模組已啟用,舊程式標記為註解,後續無需再修改
'Private Function UpdatePatentDatabase(ByRef intSaveMode As Integer, ByRef pa01 As String, _
'             ByRef pa02 As String, ByRef pa03 As String, ByRef pa04 As String, ByRef pa05 As String, _
'             ByRef pa06 As String, ByRef pa07 As String, ByRef PA08 As String, ByRef PA09 As String, _
'             ByRef pa26 As String, ByRef pa27 As String, ByRef pa28 As String, _
'             ByRef pa29 As String, ByRef pa30 As String, ByRef pa75 As String, _
'             ByRef CP09 As String, _
'             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, _
'             ByRef cp11 As String, ByRef cp13 As String, ByRef cp16 As String, _
'             ByRef cp17 As String, ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, _
'             ByRef cp56 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef cp33 As Double, _
'             ByRef cp34 As Double, ByRef CP64 As String, ByRef CP89 As String, ByRef CP90 As String, _
'             ByRef CP91 As String, ByRef CP92 As String, ByRef PA77 As String, ByRef PA149 As String, ByRef PA176 As String) As Boolean
'Dim strSql As String, cp55 As String, strCustomer(4) As String, i As Integer
'Dim cp93 As String, cp94 As String, cp95 As String, cp96 As String 'Add by Morgan 2006/6/26
'Dim adoquery As New ADODB.Recordset
''Dim strInventor(9) As String     'add by toni 2008/8/26 寫發明人data use
'Dim cp48 As String 'Add by Morgan 2008/8/28
'Dim stUpdate As String 'Add by Morgan 2008/8/28
'
'   'add by nickc 2007/12/12
'   If IsSaveData = True Then
'      Exit Function
'   End If
'   IsSaveData = True
'
'On Error GoTo ErrHand
'   cp05 = ChangeTStringToWString(cp05)
'   cp06 = ChangeTStringToWString(cp06)
'   cp07 = ChangeTStringToWString(cp07)
'   pa26 = ChangeCustomerL(pa26)
'   pa27 = ChangeCustomerL(pa27)
'   pa28 = ChangeCustomerL(pa28)
'   pa29 = ChangeCustomerL(pa29)
'   pa30 = ChangeCustomerL(pa30)
'   pa75 = ChangeCustomerL(pa75)
'   cnnConnection.BeginTrans
'   'edit by nickc 2007/03/27 加入彼所案號
'   'strSQL = "update patent set pa05=" + CNULL(ChgSQL(pa05)) + ",pa06=" + CNULL(Replace(pa06, "'", "''")) + _
'      ",pa07=" + CNULL(ChgSQL(pa07)) + ",pa08=" + CNULL(PA08) + ",pa09=" + CNULL(PA09) + _
'      ",pa26=" + CNULL(pa26) + ",pa27=" + CNULL(pa27) + _
'      ",pa28=" + CNULL(pa28) + ",pa29=" + CNULL(pa29) + ",pa30=" + CNULL(pa30) + _
'      ",pa75=" + CNULL(pa75) + " where pa01=" + CNULL(PA01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)
'   'strSQL = "update patent set pa05=" + CNULL(ChgSQL(pa05)) + ",pa06=" + CNULL(Replace(pa06, "'", "''")) + _
'   '   ",pa07=" + CNULL(ChgSQL(pa07)) + ",pa08=" + CNULL(PA08) + ",pa09=" + CNULL(PA09) + _
'   '   ",pa26=" + CNULL(pa26) + ",pa27=" + CNULL(pa27) + _
'   '   ",pa28=" + CNULL(pa28) + ",pa29=" + CNULL(pa29) + ",pa30=" + CNULL(pa30) + _
'   '   ",pa75=" + CNULL(pa75) + ",pa77=" + CNULL(PA77)
'
'        'Add by Lydia 2014/10/31 開放外專程序人員可進入專利處系統操作FMP寰華案件=>寫入代理人
'        If mFMPchk = True Then
'            strSql = "update caseprogress set cp44='Y53374000' where cp09='" & txtRecieveCode & "' "
'        Else
'            strSql = "update caseprogress set cp44='' where cp09='" & txtRecieveCode & "' "
'        End If
'        cnnConnection.Execute strSql
'        'end. 'Add by Lydia 2014/10/31
'
'            varInventorNo = Split(strInventorNo, ",")
'            For i = 0 To UBound(varInventorNo)
'               strInventor(i) = varInventorNo(i)
'            Next
'            For i = i + 1 To 99 '9
'               strInventor(i) = ""
'            Next
'            'Add By Sindy 2014/11/6 更新專利發明人檔
'            If strInventorNo_Old <> strInventorNo Then
'               strSql = "delete from patentInventor where pi01=" + CNULL(pa01) + " and pi02=" + CNULL(pa02) + " and pi03=" + CNULL(pa03) + " and pi04=" + CNULL(pa04)
'               Pub_SeekTbLog strSql 'Add By Sindy 2017/8/23
'               cnnConnection.Execute strSql
'               For i = 0 To 99
'                  If strInventor(i) <> "" Then
'                     strSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
'                              CNULL(pa01) & "," & CNULL(pa02) & "," & CNULL(pa03) & "," & CNULL(pa04) & "," & i + 1 & ",'" & strInventor(i) & "')"
'                     Pub_SeekTbLog strSql 'Add By Sindy 2017/8/23
'                     cnnConnection.Execute strSql
'                  Else
'                     Exit For
'                  End If
'               Next i
'            End If
'            '2014/11/6 END
'
'            'Memo by Lydia 2021/08/17 刪除舊程式碼：專利發明人在專利基本檔60~69
'            strSql = "update patent set pa05=" + CNULL(ChgSQL(pa05)) + ",pa06=" + CNULL(Replace(pa06, "'", "''")) + _
'               ",pa07=" + CNULL(ChgSQL(pa07)) + ",pa08=" + CNULL(PA08) + ",pa09=" + CNULL(PA09) + _
'               ",pa26=" + CNULL(pa26) + ",pa27=" + CNULL(pa27) + _
'               ",pa28=" + CNULL(pa28) + ",pa29=" + CNULL(pa29) + ",pa30=" + CNULL(pa30) + _
'               ",pa75=" + CNULL(pa75) + ",pa77=" + CNULL(PA77) + _
'               ",pa158=" + CNULL(Left(Combo3, 1))
'   'Add by Morgan 2008/8/5 +PA149
'   If UCase(PA149) <> "PA149" Then
'      strSql = strSql + ",PA149=" + CNULL(PA149)
'   End If
'
'   'Added by Lydia 2017/11/14 +PA150
'   If fraTCT.Visible = True And txtData(2).Text <> txtData(2).Tag Then
'      strSql = strSql + ",pa150=" + CNULL(IIf(txtData(2).Text = "B", "", txtData(2).Text))
'   End If
'   'end 2017/11/14
'
'   'Added by Morgan 2021/7/21
'   'Modified by Lydia 2022/08/22 debug
'   'If PA149 <> "" Then
'   If PA176 <> "" Then
'      strSql = strSql + ",PA176='" & PA176 & "'"
'   End If
'   'end 2021/7/21
'
'   'Add By Sindy 2010/3/8 增加聯絡人pa51~pa56欄位
'   If bolCancel = True Then
'      strSql = strSql + ",PA51=" + CNULL(ChgSQL(strPA51s)) + ",PA52=" + CNULL(ChgSQL(strPA52s)) + _
'                                ",PA53=" + CNULL(ChgSQL(strPA53s)) + ",PA54=" + CNULL(ChgSQL(strPA54s)) + _
'                                ",PA55=" + CNULL(ChgSQL(strPA55s)) + ",PA56=" + CNULL(ChgSQL(strPA56s))
'   End If
'   strSql = strSql + " where pa01=" + CNULL(pa01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)
'   cnnConnection.Execute strSql
'   strCustomer(0) = pa26
'   strCustomer(1) = pa27
'   strCustomer(2) = pa28
'   strCustomer(3) = pa29
'   strCustomer(4) = pa30
'
'   For i = 0 To 4
'          strSql = "update patent set pa" + Format(31 + i) + "=(select cu23 from customer where cu01=" + CNULL(Mid(strCustomer(i), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(i), 9, 1)) + _
'             "),pa" + Format(36 + i) + "=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(strCustomer(i), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(i), 9, 1)) + _
'             "),pa" + Format(41 + i) + "=(select cu29 from customer where cu01=" + CNULL(Mid(strCustomer(i), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(i), 9, 1)) + ") where pa01=" + CNULL(pa01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)
'          cnnConnection.Execute strSql
'   Next
'   'Add By Cheng 2003/08/28
'   'Begin
'   strSql = "Update Patent Set PA47='" & ChgSQL(Me.txtPatent(26).Text) & "', PA48='" & ChgSQL(Me.txtPatent(25).Text) & "' Where pa01=" + CNULL(pa01) + " and pa02=" + CNULL(pa02) + " and pa03=" + CNULL(pa03) + " and pa04=" + CNULL(pa04)
'   cnnConnection.Execute strSql
'   'End
'   If cp56 <> "" Then
'   '   cp55 = pa26
'      'Add by Morgan 2006/6/23
'      '讓與人2-5,受讓人2-5
'      cp56 = ChangeCustomerL(cp56)
'      CP89 = ChangeCustomerL(CP89)
'      CP90 = ChangeCustomerL(CP90)
'      CP91 = ChangeCustomerL(CP91)
'      CP92 = ChangeCustomerL(CP92)
'   '   cp93 = pa27
'   '   cp94 = pa28
'   '   cp95 = pa29
'   '   cp96 = pa30
'      'end 2006/6/23
'   End If
'
'      'Add by Morgan 2008/8/28 預設承辦期限
'      '2010/3/11 MODIFY BY SONIA
'      'If Me.txtSystem.Text = "FCP" Then
'      If Me.txtSystem.Text = "FCP" Or Me.txtSystem.Text = "FG" Then
'      'Modify by Morgan 2008/10/23
'         'Ciba Y45697的年費承辦期限掛15個工作天
'         If pa75 = "Y45697000" And CP10 = "605" Then
'            cp48 = CompWorkDay(15, strSrvDate(1))
'            If Val(cp06) > 0 And Val(cp48) > Val(cp06) Then
'               cp48 = cp06
'            End If
'
'         'Added by Morgan 2012/7/13
'         '加速審查要判斷已輸入通知實審日才掛承辦期限
'         ElseIf CP10 = "422" Then
'            strExc(1) = pa01
'            strExc(2) = pa02
'            strExc(3) = pa03
'            strExc(4) = pa04
'            If PUB_ChkCPExist(strExc(), "1204") Then
'               cp48 = Pub_GetHandleDay("FCP", "000", CP10, , cp06)
'            End If
'         'end 2012/7/13
'
'         'Add By Sindy 2021/6/24 968回復說明書校閱
'         ElseIf CP10 = "968" Then
'            cp48 = Pub_GetHandleDay("FCP", "000", CP10, , cp06, , , pa01 & "-" & pa02 & "-" & pa03 & "-" & pa04)
'
'         'end 2008/10/23
'         ElseIf InStr(SkipCasePtyList, CP10) = 0 Then
'            cp48 = Pub_GetHandleDay("FCP", "000", CP10, , cp06)
'         End If
'         stUpdate = ",cp48=" & CNULL(cp48, True)
'         'Y54732000 & X30299000組合之回代承辦期限下面另有更新
'
'         'Add By Sindy 2021/4/29 不是主管機關期限
'         If m_strCPM34 = "N" And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
'            '(2)收文時無設本所期限，以承辦期限＋5個工作天為本所期限
'            If Val(cp06) = 0 Then
'               cp06 = PUB_GetFCPOurDeadline(DBDATE(cp48), , , , "N")
'            '(1)收文時有設本所期限，自動備註:本所期限為yyy/mm/dd(本所期限)
'            ElseIf Val(DBDATE(m_strCP06)) <> Val(cp06) Then '有異動時
'               If InStr(CP64, "原本所期限為" & ChangeWStringToTDateString(DBDATE(m_strCP06)) & "已修改;") = 0 Then
'                  CP64 = "原本所期限為" & ChangeWStringToTDateString(DBDATE(m_strCP06)) & "已修改;" & CP64
'               End If
'            End If
'         End If
'      End If
'
'      'Added by Morgan 2012/6/20
'      If chkWebApp.Visible = True Then
'         stUpdate = ",cp118='" & IIf(chkWebApp.Value = 1, "Y", "") & "'"
'      End If
'      'end 2012/6/20
'
'   'Modify By Sindy 2009/10/19
'   'strSQL = "update caseprogress set cp05=" + CNULL(cp05) + ",cp06=" + CNULL(cp06) + ",cp07=" + CNULL(cp07) + ",cp10=" + CNULL(CP10) + _
'   '         ",cp11=" + CNULL(cp11) + ",cp13=" + CNULL(cp13) + ",cp14=" + CNULL(cp14) + ",cp16=" + CNULL(cp16) + ",cp17=" + CNULL(cp17) + _
'   '         ",cp18=" + CNULL(cp18) + ",cp19=" + CNULL(cp19) + ",cp32=" + CNULL(cp32) + ",cp55=" + CNULL(cp55) + ",cp56=" + CNULL(cp56) + _
'   '         ",cp33=" & cp33 & ",cp34=" & cp34 & ",CP64=" + CNULL(ChgSQL(CP64)) + ",cp89=" + CNULL(CP89) + ",cp90=" + CNULL(CP90) + _
'   '         ",cp91=" + CNULL(CP91) + ",cp92=" + CNULL(CP92) + ",cp93=" + CNULL(cp93) + ",cp94=" + CNULL(cp94) + _
'   '         ",cp95=" + CNULL(cp95) + ",cp96=" + CNULL(cp96) & stUpdate & " where cp09='" + CP09 + "'"
'   'Add By Sindy 2012/11/06 有★★的應收帳款簽核控管
'   m_CP150 = ""
'   If Check2.Value = 1 Then m_CP150 = "Y"
'   '2012/11/06 End
'   'Modify By Sindy 2012/11/06 +CP150
'   strSql = "update caseprogress set cp05=" + CNULL(cp05) + ",cp06=" + CNULL(cp06) + ",cp07=" + CNULL(cp07) + ",cp10=" + CNULL(CP10) + _
'            ",cp11=" + CNULL(cp11) + ",cp13=" + CNULL(cp13) + ",cp14=" + CNULL(cp14) + ",cp16=" + CNULL(cp16) + ",cp17=" + CNULL(cp17) + _
'            ",cp18=" + CNULL(cp18) + ",cp19=" + CNULL(cp19) + ",cp32=" + CNULL(cp32) + ",cp56=" + CNULL(cp56) + _
'            ",cp33=" & cp33 & ",cp34=" & cp34 & ",CP64=" + CNULL(ChgSQL(CP64)) + ",cp89=" + CNULL(CP89) + ",cp90=" + CNULL(CP90) + _
'            ",cp91=" + CNULL(CP91) + ",cp92=" + CNULL(CP92) + stUpdate & ",cp150=" & CNULL(m_CP150) & " where cp09='" + CP09 + "'"
'   cnnConnection.Execute strSql
'
'   'add by sonia 2019/7/31 Y54732000 & X30299000組合,且會稿924發文後,新案翻譯201發文前收文之回代902,設回代相關收文號掛會稿,承辦期限掛新案翻譯的本所期限
'   If pa75 = "Y54732000" And Left(pa26, 8) = "X3029900" And CP10 = "902" Then
'      strExc(0) = "select c2.cp06,c1.cp09 from caseprogress c1,caseprogress c2 where c1.cp01='" & pa01 & "' and c1.cp02='" & pa02 & "' and c1.cp03='" & pa03 & "' and c1.cp04='" & pa04 & "' and c1.cp10='924' and c1.cp27>0 " & _
'                  "   and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and '201'=c2.cp10(+) and c2.cp158=0"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         strSql = "update caseprogress set cp43='" & "" & RsTemp(1) & "',cp48=" & "" & RsTemp(0) & " where cp09=" & CNULL(CP09)
'         cnnConnection.Execute strSql
'      End If
'   End If
'   'end 2019/7/31
'
'   'Modified by Morgan 2012/4/25 +cp71(優先權份數)
'   strSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(cp13) + ")" & IIf(txtCopy.Visible, ",cp71=" & Val(txtCopy), "") & " where cp09=" + CNULL(CP09)
'   'end 2012/4/25
'   cnnConnection.Execute strSql
'
'   'Add By Sindy 2009/07/06
'   If textYear.Visible = True Then
'      If m_CP10 = "601" Then
'         If Val(Text1(0)) > 0 Then
'            strSql = "update caseprogress set cp53=" & CNULL(textYear, True) & ",cp54=" & CNULL(Text1(0), True) & " where cp09=" & CNULL(CP09)
'         End If
'      Else
'         strSql = "update caseprogress set cp53=" & CNULL(textYear, True) & ",cp54=" & CNULL(Text1(0), True) & " where cp09=" & CNULL(CP09)
'      End If
'      cnnConnection.Execute strSql
'   End If
'   '2009/07/06 End
'   'Added by Lydia 2022/11/29 非內部收文並且有費用，先統一設定CP20=Null ;
'   If frm010001.intChoose = 0 And Val(cp16) > 0 Then
'      If (m_PA16 = "" And InStr("FCP062174000", Me.txtSystem & Me.txtCode(0) & IIf(Me.txtCode(1) = "", "0", Me.txtCode(1)) & IIf(Me.txtCode(2) = "", "00", Me.txtCode(2))) > 0) Or _
'        (m_PA16 <> "1" And InStr("FCP067004000", Me.txtSystem & Me.txtCode(0) & IIf(Me.txtCode(1) = "", "0", Me.txtCode(1)) & IIf(Me.txtCode(2) = "", "00", Me.txtCode(2))) > 0) Then
'          '排除特定案件
'      Else
'          stUpdate = ""
'          If Left(m_SalesST15, 1) = "F" And (Me.txtSystem.Text = "FCP" Or Me.txtSystem.Text = "FG" Or Me.txtSystem.Text = "P") Then
'             stUpdate = PUB_GetCP20(txtSystem, CP10)
'          End If
'          If stUpdate = "" Then
'             strSql = "update caseprogress set cp20=null where cp09=" + CNULL(CP09)
'             cnnConnection.Execute strSql
'          End If
'      End If
'   End If
'   'end 2022/11/29
'
'           'Add By nickc 2007/08/21
'           '若為接洽記錄單(櫃台收文)
'           'Modify by Morgan 2007/10/26 費用可改時才做，否則已收款資料會被還原
'           'If frm010001.intChoose = 0 Then
'           If frm010001.intChoose = 0 And txtPatent(17).Enabled = True Then
'           'end 2007/10/26
'               '未收金額 = 費用
'               strSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(CP09)
'               cnnConnection.Execute strSql
'           End If
'   'Add By Cheng 2002/05/10
'   '若為內部收文作業時, 案件進度檔的是否向客戶收款設定為"N"
'   If frm010001.intChoose = 1 Then
'      strSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(CP09)
'      cnnConnection.Execute strSql
'   End If
'   strSql = "update customer set cu30=" + CNULL(cu30) + " where cu01=" + CNULL(Mid(pa26, 1, 8)) + " and cu02=" + CNULL(Mid(pa26, 9, 1))
'   cnnConnection.Execute strSql
'   UpdatePatentDatabase = True
'   adoquery.CursorLocation = adUseClient
'   'adoquery.Open "select np01 from nextprogress where np02 = '" & pa01 & "' and np03 = '" & pa02 & "' and np04 = '" & pa03 & "' and np05 = '" & pa04 & "' and np07 = '" & cp10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
'   adoquery.Open "select np01 from nextprogress where np02 = '" & pa01 & "' and np03 = '" & pa02 & "' and np04 = '" & pa03 & "' and np05 = '" & pa04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
'   'Modify By Cheng 2002/05/10
'   '若在下一程序檔只抓到一筆資料時, 才要抓下一程序檔的總收文號更新案件進度檔的相關總收文號
'   'If adoquery.RecordCount <> 0 Then
'   If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
'      If IsNull(adoquery.Fields(0).Value) = False Then
'         'Add by Morgan 2010/6/30 異議答辯、舉發答辯要一並更新對造資料
'         If (CP10 = "802" Or CP10 = "804") Then
'            cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & CP09 & "'", intI
'         Else
'         'End 2010/6/30
'            cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & CP09 & "'"
'         End If
'      End If
'   End If
'   adoquery.Close
'   'add by nickc 2008/05/02 儲存預定收款日
'   'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
''   Dim rtCnt As Integer
''   'Modify by Morgan 2010/12/9
''   'If txtPatent(28) <> "" Then
''   '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " ", rtCnt
''   If txtPatent(28) <> "" And txtPatent(28) <> txtPatent(28).Tag Then
''       cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0) + 1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
''   'end 2010/12/9
''       If rtCnt = 0 Then
''           cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from dual "
''       End If
''   End If
'   'end 2018/08/22
'
'   'Added by Lydia 2017/11/14 FCP案件命名電子化：中說輸入相關設定-存檔
'   'Modified by Lydia 2019/06/11 判斷走命名流程才檢查; FCP-62285收307分割,修改時增加申請人2-X76639會彈未輸入分案組別
'   'If fraTCT.Visible = True And fraTCT.Enabled = True Then
'   'Modified by Lydia 2019/07/04 分割案不走命名流程，但是要能勾選其他收文
'   'If fraTCT.Visible = True And fraTCT.Enabled = True And InStr(FcpAddTct, txtPatent(1)) > 0 Then
'   'Modified by Lydia 2022/09/01 改成共用模組
'   'If fraTCT.Visible = True And fraTCT.Enabled = True And InStr(AddTrackingNo, txtPatent(1)) > 0 Then
'   '   Call UpdateTCTrecord(pa01, pa02, pa03, pa04, IIf(m_TCT01 <> "", m_TCT01, CP09))
'   'End If
'      '勾選的收文性質
'      strExc(1) = "": strExc(2) = ""
'      If ChkAdd416.Visible = True And ChkAdd416.Value = True Then strExc(1) = strExc(1) & "416,"
'      If ChkAdd203.Visible = True And ChkAdd203.Value = True Then strExc(1) = strExc(1) & "203,"
'      If ChkAdd902.Visible = True And ChkAdd902.Value = True Then strExc(1) = strExc(1) & "902,"
'      If ChkAdd924.Visible = True And ChkAdd924.Value = True Then strExc(1) = strExc(1) & "924,"
'      If ChkAdd968.Visible = True And ChkAdd968.Value = True Then strExc(1) = strExc(1) & "968,"
'      If ChkAdd414.Visible = True And ChkAdd414.Value = True Then strExc(1) = strExc(1) & "414,"
'      If ChkAdd938.Visible = True And ChkAdd938.Value = True Then strExc(1) = strExc(1) & "938,"
'      If ChkAdd939.Visible = True And ChkAdd939.Value = True Then strExc(1) = strExc(1) & "939,"
'      If ChkAdd106.Visible = True And ChkAdd106.Value = True Then strExc(1) = strExc(1) & "106,"
'      If ChkAdd228.Visible = True And ChkAdd228.Value = True Then strExc(1) = strExc(1) & "228,"
'      If ChkAdd435.Visible = True And ChkAdd435.Value = True Then strExc(1) = strExc(1) & "435,"
'      Call PUB_UpdTCTrecord(Trim(txtData(3)), strExc(1), Trim(txtTCN01.Text), strExc(2), pa01, pa02, pa03, pa04, pa05, pa06, _
'               CP09, CP10, cp06, cp07, cp13, PA08, PA09, m_PA16, m_PA14, ChangeCustomerL(pa26) & ChangeCustomerL(pa27) & ChangeCustomerL(pa28) & ChangeCustomerL(pa29) & ChangeCustomerL(pa30), _
'               pa75, IIf(fraTCT.Visible = True And txtData(2).Text <> "" And txtData(2).Text <> "B", txtData(2), ""), IIf(Trim(txtData(2)) <> "", txtData(2), "B"), Trim(txtData(0)) & Trim(txtData(1)))
'   'end 2022/09/01
'
'   cnnConnection.CommitTrans
'   Exit Function
'ErrHand:
'   cnnConnection.RollbackTrans
'   ShowMsg MsgText(9004)
'   'add by nickc 2007/12/12
'   IsSaveData = False
'End Function

'從下一程序檔取回本所期限、法定期限
Private Function GetNextProgressDate(ByVal np02 As String, ByVal np03 As String, ByVal np04 As String, _
       ByVal np05 As String, ByVal NP07 As String, ByRef strDate1 As String, ByRef strDate2 As String, ByRef strNP15 As String, _
       ByRef strNP10 As String, Optional ByRef strPA09 As String) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset
Dim mStr As String 'Added by Lydia 2017/12/21

On Error GoTo ErrHand
   'NICK 900803 **********************
   'strSQL = "select np08,np09,NP15 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
   '          CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
   '          " and np07=" + CNULL(np07) + " and (np06<>'Y' or np06 is null)"
   '911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
   'strSQL = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
             CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
             " and np07=" + CNULL(NP07) + " and (np06<>'Y' or np06 is null)"
             
   strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
             CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
             " and np07=" + CNULL(NP07) + " and np06 is null "
   mStr = NP07 'Added by Lydia 2017/12/21
   'Add by Morgan 2007/1/12 台灣專利的申復或修正時下一程序兩個都要抓
   If strPA09 = "000" And (NP07 = "204" Or NP07 = "205") Then
      strExc(0) = "select * from patent where pa01='" & np02 & "' and pa02='" & np03 & "' and pa03='" & np04 & "' and pa04='" & np05 & "' and pa09='000'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
             CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
             " and np07 in ('204','205') and np06 is null "
             mStr = "'204','205'" 'Added by Lydia 2017/12/21
      End If
   End If
   'end 2007/1/12
   
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection, adOpenStatic
   If rsRecordset.RecordCount > 0 Then
      rsRecordset.MoveLast
      rsRecordset.MoveFirst
      If rsRecordset.RecordCount = 1 Then
         strDate1 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
         strDate2 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
         strNP15 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
         '取得下一程序資料檔之智權人員代號
         strNP10 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
         'Added by Lydia 2017/12/21 檢查是否有不續辦相同性質且未到期的期限，若有則提醒操作人員注意要輸入接洽單上填寫的期限
         If frm010001.mRole = "" Then 'Added by Lydia 2024/10/18 排除外專/外商自行收文
            strExc(1) = Pub_GetNPDoubleMsg(DBDATE(txtPatent(0).Text), np02, np03, np04, np05, mStr)
            If strExc(1) <> "" Then MsgBox strExc(1), vbExclamation + vbOKOnly
         End If 'Added by Lydia 2024/10/18
         'end 2017/12/21
         'Added by Lydia 2023/06/08
         If strDate1 <> "" Or strDate2 <> "" Then
            bolisNP0809 = True
         End If
         'end 2023/06/08
      End If
   Else
      '911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
      strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
                CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
               " and np07=" + CNULL(NP07) + " and np06 <>'Y' "
      'Add by Morgan 2007/1/25 台灣專利的申復或修正時下一程序兩個都要抓
      If strPA09 = "000" And (NP07 = "204" Or NP07 = "205") Then
         strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
                CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
               " and np07 in ('204','205') and np06 <>'Y' "
      End If
      'end 2007/1/25
      Set rsRecordset = New ADODB.Recordset
      rsRecordset.CursorLocation = adUseClient
      rsRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsRecordset.RecordCount = 1 Then
         strDate1 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
         strDate2 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
         strNP15 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
         '取得下一程序資料檔之智權人員代號
         strNP10 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
         'Added by Lydia 2023/06/08
         If strDate1 <> "" Or strDate2 <> "" Then
            bolisNP0809 = True
         End If
         'end 2023/06/08
      End If
   End If
   '**********************
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
   For Each objTxt In Me.txtPatent
      
      'Add By Cheng 2002/06/11
      If objTxt.Index = 24 And Me.fraPatition.Visible = False Then GoTo NextTxt
      
      '93.3.7 add by sonia
      'If objTxt.Index = 4 Then CheckKeyIn 21
      '93.3.7 end
      
      If objTxt.Enabled = True Then
         If objTxt.Index < 8 Or objTxt.Index > 12 Then 'Added by Lydia 2024/02/16 因為在cmdOK_Click已呼叫CheckKeyin檢查
            Cancel = False
            txtPatent_Validate objTxt.Index, Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If 'Added by Lydia 2024/02/16
      End If
NextTxt:
   Next

   '2008/9/10 ADD BY SONIA
   'Modified by Morgan 2012/11/8 有可能沒有故改只提醒
   If txtPatent(15).Text = "79075" And txtCode(0) = "" And txtPatent(13) = "" Then
      'MsgBox "郭雅娟案件請輸入代理人編號！"
      If MsgBox("本案為郭雅娟案件是否確定不輸入代理人編號！", vbYesNo + vbDefaultButton2) = vbNo Then
         txtPatent(13).SetFocus
         Exit Function
      End If
   End If
   If txtPatent(15).Text = "79075" And txtCode(0) = "" And txtPatent(27) = "" Then
      'MsgBox "郭雅娟案件請輸入代理人彼所案號！"
      If MsgBox("本案為郭雅娟案件是否確定不輸入彼所案號！", vbYesNo + vbDefaultButton2) = vbNo Then
         txtPatent(27).SetFocus
         Exit Function
      End If
   End If
   '2008/9/10 END

   'Added by Morgan 2012/4/25
   If txtCopy.Visible And Val(txtCopy) = 0 Then
      MsgBox "請輸入優先權份數!!", vbExclamation
      txtCopy.SetFocus
      Exit Function
   End If
   'end 2012/4/25
   
   'Add by Amy 2013/07/19 lblCaseProperty顯示（無）不可以存檔
   If lblCaseProperty = "（無）" Then
      MsgBox "案件性質錯誤!!", vbExclamation
      Exit Function
   End If
   'end 2013/07/19
   
   'Add by Amy 2013/06/26 FCP/P 新申請案101/102/103/衍生設計125 則追蹤流水號不可為空
   '2015/4/14 MODIFY BY SONIA 林總指示開放投資法務人員可收文L案及P案,故改控制外專收文
   'If frm010001.intModifyKind = 0 And strSrvDate(2) >= Val("1020723") And Left(GetST15(txtPatent(15).Text), 1) = "F" And (txtSystem = "FCP" Or txtSystem = "P") And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "125") Then
   'Modified by Lydia 2018/09/06 改成變數
   'If frm010001.intModifyKind = 0 And strSrvDate(2) >= Val("1020723") And Left(GetST15(txtPatent(15).Text), 2) = "F2" And (txtSystem = "FCP" Or txtSystem = "P") And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "125") Then
   If frm010001.intModifyKind = 0 And txtTCN01.Visible = True And Left(m_SalesST15, 2) = "F2" And (txtSystem = "FCP" Or txtSystem = "P") And InStr(AddTrackingNo, txtPatent(1)) > 0 Then
        If Trim(txtTCN01) = "" Then
            MsgBox "追蹤流水號不可為空!!", vbExclamation
            txtTCN01.SetFocus
            Exit Function
        End If
   End If
   
   Cancel = False
   txtTCN01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   'end 2013/06/26
   
   TxtValidate = True
End Function

' 91.09.11 modify by louis
' 更新畫面中費用及規費的欄位內容
'Modified  by Lydia 2018/03/01 傳案件性質算費用
Private Sub OnUpdateFee(Optional ByVal tCP10 As String = "")
Dim strSG07 As String, strSG08 As String 'Add By Sindy 2012/11/22
Dim strTmpVal As String 'Added by Lydia 2018/03/01
   
   'Modified by Lydia 2019/09/16
   'm_SalesST15 = GetST15(txtPatent(15).Text) 'Added by Lydia 2018/09/06
   m_SalesST15 = GetST15(txtPatent(15).Text, , , m_SalesST06)
   
   'Added by Lydia 2018/12/06 預設先清空費用等變數, 因為同時收文若遇到無費用的性質,所以未清空上一筆收文的數值(ex.P-121437)
   txtPatent(17).Tag = ""  '費用 CP16
   txtPatent(21).Tag = ""  '規費 CP17
   txtPatent(18).Tag = ""  '點數 CP18
   'end 2018/12/06
   
   'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
   'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
   'modify by sonia 2014/9/11 取消X69514,已轉外專
   'Modified by Lydia 2020/01/13 X38120030寧遠縣碩寧電子有限公司,其大陸年費605不要控制金額
   'If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" And _
   '   Mid(txtPatent(8), 1, 8) <> "X3928904" And Mid(txtPatent(9), 1, 8) <> "X3928904" And Mid(txtPatent(10), 1, 8) <> "X3928904" And Mid(txtPatent(11), 1, 8) <> "X3928904" And Mid(txtPatent(12), 1, 8) <> "X3928904" Then
   'Modified by Lydia 2022/08/19 改共用模組
   'If CheckExcept = False Then
   'end 2020/01/13
   If PUB_CheckExceptFrm010005(txtSystem, txtPatent(4), txtPatent(1), ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10))) = False Then
      'Modified by Morgan 2013/2/4 +判斷非國外部收文的不預設 Ex.FCP-039919
      If txtSystem = "FCP" And Mid(m_SalesST15, 1, 1) = "F" Then
         '規費
         '2009/12/31 modify by sonia 台灣發明實審調整規費
         'txtPatent(21) = GetPatentOfficialFee(txtSystem, txtPatent(1), txtPatent(19), txtPatent(3), txtPatent(4), m_PA16)
         If txtCode(1) = "" Then txtCode(1).Text = "0"
         If txtCode(2) = "" Then txtCode(2).Text = "00"
         
'         'Add By Sindy 2012/11/22 取得特殊客戶/代理人收文費用
'         If GetSpecGuestFee(Mid(txtPatent(8), 1, 8), Mid(txtPatent(13), 1, 8), txtSystem, txtPatent(4), txtPatent(1), DBDATE(txtPatent(0)), strSG07, strSG08) = True Then
'            '規費
'            txtPatent(21) = Val(strSG08)
'            '費用
'            txtPatent(17) = Val(strSG07) + Val(strSG08)
'            '點數
'            txtPatent(18) = Format((Val(txtPatent(17)) - Val(txtPatent(21))) / 1000, "0.0")
'         Else
'         '2012/11/22 End
            '2010/8/17 MODIFY BY SONIA
            'txtPatent(21) = GetPatentOfficialFee(txtSystem, txtPatent(1), txtPatent(19), txtPatent(3), txtPatent(4), m_PA16, , txtCode(0), txtCode(1), txtCode(2))
            'Modified by Lydia 2017/03/01 +是否電子送件,+判斷
            'txtPatent(21) = GetPatentOfficialFee(txtSystem, txtPatent(1), txtPatent(19), txtPatent(3), txtPatent(4), m_PA16, m_PA14, txtCode(0), txtCode(1), txtCode(2))
             '2009/12/31 end
            strTmpVal = GetPatentOfficialFee(txtSystem, IIf(tCP10 <> "", tCP10, txtPatent(1)), txtPatent(19), txtPatent(3), txtPatent(4), m_PA16, m_PA14, txtCode(0), txtCode(1), txtCode(2), IIf(tCP10 = "" And chkWebApp.Visible = True And chkWebApp.Value = 1, "Y", ""))
            If tCP10 = "" Then
                 txtPatent(21).Text = strTmpVal
            Else
                 txtPatent(21).Tag = strTmpVal
            End If
            'end 2018/03/01
            
            '費用
            '911017 nick 邱小姐說費用還要加上規費
            'txtPatent(17) = GetFCPFee(txtPatent(1))
            'Modified by Lydia 2018/03/01 +判斷
            'If Val(GetFCPFee(txtSystem, txtPatent(1))) + Val(txtPatent(21)) > 0 Then
            '   txtPatent(17) = Val(GetFCPFee(txtSystem, txtPatent(1))) + Val(txtPatent(21))
            '   '點數
            '   txtPatent(18) = Format((Val(txtPatent(17)) - Val(txtPatent(21))) / 1000, "0.0")
            'End If
            If Val(GetFCPFee(txtSystem, IIf(tCP10 <> "", tCP10, txtPatent(1)))) + Val(strTmpVal) > 0 Then
               If tCP10 = "" Then
                    txtPatent(17).Text = Val(GetFCPFee(txtSystem, txtPatent(1))) + Val(strTmpVal)
                    '點數
                    txtPatent(18).Text = Format((Val(txtPatent(17).Text) - Val(strTmpVal)) / 1000, "0.0")
               Else
                    txtPatent(17).Tag = Val(GetFCPFee(txtSystem, tCP10)) + Val(strTmpVal)
                    '點數
                    txtPatent(18).Tag = Format((Val(txtPatent(17).Tag) - Val(strTmpVal)) / 1000, "0.0")
               End If
            End If
            'end 2018/03/01
'         End If
      '2009/10/15 ADD BY SONIA FMP案件也要預設費用,抓CASEFEE則以'FCP'+申請國家+案件性質抓
      ElseIf txtSystem = "P" And Mid(m_SalesST15, 1, 1) = "F" Then
'         'Add By Sindy 2012/11/22 取得特殊客戶/代理人收文費用
'         If GetSpecGuestFee(Mid(txtPatent(8), 1, 8), Mid(txtPatent(13), 1, 8), "FCP", txtPatent(4), txtPatent(1), DBDATE(txtPatent(0)), strSG07, strSG08) = True Then
'            '規費
'            If txtPatent(21) = "" Or Val(txtPatent(21)) = 0 Then
'               txtPatent(21) = Val(strSG08)
'            End If
'            '費用
'            If txtPatent(17) = "" Or Val(txtPatent(17)) = 0 Then
'               txtPatent(17) = Val(strSG07) + Val(strSG08)
'               '點數
'               txtPatent(18) = Format((Val(txtPatent(17)) - Val(txtPatent(21))) / 1000, "0.0")
'            End If
'         Else
'         '2012/11/22 End
            '規費
            'Remove by Lydia 2017/12/08 預設-重新計算
            'If txtPatent(21) = "" Or Val(txtPatent(21)) = 0 Then
            If tCP10 = "" Then  'Added by Lydia 2018/05/09
                txtPatent(21).Text = GetFMPOfficialFee("FCP", txtPatent(1), txtPatent(4))
            'Added by Lydia 2018/05/09 傳案件性質
            Else
                txtPatent(21).Tag = GetFMPOfficialFee("FCP", tCP10, txtPatent(4))
            End If
            'end 2018/05/09
            
            'End If
            '費用
            'Remove by Lydia 2017/12/08 預設-重新計算
            'If txtPatent(17) = "" Or Val(txtPatent(17)) = 0 Then
            If tCP10 = "" Then  'Added by Lydia 2018/05/09
               If Val(GetFMPFee("FCP", txtPatent(1), txtPatent(4))) > 0 Then
                  txtPatent(17).Text = Val(GetFMPFee("FCP", txtPatent(1), txtPatent(4)))
                  '點數
                  txtPatent(18).Text = Format((Val(txtPatent(17).Text) - Val(txtPatent(21).Text)) / 1000, "0.0")
               End If
            'Added by Lydia 2018/05/09 傳案件性質
            Else
               If Val(GetFMPFee("FCP", tCP10, txtPatent(4))) > 0 Then
                  txtPatent(17).Tag = Val(GetFMPFee("FCP", tCP10, txtPatent(4)))
                  '點數
                  txtPatent(18).Tag = Format((Val(txtPatent(17).Tag) - Val(txtPatent(21).Tag)) / 1000, "0.0")
               End If
            End If
            'end 2018/05/09
            'End If
'         End If
      '2009/10/15 END
      End If
   End If
   
   'Added by Lydia 2020/03/27 FCP-062174審定前不收費控制: 判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
   If m_PA16 = "" And InStr("FCP062174000", txtSystem & txtCode(0) & IIf(txtCode(1) = "", "0", txtCode(1)) & IIf(txtCode(2) = "", "00", txtCode(2))) > 0 Then
        txtPatent(17).Text = ""
        txtPatent(18).Text = ""
        txtPatent(21).Text = ""
   End If
   'Added by Lydia 2022/05/03 FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
   If m_PA16 <> "1" And InStr("FCP067004000", txtSystem & txtCode(0) & IIf(txtCode(1) = "", "0", txtCode(1)) & IIf(txtCode(2) = "", "00", txtCode(2))) > 0 Then
        txtPatent(17).Text = ""
        txtPatent(18).Text = ""
        txtPatent(21).Text = ""
   End If

End Sub

''Add By Cheng 2003/08/28
''比較點數與底價
'Private Function ChkPointValue(strCF01 As String, strCF02 As String, strCF03 As String, strPointValue As String, strClerk As String) As Boolean
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'Dim dblFreePoV As Double 'Add By Sindy 2010/3/23
'
'   ChkPointValue = True
'   'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'   If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" Then
'       StrSQLa = "Select * From CaseFee Where CF01='" & strCF01 & "' And CF02='" & strCF02 & "' And CF03='" & strCF03 & "' "
'       rsA.CursorLocation = adUseClient
'       rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'       If rsA.RecordCount > 0 Then
'           If Val("" & rsA("CF14").Value) > 0 Then
'               'Add By Sindy 2010/3/23
'               dblFreePoV = 0
'               StrSQLa = "SELECT * FROM staff WHERE ST01='" & Trim(strClerk) & "' "
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, StrSQLa)
'               If intI = 1 Then
'                  If Not IsNull(RsTemp("ST20")) Then
'                     If Left(Trim(RsTemp("ST20")), 1) = "4" Then dblFreePoV = 1
'                     If Left(Trim(RsTemp("ST20")), 1) = "3" Then dblFreePoV = 3
'                  End If
'               End If
'               '2010/3/23 End
'               'Modify By Sindy 2010/3/23
'               'If Val("" & rsA("CF14").Value) > Val(strPointValue) Then
'               If Val("" & rsA("CF14").Value) > (Val(strPointValue) + dblFreePoV) Then
'               '2010/3/23 End
'                   If MsgBox("您輸入的點數 (" & Val(strPointValue) & ") 低於底價 (" & Val("" & rsA("CF14").Value) & ")，請確認此客戶接洽單主管是否核示???", vbExclamation + vbYesNo) = vbNo Then
'                       ChkPointValue = False
'                   End If
'               End If
'           End If
'       End If
'       If rsA.State <> adStateClosed Then rsA.Close
'       Set rsA = Nothing
'   End If
'End Function

'2007/7/4 add by sonia 檢查是否同意重新委任
Private Function ChkAgree928(strCustNo As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   ChkAgree928 = True
   If txtPatent(1) <> "928" Then Exit Function
   StrSQLa = "Select * From LinReasignRec Where LR01='" & ChangeCustomerL(strCustNo) & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If rsA("LR09").Value = "N" Then
         MsgBox "此客戶不同意重新委任, 請退回原智權人員 !!!", vbExclamation + vbOKOnly
         ChkAgree928 = False
      End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

End Function
'2007/7/4 end
Private Sub txtPetitionx_Change(Index As Integer)
   lblPetitionNamex(Index) = ""
End Sub

Private Sub txtPetitionx_GotFocus(Index As Integer)
   TextInverse txtPetitionx(Index)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If pub_OS = "1" Then
   '   txtPetitionx(Index).IMEMode = 2
   'End If
   CloseIme
End Sub

'Modify by Amy 2021/12/16 原:Integer
Private Sub txtPetitionx_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPetitionx_Validate(Index As Integer, Cancel As Boolean)
   Dim strTemp As String, strPetition As String
   
   If Len(txtPetitionx(Index)) > 0 Then
      strPetition = txtPetitionx(Index)
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCustomer(strPetition, strTemp) Then
      If ClsPDGetCustomer(strPetition, strTemp) Then
         txtPetitionx(Index) = strPetition
         lblPetitionNamex(Index) = strTemp
      Else
         Cancel = True
         txtPetitionx(Index).SetFocus
      End If
      'Add By Sindy 2021/2/1
      'Modified by Lydia 2023/03/06 傳入本所案號 , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
      If GetCustomerAndState(strPetition, strTemp, , , , txtSystem, strXState(Index), IIf(frm010001.intSaveMode = 0, True, False), , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = False Then
         Cancel = True
         txtPetitionx(Index).SetFocus
      End If
      '2021/2/1 END
   End If
End Sub

'Modified by Lydia 2024/12/13 改成共用模組Pub_GetTCN01
'Private Function GetTCN01() As String
'    GetTCN01 = ""
'    strExc(0) = "Select TCN01 From TrackingCaseName Where TCN05='" & txtRecieveCode & "' And TCN05<>'111111' Order by TCN01"
'    If RsTemp.State <> adStateClosed Then RsTemp.Close
'        RsTemp.CursorLocation = adUseClient
'        RsTemp.Open strExc(0), cnnConnection
'        If RsTemp.RecordCount > 0 Then
'            For ii = 0 To RsTemp.RecordCount - 1
'                txtTCN01 = txtTCN01 & RsTemp.Fields("TCN01") & ","
'                If Not RsTemp.BOF And Not RsTemp.EOF Then RsTemp.MoveNext
'            Next ii
'            GetTCN01 = Left(txtTCN01, Len(txtTCN01) - 1)
'    End If
'End Function
'end 2024/12/13

Private Sub txtTCN01_GotFocus()
   TextInverse txtTCN01
End Sub

'智權人ST15為F開頭且FCP/P 新申請案(101/102/103)及衍生設計(125) 案件命名追蹤流水號不可為空
'且管制人或新增人需為智權人(2013/7/23開始使用)
Private Sub txtTCN01_Validate(Cancel As Boolean)

    '2015/4/14 MODIFY BY SONIA 林總指示開放投資法務人員可收文L案及P案,故改控制外專收文
    'If frm010001.intModifyKind = 0 And strSrvDate(2) >= Val("1020723") And Left(GetST15(txtPatent(15).Text), 1) = "F" And (txtSystem = "FCP" Or txtSystem = "P") And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "125") Then
    'Modified by Lydia 2018/09/06 改成變數
    'If frm010001.intModifyKind = 0 And strSrvDate(2) >= Val("1020723") And Left(GetST15(txtPatent(15).Text), 2) = "F2" And (txtSystem = "FCP" Or txtSystem = "P") And (txtPatent(1) = "101" Or txtPatent(1) = "102" Or txtPatent(1) = "103" Or txtPatent(1) = "125") Then
    If frm010001.intModifyKind = 0 And Left(m_SalesST15, 2) = "F2" And (txtSystem = "FCP" Or txtSystem = "P") And InStr(AddTrackingNo, txtPatent(1)) > 0 Then
        If Len(Trim(txtTCN01)) > 0 Then
            'Modified by Lydia 2024/12/13 改成共用模組，舊程式刪除
            If Pub_ChkTCN01Status(Trim(txtTCN01), Trim(txtPatent(15))) = False Then
               Cancel = True
               txtTCN01.SetFocus
               txtTCN01_GotFocus
               Exit Sub
            End If
        End If
    End If

End Sub
'end 2013/06/26

'Added by Lydia 2017/11/14 FCP案件命名電子化：中說輸入相關設定
Private Sub SetTCTdata()
Dim rsA As New ADODB.Recordset
Dim Str01 As String
Dim intA As Integer
     
    If frm010001.intModifyKind = 0 Then Exit Sub
    
    Str01 = "select TCT01 AS CP09, TCT01,TCT02,TCT03,TCT04,NVL(TCT10,TCT07) TYPE " & _
            "FROM TransCaseTitle WHERE TCT01='" & txtRecieveCode & "' "
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, Str01)
    If intA = 1 Then
      m_TCT01 = "" & rsA.Fields("TCT01")
      If m_TCT01 <> "" Then
        bolExistTCT = True
        
        With rsA
           '急件
           If Val("" & .Fields("TCT02")) > 0 Then
              ChkExpDate.Value = 1
           End If
           '譯畢期限
           If "" & .Fields("TCT02") <> "" Then
              txtData(0).Text = TransDate(.Fields("TCT02"), 1)
           End If
           txtData(0).Tag = txtData(0).Text
           If "" & .Fields("TCT03") <> "" Then
              'Modified by Lydia 2018/03/06 +format
              txtData(1).Text = Format(.Fields("TCT03"), "0000")
           End If
           txtData(1).Tag = txtData(1).Text
           '分案->因為直接寫PA150,所以退程序B=空白
           txtData(2) = IIf(txtData(2).Text <> "", txtData(2).Text, "B")
           txtData(2).Tag = txtData(2).Text
           m_TCT04 = "" & .Fields("TCT04")
           
           '工程師主管已分案後,櫃台不可變更
           'Mark by Lydia 2022/10/03 因為在修改模式只剩下急件翻譯可輸入,所以先開放;
           'ex. FCP-68033當天收文又限11點前完成命名, 但是櫃台漏輸入急件翻譯, 要求主管補輸有點困難; 考慮後決定先開放
           'If "" & .Fields("TYPE") <> "" Then
           '   fraTCT.Enabled = False
           'End If
           'end 2022/10/03
        End With
      End If
    End If
    Set rsA = Nothing
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 2 Then
      KeyAscii = UpperCase(KeyAscii)
   Else
      KeyAscii = Pub_NumAscii(KeyAscii)
   End If
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse txtData(Index)
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)

   If Trim(txtData(Index).Text) = "" And Index <> 1 Then Exit Sub
   Select Case Index
      Case 0 '譯畢期限日期
          If CheckIsTaiwanDate(txtData(Index)) = False Then
             GoTo ExceptRun
          End If
          strExc(0) = TransDate(txtData(Index), 2)
          If CompWorkDay(1, strExc(0)) <> strExc(0) Then
             MsgBox "請輸入上班日!"
             GoTo ExceptRun
          End If
          If ChkExpDate.Value = 0 Then ChkExpDate.Value = 1
      Case 1 '譯畢期限時間
          If Trim(txtData(0).Text & txtData(1).Text) = "" Then
             ChkExpDate.Value = 0
             Exit Sub
          End If
          If Len(txtData(Index)) < 4 Or Val(txtData(Index)) < 900 Or Val(txtData(Index)) > 1700 Or Val(txtData(Index)) < 0 Or Val(Right(txtData(Index), 2)) > 60 Then
             MsgBox "請輸入正確時間(ex.0900~1700) ! ", vbExclamation
             GoTo ExceptRun
          End If
          If ChkExpDate.Value = 0 Then ChkExpDate.Value = 1
      Case 2 '分組
          If InStr("1,2,3,4,B", txtData(Index)) = 0 Or txtData(Index) = "," Then
             MsgBox "請輸入1~4或B(退程序) ! ", vbExclamation
             GoTo ExceptRun
          End If
      Case 3 '中說類型
          'Modified by Lydia 2018/05/07 +6 檢視PCT公開本與FCP相異處
          'Modified by Lydia 2022/10/07 +FCP案限制輸入1~5
          If txtSystem = "FCP" And InStr("1,2,3,4,5", txtData(Index)) = 0 Or txtData(Index) = "," Then
             MsgBox "請輸入1~5 ! ", vbExclamation
             GoTo ExceptRun
          End If
          'Added by Lydia 2022/10/07 P案無外文提申本242
          If txtSystem = "P" And InStr("1,2,3,5,6", txtData(Index)) = 0 Or txtData(Index) = "," Then
             MsgBox "請輸入1,2,3,5,6 ! ", vbExclamation
             GoTo ExceptRun
          End If
          'end 2022/10/07
          If txtData(Index) = "4" And txtPatent(3) <> "3" Then
             MsgBox "專利種類必須為設計案 ! ", vbExclamation
             GoTo ExceptRun
          End If
          'Added by Lydia 2018/05/09 檢視PCT公開本與FCP相異處,限FMP案
          If txtData(Index) = "6" And txtSystem <> "P" Then
             MsgBox "檢視PCT公開本與FCP相異處,限FMP案 ! ", vbExclamation
             GoTo ExceptRun
          End If
   End Select
   
   Exit Sub
   
ExceptRun:
   txtData(Index).SetFocus
   Txtdata_GotFocus Index
   Cancel = True
End Sub

'Added by Lydia 2017/11/14 FCP案件命名電子化：中說輸入相關設定-存檔
'Mark by Lydia 2025/06/19
'Private Sub UpdateTCTrecord(ByVal mCP01 As String, ByVal mCP02 As String, ByVal mCP03 As String, ByVal mCP04 As String, ByVal mCP09 As String)
'Dim strPK As String, strANo As String
'Dim strPKList As String
'Dim m_Cp33 As Double, m_Cp34 As Double 'Added by Lydia 2018/05/07標準價和底價
'Dim m_Cp26 As String 'Added by Lydia 2018/05/10 是否算案件
'Dim strTF As String 'Added by Lydia 2018/06/28 中說第一個收文號
''Added by Lydia 2019/12/23
'Dim strReceiver As String '急件翻譯的收件者
'Dim strContent As String '急件翻譯的email內容
'
'      'If txtData(3).Visible = True Then 'Remove by Lydia 2018/09/06 增加分割案可勾選其他性質收文
'         '中說類型
'         Select Case Trim(txtData(3))
'             Case "1" '翻譯中說201
'                 strPK = AutoNo("A", 6)
'                 strTF = strPK & "-201" 'Added by Lydia 2018/06/28
'                 strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
'                          "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','201',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "201")) & " from caseprogress where cp09='" & mCP09 & "' "
'                 cnnConnection.Execute strSql, intI
'                 'Modified by Lydia 2018/04/18
'                 'strPKList = strPKList & strPK & ","
'                 strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
'             Case "2" '檢視中說209
'                 strPK = AutoNo("A", 6)
'                 strTF = strPK & "-209" 'Added by Lydia 2018/06/28
'                 strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
'                          "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','209',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "209")) & " from caseprogress where cp09='" & mCP09 & "' "
'                 cnnConnection.Execute strSql, intI
'                 'Modified by Lydia 2018/04/18
'                 'strPKList = strPKList & strPK & ","
'                 strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
'             Case "3", "4" '製作中說/撰稿210、製作中說210＆外文提申本242
'                 strPK = AutoNo("A", 6)
'                 strTF = strPK & "-210" 'Added by Lydia 2018/06/28
'                 strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
'                          "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','210',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "210")) & " from caseprogress where cp09='" & mCP09 & "' "
'                 cnnConnection.Execute strSql, intI
'                 'Modified by Lydia 2018/04/18
'                 'strPKList = strPKList & strPK & ","
'                 strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
'                 If Trim(txtData(3)) = "4" Then
'                    strPK = AutoNo("A", 6)
'                    strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
'                             "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','242',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "242")) & " from caseprogress where cp09='" & mCP09 & "' "
'                    cnnConnection.Execute strSql, intI
'                    'Modified by Lydia 2018/04/18
'                    'strPKList = strPKList & strPK & ","
'                    strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
'                 End If
'             Case "5" '核對中說235
'                 strPK = AutoNo("A", 6)
'                 strTF = strPK & "-235" 'Added by Lydia 2018/06/28
'                 strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
'                          "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','235',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "235")) & " from caseprogress where cp09='" & mCP09 & "' "
'                 cnnConnection.Execute strSql, intI
'                 'Modified by Lydia 2018/04/18
'                 'strPKList = strPKList & strPK & ","
'                 strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
'             'Added by Lydia 2018/05/07
'             Case "6" '檢視PCT公開本與FCP相異處942
'                 strPK = AutoNo("A", 6)
'                 strTF = strPK & "-942" 'Added by Lydia 2018/06/28
'                 strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
'                          "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','942',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "942")) & " from caseprogress where cp09='" & mCP09 & "' "
'                 cnnConnection.Execute strSql, intI
'                 strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
'         End Select
''--------------------------------勾選項1
'         'Added by Lydia 2018/03/01 須同時提實審
'         If ChkAdd416.Value = True Then
'            strANo = AutoNo("A", 6)
'            'Modified by Morgan 2018/6/15 P案不自動分案 Ex:P-120468
'            'strExc(1) = PUB_GetFCPHandler(mCP01, mCP02, mCP03, mCP04) '承辦人CP14
'            If mCP01 = "FCP" Then
'               strExc(1) = PUB_GetFCPHandler(mCP01, mCP02, mCP03, mCP04) '承辦人CP14
'            Else
'               strExc(1) = ""
'            End If
'            'end 2018/6/15
'
'            'Modified by Lydia 2022/08/25改成共用模組
'            'Call OnUpdateFee("416")
'            'strExc(2) = txtPatent(17).Tag  '費用 CP16
'            'strExc(3) = txtPatent(21).Tag  '規費 CP17
'            'strExc(4) = txtPatent(18).Tag  '點數 CP18
'            'Memo --- 因為是直接新增其他收文,所以不設定電子送件
'            If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), "416", txtPatent(15), _
'                 "", ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
'            End If
'            strExc(2) = m_NowCP16
'            strExc(3) = m_NowCP17
'            strExc(4) = m_NowCP18
'            'end 2022/08/25
'            strExc(5) = strExc(2)              '未收金額 CP79
'            'Added by Lydia 2018/05/07標準價和底價
'            If ClsPDGetCaseLowPrice(mCP01, txtPatent(4), "416", m_Cp33, m_Cp34) = 1 Then
'            End If
'            'end 2018/05/08
'            Pub_SetPAIsCase mCP01, "416", m_Cp26 'Added by Lydia 2018/05/10 是否算案件數
'
'            'Modified by Lydia 2018/03/15 自動上已分案(CP122)
'            'Modified by Lydia 2018/03/27 判斷退程序不上已分案
'            'strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp20,cp16,cp17,cp18,cp79,cp122) " & _
'                        "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','416',cp11,cp12,cp13,'" & strExc(1) & "'," & CNULL(PUB_GetCP20(mCP01, "416")) & _
'                                 ", " & CNULL(strExc(2), True) & ", " & CNULL(strExc(3), True) & ", " & CNULL(strExc(4), True) & ", " & CNULL(strExc(5), True) & ",'Y' from caseprogress where cp09='" & mCP09 & "' "
'            'Modified by Lydia 2018/05/10 +cp33,cp34 ,cp26
'            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp20,cp16,cp17,cp18,cp79,cp122,cp33,cp34,cp26) " & _
'                        "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','416',cp11,cp12,cp13,'" & strExc(1) & "'," & CNULL(PUB_GetCP20(mCP01, "416")) & _
'                                 ", " & CNULL(strExc(2), True) & ", " & CNULL(strExc(3), True) & ", " & CNULL(strExc(4), True) & ", " & CNULL(strExc(5), True) & _
'                                 "," & IIf(txtData(2).Text <> "" And txtData(2).Text <> "B", "'Y'", "NULL") & ", " & m_Cp33 & ", " & m_Cp34 & "," & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
'            cnnConnection.Execute strSql, intI
'            'Modified by Lydia 2018/04/18
'            'strPKList = strPKList & strPK & ","
'            strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
'         End If
''--------------------------------勾選項2
'         'Added by Lydia 2018/04/12 收文主動修正
'         'Move by Lydia 2018/05/07 從第4改到第2
'         If ChkAdd203.Value = True Then
'            strANo = AutoNo("A", 6)
'            'Modified by Lydia 2022/08/25改成共用模組
'            'Call OnUpdateFee("203")
'            'strExc(2) = txtPatent(17).Tag  '費用 CP16
'            'strExc(3) = txtPatent(21).Tag  '規費 CP17
'            'strExc(4) = txtPatent(18).Tag  '點數 CP18
'            'Memo --- 因為是直接新增其他收文,所以不設定電子送件
'            If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), "203", txtPatent(15), _
'                 "", ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
'            End If
'            strExc(2) = m_NowCP16
'            strExc(3) = m_NowCP17
'            strExc(4) = m_NowCP18
'            'end 2022/08/25
'            strExc(5) = strExc(2)              '未收金額 CP79
'            If mCP01 = "FCP" Then  'Added by Lydia 2018/05/10 判斷FCP案才有承辦期限
'                strExc(7) = Pub_GetHandleDay("FCP", "000", "203", , TransDate(txtPatent(14), 2)) '承辦期限
'            Else 'Added by Lydia 2018/05/10 判斷FCP案才有承辦期限
'                strExc(7) = ""
'            End If
'            'end 2018/05/10
'            'Added by Lydia 2018/05/07標準價和底價
'            If ClsPDGetCaseLowPrice(mCP01, txtPatent(4), "203", m_Cp33, m_Cp34) = 1 Then
'            End If
'            'end 2018/05/08
'            Pub_SetPAIsCase mCP01, "203", m_Cp26 'Added by Lydia 2018/05/10 是否算案件數
'
'            'Modified by Lydia 2018/05/10 +cp33,cp34 ,cp26
'            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20,cp48,cp16,cp17,cp18,cp79,cp33,cp34,cp26) " & _
'                     "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','203',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "203")) & ", " & CNULL(strExc(7), True) & _
'                     ", " & CNULL(strExc(2), True) & ", " & CNULL(strExc(3), True) & ", " & CNULL(strExc(4), True) & ", " & CNULL(strExc(5), True) & _
'                     ", " & m_Cp33 & ", " & m_Cp34 & ", " & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
'            cnnConnection.Execute strSql, intI
'            'Modified by Lydia 2018/04/18
'            'strPKList = strPKList & strPK & ","
'            strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
'         End If
'         'end 2018/04/12
''--------------------------------勾選項3
'         'Added by Lydia 2018/04/13 收文回代(回覆代理人)
'         If ChkAdd902.Value = True Then
'            strANo = AutoNo("A", 6)
'            If mCP01 = "FCP" Then  'Added by Lydia 2018/05/10 判斷FCP案才有承辦期限
'                strExc(7) = Pub_GetHandleDay("FCP", "000", "902", , TransDate(txtPatent(14), 2)) '承辦期限
'            Else 'Added by Lydia 2018/05/10 判斷FCP案才有承辦期限
'                strExc(7) = ""
'            End If
'            'end 2018/05/10
'            Pub_SetPAIsCase mCP01, "902", m_Cp26 'Added by Lydia 2018/05/10 是否算案件數
'
'            'Modified by Lydia 2018/05/10 +cp26
'            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20,cp48,cp26) " & _
'                     "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','902',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "902")) & _
'                     ", " & CNULL(strExc(7), True) & ", " & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
'            cnnConnection.Execute strSql, intI
'            'Modified by Lydia 2018/04/18
'            'strPKList = strPKList & strPK & ","
'            strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
'         End If
'         'end 2018/04/13
''--------------------------------勾選項4
'         '須會稿
'          'Move by Lydia 2018/05/07 從第2改到第4
'         If ChkAdd924.Value = True Then
'            strANo = AutoNo("A", 6)
'            Pub_SetPAIsCase mCP01, "924", m_Cp26 'Added by Lydia 2018/05/10 是否算案件數
'
'            'Modified by Lydia 2018/05/10 +cp26
'            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20,cp26) " & _
'                        "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','924',cp11,cp12,cp13," & _
'                         CNULL(PUB_GetCP20(mCP01, "924")) & ", " & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
'            cnnConnection.Execute strSql, intI
'            'Modified by Lydia 2018/04/18
'            'strPKList = strPKList & strPK & ","
'            strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
'         End If
''--------------------------------
'         'Added by Lydia 2021/08/27 回復說明書校閱968
'         If ChkAdd968.Value = True Then
'            strANo = AutoNo("A", 6)
'            Pub_SetPAIsCase mCP01, "968", m_Cp26
'
'            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20,cp26) " & _
'                        "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','968',cp11,cp12,cp13," & _
'                         CNULL(PUB_GetCP20(mCP01, "968")) & ", " & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
'            cnnConnection.Execute strSql, intI
'            strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
'         End If
''--------------------------------勾選項5
'         'Added by Lydia 2018/05/07 恢復權利
'         If ChkAdd414.Value = True Then
'            strANo = AutoNo("A", 6)
'            'Modified by Lydia 2022/08/25改成共用模組
'            'Call OnUpdateFee("414")
'            'strExc(2) = txtPatent(17).Tag  '費用 CP16
'            'strExc(3) = txtPatent(21).Tag  '規費 CP17
'            'strExc(4) = txtPatent(18).Tag  '點數 CP18
'            'Memo --- 因為是直接新增其他收文,所以不設定電子送件
'            If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), "414", txtPatent(15), _
'                 "", ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
'            End If
'            strExc(2) = m_NowCP16
'            strExc(3) = m_NowCP17
'            strExc(4) = m_NowCP18
'            'end 2022/08/25
'            strExc(5) = strExc(2)              '未收金額 CP79
'            '標準價和底價
'            If ClsPDGetCaseLowPrice(mCP01, txtPatent(4), "414", m_Cp33, m_Cp34) = 1 Then
'            End If
'            Pub_SetPAIsCase mCP01, "414", m_Cp26 '是否算案件數
'
'            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20,cp16,cp17,cp18,cp79,cp33,cp34,cp26) " & _
'                     "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','414',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "414")) & _
'                     ", " & CNULL(strExc(2), True) & ", " & CNULL(strExc(3), True) & ", " & CNULL(strExc(4), True) & ", " & CNULL(strExc(5), True) & _
'                     ", " & m_Cp33 & ", " & m_Cp34 & ", " & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
'            cnnConnection.Execute strSql, intI
'            strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
'         End If
'         'end 2018/05/07
''--------------------------------勾選項6
'         'Added by Lydia 2018/05/07 超頁費
'         If ChkAdd938.Value = True Then
'            strANo = AutoNo("A", 6)
'            'Modified by Lydia 2022/08/25改成共用模組
'            'Call OnUpdateFee("938")
'            'strExc(2) = txtPatent(17).Tag  '費用 CP16
'            'strExc(3) = txtPatent(21).Tag  '規費 CP17
'            'strExc(4) = txtPatent(18).Tag  '點數 CP18
'            'Memo --- 因為是直接新增其他收文,所以不設定電子送件
'            If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), "938", txtPatent(15), _
'                 "", ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
'            End If
'            strExc(2) = m_NowCP16
'            strExc(3) = m_NowCP17
'            strExc(4) = m_NowCP18
'            'end 2022/08/25
'            strExc(5) = strExc(2)              '未收金額 CP79
'            '標準價和底價
'            If ClsPDGetCaseLowPrice(mCP01, txtPatent(4), "938", m_Cp33, m_Cp34) = 1 Then
'            End If
'            Pub_SetPAIsCase mCP01, "938", m_Cp26 '是否算案件數
'
'            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20,cp16,cp17,cp18,cp79,cp33,cp34,cp26) " & _
'                     "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','938',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "938")) & _
'                     ", " & CNULL(strExc(2), True) & ", " & CNULL(strExc(3), True) & ", " & CNULL(strExc(4), True) & ", " & CNULL(strExc(5), True) & _
'                     ", " & m_Cp33 & ", " & m_Cp34 & ", " & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
'            cnnConnection.Execute strSql, intI
'            strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
'         End If
'         'end 2018/05/07
''--------------------------------勾選項7
'         'Added by Lydia 2018/05/07 超項費
'         If ChkAdd939.Value = True Then
'            strANo = AutoNo("A", 6)
'            'Modified by Lydia 2022/08/25改成共用模組
'            'Call OnUpdateFee("939")
'            'strExc(2) = txtPatent(17).Tag  '費用 CP16
'            'strExc(3) = txtPatent(21).Tag  '規費 CP17
'            'strExc(4) = txtPatent(18).Tag  '點數 CP18
'            'Memo --- 因為是直接新增其他收文,所以不設定電子送件
'            If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), "939", txtPatent(15), _
'                 "", ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) Then
'            End If
'            strExc(2) = m_NowCP16
'            strExc(3) = m_NowCP17
'            strExc(4) = m_NowCP18
'            'end 2022/08/25
'            strExc(5) = strExc(2)              '未收金額 CP79
'            '標準價和底價
'            If ClsPDGetCaseLowPrice(mCP01, txtPatent(4), "939", m_Cp33, m_Cp34) = 1 Then
'            End If
'            Pub_SetPAIsCase mCP01, "939", m_Cp26 '是否算案件數
'
'            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20,cp16,cp17,cp18,cp79,cp33,cp34,cp26) " & _
'                     "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','939',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "939")) & _
'                     ", " & CNULL(strExc(2), True) & ", " & CNULL(strExc(3), True) & ", " & CNULL(strExc(4), True) & ", " & CNULL(strExc(5), True) & _
'                     ", " & m_Cp33 & ", " & m_Cp34 & ", " & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
'            cnnConnection.Execute strSql, intI
'            strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
'         End If
'         'end 2018/05/07
''--------------------------------勾選項8
'         'Added by Lydia 2018/05/07 主張國際優先權
'         If ChkAdd106.Value = True Then
'            strANo = AutoNo("A", 6)
'            'Modified by Lydia 2022/08/25改成共用模組
'            'Call OnUpdateFee("106")
'            'strExc(2) = txtPatent(17).Tag  '費用 CP16
'            'strExc(3) = txtPatent(21).Tag  '規費 CP17
'            'strExc(4) = txtPatent(18).Tag  '點數 CP18
'            'Memo --- 因為是直接新增其他收文,所以不設定電子送件
'            If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), "106", txtPatent(15), _
'                 "", ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
'            End If
'            strExc(2) = m_NowCP16
'            strExc(3) = m_NowCP17
'            strExc(4) = m_NowCP18
'            'end 2022/08/25
'            strExc(5) = strExc(2)              '未收金額 CP79
'            '標準價和底價
'            If ClsPDGetCaseLowPrice(mCP01, txtPatent(4), "106", m_Cp33, m_Cp34) = 1 Then
'            End If
'            Pub_SetPAIsCase mCP01, "106", m_Cp26 '是否算案件數
'
'            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20,cp16,cp17,cp18,cp79,cp33,cp34,cp26) " & _
'                     "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','106',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "106")) & _
'                     ", " & CNULL(strExc(2), True) & ", " & CNULL(strExc(3), True) & ", " & CNULL(strExc(4), True) & ", " & CNULL(strExc(5), True) & _
'                     ", " & m_Cp33 & ", " & m_Cp34 & ", " & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
'            cnnConnection.Execute strSql, intI
'            strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
'         End If
'         'end 2018/05/07
''--------------------------------勾選項9
'         'Added by Lydia 2018/05/07 呈國際階段修正內容
'         If ChkAdd228.Value = True Then
'            strANo = AutoNo("A", 6)
'            'Modified by Lydia 2022/08/25改成共用模組
'            'Call OnUpdateFee("228")
'            'strExc(2) = txtPatent(17).Tag  '費用 CP16
'            'strExc(3) = txtPatent(21).Tag  '規費 CP17
'            'strExc(4) = txtPatent(18).Tag  '點數 CP18
'            'Memo --- 因為是直接新增其他收文,所以不設定電子送件
'            If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), "228", txtPatent(15), _
'                 "", ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
'            End If
'            strExc(2) = m_NowCP16
'            strExc(3) = m_NowCP17
'            strExc(4) = m_NowCP18
'            'end 2022/08/25
'            strExc(5) = strExc(2)              '未收金額 CP79
'            '標準價和底價
'            If ClsPDGetCaseLowPrice(mCP01, txtPatent(4), "228", m_Cp33, m_Cp34) = 1 Then
'            End If
'            Pub_SetPAIsCase mCP01, "228", m_Cp26 '是否算案件數
'
'            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20,cp16,cp17,cp18,cp79,cp33,cp34,cp26) " & _
'                     "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','228',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mCP01, "228")) & _
'                     ", " & CNULL(strExc(2), True) & ", " & CNULL(strExc(3), True) & ", " & CNULL(strExc(4), True) & ", " & CNULL(strExc(5), True) & _
'                     ", " & m_Cp33 & ", " & m_Cp34 & ", " & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
'            cnnConnection.Execute strSql, intI
'            strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
'         End If
'         'end 2018/05/07
''--------------------------------勾選項10
'         'Added by Lydia 2018/09/06 (FCP分割案) 續行母案再審
'         If ChkAdd435.Value = True Then
'            strANo = AutoNo("A", 6)
'            strExc(1) = ""  '承辦人CP14
'            'Modified by Lydia 2022/08/25改成共用模組
'            'Call OnUpdateFee("435")
'            'strExc(2) = txtPatent(17).Tag  '費用 CP16
'            'strExc(3) = txtPatent(21).Tag  '規費 CP17
'            'strExc(4) = txtPatent(18).Tag  '點數 CP18
'            'Memo --- 因為是直接新增其他收文,所以不設定電子送件
'            If PUB_Frm010005OnUpdFee(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtPatent(4), txtPatent(3), m_PA16, m_PA14, txtPatent(19), "435", txtPatent(15), _
'                 "", ChangeCustomerL(txtPatent(8)) & "," & ChangeCustomerL(txtPatent(11)) & "," & ChangeCustomerL(txtPatent(9)) & "," & ChangeCustomerL(txtPatent(12)) & "," & ChangeCustomerL(txtPatent(10)), m_NowCP16, m_NowCP17, m_NowCP18) = True Then
'            End If
'            strExc(2) = m_NowCP16
'            strExc(3) = m_NowCP17
'            strExc(4) = m_NowCP18
'            'end 2022/08/25
'            strExc(5) = strExc(2)              '未收金額 CP79
'            '標準價和底價
'            If ClsPDGetCaseLowPrice(mCP01, txtPatent(4), "435", m_Cp33, m_Cp34) = 1 Then
'            End If
'
'            Pub_SetPAIsCase mCP01, "435", m_Cp26 '是否算案件數
'
'            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp20,cp16,cp17,cp18,cp79,cp122,cp33,cp34,cp26) " & _
'                        "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "','435',cp11,cp12,cp13,'" & strExc(1) & "'," & CNULL(PUB_GetCP20(mCP01, "435")) & _
'                                 ", " & CNULL(strExc(2), True) & ", " & CNULL(strExc(3), True) & ", " & CNULL(strExc(4), True) & ", " & CNULL(strExc(5), True) & _
'                                 "," & IIf(txtData(2).Text <> "" And txtData(2).Text <> "B", "'Y'", "NULL") & ", " & m_Cp33 & ", " & m_Cp34 & "," & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
'            cnnConnection.Execute strSql, intI
'            strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
'         End If
''--------------------------------end
'         If strPKList <> "" Then
'            'Added by Lydia 2018/06/28 急件翻譯將新案翻譯收文號回寫到翻譯費用檔和命名記錄檔(TCN14)
'            If txtTCN01.Text <> "" And strTF <> "" Then
'                'Modified by Lydia 2020/01/09 FMP案收文，新案建檔未提申先翻譯自動勾選
'                'strSql = "update TransFee set TF01=" & CNULL(Mid(strTF, 1, 9)) & " where TF01=" & CNULL(txtTCN01.Text)
'                strSql = "update TransFee set TF01=" & CNULL(Mid(strTF, 1, 9)) & IIf(mCP01 = "P", " ,TF31='Y' ", "") & _
'                            " where TF01=" & CNULL(txtTCN01.Text)
'                cnnConnection.Execute strSql, intI
'                If intI > 0 Then
'                     strReceiver = Pub_GetSpecMan("M") 'Added by Lydia 2019/12/23 急件翻譯的收件者
'                     strSql = "update TrackingCaseName set TCN14=" & CNULL(Mid(strTF, 1, 9)) & " where TCN01=" & CNULL(txtTCN01.Text)
'                     cnnConnection.Execute strSql, intI
'                     If Right(strTF, 3) <> "201" Then
'                         strExc(1) = Pub_GetSpecMan("M")
'                         If strExc(1) <> "" Then
'                            'Modified by Lydia 2019/12/23 +CC:急件翻譯的收件者
'                            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'                               " values( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd') ,to_char(sysdate,'hh24miss'),'" & _
'                               "急件翻譯(" & txtTCN01 & ")未收新案翻譯,請檢查" & mCP01 & "-" & mCP02 & IIf(mCP03 & mCP04 <> "000", "-" & mCP03 & "-" & mCP04, "") & "' ,'同主旨','" & strReceiver & "')"
'                            cnnConnection.Execute strSql
'                         End If
'                     End If
'                     'Added by Lydia 2018/08/13 急件翻譯立案時,有交稿期限產生國外部行事曆(比照新案建檔)
'                     'Modified by Lydia 2018/12/04 +只交Claims期限
'                     'Modified by Lydia 2019/12/23 改成同時抓命名追蹤
'                     'strExc(0) = "select TF01,TF26,TF32 from TransFee where TF01=" & CNULL(Mid(strTF, 1, 9))
'                     strExc(0) = "select TF01,TF26,TF27,TF28,TF32,b2.* from TransFee a1, TrackingCaseName b2 where TF01=" & CNULL(Mid(strTF, 1, 9)) & " and tf01=tcn14(+) "
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                         If Val("" & RsTemp.Fields("TF26")) > 0 Then
'                            strExc(1) = PUB_GetFCPHandler(mCP01, mCP02, mCP03, mCP04)
'                            strExc(0) = Pub_GetSpecMan("M")
'                            If PUB_AddFCPStaffCalendar("" & RsTemp.Fields("TF26"), "1", strExc(1) & IIf(strExc(1) <> strExc(0), "," & strExc(0), ""), "譯者翻譯交稿期限", strExc(1), "1", mCP01, mCP02, mCP03, mCP04) = True Then
'                            End If
'                         End If
'                         'Added by Lydia 2018/12/04 只交Claims期限
'                         If Val("" & RsTemp.Fields("TF32")) > 0 Then
'                            strExc(1) = PUB_GetFCPHandler(mCP01, mCP02, mCP03, mCP04)
'                            strExc(0) = Pub_GetSpecMan("M")
'                            If PUB_AddFCPStaffCalendar("" & RsTemp.Fields("TF32"), "1", strExc(1) & IIf(strExc(1) <> strExc(0), "," & strExc(0), ""), "譯者Claims交稿期限", strExc(1), "1", mCP01, mCP02, mCP03, mCP04) = True Then
'                            End If
'                         End If
'                         'Added by Lydia 2019/12/23 急件翻譯加註在email內文
'                         If Right(strTF, 3) = "201" And "" & RsTemp.Fields("TCN01") <> "" Then
'                              strExc(9) = "" & RsTemp.Fields("tcn15") & " " & GetStaffName("" & RsTemp.Fields("tcn15"))
'                              strContent = mCP01 & "-" & mCP02 & IIf(mCP03 & mCP04 <> "000", "-" & mCP03 & "-" & mCP04, "") & "已提供檔案給" & strExc(9) & "，進行翻譯" & vbCrLf
'                              strContent = strContent & "追蹤號：" & RsTemp.Fields("tcn01") & vbCrLf
'                              strContent = strContent & "翻譯人員：" & strExc(9) & vbCrLf
'                              strContent = strContent & "交稿期限：" & ChangeTStringToTDateString("" & RsTemp.Fields("tf26")) & vbCrLf
'                              If "" & RsTemp.Fields("tf32") <> "" Then strContent = strContent & "只交Claims期限：" & ChangeTStringToTDateString("" & RsTemp.Fields("tf32")) & vbCrLf
'                              strContent = strContent & "原文語種：" & Pub_GetTransFeeL("1", "" & RsTemp.Fields("tf27")) & vbCrLf
'                              strContent = strContent & "翻譯語種：" & Pub_GetTransFeeL("2", "" & RsTemp.Fields("tf28")) & vbCrLf
'                              strContent = strContent & String(30, "-") & vbCrLf
'                              strContent = strContent & "管制人：" & GetStaffName("" & RsTemp.Fields("tcn03")) & vbCrLf
'                              strContent = strContent & "備　註：" & RsTemp.Fields("tcn04") & vbCrLf
'                         End If
'                     End If
'                     'end 2018/08/13
'                'Added by Lydia 2020/01/06 FMP案收文，新案建檔未提申先翻譯自動勾選
'                ElseIf mCP01 = "P" And Trim(txtData(3)) = "1" Then
'                     strSql = "Insert into TransFee(TF01,TF31) values(" & CNULL(Mid(strTF, 1, 9)) & ", 'Y' )"
'                     cnnConnection.Execute strSql
'                     'Move by Lydia 2020/12/09 移到下方
'                'end 2020/01/06
'                End If
'                'Added by Lydia 2018/08/27 新案翻譯(201)預設個案的固定報價(PA62)
'                If Trim(txtData(3)) = "1" Then
'                     strExc(0) = Pub_GetPa62Flag(mCP01 & mCP02 & mCP03 & mCP04)
'                     If strExc(0) <> "" Then
'                          strSql = "update patent set pa62='" & strExc(0) & "' where " & ChgPatent(mCP01 & mCP02 & mCP03 & mCP04)
'                          cnnConnection.Execute strSql, intI
'                     End If
'                     'Move by Lydia 2020/12/09 從上面移過來 ex.P-126344在立卷前有急件翻譯
'                     If mCP01 = "P" Then 'Added by Lydia 2020/12/09
'                         'Added by Lydia 2020/08/24 FMP案預設發"未提申先翻譯"email ; 參考frm060102
'                         strExc(0) = Pub_GetSpecMan("M")
'                         If strExc(0) <> "" Then
'                             strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'                                " values( '" & strUserNum & "','" & strExc(0) & "',to_char(sysdate,'yyyymmdd')" & _
'                                ",to_char(sysdate,'hh24miss'),'" & mCP01 & mCP02 & IIf(mCP03 & mCP04 <> "000", mCP03 & mCP04, "") & " 未提申先翻譯" & "','同主旨',null)"
'                             cnnConnection.Execute strSql
'                         End If
'                         'end 2020/08/24
'                     End If 'Added by Lydia 2020/12/09
'                     'end------Move by Lydia 2020/12/09
'                End If
'                'end 2018/08/27
'            End If
'            'end 2018/06/28
'            strPKList = Replace(Mid(strPKList, 1, Len(strPKList) - 1), ",", "、")
'            'Modified by Lydia 2018/05/07 改說明
'            'frm010001.lblTCT.Caption = "中說、實審或會稿收文號：" & strPKList
'            frm010001.lblTCT.Caption = "中說或其他收文號："
'            frm010001.lblTCTNO.Caption = strPKList
'         End If
'      'Remove by Lydia 2018/09/06 增加分割案可勾選其他性質收文
'      'Else
'      '   strPK = mCP09
'      'End If
'      'end 2018/09/06
'
'      'Added by Lydia 2018/09/06 判斷有走命名流程的新案性質才產生記錄
'      If InStr(FcpAddTct, txtPatent(1)) > 0 Then
'          '抓分案組別-主管
'          strExc(6) = ""
'          'Modified by Lydia 2019/01/09
''          Select Case Trim(txtData(2))
''              Case "1" '電子組
''                   strExc(6) = Pub_GetSpecMan("T")
''              Case "2" '化學組
''                   strExc(6) = Pub_GetSpecMan("R")
''              Case "3" '日文組
''                   strExc(6) = Pub_GetSpecMan("S")
''              Case "4" '機械組
''                   strExc(6) = Pub_GetSpecMan("T1")
''              Case "B": 'FCP程序管制人
''                   strExc(6) = "B"
''          End Select
'          strExc(6) = Pub_GetFCPGrpMan(Trim(txtData(2)))
'          strExc(6) = PUB_GetStateForMan(strExc(6)) 'Added by Lydia 2022/10/12 特殊情況之指定職代
'          If strExc(6) = "" Then strExc(6) = "B"
'          'end 2019/01/09
'
'          '譯畢期限
'          strExc(4) = IIf(txtData(0).Text <> "", TransDate(txtData(0).Text, 2), "")
'          strExc(5) = IIf(txtData(1).Text <> "", txtData(1).Text, "")
'
'          strSql = ""
'          If bolExistTCT = False Then
'             'Modified by Lydia 2018/06/26 去除單引號
'             'strSql = "INSERT INTO TransCaseTitle(TCT01,TCT02,TCT03,TCT04,TCT16,TCT17,TCT112,TCT113,TCT114) " & _
'                      "VALUES ('" & mCP09 & "'," & CNULL(strExc(4), True) & "," & CNULL(strExc(5), True) & "," & CNULL(IIf(strExc(6) <> "B", strExc(6), "")) & _
'                      ",'" & txtPatent(5).Text & "','" & txtPatent(6).Text & "','" & strUserNum & "'," & strSrvDate(1) & "," & Mid(Format(ServerTime, "000000"), 1, 4) & ")"
'             strSql = "INSERT INTO TransCaseTitle(TCT01,TCT02,TCT03,TCT04,TCT16,TCT17,TCT112,TCT113,TCT114) " & _
'                      "VALUES ('" & mCP09 & "'," & CNULL(strExc(4), True) & "," & CNULL(strExc(5), True) & "," & CNULL(IIf(strExc(6) <> "B", strExc(6), "")) & _
'                      ",'" & ChgSQL(txtPatent(5).Text) & "','" & ChgSQL(txtPatent(6).Text) & "','" & strUserNum & "'," & strSrvDate(1) & "," & Mid(Format(ServerTime, "000000"), 1, 4) & ")"
'          Else
'             '譯畢期限
'             If txtData(0).Text <> txtData(0).Tag Then
'                strSql = strSql & ", TCT02=" & CNULL(strExc(4), True)
'             End If
'             If txtData(1).Text <> txtData(1).Tag Then
'                strSql = strSql & ", TCT03=" & CNULL(strExc(5), True)
'             End If
'             '分案
'             If txtData(2).Text <> txtData(2).Tag Or m_TCT04 <> strExc(6) Then
'                strSql = strSql & ", TCT04=" & CNULL(IIf(strExc(6) <> "B", strExc(6), ""))
'             End If
'             '案件名稱
'             If txtPatent(5).Text <> txtPatent(5).Tag Then
'                  'Modified by Lydia 2018/06/26 +去除單引號ChgSQL
'                  strSql = strSql & ", TCT16=" & CNULL(ChgSQL(txtPatent(5).Text))
'             End If
'             If txtPatent(6).Text <> txtPatent(6).Tag Then
'                  'Modified by Lydia 2018/06/26 +去除單引號ChgSQL
'                  strSql = strSql & ", TCT17=" & CNULL(ChgSQL(txtPatent(6).Text))
'             End If
'             If strSql <> "" Then
'                strSql = "UPDATE TransCaseTitle SET TCT112='" & strUserNum & "', TCT113=" & strSrvDate(1) & ", TCT114=" & Mid(Format(ServerTime, "000000"), 1, 4) & strSql & " WHERE TCT01='" & mCP09 & "' "
'             End If
'          End If
'          If strSql <> "" Then
'            cnnConnection.Execute strSql, intI
'            '發email通知各組主管
'             '新增
'            If bolExistTCT = False Then
'                  'Modified by Lydia 2019/12/23 + strContent,副本收受者
'                  'If PUB_GetTCTmail(True, 1, mCP01, mCP02, mCP03, mCP04, mCP09, strExc(6)) Then
'                  'Modified by Lydia 2022/08/08 將David加入(英文組)所有FCP, P櫃台新案立卷 (101-103)通知收件人之一
'                  'If PUB_GetTCTmail(True, 1, mCP01, mCP02, mCP03, mCP04, mCP09, strExc(6), , , , strContent, strReceiver) Then
'                  strExc(7) = ""
'                  If InStr("101,102,103,", txtPatent(1) & ",") > 0 And strSrvDate(1) >= 外專信件沖銷啟用日 Then
'                      If Trim(txtPatent(13)) <> "" Then  'FC代理人
'                          strExc(7) = GetPrjNationNumber(ChangeCustomerL(Trim(txtPatent(13))))
'                      ElseIf Trim(txtPatent(8)) <> "" Then '申請人1
'                          strExc(7) = GetPrjNationNumber1(ChangeCustomerL(Trim(txtPatent(8))))
'                      End If
'                      If Left(strExc(7), 3) <> "011" Then
'                          strExc(7) = "+77015"
'                      'Added by Lydia 2022/08/19 debug: 日本案要清空
'                      Else
'                          strExc(7) = ""
'                      'end 2022/08/19
'                      End If
'                  End If
'                  'Modified by Lydia 2023/02/17區分「櫃台收新案=新案立卷」 iSta=1=>0
'                  'If PUB_GetTCTmail(True, 1, mCP01, mCP02, mCP03, mCP04, mCP09, strExc(6), strExc(7), , , strContent, strReceiver) Then
'                  If PUB_GetTCTmail(True, 0, mCP01, mCP02, mCP03, mCP04, mCP09, strExc(6), strExc(7), , , strContent, strReceiver) Then
'                  'end 2022/08/08
'                  End If
'            '修改
'            ElseIf txtData(2).Text <> txtData(2).Tag Or txtData(0).Tag <> txtData(0).Text Or txtData(1).Tag <> txtData(1).Text Then
'                  If txtData(2).Text <> txtData(2).Tag Then
'                      '改組別
'                      strExc(6) = m_TCT04 & ";" & strExc(6)
'                      If PUB_GetTCTmail(True, 2, mCP01, mCP02, mCP03, mCP04, mCP09, , strExc(6), txtData(2).Tag & "-" & txtData(2).Text, "修改補發: ") = True Then
'                      End If
'                  Else
'                      If PUB_GetTCTmail(True, 1, mCP01, mCP02, mCP03, mCP04, mCP09, strExc(6), , , "修改補發: ") Then
'                      End If
'                  End If
'            End If
'          End If
'      End If
'End Sub
'end 2025/06/19

'Added by Lydia 2018/03/01 改用FTP從Tracking_no 搬到English_Vers案號資料夾
'Mark by Lydia 2020/04/29 改寫法，暫時保留
'Memo by Lydia 2022/09/05 刪除舊寫法MoveFtpFile


'Added by Lydia 2018/11/22 區分專利設計案案件性質
Private Sub SetCombo3(ByVal pKind As String)
    If pKind <> Combo3.Tag Then
        If pKind = "3" Then
             Combo3.Clear
             For ii = 1 To 4
                  Combo3.AddItem ii & "." & PUB_GetCaseAttributeName(Trim(ii), "3")
             Next
        'Added by Lydia 2018/12/19 發明和新型案的屬性相同,不用再變更
        ElseIf (pKind = "1" Or pKind = "2") And (Combo3.Tag = "1" Or Combo3.Tag = "2") Then
        'end 2018/12/19
        Else
             Combo3.Clear
             For ii = 1 To 3
                  Combo3.AddItem ii & "." & PUB_GetCaseAttributeName(Trim(ii), pKind)
             Next
        End If
    End If
    Combo3.Tag = pKind
End Sub

'Added by Lydia 2020/01/13 判斷例外客戶
'Modified by Lydia 2022/08/19 改成PUB_CheckExceptFrm010005
'Public Function CheckExcept() As Boolean
'
'     CheckExcept = False
'    'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
'    'modify by sonia 2014/9/11 取消X69514,已轉外專
'    'Modified by Lydia 2020/01/13 X3812003寧遠縣碩寧電子有限公司,其大陸年費605不要控制金額
'    If Mid(txtPatent(8), 1, 8) <> "X1484305" And Mid(txtPatent(9), 1, 8) <> "X1484305" And Mid(txtPatent(10), 1, 8) <> "X1484305" And Mid(txtPatent(11), 1, 8) <> "X1484305" And Mid(txtPatent(12), 1, 8) <> "X1484305" And _
'       Mid(txtPatent(8), 1, 8) <> "X3928904" And Mid(txtPatent(9), 1, 8) <> "X3928904" And Mid(txtPatent(10), 1, 8) <> "X3928904" And Mid(txtPatent(11), 1, 8) <> "X3928904" And Mid(txtPatent(12), 1, 8) <> "X3928904" Then
'       'Added by Lydia 2020/01/13 X3812003寧遠縣碩寧電子有限公司的P大陸案年費605不要控制金額
'       If txtSystem = "P" And txtPatent(4) = "020" And Trim(txtPatent(1)) = "605" And _
'            InStr(Mid(txtPatent(8), 1, 8) & "," & Mid(txtPatent(9), 1, 8) & "," & Mid(txtPatent(10), 1, 8) & "," & Mid(txtPatent(11), 1, 8) & "," & Mid(txtPatent(12), 1, 8), "X3812003") > 0 Then
'                CheckExcept = True
'       End If
'    Else
'                CheckExcept = True
'    End If
'End Function
'end 2022/08/19

'Added by Lydia 2020/05/20 法律所案源收文：讀取法務案源檔
Private Sub ReadLOS()
Dim strR As String, intR As Integer
Dim rsRd As New ADODB.Recordset
     
    m_LOS15 = ""
    If frm010001.intModifyKind = 0 And strSrvDate(1) >= 法律所案源收文啟用日 Then
        If frm010001.txtLOS15 <> "" Then
            strR = "select X.*,cp01,cp02,cp03,cp04 from LawOfficeSource X,caseprogress where los15=" & CNULL(frm010001.txtLOS15) & " and los01=cp09(+) "
            intR = 1
            Set rsRd = ClsLawReadRstMsg(intR, strR)
            If intR = 1 Then
                '(原)案源案件類型
                m_LOS02 = "" & rsRd.Fields("LOS02")
                t_LOSkind = m_LOS02
                '案源單號
                m_LOS15 = "" & rsRd.Fields("LOS15")
            End If
        'Mark by Lydia 2020/06/10 先保留
        'Else
        '    '案件性質=>案源案件類型 ; 前畫面未輸入案源編號也要另外判斷
        '    t_LOSkind = PUB_GetLOSkind(txtSystem, txtPatent(1), txtPatent(4))
        'end 2020/06/10
        End If
    End If
    Set rsRd = Nothing
End Sub

