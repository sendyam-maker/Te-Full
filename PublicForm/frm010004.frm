VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010004 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   7050
   ClientLeft      =   -1120
   ClientTop       =   -100
   ClientWidth     =   8720
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   8720
   Begin VB.CommandButton cmdTSMap 
      BackColor       =   &H0000FF00&
      Caption         =   "查名代號"
      Height          =   400
      Left            =   4440
      Style           =   1  '圖片外觀
      TabIndex        =   126
      Top             =   15
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCP64 
      Height          =   315
      Left            =   6600
      TabIndex        =   94
      Top             =   6780
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5820
      TabIndex        =   47
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7785
      TabIndex        =   49
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6660
      TabIndex        =   48
      Top             =   15
      Width           =   1100
   End
   Begin VB.Frame fraWindow1 
      BorderStyle     =   0  '沒有框線
      Height          =   6435
      Left            =   30
      TabIndex        =   50
      Top             =   480
      Width           =   8625
      Begin VB.Frame fraPatition 
         BorderStyle     =   0  '沒有框線
         Height          =   915
         Left            =   10
         TabIndex        =   90
         Top             =   5565
         Visible         =   0   'False
         Width           =   8505
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   31
            Left            =   1170
            TabIndex        =   46
            Top             =   570
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "8555;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   30
            Left            =   5370
            TabIndex        =   45
            Top             =   300
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "8555;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   29
            Left            =   1170
            TabIndex        =   44
            Top             =   300
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "8555;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   28
            Left            =   5370
            TabIndex        =   43
            Top             =   30
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "8555;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   20
            Left            =   1170
            TabIndex        =   42
            Top             =   30
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "8555;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "移轉申請人5："
            Height          =   180
            Left            =   30
            TabIndex        =   116
            Top             =   630
            Width           =   1170
         End
         Begin MSForms.Label lblPetitionName 
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   115
            Top             =   600
            Width           =   1875
            VariousPropertyBits=   27
            Size            =   "3307;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "移轉申請人4："
            Height          =   180
            Left            =   4230
            TabIndex        =   114
            Top             =   360
            Width           =   1170
         End
         Begin MSForms.Label lblPetitionName 
            Height          =   255
            Index           =   3
            Left            =   6600
            TabIndex        =   113
            Top             =   330
            Width           =   1875
            VariousPropertyBits=   27
            Size            =   "3307;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "移轉申請人3："
            Height          =   180
            Left            =   30
            TabIndex        =   112
            Top             =   360
            Width           =   1170
         End
         Begin MSForms.Label lblPetitionName 
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   111
            Top             =   330
            Width           =   1815
            VariousPropertyBits=   27
            Size            =   "3201;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "移轉申請人2："
            Height          =   180
            Left            =   4230
            TabIndex        =   110
            Top             =   90
            Width           =   1170
         End
         Begin MSForms.Label lblPetitionName 
            Height          =   255
            Index           =   1
            Left            =   6600
            TabIndex        =   109
            Top             =   60
            Width           =   1875
            VariousPropertyBits=   27
            Size            =   "3307;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "移轉申請人1："
            Height          =   180
            Left            =   30
            TabIndex        =   108
            Top             =   90
            Width           =   1170
         End
         Begin MSForms.Label lblPetitionName 
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   107
            Top             =   60
            Width           =   1815
            VariousPropertyBits=   27
            Size            =   "3201;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.TextBox txtCFTNa048 
         Height          =   300
         Left            =   3270
         MaxLength       =   12
         TabIndex        =   30
         Top             =   4830
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CheckBox chkWebApp 
         Caption         =   "電子送件"
         Height          =   255
         Left            =   4530
         TabIndex        =   39
         Top             =   4980
         Width           =   1050
      End
      Begin VB.CheckBox Check2 
         Caption         =   "有★★的應收帳款簽核控管"
         Height          =   285
         Left            =   5820
         TabIndex        =   25
         Top             =   4350
         Width           =   2505
      End
      Begin VB.Frame Frame21 
         BorderStyle     =   0  '沒有框線
         Height          =   525
         Left            =   10
         TabIndex        =   121
         Top             =   5220
         Visible         =   0   'False
         Width           =   5925
         Begin VB.TextBox textCP143 
            Height          =   270
            Left            =   3420
            MaxLength       =   1
            TabIndex        =   34
            Top             =   180
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox textCP122 
            Height          =   270
            Left            =   5130
            MaxLength       =   1
            TabIndex        =   35
            Top             =   60
            Width           =   255
         End
         Begin VB.TextBox textEP34 
            Height          =   270
            Left            =   3270
            MaxLength       =   1
            TabIndex        =   33
            Top             =   60
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox textEP06 
            Height          =   270
            Left            =   1350
            MaxLength       =   1
            TabIndex        =   32
            Top             =   60
            Width           =   255
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "查名是否齊備：       (Y/N)"
            Height          =   180
            Left            =   2130
            TabIndex        =   128
            Top             =   240
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "是否急件：       (Y/N)"
            Height          =   180
            Left            =   4200
            TabIndex        =   124
            Top             =   105
            Width           =   1620
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "是否會稿：       (Y/N)"
            Height          =   180
            Left            =   2340
            TabIndex        =   123
            Top             =   105
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "資料是否齊備：       (Y/N)"
            Height          =   180
            Left            =   60
            TabIndex        =   122
            Top             =   105
            Width           =   1980
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "現金或支票"
         Height          =   285
         Left            =   5700
         TabIndex        =   40
         Top             =   4980
         Width           =   1215
      End
      Begin VB.Frame fraPromoter 
         BorderStyle     =   0  '沒有框線
         Height          =   315
         Left            =   4560
         TabIndex        =   91
         Top             =   4590
         Width           =   3675
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   21
            Left            =   780
            TabIndex        =   28
            Top             =   90
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   6
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label15 
            Caption         =   "承辦人："
            Height          =   255
            Left            =   0
            TabIndex        =   92
            Top             =   90
            Width           =   855
         End
         Begin MSForms.Label lblPromoter 
            Height          =   285
            Left            =   2040
            TabIndex        =   93
            Top             =   90
            Width           =   1575
            VariousPropertyBits=   27
            Size            =   "2778;503"
            SpecialEffect   =   2
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
         Height          =   3155
         Left            =   120
         TabIndex        =   51
         Top             =   570
         Width           =   8352
         Begin VB.TextBox txtTS 
            Height          =   280
            Index           =   1
            Left            =   5760
            MaxLength       =   6
            TabIndex        =   38
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtTS 
            Height          =   280
            Index           =   0
            Left            =   5310
            MaxLength       =   3
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   720
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Frame fraTM15 
            BorderStyle     =   0  '沒有框線
            Height          =   330
            Left            =   105
            TabIndex        =   97
            Top             =   735
            Visible         =   0   'False
            Width           =   2595
            Begin MSForms.TextBox txtTrademark 
               Height          =   300
               Index           =   7
               Left            =   1110
               TabIndex        =   6
               Top             =   0
               Width           =   1332
               VariousPropertyBits=   679493659
               MaxLength       =   20
               Size            =   "2350;529"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "商標審定號："
               Height          =   180
               Left            =   15
               TabIndex        =   98
               Top             =   15
               Width           =   1080
            End
         End
         Begin VB.TextBox txtSystem 
            Enabled         =   0   'False
            Height          =   288
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   52
            Top             =   180
            Width           =   732
         End
         Begin VB.Frame fraElse 
            BorderStyle     =   0  '沒有框線
            Caption         =   "Frame2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1800
            TabIndex        =   53
            Top             =   180
            Width           =   2388
            Begin VB.TextBox txtCode 
               Enabled         =   0   'False
               Height          =   288
               Index           =   2
               Left            =   1560
               MaxLength       =   2
               TabIndex        =   56
               Top             =   0
               Width           =   492
            End
            Begin VB.TextBox txtCode 
               Enabled         =   0   'False
               Height          =   288
               Index           =   1
               Left            =   1230
               MaxLength       =   1
               TabIndex        =   55
               Top             =   0
               Width           =   372
            End
            Begin VB.TextBox txtCode 
               Enabled         =   0   'False
               Height          =   288
               Index           =   0
               Left            =   0
               MaxLength       =   6
               TabIndex        =   54
               Top             =   0
               Width           =   1212
            End
         End
         Begin VB.Frame fraTF 
            BorderStyle     =   0  '沒有框線
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   1800
            TabIndex        =   57
            Top             =   180
            Width           =   2352
            Begin VB.TextBox txtTFCode 
               Enabled         =   0   'False
               Height          =   288
               Index           =   3
               Left            =   1680
               TabIndex        =   61
               Top             =   0
               Width           =   492
            End
            Begin VB.TextBox txtTFCode 
               Enabled         =   0   'False
               Height          =   288
               Index           =   2
               Left            =   1320
               TabIndex        =   60
               Top             =   0
               Width           =   372
            End
            Begin VB.TextBox txtTFCode 
               Enabled         =   0   'False
               Height          =   288
               Index           =   1
               Left            =   960
               TabIndex        =   59
               Top             =   0
               Width           =   372
            End
            Begin VB.TextBox txtTFCode 
               Enabled         =   0   'False
               Height          =   288
               Index           =   0
               Left            =   0
               TabIndex        =   58
               Top             =   0
               Width           =   972
            End
         End
         Begin MSForms.ComboBox cboContact 
            Height          =   300
            Left            =   4950
            TabIndex        =   11
            Top             =   1890
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
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   5
            Left            =   5310
            TabIndex        =   5
            Top             =   180
            Width           =   612
            VariousPropertyBits=   679493659
            MaxLength       =   3
            Size            =   "1080;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   33
            Left            =   5700
            TabIndex        =   17
            Top             =   2820
            Width           =   2565
            VariousPropertyBits=   679493659
            MaxLength       =   50
            Size            =   "4524;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   285
            Index           =   32
            Left            =   1680
            TabIndex        =   9
            Top             =   1620
            Width           =   6492
            VariousPropertyBits=   -1466941413
            ScrollBars      =   2
            Size            =   "11451;503"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   27
            Left            =   4965
            TabIndex        =   15
            Top             =   2520
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   26
            Left            =   870
            TabIndex        =   14
            Top             =   2520
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   25
            Left            =   4965
            TabIndex        =   13
            Top             =   2220
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "2138;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   24
            Left            =   870
            TabIndex        =   12
            Top             =   2220
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "2138;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   3
            Left            =   1080
            TabIndex        =   4
            Top             =   450
            Width           =   372
            VariousPropertyBits=   679493659
            MaxLength       =   1
            Size            =   "656;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   285
            Index           =   4
            Left            =   1680
            TabIndex        =   7
            Top             =   1050
            Width           =   6492
            VariousPropertyBits=   -1466941413
            ScrollBars      =   2
            Size            =   "11451;503"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   9
            Left            =   870
            TabIndex        =   10
            Top             =   1920
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "2138;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   300
            Index           =   10
            Left            =   870
            TabIndex        =   16
            Top             =   2820
            Width           =   1215
            VariousPropertyBits=   679493659
            MaxLength       =   9
            Size            =   "2143;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTrademark 
            Height          =   285
            Index           =   6
            Left            =   1920
            TabIndex        =   8
            Top             =   1320
            Width           =   6252
            VariousPropertyBits=   -1466941413
            MaxLength       =   200
            ScrollBars      =   2
            Size            =   "11028;503"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label lblTS 
            Caption         =   "查名代號："
            Height          =   240
            Left            =   4350
            TabIndex        =   127
            Top             =   735
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "申請國家："
            Height          =   255
            Left            =   4350
            TabIndex        =   69
            Top             =   210
            Width           =   975
         End
         Begin VB.Label lblNation 
            Height          =   255
            Left            =   5970
            TabIndex        =   70
            Top             =   210
            Width           =   2175
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "接洽人："
            Height          =   180
            Left            =   4185
            TabIndex        =   120
            Top             =   1980
            Width           =   720
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "代理人彼所案號："
            Height          =   180
            Left            =   4200
            TabIndex        =   118
            Top             =   2880
            Width           =   1440
         End
         Begin VB.Label Label31 
            Caption         =   "商品組群（699）："
            Height          =   255
            Left            =   90
            TabIndex        =   117
            Top             =   1650
            Width           =   1575
         End
         Begin MSForms.Label lblPetition 
            Height          =   255
            Index           =   4
            Left            =   6240
            TabIndex        =   106
            Top             =   2535
            Width           =   1935
            BackColor       =   16777215
            VariousPropertyBits=   27
            Size            =   "3413;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "申請人5："
            Height          =   180
            Left            =   4185
            TabIndex        =   105
            Top             =   2580
            Width           =   810
         End
         Begin MSForms.Label lblPetition 
            Height          =   255
            Index           =   3
            Left            =   2145
            TabIndex        =   104
            Top             =   2550
            Width           =   1935
            BackColor       =   16777215
            VariousPropertyBits=   27
            Size            =   "3413;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "申請人4："
            Height          =   180
            Left            =   90
            TabIndex        =   103
            Top             =   2580
            Width           =   810
         End
         Begin MSForms.Label lblPetition 
            Height          =   255
            Index           =   2
            Left            =   6240
            TabIndex        =   102
            Top             =   2235
            Width           =   1935
            BackColor       =   16777215
            VariousPropertyBits=   27
            Size            =   "3413;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "申請人3："
            Height          =   180
            Index           =   0
            Left            =   4185
            TabIndex        =   101
            Top             =   2280
            Width           =   810
         End
         Begin MSForms.Label lblPetition 
            Height          =   255
            Index           =   1
            Left            =   2145
            TabIndex        =   100
            Top             =   2235
            Width           =   1935
            BackColor       =   16777215
            VariousPropertyBits=   27
            Size            =   "3413;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "申請人2："
            Height          =   180
            Left            =   90
            TabIndex        =   99
            Top             =   2280
            Width           =   810
         End
         Begin VB.Label Label8 
            Caption         =   "商標名稱（140）："
            Height          =   255
            Left            =   90
            TabIndex        =   86
            Top             =   1350
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "本所案號："
            Height          =   252
            Left            =   120
            TabIndex        =   71
            Top             =   180
            Width           =   972
         End
         Begin VB.Label Label10 
            Caption         =   "商標種類："
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   450
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "商品類別（395）："
            Height          =   255
            Left            =   90
            TabIndex        =   67
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "代理人："
            Height          =   180
            Left            =   90
            TabIndex        =   66
            Top             =   2880
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "申請人1："
            Height          =   180
            Left            =   90
            TabIndex        =   65
            Top             =   1980
            Width           =   810
         End
         Begin MSForms.Label lblPetition 
            Height          =   255
            Index           =   0
            Left            =   2145
            TabIndex        =   64
            Top             =   1920
            Width           =   1935
            BackColor       =   16777215
            VariousPropertyBits=   27
            Size            =   "3413;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblAgent 
            Height          =   255
            Left            =   2145
            TabIndex        =   63
            Top             =   2850
            Width           =   1935
            VariousPropertyBits=   27
            Size            =   "3413;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label lblTrademarkKind 
            Height          =   210
            Left            =   1560
            TabIndex        =   62
            Top             =   495
            Width           =   2655
         End
      End
      Begin VB.TextBox txtRecieveCode 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1092
         TabIndex        =   0
         Top             =   0
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "證書形式："
         Height          =   285
         Index           =   141
         Left            =   6090
         TabIndex        =   131
         Top             =   5370
         Width           =   945
      End
      Begin VB.Label Label28 
         Caption         =   "(1.電子 2.紙本)"
         Height          =   315
         Index           =   1
         Left            =   7410
         TabIndex        =   130
         Top             =   5370
         Width           =   1215
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   35
         Left            =   7050
         TabIndex        =   36
         Top             =   5310
         Width           =   315
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "556;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   19
         Left            =   7470
         TabIndex        =   41
         Top             =   4980
         Width           =   1095
         VariousPropertyBits=   679493659
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   34
         Left            =   3270
         TabIndex        =   31
         Top             =   4950
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   23
         Left            =   6360
         TabIndex        =   22
         Top             =   4050
         Width           =   2085
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "3678;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   22
         Left            =   5640
         TabIndex        =   20
         Top             =   3750
         Width           =   2805
         VariousPropertyBits=   679493659
         MaxLength       =   30
         Size            =   "4948;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   2
         Left            =   5040
         TabIndex        =   3
         Top             =   300
         Width           =   375
         VariousPropertyBits=   679493659
         MaxLength       =   2
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   1
         Left            =   1095
         TabIndex        =   2
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
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   16
         Left            =   3120
         TabIndex        =   19
         Top             =   3750
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   17
         Left            =   3660
         TabIndex        =   24
         Top             =   4320
         Width           =   495
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   18
         Left            =   3270
         TabIndex        =   27
         Top             =   4650
         Width           =   1095
         VariousPropertyBits=   679493659
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   12
         Left            =   1050
         TabIndex        =   21
         Top             =   4050
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   6
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   14
         Left            =   1050
         TabIndex        =   26
         Top             =   4650
         Width           =   1095
         VariousPropertyBits=   679493659
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   15
         Left            =   1050
         TabIndex        =   29
         Top             =   4950
         Width           =   1095
         VariousPropertyBits=   679493659
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   11
         Left            =   1050
         TabIndex        =   18
         Top             =   3750
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   13
         Left            =   1050
         TabIndex        =   23
         Top             =   4350
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   5
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTrademark 
         Height          =   300
         Index           =   0
         Left            =   5040
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCFTna048 
         AutoSize        =   -1  'True
         Caption         =   "緬甸舊案號："
         Height          =   180
         Left            =   2190
         TabIndex        =   129
         Top             =   4830
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label27 
         Caption         =   "後金:"
         Height          =   255
         Left            =   7020
         TabIndex        =   125
         Top             =   5010
         Width           =   435
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "預定收款日："
         Height          =   180
         Left            =   2190
         TabIndex        =   119
         Top             =   4980
         Width           =   1080
      End
      Begin VB.Label Label19 
         Caption         =   "分所案號："
         Height          =   255
         Left            =   5400
         TabIndex        =   96
         Top             =   4050
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "客戶案件案號："
         Height          =   255
         Left            =   4290
         TabIndex        =   95
         Top             =   3750
         Width           =   1305
      End
      Begin VB.Label lblDepartment 
         Height          =   255
         Left            =   4140
         TabIndex        =   89
         Top             =   4050
         Width           =   975
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "是否開電腦收據：           （N：不開)"
         Height          =   180
         Left            =   2190
         TabIndex        =   88
         Top             =   4350
         Width           =   2835
      End
      Begin VB.Label Label11 
         Caption         =   "法定期限："
         Height          =   255
         Left            =   2190
         TabIndex        =   87
         Top             =   3750
         Width           =   975
      End
      Begin MSForms.Label lblSales 
         Height          =   255
         Left            =   2190
         TabIndex        =   85
         Top             =   4050
         Width           =   975
         VariousPropertyBits=   27
         Size            =   "1720;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCaseSource 
         Height          =   255
         Left            =   5460
         TabIndex        =   83
         Top             =   345
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "案件來源："
         Height          =   255
         Left            =   4080
         TabIndex        =   82
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "案件性質："
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "規費："
         Height          =   255
         Left            =   2550
         TabIndex        =   80
         Top             =   4650
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "業務區："
         Height          =   255
         Left            =   3360
         TabIndex        =   79
         Top             =   4050
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "智權人員："
         Height          =   255
         Left            =   90
         TabIndex        =   78
         Top             =   4050
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "點數："
         Height          =   255
         Left            =   90
         TabIndex        =   77
         Top             =   4980
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "費用："
         Height          =   255
         Left            =   90
         TabIndex        =   76
         Top             =   4650
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "郵遞區號："
         Height          =   255
         Left            =   90
         TabIndex        =   75
         Top             =   4350
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "本所期限："
         Height          =   255
         Left            =   90
         TabIndex        =   74
         Top             =   3750
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "收文日："
         Height          =   255
         Left            =   4080
         TabIndex        =   73
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "收文號："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   72
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblCaseProperty 
         Height          =   255
         Left            =   1755
         TabIndex        =   84
         Top             =   315
         Width           =   2160
      End
   End
End
Attribute VB_Name = "frm010004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/16 Form2.0已修改 txtTradeMark()/lblPetition()/lblPetitionName()/lblSales/lblAgent/lblPromoter/cboContact
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
Dim oMailCount As String

'Add by Morgan 2004/4/15
'是否已觸發 Form Active 事件
Dim bolActive As Boolean
'add by nickc 2007/12/1z2
Dim IsSaveData As Boolean
Dim strAppNo1 As String '申請人1編號
Dim m_EP06 As String 'Add By Sindy 2012/5/8
Dim dblAmt As Double, dblPFee As Double, dblTFee As Double, m_CP150 As String 'Add By Sindy 2012/11/06
Dim dblChkAmt As Double 'Add By Sindy 2012/12/10
'Added by Lydia 2020/02/03
Dim dblCu183 As Double '個人之應收帳款上限
Dim dblAmtR As Double, dblPFeeR As Double, dblTFeeR As Double '關係企業之應收帳款金額
'end 2020/02/03

'Added by Lydia 2015/11/12 新增查名單對應
Dim m_AttachPath As String
Dim m_PrevForm As Form
Public Tmpfrm090130 As Form
Public TMQList As String
Dim bolOpen130 As Boolean 'Added by Lydia 2016/03/28 是否開啟過查名代號表單
'Added by Lydia 2019/02/14
Dim m_SalesST15 As String '畫面上智權人員的收文部門
Dim m_Tuser As String '創新業務部預設收文人員
'Added by Lydia 2019/09/16
Dim m_SalesST06 As String '智權人員的所別
'Added by Lydia 2020/05/20 法律所案源收文
Dim m_LOS02 As String '案源案件類型
Dim t_LOSkind As String '案件性質=>案源案件類型
Dim m_LOS15 As String '案源單號
Dim m_CaseNa239() As String ''Added by Lydia 2020/11/19 CFT英國脫歐案管制：歐盟案案號
Dim strXState(1 To 50) As String, strYState As String 'Add By Sindy 2021/2/1 回傳客戶狀態
Dim m_TM22 As String  'add by sonia 2021/3/31
Dim m_TM58 As String 'Added by Lydia 2021/04/15 案件備註
'Mark by Lydia 2022/09/06 改抓特殊設定
'Private Const cnt應收帳款檢查排除 As String = "74018,70005" 'Added by Lydia 2022/06/15 應收帳款上限檢查排除特定人員: 如果人員有異動, 請一併修改接洽單frm090801和收文frm010004~frm010007

'Added by Lydia 2022/09/05 櫃台收文模組化
Private Const 收文存檔模組化啟用日 = 20220906 '完成後先開始使用
Dim modCP() As String, modBase() As String ' 收文 和 基本檔
Dim m_strControl As String  '齊備日管制
Dim mType As String, mCaseNo As String  '特殊管制
'Add By Sindy 2023/5/30
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_PrevFormIR As Form '前一畫面
Dim m_bolRecvOK As Boolean '是否收完文
Dim m_strMCR11 As String '多案收文時,第一筆的總收文號
'2023/5/30 END


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
      modBase(5) = txtTrademark(6)  '案件名稱(中)
      modBase(8) = txtTrademark(3) '商標種類
      modBase(9) = Trim(txtTrademark(4)) '商品類別
      modBase(10) = Trim(txtTrademark(5)) '申請國家
      modBase(32) = Trim(txtTrademark(32))  '商品組群
      modBase(35) = Trim(txtTrademark(22))  '客戶案件案號
      modBase(34) = Trim(txtTrademark(23))  '分所案號
      modBase(15) = Trim(txtTrademark(7)) '審定號
      '申請人1~5
      modBase(23) = ChangeCustomerL(txtTrademark(9))
      modBase(78) = ChangeCustomerL(txtTrademark(24))
      modBase(79) = ChangeCustomerL(txtTrademark(25))
      modBase(80) = ChangeCustomerL(txtTrademark(26))
      modBase(81) = ChangeCustomerL(txtTrademark(27))
      '代理人
      modBase(44) = ChangeCustomerL(txtTrademark(10))
      modBase(45) = Trim(txtTrademark(33)) '彼所案號
      '申請人聯絡人編號
      If cboContact.Locked = False Then
         If cboContact.ListIndex >= 0 Then
            modBase(123) = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
            If Val(modBase(123)) > 0 Then
            'Add by Morgan 2008/8/7 若個案接洽人與客戶檔的預設接洽人相同時不必設定
               PUB_GetContact modBase(23), strTmpA, True
               If modBase(123) = strTmpA Then
                  modBase(123) = ""
               End If
            '排除空白=00
            ElseIf modBase(123) = "00" And Trim(cboContact.Text) = "" Then
               modBase(123) = ""
            End If
         End If
      End If
      'Add By Sindy 2022/12/7
      If txtTrademark(35).Visible = True Then
         modBase(136) = Trim(txtTrademark(35)) '證書形式
      End If
      '2022/12/7
      
      modCP(9) = txtRecieveCode  '收文號
      modCP(5) = ChangeTStringToWString(txtTrademark(0)) '收文日
      modCP(6) = ChangeTStringToWString(txtTrademark(11)) '本所期限
      modCP(7) = ChangeTStringToWString(txtTrademark(16))  '法定期限
      modCP(10) = Trim(txtTrademark(1)) '案件性質
      modCP(11) = Trim(txtTrademark(2)) '案件來源
      modCP(12) = GetST15(txtTrademark(12))
      modCP(13) = Trim(txtTrademark(12))       '智權人員
      modCP(14) = Trim(txtTrademark(21))    '承辦人
      modCP(16) = txtTrademark(14)    '費用
      modCP(17) = txtTrademark(18)    '規費
      modCP(18) = txtTrademark(15)    '點數
      modCP(19) = txtTrademark(19)    '後金
      modCP(32) = txtTrademark(17) '是否開電腦收據
      modCP(33) = douStPrice '標準價
      modCP(34) = douLowPrice '底價
      modCP(64) = txtCP64
      '讓與人1-5,受讓人1-5
      modCP(56) = ChangeCustomerL(txtTrademark(20))
      modCP(89) = ChangeCustomerL(txtTrademark(28))
      modCP(90) = ChangeCustomerL(txtTrademark(29))
      modCP(91) = ChangeCustomerL(txtTrademark(30))
      modCP(92) = ChangeCustomerL(txtTrademark(31))
      '有★★的應收帳款簽核控管
      If Check2.Visible = True Then
         modCP(150) = IIf(Check2.Value = 1, "Y", "")
      End If
      '電子送件
      'Modified by Lydia 2022/09/20 因為P案有Trigger會自動設定電子送件CP118 = Y, 所以改成兩個判斷
      'If chkWebApp.Visible = True And chkWebApp.Value = 1 Then
      '    modCP(118) = "Y"
      If chkWebApp.Visible = True Then
         If chkWebApp.Value = 1 Then
             modCP(118) = "YY"
         Else
             modCP(118) = "YN"
         End If
      'end 2022/09/20
      End If
      '特殊管制
      mType = "": mCaseNo = ""
      If txtSystem = "CFT" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" Then
          mType = "CFT英國脫歐案"
          mCaseNo = m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4)
      ElseIf m_LOS02 <> "" And m_LOS15 <> "" Then
          mType = "LOS案源收文"
          mCaseNo = m_LOS02 & "," & m_LOS15
      ElseIf txtSystem = "CFT" And txtCFTNa048 <> "" Then
          mType = "CFT緬甸重新申請案"
          mCaseNo = Trim(txtCFTNa048)
      ElseIf txtSystem = "T" And TMQList <> "" Then
          mType = "T查名單"
          mCaseNo = TMQList
      'Modify By Sindy 2025/8/18 發生了案源+信件沖銷 ex:FCP-057445/FCL-011034
      'Add By Sindy 2023/5/30
      'ElseIf m_strIR01 <> "" Then
      End If
      If m_strIR01 <> "" Then
      '2025/8/18 END
          'mType = "信件沖銷"
          mType = mType & "-信件沖銷" 'Modify By Sindy 2025/8/18 + "-"
          'If m_bMRecvBatch = True Then mType = mType & "-多案收文"
          'Modify By Sindy 2025/8/18 + IIf(mCaseNo <> "", mCaseNo & "-", "") &
          mCaseNo = IIf(mCaseNo <> "", mCaseNo & "-", "") & m_strIR01 & "," & m_strIR02 & "," & m_strIR03 & "," & m_strIR04
      End If
      
      '齊備日  --m_strControl
       Call GetStrControl
   End If
End Sub
                                  
'Added by Lydia 2022/09/05 處理齊備日相關欄位的變數
Private Sub GetStrControl()
      m_strControl = ""
      If Frame21.Visible = True Then
         '文件是否齊備(101申請)、資料是否齊備
         If textEP06.Visible = True Then
             m_strControl = m_strControl & ",EP06|" & Trim(textEP06) & "|" & m_EP06
         End If
         '是否會稿
         If textEP34.Visible = True Then
             m_strControl = m_strControl & ",EP34|" & Trim(textEP34)
         End If
         '是否急件
         If textCP122.Visible = True Then
             m_strControl = m_strControl & ",CP122|" & Trim(textCP122)
         End If
         '查名是否齊備
         If textCP143.Visible = True Then
             m_strControl = m_strControl & ",CP143|" & IIf(textCP143 = "Y", strSrvDate(1), IIf(textCP143 = "N", "0", ""))
         End If
         If m_strControl <> "" Then m_strControl = Mid(m_strControl, 2)
      End If
End Sub

'Added by Lydia 2015/11/12
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub
'end 2015/11/12

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      '2011/4/22 MODIFY BY SONIA 分所智權人員則多一天
      'txtTrademark(34) = PUB_GetWorkDayAfterSysDate(CDbl(txtTrademark(0)) + 19110000, 5)
      'Modified by Lydia 2019/09/16
      'If PUB_GetST06(txtTrademark(12)) <> "1" Then
      If m_SalesST06 <> "1" Then
         txtTrademark(34) = PUB_GetWorkDayAfterSysDate(CDbl(txtTrademark(0)) + 19110000, 6)
      Else
         txtTrademark(34) = PUB_GetWorkDayAfterSysDate(CDbl(txtTrademark(0)) + 19110000, 5)
      End If
      '2011/4/22 END
      txtTrademark(34).Locked = True
   Else
      txtTrademark(34).Locked = False
   End If
End Sub

'Add By Sindy 2013/8/26
Private Sub chkWebApp_LostFocus()
   If chkWebApp.Visible = True Then
      'Modified by Lydia 2019/12/24 +限制:101申請才預設所限和法限為收文日
      'Modify By Sindy 2020/1/16 申請案才檢查
      'Modified by Lydia 2020/03/12 請取消商申電子送件有關當日期限的管制. 若該案須設期限管制時, 由承辦人載入接洽單, 於收文時輸入即可.
      'If txtTrademark(1) = "101" Then
      '   If chkWebApp.Value = 1 Then
     '    'If chkWebApp.Visible = True And txtTrademark(1) = "101" Then
      ''2020/1/16 END
      '      txtTrademark(11) = txtTrademark(0)
      '      txtTrademark(16) = txtTrademark(0)
      '   Else
      '      txtTrademark(11) = ""
      '      txtTrademark(16) = ""
      '   End If
     ' End If
      'end 2020/03/12
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim varSaveCursor, strAuto1 As String, strAuto2 As String, i As Integer
Dim tmpArr As Variant
Dim intMoney As Long  '倍數
Dim strFee As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim intGoodCnt As Integer 'Add By Sindy 2012/3/1
Dim bolChkFee18 As Boolean 'Add By Sindy 2013/12/27
Dim mBillNo As String, mMemo As String 'Added by Lydia 2019/05/13
Dim bolSaveOK As Boolean, mRetVal As String 'Added by Lydia 2022/09/05

   If Index = 0 Then '確定
      'Add by Amy 2021/12/16檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me, , True) = False Then
         Exit Sub
      End If
     
      'Added by Lydia 2017/07/31 預設和檢查-所有內部收文, 若有輸入本所期限或法定期限者
      'Modified by Lyddia 2023/11/08 傳入必需欄位
      'If PUB_CheckCP0607(0, txtTrademark(11), txtTrademark(16)) = False Then Exit Sub
      If PUB_CheckCP0607(0, txtTrademark(11), txtTrademark(16), IIf(frm010001.intModifyKind = 0, "Y", ""), txtTrademark(5), txtSystem, txtTrademark(1)) = False Then Exit Sub
      
      'Added by Lydia 2020/05/20 法律所案源收文：5/28台灣案之B1、B2及C收文時，增加"案源單號"欄位，B1、B2一定要輸入，C若未輸入則提醒'請確認接洽單沒有案源單號？'，案源單號更新至該筆收文的CP162。
      'Mark by Lydia 2020/06/10 重整判斷,以案源單的案源類型為準; 保留舊程式
'      If frm010001.intModifyKind = 0 And strSrvDate(1) >= 法律所案源收文啟用日 And (txtSystem = "FCT" Or txtSystem = "T" Or txtSystem = "TC") Then
'           t_LOSkind = PUB_GetLOSkind(txtSystem, txtTrademark(1), txtTrademark(5))
'           If Left(t_LOSkind, 1) = "B" And m_LOS15 = "" Then
'               MsgBox "請先回前畫面輸入案源單號！", vbCritical, "檢核案源單號"
'               Exit Sub
'           End If
'           'Added by Lydia 2020/06/04 法律所案源收文：判斷是否為補收文=>案源類別
'           strExc(1) = PUB_GetLOSplus(txtSystem, txtCode(0), txtCode(1), txtCode(2), txtTrademark(1), txtTrademark(5), t_LOSkind)
'           If m_LOS02 = "" And Left(strExc(1), 1) = "B" Then
'               If MsgBox("請確認接洽單沒有案源單號？" & vbCrLf & "選擇""是""會繼續作業，選擇""否""回前畫面輸入案源單號", vbInformation + vbYesNo + vbDefaultButton1, "檢核案源單號") = vbNo Then
'                   Exit Sub
'               End If
'           End If
'           'end 2020/06/04
'           'Modified by Lydia 2020/06/04
'           'If Left(t_LOSkind, 1) = "C" And m_LOS15 = "" Then
'           If ((Left(t_LOSkind, 1) = "C" And txtCode(0) = "") Or (Left(strExc(1), 1) = "C" And txtCode(0) <> "")) And m_LOS15 = "" Then
'               If MsgBox("請確認接洽單沒有案源單號？" & vbCrLf & "選擇""是""會繼續作業，選擇""否""回前畫面輸入案源單號", vbInformation + vbYesNo + vbDefaultButton1, "檢核案源單號") = vbNo Then
'                   Exit Sub
'               End If
'           End If
'      End If
'      'end 2020/05/20
      
      'Add By Sindy 2024/9/4 CF案申請國家不可為台灣
      If Left(txtSystem, 2) = "CF" And txtTrademark(5) = "000" Then
         MsgBox "CF案申請國家不可為台灣！", vbExclamation
         txtTrademark(5).SetFocus
         Call txtTrademark_GotFocus(5)
         Exit Sub
      End If
      '2024/9/4 END
      
      If frm010001.intModifyKind = 0 And strSrvDate(1) >= 法律所案源收文啟用日 And (txtSystem = "FCT" Or txtSystem = "T" Or txtSystem = "TC") Then
           If txtTrademark(5) <> "000" Then '非台灣案=>清空資料
               m_LOS02 = ""
               m_LOS15 = ""
           Else
                t_LOSkind = PUB_GetLOSkind(txtSystem, txtTrademark(1), txtTrademark(5))
                If Left(t_LOSkind, 1) = "B" And m_LOS15 = "" Then
                    MsgBox "請先回前畫面輸入案源單號！", vbCritical, "檢核案源單號"
                    Exit Sub
                End If
                '判斷是否為補收文=>案源類別
                strExc(1) = PUB_GetLOSplus(txtSystem, txtCode(0), txtCode(1), txtCode(2), txtTrademark(1), txtTrademark(5), IIf(t_LOSkind = "", "C", t_LOSkind))
                If m_LOS02 = "" And strExc(1) <> "" And m_LOS15 = "" Then
                    'Modified by Lydia 2020/07/20 預設"否"要輸入案源單號 vbDefaultButton2 (原本預設vbDefaultButton1)
                    If MsgBox("請確認接洽單左上角是否沒有案源單號？" & vbCrLf & "選擇""是""會繼續作業，選擇""否""回前畫面輸入案源單號", vbInformation + vbYesNo + vbDefaultButton2, "檢核案源單號") = vbNo Then
                        Exit Sub
                    End If
                End If
           End If
      End If
      'end 2020/06/10
      
      'Modified by Lydia 2019/09/16 +st06
      'm_SalesST15 = GetST15(txtTrademark(12)) 'Added by Lydia 2019/02/14
      m_SalesST15 = GetST15(txtTrademark(12), , , m_SalesST06)
      
      'Added by Lydia 2020/04/08 檢查案件或智權人員是否為法務部
      If PUB_ChkSalesL(txtSystem, txtTrademark(12).Text) = False Then
          txtTrademark(12).SetFocus
          Call txtTrademark_GotFocus(12)
          Exit Sub
      End If
      'end 2020/04/08
      
      'Added by Lydia 2020/12/15 CFT緬甸重新申請案：檢查
      If CheckCFTna048 = False Then
          Exit Sub
      End If
      'end 2020/12/15
      
      'Added by Lydia 2021/09/10 修正畫面所有含跳行符號的文字框; 9/10 FCT-47909收文申請,彼所案號中間有換行
      PUB_FilterFormText Me
      
      varSaveCursor = Screen.MousePointer
      Screen.MousePointer = vbHourglass
      For i = 0 To 21
         If i = 7 Or i = 8 Then GoTo GoToNext
         'Add By Cheng 2001/12/12
         If i = 11 Then
            If txtTrademark(11).Text <> "" Then
               If CheckIsTaiwanDate(txtTrademark(11).Text) Then
                  If CheckReKey(txtTrademark(11)) Then
                     If Val(txtTrademark(11)) = Val(GetTaiwanTodayDate) Then
                        ShowMsg "此案件已屆本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
                     End If
                     If Val(txtTrademark(11)) < Val(GetTaiwanTodayDate) Then
                        ShowMsg "此案件已逾本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
                     End If
                  Else
                     Screen.MousePointer = varSaveCursor
                     Exit Sub
                  End If
                End If
            '93.3.24 ADD BY SONIA
            Else
               If txtCode(0) = "" And txtTrademark(1).Text = "102" Then
                  ShowMsg "延展新案, 請輸入延展本所期限!"
                  txtTrademark(11).SetFocus
                  txtTrademark_GotFocus (11)
                  Exit For
               End If
            '93.3.24 END
               'Add by Amy 2014/10/22 +大陸分割案控制
               If txtTrademark(1).Text = "308" And txtTrademark(5).Text = "020" Then
                  ShowMsg "大陸分割案, 請輸入分割本所期限!"
                  txtTrademark(11).SetFocus
                  txtTrademark_GotFocus (11)
                  Exit For
               End If
               'end 2014/10/22
            End If
         End If
         '93.3.24 ADD BY SONIA
         If i = 16 Then
            If txtTrademark(16).Text = "" Then
               If txtCode(0) = "" And txtTrademark(1).Text = "102" Then
                  ShowMsg "延展新案, 請輸入延展法定期限!"
                  txtTrademark(16).SetFocus
                  txtTrademark_GotFocus (16)
                  Exit For
               End If
               'Add by Amy 2014/10/22 +大陸分割案控制
                If txtTrademark(1).Text = "308" And txtTrademark(5).Text = "020" Then
                    ShowMsg "大陸分割案, 請輸入分割法定期限!"
                    txtTrademark(16).SetFocus
                    txtTrademark_GotFocus (16)
                    Exit For
                End If
                'end 2014/10/22
            End If
         End If
         '93.3.24 END
         
         'Add By Sindy 2010/12/31 費用檢查提到存檔前檢查
         If i = 14 Then
            '郭 請作單 X14843050 不管
            'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
            'modify by sonia 2014/9/11 取消X69514,已轉外專
            If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" And _
               Mid(txtTrademark(9), 1, 8) <> "X3928904" And Mid(txtTrademark(24), 1, 8) <> "X3928904" And Mid(txtTrademark(25), 1, 8) <> "X3928904" And Mid(txtTrademark(26), 1, 8) <> "X3928904" And Mid(txtTrademark(27), 1, 8) <> "X3928904" Then
                If ClsPDGetCaseLowPrice(txtSystem, txtTrademark(5), txtTrademark(1), douStPrice, douLowPrice) = 1 Then
                End If
                If txtTrademark(14) <> "" Then
                   If txtTrademark(5) = "000" And _
                     (Trim(txtTrademark(1)) = "101" Or _
                      Trim(txtTrademark(1)) = "102" Or _
                      Trim(txtTrademark(1)) = "715" Or _
                      Trim(txtTrademark(1)) = "716" Or _
                      Trim(txtTrademark(1)) = "717" Or _
                      Trim(txtTrademark(1)) = "601" Or _
                      Trim(txtTrademark(1)) = "603" Or _
                      Trim(txtTrademark(1)) = "605") Then
                     If Val(txtTrademark(18)) > 0 And txtTrademark(4) <> "" Then
                        tmpArr = Split(IIf(Right(txtTrademark(4), 1) = ",", Mid(txtTrademark(4), 1, Len(txtTrademark(4)) - 1), txtTrademark(4)), ",")
                        If ClsPDGetCaseFee_T(txtSystem, txtTrademark(5), txtTrademark(1), Val(txtTrademark(14)), Val(txtTrademark(18)), Val(UBound(tmpArr))) = 0 Then
                           Screen.MousePointer = varSaveCursor
                           Exit Sub
                        End If
                     End If
                   'Added by Lydia 2020/05/20 法律所案源收文：台灣案之B1及B2案件性質都不可收費用。
                   'Modified by Lydia 2020/07/03 不確定國外部是否收費, 先排除
                   'ElseIf txtTrademark(5) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
                   ElseIf txtSystem <> "FCT" And txtTrademark(5) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
                       If Val(txtTrademark(14)) > 0 Then
                           MsgBox "【B類】案源接洽單之費用、規費、點數都必須為 0", vbExclamation, "檢核案源單號"
                           Screen.MousePointer = varSaveCursor
                           Exit Sub
                       End If
                   'end 2020/05/20
                   Else
                     'Add by Amy 2014/10/22 +大陸分割控制
                     If txtCode(0) <> "" And txtTrademark(1).Text = "308" And txtTrademark(5).Text = "020" Then
                        '大陸分割母案不需檢查CaseFree費用
                     Else
                        'MODIFY BY SONIA 2014/7/17 +傳規費 CFP-027024
                        If ClsPDGetCaseFee(txtSystem, txtTrademark(5), txtTrademark(1), Val(txtTrademark(14)), Val(txtTrademark(18))) = 0 Then
                            Screen.MousePointer = varSaveCursor
                            Exit Sub
                        End If
                     End If
                     'end 2014/10/22
                   End If
                'txtTrademark(14)="" Add by Amy 2014/10/22
                Else
                    If txtCode(0) = "" And txtTrademark(1).Text = "308" And txtTrademark(5).Text = "020" Then
                        '大陸分割子案費用需大於0
                        If Val(txtTrademark(14)) = 0 Then
                            Screen.MousePointer = varSaveCursor
                            ShowMsg "大陸分割案, 請輸入分割費用!"
                            txtTrademark(14).SetFocus
                            Exit Sub
                        End If
                    End If
                'end 2014/10/22
                End If
            End If
         End If
         
         If i = 15 Then
            'Add By Sindy 2010/12/31 點數檢查提到存檔前檢查
            '郭 請作單 X14843050 不管
            'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
            'modify by sonia 2014/9/11 取消X69514,已轉外專
            If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" And _
               Mid(txtTrademark(9), 1, 8) <> "X3928904" And Mid(txtTrademark(24), 1, 8) <> "X3928904" And Mid(txtTrademark(25), 1, 8) <> "X3928904" And Mid(txtTrademark(26), 1, 8) <> "X3928904" And Mid(txtTrademark(27), 1, 8) <> "X3928904" Then
               If txtTrademark(15) = "" Then
                  If txtTrademark(14) <> "" Or txtTrademark(18) <> "" Then
                     ShowMsg MsgText(1035)
                     Screen.MousePointer = varSaveCursor
                     Exit Sub
                  End If
               'Added by Lydia 2020/05/20 法律所案源收文：台灣案之B1及B2案件性質都不可收費用。
               'Modified by Lydia 2020/07/03 不確定國外部是否收費, 先排除
               'ElseIf txtTrademark(5) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
               ElseIf txtSystem <> "FCT" And txtTrademark(5) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
                   If Val(txtTrademark(15)) > 0 Then
                       MsgBox "【B類】案源接洽單之費用、規費、點數都必須為 0", vbExclamation, "檢核案源單號"
                       Screen.MousePointer = varSaveCursor
                       Exit Sub
                   End If
               'end 2020/05/20
               ElseIf txtTrademark(14) <> "" Or txtTrademark(18) <> "" Then
   '               If Format((Val(txtTrademark(14)) - Val(txtTrademark(18))) / 1000, "0.0") <> Format(Val(txtTrademark(15)), "0.0") Then
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
            'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
            'modify by sonia 2014/9/11 取消X69514,已轉外專
            If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" And _
               Mid(txtTrademark(9), 1, 8) <> "X3928904" And Mid(txtTrademark(24), 1, 8) <> "X3928904" And Mid(txtTrademark(25), 1, 8) <> "X3928904" And Mid(txtTrademark(26), 1, 8) <> "X3928904" And Mid(txtTrademark(27), 1, 8) <> "X3928904" Then
               If txtTrademark(14) <> "" Or txtTrademark(18) <> "" Then
                  If Format((Val(txtTrademark(14)) - Val(txtTrademark(18))) / 1000, "0.0") <> Format(Val(txtTrademark(15)), "0.0") Then
                     ShowMsg MsgText(1036)
                     Screen.MousePointer = varSaveCursor
   '                  txtTrademark(i).SetFocus
   '                  txtTrademark_GotFocus (i)
                     Exit Sub
                  End If
               End If
           End If
         End If
         
         'Add By Sindy 2013/7/17
         '台灣案時, 檢查商標種類不可輸入2,4,5,6
         If i = 3 Then
            If Trim(txtTrademark(5)) = "000" And _
               (Trim(txtTrademark(3)) = "2" Or Trim(txtTrademark(3)) = "4" Or Trim(txtTrademark(3)) = "5" Or Trim(txtTrademark(3)) = "6") Then
               MsgBox "台灣案時, 商標種類不可輸入2,4,5,6！", vbExclamation + vbOKOnly
               txtTrademark(3).SetFocus
               Screen.MousePointer = varSaveCursor
               Exit Sub
            End If
         End If
         '2013/7/17 End
         
         'Add By Sindy 2010/12/31 規費檢查提到存檔前檢查
         bolChkFee18 = False 'Add By Sindy 2013/12/27
         If i = 18 Then
            If Val(txtTrademark(18)) > 0 Or txtTrademark(5) = "000" Then     'ADD by sonia 2014/7/17 加入未輸規費時不檢查此段,因為可能是依代理人帳單請款CFP-027024
               '郭 請作單 X14843050 不管
               'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
               'modify by sonia 2014/9/11 取消X69514,已轉外專
               If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" And _
                  Mid(txtTrademark(9), 1, 8) <> "X3928904" And Mid(txtTrademark(24), 1, 8) <> "X3928904" And Mid(txtTrademark(25), 1, 8) <> "X3928904" And Mid(txtTrademark(26), 1, 8) <> "X3928904" And Mid(txtTrademark(27), 1, 8) <> "X3928904" Then
                  If txtTrademark(4) <> "" Then
                     tmpArr = Split(IIf(Right(txtTrademark(4), 1) = ",", Mid(txtTrademark(4), 1, Len(txtTrademark(4)) - 1), txtTrademark(4)), ",")
                  'Add By Sindy 2012/3/1
                     intGoodCnt = UBound(tmpArr)
                  Else
                     intGoodCnt = 0
                  '2012/3/1 End
                  End If
   '2011/7/22 cancel by sonia 因fct超項費不分開收文
   '               If txtTrademark(1) = "101" And txtTrademark(5) = "000" And txtTrademark(18) = "4000" Then
   '                  MsgBox "台灣商標已修法, 申請案規費不可為 4000, 請與智權人員確認！", vbCritical
   '                  Screen.MousePointer = varSaveCursor
   '                  Exit Sub
   '               End If
                  If (txtTrademark(1) = "501" Or txtTrademark(1) = "502") And txtTrademark(5) = "000" Then
                     If txtTrademark(18) <> "2000" Then
                        MsgBox "台灣商標移轉或授權規費為 2000, 請確認！", vbCritical
                        Screen.MousePointer = varSaveCursor
                        Exit Sub
                     Else
                        bolChkFee18 = True 'Add By Sindy 2013/12/27
                     End If
                  End If
                  
   '               '2012/1/4 add by sonia 內商程序說要再加入,但須為2400,2700,3000的倍數
   '               If Me.txtSystem.Text = "T" And txtTrademark(5) = "000" And txtTrademark(1) = "101" Then
   ''                  If (Val(txtTrademark(18)) <> (3000# * (Val(UBound(TmpArr)) + 1))) And (Val(txtTrademark(18)) <> (2700# * (Val(UBound(TmpArr)) + 1))) And (Val(txtTrademark(18)) <> (2400# * (Val(UBound(TmpArr)) + 1))) Then
   ''                     MsgBox "台灣商標申請規費要" & str(3000# * (Val(UBound(TmpArr)) + 1)) & "或" & str(2700# * (Val(UBound(TmpArr)) + 1)) & "(電子申請) 或" & str(2400# * (Val(UBound(TmpArr)) + 1)) & "(電子申請), 請確認！", vbCritical
   '                  If (Val(txtTrademark(18)) <> (3000# * (Val(intGoodCnt) + 1))) And (Val(txtTrademark(18)) <> (2700# * (Val(intGoodCnt) + 1))) And (Val(txtTrademark(18)) <> (2400# * (Val(intGoodCnt) + 1))) Then
   '                     MsgBox "台灣商標申請規費要" & str(3000# * (Val(intGoodCnt) + 1)) & "或" & str(2700# * (Val(intGoodCnt) + 1)) & "(電子申請) 或" & str(2400# * (Val(intGoodCnt) + 1)) & "(電子申請), 請確認！", vbCritical
   '                     Screen.MousePointer = varSaveCursor
   '                     Exit Sub
   '                  End If
   '               End If
   '               '2012/1/4 end
                  'Add By Sindy 2012/3/16
                  '檢查時間點：規費欄跳離；
                  '控制：T、台灣、101申請時，若規費欄非 2400, 2700, 3000 時，彈訊息
                  If Trim(txtSystem) = "T" And Trim(txtTrademark(5)) = "000" And Trim(txtTrademark(1)) = "101" Then
                     If txtTrademark(3) <> "7" And txtTrademark(3) <> "8" Then 'Add By Sindy 2012/4/26 +if
                        If Val(txtTrademark(18)) <> 2400 And Val(txtTrademark(18)) <> 2700 And Val(txtTrademark(18)) <> 3000 Then
                           MsgBox "台灣商標申請案規費不符，只可為 2,400元, 2,700元, 3,000元，請退回智權人員修改！", vbExclamation + vbOKOnly
                           txtTrademark(18).SetFocus
                           Screen.MousePointer = varSaveCursor
                           Exit Sub
                        Else
                           bolChkFee18 = True 'Add By Sindy 2013/12/27
                        End If
                     End If
                  End If
                  '2012/3/16 End
                  
                  'Added by Lydia 2020/07/20 T台灣案的團標標章及證明標章，固定規費為4700,5000
                  If Trim(txtSystem) = "T" And Trim(txtTrademark(5)) = "000" And Trim(txtTrademark(1)) = "101" Then
                     If txtTrademark(3) = "7" Or txtTrademark(3) = "8" Then
                        If Val(txtTrademark(18)) <> 5000 And Val(txtTrademark(18)) <> 4700 Then
                           MsgBox "台灣商標申請案規費不符，只可為 4,700元, 5,000元，請退回智權人員修改！", vbExclamation + vbOKOnly
                           txtTrademark(18).SetFocus
                           Screen.MousePointer = varSaveCursor
                           Exit Sub
                        Else
                           bolChkFee18 = True
                        End If
                     End If
                  End If
                  'end 2020/07/20
                  If txtTrademark(5) = "000" And txtTrademark(4) <> "" Then
                     intMoney = 1
                     'Modify By Sindy 2011/3/14 若系統日的昨天為非工作天, 則以系統日的前一個工作天做比較
                     If ChkWorkDay(DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1))))) = False Then
                        If Val(txtTrademark(16)) < (Val(CompWorkDay(1, DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1)))), 1)) - 19110000) And Val(txtTrademark(16)) <> 0 Then
                           intMoney = 2
                        End If
                     '2011/3/14 End
                     Else
                        If Val(txtTrademark(16)) < Val(GetTaiwanTodayDate) And Val(txtTrademark(16)) <> 0 Then
                           intMoney = 2
                        End If
                     End If
                     'add by sonia 2021/3/31 台灣舊案延展以基本檔之專用期止日判斷是否過期FCT-029952
                     If txtTrademark(5) = "000" And txtCode(0) <> "" And txtTrademark(1) = "102" Then
                        If ChkWorkDay(DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1))))) = False Then
                           If Val(m_TM22) < (Val(CompWorkDay(1, DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1)))), 1)) - 19110000) Then
                              intMoney = 2
                           End If
                        Else
                           'modify by sonia 2021/7/20 延展新案修改時會錯誤,故加Val(m_TM22)>0條件T-234834
                           'If Val(m_TM22) < Val(GetTodayDate) Then
                           If Val(m_TM22) > 0 And Val(m_TM22) < Val(GetTodayDate) Then
                              intMoney = 2
                           End If
                        End If
                     End If
                     'end
                     If txtTrademark(1) = "715" Then
                        If Val(txtTrademark(18)) <> (1000# * (Val(UBound(tmpArr)) + 1) * intMoney) Then
                           MsgBox "台灣商標第一期註冊費規費要" & str(1000# * (Val(UBound(tmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
                           Screen.MousePointer = varSaveCursor
                           Exit Sub
                        Else
                           bolChkFee18 = True 'Add By Sindy 2013/12/27
                        End If
                     End If
                     If txtTrademark(1) = "716" Then
                        If Val(txtTrademark(18)) <> (1500# * (Val(UBound(tmpArr)) + 1) * intMoney) Then
                           MsgBox "台灣商標第二期註冊費規費要" & str(1500# * (Val(UBound(tmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
                           Screen.MousePointer = varSaveCursor
                           Exit Sub
                        Else
                           bolChkFee18 = True 'Add By Sindy 2013/12/27
                        End If
                     End If
                     If txtTrademark(1) = "717" Then
                        If Val(txtTrademark(18)) <> (2500# * (Val(UBound(tmpArr)) + 1) * intMoney) Then
   '                       MsgBox "台灣商標全期註冊費規費要" & str(2500# * (Val(UBound(TmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
                           'Modify By Sindy 2012/6/27 商標修法全期字樣拿掉
                           MsgBox "台灣商標註冊費規費要" & str(2500# * (Val(UBound(tmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
                           '2012/6/27 End
                           Screen.MousePointer = varSaveCursor
                           Exit Sub
                        Else
                           bolChkFee18 = True 'Add By Sindy 2013/12/27
                        End If
                     End If
                     '2014/5/5 modify by sonia 延展新案不檢查
                     If txtTrademark(1) = "102" Then
                        'Add By Sindy 2014/2/19
                        '檢查台灣商標延展跨類的規費
                        If Trim(txtTrademark(5)) = "000" And (Val(UBound(tmpArr)) + 1) > 1 Then
                           If PUB_ChkTawT102MClassFee(txtSystem, txtCode(0), txtCode(1), txtCode(2), Val(txtTrademark(18)), txtTrademark(16)) = "N" Then
                              Screen.MousePointer = varSaveCursor
                              Exit Sub
                           Else
                              bolChkFee18 = True
                           End If
                        Else
                        '2014/2/19 END
                           If Val(txtTrademark(18)) <> (4000# * (Val(UBound(tmpArr)) + 1) * intMoney) Then
                              MsgBox "台灣商標延展規費要" & str(4000# * (Val(UBound(tmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
                              Screen.MousePointer = varSaveCursor
                              Exit Sub
                           Else
                              bolChkFee18 = True 'Add By Sindy 2013/12/27
                           End If
                        End If
                     End If
                     If txtTrademark(1) = "601" Then
                        If Val(txtTrademark(18)) <> (4000# * (Val(UBound(tmpArr)) + 1) * intMoney) Then
                           MsgBox "台灣商標異議規費要" & str(4000# * (Val(UBound(tmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
                           Screen.MousePointer = varSaveCursor
                           Exit Sub
                        Else
                           bolChkFee18 = True 'Add By Sindy 2013/12/27
                        End If
                     End If
                     If (txtTrademark(1) = "603" Or txtTrademark(1) = "605") Then
                        If Val(txtTrademark(18)) <> (7000# * (Val(UBound(tmpArr)) + 1) * intMoney) Then
                           MsgBox "台灣商標評定或廢止規費要" & str(7000# * (Val(UBound(tmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
                           Screen.MousePointer = varSaveCursor
                           Exit Sub
                        Else
                           bolChkFee18 = True 'Add By Sindy 2013/12/27
                        End If
                     End If
                  End If
               End If
               
               'Added by Lydia 2020/05/20 法律所案源收文：台灣案之B1及B2案件性質都不可收費用。
               'Modified by Lydia 2020/07/03 不確定國外部是否收費, 先排除
               'If txtTrademark(5) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
               If txtSystem <> "FCT" And txtTrademark(5) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
                   If Val(txtTrademark(18)) > 0 Then
                       MsgBox "【B類】案源接洽單之費用、規費、點數都必須為 0", vbExclamation, "檢核案源單號"
                       Screen.MousePointer = varSaveCursor
                       Exit Sub
                   End If
                   bolChkFee18 = True 'B類不檢查規費
               End If
               'end 2020/05/20
               
               '郭 請作單 X14843050 不管
               'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
               'modify by sonia 2014/9/11 取消X69514,已轉外專
               If bolChkFee18 = False Then 'Add By Sindy 2013/12/27 +if
                  If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" And _
                     Mid(txtTrademark(9), 1, 8) <> "X3928904" And Mid(txtTrademark(24), 1, 8) <> "X3928904" And Mid(txtTrademark(25), 1, 8) <> "X3928904" And Mid(txtTrademark(26), 1, 8) <> "X3928904" And Mid(txtTrademark(27), 1, 8) <> "X3928904" Then
                     If txtTrademark(5) = "000" Then
                        strFee = GetTrademarkOfficialFee(txtSystem, txtTrademark(1), txtTrademark(16))
                     End If
                     If Val(strFee) > 0 Then
                        If Val(txtTrademark(18)) <> Val(strFee) Then
                           strTit = "檢核資料"
                           strMsg = "規費應為<" & strFee & ">"
                           nResponse = MsgBox(strMsg, vbOKCancel + vbCritical, strTit)
                           Screen.MousePointer = varSaveCursor
                           Exit Sub
                        End If
                     End If
                     If GetCaseMoney(txtSystem, txtTrademark(5), txtTrademark(1), Val(txtTrademark(18))) = 0 Then
                        Screen.MousePointer = varSaveCursor
                        Exit Sub
                     End If
                     'T及FCT第二期註冊費收文時, 若收文日大於法定期限時, 則控制規費加倍
                     If (Me.txtSystem.Text = "T" Or Me.txtSystem.Text = "FCT") And Me.txtTrademark(1).Text = "716" And (Me.txtTrademark(0).Text <> "" And Me.txtTrademark(16).Text <> "") And (Val(Me.txtTrademark(0).Text) > Val(Me.txtTrademark(16).Text)) Then
                        'Add By Sindy 2011/3/14 若收文日的昨天為非工作天, 則計算出收文日的前一個工作天
                        intMoney = 1
                        If ChkWorkDay(DBDATE(DateAdd("d", -1, ChangeWStringToWDateString((Val(Me.txtTrademark(0).Text) + 19110000))))) = False Then
                           If (Val(CompWorkDay(1, DBDATE(DateAdd("d", -1, ChangeWStringToWDateString((Val(Me.txtTrademark(0).Text) + 19110000)))), 1)) - 19110000) > Val(Me.txtTrademark(16).Text) Then
                              intMoney = 2
                           End If
                        Else
                           intMoney = 2
                        End If
                        If intMoney = 2 Then
                        '2011/3/14 End
                           strFee = Val(GetOfficalFee(Me.txtSystem.Text, Me.txtTrademark(5).Text, Me.txtTrademark(1).Text)) * intMoney
                           If Val(Me.txtTrademark(18).Text) <> Val(strFee) Then
                              MsgBox "規費應為<" & strFee & ">", vbExclamation + vbOKOnly
                              Screen.MousePointer = varSaveCursor
                              Exit Sub
                           End If
                        End If
                     End If
                  End If
               End If '2013/12/27 +if END
            End If 'ADD by sonia 2014/7/17 加入未輸規費時不檢查上面這段,因為可能是依代理人帳單請款CFP-027024
         End If
         
         'Add by Sindy 2015/4/1 收”延期”時一定要輸本所期限及法定期限
         If (txtSystem = "T" Or txtSystem = "CFT") And txtTrademark(1).Text = "303" Then
            If Val(txtTrademark(11).Text) = 0 Or Val(txtTrademark(16).Text) = 0 Then
               MsgBox "延期一定要有期限，請退回智權人員補填期限！", vbExclamation + vbOKOnly
               Screen.MousePointer = varSaveCursor
               Exit Sub
            End If
         End If
         '2015/4/1 END
         
         'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
         'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
         'modify by sonia 2014/9/11 取消X69514,已轉外專
         If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" And _
            Mid(txtTrademark(9), 1, 8) <> "X3928904" And Mid(txtTrademark(24), 1, 8) <> "X3928904" And Mid(txtTrademark(25), 1, 8) <> "X3928904" And Mid(txtTrademark(26), 1, 8) <> "X3928904" And Mid(txtTrademark(27), 1, 8) <> "X3928904" Then
               If Val(txtTrademark(14)) = 0 And Val(txtTrademark(18)) = 0 And Val(txtTrademark(15)) = 0 Then
                  If i = 14 Or i = 18 Or i = 15 Then
                     GoTo GoToNext
                  End If
               End If
         End If
         'Add By Sindy 2010/5/26 檢查申請人及移轉申請人的輸入順序
         If (Trim(txtTrademark(24)) <> "" And Trim(txtTrademark(9)) = "") Or _
            (Trim(txtTrademark(25)) <> "" And Trim(txtTrademark(24)) = "") Or _
            (Trim(txtTrademark(26)) <> "" And Trim(txtTrademark(25)) = "") Or _
            (Trim(txtTrademark(27)) <> "" And Trim(txtTrademark(26)) = "") Then
            ShowMsg "請依序輸入申請人!"
            If Trim(txtTrademark(27)) <> "" Then txtTrademark(27).SetFocus: Call txtTrademark_GotFocus(27): Exit For
            If Trim(txtTrademark(26)) <> "" Then txtTrademark(26).SetFocus: Call txtTrademark_GotFocus(26): Exit For
            If Trim(txtTrademark(25)) <> "" Then txtTrademark(25).SetFocus: Call txtTrademark_GotFocus(25): Exit For
            If Trim(txtTrademark(24)) <> "" Then txtTrademark(24).SetFocus: Call txtTrademark_GotFocus(24): Exit For
         End If
         If (Trim(txtTrademark(28)) <> "" And Trim(txtTrademark(20)) = "") Or _
            (Trim(txtTrademark(29)) <> "" And Trim(txtTrademark(28)) = "") Or _
            (Trim(txtTrademark(30)) <> "" And Trim(txtTrademark(29)) = "") Or _
            (Trim(txtTrademark(31)) <> "" And Trim(txtTrademark(30)) = "") Then
            ShowMsg "請依序輸入移轉申請人!"
            If Trim(txtTrademark(31)) <> "" Then txtTrademark(31).SetFocus: Call txtTrademark_GotFocus(31): Exit For
            If Trim(txtTrademark(30)) <> "" Then txtTrademark(30).SetFocus: Call txtTrademark_GotFocus(30): Exit For
            If Trim(txtTrademark(29)) <> "" Then txtTrademark(29).SetFocus: Call txtTrademark_GotFocus(29): Exit For
            If Trim(txtTrademark(28)) <> "" Then txtTrademark(28).SetFocus: Call txtTrademark_GotFocus(28): Exit For
         End If
         '2010/5/26 End
         'Modify By Cheng 2001/12/27
         If i = 9 Then
            If txtTrademark(9) = "" And txtTrademark(10) = "" Then
               ShowMsg "申請人或代理人不可同時空白!"
               txtTrademark(9).SetFocus
               txtTrademark_GotFocus (9)
               Exit For
           End If
         End If
         'modify by sonia 2017/1/23 +第二~五申請人
         'If i = 9 Then
         If i = 9 Or i = 24 Or i = 25 Or i = 26 Or i = 27 Then
            If Len(Trim(Me.txtTrademark(9).Text)) > 0 Then
               If CheckKeyIn(i) <> 1 Then
                  If txtTrademark(i).Enabled = True Then
                     txtTrademark(i).SetFocus
                     txtTrademark_GotFocus (i)
                  End If
                  Exit For
               End If
            End If
         ElseIf i = 10 Then
            If Len(Trim(Me.txtTrademark(10).Text)) > 0 Then
               If CheckKeyIn(i) <> 1 Then
                  If txtTrademark(i).Enabled = True Then
                     txtTrademark(i).SetFocus
                     txtTrademark_GotFocus (i)
                  End If
                  Exit For
               End If
            End If
         ElseIf txtTrademark(i).Enabled And txtTrademark(i).Visible Then
            If CheckKeyIn(i) <> 1 Then
               txtTrademark(i).SetFocus
               txtTrademark_GotFocus (i)
               Exit For
            End If
         End If
GoToNext:
      Next
      If i = 22 Then
         'Add By Cheng 2003/11/20
         '檢查規費欄位
         'edit by nick 2004/09/14
         'edit by nick 2004/09/15 台灣的才要
         'If (txtSystem = "T" Or txtSystem = "FCT") And (txtTrademark(1) = "602" Or txtTrademark(1) = "604" Or txtTrademark(1) = "606") Then
         If (txtSystem = "T" Or txtSystem = "FCT") And (txtTrademark(1) = "602" Or txtTrademark(1) = "604" Or txtTrademark(1) = "606") And txtTrademark(5).Text = "000" Then
             If Val(txtTrademark(18)) <> 0 Then
                 ShowMsg "異議答辯、裁定答辯、撤銷答辯..規費應該是 0!"
                 Me.txtTrademark(18).SetFocus
                 txtTrademark_GotFocus 18
                 Exit Sub
             End If
         'Added by Lydia 2020/05/20 法律所案源收文：台灣案之B1及B2案件性質都不可收費用。
         ElseIf txtTrademark(5) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
         'end 2020/05/20
         Else
            If Val(txtTrademark(18)) > 0 Or txtTrademark(5) = "000" Then     'ADD by sonia 2014/7/17 加入未輸規費時不檢查此段,因為可能是依代理人帳單請款CFP-027024
               'edit by nickc 2006/12/05 change call basquery
               'If objPublicData.GetCaseMoney(txtSystem, txtTrademark(5), txtTrademark(1), Val(txtTrademark(18))) <> 1 Then
               'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
               'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
               'modify by sonia 2014/9/11 取消X69514,已轉外專
               If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" And _
                  Mid(txtTrademark(9), 1, 8) <> "X3928904" And Mid(txtTrademark(24), 1, 8) <> "X3928904" And Mid(txtTrademark(25), 1, 8) <> "X3928904" And Mid(txtTrademark(26), 1, 8) <> "X3928904" And Mid(txtTrademark(27), 1, 8) <> "X3928904" Then
                   If GetCaseMoney(txtSystem, txtTrademark(5), txtTrademark(1), Val(txtTrademark(18))) <> 1 Then
                       Screen.MousePointer = varSaveCursor
                       If txtTrademark(18).Enabled = True Then 'Add by Morgan 2006/11/9 要檢查enabled否則會當
                         Me.txtTrademark(18).SetFocus
                         txtTrademark_GotFocus 18
                       End If
                       Exit Sub
                   End If
               End If
            End If 'ADD by sonia 2014/7/17 加入未輸規費時不檢查上面這段,因為可能是依代理人帳單請款CFP-027024
         End If
         'Add By Cheng 2003/08/28
         '檢查點數是否低於底價
         'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
         'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
         'modify by sonia 2014/9/11 取消X69514,已轉外專
         If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" And _
            Mid(txtTrademark(9), 1, 8) <> "X3928904" And Mid(txtTrademark(24), 1, 8) <> "X3928904" And Mid(txtTrademark(25), 1, 8) <> "X3928904" And Mid(txtTrademark(26), 1, 8) <> "X3928904" And Mid(txtTrademark(27), 1, 8) <> "X3928904" Then
             'Added by Lydia 2020/05/20 法律所案源收文：B類不檢查規費
             If txtTrademark(5) = "000" And m_LOS02 <> "" And Left(m_LOS02, 1) = "B" Then
             Else
             'end 2020/05/20
                'Add by Amy 2014/10/23 +T大陸分割母案不需控制
                If Not (txtCode(0) <> "" And txtTrademark(1).Text = "308" And txtTrademark(5).Text = "020") Then
                   If ChkPointValue(Me.txtSystem.Text, Me.txtTrademark(5).Text, Me.txtTrademark(1).Text, Me.txtTrademark(15).Text, Me.txtTrademark(12).Text) = False Then
                        Screen.MousePointer = varSaveCursor
                       txtTrademark(15).SetFocus
                       txtTrademark_GotFocus (15)
                       Exit Sub
                   End If
                End If
                'end 2014/10/23
             End If 'Added by Lydia 2020/05/20
         End If
         'add by nickc 2006/11/30 查名時，類別一定要輸
         If txtTrademark(4).Visible = True And txtTrademark(4).Enabled = True And txtTrademark(1) = "001" Then
             If Trim(txtTrademark(4)) = "" Then
                 Screen.MousePointer = varSaveCursor
                 MsgBox "查名必須要有商品類別！", , "注意！"
                 Me.txtTrademark(25).SetFocus
                 txtTrademark_GotFocus (25)
                 Exit Sub
             End If
         End If
         strAuto1 = txtRecieveCode
         'Add By Cheng 2002/05/23
         '重新檢查欄位有效性
         If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
         'add by nickc 2007/11/12 加入檢查特殊客戶
         Dim IsSpecCu As Boolean
         IsSpecCu = False
         If fraPatition.Visible = True Then
               If IsSpecCu = False And txtTrademark(20) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtTrademark(20)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtTrademark(20)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtTrademark(28) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtTrademark(28)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtTrademark(28)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtTrademark(29) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtTrademark(29)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtTrademark(29)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtTrademark(30) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtTrademark(30)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtTrademark(30)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtTrademark(31) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtTrademark(31)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtTrademark(31)), 9, 1) & "' "
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
               If IsSpecCu = False And txtTrademark(9) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtTrademark(9)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtTrademark(9)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtTrademark(24) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtTrademark(24)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtTrademark(24)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtTrademark(25) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtTrademark(25)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtTrademark(25)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtTrademark(26) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtTrademark(26)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtTrademark(26)), 9, 1) & "' "
                   CheckOC3
                   AdoRecordSet3.CursorLocation = adUseClient
                   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                   If AdoRecordSet3.RecordCount <> 0 Then
                       If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                           IsSpecCu = True
                       End If
                   End If
               End If
               If IsSpecCu = False And txtTrademark(27) <> "" Then
                   strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtTrademark(27)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtTrademark(27)), 9, 1) & "' "
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
               If MsgBox("請確認此客戶接洽單主管是否核示??", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                   Exit Sub
               End If
         End If
         
         'add by nickc 2007/03/27 非台灣要詢問
         '2009/11/25 MODIFY BY SONIA 新案才要詢問
         'If GetPrjNationNumber1(ChangeCustomerL(txtTrademark(9))) > "010" Then
         '2010/10/20 modify by sonia 非智權部收文才要問 CFP-023621
         'If GetPrjNationNumber1(ChangeCustomerL(txtTrademark(9))) > "010" And txtCode(0) = "" Then
         'Modified by Lydia 2019/02/14
         'If GetPrjNationNumber1(ChangeCustomerL(txtTrademark(9))) > "010" And txtCode(0) = "" And Left(Trim(GetST15(txtTrademark(12).Text)), 1) <> "S" Then
         If GetPrjNationNumber1(ChangeCustomerL(txtTrademark(9))) > "010" And txtCode(0) = "" And Left(m_SalesST15, 1) <> "S" Then
               If txtTrademark(10) = "" Then
                   If MsgBox("請確定  無代理人   !!", vbYesNo, "警告！") = vbNo Then
                       Screen.MousePointer = varSaveCursor
                       txtTrademark(10).SetFocus
                       txtTrademark_GotFocus (10)
                       Exit Sub
                   End If
               'Modify by Amy 2017/01/03 從下面搬上來,上面訊息若選擇"是",就不要再詢問下列訊息-秀玲
               ElseIf txtTrademark(33) = "" Then
                    If MsgBox("請確定  無代理人彼所案號  !!", vbYesNo, "警告！") = vbNo Then
                        Screen.MousePointer = varSaveCursor
                        txtTrademark(33).SetFocus
                        txtTrademark_GotFocus (33)
                        Exit Sub
                    End If
               End If
         End If
         
         'Add By Sindy 2011/01/05
         'Modified by Lydia 2019/04/08 改成變數
         'strSql = "select st15 from staff where st01='" & txtTrademark(12) & "'"
         'intI = 1
         'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         'If intI = 1 Then
         '   If Not IsNull(RsTemp.Fields("st15")) Then
               '國外部收文台灣案必須收FCT案號
               'If Left(Trim(RsTemp.Fields("st15")), 1) = "F" And txtTrademark(5) = "000" And txtSystem <> "FCT" Then
               If Left(m_SalesST15, 1) = "F" And txtTrademark(5) = "000" And txtSystem <> "FCT" Then
                  '2015/3/17 MODIFY BY SONIA 江律師要收T案
                  'MsgBox "國外部台灣案必須收 FCT 案號 !!!", vbExclamation + vbOKOnly
                  'Screen.MousePointer = varSaveCursor
                  'Exit Sub
                  If MsgBox("國外部台灣案應該收 FCT 案號, 是否修改系統類別？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                     Screen.MousePointer = varSaveCursor
                     Exit Sub
                  End If
               End If
         '   End If
         'End If
         'end 2019/04/08
         
         'Add By Sindy 2012/5/8
         '台灣商標Ｔ,FCT案若收文爭議案件性質時，若未填寫則不可存檔；FCT案爭議案件性質則可不輸入但要提醒
         If Frame21.Visible = True Then
            If txtSystem = "T" Then
               'Modified by Lydia 2018/12/10 延期303、放棄專用權206、暫緩審理310不必填「文件是否齊備」
               'If textEP06 = "" Then
               '   MsgBox "資料是否齊備不可空白!!!", vbExclamation + vbOKOnly
               If textEP06 = "" And textEP06.Visible = True And InStr(T案收文齊備排除, txtTrademark(1)) = 0 Then
                  MsgBox Left(Label41.Caption, 2) & "是否齊備不可空白!!!", vbExclamation + vbOKOnly
               'end 2018/12/10
                  Me.textEP06.SetFocus
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
               If textEP34 = "" And textEP34.Visible = True Then   'Modified by Lydia 2018/12/10 +判斷顯示
                  MsgBox "是否會稿不可空白!!!", vbExclamation + vbOKOnly
                  Me.textEP34.SetFocus
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
               If textCP122 = "" And textCP122.Visible = True Then  'Modified by Lydia 2018/12/10 +判斷顯示
                  MsgBox "是否急件不可空白!!!", vbExclamation + vbOKOnly
                  Me.textCP122.SetFocus
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
               'Added by Lydia 2018/12/10
               If textCP143 = "" And textCP143.Visible = True Then
                  MsgBox "查名齊備不可空白!!!", vbExclamation + vbOKOnly
                  Me.textCP143.SetFocus
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
               'end 2018/12/10
            Else
               strMsg = ""
               If textEP06 = "" Then strMsg = strMsg & "資料是否齊備、"
               If textEP34 = "" Then strMsg = strMsg & "是否會稿、"
               If textCP122 = "" Then strMsg = strMsg & "是否急件、"
               If strMsg <> "" Then
                  strMsg = Left(strMsg, Len(strMsg) - 1)
                  If MsgBox("未輸入" & strMsg & "，是否繼續收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                     If textEP06 = "" Then
                        Me.textEP06.SetFocus
                     ElseIf textEP34 = "" Then
                        Me.textEP34.SetFocus
                     ElseIf textCP122 = "" Then
                        Me.textCP122.SetFocus
                     End If
                     Screen.MousePointer = varSaveCursor
                     Exit Sub
                  End If
               End If
            End If
'            '案件性質為613補充答辯或612補充理由時，則只可不會稿
'            If txtTrademark(1) = "613" Or txtTrademark(1) = "612" Then
'               If textEP34.Text <> "N" Then
'                  MsgBox "案件性質為補充答辯或補充理由時，不需會稿!!!", vbExclamation + vbOKOnly
'                  Me.textEP34.SetFocus
'                  Screen.MousePointer = varSaveCursor
'                  Exit Sub
'               End If
'            End If
            '存檔時檢查若本所期限在7個日曆天內將到期或急件，要會稿且費用<8000元者，彈訊息讓使用者可選擇繼續收文。
            If ((Val(DBDATE(txtTrademark(11))) > 0 And _
                 Val(DBDATE(txtTrademark(11))) <= Val(CompDate(2, 7, strSrvDate(1)))) Or _
                textCP122 = "Y") And _
               textEP34 = "Y" And _
               Val(txtTrademark(14)) < 8000 Then
               'Modified by Lydia 2019/07/02 +可以延期案件
               'If MsgBox("本所期限在7天內或急件且收費低於8,000元且要會稿，此為特殊案件，請注意有主管核可才可收文！是否繼續收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               If MsgBox("本所期限在7天內或急件且收費低於8,000元且要會稿，此為特殊案件，請注意有主管核可或可以延期案件才可收文！是否繼續收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
            End If
         End If
         '2012/5/8 End
         
'2011/4/21 add by sonia
Dim strTM123 As String, strContact As String
   
         If cboContact.Locked = False Then
            strContact = ""
            If cboContact.ListCount > 2 Then
               'Modify by Amy 2021/12/20 改成Form 2.0
               'strTM123 = Format(cboContact.ItemData(cboContact.ListIndex), "00")
               strTM123 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
               PUB_GetContact strAppNo1, strContact, True
               If strTM123 = strContact Or strTM123 = "00" Then
                  If MsgBox("請確定接洽人欄是否有為★, 是否要選擇其他接洽人!!", vbYesNo, "警告！") = vbYes Then
                      Screen.MousePointer = varSaveCursor
                      cboContact.SetFocus
                      Exit Sub
                  End If
               End If
            End If
         End If
         '2011/4/21 end
         
         'Add By Sindy 2022/12/7 證書形式
         If txtTrademark(35).Visible = True Then
            If Len(txtTrademark(35)) = 0 And strSrvDate(1) >= "20230101" Then
               MsgBox "證書形式不可空白！", vbExclamation, "檢核資料"
               txtTrademark(35).SetFocus
               Screen.MousePointer = varSaveCursor
               Exit Sub
            End If
         End If
         
         'Added by Lydia 2019/05/13 改模組(一併取得)
         If txtTrademark(9).Text <> "" And Val(txtTrademark(14)) > 0 And Left(m_SalesST15, 1) <> "F" Then
             'Modified by Lydia 2022/06/13 傳入收文之本所案號,案件性質(可用,串接)
             'Call PUB_GetBillDataAll("3", txtTrademark(9), dblAmt, dblPFee, dblTFee, , , TransDate(txtTrademark(0), 2), mBillNo, mMemo)
             'Modified by Lydia 2022/06/15 傳入收文之智權人員
             If txtSystem.Text = 馬德里案 Then
                 Call PUB_GetBillDataAll("3", txtTrademark(9), txtSystem & IIf(txtTFCode(0) <> "", txtTFCode(0) & Left(txtTFCode(1) & "0", 1) & Left(txtTFCode(2) & "00", 2), ""), txtTrademark(1), Trim(txtTrademark(12)), dblAmt, dblPFee, dblTFee, , , TransDate(txtTrademark(0), 2), mBillNo, mMemo)
             Else
                 Call PUB_GetBillDataAll("3", txtTrademark(9), txtSystem & IIf(txtCode(0) <> "", txtCode(0) & Left(txtCode(1) & "0", 1) & Left(txtCode(2) & "00", 2), ""), txtTrademark(1), Trim(txtTrademark(12)), dblAmt, dblPFee, dblTFee, , , TransDate(txtTrademark(0), 2), mBillNo, mMemo)
             End If
             'end 2022/06/13
         End If
         
         'Add By Sindy 2012/11/06 非T*案件(TF要含)若已送件之應收款超過15萬以上,智權人員非國外部且有費用者須做下列控管
         'Modified by Lydia 2017/06/19 +判斷有申請人編號
         'If (Left(Trim(txtSystem), 1) <> "T" Or Trim(txtSystem) = "TF") And _
            Left(PUB_GetStaffST15(Trim(txtTrademark(12)), "1"), 1) <> "F" And _
            Val(txtTrademark(14)) > 0 And _
            Check2.Value = 0 Then
         'Modified by Lydia 2019/04/08 PUB_GetStaffST15(Trim(txtTrademark(12)), "1") => m_SalesST15
         If (Left(Trim(txtSystem), 1) <> "T" Or Trim(txtSystem) = "TF") And _
            Left(m_SalesST15, 1) <> "F" And _
            Val(txtTrademark(14)) > 0 And _
            Check2.Value = 0 And Trim(txtTrademark(9)) <> "" Then
         'end 2017/06/19
            'Mark by Lydia 2019/05/13 改模組(一併取得)
            'GetBillData txtTrademark(9), dblAmt, dblPFee, dblTFee
            
            'Add By Sindy 2012/12/10 取得客戶應收帳款收文檢查上限
            'Modified by Lydia 2020/02/03 應收帳款上限分開管制為個人"應收帳款上限"和"集團應收帳款上限"
            'dblChkAmt = PUB_GetCustRecAmtLmt(txtTrademark(9))
            ''2012/12/10 End
            dblCu183 = PUB_GetCustRecAmtLmt(txtTrademark(9), dblChkAmt)
            'Added by Lydia 2020/02/03 判斷是否有集團上限
            If dblChkAmt = 0 Then
                dblAmtR = 0: dblPFeeR = 0: dblTFeeR = 0
            Else   '有集團上限才抓關係企業的應收帳款金額
                GetBillData txtTrademark(9), dblAmtR, dblPFeeR, dblTFeeR
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
            'If InStr(cnt應收帳款檢查排除, Trim(txtTrademark(12))) = 0 And dblAmt >= dblCu183 Or (dblAmtR >= dblChkAmt And dblChkAmt > 0) Then
            If InStr(Pub_GetSpecMan("應收帳款上限檢查排除"), Trim(txtTrademark(12))) = 0 And dblAmt >= dblCu183 Or (dblAmtR >= dblChkAmt And dblChkAmt > 0) Then
               'Modified by Lydia 2018/09/20 預設按鈕改成"否" vbDefaultButton1=>vbDefaultButton2
               If MsgBox("請注意接洽單上是否有註明應收帳款超額，需主管簽核才可收文！是否可收文？" & vbCrLf & _
                         "（接洽單上若有★★的應收帳款簽核控管，是否已勾選畫面上的註記欄位了？）", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
'            '已送件之應收款超過15萬以上(不含T*案件應收款),提醒
'            ElseIf dblAmt >= 150000 Then
'               If MsgBox("請注意接洽單上是否有註明應收帳款超額，需主管簽核才可收文！是否可收文？" & vbCrLf & _
'                         "（接洽單上若有★★的應收帳款簽核控管，是否已勾選畫面上的註記欄位了？）", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
'                  Screen.MousePointer = varSaveCursor
'                  Exit Sub
'               End If
            End If
         End If
         
         'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
         'Modified by Lydia 2019/04/08 智權人員非國外部
         'If txtTrademark(9).Text <> "" And Val(txtTrademark(14)) > 0 Then
         If txtTrademark(9).Text <> "" And Val(txtTrademark(14)) > 0 And Left(m_SalesST15, 1) <> "F" Then
            'Modified by Lydia 2019/05/13 改模組(一併取得)
            'If GetBillDate(txtTrademark(9), TransDate(txtTrademark(0), 2), strExc(1), strExc(2)) = True Then
            If mMemo <> "" Then
                'Modified by Lydia 2018/10/29 改訊息
                'If MsgBox("請注意接洽單上是否有註明" & vbCrLf & strExc(2) & vbCrLf & "，請交主管簽核並且有主管簽核。" & vbCrLf & "是否可收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                If MsgBox("請注意接洽單上是否有註明" & vbCrLf & mMemo & "，請交主管簽核。" & vbCrLf & "並且有主管簽核，是否可收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
                End If
            End If
         End If
         'end 2018/08/22
         
         '2012/11/06 End
         'Added by Lydia 2015/11/12 提示輸入收文組群
         'Modified by Lydia 2016/03/28
         'Modified by Lydia 2016/04/18 收文組群更名為查名代號
         'Modified by Lydia 2016/04/27 改成直接在畫面輸入查名代號
         'If cmdTSMap.Visible = True And TMQList = "" And bolOpen130 = False Then
         'Modified by Lydia 2016/05/09 +台灣案
         If txtTrademark(5) = "000" And lblTS.Visible = True And txtTS(0).Visible = True And txtTS(1).Enabled = True Then
            'Modified by Lydia 2016/03/09
            'If MsgBox("是否要輸入收文組群?", vbInformation + vbYesNo, "輸入收文組群") = vbYes Then
            '   Screen.MousePointer = varSaveCursor
               'Call cmdTSMap_Click
            If Len(txtTS(0) & txtTS(1)) <> 9 Then
                If MsgBox("查名代號應輸入?", vbInformation + vbYesNo, "輸入查名代號") = vbYes Then
                   Screen.MousePointer = varSaveCursor
                   txtTS(1).SetFocus
                   Exit Sub
                End If
            Else
                'Modified dy Lydia 2019/08/23 +智權人員的部門
                'If PUB_TQCtoTMQ(txtTrademark(12).Text, txtTS(0).Text & txtTS(1).Text, TMQList) = False Then
                'Modified by Lydia 2024/03/14 +True
                If PUB_TQCtoTMQ(True, m_SalesST15, txtTrademark(12).Text, txtTS(0).Text & txtTS(1).Text, TMQList) = False Then
                    Screen.MousePointer = varSaveCursor
                    txtTS(1).SetFocus
                    Exit Sub
                End If
                TMQList = txtTS(0).Text & txtTS(1).Text & "|" & TMQList
            End If
            'end 2016/04/27
         End If
         'end 2015/11/12
         'Modified by Lydia 2022/09/05 判斷啟用日
         'If SaveDatabase(strAuto1, strAuto2) Then
         bolSaveOK = False
         If strSrvDate(1) < 收文存檔模組化啟用日 Then
             bolSaveOK = SaveDatabase(strAuto1, strAuto2)
         Else
             If txtSystem.Text = 馬德里案 Then
                Call SetDBArray(False, txtRecieveCode, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)))
             Else
                Call SetDBArray(False, txtRecieveCode, txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
             End If
             bolSaveOK = PUB_SaveFrm010004(Me.Name, frm010001.intSaveMode, frm010001.intModifyKind, frm010001.intChoose, modBase, modCP, txtTrademark(13), m_strControl, IsSaveData, mType, mCaseNo, mRetVal)
                 
             If frm010001.intModifyKind = 0 And bolSaveOK = True Then
                If txtSystem = "TF" Then
                   txtTFCode(0) = Mid(modBase(2), 1, 5)
                   txtTFCode(1) = Mid(modBase(2), 6, 1)
                Else
                   txtCode(0) = modBase(2)
                End If
                strAuto1 = modCP(9)
                strAuto2 = modBase(2)
             End If
'             'Add By Sindy 2023/5/31
'             If bolSaveOK = True Then
'                 '外專信件沖銷: 收完文
'                 If InStr("," & mRetVal, "m_bolRecvOK = True") > 0 Then
'                    m_bolRecvOK = True
'                 End If
'                 '多案收文的總收文號要傳入第一筆總收文號
'                 If InStr("," & mRetVal, "MCR11:") > 0 Then
'                     m_strMCR11 = Mid(mRetVal, InStr(mRetVal, "MCR11:") + 6, 9)
'                 End If
'             End If
'             '2023/5/31 END
         End If
'-----------------------------------------------
         If bolSaveOK = True Then
         'end 2022/09/05
            'Add By Sindy 2009/06/12
            If Trim(txtSystem.Text) = "T" And txtTrademark(5).Text = "000" And txtTrademark(1).Text = "403" And Val(txtTrademark(14).Text) > 20000 Then
               MsgBox "請 T 國內行政訴訟案，請同時收文 204 準備程序及 205 言詞辯論！", , "請注意！"
            End If
            '2009/06/12 End
            'Added by Lydia 2015/11/12 查名單對應存檔
            'Modified by Lydia 2022/09/05 判斷啟用日
            If TMQList <> "" And strSrvDate(1) < 收文存檔模組化啟用日 Then
               strExc(1) = Mid(TMQList, 1, InStr(TMQList, "|") - 1)
               strExc(2) = Mid(TMQList, InStr(TMQList, "|") + 1)
               'Added by Lydia 2018/12/10 查名是否已齊備
               'Mark by Lydia 2019/01/30 收文依照列印的接洽單輸入,不檢查查名是否已齊備
               'If Trim(txtSystem.Text) = "T" And txtTrademark(5).Text = "000" And txtTrademark(1).Text = 申請 And DBDATE(txtTrademark(0).Text) >= T案收文齊備啟用日 Then
               '     strExc(0) = PUB_TMQchkCP143(Replace(Replace(strExc(2), "+", ""), "-", "")) '傳入每一張查名單號,非接洽單印的查名代號(流水號)
               '     If textCP143.Text <> strExc(0) And strExc(0) <> "" Then
               '         MsgBox "查名" & IIf(strExc(0) = "Y", "已齊備", "未齊備") & "！", vbInformation
               '         textCP143.Text = strExc(0)
               '     End If
               'End If
               'end 2018/12/10
               'end 2019/01/30
               'Memo by Lydia 2022/09/05 因為最初查名單資料是在Table存放檔案，所以放在外面；現在直接併入存檔模組
               'Modified by Lydia 2024/03/14 +False
               'If PUB_TMQtoCP(m_AttachPath, strAuto1, strExc(2), strExc(1)) = False Then
               If PUB_TMQtoCP(False, m_AttachPath, strAuto1, strExc(2), strExc(1)) = False Then
               End If
            End If
            
            PUB_SendMailCache 'Add by Sindy 2022/9/29
            
            frm010001.ClearForm strAuto1, strAuto2
            bolLeave = True
            intLeaveKind = 1
            If frm010001.intModifyKind = 0 Then LastDate = txtTrademark(0).Text
            
            'Modify By Sindy 2023/5/30 信件內部收文執行完畢後,關閉視窗
            If m_strIR01 <> "" Then
               If Not m_PrevFormIR Is Nothing Then
                  Call m_PrevFormIR.GoNext
               End If
               Unload Me
               Unload frm010001
            Else
            '2023/5/30 END
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
End Sub

Private Function SaveDatabase(ByRef strRecieveAuto As String, ByRef strCaseAuto As String) As Boolean
Dim adoquery As New ADODB.Recordset
Dim strTM123 As String, strContact As String

   'Add by Morgan 2008/8/5
   If cboContact.Locked = False Then
      If cboContact.ListIndex >= 0 Then
         'Modify by Amy 2021/12/20 改成Form 2.0
         'If Val(cboContact.ItemData(cboContact.ListIndex)) > 0 Then
         '   strTM123 = Format(cboContact.ItemData(cboContact.ListIndex), "00")
         strTM123 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
         If Val(strTM123) > 0 Then
            'Add by Morgan 2008/8/7 若個案接洽人與客戶檔的預設接洽人相同時不必設定
            PUB_GetContact strAppNo1, strContact, True
            If strTM123 = strContact Then
               strTM123 = ""
            End If
         'Added by Lydia 2022/09/16 排除空白=00
         ElseIf strTM123 = "00" And Trim(cboContact.Text) = "" Then
             strTM123 = ""
         'end 2022/09/16
         End If
      End If
   Else
      strTM123 = "TM123"
   End If
   
   If frm010001.intModifyKind = 0 Then
      If strTM123 = "TM123" Then strTM123 = "" 'Add by Morgan 2008/8/7
      If txtSystem.Text = 馬德里案 Then
         SaveDatabase = InsertTrademarkDatabase(frm010001.intSaveMode, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
                  IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), txtTrademark(6), _
                  txtTrademark(3), txtTrademark(4), txtTrademark(5), txtTrademark(9), txtTrademark(10), txtTrademark(0), txtTrademark(11), txtTrademark(16), _
                  txtTrademark(1), txtTrademark(2), txtTrademark(12), txtTrademark(14), txtTrademark(18), txtTrademark(15), txtTrademark(19), _
                  txtTrademark(17), txtTrademark(20), txtTrademark(13), txtTrademark(21), strRecieveAuto, strCaseAuto, douStPrice, douLowPrice, txtCP64, txtTrademark(7), txtTrademark(24), txtTrademark(25), txtTrademark(26), txtTrademark(27), txtTrademark(28), txtTrademark(29), txtTrademark(30), txtTrademark(31), txtTrademark(32), txtTrademark(33), strTM123)
      Else
         SaveDatabase = InsertTrademarkDatabase(frm010001.intSaveMode, txtSystem, txtCode(0), _
                  IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtTrademark(6), _
                  txtTrademark(3), txtTrademark(4), txtTrademark(5), txtTrademark(9), txtTrademark(10), txtTrademark(0), txtTrademark(11), txtTrademark(16), _
                  txtTrademark(1), txtTrademark(2), txtTrademark(12), txtTrademark(14), txtTrademark(18), txtTrademark(15), txtTrademark(19), _
                  txtTrademark(17), txtTrademark(20), txtTrademark(13), txtTrademark(21), strRecieveAuto, strCaseAuto, douStPrice, douLowPrice, txtCP64, txtTrademark(7), txtTrademark(24), txtTrademark(25), txtTrademark(26), txtTrademark(27), txtTrademark(28), txtTrademark(29), txtTrademark(30), txtTrademark(31), txtTrademark(32), txtTrademark(33), strTM123)
      End If
   Else
      If txtSystem.Text = 馬德里案 Then
         SaveDatabase = UpdateTrademarkDatabase(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
                  IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), txtTrademark(6), _
                  txtTrademark(3), txtTrademark(4), txtTrademark(5), txtTrademark(9), txtTrademark(10), txtRecieveCode, txtTrademark(0), txtTrademark(11), txtTrademark(16), _
                  txtTrademark(1), txtTrademark(2), txtTrademark(12), txtTrademark(21), txtTrademark(14), txtTrademark(18), txtTrademark(15), txtTrademark(19), _
                  txtTrademark(17), txtTrademark(20), txtTrademark(13), douStPrice, douLowPrice, txtCP64, txtTrademark(24), txtTrademark(25), txtTrademark(26), txtTrademark(27), txtTrademark(28), txtTrademark(29), txtTrademark(30), txtTrademark(31), txtTrademark(32), txtTrademark(33), strTM123)
      Else
         SaveDatabase = UpdateTrademarkDatabase(txtSystem, txtCode(0), _
                  IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtTrademark(6), _
                  txtTrademark(3), txtTrademark(4), txtTrademark(5), txtTrademark(9), txtTrademark(10), txtRecieveCode, txtTrademark(0), txtTrademark(11), txtTrademark(16), _
                  txtTrademark(1), txtTrademark(2), txtTrademark(12), txtTrademark(21), txtTrademark(14), txtTrademark(18), txtTrademark(15), txtTrademark(19), _
                  txtTrademark(17), txtTrademark(20), txtTrademark(13), douStPrice, douLowPrice, txtCP64, txtTrademark(24), txtTrademark(25), txtTrademark(26), txtTrademark(27), txtTrademark(28), txtTrademark(29), txtTrademark(30), txtTrademark(31), txtTrademark(32), txtTrademark(33), strTM123)
      End If
   End If
    
   'add by nickc 2007/11/09 測試解決mail 發不到的時候會存兩筆的錯誤
   On Error GoTo 0    '歸零
   
   'Add By Sindy 2020/2/7
   '外商收文時, 案件有 FC代理人, 且代理人國籍為日本時,
   '不論是商標案件或服務業務案件, 存檔時定稿語文欄若為空值時,
   '一律更新為3.日文
   If txtTrademark(10) <> "" And (txtSystem = "FCT" Or txtSystem = "T" Or txtSystem = "S") Then
      strExc(10) = ChangeCustomerL(txtTrademark(10))
      strSql = "SELECT fa01,fa02,fa10 FROM fagent" & _
               " WHERE fa01=" & CNULL(Left(strExc(10), 8)) & _
               " and fa02=" & CNULL(Mid(strExc(10), 9, 1))
      adoquery.CursorLocation = adUseClient
      adoquery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount > 0 Then
         If Left("" & adoquery.Fields("fa10"), 3) = "011" Then '日本
            strSql = "UPDATE TradeMark SET TM53='3'" & _
                     " WHERE TM01 = '" & txtSystem & "' AND " & _
                        "TM02 = '" & txtCode(0) & "' AND " & _
                        "TM03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' AND " & _
                        "TM04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' AND " & _
                        "TM53 is null"
            cnnConnection.Execute strSql
         End If
      End If
      adoquery.Close
   End If
   '2020/2/7 END
   
   'Added by Lydia 2021/01/08 CFT英國脫歐案：一併複製商標圖、指定商品或服務名稱(InsertTrademarkDatabase)
   'Memo by Lydia 2021/03/05 含CFT歐盟尚未註冊案轉換英國申請案收文控管
   If txtSystem = "CFT" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" Then
      strExc(9) = ""
      If GetImgByteFile_Case(m_CaseNa239(1), m_CaseNa239(2), m_CaseNa239(3), m_CaseNa239(4), strExc(9), 0, strExc(5), strExc(6)) = True Then
          Call SaveImgByteFile(strExc(9), txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strExc(5), strExc(6))
      End If
   End If
   'end 2021/01/08
   
   'Added by Lydia 2021/02/01 CFT緬甸重新申請案：緬甸商標重新申請收文時，一併將舊案之「商標圖樣」、「商品/服務類別及名稱」、「優先權資料」帶入新案號
   If txtSystem = "CFT" And txtTrademark(5) = "048" And txtCFTNa048 <> "" Then
      strExc(9) = ""
      strSql = txtCFTNa048.Text
      strExc(1) = "": strExc(2) = "": strExc(3) = "": strExc(4) = ""
      Call ChgCaseNo(strSql, strExc)
      If GetImgByteFile_Case(strExc(1), strExc(2), strExc(3), strExc(4), strExc(9), 0, strExc(5), strExc(6)) = True Then
          Call SaveImgByteFile(strExc(9), txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strExc(5), strExc(6))
      End If
   End If
   'end 2021/02/01
   
   'Modified by Lydia 2019/09/16 +st06
   'm_SalesST15 = GetST15(txtTrademark(12)) 'Added by Lydia 2019/02/14
   m_SalesST15 = GetST15(txtTrademark(12), , , m_SalesST06)
   
'add by nickc 2005/09/05
If frm010001.intModifyKind = 0 Then
   Dim oContext As String, strCaseNo As String
   Dim strTemp As String
   'Add By Sindy 2021/2/1 不得代理的後續舊案收文控管，通知收文人員（CP13）
   If InStr(strYState, "不得代理") > 0 Then
      oContext = oContext & vbCrLf + "代理人： " + ChangeCustomerL(txtTrademark(10).Text) + " " + lblAgent.Caption + vbCrLf
      strTemp = strTemp & "," & ChangeCustomerL(txtTrademark(10).Text)
   End If
   If InStr(strXState(9), "不得代理") > 0 Then
      oContext = oContext & vbCrLf + "申請人1： " + ChangeCustomerL(txtTrademark(9).Text) + " " + lblPetition(0).Caption + vbCrLf
      strTemp = strTemp & "," & ChangeCustomerL(txtTrademark(9).Text)
   End If
   If InStr(strXState(24), "不得代理") > 0 Then
      oContext = oContext & vbCrLf + "申請人2： " + ChangeCustomerL(txtTrademark(24).Text) + " " + lblPetition(1).Caption + vbCrLf
      strTemp = strTemp & "," & ChangeCustomerL(txtTrademark(24).Text)
   End If
   If InStr(strXState(25), "不得代理") > 0 Then
      oContext = oContext & vbCrLf + "申請人3： " + ChangeCustomerL(txtTrademark(25).Text) + " " + lblPetition(2).Caption + vbCrLf
      strTemp = strTemp & "," & ChangeCustomerL(txtTrademark(25).Text)
   End If
   If InStr(strXState(26), "不得代理") > 0 Then
      oContext = oContext & vbCrLf + "申請人4： " + ChangeCustomerL(txtTrademark(26).Text) + " " + lblPetition(3).Caption + vbCrLf
      strTemp = strTemp & "," & ChangeCustomerL(txtTrademark(26).Text)
   End If
   If InStr(strXState(27), "不得代理") > 0 Then
      oContext = oContext & vbCrLf + "申請人5： " + ChangeCustomerL(txtTrademark(27).Text) + " " + lblPetition(4).Caption + vbCrLf
      strTemp = strTemp & "," & ChangeCustomerL(txtTrademark(27).Text)
   End If
   If InStr(strXState(20), "不得代理") > 0 Then
      oContext = oContext & vbCrLf + "移轉申請人1： " + ChangeCustomerL(txtTrademark(20).Text) + " " + lblPetitionName(0).Caption + vbCrLf
      strTemp = strTemp & "," & ChangeCustomerL(txtTrademark(20).Text)
   End If
   If InStr(strXState(28), "不得代理") > 0 Then
      oContext = oContext & vbCrLf + "移轉申請人2： " + ChangeCustomerL(txtTrademark(28).Text) + " " + lblPetitionName(1).Caption + vbCrLf
      strTemp = strTemp & "," & ChangeCustomerL(txtTrademark(28).Text)
   End If
   If InStr(strXState(29), "不得代理") > 0 Then
      oContext = oContext & vbCrLf + "移轉申請人3： " + ChangeCustomerL(txtTrademark(29).Text) + " " + lblPetitionName(2).Caption + vbCrLf
      strTemp = strTemp & "," & ChangeCustomerL(txtTrademark(29).Text)
   End If
   If InStr(strXState(30), "不得代理") > 0 Then
      oContext = oContext & vbCrLf + "移轉申請人4： " + ChangeCustomerL(txtTrademark(30).Text) + " " + lblPetitionName(3).Caption + vbCrLf
      strTemp = strTemp & "," & ChangeCustomerL(txtTrademark(30).Text)
   End If
   If InStr(strXState(31), "不得代理") > 0 Then
      oContext = oContext & vbCrLf + "移轉申請人5： " + ChangeCustomerL(txtTrademark(31).Text) + " " + lblPetitionName(4).Caption + vbCrLf
      strTemp = strTemp & "," & ChangeCustomerL(txtTrademark(31).Text)
   End If
   If oContext <> "" Then
      strTemp = Mid(strTemp, 2)
      If txtSystem.Text = 馬德里案 Then
         strCaseNo = txtSystem + "-" + txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)) + "-" + IIf(txtTFCode(2) = "", "0", txtTFCode(2)) + "-" + IIf(txtTFCode(3) = "", "00", txtTFCode(3))
         oContext = "本所案號： " + strCaseNo + vbCrLf + _
                    "案件名稱： " + txtTrademark(6) + vbCrLf + _
                    "申請國家： " + txtTrademark(5) + " " + lblNation.Caption + vbCrLf + _
                    "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + _
                    "案件性質： " + lblCaseProperty.Caption + vbCrLf + vbCrLf + _
                    "【不得代理】" + vbCrLf + _
                    oContext
      Else
         strCaseNo = IIf("-" + txtCode(1) + "-" + txtCode(2) = "-0-00", txtSystem + "-" + txtCode(0), txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2))
         oContext = "本所案號： " + strCaseNo + vbCrLf + _
                    "案件名稱： " + txtTrademark(6) + vbCrLf + _
                    "申請國家： " + txtTrademark(5) + " " + lblNation.Caption + vbCrLf + _
                    "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + _
                    "案件性質： " + lblCaseProperty.Caption + vbCrLf + vbCrLf + _
                    "【不得代理】" + vbCrLf + _
                    oContext
      End If
      oMailCount = Trim(txtTrademark(12).Text) & ";" & PUB_GetFCPProSup(Trim(txtTrademark(12).Text))
      PUB_SendMail strUserNum, oMailCount, "", strCaseNo & _
         " 已確認續行收文，請注意該" & strTemp & "編號已設為不得代理。", oContext
   End If
   '2021/2/1 END
   
      'add by nick 2004/10/15  當收文業務區與客戶檔業務區不同時發 mail  及提示
'edit by nickc 2006/11/22 加入申請人
'      Dim oStrCuSales As String
'      oStrCuSales = ""
'      '2006/8/2 MODIFY BY SONIA TXTSYSTEM只判斷1碼,因為FG
'      'If UCase(Mid(txtSystem, 1, 2)) <> "FC" And GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(9).Text), oStrCuSales) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(9).Text) <> "" Then
'      If UCase(Mid(txtSystem, 1, 1)) <> "F" And GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(9).Text), oStrCuSales) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(9).Text) <> "" Then
'          'edit by nickc 2005/08/10
'          'MsgBox "收文智權人員與客戶智權人員不同業務區，準備發 mail ，請定時刪除郵件備份！", , "注意！"
'          MsgBox "收文智權人員與客戶智權人員不同業務區，準備發 mail ！", , "注意！"
'          'edit by nickc 2005/08/10 加發秀玲
'          'PUB_SendMail strUserNum, Trim(txtTrademark(12).Text) & ";" & oStrCuSales, "", "案件收文通知--此案收文非原智權人員(區)！", vbCrLf + "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtTrademark(6) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf + "申請人： " + GetCustomerName(ChangeCustomerL(txtTrademark(9).Text)) + "原智權人員： " + oStrCuSales + vbCrLf + "收文智權人員： " + lblSales.Caption + vbCrLf + vbCrLf + "智權人員(區)不同！", ""
'          'edit by nickc 2005/08/16
'          'PUB_SendMail strUserNum, Trim(txtTrademark(12).Text) & ";" & oStrCuSales & ";83002", "", "案件收文通知--此案收文非原智權人員(區)！", vbCrLf + "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtTrademark(6) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf + "申請人： " + GetCustomerName(ChangeCustomerL(txtTrademark(9).Text)) + "原智權人員： " + oStrCuSales + vbCrLf + "收文智權人員： " + lblSales.Caption + vbCrLf + vbCrLf + "智權人員(區)不同！", ""
'          'add by nickc 2006/10/25
'          If txtSystem.Text = 馬德里案 Then
'            PUB_SendMail strUserNum, Trim(txtTrademark(12).Text) & ";" & oStrCuSales & ";83002", "", "案件收文通知--此案收文非原智權人員(區)！", vbCrLf + "本所案號： " + txtSystem + "-" + txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)) + "-" + IIf(txtTFCode(2) = "", "0", txtTFCode(2)) + "-" + IIf(txtTFCode(3) = "", "00", txtTFCode(3)) + vbCrLf + "案件名稱： " + txtTrademark(6) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf + "申請人： " + GetCustomerName(ChangeCustomerL(txtTrademark(9).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales) + vbCrLf + "收文智權人員： " + lblSales.Caption + vbCrLf + vbCrLf + "智權人員(區)不同！", ""
'          Else
'            PUB_SendMail strUserNum, Trim(txtTrademark(12).Text) & ";" & oStrCuSales & ";83002", "", "案件收文通知--此案收文非原智權人員(區)！", vbCrLf + "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtTrademark(6) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf + "申請人： " + GetCustomerName(ChangeCustomerL(txtTrademark(9).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales) + vbCrLf + "收文智權人員： " + lblSales.Caption + vbCrLf + vbCrLf + "智權人員(區)不同！", ""
'          End If
'      End If

Dim oStrCuSales1 As String
Dim oStrCuSales2 As String
Dim oStrCuSales3 As String
Dim oStrCuSales4 As String
Dim oStrCuSales5 As String
'Dim oContext As String
'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
Dim IsMail As Boolean

   IsMail = True
   'add by nickc 2007/05/16 加入若是本所期限小於等於當天，要發mail  通知
   Dim oContext2 As String
   oContext = "": oContext2 = ""
   
   oStrCuSales1 = ""
   oStrCuSales2 = ""
   oStrCuSales3 = ""
   oStrCuSales4 = ""
   oStrCuSales5 = ""
   If txtSystem.Text = 馬德里案 Then
     oContext = "本所案號： " + txtSystem + "-" + txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)) + "-" + IIf(txtTFCode(2) = "", "0", txtTFCode(2)) + "-" + IIf(txtTFCode(3) = "", "00", txtTFCode(3)) + vbCrLf + "案件名稱： " + txtTrademark(6) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
     'add by nickc 2007/05/16 加入若是本所期限小於等於當天，要發mail  通知
     'edit by nickc 2008/04/23 加入國家
     'oContext2 = "本所案號： " + txtSystem + "-" + txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)) + "-" + IIf(txtTFCode(2) = "", "0", txtTFCode(2)) + "-" + IIf(txtTFCode(3) = "", "00", txtTFCode(3)) + vbCrLf + "案件名稱： " + txtTrademark(6) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
     oContext2 = "本所案號： " + txtSystem + "-" + txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)) + "-" + IIf(txtTFCode(2) = "", "0", txtTFCode(2)) + "-" + IIf(txtTFCode(3) = "", "00", txtTFCode(3)) + vbCrLf + "案件名稱： " + txtTrademark(6) + vbCrLf + "申請國家：" + txtTrademark(5) + " " + lblNation.Caption + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
   Else
     oContext = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtTrademark(6) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
     'add by nickc 2007/05/16 加入若是本所期限小於等於當天，要發mail  通知
     'edit by nickc 2008/04/23 加入國家
     'oContext2 = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtTrademark(6) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
     oContext2 = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtTrademark(6) + vbCrLf + "申請國家：" + txtTrademark(5) + " " + lblNation.Caption + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtTrademark(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
   End If
   oMailCount = ""
   
   'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
   'If GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(9).Text), oStrCuSales1) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(9).Text) <> "" Then
   'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
   If ChkSameCuArea(Trim(txtTrademark(9)), Trim(txtTrademark(12)), , , , , Trim(txtTrademark(10))) = False And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(9).Text) <> "" Then
      'Add By Sindy 2009/10/19
      'Modified by Lydia 2019/02/14
      'If Left(Trim(GetST15(txtTrademark(12).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(9).Text), oStrCuSales1)), 1) = "F" Then
      If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(9).Text), oStrCuSales1)), 1) = "F" Then
         '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
      Else
         oMailCount = oMailCount & oStrCuSales1 & ";"
         oContext = oContext & vbCrLf + "申請人1： " + GetCustomerName(ChangeCustomerL(txtTrademark(9).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales1)
      End If
   'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
   Else
        If Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(9).Text) <> "" Then
            IsMail = False
        End If
   End If
   'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
   If m_SalesST06 <> "" And Trim(txtTrademark(9)) <> "" And Trim(txtTrademark(12).Text) <> "" Then
       If PUB_ChkOldCustomer(True, txtTrademark(9), Trim(txtTrademark(12)), m_SalesST15, m_SalesST06) = True Then
           IsMail = False
       End If
   End If
   
   'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
   'If GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(24).Text), oStrCuSales2) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(24).Text) <> "" Then
   'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
   If ChkSameCuArea(Trim(txtTrademark(24)), Trim(txtTrademark(12)), , , , , Trim(txtTrademark(10))) = False And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(24).Text) <> "" Then
      'Add By Sindy 2009/10/19
      'Modified by Lydia 2019/02/14
      'If Left(Trim(GetST15(txtTrademark(12).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(24).Text), oStrCuSales2)), 1) = "F" Then
      If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(24).Text), oStrCuSales2)), 1) = "F" Then
         '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
      Else
         oMailCount = oMailCount & oStrCuSales2 & ";"
         oContext = oContext & vbCrLf + "申請人2： " + GetCustomerName(ChangeCustomerL(txtTrademark(24).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales2)
      End If
   'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
   Else
        If Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(24).Text) <> "" Then
            IsMail = False
        End If
   End If
   'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
   If m_SalesST06 <> "" And Trim(txtTrademark(24)) <> "" And Trim(txtTrademark(12).Text) <> "" Then
       If PUB_ChkOldCustomer(True, txtTrademark(24), Trim(txtTrademark(12)), m_SalesST15, m_SalesST06) = True Then
           IsMail = False
       End If
   End If
   
   'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
   'If GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(25).Text), oStrCuSales3) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(25).Text) <> "" Then
   'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
   If ChkSameCuArea(Trim(txtTrademark(25)), Trim(txtTrademark(12)), , , , , Trim(txtTrademark(10))) = False And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(25).Text) <> "" Then
      'Add By Sindy 2009/10/19
      'Modified by Lydia 2019/02/14
      'If Left(Trim(GetST15(txtTrademark(12).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(25).Text), oStrCuSales3)), 1) = "F" Then
      If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(25).Text), oStrCuSales3)), 1) = "F" Then
         '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
      Else
         oMailCount = oMailCount & oStrCuSales3 & ";"
         oContext = oContext & vbCrLf + "申請人3： " + GetCustomerName(ChangeCustomerL(txtTrademark(25).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales3)
      End If
   'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
   Else
        If Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(25).Text) <> "" Then
            IsMail = False
        End If
   End If
   'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
   If m_SalesST06 <> "" And Trim(txtTrademark(25)) <> "" And Trim(txtTrademark(12).Text) <> "" Then
       If PUB_ChkOldCustomer(True, txtTrademark(25), Trim(txtTrademark(12)), m_SalesST15, m_SalesST06) = True Then
           IsMail = False
       End If
   End If
  
   'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
   'If GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(26).Text), oStrCuSales4) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(26).Text) <> "" Then
   'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
   If ChkSameCuArea(Trim(txtTrademark(26)), Trim(txtTrademark(12)), , , , , Trim(txtTrademark(10))) = False And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(26).Text) <> "" Then
      'Add By Sindy 2009/10/19
      'Modified by Lydia 2019/02/14
      'If Left(Trim(GetST15(txtTrademark(12).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(26).Text), oStrCuSales4)), 1) = "F" Then
      If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(26).Text), oStrCuSales4)), 1) = "F" Then
         '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
      Else
         oMailCount = oMailCount & oStrCuSales4 & ";"
         oContext = oContext & vbCrLf + "申請人4： " + GetCustomerName(ChangeCustomerL(txtTrademark(26).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales4)
      End If
   'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
   Else
        If Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(26).Text) <> "" Then
            IsMail = False
        End If
   End If
   'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
   If m_SalesST06 <> "" And Trim(txtTrademark(26)) <> "" And Trim(txtTrademark(12).Text) <> "" Then
       If PUB_ChkOldCustomer(True, txtTrademark(26), Trim(txtTrademark(12)), m_SalesST15, m_SalesST06) = True Then
           IsMail = False
       End If
   End If
   
   'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
   'If GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(27).Text), oStrCuSales5) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(27).Text) <> "" Then
   'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
   If ChkSameCuArea(Trim(txtTrademark(27)), Trim(txtTrademark(12)), , , , , Trim(txtTrademark(10))) = False And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(27).Text) <> "" Then
      'Add By Sindy 2009/10/19
      'Modified by Lydia 2019/02/14
      'If Left(Trim(GetST15(txtTrademark(12).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(27).Text), oStrCuSales5)), 1) = "F" Then
      If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(27).Text), oStrCuSales5)), 1) = "F" Then
         '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
      Else
         oMailCount = oMailCount & oStrCuSales5 & ";"
         oContext = oContext & vbCrLf + "申請人5： " + GetCustomerName(ChangeCustomerL(txtTrademark(27).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales5)
      End If
   'add by nickc 2007/05/08 秀玲說，其中一個符合就不發了
   Else
        If Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(27).Text) <> "" Then
            IsMail = False
        End If
   End If
   'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
   If m_SalesST06 <> "" And Trim(txtTrademark(27)) <> "" And Trim(txtTrademark(12).Text) <> "" Then
       If PUB_ChkOldCustomer(True, txtTrademark(27), Trim(txtTrademark(12)), m_SalesST15, m_SalesST06) = True Then
           IsMail = False
       End If
   End If
   
'Remove by Morgan 2009/8/20 國外部智權人員改可收所內信件
'   '2008/12/3 ADD BY SONIA 客戶檔之智權人員為國外部者不發mail
'   If IsMail = True Then
'      IsMail = PUB_CHKcusales(Trim(ChangeCustomerL(txtTrademark(9))) & "," & Trim(ChangeCustomerL(txtTrademark(24))) & "," & Trim(ChangeCustomerL(txtTrademark(25))) & "," & Trim(ChangeCustomerL(txtTrademark(26))) & "," & Trim(ChangeCustomerL(txtTrademark(27))))
'   End If
'   '2008/12/3 END
   
   'edit by nickc 2007/08/21 若申請人全空白，不發
   'If IsMail = False Then
   If IsMail = False Or (Trim(txtTrademark(9)) = "" And Trim(txtTrademark(24)) = "" And Trim(txtTrademark(25)) = "" And Trim(txtTrademark(26)) = "" And Trim(txtTrademark(27)) = "") Then
        oMailCount = ""
   End If
   
   If UCase(Mid(txtSystem, 1, 1)) <> "F" And oMailCount <> "" Then
      'Modify By Sindy 2010/11/26 申請人1~5為 X65299 或 X03072 的所有關係企業都不檢查業務區
      If Left(Trim(txtTrademark(9)), 6) <> "X65299" And Left(Trim(txtTrademark(9)), 6) <> "X03072" And _
         Left(Trim(txtTrademark(24)), 6) <> "X65299" And Left(Trim(txtTrademark(24)), 6) <> "X03072" And _
         Left(Trim(txtTrademark(25)), 6) <> "X65299" And Left(Trim(txtTrademark(25)), 6) <> "X03072" And _
         Left(Trim(txtTrademark(26)), 6) <> "X65299" And Left(Trim(txtTrademark(26)), 6) <> "X03072" And _
         Left(Trim(txtTrademark(27)), 6) <> "X65299" And Left(Trim(txtTrademark(27)), 6) <> "X03072" Then
         MsgBox "收文智權人員與客戶智權人員不同業務區，準備發 mail ！", , "注意！"
         oMailCount = oMailCount & Trim(txtTrademark(12).Text) & ";83002"
         oContext = oContext & vbCrLf + "收文智權人員： " + lblSales.Caption + vbCrLf + vbCrLf + "智權人員(區)不同！"
         PUB_SendMail strUserNum, oMailCount, "", "案件收文通知--此案收文非原智權人員(區)！", oContext
      End If
   End If
   'add by nickc 2007/05/16 加入若是本所期限小於等於當天，要發mail  通知
   oMailCount = ""
   If Mid(txtSystem, 1, 1) = "T" Then   'T,TF案
        'edit by nickc 2007/10/16 修改到table
        'oMailCount = "84027;69008"
        oMailCount = Pub_GetSpecMan("E")
   ElseIf Right(txtSystem, 1) = "T" Then
        'edit by nickc 2007/10/16 修改到table
        'oMailCount = "68005;72012"
        'Modified by Lydia 2021/07/30 商標及商標服務業務收文-因外商陳經理退休而修改程式控制
        'oMailCount = Pub_GetSpecMan("D")
        If txtSystem = "CFT" Then
           '以本所案號呼叫GetCFTSt16Manager抓主管
            oMailCount = GetCFTSt16Manager(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
        Else
            '先以本所案號呼叫PUB_GetFCTSalesNo抓出負責的人，再抓該員除ST55之外的最高主管NVL(NVL(ST54,ST53),ST52)
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
        End If
        'end 2021/07/30
        'add by nickc 2007/06/23 加入FCT 爭議案通知內商商爭  84027;69008     案件性質 202 除外，還是送外商，阿蓮說自己判斷，若為內商案件，他會再轉過來
        If txtSystem = "FCT" Then    'add by nickc 2007/08/01 不判斷的話  CFT 也會進入
            'edit by nickc 2007/08/10  內外商協議 202 都發
'2011/8/2 cancel by sonia
'            If txtTrademark(1) = "202" Then
'                'edit by nickc 2007/10/16 修改到table
'                'oMailCount = "68005;72012;84027;69008"
'                oMailCount = Pub_GetSpecMan("F")
'            End If
            Dim tmp960623 As New ADODB.Recordset
            Set tmp960623 = New ADODB.Recordset
            If tmp960623.State = 1 Then tmp960623.Close
            tmp960623.CursorLocation = adUseClient
            tmp960623.Open "select * from staff_group where sg01='C1' and sg02='FCT' and sg03='" & txtTrademark(1) & "' and sg03<>'202'  ", cnnConnection, adOpenStatic, adLockReadOnly
            If tmp960623.RecordCount <> 0 Then
                 If Trim(txtTrademark(11)) <> "" And Trim(txtTrademark(16)) <> "" Then
                     '2011/8/2 modify by sonia  FCT-030368 前2007/6/23mark此段造成FCT 爭議案都沒通知內商,只有202才全發
                     'oMailCount = "84027;69008"
                     'Modified by Lydia 2021/07/30 加發最高主管
                     'oMailCount = Pub_GetSpecMan("F")
                     strExc(1) = Pub_GetSpecMan("F")
                     If InStr(strExc(1), oMailCount) = 0 Then
                         oMailCount = strExc(1) & IIf(oMailCount = "", "", ";" & oMailCount)
                     Else
                         oMailCount = strExc(1)
                     End If
                     'end 2021/07/30
                 End If
            End If
            tmp960623.Close
            Set tmp960623 = Nothing
        End If
   ElseIf txtSystem = "S" Then
        'edit by nickc 2007/10/16 修改到table
        'oMailCount = "68005;72012"
        oMailCount = Pub_GetSpecMan("D")
   'edit by nickc 2008/04/23 CFC & CFT 都改發 68005
   'ElseIf txtSystem = "CFC" Then
   'Modified by Lydia 2021/07/30 debug: CFT在前面
   'ElseIf txtSystem = "CFC" Or txtSystem = "CFT" Then
   ElseIf txtSystem = "CFC" Then
        'edit by nickc 2007/10/16 修改到table
        'oMailCount = "68005;72012"
        oMailCount = Pub_GetSpecMan("L")
   End If
   If DBDATE(txtTrademark(11).Text) < strSrvDate(1) And Trim(txtTrademark(11).Text) <> "" And Trim(oMailCount) <> "" Then
      '2007/8/13 MODIFY BY SONIA 加智權人員
      'Modify By Sindy 2010/12/16 加業務區,費用
      'Modify By Sindy 2013/4/11 +規費,點數
      PUB_SendMail strUserNum, oMailCount, "", "案件收文 緊急 通知--此案已逾本所期限，請儘速辦理！", oContext2 & vbCrLf & "本所期限：" & ChangeWStringToTDateString(DBDATE(txtTrademark(11).Text)) & vbCrLf & "法定期限：" & ChangeWStringToTDateString(DBDATE(txtTrademark(16).Text)) & vbCrLf & "智權人員　：" & lblSales & vbCrLf & "業務區　：" & lblDepartment & vbCrLf & "費用　　：" & Format(txtTrademark(14), "##,##0") & vbCrLf & "規費　　：" & Format(txtTrademark(18), "##,##0") & vbCrLf & "點數　　：" & txtTrademark(15)
   End If
   If DBDATE(txtTrademark(11).Text) = strSrvDate(1) And Trim(txtTrademark(11).Text) <> "" And Trim(oMailCount) <> "" Then
      '2007/8/13 MODIFY BY SONIA 加智權人員
      'Modify By Sindy 2010/12/16 加業務區,費用
      'Modify By Sindy 2013/4/11 +規費,點數
      PUB_SendMail strUserNum, oMailCount, "", "案件收文 緊急 通知--此案已屆本所期限，請儘速辦理！", oContext2 & vbCrLf & "本所期限：" & ChangeWStringToTDateString(DBDATE(txtTrademark(11).Text)) & vbCrLf & "法定期限：" & ChangeWStringToTDateString(DBDATE(txtTrademark(16).Text)) & vbCrLf & "智權人員　：" & lblSales & vbCrLf & "業務區　：" & lblDepartment & vbCrLf & "費用　　：" & Format(txtTrademark(14), "##,##0") & vbCrLf & "規費　　：" & Format(txtTrademark(18), "##,##0") & vbCrLf & "點數　　：" & txtTrademark(15)
   End If
   'add by nickc 2007/10/16 假日前收文，且期限為假日
   'edit by nickc 2008/03/06 秀玲說判斷未到期的就好
   'If txtTrademark(11).Text <> "" Then
   If DBDATE(txtTrademark(11).Text) > strSrvDate(1) Then
        If (txtSystem = "T" Or txtSystem = "FCT") And ChkMyWeek(DBDATE(txtTrademark(11).Text)) = True And Trim(oMailCount) <> "" Then
           'Modify By Sindy 2010/12/16 加業務區,費用
           'Modify By Sindy 2013/4/11 +規費,點數
           PUB_SendMail strUserNum, oMailCount, "", "案件收文 緊急 通知--此案即將屆本所期限，且本所期限為假日，請儘速辦理！", oContext2 & vbCrLf & "本所期限：" & ChangeWStringToTDateString(DBDATE(txtTrademark(11).Text)) & vbCrLf & "法定期限：" & ChangeWStringToTDateString(DBDATE(txtTrademark(16).Text)) & vbCrLf & "智權人員　：" & lblSales & vbCrLf & "業務區　：" & lblDepartment & vbCrLf & "費用　　：" & Format(txtTrademark(14), "##,##0") & vbCrLf & "規費　　：" & Format(txtTrademark(18), "##,##0") & vbCrLf & "點數　　：" & txtTrademark(15)
        End If
        'add by nickc 2008/01/24 若是分所收文，期限為工作天且為隔天也要通知
        If (txtSystem = "T" Or txtSystem = "FCT") And pub_strUserOffice > "1" And Val(CompWorkDay(2, strSrvDate(1), 0)) = Val(DBDATE(txtTrademark(11).Text)) And Trim(oMailCount) <> "" Then
           'Modify By Sindy 2010/12/16 加業務區,費用
           'Modify By Sindy 2013/4/11 +規費,點數
           PUB_SendMail strUserNum, oMailCount, "", "案件收文 緊急 通知--此案為分所案件且將屆本所期限，本所期限為下一工作日，請儘速辦理！", oContext2 & vbCrLf & "本所期限：" & ChangeWStringToTDateString(DBDATE(txtTrademark(11).Text)) & vbCrLf & "法定期限：" & ChangeWStringToTDateString(DBDATE(txtTrademark(16).Text)) & vbCrLf & "智權人員　：" & lblSales & vbCrLf & "業務區　：" & lblDepartment & vbCrLf & "費用　　：" & Format(txtTrademark(14), "##,##0") & vbCrLf & "規費　　：" & Format(txtTrademark(18), "##,##0") & vbCrLf & "點數　　：" & txtTrademark(15)
        End If
   End If
End If
   Set adoquery = Nothing
End Function

Private Sub Form_Activate()

'Add by Morgan 2004/4/15
If bolActive Then
   Exit Sub
Else
   bolActive = True
End If

Dim strTKindName As String, strDate1 As String, strDate2 As String, strCode(5) As String, i As Integer
Dim bolAdd As Boolean 'Added by Lydia 2016/04/27

   Me.Refresh
   '判斷是否為TF，調整顯示farTF，fraElse
   If txtSystem.Text = 馬德里案 Then
      fraTF.Visible = True
      fraElse.Visible = False
      Check2.Visible = True 'Add By Sindy 2012/11/12
   Else
      fraTF.Visible = False
      fraElse.Visible = True
      'Add By Sindy 2012/11/12
      If Left(Trim(txtSystem), 1) <> "T" Then
         Check2.Visible = True
      Else
         Check2.Visible = False
      End If
      '2012/11/12 End
   End If
   
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
                                                   'Modify by Amy 2014/10/23 分割案不預帶
                                                   If intWhere <> 國外_CF And txtTrademark(1) <> "308" Then
                                                      txtTrademark(5) = 台灣國家代號
                                                      CheckKeyIn 5
                                                   End If
                                                End If
                        End Select
                        If LastDate = "" Then
                           txtTrademark(0).Text = GetTaiwanTodayDate
                        Else
                           txtTrademark(0).Text = LastDate
                        End If
                        txtTrademark_GotFocus 0
             Case 1
                        '修改：中間欄位不可輸入
                        fraWindow1.Enabled = True
                        Dim bolNew As Boolean
                        'edit by nickc 2007/02/06 不用 dll 了
                        'If obj001.IsNewCase(txtRecieveCode, bolNew) Then
                        If Cls001IsNewCase(txtRecieveCode, bolNew) Then
                           If bolNew Then
                              fraWindow2.Enabled = True
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
   
   ReDim m_CaseNa239(1 To TF_TM) 'Move by Lydia 2020/11/29 從CFP英國脫歐案管制內移上來
   
   If frm010001.intModifyKind <> 0 Or frm010001.intSaveMode <> 1 Then 'Memo by Lydia 2020/11/19 非新增，屬於查詢／修改
      ooReadTrademarkDatabaseR
   End If
   'Added by Lydia 2020/11/19 CFT英國脫歐案管制：新增英國脫歐案,先讀取歐盟案
   'Modified by Lydia 2020/12/01 判斷新案
   'If txtSystem = "CFT" And frm010001.txtCaseNa239 <> "" Then
   If txtSystem = "CFT" And frm010001.txtCode(0) & frm010001.txtCode(1) & frm010001.txtCode(2) = "" And frm010001.txtCaseNa239 <> "" Then
       Call ChgCaseNo(frm010001.txtCaseNa239.Text, m_CaseNa239)
           If txtTrademark(1) <> "101" Then  'Added by Lydia 2021/03/05 判斷非申請案；CFT歐盟尚未註冊案轉換英國申請案收文控管：針對2021.9.30前收文之英國新「申請101」案建立關聯案
               ooReadTrademarkDatabaseR
           End If
   End If
   'end 2020/11/19
   
   'Add By Sindy 2013/8/23
   '台灣新申請案電子送件
   'Modified by Lydia 2019/08/01 開放FCT舊案可勾選電子送件
   'If txtSystem = "FCT" And txtTrademark(5) = "000" And txtTrademark(1) = "101" Then
   If txtSystem = "FCT" And txtTrademark(5) = "000" Then
      chkWebApp.Visible = True
      'Added by Lydia 2021/06/23 因FCT之「註冊費」案件已全面採電子送件，請於收文時，系統預設「以電子送件」。
      'Modified by Lydia 2022/11/15 外商人員F11自行操作之FCT舊案收文時，一律預設電子送件
      'Modfied by Lydia 2022/12/12 外商人員自行操作之FCT舊案收文時，一律預設電子送件排除特定案件性質
      'If txtTrademark(1) = "717" Or frm010001.mRole = "F11" Then
      'Modified by Lydia 2023/11/09 1.請預設FCT新、舊案所有案件性質收文時為電子送件
                        '           2.預設電子送件時排除案件性質:705補收款,706其他,707調查,709協調,711文件公／簽證,713諮詢,721翻譯,723出具同意書,724徵求同意書
      'If txtTrademark(1) = "717" Or (frm010001.mRole = "F11" And InStr("705補收款,706其他,707調查,709協調,711文件公／簽證,713諮詢,721翻譯,723出具同意書,724徵求同意書", txtTrademark(1)) = 0) Then
      'Modified by Lydia 2024/10/18 +702刊登廣告,727分析 ---from 徐湘?
      If InStr("702刊登廣告,705補收款,706其他,707調查,709協調,711文件公／簽證,713諮詢,721翻譯,723出具同意書,724徵求同意書,727分析", txtTrademark(1)) = 0 Then
         'Added by Lydia 2023/11/29 取消外商1121110-05上線之請作，並修改如下：
                                    '煩請將電腦系統FCT下之101(商申) 102(延展) 201(補正) 202(申請意見書) 208 (補優先權證明) 301(變更) 302(更正) 304(申請英文證明) 305(催審) 307(自請拋棄專用權) 308(分割) 313(減縮商品) 501(移轉) 502(授權) 504(再授權)  共15個案件性質預設電子送件, 謝謝!
         'Modified by Lydia 2023/11/30 + 717(註冊費)
         If InStr("101(商申),102(延展),201(補正),202(申請意見書),208(補優先權證明),301(變更),302(更正),304(申請英文證明),305(催審),307(自請拋棄專用權),308(分割),313(減縮商品),501(移轉),502(授權),504(再授權),717(註冊費)", txtTrademark(1)) > 0 Then
         'end 2023/11/29
            chkWebApp.Value = 1
         End If 'Added by Lydia 2023/11/29
      End If
      'end 2021/06/23
      'Added by Lydia 2023/10/27 FCT舊案收文案件性質為「延展」，請鎖住收文畫面上之本所及法定期限。
      If frm010001.mRole = "F11" And txtCode(0) <> "" And txtTrademark(1) = "102" Then
          txtTrademark(11).Locked = True
          txtTrademark(16).Locked = True
      End If
      'end 2023/10/27
   Else
      chkWebApp.Visible = False
      chkWebApp.Value = 0
   End If
   '2013/8/23 END
   
   'Added by Lydia 2015/11/12
    TMQList = ""
    'Modified by Lydia 2016/03/28
    bolOpen130 = False
    If strSrvDate(1) >= TMQ電子化啟用日 And TypeName(Tmpfrm090130) <> "" And txtTrademark(1) = "101" And txtSystem = "T" Then
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
    End If
   'end 2015/11/12
   
   'Added by Lydia 2020/05/20 法律所案源收文：讀取法務案源檔
   If strSrvDate(1) >= 法律所案源收文啟用日 And frm010001.intModifyKind = 0 And txtTrademark(5) = "000" And (txtSystem = "FCT" Or txtSystem = "T" Or txtSystem = "TC") Then
      Call ReadLOS
   End If
   'end 2020/05/20
   'Add by Amy 2021/12/20 改form2.0 TopIndex 有問題,因判斷複雜,故先寫死
    If UCase(App.EXEName) = "TEWRITER" Or UCase(App.EXEName) = "WRITER" Then
        'Modify by Amy 2022/01/06 +txtTrademark(2).Enabled = True,AB1000598 從查詢進入txtTrademark(2).Enabled = false 會出現引號異常就跳離開-唐
        If txtTrademark(0) <> MsgText(601) And txtTrademark(1) <> MsgText(601) And txtTrademark(2).Enabled = True Then
            txtTrademark(2).SetFocus
        End If
    End If
    
   'Added by Lydia 2022/08/05
   If strSrvDate(1) >= 收文存檔模組化啟用日 Then
      If txtSystem.Text = 馬德里案 Then
          Call SetDBArray(True, txtRecieveCode, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), txtTFCode(2), txtTFCode(3))
      Else
          Call SetDBArray(True, txtRecieveCode, txtSystem, txtCode(0), txtCode(1), txtCode(2))
      End If
   End If
       
End Sub

'modify by sonia 2021/3/31 加TM22專用期止日
'Modifeid by Lydia 2021/04/15 +TM58案件備註
Private Sub ooReadTrademarkDatabaseR()
Dim tm01 As String, tm02 As String, tm03 As String, tm04 As String, tm05 As String, _
              tm08 As String, tm09 As String, _
              tm10 As String, tm23 As String, tm44 As String, tm22 As String, _
              cp05 As String, cp06 As String, cp07 As String, CP10 As String, cp11 As String, _
              cp13 As String, cp14 As String, cp16 As String, cp17 As String, _
              cp18 As String, cp19 As String, cp32 As String, cp56 As String, cu30 As String, CP64 As String, rt As Boolean
'add by nickc 2006/11/22
Dim TM78 As String, TM79 As String, TM80 As String, TM81 As String, CP89 As String, CP90 As String, CP91 As String, CP92 As String, TM32 As String
'add by nickc 2007/03/27
Dim TM45 As String
Dim strTemp As String, bolIsChina As Boolean
Dim CP150 As String 'Add By Sindy 2012/11/08
Dim nTM01 As String, nTM02 As String, nTM03 As String, nTM04 As String 'Added by Lydia 2020/11/19 要讀取之案號
Dim tm58 As String 'Added by Lydia 2021/04/15
Dim tm136 As String 'Added by Sindy 2022/12/7

   'Added by Lydia 2020/11/19  CFT英國脫歐案管制：要讀取之案號
   nTM01 = IIf(m_CaseNa239(1) <> "", m_CaseNa239(1), txtSystem.Text)
   nTM02 = IIf(m_CaseNa239(2) <> "", m_CaseNa239(2), txtCode(0))
   nTM03 = IIf(m_CaseNa239(3) <> "", m_CaseNa239(3), IIf(txtCode(1) = "", "0", txtCode(1)))
   nTM04 = IIf(m_CaseNa239(4) <> "", m_CaseNa239(4), IIf(txtCode(2) = "", "00", txtCode(2)))
   'end 2020/11/19
   
   If txtSystem.Text = 馬德里案 Then
       'add by nickc 2006/06/20
       CP10 = txtTrademark(1)
   '   rt = ReadTrademarkDatabase(frm010001.intModifyKind, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
   '          IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), tm05, tm06, _
   '          tm07, tm08, tm09, tm10, tm23, tm44, txtRecieveCode, cp05, cp06, cp07, CP10, cp11, cp13, cp14, cp16, cp17, _
   '          cp18, cp19, cp32, cp56, cu30, CP64)
   'edit by nickc 2006/11/22 加申請人
   '   rt = ReadTrademarkDatabase(frm010001.intModifyKind, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
             IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), tm05, _
             tm08, tm09, tm10, tm23, tm44, txtRecieveCode, cp05, cp06, cp07, CP10, cp11, cp13, cp14, cp16, cp17, _
             cp18, cp19, cp32, cp56, cu30, CP64)
   'edit by nickc 2007/03/27 加入彼所案號
   '   rt = ooReadTrademarkDatabase(frm010001.intModifyKind, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
             IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), tm05, _
             tm08, tm09, tm10, tm23, tm44, txtRecieveCode, cp05, cp06, cp07, CP10, cp11, cp13, cp14, cp16, cp17, _
             cp18, cp19, cp32, cp56, cu30, CP64, TM78, TM79, TM80, TM81, CP89, CP90, CP91, CP92, TM32)
      'modify by sonia 2021/3/31 加TM22專用期止日
      'Modifeid by Lydia 2021/04/15 +TM58案件備註
      rt = ooReadTrademarkDatabase(frm010001.intModifyKind, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
             IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), tm05, _
             tm08, tm09, tm10, tm23, tm44, txtRecieveCode, cp05, cp06, cp07, CP10, cp11, cp13, cp14, cp16, cp17, _
             cp18, cp19, cp32, cp56, cu30, CP64, TM78, TM79, TM80, TM81, CP89, CP90, CP91, CP92, TM32, TM45, CP150, tm22, tm58, tm136)
   Else
      CP10 = txtTrademark(1)
   '   rt = ReadTrademarkDatabase(frm010001.intModifyKind, txtSystem, txtCode(0), _
   '          IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), tm05, tm06, _
   '          tm07, tm08, tm09, tm10, tm23, tm44, txtRecieveCode, cp05, cp06, cp07, CP10, cp11, cp13, cp14, cp16, cp17, _
   '          cp18, cp19, cp32, cp56, cu30, CP64)
   'edit by nickc 2006/11/22 加申請人
   '   rt = ReadTrademarkDatabase(frm010001.intModifyKind, txtSystem, txtCode(0), _
             IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), tm05, _
             tm08, tm09, tm10, tm23, tm44, txtRecieveCode, cp05, cp06, cp07, CP10, cp11, cp13, cp14, cp16, cp17, _
             cp18, cp19, cp32, cp56, cu30, CP64)
      'edit by nickc 2007/3/27 加入彼所案號
      'rt = ooReadTrademarkDatabase(frm010001.intModifyKind, txtSystem, txtCode(0), _
             IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), tm05, _
             tm08, tm09, tm10, tm23, tm44, txtRecieveCode, cp05, cp06, cp07, CP10, cp11, cp13, cp14, cp16, cp17, _
             cp18, cp19, cp32, cp56, cu30, CP64, TM78, TM79, TM80, TM81, CP89, CP90, CP91, CP92, TM32)
      'Modified by Lydia 2020/11/19 改變數
      'rt = ooReadTrademarkDatabase(frm010001.intModifyKind, txtSystem, txtCode(0), _
             IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), tm05, _
             tm08, tm09, tm10, tm23, tm44, txtRecieveCode, cp05, cp06, cp07, CP10, cp11, cp13, cp14, cp16, cp17, _
             cp18, cp19, cp32, cp56, cu30, CP64, TM78, TM79, TM80, TM81, CP89, CP90, CP91, CP92, TM32, TM45, CP150)
      'modify by sonia 2021/3/31 加TM22專用期止日
      'Modifeid by Lydia 2021/04/15 +TM58案件備註
      rt = ooReadTrademarkDatabase(frm010001.intModifyKind, nTM01, nTM02, _
             nTM03, nTM04, tm05, _
             tm08, tm09, tm10, tm23, tm44, txtRecieveCode, cp05, cp06, cp07, CP10, cp11, cp13, cp14, cp16, cp17, _
             cp18, cp19, cp32, cp56, cu30, CP64, TM78, TM79, TM80, TM81, CP89, CP90, CP91, CP92, TM32, TM45, CP150, tm22, tm58, tm136)
   End If
   'NICK 900803 **********************
   txtCP64 = CP64
   ' **********************
   If rt Then
      If frm010001.intModifyKind <> 0 Then
         txtTrademark(0) = cp05
         txtTrademark(1) = CP10
         txtTrademark(2) = cp11
         txtTrademark(12) = cp13
         txtTrademark(13) = cu30
         txtTrademark(14) = cp16
         txtTrademark(15) = cp18
         txtTrademark(17) = cp32
         txtTrademark(18) = cp17
         txtTrademark(19) = cp19
         txtTrademark(21) = cp14
         CheckKeyIn 1
         CheckKeyIn 2
         CheckKeyIn 12
         CheckKeyIn 14
         If txtTrademark(1) = 移轉 Then
            fraPatition.Visible = True
            txtTrademark(20) = cp56
            CheckKeyIn 20
            txtTrademark(28) = CP89
            CheckKeyIn 28
            txtTrademark(29) = CP90
            CheckKeyIn 29
            txtTrademark(30) = CP91
            CheckKeyIn 30
            txtTrademark(31) = CP92
            CheckKeyIn 31
         Else
            fraPatition.Visible = False
         End If
         'Add By Sindy 2012/11/08
         If CP150 = "Y" Then
            Me.Check2.Value = 1
         End If
         '2012/11/08 End
      End If
      txtTrademark(11) = cp06
      txtTrademark(16) = cp07
      txtTrademark(3) = tm08
      txtTrademark(4) = tm09
      
      txtTrademark(35) = tm136 'Add By Sindy 2022/12/7
      'Modified by Morgan 2022/12/26
      If txtTrademark(35) = "" Then
         txtTrademark(35) = PUB_GetCertType(nTM01, nTM02, nTM03, nTM04)
      End If
      'end 2022/12/26
      'Added by Lydia 2020/11/19 CFT英國脫歐案管制
      If nTM01 = "CFT" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" Then
           txtTrademark(5) = "201"  '預設:英國
           m_CaseNa239(5) = tm05
      Else
      'end 2020/11/19
           txtTrademark(5) = tm10
      End If 'Added by Lydia 2020/11/19
      txtTrademark(6) = tm05
   '   txtTrademark(7) = tm06
   '   txtTrademark(8) = tm07
      txtTrademark(9) = tm23
      'add by nickc 2006/11/22
      txtTrademark(24) = TM78
      txtTrademark(25) = TM79
      txtTrademark(26) = TM80
      txtTrademark(27) = TM81
      txtTrademark(32) = TM32
      'add by nickc 2007/03/27
      txtTrademark(33) = TM45
      txtTrademark(10) = tm44
      m_TM22 = tm22 'add by sonia 2021/3/23
      m_TM58 = tm58 'Added by Lydia 2021/04/15 +TM58案件備註
      
      'Add By Cheng 2001/12/17
      '顯示智權人員代碼
      'txtTrademark(12) = cp13  '2011/5/11 cancel by sonia 偶而改智權人員收文會忘記打所以不自動帶
      'Modify By Cheng 2002/01/03
      If Len("" & txtTrademark(9).Text) > 0 Then CheckKeyIn 9
      'Add By Sindy 2011/01/07
      If Len("" & txtTrademark(24).Text) > 0 Then CheckKeyIn 24
      If Len("" & txtTrademark(25).Text) > 0 Then CheckKeyIn 25
      If Len("" & txtTrademark(26).Text) > 0 Then CheckKeyIn 26
      If Len("" & txtTrademark(27).Text) > 0 Then CheckKeyIn 27
      '2011/01/07 End
      
      CheckKeyIn 10
      CheckKeyIn 3
      CheckKeyIn 5
      'Add By Cheng 2001/12/17
      '顯示智權人員姓名
      If txtTrademark(12) <> "" Then CheckKeyIn 12
   Else
      If frm010001.intModifyKind <> 0 Then
         MsgBox "讀取資料時發生錯誤!!", vbCritical
         bolLeave = True
         Unload Me
      Else
         txtTrademark(11) = cp06
         txtTrademark(16) = cp07
         txtTrademark(3) = tm08
         txtTrademark(4) = tm09
         txtTrademark(5) = tm10
         txtTrademark(6) = tm05
   '      txtTrademark(7) = tm06
   '      txtTrademark(8) = tm07
         txtTrademark(9) = tm23
         'add by nickc 2006/11/22
         txtTrademark(24) = TM78
         txtTrademark(25) = TM79
         txtTrademark(26) = TM80
         txtTrademark(27) = TM81
         txtTrademark(32) = TM32
         'add by nickc 2007/03/27
         txtTrademark(33) = TM45
         
         txtTrademark(10) = tm44
         CheckKeyIn 9
         CheckKeyIn 10
         If frm010001.intSaveMode <> 1 Then
            CheckKeyIn 3
            CheckKeyIn 5
         End If
      End If
   End If
   'NICK 900803 **********************
   If frm010001.intChoose = 1 Then
      txtTrademark(2) = "90"
      CheckKeyIn (2)
   End If
   ' **********************
End Sub

Private Sub Form_Load()
Dim oLbl 'Modify by Amy 2021/12/16 原:As LABEL
   
   MoveFormToCenter Me
   
   bolLeave = False
   intLeaveKind = 1
   If frm010001.intChoose = 1 Then
      fraPromoter.Visible = True
      txtTrademark(17) = "N"
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
   
   'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
   Label35.Visible = False
   txtTrademark(34).Visible = False
   Check1.Visible = False
   
   'Added by Lydia 2020/12/15 CFT緬甸重新申請案：預設位置
   lblCFTna048.Top = Label35.Top
   txtCFTNa048.Top = txtTrademark(34).Top
   
   'Added by Lydia 2018/12/10 商標種類：預設位置
   Label42.Top = Label39.Top: textCP143.Top = textEP34.Top
   
   'Add By Sindy 2023/5/30
   m_strIR01 = frm010001.m_strIR01
   m_strIR02 = frm010001.m_strIR02
   m_strIR03 = frm010001.m_strIR03
   m_strIR04 = frm010001.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2023/5/30 END
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

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2021/6/22
   Where01ToGo intLeaveKind
   'Add By Cheng 2002/07/18
   'Modify by Amy 2021/12/20 改Form2.0後,存檔按Enter會當掉,改在呼叫時清除記憶體變數
   'Set frm010004 = Nothing
   'Added by Lydia 2015/11/12
   If TypeName(m_PrevForm) <> "Nothing" Then
      Set m_PrevForm.Tmpfrm090130 = Tmpfrm090130
   End If
    
   'Add By Sindy 2023/5/30
   If m_strIR01 <> "" Then
      If Not m_PrevFormIR Is Nothing Then
         Set m_PrevFormIR = Nothing
      End If
   End If
   '2023/5/30 END
   
   stChkForm = Me.Name 'Add by Amy 2021/12/21
End Sub

'Add By Sindy 2012/5/8
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
Private Sub textCP122_GotFocus()
   TextInverse textCP122
End Sub
Private Sub textCP122_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'2012/5/8 End

'Added by Lydia 2018/12/10
Private Sub textCP143_GotFocus()
   TextInverse textCP143
End Sub
Private Sub textCP143_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'end 2018/12/10

'Modify by Amy 2021/12/16 原:Integer
Private Sub txtTrademark_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
        'edit by nickc 2006/11/23
        'Case 9, 10, 12, 17, 20
        Case 9, 10, 12, 17, 21, 20, 24, 25, 26, 27, 28, 29, 30, 31
                KeyAscii = UpperCase(KeyAscii)
        Case 13
                'Modify by Amy 2021/12/16 +txtTrademark(Index)
                KeyAscii = ChangeZIP(KeyAscii, txtTrademark(Index))
        'Modify By Sindy 2022/12/7 證書形式
        Case 35
              If (KeyAscii < 49 Or KeyAscii > 50) And KeyAscii <> 8 Then
                 KeyAscii = 0
                 Beep
              End If
              '2022/12/7 END
   End Select
End Sub
Private Sub txtTrademark_Change(Index As Integer)
   Select Case Index
             Case 2
                        lblCaseSource.Caption = ""
             Case 3
                        lblTrademarkKind = ""
             Case 5
                        lblNation.Caption = ""
             Case 9
                        'lblPetition.Caption = ""
                        lblPetition(0).Caption = ""
                        txtTrademark(13).Text = ""
             'add by nickc 2006/11/22
             Case 24, 25, 26, 27
                        lblPetition(Index - 23).Caption = ""
             Case 28, 29, 30, 31
                    lblPetitionName(Index - 27).Caption = ""
             Case 10
                        lblAgent.Caption = ""
             Case 12
                        lblSales.Caption = ""
                        lblDepartment = ""
                        m_SalesST15 = "" 'Added by Lydia 2019/02/14
                        m_SalesST06 = "" 'Added by Lydia 2019/09/16
             Case 20
                        'edit by nickc 2006/11/22
                        'lblPetitionName = ""
                        lblPetitionName(0) = ""
             'add  by nickc 2006/11/22
             Case 28
                        lblPetitionName(1) = ""
             Case 29
                        lblPetitionName(2) = ""
             Case 30
                        lblPetitionName(3) = ""
             Case 31
                        lblPetitionName(4) = ""
             
             Case 21
                        lblPromoter = ""
   End Select
End Sub

Private Sub txtTrademark_Validate(Index As Integer, Cancel As Boolean)
Dim ii As Integer
Dim arrTM09 '商品類別
'add by nickc 2006/12/15
Dim intMoney As Long  '倍數
Dim strRetrunText As String 'Add By Sindy 2017/5/17
Dim strPassTM09 As String  'Added by Lydia 2023/02/10

   Select Case Index
      Case 12
           'add by nick 2005/01/04
           If txtTrademark(Index).Text <> "" And txtTrademark(Index) < "63001" Then
                MsgBox "智權人員編號不可小於 63001！", , "注意！"
                Cancel = True
                Exit Sub
           End If
          'add by nick 2004/12/08 因為之前的 智權人員並沒有抓
              Dim strTemp As String, strTemp1 As String
              'edit by nickc 2007/02/02 不用 dll 了
              'If Not objPublicData.GetStaff(txtTrademark(12).Text, strTemp, strTemp1) Then
              If Not ClsPDGetStaff(txtTrademark(12).Text, strTemp, strTemp1) Then
                  Cancel = True
                  Exit Sub
              End If
              'add by nickc 2006/11/02
              'Modified by Lydia 2019/02/14
              'GetST15 txtTrademark(12).Text, strTemp1
              'Modified by Lydia 2019/09/16 +st06
              'm_SalesST15 = GetST15(txtTrademark(12).Text, strTemp1)
              m_SalesST15 = GetST15(txtTrademark(12).Text, strTemp1, , m_SalesST06)
   
              lblSales.Caption = strTemp
              lblDepartment = strTemp1
              
              'Added by Lydia 2020/04/08 檢查案件或智權人員是否為法務部
              If PUB_ChkSalesL(txtSystem, txtTrademark(12).Text) = False Then
              End If
              'end 2020/04/08
              
              'Added by Lydia 2019/02/14 創新業務部人員收文控管
              If PUB_ChkIsT10T20("2", txtTrademark(12).Text, m_Tuser, strTemp) = True Then
                  txtTrademark(12) = m_Tuser
                  lblSales.Caption = strTemp
                  txtTrademark(12).SetFocus
                  Call txtTrademark_GotFocus(12)
                  Cancel = True
                  Exit Sub
              End If
              'end 2019/02/14
              
          'add by nick 2004/12/08  當收文業務區與客戶檔業務區不同時提示
          'edit by nickc 2007/05/10 改成檢查所有申請人，一個符合就不提醒
          'Dim oStrCuSales As String
          'oStrCuSales = ""
          Dim oStrCuSales1 As String
          Dim oStrCuSales2 As String
          Dim oStrCuSales3 As String
          Dim oStrCuSales4 As String
          Dim oStrCuSales5 As String
          Dim oContext As String
          Dim IsMail As Boolean
          oStrCuSales1 = ""
          oStrCuSales2 = ""
          oStrCuSales3 = ""
          oStrCuSales4 = ""
          oStrCuSales5 = ""
          IsMail = True
   
          'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
          'If GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(9).Text), oStrCuSales1) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(9).Text) <> "" Then
          'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
          If ChkSameCuArea(Trim(txtTrademark(9)), Trim(txtTrademark(12)), , , , , Trim(txtTrademark(10))) = False And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(9).Text) <> "" Then
          Else
               If Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(9).Text) <> "" Then
                   IsMail = False
               End If
          End If
          'Added by Lydia 2019/09/16 檢查是否為待活化客戶
          If m_SalesST06 <> "" And Trim(txtTrademark(9)) <> "" And Trim(txtTrademark(12).Text) <> "" Then
              If PUB_ChkOldCustomer(False, txtTrademark(9), Trim(txtTrademark(12)), m_SalesST15, m_SalesST06) = True Then
                  IsMail = False
              End If
          End If
          
          'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
          'If GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(24).Text), oStrCuSales2) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(24).Text) <> "" Then
          'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
          If ChkSameCuArea(Trim(txtTrademark(24)), Trim(txtTrademark(12)), , , , , Trim(txtTrademark(10))) = False And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(24).Text) <> "" Then
          Else
               If Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(24).Text) <> "" Then
                   IsMail = False
               End If
          End If
          'Added by Lydia 2019/09/16 檢查是否為待活化客戶
          If m_SalesST06 <> "" And Trim(txtTrademark(24)) <> "" And Trim(txtTrademark(12).Text) <> "" Then
              If PUB_ChkOldCustomer(False, txtTrademark(24), Trim(txtTrademark(12)), m_SalesST15, m_SalesST06) = True Then
                  IsMail = False
              End If
          End If
          
          'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
          'If GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(25).Text), oStrCuSales3) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(25).Text) <> "" Then
          'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
          If ChkSameCuArea(Trim(txtTrademark(25)), Trim(txtTrademark(12)), , , , , Trim(txtTrademark(10))) = False And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(25).Text) <> "" Then
          Else
               If Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(25).Text) <> "" Then
                   IsMail = False
               End If
          End If
          'Added by Lydia 2019/09/16 檢查是否為待活化客戶
          If m_SalesST06 <> "" And Trim(txtTrademark(25)) <> "" And Trim(txtTrademark(12).Text) <> "" Then
              If PUB_ChkOldCustomer(False, txtTrademark(25), Trim(txtTrademark(12)), m_SalesST15, m_SalesST06) = True Then
                  IsMail = False
              End If
          End If
          
          'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
          'If GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(26).Text), oStrCuSales4) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(26).Text) <> "" Then
          'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
          If ChkSameCuArea(Trim(txtTrademark(26)), Trim(txtTrademark(12)), , , , , Trim(txtTrademark(10))) = False And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(26).Text) <> "" Then
          Else
               If Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(26).Text) <> "" Then
                   IsMail = False
               End If
          End If
          'Added by Lydia 2019/09/16 檢查是否為待活化客戶
          If m_SalesST06 <> "" And Trim(txtTrademark(26)) <> "" And Trim(txtTrademark(12).Text) <> "" Then
              If PUB_ChkOldCustomer(False, txtTrademark(26), Trim(txtTrademark(12)), m_SalesST15, m_SalesST06) = True Then
                  IsMail = False
              End If
          End If
          
          'Modify by Amy 2017/01/03 因加MCTF判斷,故改判斷ChkSameCuArea
          'If GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(27).Text), oStrCuSales5) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(27).Text) <> "" Then
          'modify by sonia 2021/11/25 MCT案加傳FC代理人來判斷ChkSameCuArea
          If ChkSameCuArea(Trim(txtTrademark(27)), Trim(txtTrademark(12)), , , , , Trim(txtTrademark(10))) = False And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(27).Text) <> "" Then
          Else
               If Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(27).Text) <> "" Then
                   IsMail = False
               End If
          End If
          'Added by Lydia 2019/09/16 檢查是否為待活化客戶
          If m_SalesST06 <> "" And Trim(txtTrademark(27)) <> "" And Trim(txtTrademark(12).Text) <> "" Then
              If PUB_ChkOldCustomer(False, txtTrademark(27), Trim(txtTrademark(12)), m_SalesST15, m_SalesST06) = True Then
                  IsMail = False
              End If
          End If
          
'Remove by Morgan 2009/8/20 國外部智權人員改可收所內信件
'                  '2008/12/3 ADD BY SONIA 客戶檔之智權人員為國外部者不發mail
'                  If IsMail = True Then
'                     IsMail = PUB_CHKcusales(Trim(ChangeCustomerL(txtTrademark(9))) & "," & Trim(ChangeCustomerL(txtTrademark(24))) & "," & Trim(ChangeCustomerL(txtTrademark(25))) & "," & Trim(ChangeCustomerL(txtTrademark(26))) & "," & Trim(ChangeCustomerL(txtTrademark(27))))
'                  End If
'                  '2008/12/3 END
            
          '2006/8/2 MODIFY BY SONIA TXTSYSTEM只判斷1碼,因為FG
          'If UCase(Mid(txtSystem, 1, 2)) <> "FC" And GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(9).Text), oStrCuSales) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(9).Text) <> "" Then
          'edit by nickc 2007/05/10 改成檢查所有申請人，一個符合就不提醒
          'If UCase(Mid(txtSystem, 1, 1)) <> "F" And GetST15(txtTrademark(12).Text) <> GetCuSales(ChangeCustomerL(txtTrademark(9).Text), oStrCuSales) And Trim(txtTrademark(12).Text) <> "" And Trim(txtTrademark(9).Text) <> "" Then
          'edit by nickc 2008/03/26 若是申請人皆空白，就不管
          'If UCase(Mid(txtSystem, 1, 1)) <> "F" And IsMail = True Then
          If UCase(Mid(txtSystem, 1, 1)) <> "F" And IsMail = True And (txtTrademark(9) <> "" Or txtTrademark(24) <> "" Or txtTrademark(25) <> "" Or txtTrademark(26) <> "" Or txtTrademark(27) <> "") Then
               'Add By Sindy 2009/10/19
               '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail，不顯示訊息
               oMailCount = ""
               If txtTrademark(9) <> "" Then
                  'Modified by Lydia 2019/02/14
                  'If Left(Trim(GetST15(txtTrademark(12).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(9).Text), oStrCuSales1)), 1) = "F" Then
                  If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(9).Text), oStrCuSales1)), 1) = "F" Then
                  Else
                     oMailCount = "Y"
                  End If
               End If
               If txtTrademark(24) <> "" Then
                  'Modified by Lydia 2019/02/14
                  'If Left(Trim(GetST15(txtTrademark(12).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(24).Text), oStrCuSales1)), 1) = "F" Then
                  If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(24).Text), oStrCuSales1)), 1) = "F" Then
                  Else
                     oMailCount = "Y"
                  End If
               End If
               If txtTrademark(25) <> "" Then
                  'Modified by Lydia 2019/02/14
                  'If Left(Trim(GetST15(txtTrademark(12).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(25).Text), oStrCuSales1)), 1) = "F" Then
                  If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(25).Text), oStrCuSales1)), 1) = "F" Then
                  Else
                     oMailCount = "Y"
                  End If
               End If
               If txtTrademark(26) <> "" Then
                  'Modified by Lydia 2019/02/14
                  'If Left(Trim(GetST15(txtTrademark(12).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(26).Text), oStrCuSales1)), 1) = "F" Then
                  If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(26).Text), oStrCuSales1)), 1) = "F" Then
                  Else
                     oMailCount = "Y"
                  End If
               End If
               If txtTrademark(27) <> "" Then
                  'Modified by Lydia 2019/02/14
                  'If Left(Trim(GetST15(txtTrademark(12).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(27).Text), oStrCuSales1)), 1) = "F" Then
                  If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtTrademark(27).Text), oStrCuSales1)), 1) = "F" Then
                  Else
                     oMailCount = "Y"
                  End If
               End If
               If Trim(oMailCount) <> "" Then
               '2009/10/19 End
                  'Modify By Sindy 2010/11/26 申請人1~5為 X65299 或 X03072 的所有關係企業都不檢查業務區
                  If Left(Trim(txtTrademark(9)), 6) <> "X65299" And Left(Trim(txtTrademark(9)), 6) <> "X03072" And _
                     Left(Trim(txtTrademark(24)), 6) <> "X65299" And Left(Trim(txtTrademark(24)), 6) <> "X03072" And _
                     Left(Trim(txtTrademark(25)), 6) <> "X65299" And Left(Trim(txtTrademark(25)), 6) <> "X03072" And _
                     Left(Trim(txtTrademark(26)), 6) <> "X65299" And Left(Trim(txtTrademark(26)), 6) <> "X03072" And _
                     Left(Trim(txtTrademark(27)), 6) <> "X65299" And Left(Trim(txtTrademark(27)), 6) <> "X03072" Then
                     MsgBox "收文智權人員與客戶智權人員不同業務區！", , "注意！"
                  End If
               End If
          End If
      Case 3 '證明標章
          'Added by Lydia 2018/12/10 檢查商標種類
          If CheckKeyIn(Index) = -1 Then
             Cancel = True
             GoTo EXITSUB
          End If
          'end 2018/12/10
          'Add By Sindy 2015/6/30 證明標章時商品類別為證
          If Me.txtTrademark(Index).Text = "7" And (txtSystem = "FCT" Or txtSystem = "T") Then
             Me.txtTrademark(4).Text = "證"
          End If
          '2015/6/30 END
          'add by sonia 2021/1/12 團體標章時商品類別為團
          If Me.txtTrademark(Index).Text = "8" And (txtSystem = "FCT" Or txtSystem = "T") Then
             Me.txtTrademark(4).Text = "團"
          End If
          'end 2021/1/12
      Case 4 '商品類別
'               cmdOK(0).Default = True
'               cmdOK(0).CausesValidation = True
          'Add By Cheng 2003/12/31
          '若有輸入商品類別
          If Me.txtTrademark(Index).Text <> "" Then
              If CheckKeyIn(Index) = -1 Then
                 Cancel = True
                 GoTo EXITSUB
              End If
              arrTM09 = Split(Me.txtTrademark(Index).Text, ",")
              strPassTM09 = "" 'Added by Lydia 2023/02/10
              For ii = LBound(arrTM09) To UBound(arrTM09)
                  If Len(arrTM09(ii)) < 2 Or Len(arrTM09(ii)) > 3 Then
                     If Me.txtTrademark(3).Text <> "7" And Me.txtTrademark(3).Text <> "8" Then 'Add By Sindy 2015/6/30 +if   'modify by sonia 2021/1/8 +Me.txtTrademark(3).Text <> "8"
                        MsgBox "商品類別 <" & arrTM09(ii) & "> 不可小於二碼且不可大於三碼!!!", vbExclamation + vbOKOnly
                        Cancel = True
                        Exit For
                     End If
                  End If
                  'Added by Lydia 2023/02/10 商品類別不可重複
                  If InStr(strPassTM09 & ",", arrTM09(ii) & ",") > 0 Then
                        MsgBox "商品類別 <" & arrTM09(ii) & "> 重覆輸入!!!", vbExclamation + vbOKOnly
                        Cancel = True
                        Exit For
                  End If
                  strPassTM09 = strPassTM09 & arrTM09(ii) & ","
                  'end 2023/02/10
              Next ii
          End If
          'add by nickc 2005/06/03
          txtTrademark(Index).Text = Replace(txtTrademark(Index).Text, " ", "")
          'End
          '2014/5/5 add by sonia T-192228
          If Trim(txtTrademark(5)) = "000" And txtTrademark(1) = "102" And txtTrademark(4) = "" Then
             MsgBox "台灣延展新案件, 請輸入商品類別以便檢查規費！", vbCritical
             Cancel = True
          End If
          '2014/5/5 end
      'add by nickc 2006/11/30
      Case 32 '商品組群
          If Me.txtTrademark(Index).Text <> "" Then
              If CheckKeyIn(Index) = -1 Then
                 Cancel = True
                 GoTo EXITSUB
              End If
              'Modify By Sindy 2024/4/18 商品組群欄人員貼上資料後將全形或半形的「；」分號，轉為半形的逗號存入TM32。
              Me.txtTrademark(Index).Text = Replace(Replace(Me.txtTrademark(Index).Text, ";", ","), "；", ",")
              '2024/4/18 END
              arrTM09 = Split(Me.txtTrademark(Index).Text, ",")
              For ii = LBound(arrTM09) To UBound(arrTM09)
                  If Len(arrTM09(ii)) < 4 Or Len(arrTM09(ii)) > 6 Then
                      MsgBox "商品組群 <" & arrTM09(ii) & "> 不可小於四碼且不可大於六碼!!!", vbExclamation + vbOKOnly
                      Cancel = True
                      Exit For
                  End If
              Next ii
          End If
          txtTrademark(Index).Text = Replace(txtTrademark(Index).Text, " ", "")
      Case 5
         If CheckKeyIn(Index) <> -1 Then
            CheckKeyIn 1
                  
            CheckKeyIn 3
'                   CheckKeyIn 14
            ' 91.09.11 marked by louis
            'CheckKeyIn 18
            
            'Added by Lydia 2020/12/15 CFT緬甸重新申請案：顯示欄位
            If txtSystem = "CFT" And txtTrademark(5) = "048" And txtTrademark(1) = "101" Then
                lblCFTna048.Visible = True
                txtCFTNa048.Visible = True
            Else
                lblCFTna048.Visible = False
                txtCFTNa048.Visible = False
                txtCFTNa048.Text = ""
            End If
            'end 2020/12/15
         Else
            Cancel = True
         End If
      'Add By Cheng 2001/12/26
      'edit by nickc 2006/11/22
      'Case 9 '申請人
      Case 9, 24, 25, 26, 27, 28, 29, 30, 31
         '若申請人有輸入才做Check動作
         If Len(Trim(Me.txtTrademark(Index).Text)) > 0 Then
            If CheckKeyIn(Index) = -1 Then
               Cancel = True
            End If
         End If
      Case 10 '代理人
         '若代理人有輸入才做Check動作
         If Len(Trim(Me.txtTrademark(Index).Text)) > 0 Then
            If CheckKeyIn(Index) = -1 Then
               Cancel = True
            End If
         End If
         '若申請人與代理人同時空白時
         If Len(Trim(Me.txtTrademark(9).Text)) <= 0 And Len(Trim(Me.txtTrademark(10).Text)) <= 0 Then
            MsgBox "申請人與代理人必須至少輸入一項!!!", vbExclamation
'                  Cancel = True
         End If
      Case 7
         
         If (fraTM15.Visible = True) Then
            'Add by Morgan 2003/11/26
            Dim adoquery As New ADODB.Recordset, strSql As String, bolErr As Boolean
            
            If Trim(txtTrademark(7).Text) = "" Then
               MsgBox "商標審定號不可為空白！", vbCritical
               Cancel = True
            Else
               'Add By Sindy 2010/8/31 檢查審定號所輸入的長度是否正確
               If bolNewAppNoFormat Then
                  'Add By Sindy 2017/5/17 + strRetrunText
                  If PUB_ChkTm12Tm15Length("2", txtTrademark(7), txtSystem, txtCode(0), txtCode(1), txtCode(2), txtTrademark(5), , , strRetrunText) = False Then
                     Cancel = True
                     Exit Sub
                  'Add By Sindy 2017/5/17
                  Else
                     txtTrademark(7) = strRetrunText
                  '2017/5/17 END
                  End If
               '2010/8/31 End
               Else
                  '2009/1/20 ADD BY SONIA 台灣及大陸案審定號要輸8碼
                  If (txtTrademark(5) = "000" Or txtTrademark(5) = "020") And GetTextLength(Trim(txtTrademark(7).Text)) <> 8 Then
                     MsgBox "台灣商標審定號, 前面補 0(零)補滿 8 碼！", vbCritical
                     Cancel = True
                     Exit Sub
                  End If
               End If
               '2009/1/20 END
               '93.8.17 modify by sonia 加入商標種類的控制
               'strSQL = "select DECODE( TM01,'TF',TM01||'-'||SUBSTR(TM02,1,5)||'-'||SUBSTR(TM02,6,1)||'-'||TM03||'-'||TM04,TM01||'-'||TM02||'-'||TM03||'-'||TM04) as C1" & _
               '   " from TRADEMARK where TM10='" & Me.txtTrademark(5) & "' AND TM16='1' and TM15='" & Me.txtTrademark(Index) & "'"
               strSql = "select DECODE( TM01,'TF',TM01||'-'||SUBSTR(TM02,1,5)||'-'||SUBSTR(TM02,6,1)||'-'||TM03||'-'||TM04,TM01||'-'||TM02||'-'||TM03||'-'||TM04) as C1" & _
                  " from TRADEMARK where TM10='" & Me.txtTrademark(5) & "' AND TM16='1' and TM15='" & Me.txtTrademark(Index) & "' and TM08='" & Me.txtTrademark(3) & "'"
               '93.8.17 end
               adoquery.CursorLocation = adUseClient
               adoquery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If Not (adoquery.BOF And adoquery.EOF) Then
                  MsgBox "此審定號已有本所資料，本所案號為 " & adoquery.Fields(0).Value & " ，不可收新案號！", vbCritical
                  Cancel = True
               Else
                  Cancel = False
               End If
               adoquery.Close
            End If
         End If
         '---End
      'Add By Sindy 2013/3/12 改為點數>=5時預設為會稿，反之預設為不會稿且鎖住欄位
      Case 15
         If Frame21.Visible = True And textEP34.Visible = True Then  'Modified by Lydia 2018/12/10 +判斷顯示
            If Trim(txtTrademark(1)) = "613" Or Trim(txtTrademark(1)) = "612" Then
               If Val(txtTrademark(15)) >= 5 Then
                  textEP34.Text = "Y" '會稿
                  textEP34.Enabled = True
               Else
                  textEP34.Text = "N" '不會稿
                  textEP34.Enabled = False
               End If
            End If
         End If
      '2013/3/12 End
      
      '92.12.21 ADD BY SONIA
'      Case 18
'          'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'          If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" Then
'            'add by nick 2005/02/15
'            Dim TmpArr As Variant
'            If txtTrademark(4) <> "" Then
'               TmpArr = Split(IIf(Right(txtTrademark(4), 1) = ",", Mid(txtTrademark(4), 1, Len(txtTrademark(4)) - 1), txtTrademark(4)), ",")
'            End If
'            If txtTrademark(1) = "101" And txtTrademark(5) = "000" And txtTrademark(18) = "4000" Then
'               MsgBox "台灣商標已修法, 申請案規費不可為 4000, 請與智權人員確認！", vbCritical
'               Cancel = True
'               'add by nick 2005/02/15
'               Exit Sub
'            End If
'            '2008/6/24 add by sonia 移轉,授權規費固定
'            If (txtTrademark(1) = "501" Or txtTrademark(1) = "502") And txtTrademark(5) = "000" And txtTrademark(18) <> "2000" Then
'               MsgBox "台灣商標移轉或授權規費為 2000, 請確認！", vbCritical
'               Cancel = True
'            End If
'            '2008/6/24 END
'
''edit by nick 2005/02/15 每類 1000
'            '93.10.2 add by sonia
''               If txtTrademark(1) = "715" And txtTrademark(5) = "000" And txtTrademark(18) <> "1000" Then
''                  MsgBox "國內商標第一期註冊費規費為 1000, 請確認！", vbCritical
''                  Cancel = True
''               End If
'            '93.10.2 END
'         '92.12.21 END
'             If txtTrademark(5) = "000" And txtTrademark(4) <> "" Then
'                '2010/1/14 MODIFY BY SONIA 葉經理說因為以電子送件故可以2700倍數收文
'                'If txtTrademark(1) = "101" And Val(txtTrademark(18)) < (3000# * (Val(UBound(TmpArr)) + 1)) Then
'                '   MsgBox "國內商標申請規費最少要" & str(3000# * (Val(UBound(TmpArr)) + 1)) & ", 請確認！", vbCritical
'                '   Cancel = True
'                '   Exit Sub
'                'End If
'                '2010/2/6 MODIFY BY SONIA 發現有規費5000者再修改FCT-30038
'                'If txtTrademark(1) = "101" Then
'                '   If (Val(txtTrademark(18)) <> (3000# * (Val(UBound(TmpArr)) + 1))) And (Val(txtTrademark(18)) <> (2700# * (Val(UBound(TmpArr)) + 1))) Then
'                '       MsgBox "台灣商標申請規費要" & str(3000# * (Val(UBound(TmpArr)) + 1)) & "或" & str(2700# * (Val(UBound(TmpArr)) + 1)) & "(電子申請), 請確認！", vbCritical
'                '   End If
'                'End If
'                'Modify By Sindy 2010/8/25 葉經理說要取消此條件控管
''                If txtTrademark(1) = "101" And Val(txtTrademark(18)) < (3000# * (Val(UBound(TmpArr)) + 1)) Then
''                   If Val(txtTrademark(18)) <> 2700 Then  '電子送件
''                      MsgBox "國內商標申請規費最少要" & str(3000# * (Val(UBound(TmpArr)) + 1)) & ", 請確認！", vbCritical
''                      Cancel = True
''                      Exit Sub
''                   End If
''                End If
'                '2010/2/6 END
'                '2010/1/14 END
'                'add by nickc 2006/12/15
'                intMoney = 1
'                If Val(txtTrademark(16)) < Val(GetTaiwanTodayDate) And Val(txtTrademark(16)) <> 0 Then
'                   intMoney = 2
'                End If
''edit by nickc 2006/12/15
''                     If txtTrademark(1) = "715" And Val(txtTrademark(18)) <> (1000# * (Val(UBound(TmpArr)) + 1)) Then
''                        MsgBox "國內商標第一期註冊費規費要" & str(1000# * (Val(UBound(TmpArr)) + 1)) & ", 請確認！", vbCritical
'                If txtTrademark(1) = "715" And Val(txtTrademark(18)) <> (1000# * (Val(UBound(TmpArr)) + 1) * intMoney) Then
'                   MsgBox "台灣商標第一期註冊費規費要" & str(1000# * (Val(UBound(TmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
'                   Cancel = True
'                   Exit Sub
'                End If
''edit by nickc 2006/12/15
''                     If txtTrademark(1) = "716" And Val(txtTrademark(18)) <> (1500# * (Val(UBound(TmpArr)) + 1)) Then
''                        MsgBox "國內商標第二期註冊費規費要" & str(1500# * (Val(UBound(TmpArr)) + 1)) & ", 請確認！", vbCritical
'                If txtTrademark(1) = "716" And Val(txtTrademark(18)) <> (1500# * (Val(UBound(TmpArr)) + 1) * intMoney) Then
'                   MsgBox "台灣商標第二期註冊費規費要" & str(1500# * (Val(UBound(TmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
'                   Cancel = True
'                   Exit Sub
'                End If
''edit by nickc 2006/12/15
''                     If txtTrademark(1) = "717" And Val(txtTrademark(18)) <> (2500# * (Val(UBound(TmpArr)) + 1)) Then
''                        MsgBox "國內商標全期註冊費規費要" & str(2500# * (Val(UBound(TmpArr)) + 1)) & ", 請確認！", vbCritical
'                If txtTrademark(1) = "717" And Val(txtTrademark(18)) <> (2500# * (Val(UBound(TmpArr)) + 1) * intMoney) Then
'                   MsgBox "台灣商標全期註冊費規費要" & str(2500# * (Val(UBound(TmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
'                   Cancel = True
'                   Exit Sub
'                End If
'                'add by nickc 2006/12/15
'                If txtTrademark(1) = "102" And Val(txtTrademark(18)) <> (4000# * (Val(UBound(TmpArr)) + 1) * intMoney) Then
'                   MsgBox "台灣商標延展規費要" & str(4000# * (Val(UBound(TmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
'                   Cancel = True
'                   Exit Sub
'                End If
'                '2007/8/30 ADD BY SONIA 加入異議,評定,廢止
'                If txtTrademark(1) = "601" And Val(txtTrademark(18)) <> (4000# * (Val(UBound(TmpArr)) + 1) * intMoney) Then
'                   MsgBox "台灣商標異議規費要" & str(4000# * (Val(UBound(TmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
'                   Cancel = True
'                   Exit Sub
'                End If
'                If (txtTrademark(1) = "603" Or txtTrademark(1) = "605") And Val(txtTrademark(18)) <> (7000# * (Val(UBound(TmpArr)) + 1) * intMoney) Then
'                   MsgBox "台灣商標評定或廢止規費要" & str(7000# * (Val(UBound(TmpArr)) + 1) * intMoney) & ", 請確認！", vbCritical
'                   Cancel = True
'                   Exit Sub
'                End If
'                '2007/8/30 END
'             End If
'             'Add By Sindy 2010/7/30
'             If CheckKeyIn(14) = -1 Then '檢查費用
'               Cancel = True
'               Exit Sub
'             End If
'          End If
      Case Else
'                Select Case Index
'                   Case 9, 10
'                      If Len(txtTrademark(Index).Text) = 6 Then
'                         txtTrademark(Index).Text = txtTrademark(Index).Text & "000"
'                      Else
'                         If Len(txtTrademark(Index).Text) = 8 Then
'                            txtTrademark(Index).Text = txtTrademark(Index).Text & "0"
'                         End If
'                      End If
'                End Select
           If CheckKeyIn(Index) = -1 Then
              Cancel = True
           End If
   End Select
EXITSUB:
   If Cancel Then txtTrademark_GotFocus (Index)
End Sub

Private Function CheckKeyIn(ByRef intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, bolIsChina As Boolean, strCusTemp As String
Static strLastCus As String
' 91.09.11 modify by louis
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strFee As String
Dim strBCase(1 To 4) As String 'Added by Lydia 2023/03/06
   
   'Added by Lydia 2023/03/06 抓本所案號
   If txtSystem.Text = 馬德里案 Then
       strBCase(1) = txtSystem
       strBCase(2) = txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1))
       strBCase(3) = IIf(txtTFCode(2) = "", "0", txtTFCode(2))
       strBCase(4) = IIf(txtTFCode(3) = "", "00", txtTFCode(3))
   Else
       strBCase(1) = txtSystem
       strBCase(2) = txtCode(0)
       strBCase(3) = IIf(txtCode(1) = "", "0", txtCode(1))
       strBCase(4) = IIf(txtCode(2) = "", "00", txtCode(2))
   End If
   'end 2023/03/06
   
   CheckKeyIn = -1
   Select Case intIndex
             Case 4
                       If CheckLengthIsOK(txtTrademark(intIndex), 395) Then
                          CheckKeyIn = 1
                       End If
             'add by nickc 2006/11/30
             Case 32
                       If CheckLengthIsOK(txtTrademark(intIndex), 349) Then
                          CheckKeyIn = 1
                       End If
             Case 6
'                       If CheckLengthIsOK(txtTrademark(intIndex), 40) Then
                       If CheckLengthIsOK(txtTrademark(intIndex), 140) Then
                          CheckKeyIn = 1
                       End If
             Case 7
                       If CheckLengthIsOK(txtTrademark(intIndex), 60) Then
                          CheckKeyIn = 1
                       End If
             Case 0
                        If CheckIsTaiwanDate(txtTrademark(intIndex).Text) Then
                            CheckKeyIn = 1
                        End If
             Case 1
                       If txtTrademark(5) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
                       'edit by nickc 2007/02/02 不用 dll 了
                       'If objPublicData.GetCaseProperty(txtSystem, txtTrademark(intIndex), strTemp, bolIsChina) Then
                       If ClsPDGetCaseProperty(txtSystem, txtTrademark(intIndex), strTemp, bolIsChina) Then
                           lblCaseProperty.Caption = strTemp
                           CheckKeyIn = 1
                       End If
                       Call SetTM136 'Add By Sindy 2022/12/7 證書形式
             Case 2
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCaseSource(txtTrademark(intIndex).Text, strTemp) Then
                        If ClsPDGetCaseSource(txtTrademark(intIndex).Text, strTemp) Then
                           lblCaseSource.Caption = strTemp
                           CheckKeyIn = 1
                        End If
             Case 3
                        If txtTrademark(5) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
'                        If objPublicData.GetPatentTrademarkKind(商標, txtTrademark(intIndex).Text, strTemp, bolIsChina) Then
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetPatentTrademarkKind(商標, txtTrademark(intIndex).Text, strTemp, bolIsChina) = 1 Then
                        If ClsPDGetPatentTrademarkKind(商標, txtTrademark(intIndex).Text, strTemp, bolIsChina) = 1 Then
                           lblTrademarkKind = strTemp
                           CheckKeyIn = 1
                        End If
             Case 5
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetNation(txtTrademark(intIndex).Text, strTemp) Then
                        If ClsPDGetNation(txtTrademark(intIndex).Text, strTemp) Then
                           lblNation.Caption = strTemp
                           CheckKeyIn = 1
                        End If
                        '91.10.25 add by sonia
                        If txtSystem = "FCT" And txtTrademark(intIndex).Text <> 台灣國家代號 Then
                           ShowMsg MsgText(9219)
                           CheckKeyIn = -1
                           Exit Function
                        End If
                        '91.10.25 END
                        If Val(txtTrademark(intIndex)) >= 1 And Val(txtTrademark(intIndex)) <= 8 Then
                           ShowMsg MsgText(38)
                           CheckKeyIn = -1
                           'add by nick 2004/12/07
                           Exit Function
                        End If
                        'add by nick 2004/12/07 不能輸入歐盟
                        If txtTrademark(intIndex).Text = "221" Then
                           ShowMsg MsgText(9220)
                           CheckKeyIn = -1
                           Exit Function
                        End If
                        'add end
                        Call setFrame21 'Add By Sindy 2012/7/23
                        Call SetTM136 'Add By Sindy 2022/12/7 證書形式
                        
                        'Added by Lydia 2020/11/19 CFT英國脫歐案管制
                        If txtSystem = "CFT" And frm010001.txtCaseNa239 <> "" And txtTrademark(5) <> "201" Then
                             'Added by Lydia 2021/03/05 CFT歐盟尚未註冊案轉換英國申請案收文控管：針對2021.9.30前收文之英國新「申請101」案建立關聯案
                             If txtTrademark(1) = "101" Then
                                 MsgBox "申請國家只可為英國！", vbCritical
                             Else
                             'end 2021/03/05
                                 MsgBox "英國脫歐案的申請國家只可為英國！", vbCritical
                             End If  'Added by Lydia 2021/03/05
                             CheckKeyIn = -1
                             Exit Function
                        End If
                        'end 2020/11/19
                        
                        'Add By Sindy 2013/8/23
                        '台灣新申請案電子送件
                        'Modified by Lydia 2019/08/01 開放FCT舊案可勾選電子送件
                        'If txtSystem = "FCT" And txtTrademark(5) = "000" And txtTrademark(1) = "101" Then
                        If txtSystem = "FCT" And txtTrademark(5) = "000" Then
                           chkWebApp.Visible = True
                        Else
                           chkWebApp.Visible = False
                           chkWebApp.Value = 0
                        End If
                        '2013/8/23 END
             Case 8
                        If txtTrademark(6) = "" And txtTrademark(7) = "" And txtTrademark(8) = "" Then
                           ShowMsg MsgText(1031)
                           intIndex = 6
                           CheckKeyIn = 0
                        ElseIf CheckLengthIsOK(txtTrademark(intIndex), 40) Then
                           CheckKeyIn = 1
                        End If
             Case 9 '申請人
                        strCusTemp = txtTrademark(intIndex)
                        'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
                        'If objPublicData.GetCustomer(strCusTemp, strTemp, strTemp1) Then
                        'Modify By Sindy 2015/8/27 +txtSystem
                        'Modify By Sindy 2021/2/1 + , strXState(9), IIf(frm010001.intSaveMode = 0, True, False)
                        'Modified by Lydia 2023/03/06 傳入本所案號 , , strBCase(2), strBCase(3), strBCase(4)
                        If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, strXState(9), IIf(frm010001.intSaveMode = 0, True, False), , strBCase(2), strBCase(3), strBCase(4)) Then
                           txtTrademark(intIndex) = strCusTemp
                           lblPetition(0).Caption = strTemp
                           If strLastCus <> strCusTemp Or txtTrademark(13).Text = "" Then
                              txtTrademark(13).Text = strTemp1
                              strLastCus = strCusTemp
                           End If
                           CheckKeyIn = 1
                           'Add by Morgan 2008/8/5
                           If ChangeCustomerL(strCusTemp) <> strAppNo1 Then
                              strAppNo1 = ChangeCustomerL(strCusTemp)
                              'Modify by Amy 2021/12/20 改成Form 2.0
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
                        If CheckKeyIn = 1 Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetCustomerNation(strCusTemp, strNation) Then
                           If ClsPDGetCustomerNation(strCusTemp, strNation) Then
                              'If strNation >= "010" Then
                              '   txtTrademark(17) = "N"
                              'Else
                              '   txtTrademark(17) = ""
                              'End If
                           End If
                        End If
                        'Add By Cheng 2003/09/08
                        If CheckKeyIn = 1 Then
                            '2010/9/30 modify by sonia 新增時才要檢查
                            'If frm010001.m_blnNewCase = True Then
                            If frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 Then
                                '若輸入9碼且最後一碼不為"0"
                                If Len(Me.txtTrademark(intIndex).Text) = 9 And Right(Me.txtTrademark(intIndex).Text, 1) <> "0" Then
                                   'Added by Lydia 2024/02/16 商標案件FCT、T、CFT、TF的分割308時，改為彈訊息並可選擇是或否？
                                   If (txtSystem = "FCT" Or txtSystem = "CFT" Or txtSystem = "T" Or txtSystem = "TF") And txtTrademark(1) = "308" Then
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
             'add by nickc 2006/11/22
             Case 24, 25, 26, 27 '申請人
                        If txtTrademark(intIndex) = "" Then
                           CheckKeyIn = 1
                           Exit Function
                        End If
                        If intIndex = 24 Then
                           If txtTrademark(intIndex) = txtTrademark(9) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(25) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(26) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(27) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        If intIndex = 25 Then
                           If txtTrademark(intIndex) = txtTrademark(9) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(24) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(26) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(27) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        If intIndex = 26 Then
                           If txtTrademark(intIndex) = txtTrademark(9) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(24) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(25) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(27) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        If intIndex = 27 Then
                           If txtTrademark(intIndex) = txtTrademark(9) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(24) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(25) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtTrademark(intIndex) = txtTrademark(26) Then
                              ShowMsg "申請人不可重複"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        If txtTrademark(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           strCusTemp = txtTrademark(intIndex)
                           'Modify By Sindy 2015/8/27 +txtSystem
                           'Modify By Sindy 2021/2/1 + , strXState(intIndex), IIf(frm010001.intSaveMode = 0, True, False)
                           'Modified by Lydia 2023/03/06 傳入本所案號  , , strBCase(2), strBCase(3), strBCase(4)
                           If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, strXState(intIndex), IIf(frm010001.intSaveMode = 0, True, False), , strBCase(2), strBCase(3), strBCase(4)) Then
                              txtTrademark(intIndex) = strCusTemp
                              lblPetition(intIndex - 23).Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        End If
                        If CheckKeyIn = 1 Then
                            '2010/9/30 modify by sonia 新增時才要檢查
                            'If frm010001.m_blnNewCase = True Then
                            If frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 Then
                                If Len(Me.txtTrademark(intIndex).Text) = 9 And Right(Me.txtTrademark(intIndex).Text, 1) <> "0" Then
                                   'Added by Lydia 2024/02/16 商標案件FCT、T、CFT、TF的分割308時，改為彈訊息並可選擇是或否？
                                   If (txtSystem = "FCT" Or txtSystem = "CFT" Or txtSystem = "T" Or txtSystem = "TF") And txtTrademark(1) = "308" Then
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
            Case 28, 29, 30, 31
                        If txtTrademark(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           strCusTemp = txtTrademark(intIndex)
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetCustomer(strCusTemp, strTemp) Then
                           If ClsPDGetCustomer(strCusTemp, strTemp) Then
                              txtTrademark(intIndex) = strCusTemp
                              lblPetitionName(intIndex - 27).Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        End If
             Case 10 '代理人
                        strCusTemp = txtTrademark(intIndex)
                        If txtTrademark(intIndex) <> "" Then
                           'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
                           'If objPublicData.GetAgent(strCusTemp, strTemp) Then
                           'Modify By Sindy 2015/8/27 +txtSystem
                           'Modify By Sindy 2021/2/1 + , strYState, IIf(frm010001.intSaveMode = 0, True, False)
                           If GetAgentAndState(strCusTemp, strTemp, , , , txtSystem, strYState, IIf(frm010001.intSaveMode = 0, True, False)) Then
                              txtTrademark(intIndex) = strCusTemp
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
                                If Len(Me.txtTrademark(intIndex).Text) = 9 And Right(Me.txtTrademark(intIndex).Text, 1) <> "0" Then
                                    MsgBox "此代理人已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
                                    CheckKeyIn = -1
                                End If
                            End If
                        End If
             Case 11
                        If txtTrademark(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           If CheckIsTaiwanDate(txtTrademark(intIndex).Text) Then
                              If CheckReKey(txtTrademark(intIndex)) Then
'                                 If txtTrademark(intIndex) = GetTaiwanTodayDate Then
'                                    ShowMsg "此案件已屆本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
'                                 End If
'                                 If txtTrademark(intIndex) < GetTaiwanTodayDate Then
'                                    ShowMsg "此案件已逾本所期限, 請立刻提醒專業部門作業! 分所請傳真接洽單至北所! "
'                                 End If
                                 CheckKeyIn = 1
                                 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                                  txtTrademark(intIndex).Text = TransDate(PUB_GetWorkDay1(txtTrademark(intIndex).Text, True), 1)
                              Else
                                 CheckKeyIn = 0
                              End If
                            End If
                        End If
             Case 16
                        If txtTrademark(intIndex) <> "" Then
                           If CheckIsTaiwanDate(txtTrademark(intIndex).Text) Then
                              If Val(txtTrademark(11)) <= Val(txtTrademark(16)) Then
                                 If CheckReKey(txtTrademark(intIndex)) Then
                                    CheckKeyIn = 1
                                 Else
                                    CheckKeyIn = 0
                                 End If
                              Else
                                 ShowMsg MsgText(1033)
                              End If
                           End If
                        ElseIf txtTrademark(11) <> "" Then
                           ShowMsg MsgText(1033)
                           CheckKeyIn = 0
                        Else
                           CheckKeyIn = 1
                        End If
             Case 12
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetStaff(txtTrademark(intIndex).Text, strTemp, strTemp1) Then
                        If ClsPDGetStaff(txtTrademark(intIndex).Text, strTemp, strTemp1) Then
                           CheckKeyIn = 1
                        End If
                        lblSales.Caption = strTemp
                        
                        'Modified by Lydia 2019/02/14
                        'strTemp = GetST15(txtTrademark(intIndex).Text, strTemp1)
                        'Modified by Lydia 2019/09/16 +st06
                        'm_SalesST15 = GetST15(txtTrademark(intIndex).Text, strTemp1)
                        m_SalesST15 = GetST15(txtTrademark(intIndex).Text, strTemp1, , m_SalesST06)
                        lblDepartment = strTemp1
'             Case 14
'                        'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'                        If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" Then
'                            'edit by nickc 2007/02/02 不用 dll 了
'                            'If objPublicData.GetCaseLowPrice(txtSystem, txtTrademark(5), txtTrademark(1), douStPrice, douLowPrice) = 1 Then
'                            If ClsPDGetCaseLowPrice(txtSystem, txtTrademark(5), txtTrademark(1), douStPrice, douLowPrice) = 1 Then
'                            End If
'                            If txtTrademark(intIndex) <> "" Then
'                               'edit by nickc 2007/02/02 不用 dll 了
'                               'If objPublicData.GetCaseFee(txtSystem, txtTrademark(5), txtTrademark(1), Val(txtTrademark(intIndex))) = 1 Then
'                               'Add By Sindy 2010/7/30
'                               If txtTrademark(5) = "000" And _
'                                 (Trim(txtTrademark(1)) = "101" Or _
'                                  Trim(txtTrademark(1)) = "102" Or _
'                                  Trim(txtTrademark(1)) = "715" Or _
'                                  Trim(txtTrademark(1)) = "716" Or _
'                                  Trim(txtTrademark(1)) = "717" Or _
'                                  Trim(txtTrademark(1)) = "601" Or _
'                                  Trim(txtTrademark(1)) = "603" Or _
'                                  Trim(txtTrademark(1)) = "605") Then
'                                 If Val(txtTrademark(18)) > 0 And txtTrademark(4) <> "" Then
'                                    Dim TmpArr As Variant
'                                    TmpArr = Split(IIf(Right(txtTrademark(4), 1) = ",", Mid(txtTrademark(4), 1, Len(txtTrademark(4)) - 1), txtTrademark(4)), ",")
'                                    If ClsPDGetCaseFee_T(txtSystem, txtTrademark(5), txtTrademark(1), Val(txtTrademark(14)), Val(txtTrademark(18)), Val(UBound(TmpArr))) = 1 Then
'                                       CheckKeyIn = 1
'                                    End If
'                                 Else
'                                    CheckKeyIn = 1
'                                 End If
'                               '2010/7/30 End
'                               Else
'                                 If ClsPDGetCaseFee(txtSystem, txtTrademark(5), txtTrademark(1), Val(txtTrademark(intIndex))) = 1 Then
'                                    CheckKeyIn = 1
'                                 End If
'                               End If
'                            'ElseIf txtTrademark(18) <> "" Then
'                            '   ShowMsg MsgText(1034)
'                            '   CheckKeyIn = 0
'                            Else
'                               CheckKeyIn = 1
'                            End If
'                        Else
'                            CheckKeyIn = 1
'                        End If
'             Case 15
'                    'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'                    If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" Then
'                        If txtTrademark(intIndex) = "" Then
'                           If txtTrademark(14) <> "" Or txtTrademark(18) <> "" Then
'                              ShowMsg MsgText(1035)
'                              CheckKeyIn = 0
'                           Else
'                              CheckKeyIn = 1
'                           End If
'                        ElseIf txtTrademark(14) <> "" Or txtTrademark(18) <> "" Then
'                           If Format((Val(txtTrademark(14)) - Val(txtTrademark(18))) / 1000, "0.0") <> Format(Val(txtTrademark(15)), "0.0") Then
''                              ShowMsg MsgText(1036)
'                              CheckKeyIn = 0
'                           Else
'                              CheckKeyIn = 1
'                           End If
'                        Else
'                           ShowMsg MsgText(1037)
'                        End If
'                    Else
'                        CheckKeyIn = 1
'                    End If
             Case 17
                        'If strNation >= "010" Then
                        '   If txtTrademark(17) <> "N" Then
                        '      ShowMsg "申請人國籍非台灣時, 是否開電腦收據必須為 N"
                        '      CheckKeyIn = -1
                        '      Exit Function
                        '   End If
                        'End If
                        If txtTrademark(intIndex) = "" Or txtTrademark(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
'             Case 18 '規費
'                    'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'                    If Mid(txtTrademark(9), 1, 8) <> "X1484305" And Mid(txtTrademark(24), 1, 8) <> "X1484305" And Mid(txtTrademark(25), 1, 8) <> "X1484305" And Mid(txtTrademark(26), 1, 8) <> "X1484305" And Mid(txtTrademark(27), 1, 8) <> "X1484305" Then
'                        '91.10.24 modify by sonia
'                        ' 91.09.11 modify by louis
'                        'strFee = GetTrademarkOfficialFee(txtSystem, txtTrademark(1), txtTrademark(16))
'                        If txtTrademark(5) = "000" Then
'                           strFee = GetTrademarkOfficialFee(txtSystem, txtTrademark(1), txtTrademark(16))
'                        End If
'                        '91.10.24 end
'
'                        If Val(strFee) > 0 Then
'                           If Val(txtTrademark(18)) <> Val(strFee) Then
'                              strTit = "檢核資料"
'                              strMsg = "規費應為<" & strFee & ">"
'                              nResponse = MsgBox(strMsg, vbOKCancel + vbCritical, strTit)
'                              GoTo EXITSUB
'                           End If
'                        End If
'
'                        '2009/6/12 modify by sonia取消有輸入才檢查的限制
'                        'If txtTrademark(intIndex) <> "" Then   '2009/6/12 cancel by sonia
'                           '91.11.21 CANCEL BY SONIA
'                           'If txtTrademark(14) = "" Then
'                           '   ShowMsg MsgText(1039)
'                           'ElseIf objPublicData.GetCaseMoney(txtSystem, txtTrademark(5), txtTrademark(1), Val(txtTrademark(intIndex))) = 1 Then
'                           '91.11.21 END
'                           'edit by nickc 2006/12/05 change call basquery
'                           'If objPublicData.GetCaseMoney(txtSystem, txtTrademark(5), txtTrademark(1), Val(txtTrademark(intIndex))) = 1 Then
'                           If GetCaseMoney(txtSystem, txtTrademark(5), txtTrademark(1), Val(txtTrademark(intIndex))) = 1 Then
'                              CheckKeyIn = 1
'                           End If
'                        'Else                                   '2009/6/12 cancel by sonia
''                           If txtTrademark(14) <> "" Then
''                              ShowMsg MsgText(1040)
''                              CheckKeyIn = 0
''                           Else
'                        '      CheckKeyIn = 1                   '2009/6/12 cancel by sonia
''                           End If
'                        'End If                                 '2009/6/12 cancel by sonia
'
'                        'Add By Cheng 2003/11/19
'                        'T及FCT第二期註冊費收文時, 若收文日大於法定期限時, 則控制規費加倍
'                        If CheckKeyIn <> -1 Then
'                            CheckKeyIn = 1
'                            If (Me.txtSystem.Text = "T" Or Me.txtSystem.Text = "FCT") And Me.txtTrademark(1).Text = "716" And (Me.txtTrademark(0).Text <> "" And Me.txtTrademark(16).Text <> "") And (Val(Me.txtTrademark(0).Text) > Val(Me.txtTrademark(16).Text)) Then
'                                strFee = Val(GetOfficalFee(Me.txtSystem.Text, Me.txtTrademark(5).Text, Me.txtTrademark(1).Text)) * 2
'                                If Val(Me.txtTrademark(intIndex).Text) <> Val(strFee) Then
'                                    MsgBox "規費應為<" & strFee & ">", vbExclamation + vbOKOnly
'                                    CheckKeyIn = -1
'                                End If
'                            End If
'                        End If
'                    Else
'                        CheckKeyIn = 1
'                    End If
             Case 20
                        If txtTrademark(intIndex) <> "" Then
                           strCusTemp = txtTrademark(intIndex)
                           'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
                           'If objPublicData.GetCustomer(strCusTemp, strTemp, strTemp1) Then
                           'Modify By Sindy 2015/8/27 +txtSystem
                           'Modify By Sindy 2021/2/1 + , strXState(20), IIf(frm010001.intSaveMode = 0, True, False)
                           'Modified by Lydia 2023/03/06 傳入本所案號  , , strBCase(2), strBCase(3), strBCase(4)
                           If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, strXState(20), IIf(frm010001.intSaveMode = 0, True, False), , strBCase(2), strBCase(3), strBCase(4)) Then
                              txtTrademark(intIndex) = strCusTemp
                              'edit by nickc 2006/11/22
                              'lblPetitionName.Caption = strTemp
                              lblPetitionName(0).Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
             Case 21
                        If txtTrademark(intIndex) <> "" Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetStaff(txtTrademark(intIndex), strTemp) Then
                           If ClsPDGetStaff(txtTrademark(intIndex), strTemp) Then
                              lblPromoter = strTemp
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
             'add by nickc 2005/10/06 加長分所號
             Case 23
                       If CheckLengthIsOK(txtTrademark(intIndex), 50) Then
                          CheckKeyIn = 1
                       End If
             Case 28
                        If txtTrademark(intIndex) <> "" Then
                           strCusTemp = txtTrademark(intIndex)
                           'Modify By Sindy 2015/8/27 +txtSystem
                           'Modify By Sindy 2021/2/1 + , strXState(28), IIf(frm010001.intSaveMode = 0, True, False)
                           'Modified by Lydia 2023/03/06 傳入本所案號  , , strBCase(2), strBCase(3), strBCase(4)
                           If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, strXState(28), IIf(frm010001.intSaveMode = 0, True, False), , strBCase(2), strBCase(3), strBCase(4)) Then
                              txtTrademark(intIndex) = strCusTemp
                              lblPetitionName(1).Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
             Case 29
                        If txtTrademark(intIndex) <> "" Then
                           strCusTemp = txtTrademark(intIndex)
                           'Modify By Sindy 2015/8/27 +txtSystem
                           'Modify By Sindy 2021/2/1 + , strXState(29), IIf(frm010001.intSaveMode = 0, True, False)
                           'Modified by Lydia 2023/03/06 傳入本所案號  , , strBCase(2), strBCase(3), strBCase(4)
                           If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, strXState(29), IIf(frm010001.intSaveMode = 0, True, False), , strBCase(2), strBCase(3), strBCase(4)) Then
                              txtTrademark(intIndex) = strCusTemp
                              lblPetitionName(2).Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
             Case 30
                        If txtTrademark(intIndex) <> "" Then
                           strCusTemp = txtTrademark(intIndex)
                           'Modify By Sindy 2015/8/27 +txtSystem
                           'Modify By Sindy 2021/2/1 + , strXState(30), IIf(frm010001.intSaveMode = 0, True, False)
                           'Modified by Lydia 2023/03/06 傳入本所案號  , , strBCase(2), strBCase(3), strBCase(4)
                           If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, strXState(30), IIf(frm010001.intSaveMode = 0, True, False), , strBCase(2), strBCase(3), strBCase(4)) Then
                              txtTrademark(intIndex) = strCusTemp
                              lblPetitionName(3).Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
             Case 31
                        If txtTrademark(intIndex) <> "" Then
                           strCusTemp = txtTrademark(intIndex)
                           'Modify By Sindy 2015/8/27 +txtSystem
                           'Modify By Sindy 2021/2/1 + , strXState(31), IIf(frm010001.intSaveMode = 0, True, False)
                           'Modified by Lydia 2023/03/06 傳入本所案號  , , strBCase(2), strBCase(3), strBCase(4)
                           If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, strXState(31), IIf(frm010001.intSaveMode = 0, True, False), , strBCase(2), strBCase(3), strBCase(4)) Then
                              txtTrademark(intIndex) = strCusTemp
                              lblPetitionName(4).Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
            'add by nickc 2008/05/02 加預定收款日
             'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'             Case 34
'                        If txtTrademark(intIndex) = "" Then
'                           CheckKeyIn = 1
'                        Else
'                           If CheckIsTaiwanDate(txtTrademark(intIndex).Text) Then
'                                'edit by nickc 2008/05/21 原先判斷系統日，改成判斷收文日，因為分所會去補分所號
'                                'If DBDATE(txtTrademark(intIndex).Text) >= strSrvDate(1) Then
'                                If DBDATE(txtTrademark(intIndex).Text) >= DBDATE(txtTrademark(0).Text) Then
'                                   CheckKeyIn = 1
'                                Else
'                                    'edit by nickc 2008/05/21 原先判斷系統日，改成判斷收文日，因為分所會去補分所號
'                                    'MsgBox "預定收款日必須>= 系統日", vbOKOnly + vbCritical, "輸入錯誤！"
'                                    MsgBox "預定收款日必須>= 收文日", vbOKOnly + vbCritical, "輸入錯誤！"
'                                End If
'                           End If
'                        End If
             'end 2018/08/22
             Case Else
                        CheckKeyIn = 1
   End Select
EXITSUB:
End Function

Private Sub txtTrademark_GotFocus(Index As Integer)
   If Index = 9 Then
   '   If txtTrademark(6) = "" And txtTrademark(7) = "" And txtTrademark(8) = "" Then
      If txtTrademark(6) = "" Then
         txtTrademark(6).SetFocus
         Exit Sub
      End If
   End If
   txtTrademark(Index).SelStart = 0
   txtTrademark(Index).SelLength = Len(txtTrademark(Index).Text)
   '儲存未修改前之值至Tag中,供再確認時使用
   txtTrademark(Index).Tag = txtTrademark(Index)
   '切換輸入法
   Select Case Index
             'Case 4 'Removed by Morgan 2016/10/20
'                 cmdOK(0).CausesValidation = False
'                 cmdOK(0).Default = False
             Case 6
                        'edit by nickc 2007/06/06 切換輸入法改用API
                        'txtTrademark(Index).IMEMode = 1
                        OpenIme
             Case Else
                        'edit by nickc 2007/06/06 切換輸入法改用API
                        'txtTrademark(Index).IMEMode = 2
                        CloseIme
   End Select
End Sub

Private Sub txtTrademark_LostFocus(Index As Integer)
   '關閉輸入法
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtTrademark(Index).IMEMode = 2
   'CloseIme 'Removed by Morgan 2016/10/20 會造成 Win7 的切換錯誤
   'Add By Cheng 2001/12/27
   If Index = 10 And Len(Trim(Me.txtTrademark(9).Text)) <= 0 And Len(Trim(Me.txtTrademark(10).Text)) <= 0 Then
      '若申請人與代理人皆未輸入, 則將游標設定在申請人欄位
      Me.txtTrademark(9).SetFocus
   End If
   'Add By Sindy 2013/8/26
   If Index = 0 Then '收文日
      'Modify By Sindy 2019/8/12 + And txtTrademark(1) = "101"
      'Remove by Lydia 2020/03/12 請取消商申電子送件有關當日期限的管制. 若該案須設期限管制時, 由承辦人載入接洽單, 於收文時輸入即可.
      'If chkWebApp.Visible = True And txtTrademark(1) = "101" Then
      '   If chkWebApp.Value = 1 Then
      '      txtTrademark(11) = txtTrademark(0)
      '      txtTrademark(16) = txtTrademark(0)
      '   Else
      '      txtTrademark(11) = ""
      '      txtTrademark(16) = ""
      '   End If
      'End If
      'end 2020/03/12
   End If
   '2013/8/26 END
End Sub

'修改商標資料庫
'edit by nickc 2006/11/22
'Private Function UpdateTrademarkDatabase(ByRef tm01 As String, _
             ByRef tm02 As String, ByRef tm03 As String, ByRef tm04 As String, ByRef tm05 As String, _
             ByRef tm08 As String, ByRef tm09 As String, _
             ByRef tm10 As String, ByRef tm23 As String, ByRef tm44 As String, _
             ByRef cp09 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, _
             ByRef cp11 As String, ByRef cp13 As String, ByRef cp14 As String, ByRef cp16 As String, _
             ByRef cp17 As String, ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, _
             ByRef cp56 As String, ByRef cu30 As String, ByRef cp33 As Double, ByRef cp34 As Double, ByRef CP64 As String) As Boolean
'edit by nickc 2007/03/27 加入彼所案號
'Private Function UpdateTrademarkDatabase(ByRef tm01 As String, _
             ByRef tm02 As String, ByRef tm03 As String, ByRef tm04 As String, ByRef tm05 As String, _
             ByRef tm08 As String, ByRef tm09 As String, _
             ByRef tm10 As String, ByRef tm23 As String, ByRef tm44 As String, _
             ByRef CP09 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, _
             ByRef cp11 As String, ByRef cp13 As String, ByRef cp14 As String, ByRef cp16 As String, _
             ByRef cp17 As String, ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, _
             ByRef cp56 As String, ByRef cu30 As String, ByRef cp33 As Double, ByRef cp34 As Double, ByRef CP64 As String, ByRef TM78 As String, ByRef TM79 As String, ByRef TM80 As String, ByRef TM81 As String, ByRef CP89 As String, ByRef CP90 As String, ByRef CP91 As String, ByRef CP92 As String, ByRef TM32 As String) As Boolean

Private Function UpdateTrademarkDatabase(ByRef tm01 As String, _
             ByRef tm02 As String, ByRef tm03 As String, ByRef tm04 As String, ByRef tm05 As String, _
             ByRef tm08 As String, ByRef tm09 As String, _
             ByRef tm10 As String, ByRef tm23 As String, ByRef tm44 As String, _
             ByRef CP09 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, _
             ByRef cp11 As String, ByRef cp13 As String, ByRef cp14 As String, ByRef cp16 As String, _
             ByRef cp17 As String, ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, _
             ByRef cp56 As String, ByRef cu30 As String, ByRef cp33 As Double, ByRef cp34 As Double, ByRef CP64 As String, ByRef TM78 As String, ByRef TM79 As String, ByRef TM80 As String, ByRef TM81 As String, ByRef CP89 As String, ByRef CP90 As String, ByRef CP91 As String, ByRef CP92 As String, ByRef TM32 As String, ByRef TM45 As String, Optional ByRef TM123 As String) As Boolean
Dim strSql As String, cp55 As String
Dim adoquery As New ADODB.Recordset
Dim cp93 As String, cp94 As String, cp95 As String, cp96 As String 'Add by Nickc 2006/11/22

   'add by nickc 2007/12/12
   If IsSaveData = True Then
       Exit Function
   End If
   IsSaveData = True
   
   On Error GoTo ErrHand
   cp05 = ChangeTStringToWString(cp05)
   cp06 = ChangeTStringToWString(cp06)
   cp07 = ChangeTStringToWString(cp07)
   tm23 = ChangeCustomerL(tm23)
   tm44 = ChangeCustomerL(tm44)
   'add by nickc 2007/11/26
   TM78 = ChangeCustomerL(TM78)
   TM79 = ChangeCustomerL(TM79)
   TM80 = ChangeCustomerL(TM80)
   TM81 = ChangeCustomerL(TM81)
   
   cp56 = ChangeCustomerL(cp56)
   
   cnnConnection.BeginTrans
   'strSQL = "update trademark set tm05=" + CNULL(ChgSQL(tm05)) + ",tm06=" + CNULL(Replace(tm06, "'", "''")) + _
   '   ",tm07=" + CNULL(ChgSQL(tm07)) + ",tm08=" + CNULL(tm08) + ",tm09=" + CNULL(tm09) + ",tm10=" + _
   '   CNULL(tm10) + ",tm23=" + CNULL(tm23) + ",tm44=" + CNULL(tm44) + " where tm01=" + _
   '   CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   'edit by nickc 2006/11/22
   'strSQL = "update trademark set tm05=" + CNULL(ChgSQL(tm05)) + _
      ",tm08=" + CNULL(tm08) + ",tm09=" + CNULL(tm09) + ",tm10=" + _
      CNULL(tm10) + ",tm23=" + CNULL(tm23) + ",tm44=" + CNULL(tm44) + " where tm01=" + _
      CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   'edit by nickc 2007/03/27 加入彼所案號
   'strSQL = "update trademark set tm05=" + CNULL(ChgSQL(tm05)) + _
      ",tm08=" + CNULL(tm08) + ",tm09=" + CNULL(tm09) + ",tm10=" + _
      CNULL(tm10) + ",tm23=" + CNULL(tm23) + ",tm44=" + CNULL(tm44) + ",tm78=" + CNULL(TM78) + ",tm79=" + CNULL(TM79) + ",tm80=" + CNULL(TM80) + ",tm81=" + CNULL(TM81) + ",tm32=" + CNULL(TM32) + " where tm01=" + _
      CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   strSql = "update trademark set tm05=" + CNULL(ChgSQL(tm05)) + _
      ",tm08=" + CNULL(tm08) + ",tm09=" + CNULL(tm09) + ",tm10=" + _
      CNULL(tm10) + ",tm23=" + CNULL(tm23) + ",tm44=" + CNULL(tm44) + ",tm78=" + CNULL(TM78) + ",tm79=" + CNULL(TM79) + ",tm80=" + CNULL(TM80) + ",tm81=" + CNULL(TM81) + ",tm32=" + CNULL(TM32) + ",tm45=" + CNULL(ChgSQL(TM45))
   'Add by Morgan 2008/8/5 +TM123
   If UCase(TM123) <> "TM123" Then
      strSql = strSql + ",tm123=" + CNULL(TM123)
   End If
   strSql = strSql & " where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   cnnConnection.Execute strSql
   
   'Add By Sindy 2012/7/19 若申請人或代理人為諾華公司者，案件備註若無"不銷卷"字樣,則要加入
   If (tm23 <> "" And InStr(strTmNovartisCust, Left(tm23, 6)) > 0) Or _
      (TM78 <> "" And InStr(strTmNovartisCust, Left(TM78, 6)) > 0) Or _
      (TM79 <> "" And InStr(strTmNovartisCust, Left(TM79, 6)) > 0) Or _
      (TM80 <> "" And InStr(strTmNovartisCust, Left(TM80, 6)) > 0) Or _
      (TM81 <> "" And InStr(strTmNovartisCust, Left(TM81, 6)) > 0) Or _
      (tm44 <> "" And InStr(strTmNovartisCust, Left(tm44, 6)) > 0) Then
      strSql = "update trademark" & _
               " set tm58=decode(tm58,null,'" & ChangeTStringToTDateString(strSrvDate(2)) & "不銷卷','" & ChangeTStringToTDateString(strSrvDate(2)) & "不銷卷,'||tm58)" & _
               " Where tm01='" & tm01 & "' and tm02='" & tm02 & "' and tm03='" & tm03 & "' and tm04='" & tm04 & "'" & _
               " and (instr(tm58,'不銷卷')=0 or tm58 is null)"
      cnnConnection.Execute strSql
   End If
   '2012/7/19 end
   'ADD BY SONIA 2015/11/24 若申請為旅狐國際及部分關係企業者，案件備註若無"不銷卷"字樣,則要加入
   If (tm23 <> "" And InStr(strTmTRAVEL_FOXCust, Left(tm23, 8)) > 0) Or _
      (TM78 <> "" And InStr(strTmTRAVEL_FOXCust, Left(TM78, 8)) > 0) Or _
      (TM79 <> "" And InStr(strTmTRAVEL_FOXCust, Left(TM79, 8)) > 0) Or _
      (TM80 <> "" And InStr(strTmTRAVEL_FOXCust, Left(TM80, 8)) > 0) Or _
      (TM81 <> "" And InStr(strTmTRAVEL_FOXCust, Left(TM81, 8)) > 0) Then
      strSql = "update trademark" & _
               " set tm58=decode(tm58,null,'" & ChangeTStringToTDateString(strSrvDate(2)) & "不銷卷','" & ChangeTStringToTDateString(strSrvDate(2)) & "不銷卷,'||tm58)" & _
               " Where tm01='" & tm01 & "' and tm02='" & tm02 & "' and tm03='" & tm03 & "' and tm04='" & tm04 & "'" & _
               " and (instr(tm58,'不銷卷')=0 or tm58 is null)"
      cnnConnection.Execute strSql
   End If
   ''END 2015/11/24
   
   strSql = "update trademark set tm24=(select cu23 from customer where cu01=" + CNULL(Mid(tm23, 1, 8)) + " and cu02=" + CNULL(Mid(tm23, 9, 1)) + _
      "),tm25=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(tm23, 1, 8)) + " and cu02=" + CNULL(Mid(tm23, 9, 1)) + _
      "),tm26=(select cu29 from customer where cu01=" + CNULL(Mid(tm23, 1, 8)) + " and cu02=" + CNULL(Mid(tm23, 9, 1)) + ") where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   cnnConnection.Execute strSql
   'add by nickc 2006/11/30
   strSql = "update trademark set tm82=(select cu23 from customer where cu01=" + CNULL(Mid(TM78, 1, 8)) + " and cu02=" + CNULL(Mid(TM78, 9, 1)) + _
      "),tm86=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(TM78, 1, 8)) + " and cu02=" + CNULL(Mid(TM78, 9, 1)) + _
      "),tm90=(select cu29 from customer where cu01=" + CNULL(Mid(TM78, 1, 8)) + " and cu02=" + CNULL(Mid(TM78, 9, 1)) + ") where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   cnnConnection.Execute strSql
   strSql = "update trademark set tm83=(select cu23 from customer where cu01=" + CNULL(Mid(TM79, 1, 8)) + " and cu02=" + CNULL(Mid(TM79, 9, 1)) + _
      "),tm87=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(TM79, 1, 8)) + " and cu02=" + CNULL(Mid(TM79, 9, 1)) + _
      "),tm91=(select cu29 from customer where cu01=" + CNULL(Mid(TM79, 1, 8)) + " and cu02=" + CNULL(Mid(TM79, 9, 1)) + ") where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   cnnConnection.Execute strSql
   strSql = "update trademark set tm84=(select cu23 from customer where cu01=" + CNULL(Mid(TM80, 1, 8)) + " and cu02=" + CNULL(Mid(TM80, 9, 1)) + _
      "),tm88=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(TM80, 1, 8)) + " and cu02=" + CNULL(Mid(TM80, 9, 1)) + _
      "),tm92=(select cu29 from customer where cu01=" + CNULL(Mid(TM80, 1, 8)) + " and cu02=" + CNULL(Mid(TM80, 9, 1)) + ") where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   cnnConnection.Execute strSql
   strSql = "update trademark set tm85=(select cu23 from customer where cu01=" + CNULL(Mid(TM81, 1, 8)) + " and cu02=" + CNULL(Mid(TM81, 9, 1)) + _
      "),tm89=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(TM81, 1, 8)) + " and cu02=" + CNULL(Mid(TM81, 9, 1)) + _
      "),tm93=(select cu29 from customer where cu01=" + CNULL(Mid(TM81, 1, 8)) + " and cu02=" + CNULL(Mid(TM81, 9, 1)) + ") where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   cnnConnection.Execute strSql
   
   
   'Add By Cheng 2003/08/28
   'Begin
   strSql = "Update Trademark Set TM34='" & ChgSQL(Me.txtTrademark(23).Text) & "', TM35='" & ChgSQL(Me.txtTrademark(22).Text) & "' Where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   cnnConnection.Execute strSql
   'End
   If cp56 <> "" Then
   '    cp55 = tm23
      'Add by nickc 2006/11/22
      '讓與人2-5,受讓人2-5
      cp56 = ChangeCustomerL(cp56)
      CP89 = ChangeCustomerL(CP89)
      CP90 = ChangeCustomerL(CP90)
      CP91 = ChangeCustomerL(CP91)
      CP92 = ChangeCustomerL(CP92)
   '   cp93 = TM78
   '   cp94 = TM79
   '   cp95 = TM80
   '   cp96 = TM81
      'end 2006/6/23
   End If
   'edit by nickc 2006/11/30
   'strSQL = "update caseprogress set cp05=" + CNULL(cp05) + ",cp06=" + CNULL(cp06) + ",cp07=" + CNULL(cp07) + ",cp10=" + CNULL(CP10) + _
            ",cp11=" + CNULL(cp11) + ",cp13=" + CNULL(cp13) + ",cp14=" + CNULL(cp14) + ",cp16=" + CNULL(cp16) + ",cp17=" + CNULL(cp17) + _
            ",cp18=" + CNULL(cp18) + ",cp19=" + CNULL(cp19) + ",cp32=" + CNULL(cp32) + ",cp55=" + CNULL(cp55) + ",cp56=" + CNULL(cp56) + ",cp33=" & cp33 & ",cp34=" & cp34 & ",CP64=" + CNULL(ChgSQL(CP64)) + " where cp09='" + cp09 + "'"
   'cnnConnection.Execute strSQL
   
   'Modify By Sindy 2009/10/19
   'strSQL = "update caseprogress set cp05=" + CNULL(cp05) + ",cp06=" + CNULL(cp06) + ",cp07=" + CNULL(cp07) + ",cp10=" + CNULL(CP10) + _
   '         ",cp11=" + CNULL(cp11) + ",cp13=" + CNULL(cp13) + ",cp14=" + CNULL(cp14) + ",cp16=" + CNULL(cp16) + ",cp17=" + CNULL(cp17) + _
   '         ",cp18=" + CNULL(cp18) + ",cp19=" + CNULL(cp19) + ",cp32=" + CNULL(cp32) + ",cp55=" + CNULL(cp55) + ",cp56=" + CNULL(cp56) + ",cp33=" & cp33 & ",cp34=" & cp34 & ",CP64=" + CNULL(ChgSQL(CP64)) + _
   '         ",cp89=" + CNULL(CP89) + ",cp90=" + CNULL(CP90) + ",cp91=" + CNULL(CP91) + ",cp92=" + CNULL(CP92) + ",cp93=" + CNULL(cp93) + ",cp94=" + CNULL(cp94) + ",cp95=" + CNULL(cp95) + ",cp96=" + CNULL(cp96) + " where cp09='" + CP09 + "'"
   'Add By Sindy 2012/11/06 有★★的應收帳款簽核控管
   m_CP150 = ""
   If Check2.Value = 1 Then m_CP150 = "Y"
   '2012/11/06 End
   'Modify By Sindy 2012/11/06 +CP150
   strSql = "update caseprogress set cp05=" + CNULL(cp05) + ",cp06=" + CNULL(cp06) + ",cp07=" + CNULL(cp07) + ",cp10=" + CNULL(CP10) + _
            ",cp11=" + CNULL(cp11) + ",cp13=" + CNULL(cp13) + ",cp14=" + CNULL(cp14) + ",cp16=" + CNULL(cp16) + ",cp17=" + CNULL(cp17) + _
            ",cp18=" + CNULL(cp18) + ",cp19=" + CNULL(cp19) + ",cp32=" + CNULL(cp32) + ",cp56=" + CNULL(cp56) + ",cp33=" & cp33 & ",cp34=" & cp34 & ",CP64=" + CNULL(ChgSQL(CP64)) + _
            ",cp89=" + CNULL(CP89) + ",cp90=" + CNULL(CP90) + ",cp91=" + CNULL(CP91) + ",cp92=" + CNULL(CP92) + ",cp150=" & CNULL(m_CP150) & " where cp09='" + CP09 + "'"
   cnnConnection.Execute strSql
   '2009/10/19 End
   
   'Modify By Sindy 2013/8/23 +cp118(電子送件)
   'strSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(cp13) + ") where cp09=" + CNULL(CP09)
   strSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(cp13) + ") " & IIf(chkWebApp.Visible, ",cp118='" & IIf(chkWebApp.Value = 1, "Y", "") & "'", "") & " where cp09=" + CNULL(CP09)
   cnnConnection.Execute strSql
           
   '若為接洽記錄單(櫃台收文)
   'Modify by Morgan 2007/10/26 費用可改時才做，否則已收款資料會被還原
   'If frm010001.intChoose = 0 Then
   If frm010001.intChoose = 0 And txtTrademark(14).Enabled = True Then
   'end 2007/10/26
       '未收金額 = 費用
       strSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(CP09)
       cnnConnection.Execute strSql
   End If
           
   'Add By Cheng 2002/01/15
   '若為內部收文, 只要系統類別為T開頭或FCT的案件
   '且案件性質為201補正, 203修正, 302更正, 305催審, 306自請撤回, 307自請撤銷, 614未補理由, 615未答辯, 706其他
   '抓系統日期更新其案件進度檔的發文日(CP27)
   If frm010001.intChoose = 1 Then
      If Left(tm01, 1) = "T" Or tm01 = "FCT" Then
         If CP10 = "201" Or CP10 = "203" Or CP10 = "302" Or CP10 = "305" Or _
            CP10 = "306" Or CP10 = "307" Or CP10 = "614" Or CP10 = "615" Or _
            CP10 = "706" Then
            strSql = "update caseprogress set cp27= '" & ServerDate & "' where cp09=" + CNULL(CP09)
            cnnConnection.Execute strSql
         End If
      End If
   End If
   
   'Added by Lydia 2022/11/29 非內部收文並且有費用，先統一設定CP20=Null ;
   If frm010001.intChoose = 0 And Val(cp16) > 0 Then
       strSql = "update caseprogress set cp20=null where cp09=" + CNULL(CP09)
       cnnConnection.Execute strSql
   End If
   'end 2022/11/29
   
   '92.5.8 ADD BY SONIA
   If tm01 = "FCT" Then
      If Val(cp16) = 0 Then
         strSql = "update caseprogress set cp20='N',CP32='N' where cp09=" + CNULL(CP09)
         cnnConnection.Execute strSql
      End If
   End If
   '92.5.8 END
   'Add By Cheng 2002/05/10
   '若為內部收文作業時, 案件進度檔的是否向客戶收款設定為"N"
   If frm010001.intChoose = 1 Then
      strSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(CP09)
      cnnConnection.Execute strSql
   End If
   
   strSql = "update customer set cu30=" + CNULL(cu30) + " where cu01=" + CNULL(Mid(tm23, 1, 8)) + " and cu02=" + CNULL(Mid(tm23, 9, 1))
   cnnConnection.Execute strSql
   UpdateTrademarkDatabase = True
   adoquery.CursorLocation = adUseClient
   'adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np07 = '" & cp10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
   'add by nickc 2007/10/24 內商全期註冊費(717)時，抓第一期註冊費(715)
   'edit by nickc 2007/10/25 加入外商
   'If tm01 = "T" And CP10 = "717" Then
   'Modify By Sindy 2012/7/13
'   If (tm01 = "T" Or tm01 = "FCT") And CP10 = "717" Then
'       adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np06 is null and np07 = '715' ", cnnConnection, adOpenStatic, adLockReadOnly
   If (tm01 = "T" Or tm01 = "FCT") And (CP10 = "717" Or CP10 = "715") Then
       adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np06 is null and np07 in('715','717') ", cnnConnection, adOpenStatic, adLockReadOnly
   Else
       adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
   End If
   'Modify By Cheng 2002/05/10
   '若在下一程序檔只抓到一筆資料時, 才要抓下一程序檔的總收文號更新案件進度檔的相關總收文號
   'If adoquery.RecordCount <> 0 Then
   If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         '2011/6/16 add by sonia 異議答辯、評定答辯、廢止答辯要一並更新對造資料
         If (CP10 = "602" Or CP10 = "604" Or CP10 = "606") Then
            cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & CP09 & "'", intI
         Else
         '2011/6/16 end
            cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & CP09 & "'"
         End If  '2011/6/16 add by sonia
      End If
   End If
   adoquery.Close
   
   ' Add by Sindy 98/03/02
   '收文時若讓案號下一程序仍有F4103且是否續辦為NULL者,
   '更新下一程序F4103為收文智權人員
   If tm01 = "FCT" Then
      strSql = "update nextprogress set np10='" & txtTrademark(12) & "' " & _
         "where np02=" + CNULL(tm01) + " and np03=" + _
         CNULL(tm02) + " and np04=" + CNULL(tm03) + " and np05=" + CNULL(tm04) + _
         " and np10='F4103' and np06 is null"
      cnnConnection.Execute strSql
   End If
   ' 98/03/02 End

   
   'add by nickc 2008/05/02 儲存預定收款日
   'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'   Dim rtCnt As Integer
'   'Modify by Morgan 2010/12/9
'   'If txtTrademark(34) <> "" Then
'   '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " ", rtCnt
'   If txtTrademark(34) <> "" And txtTrademark(34) <> txtTrademark(34).Tag Then
'       cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
'   'end 2010/12/9
'       If rtCnt = 0 Then
'           cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from dual "
'       End If
'   End If
   'end 2018/08/22
   
   'Modified by Lydia 2022/08/22 改成共用模組
   'Call SaveFrame21(CP09) 'Add By Sindy 2012/5/8
   Call GetStrControl
   'Modified by Lydia 2022/09/29 傳入系統別,國家,案件性質 => tm01, tm10, CP10
   Call PUB_SaveByControl(CP09, m_strControl, tm01, tm10, CP10)
   'end 2022/08/22
   cnnConnection.CommitTrans
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   ShowMsg MsgText(9004)
   'add by nickc 2007/12/12
   IsSaveData = False
End Function

'讀取Trademark資料庫
'edit by nickc 2006/11/22
'Private Function ReadTrademarkDatabase(ByRef intModifyKind As Integer, ByRef tm01 As String, _
             ByRef tm02 As String, ByRef tm03 As String, ByRef tm04 As String, ByRef tm05 As String, _
             ByRef tm08 As String, ByRef tm09 As String, _
             ByRef tm10 As String, ByRef tm23 As String, ByRef tm44 As String, ByRef cp09 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, ByRef cp11 As String, _
             ByRef cp13 As String, ByRef cp14 As String, ByRef cp16 As String, ByRef cp17 As String, _
             ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, ByRef cp56 As String, ByRef cu30 As String, ByRef CP64 As String) As Boolean
'edit by nickc 2007/03/27 加入彼所案號
'Private Function ooReadTrademarkDatabase(ByRef intModifyKind As Integer, ByRef tm01 As String, _
             ByRef tm02 As String, ByRef tm03 As String, ByRef tm04 As String, ByRef tm05 As String, _
             ByRef tm08 As String, ByRef tm09 As String, _
             ByRef tm10 As String, ByRef tm23 As String, ByRef tm44 As String, ByRef CP09 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, ByRef cp11 As String, _
             ByRef cp13 As String, ByRef cp14 As String, ByRef cp16 As String, ByRef cp17 As String, _
             ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, ByRef cp56 As String, ByRef cu30 As String, ByRef CP64 As String, ByRef TM78 As String, ByRef TM79 As String, ByRef TM80 As String, ByRef TM81 As String, ByRef CP89 As String, ByRef CP90 As String, ByRef CP91 As String, ByRef CP92 As String, ByRef TM32 As String) As Boolean
'modify by sonia 2021/3/31 加TM22專用期止日
'Modifeid by Lydia 2021/04/15 +TM58案件備註
'Modifeid by Sindy 2022/12/7 +TM136證書形式
Private Function ooReadTrademarkDatabase(ByRef intModifyKind As Integer, ByRef tm01 As String, _
             ByRef tm02 As String, ByRef tm03 As String, ByRef tm04 As String, ByRef tm05 As String, _
             ByRef tm08 As String, ByRef tm09 As String, _
             ByRef tm10 As String, ByRef tm23 As String, ByRef tm44 As String, ByRef CP09 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, ByRef cp11 As String, _
             ByRef cp13 As String, ByRef cp14 As String, ByRef cp16 As String, ByRef cp17 As String, _
             ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, ByRef cp56 As String, ByRef cu30 As String, _
             ByRef CP64 As String, ByRef TM78 As String, ByRef TM79 As String, ByRef TM80 As String, ByRef TM81 As String, _
             ByRef CP89 As String, ByRef CP90 As String, ByRef CP91 As String, ByRef CP92 As String, ByRef TM32 As String, _
             ByRef TM45 As String, ByRef CP150 As String, ByRef tm22 As String, ByRef tm58 As String, ByRef tm136 As String) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset, strTemp As String
'Add by Morgan 2004/4/15
'收據號碼
Dim stCP60 As String
   
   
On Error GoTo ErrHand
   If intModifyKind <> 0 Then
      'Add by Morgan 2004/4/15
      '收據號碼
      'strSQL = "select cp05,cp06,cp07,cp10,cp11,cp13,cp16,cp17,cp18,cp19,cp32,cp56,cp14 from caseprogress where cp09='" + cp09 + "'"
      'Modify by Morgan 2005/12/13 加cp33,cp34
      'edit by nickc 2006/11/22
      'strSQL = "select cp05,cp06,cp07,cp10,cp11,cp13,cp16,cp17,cp18,cp19,cp32,cp56,cp14,cp60,cp33,cp34 from caseprogress where cp09='" + cp09 + "'"
      'Modify By Sindy 2013/8/23 +cp118
      strSql = "select cp05,cp06,cp07,cp10,cp11,cp13,cp16,cp17,cp18,cp19,cp32,cp56,cp14,cp60,cp33,cp34,cp89,cp90,cp91,cp92,cp150,cp118 from caseprogress where cp09='" + CP09 + "'"
      
      rsRecordset.CursorLocation = adUseClient
      rsRecordset.Open strSql, cnnConnection
      If rsRecordset.RecordCount > 0 Then
         'Add By Sindy 2013/8/23
         If rsRecordset.Fields("cp118") = "Y" Then
            chkWebApp.Value = 1
         End If
         '2013/8/23 END
         
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
         'add by nickc 2006/11/22
         CP89 = IIf(IsNull(rsRecordset.Fields("cp89")), "", rsRecordset.Fields("cp89"))
         CP90 = IIf(IsNull(rsRecordset.Fields("cp90")), "", rsRecordset.Fields("cp90"))
         CP91 = IIf(IsNull(rsRecordset.Fields("cp91")), "", rsRecordset.Fields("cp91"))
         CP92 = IIf(IsNull(rsRecordset.Fields("cp92")), "", rsRecordset.Fields("cp92"))
         'Add by Morgan 2005/12/13
         douStPrice = Val("" & rsRecordset("CP33"))
         douLowPrice = Val("" & rsRecordset("CP34"))
         '2005/12/13 end
            
         'Add by Morgan 2004/4/15
         stCP60 = "" & rsRecordset.Fields("cp60")
         If stCP60 <> "" Then
            txtTrademark(14).Enabled = False: txtTrademark(15).Enabled = False: txtTrademark(18).Enabled = False
            'add by nickc 2006/12/25 加鎖智權人員
            txtTrademark(12).Enabled = False
         End If
         'add by nickc 2006/11/22
         CP89 = "" & rsRecordset.Fields("cp89")
         CP90 = "" & rsRecordset.Fields("cp90")
         CP91 = "" & rsRecordset.Fields("cp91")
         CP92 = "" & rsRecordset.Fields("cp92")
         CP150 = "" & rsRecordset.Fields("cp150") 'Add By Sindy 2012/11/08
      Else
         ShowMsg MsgText(1502)
         rsRecordset.Close
         Exit Function
      End If
      rsRecordset.Close
   Else
      'Modify By Cheng 2001/12/17
   '   If GetNextProgressDate(tm01, tm02, tm03, tm04, cp10, cp06, cp07, CP64) = False Then
      'edit by nickc 2007/10/16 內商全期註冊費(717)時，抓第一期註冊費(715)
      'edit by nickc 2007/10/25 加入外商
      'If tm01 = "T" And CP10 = "717" Then
      'Modify By Sindy 2012/7/13
      'If (tm01 = "T" Or tm01 = "FCT") And CP10 = "717" Then
      If (tm01 = "T" Or tm01 = "FCT") And (CP10 = "717" Or CP10 = "715") Then
           'Modify By Sindy 2012/7/13
           'If GetNextProgressDate(tm01, tm02, tm03, tm04, "715", cp06, cp07, CP64, cp13) = False Then
           If GetNextProgressDate(tm01, tm02, tm03, tm04, "715,717", cp06, cp07, CP64, cp13) = False Then
              Exit Function
           End If
      Else
           'Added by Lydia 2020/11/19 CFT英國脫歐案管制：改抓歐盟案之下一程序性質
           'Modified by Lydia 2021/03/05 判斷非申請案txtTrademark(1) <> "101"
           If tm01 = "CFT" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" And txtTrademark(1) <> "101" Then
                If GetNextProgressDate(tm01, tm02, tm03, tm04, "110", cp06, cp07, CP64, cp13) = False Then
                   Exit Function
                End If
           Else
           'end 2020/11/19
                If GetNextProgressDate(tm01, tm02, tm03, tm04, CP10, cp06, cp07, CP64, cp13) = False Then
                   Exit Function
                End If
           End If 'Added by Lydia 2020/11/19
      End If
   End If
   If cp06 <> "" Then cp06 = ChangeWStringToTString(cp06)
   If cp07 <> "" Then cp07 = ChangeWStringToTString(cp07)
   'Modify By Cheng 2003/08/28
   'strSQL = "select tm05,tm06,tm07,tm08,tm09,tm10,tm23,tm44 from trademark where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   'edit by nickc 2006/11/22
   'strSQL = "select tm05,tm06,tm07,tm08,tm09,tm10,tm23,tm44, TM35, TM34 from trademark where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   'edit by nickc 2007/03/27 加入彼所案號
   'strSQL = "select tm05,tm06,tm07,tm08,tm09,tm10,tm23,tm44, TM35, TM34,tm78,tm79,tm80,tm81,tm32 from trademark where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   'modify by sonia 2021/3/31 加TM22專用期止日
   'Modifeid by Lydia 2021/04/15 +TM58案件備註
   'Modifeid by Sindy 2022/12/7 +TM136證書形式
   strSql = "select tm05,tm06,tm07,tm08,tm09,tm10,tm23,tm44, TM35, TM34,tm78,tm79,tm80,tm81,tm32,tm45,tm123,tm22,tm58,TM136 from trademark where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      tm05 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
   '   tm06 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
   '   tm07 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
      tm08 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
      tm09 = IIf(IsNull(rsRecordset.Fields(4)), "", rsRecordset.Fields(4))
      tm10 = IIf(IsNull(rsRecordset.Fields(5)), "", rsRecordset.Fields(5))
      tm22 = IIf(IsNull(rsRecordset.Fields("tm22")), "", rsRecordset.Fields("tm22"))  'add by sonia 2021/3/31
      tm23 = IIf(IsNull(rsRecordset.Fields(6)), "", rsRecordset.Fields(6))
      tm44 = IIf(IsNull(rsRecordset.Fields(7)), "", rsRecordset.Fields(7))
      'add by nickc 2006/11/22
      TM78 = IIf(IsNull(rsRecordset.Fields("tm78")), "", rsRecordset.Fields("tm78"))
      TM79 = IIf(IsNull(rsRecordset.Fields("tm79")), "", rsRecordset.Fields("tm79"))
      TM80 = IIf(IsNull(rsRecordset.Fields("tm80")), "", rsRecordset.Fields("tm80"))
      TM81 = IIf(IsNull(rsRecordset.Fields("tm81")), "", rsRecordset.Fields("tm81"))
      'add by nickc 2007/03/27
      TM45 = IIf(IsNull(rsRecordset.Fields("tm45")), "", rsRecordset.Fields("tm45"))
      tm58 = "" & rsRecordset.Fields("tm58") 'Added by Lydia 2021/04/15 TM58案件備註
      tm136 = "" & rsRecordset.Fields("TM136") 'Add by Sindy 2022/12/7 TM136證書形式
      
       'Add By Cheng 2003/08/28
       Me.txtTrademark(22).Text = "" & rsRecordset("TM35").Value
       Me.txtTrademark(23).Text = "" & rsRecordset("TM34").Value
       
      'add by nickc 2006/11/30
      Me.txtTrademark(32).Text = "" & rsRecordset("TM32").Value
       
      'Add by Morgan 2008/8/5
      strAppNo1 = "" & rsRecordset("TM23")
      'Modify by Amy 2021/12/20 改成Form 2.0
      'PUB_AddContact strAppNo1, cboContact, "" & rsRecordset("TM123"), True
      strExc(10) = cboContact.Tag
      PUB_AddContact strAppNo1, cboContact, "" & rsRecordset("TM123"), True, True, strExc(10)
      cboContact.Tag = strExc(10)
      'end 2008/8/5
      
      'Midify By Cheng 2002/01/03
      '若有申請人
      If Len("" & tm23) > 0 Then
         rsRecordset.Close
         strSql = "select cu30 from customer where cu01=" + CNULL(Mid(tm23, 1, 8)) + " AND cu02=" + CNULL(Mid(tm23, 9, 1))
         rsRecordset.CursorLocation = adUseClient
         rsRecordset.Open strSql, cnnConnection
         If rsRecordset.RecordCount > 0 Then
            cu30 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
            tm23 = ChangeCustomerS(tm23)
            tm44 = ChangeCustomerS(tm44)
            TM78 = ChangeCustomerS(TM78)
            TM79 = ChangeCustomerS(TM79)
            TM80 = ChangeCustomerS(TM80)
            TM81 = ChangeCustomerS(TM81)
            cp56 = ChangeCustomerS(cp56)
            ooReadTrademarkDatabase = True
         Else
            ShowMsg MsgText(1503)
            Exit Function
         End If
      
      Else
         ooReadTrademarkDatabase = True
      End If
   
   Else
      If intModifyKind <> 0 Then
         ShowMsg MsgText(1504)
      End If
   End If
  'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'   rsRecordset.Close
   'add by nickc 2008/05/02 抓預定收款日
'   strSql = "select rd05 from ReceivablesDay where (rd01,rd02,rd03) in (select rd01,rd02,max(rd03) from ReceivablesDay where (rd02) in (select max(rd02) from ReceivablesDay where rd01='" & CP09 & "' ) and rd01='" & CP09 & "' group by rd01,rd02) "
'   rsRecordset.CursorLocation = adUseClient
'   rsRecordset.Open strSql, cnnConnection
'   If rsRecordset.RecordCount > 0 Then
'      txtTrademark(34) = IIf(IsNull(rsRecordset.Fields(0)), "", TAIWANDATE(rsRecordset.Fields(0)))
'   Else
'      txtTrademark(34) = ""
'   End If
'   txtTrademark(34).Tag = txtTrademark(34) 'Add by Morgan 2010/12/9
   'end 2018/08/22
   
   rsRecordset.Close
   Exit Function
ErrHand:
      ShowMsg "資料讀取失敗,請洽系統管理者!"  '2010/8/18 add by sonia
End Function

'新增Trademark至資料庫
'edit by nickc 2006/11/22
'Private Function InsertTrademarkDatabase(ByRef intSaveMode As Integer, ByRef tm01 As String, _
             ByRef tm02 As String, ByRef tm03 As String, ByRef tm04 As String, ByRef tm05 As String, _
             ByRef tm08 As String, ByRef tm09 As String, _
             ByRef tm10 As String, ByRef tm23 As String, ByRef tm44 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, _
             ByRef cp11 As String, ByRef cp13 As String, ByRef cp16 As String, _
             ByRef cp17 As String, ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, _
             ByRef cp56 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef cp09 As String, ByRef cp02 As String, ByRef cp33 As Double, ByRef cp34 As Double, ByRef CP64 As String, ByRef TM15 As String) As Boolean
'edit by nickc 2007/03/27 加入彼所案號
'Private Function InsertTrademarkDatabase(ByRef intSaveMode As Integer, ByRef tm01 As String, _
             ByRef tm02 As String, ByRef tm03 As String, ByRef tm04 As String, ByRef tm05 As String, _
             ByRef tm08 As String, ByRef tm09 As String, _
             ByRef tm10 As String, ByRef tm23 As String, ByRef tm44 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, _
             ByRef cp11 As String, ByRef cp13 As String, ByRef cp16 As String, _
             ByRef cp17 As String, ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, _
             ByRef cp56 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef CP09 As String, ByRef cp02 As String, ByRef cp33 As Double, ByRef cp34 As Double, ByRef CP64 As String, ByRef TM15 As String, ByRef TM78 As String, ByRef TM79 As String, ByRef TM80 As String, ByRef TM81 As String, ByRef CP89 As String, ByRef CP90 As String, ByRef CP91 As String, ByRef CP92 As String, ByRef TM32 As String) As Boolean

Private Function InsertTrademarkDatabase(ByRef intSaveMode As Integer, ByRef tm01 As String, _
             ByRef tm02 As String, ByRef tm03 As String, ByRef tm04 As String, ByRef tm05 As String, _
             ByRef tm08 As String, ByRef tm09 As String, _
             ByRef tm10 As String, ByRef tm23 As String, ByRef tm44 As String, _
             ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef CP10 As String, _
             ByRef cp11 As String, ByRef cp13 As String, ByRef cp16 As String, _
             ByRef cp17 As String, ByRef cp18 As String, ByRef cp19 As String, ByRef cp32 As String, _
             ByRef cp56 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef CP09 As String, ByRef cp02 As String, ByRef cp33 As Double, ByRef cp34 As Double, ByRef CP64 As String, ByRef TM15 As String, ByRef TM78 As String, ByRef TM79 As String, ByRef TM80 As String, ByRef TM81 As String, ByRef CP89 As String, ByRef CP90 As String, ByRef CP91 As String, ByRef CP92 As String, ByRef TM32 As String, ByRef TM45 As String, ByRef TM123 As String) As Boolean

Dim strSql As String, tm28 As String, tm34 As String, cp31 As String, strAutoNumber As String
Dim np13 As String, np14 As String, bolRt As Boolean, cp55 As String, tm53 As String
Dim cp93 As String, cp94 As String, cp95 As String, cp96 As String 'Add by Nickc 2006/11/22
Dim bolError As Boolean, itm53 As Integer
Dim adoquery As New ADODB.Recordset
Dim strBKindCP09 As String 'B類收文號
Dim i As Integer, intField As Integer ', strFaData(1) As String, strCuData(1) As String 'Add by Amy 2017/01/03
Dim strApply As Variant, strAllApp As String 'Add by Amy 2017/03/09
Dim strTM130 As String 'Add by Amy 2018/10/11 收據公司別

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
   tm23 = ChangeCustomerL(tm23)
   tm44 = ChangeCustomerL(tm44)
   'add by nickc 2006/11/22
   TM78 = ChangeCustomerL(TM78)
   TM79 = ChangeCustomerL(TM79)
   TM80 = ChangeCustomerL(TM80)
   TM81 = ChangeCustomerL(TM81)
   'edit by nickc 2007/02/06 不用 dll 了
   'Dim objPublicData As Object
   'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
   cnnConnection.BeginTrans
   If intSaveMode = 1 Then
      'edit by nickc 2007/02/06 不用 dll 了
      'obj001.SetTMFileProperty CP10, tm28, tm34
      'Modified by Lydia 2022/08/11 拿掉tm34
      'Cls001SetTMFileProperty CP10, tm28, tm34
      Cls001SetTMFileProperty CP10, tm28
      If tm02 = "" Or tm02 = "0" Then
         'edit by nickc 2007/02/06 不用 dll 了
         'If objPublicData.GetAutoNumber(tm01, strAutoNumber, True, False) Then
         If ClsPDGetAutoNumber(tm01, strAutoNumber, True, False) Then
            tm02 = strAutoNumber
         Else
            bolError = True
         End If
      End If
      If bolError = False Then
         'edit by nickc 2007/02/06 不用 dll 了
         'If objPublicData.GetSystemKind(tm01, , , itm53) Then
         If ClsPDGetSystemKind(tm01, , , itm53) Then
            tm53 = IIf(itm53 = 2, 2, 1)
            cp02 = tm02
            '91.12.6 modify by sonia TM17預設null
            'strSQL = "insert into trademark (tm01,tm02,tm03,tm04,tm05,tm06,tm07,tm08,tm09,tm10," + _
            '  "tm23,tm28,tm34,tm44,tm17) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(ChgSQL(tm05)) + "," + _
            '  CNULL(Replace(tm06, "'", "''")) + "," + CNULL(ChgSQL(tm07)) + "," + CNULL(tm08) + "," + CNULL(tm09) + "," + CNULL(tm10) + "," + CNULL(tm23) + "," + CNULL(tm28) + "," + CNULL(tm34) + "," + CNULL(tm44) + ", 'N')"
   '         strSQL = "insert into trademark (tm01,tm02,tm03,tm04,tm05,tm06,tm07,tm08,tm09,tm10," + _
   '           "tm23,tm28,tm34,tm44,tm17) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(ChgSQL(tm05)) + "," + _
   '           CNULL(Replace(tm06, "'", "''")) + "," + CNULL(ChgSQL(tm07)) + "," + CNULL(tm08) + "," + CNULL(tm09) + "," + CNULL(tm10) + "," + CNULL(tm23) + "," + CNULL(tm28) + "," + CNULL(tm34) + "," + CNULL(tm44) + ", '')"
   'edit by nickc 2006/11/22 加申請人
   '         strSQL = "insert into trademark (tm01, tm02, tm03, tm04, tm05, tm08, tm09, tm10," + _
              "tm23,tm28,tm34,tm44,tm17,TM15) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(ChgSQL(tm05)) + "," + _
              CNULL(tm08) + "," + CNULL(tm09) + "," + CNULL(tm10) + "," + CNULL(tm23) + "," + CNULL(tm28) + "," + CNULL(tm34) + "," + CNULL(tm44) + ", ''" + "," + CNULL(ChgSQL(TM15)) + ")"
   'edit by  nickc 2007/03/27 加入彼所案號
   '         strSQL = "insert into trademark (tm01, tm02, tm03, tm04, tm05, tm08, tm09, tm10," + _
              "tm23,tm28,tm34,tm44,tm17,TM15,tm78,tm79,tm80,tm81,tm32) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(ChgSQL(tm05)) + "," + _
              CNULL(tm08) + "," + CNULL(tm09) + "," + CNULL(tm10) + "," + CNULL(tm23) + "," + CNULL(tm28) + "," + CNULL(tm34) + "," + CNULL(tm44) + ", ''" + "," + CNULL(ChgSQL(TM15)) + "," + CNULL(TM78) + "," + CNULL(TM79) + "," + CNULL(TM80) + "," + CNULL(TM81) + "," + CNULL(TM32) + ")"
            'Modify by Morgan 2008/8/5 +TM123
            'Add By Sindy 2012/7/19 若申請人或代理人為諾華公司者，案件備註若無"不銷卷"字樣,則要加入
            Dim strTM58 As String
            strTM58 = ""
            If (txtTrademark(9) <> "" And InStr(strTmNovartisCust, Left(txtTrademark(9), 6)) > 0) Or _
               (txtTrademark(24) <> "" And InStr(strTmNovartisCust, Left(txtTrademark(24), 6)) > 0) Or _
               (txtTrademark(25) <> "" And InStr(strTmNovartisCust, Left(txtTrademark(25), 6)) > 0) Or _
               (txtTrademark(26) <> "" And InStr(strTmNovartisCust, Left(txtTrademark(26), 6)) > 0) Or _
               (txtTrademark(27) <> "" And InStr(strTmNovartisCust, Left(txtTrademark(27), 6)) > 0) Or _
               (txtTrademark(10) <> "" And InStr(strTmNovartisCust, Left(txtTrademark(10), 6)) > 0) Then
               strTM58 = ChangeTStringToTDateString(strSrvDate(2)) & "不銷卷"
            End If
            '2012/7/19 end
            'ADD BY SONIA 2015/11/24 若申請為旅狐國際及部分關係企業者，案件備註若無"不銷卷"字樣,則要加入
            strTM58 = ""
            If (txtTrademark(9) <> "" And InStr(strTmTRAVEL_FOXCust, Left(txtTrademark(9), 8)) > 0) Or _
               (txtTrademark(24) <> "" And InStr(strTmTRAVEL_FOXCust, Left(txtTrademark(24), 8)) > 0) Or _
               (txtTrademark(25) <> "" And InStr(strTmTRAVEL_FOXCust, Left(txtTrademark(25), 8)) > 0) Or _
               (txtTrademark(26) <> "" And InStr(strTmTRAVEL_FOXCust, Left(txtTrademark(26), 8)) > 0) Or _
               (txtTrademark(27) <> "" And InStr(strTmTRAVEL_FOXCust, Left(txtTrademark(27), 8)) > 0) Then
               strTM58 = ChangeTStringToTDateString(strSrvDate(2)) & "不銷卷"
            End If
            'END 2015/11/24
            'Modify by Sindy 2012/7/19 +TM58
            'MODIFY BY SONIA 2014/8/12 +TM16(延展新案要存tm16,否則重覆收文檢查不出來)
            'strSql = "insert into trademark (tm01, tm02, tm03, tm04, tm05, tm08, tm09, tm10," + _
              "tm23,tm28,tm34,tm44,tm17,TM15,tm78,tm79,tm80,tm81,tm32,tm45,tm123,tm58) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(ChgSQL(tm05)) + "," + _
              CNULL(tm08) + "," + CNULL(tm09) + "," + CNULL(tm10) + "," + CNULL(tm23) + "," + CNULL(tm28) + "," + CNULL(tm34) + "," + CNULL(tm44) + ", ''" + "," + CNULL(ChgSQL(TM15)) + "," + CNULL(TM78) + "," + CNULL(TM79) + "," + CNULL(TM80) + "," + CNULL(TM81) + "," + CNULL(TM32) + "," + CNULL(ChgSQL(TM45)) + "," + CNULL(TM123) + _
              "," + CNULL(strTM58) + ")"
            strTM130 = GetReceiptCmp(Left(tm23, 8), Mid(tm23, 9, 1), tm01, tm10) 'Add by Amy 2018/10/11 +收據公司別tm130
            'Added by Lydia 2020/11/19 CFT英國脫歐案管制：新增英國案時同時把歐盟案相關欄位帶過來(參考PUB_SaveCountry)
            'Modified by Lydia 2021/03/05 判斷非申請案；CFT歐盟尚未註冊案轉換英國申請案收文控管
            If tm01 = "CFT" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" And CP10 <> "101" Then
                If PUB_ReadTradeMarkData(m_CaseNa239(), m_CaseNa239(1), m_CaseNa239(2), m_CaseNa239(3), m_CaseNa239(4)) Then
                   strExc(0) = "": strExc(1) = ""
                   'Added by Lydia 2020/12/09 專用期間(止)=歐盟案之法限; 避免歐盟案先收延展,更新商標基本檔的專用期間(止)
                   strExc(9) = ""
                   strExc(2) = "select np09 from nextprogress where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07=" & CNULL(IIf(CP10 = "102", "110", CP10)) & " and np06 is null "
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(2))
                   If intI = 1 Then
                       strExc(9) = "" & RsTemp.Fields("np09")
                   End If
                   'end 2020/12/09
                   For i = 5 To TF_TM
                       Select Case i
                          Case 59, 60, 61, 62, 63, 64, 57, 73, 74, 75 'Create + Update, 銷卷日
                          Case 10 '申請國家
                              strExc(0) = strExc(0) & "TM" & Format(i, "00") & "," 'Insert
                              strExc(1) = strExc(1) & " '201' as TM10, " 'Select
                          Case 12, 15 '申請號、審定號：UK009+歐盟號數後8碼(拿掉第1碼0)
                              strExc(0) = strExc(0) & "TM" & Format(i, "00") & ","
                              strExc(1) = strExc(1) & " 'UK009'||substr(tm15,2,8) AS TM" & Format(i, "00") & ","
                          'Added by Lydia 2020/12/09 專用期間(止)=歐盟案之法限; 避免歐盟案先收延展,更新商標基本檔的專用期間(止)
                          Case 22    '專用期間(止日)
                              strExc(0) = strExc(0) & "TM" & Format(i, "00") & ","
                              strExc(1) = strExc(1) & " " & IIf(strExc(9) <> "", CNULL(strExc(9)), "TM22") & " as TM22, "
                          'end 2020/12/09
                          Case 58  '案件備註: 加註歐盟案案號
                              strExc(0) = strExc(0) & "TM" & Format(i, "00") & ","
                              strExc(1) = strExc(1) & CNULL("歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";") & "||TM" & Format(i, "00") & " AS TM" & Format(i, "00") & ","
                          Case 130
                              strExc(0) = strExc(0) & "TM" & Format(i, "00") & ","
                              strExc(1) = strExc(1) & " '" & strTM130 & "' as TM130, "
                          Case Else
                              strExc(0) = strExc(0) & "TM" & Format(i, "00") & ","
                              strExc(1) = strExc(1) & "TM" & Format(i, "00") & ","
                       End Select
                   Next
                   strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
                   strExc(1) = Left(strExc(1), Len(strExc(1)) - 1)
                   strSql = "INSERT INTO TRADEMARK (TM01,TM02,TM03,TM04," & strExc(0) & ") " & _
                               "SELECT '" & tm01 & "' as TM01,'" & tm02 & "' as TM02,'" & tm03 & "' as TM03,'" & tm04 & "' as TM04, " & strExc(1) & _
                               " FROM TRADEMARK WHERE TM01='" & m_CaseNa239(1) & "' and TM02='" & m_CaseNa239(2) & "' and TM03='" & m_CaseNa239(3) & "' and TM04='" & m_CaseNa239(4) & "' "
                   cnnConnection.Execute strSql
                   'Added by Lydia 2021/01/08 複製指定商品或服務名稱
                   strSql = "INSERT INTO TMGOODS(TG01,TG02,TG03,TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17) " & _
                               "SELECT '" & tm01 & "' as TG01, '" & tm02 & "' as TG02, '" & tm03 & "' as  TG03, '" & tm04 & "' as  TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17 " & _
                               "FROM TMGOODS WHERE TG01='" & m_CaseNa239(1) & "'  AND TG02='" & m_CaseNa239(2) & "' AND TG03='" & m_CaseNa239(3) & "'  AND TG04='" & m_CaseNa239(4) & "'  AND TG18 IS NULL "
                   cnnConnection.Execute strSql
                   'end 2021/01/08
                End If
            Else
            'end 2020/11/19
                'Modify by Amy 2018/10/11 +收據公司別tm130
                strSql = "insert into trademark (tm01, tm02, tm03, tm04, tm05, tm08, tm09, tm10," + _
                  "tm23,tm28,tm34,tm44,tm17,TM15,tm78,tm79,tm80,tm81,tm32,tm45,tm123,tm58,tm16,tm130) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(ChgSQL(tm05)) + "," + _
                  CNULL(tm08) + "," + CNULL(tm09) + "," + CNULL(tm10) + "," + CNULL(tm23) + "," + CNULL(tm28) + "," + CNULL(tm34) + "," + CNULL(tm44) + ", ''" + "," + CNULL(ChgSQL(TM15)) + "," + CNULL(TM78) + "," + CNULL(TM79) + "," + CNULL(TM80) + "," + CNULL(TM81) + "," + CNULL(TM32) + "," + CNULL(ChgSQL(TM45)) + "," + CNULL(TM123) + _
                  "," + CNULL(strTM58) + "," + IIf(ChgSQL(TM15) = "", "NULL", "1") + "," + CNULL(strTM130) + ")"
                'end 2018/10/11
                 cnnConnection.Execute strSql
                 'edit by nickc 2006/01/26 新增時要將郵遞區號放在前面
                 'strSQL = "update trademark set tm24=(select cu23 from customer where cu01=" + CNULL(Mid(tm23, 1, 8)) + " and cu02=" + CNULL(Mid(tm23, 9, 1)) + _
                    "),tm25=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(tm23, 1, 8)) + " and cu02=" + CNULL(Mid(tm23, 9, 1)) + _
                    "),tm26=(select cu29 from customer where cu01=" + CNULL(Mid(tm23, 1, 8)) + " and cu02=" + CNULL(Mid(tm23, 9, 1)) + ") where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
                'Memo by Lydia 2020/11/19 CFT英國脫歐案管制：新增英國案時同時把歐盟案相關欄位帶過來，所以不要變更資料
                 strSql = "update trademark set tm24=(select nvl(cu112,'')||cu23 from customer where cu01=" + CNULL(Mid(tm23, 1, 8)) + " and cu02=" + CNULL(Mid(tm23, 9, 1)) + _
                    "),tm25=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(tm23, 1, 8)) + " and cu02=" + CNULL(Mid(tm23, 9, 1)) + _
                    "),tm26=(select cu29 from customer where cu01=" + CNULL(Mid(tm23, 1, 8)) + " and cu02=" + CNULL(Mid(tm23, 9, 1)) + ") where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
                 cnnConnection.Execute strSql
                'add by nickc 2006/11/30
                strSql = "update trademark set tm82=(select cu23 from customer where cu01=" + CNULL(Mid(TM78, 1, 8)) + " and cu02=" + CNULL(Mid(TM78, 9, 1)) + _
                   "),tm86=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(TM78, 1, 8)) + " and cu02=" + CNULL(Mid(TM78, 9, 1)) + _
                   "),tm90=(select cu29 from customer where cu01=" + CNULL(Mid(TM78, 1, 8)) + " and cu02=" + CNULL(Mid(TM78, 9, 1)) + ") where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
                cnnConnection.Execute strSql
                strSql = "update trademark set tm83=(select cu23 from customer where cu01=" + CNULL(Mid(TM79, 1, 8)) + " and cu02=" + CNULL(Mid(TM79, 9, 1)) + _
                   "),tm87=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(TM79, 1, 8)) + " and cu02=" + CNULL(Mid(TM79, 9, 1)) + _
                   "),tm91=(select cu29 from customer where cu01=" + CNULL(Mid(TM79, 1, 8)) + " and cu02=" + CNULL(Mid(TM79, 9, 1)) + ") where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
                cnnConnection.Execute strSql
                strSql = "update trademark set tm84=(select cu23 from customer where cu01=" + CNULL(Mid(TM80, 1, 8)) + " and cu02=" + CNULL(Mid(TM80, 9, 1)) + _
                   "),tm88=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(TM80, 1, 8)) + " and cu02=" + CNULL(Mid(TM80, 9, 1)) + _
                   "),tm92=(select cu29 from customer where cu01=" + CNULL(Mid(TM80, 1, 8)) + " and cu02=" + CNULL(Mid(TM80, 9, 1)) + ") where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
                cnnConnection.Execute strSql
                strSql = "update trademark set tm85=(select cu23 from customer where cu01=" + CNULL(Mid(TM81, 1, 8)) + " and cu02=" + CNULL(Mid(TM81, 9, 1)) + _
                   "),tm89=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(TM81, 1, 8)) + " and cu02=" + CNULL(Mid(TM81, 9, 1)) + _
                   "),tm93=(select cu29 from customer where cu01=" + CNULL(Mid(TM81, 1, 8)) + " and cu02=" + CNULL(Mid(TM81, 9, 1)) + ") where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
                cnnConnection.Execute strSql
            End If 'Added by Lydia 2020/11/19
            '91.12.6 end
             
            'Added by Lydia 2021/03/05 CFT歐盟尚未註冊案轉換英國申請案收文控管：一併將歐盟關聯案之「商標圖樣」、「商品/服務類別及名稱」、「申請日」、「優先權資料」帶入新案號
            If tm01 = "CFT" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" And CP10 = "101" Then
                strSql = "update trademark set (tm09,tm11) = (select tm09,tm11 from trademark where TM01='" & m_CaseNa239(1) & "' and TM02='" & m_CaseNa239(2) & "' and TM03='" & m_CaseNa239(3) & "' and TM04='" & m_CaseNa239(4) & "' " & _
                             ") where TM01='" & tm01 & "' and TM02='" & tm02 & "' and TM03='" & tm03 & "' and TM04='" & tm04 & "' "
                cnnConnection.Execute strSql
                '複製指定商品或服務名稱
                strSql = "INSERT INTO TMGOODS(TG01,TG02,TG03,TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17) " & _
                            "SELECT '" & tm01 & "' as TG01, '" & tm02 & "' as TG02, '" & tm03 & "' as  TG03, '" & tm04 & "' as  TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17 " & _
                            "FROM TMGOODS WHERE TG01='" & m_CaseNa239(1) & "'  AND TG02='" & m_CaseNa239(2) & "' AND TG03='" & m_CaseNa239(3) & "'  AND TG04='" & m_CaseNa239(4) & "'  AND TG18 IS NULL "
                cnnConnection.Execute strSql
            End If
            'end 2021/03/05
             
           'Add By Cheng 2003/08/28
           'Begin
   'edit by nick 2005/01/07 搬到下面
   '        strSQL = "Update Trademark Set TM34='" & ChgSQL(Me.txtTrademark(23).Text) & "', TM35='" & ChgSQL(Me.txtTrademark(22).Text) & "' Where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
   '        cnnConnection.Execute strSQL
           'End
            cp31 = "Y"
         Else
            bolError = True
         End If
      End If
      'Add by Amy 2017/01/03 MCTF控管 T字頭新案且有輸FC代理人且收文業務區為 P2字頭且申請國家是台灣時,依FC代理人之管控智權人員更新客戶檔之CU12,CU13
      'Modify by Amy 2017/03/09 原程式改寫至Function
      'Memo by Amy 2017/03/22 修改UpdMCTF_Cu13 拿掉申請國家是台灣的判斷
      If Len(Trim(txtTrademark(10))) > 0 Then
        For i = 0 To 4
            If i = 0 Then
                If Len(Trim(txtTrademark(9))) = 0 Then Exit For
                strAllApp = strAllApp & "," & ChangeCustomerL(txtTrademark(9))
            ElseIf Len(Trim(txtTrademark(23 + i))) = 0 Then
               Exit For
            Else
                strAllApp = strAllApp & "," & ChangeCustomerL(txtTrademark(23 + i))
            End If
        Next i
        If strAllApp <> MsgText(601) Then
            strApply = Split(Mid(strAllApp, 2), ",")
            If UpdMCTF_Cu13(tm01, ChangeCustomerL(txtTrademark(10)), Trim(txtTrademark(5)), strApply, PUB_GetST03(txtTrademark(12))) = False Then
                  GoTo ErrHand
            End If
        End If
      End If
      'end 2017/03/09
   End If
   If bolError = False Then
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetAutoNumber(Left(CP09, 1), strAutoNumber, True, True) Then
      If ClsPDGetAutoNumber(Left(CP09, 1), strAutoNumber, True, True) Then
         cp56 = ChangeCustomerL(cp56)
         If cp56 <> "" Then
           cp55 = tm23
            'Add by Morgan 2006/11/22
            '讓與人2-5,受讓人2-5
            CP89 = ChangeCustomerL(CP89)
            CP90 = ChangeCustomerL(CP90)
            CP91 = ChangeCustomerL(CP91)
            CP92 = ChangeCustomerL(CP92)
            cp93 = TM78
            cp94 = TM79
            cp95 = TM80
            cp96 = TM81
            'end 2006/6/23
         End If
         CP09 = CP09 + strAutoNumber
         'edit by nickc 2007/02/06 不用 dll 了
         'bolRt = obj001.GetNextProgressData(tm01, tm02, tm03, tm04, CP10, np13, np14)
         bolRt = Cls001GetNextProgressData(tm01, tm02, tm03, tm04, CP10, np13, np14)
          'add by nick 2005/01/07 從上面搬下來
          If Me.txtTrademark(23).Text & Me.txtTrademark(22).Text <> "" Then
               strSql = "Update Trademark Set TM34='" & ChgSQL(Me.txtTrademark(23).Text) & "', TM35='" & ChgSQL(Me.txtTrademark(22).Text) & "' Where tm01=" + CNULL(tm01) + " and tm02=" + CNULL(tm02) + " and tm03=" + CNULL(tm03) + " and tm04=" + CNULL(tm04)
               cnnConnection.Execute strSql
          End If
         
         'Add By Sindy 2012/11/06 有★★的應收帳款簽核控管
         m_CP150 = ""
         If Check2.Value = 1 Then m_CP150 = "Y"
         '2012/11/06 End
         
         'Modify By Sindy 2012/11/06 +CP150
         If tm28 <> "1" And cp31 = "Y" Then
            If bolRt Then
               'Modify By Cheng 2002/09/25
   '            strSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," + _
   '                "cp16,cp17,cp18,cp19,cp31,cp32,cp40,cp55,cp56,cp33,cp34,cp37,cp38,cp39,CP64) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
   '                CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(np13) + "," + CNULL(cp09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
   '                CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(np14) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + cnull(chgsql(tm05)) + "," + cnull(chgsql(tm06)) + "," + cnull(chgsql(tm07)) + "," + cnull(chgsql(CP64)) + ")"
   '            strSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," + _
   '                "cp16,cp17,cp18,cp19,cp31,cp32,cp40,cp55,cp56,cp33,cp34,cp37,cp38,cp39,CP64) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
   '                CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(np13) + "," + CNULL(cp09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
   '                CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(ChgSQL(np14)) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(tm05)) + "," + CNULL(ChgSQL(tm06)) + "," + CNULL(ChgSQL(tm07)) + "," + CNULL(ChgSQL(CP64)) + ")"
   'edit by nickc 2006/11/22 加cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96
   '            strSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," + _
                   "cp16,cp17,cp18,cp19,cp31,cp32,cp40,cp55,cp56,cp33,cp34,cp37, CP64) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
                   CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(np13) + "," + CNULL(cp09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
                   CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(ChgSQL(np14)) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(tm05)) + "," + CNULL(ChgSQL(CP64)) + ")"
               strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," + _
                   "cp16,cp17,cp18,cp19,cp31,cp32,cp40,cp55,cp56,cp33,cp34,cp37, CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
                   CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(np13) + "," + CNULL(CP09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
                   CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(ChgSQL(np14)) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(tm05)) + "," + CNULL(ChgSQL(CP64)) + "," + CNULL(CP89) + "," + CNULL(CP90) + "," + CNULL(CP91) + "," + CNULL(CP92) + "," + CNULL(cp93) + "," + CNULL(cp94) + "," + CNULL(cp95) + "," + CNULL(cp96) + "," + CNULL(m_CP150) + ")"
            Else
   '            strSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp13,cp14," + _
   '                "cp16,cp17,cp18,cp19,cp31,cp32,cp55,cp56,cp33,cp34,cp37,cp38,cp39,CP64) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
   '                CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(cp09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
   '                CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(tm05)) + "," + CNULL(ChgSQL(tm06)) + "," + CNULL(ChgSQL(tm07)) + "," + CNULL(ChgSQL(CP64)) + ")"
   'edit by nickc 2006/11/22 加cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96
   '            strSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp13,cp14," + _
                   "cp16,cp17,cp18,cp19,cp31,cp32,cp55,cp56,cp33,cp34,cp37, CP64) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
                   CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(cp09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
                   CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(tm05)) + "," + CNULL(ChgSQL(CP64)) + ")"
               strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp13,cp14," + _
                   "cp16,cp17,cp18,cp19,cp31,cp32,cp55,cp56,cp33,cp34,cp37, CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
                   CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(CP09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
                   CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(tm05)) + "," + CNULL(ChgSQL(CP64)) + "," + CNULL(CP89) + "," + CNULL(CP90) + "," + CNULL(CP91) + "," + CNULL(CP92) + "," + CNULL(cp93) + "," + CNULL(cp94) + "," + CNULL(cp95) + "," + CNULL(cp96) + "," + CNULL(m_CP150) + ")"
            End If
         Else
            If bolRt Then
               'Modify By Cheng 2002/09/25
   '            strSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," + _
   '                "cp16,cp17,cp18,cp19,cp31,cp32,cp40,cp55,cp56,cp33,cp34,CP64) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
   '                CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(np13) + "," + CNULL(cp09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
   '                CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(np14) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + cnull(chgsql(CP64)) + ")"
   'edit by nickc 2006/11/22 加cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96
   '            strSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," + _
                   "cp16,cp17,cp18,cp19,cp31,cp32,cp40,cp55,cp56,cp33,cp34,CP64) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
                   CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(np13) + "," + CNULL(cp09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
                   CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(ChgSQL(np14)) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(CP64)) + ")"
               strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," + _
                   "cp16,cp17,cp18,cp19,cp31,cp32,cp40,cp55,cp56,cp33,cp34,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
                   CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(np13) + "," + CNULL(CP09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
                   CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(ChgSQL(np14)) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(CP64)) + "," + CNULL(CP89) + "," + CNULL(CP90) + "," + CNULL(CP91) + "," + CNULL(CP92) + "," + CNULL(cp93) + "," + CNULL(cp94) + "," + CNULL(cp95) + "," + CNULL(cp96) + "," + CNULL(m_CP150) + ")"
            Else
   'edit by nickc 2006/11/22 加cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96
   '            strSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp13,cp14," + _
                   "cp16,cp17,cp18,cp19,cp31,cp32,cp55,cp56,cp33,cp34,CP64) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
                   CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(cp09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
                   CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(CP64)) + ")"
               strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp13,cp14," + _
                   "cp16,cp17,cp18,cp19,cp31,cp32,cp55,cp56,cp33,cp34,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150) values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
                   CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(CP09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL(cp16) + "," + _
                   CNULL(cp17) + "," + CNULL(cp18) + "," + CNULL(cp19) + "," + CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(cp55) + "," + CNULL(cp56) + ", " & cp33 & ", " & cp34 & "," + CNULL(ChgSQL(CP64)) + "," + CNULL(CP89) + "," + CNULL(CP90) + "," + CNULL(CP91) + "," + CNULL(CP92) + "," + CNULL(cp93) + "," + CNULL(cp94) + "," + CNULL(cp95) + "," + CNULL(cp96) + "," + CNULL(m_CP150) + ")"
            End If
         End If
         cnnConnection.Execute strSql
           'Add By Cheng 2004/03/16
           '若為CFT的商申案
           If tm01 = "CFT" And CP10 = "101" Then
               '若申請國家為�揮Q亞(032), 科威特(028), 伊朗(025), 巴林(041), 冰島(218), 丹麥(216), 希臘(212), 土耳其(235), 挪威(215), 瑞典(214), 應產生B類"申請英文證明"(304)
               '2009/1/10 modify by sonia 取消伊朗(025)，土耳其(235)，挪威(215)
               'If tm10 = "032" Or tm10 = "028" Or tm10 = "025" Or tm10 = "041" Or tm10 = "218" Or tm10 = "216" Or tm10 = "212" Or tm10 = "235" Or tm10 = "215" Or tm10 = "214" Then
               '2010/5/10 modify by sonia 增加058尼泊爾
               '2010/6/22 modify by sonia 取消216丹麥
               '2010/9/7  modify by sonia 取消028科威特
               '2016/9/22 modify by sonai 取消218冰島
               'modify by sonia 2017/4/10 取消041巴林
               'modify by sonia 2018/9/18 取消212希臘
               If tm10 = "032" Or tm10 = "214" Or tm10 = "058" Then
                   strBKindCP09 = AutoNo("B", 6)
                   strSql = "Insert Into Caseprogress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP11,CP13,CP14,CP20,CP32) " & _
                                   " Values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
                                   CNULL(strBKindCP09) + "," + CNULL("304") + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL("N") + "," + CNULL("N") + ")"
                   cnnConnection.Execute strSql
                   strSql = "Update Caseprogress Set CP12=(Select ST15 From Staff Where ST01=" + CNULL(cp13) + ") Where CP09=" + CNULL(strBKindCP09)
                   cnnConnection.Execute strSql
               End If
               '若申請國家為緬甸(048), 阿根廷(118), 秘魯(116), 委內瑞拉(113), 埃及(303), 應產生B類"文件簽證"(711)
               '2007/10/5 MODIFY BY SONIA 加 玻利維亞(120)
               '2010/4/2 MODIFY BY SONIA 取消秘魯(116)
               'Mark by Amy 2016/05/24 取消「文件簽證(711)收文」
    '               If tm10 = "048" Or tm10 = "118" Or tm10 = "113" Or tm10 = "303" Or tm10 = "120" Then
    '                   strBKindCP09 = AutoNo("B", 6)
    '                   strSql = "Insert Into Caseprogress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP11,CP13,CP14,CP20,CP32) " & _
    '                                   " Values (" + CNULL(tm01) + "," + CNULL(tm02) + "," + CNULL(tm03) + "," + CNULL(tm04) + "," + CNULL(cp05) + "," + _
    '                                   CNULL(strBKindCP09) + "," + CNULL("711") + "," + CNULL(cp11) + "," + CNULL(cp13) + "," + CNULL(cp14) + "," + CNULL("N") + "," + CNULL("N") + ")"
    '                   cnnConnection.Execute strSql
    '                   strSql = "Update Caseprogress Set CP12=(Select ST15 From Staff Where ST01=" + CNULL(cp13) + ") Where CP09=" + CNULL(strBKindCP09)
    '                   cnnConnection.Execute strSql
    '               End If
             
           End If
           'End
         
         'Modify By Sindy 2013/8/23 +cp118(電子送件)
         'strSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(cp13) + ") where cp09=" + CNULL(CP09)
         strSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(cp13) + ") " & IIf(chkWebApp.Visible, ",cp118='" & IIf(chkWebApp.Value = 1, "Y", "") & "'", "") & " where cp09=" + CNULL(CP09)
         cnnConnection.Execute strSql
           
         'Added by Lydia 2020/05/20 法律所案源收文：台灣案B1、B2及C收文時，增加"案源單號"欄位一定要輸入，並將案源單號更新至該筆收文的CP162。
         If frm010001.intModifyKind = 0 And txtTrademark(5) = "000" And (txtSystem = "FCT" Or txtSystem = "T" Or txtSystem = "TC") And m_LOS02 <> "" And m_LOS15 <> "" Then
              If Left(m_LOS02, 1) = "B" Or Left(m_LOS02, 1) = "C" Then
                  strSql = "update caseprogress set CP162='" & m_LOS15 & "' where cp09='" & CP09 & "' "
                  cnnConnection.Execute strSql
              End If
         End If
         'end 2020/05/20
      
           '若為接洽記錄單(櫃台收文)
           'Modify by Morgan 2007/10/26 費用可改時才做，否則已收款資料會被還原
           'If frm010001.intChoose = 0 Then
           If frm010001.intChoose = 0 And txtTrademark(14).Enabled = True Then
           'end 2007/10/26
               '未收金額 = 費用
               strSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(CP09)
               cnnConnection.Execute strSql
           End If
         'Add By Cheng 2002/01/15
         '若為內部收文, 只要系統類別為T開頭或FCT的案件
         '且案件性質為201補正, 203修正, 302更正, 305催審, 306自請撤回, 307自請撤銷, 614未補理由, 615未答辯, 706其他
         '抓系統日期更新其案件進度檔的發文日(CP27)
         If frm010001.intChoose = 1 Then
            If Left(tm01, 1) = "T" Or tm01 = "FCT" Then
               If CP10 = "201" Or CP10 = "203" Or CP10 = "302" Or CP10 = "305" Or _
                  CP10 = "306" Or CP10 = "307" Or CP10 = "614" Or CP10 = "615" Or _
                  CP10 = "706" Then
                  strSql = "update caseprogress set cp27= '" & ServerDate & "' where cp09=" + CNULL(CP09)
                  cnnConnection.Execute strSql
               End If
            End If
         End If
         'Added by Lydia 2022/11/29 非內部收文並且有費用，先統一設定CP20=Null ;
         If frm010001.intChoose = 0 And Val(cp16) > 0 Then
             strSql = "update caseprogress set cp20=null where cp09=" + CNULL(CP09)
             cnnConnection.Execute strSql
         End If
         'end 2022/11/29
         
         '92.5.8 ADD BY SONIA
         If tm01 = "FCT" Then
            If Val(cp16) = 0 Then
               strSql = "update caseprogress set cp20='N',CP32='N' where cp09=" + CNULL(CP09)
               cnnConnection.Execute strSql
            End If
         End If
         '92.5.8 END
         'Add By Cheng 2002/05/10
         '若為內部收文作業時, 案件進度檔的是否向客戶收款設定為"N"
         If frm010001.intChoose = 1 Then
            strSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(CP09)
            cnnConnection.Execute strSql
         End If

         
         'Added by Lydia 2020/11/19 CFT英國脫歐案管制
         'Modified by Lydia 2021/03/05 判斷非申請案
        If tm01 = "CFT" And cp31 = "Y" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" And CP10 <> "101" Then
             strExc(0) = "select cp09,cp30 from caseprogress where cp01='" & m_CaseNa239(1) & "' and cp02='" & m_CaseNa239(2) & "' and cp03='" & m_CaseNa239(3) & "' and cp04='" & m_CaseNa239(4) & "' " & _
                              "and substr(cp09,1,1) ='C' and cp10='1730' and cp159=0 order by cp05 desc "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
                'A. 歐盟案若有「通知英國再註冊」的C類來函1730之CP30存至新英國案之審定號TM15
                If "" & RsTemp.Fields("CP30") <> "" Then
                    strSql = "update TRADEMARK set TM15='" & ChgSQL(RsTemp.Fields("cp30")) & "' where TM01='" & tm01 & "' and TM02='" & tm02 & "' and TM03='" & tm03 & "' and TM04='" & tm04 & "' "
                    cnnConnection.Execute strSql
                End If
                'B. 歐盟案若有「通知英國再註冊」的C類來函也轉至新英國案號
                If "" & RsTemp.Fields("cp09") <> "" Then
                     strSql = "update caseprogress set cp01='" & tm01 & "', cp02='" & tm02 & "', cp03='" & tm03 & "', cp04='" & tm04 & "' where cp09='" & RsTemp.Fields("cp09") & "' "
                     cnnConnection.Execute strSql
                End If
             End If
             'Added by Lydia 2020/12/01
             If CP10 = "710" Then
                  '委任代理人上續辦；下一程序備註加註「英國案案號」
                  strSql = "update nextprogress set np06='Y', np24='" & CP09 & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "英國案案號：" & tm01 & tm02 & tm03 & tm04 & ";'||np15 " & _
                              "where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='710' and np06 is null "
                  cnnConnection.Execute strSql
                  'E. 若收文「委任代理人(CFT.710)」時歐盟案下一程序之/「延展(英國)110」期限轉至新案號並改案件性質為「延展102」；下一程序備註加註「歐盟案案號」
                  'Modified by Lydia 2020/12/16 將NP01改為英國案收文號; 否則分案作業會錯誤(本所案號不同)
                  'strSql = "update nextprogress set np02='" & tm01 & "', np03='" & tm02 & "', np04='" & tm03 & "', np05='" & tm04 & "', np07='102', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";'||np15 " & _
                              "where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='110' and np06 is null "
                  strSql = "update nextprogress set np01='" & CP09 & "', np02='" & tm01 & "', np03='" & tm02 & "', np04='" & tm03 & "', np05='" & tm04 & "', np07='102', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";'||np15 " & _
                              "where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='110' and np06 is null "
                  cnnConnection.Execute strSql
             Else
             'end 2020/12/01
                  'C. 歐盟案下一程序之「延展(英國)」(CFT.110)期限上續辦NP06，下一單據編號NP24記錄新英國案之總收文號；下一程序備註加註「英國案案號」
                  strSql = "update nextprogress set np06='Y', np24='" & CP09 & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "英國案案號：" & tm01 & tm02 & tm03 & tm04 & ";'||np15 where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='110' and np06 is null "
                  cnnConnection.Execute strSql
                  'Added by Lydia 2020/12/04 將歐盟案「委任代理人」期限轉至新英國案；下一程序備註加註「歐盟案案號」
                  'Modified by Lydia 2020/12/16 將NP01改為英國案收文號; 否則分案作業會錯誤(本所案號不同)
                  'strSql = "update nextprogress set np02='" & tm01 & "', np03='" & tm02 & "', np04='" & tm03 & "', np05='" & tm04 & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";'||np15 " & _
                              "where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='710' and np06 is null "
                  strSql = "update nextprogress set np01='" & CP09 & "', np02='" & tm01 & "', np03='" & tm02 & "', np04='" & tm03 & "', np05='" & tm04 & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";'||np15 " & _
                              "where np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "' and np04='" & m_CaseNa239(3) & "' and np05='" & m_CaseNa239(4) & "' and np07='710' and np06 is null "
                  cnnConnection.Execute strSql
                  'end 2020/12/04
             End If 'Added by Lydia 2020/12/01
             'D. 建立歐盟案及英國案之關聯(相關卷號)
             strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(tm01) & ", " & CNULL(tm02) & ", " & CNULL(tm03) & ", " & CNULL(tm04) & ", " & CNULL(m_CaseNa239(1)) & ", " & CNULL(m_CaseNa239(2)) & ", " & CNULL(m_CaseNa239(3)) & ", " & CNULL(m_CaseNa239(4)) & " ) "
             cnnConnection.Execute strSql
             strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_CaseNa239(1)) & ", " & CNULL(m_CaseNa239(2)) & ", " & CNULL(m_CaseNa239(3)) & ", " & CNULL(m_CaseNa239(4)) & ", " & CNULL(tm01) & ", " & CNULL(tm02) & ", " & CNULL(tm03) & ", " & CNULL(tm04) & " ) "
             cnnConnection.Execute strSql
             'Added by Lydia 2020/12/04 歐盟案案件備註加註「英國案案號」；新英國案之新案收文的進度備註加註「歐盟案案號」
             strSql = "Update Trademark set TM58=" & CNULL("英國案案號：" & tm01 & tm02 & tm03 & tm04 & ";") & "||TM58 where tm01='" & m_CaseNa239(1) & "' and tm02='" & m_CaseNa239(2) & "' and tm03='" & m_CaseNa239(3) & "' and tm04='" & m_CaseNa239(4) & "' "
             cnnConnection.Execute strSql
             strSql = "Update CaseProgress set CP64=" & CNULL("歐盟案案號：" & m_CaseNa239(1) & m_CaseNa239(2) & m_CaseNa239(3) & m_CaseNa239(4) & ";") & "||CP64 where CP09='" & CP09 & "' "
             cnnConnection.Execute strSql
             'end 2020/12/04
             'Added by Lydia 2021/01/11 複製優先權資料
             strSql = "insert into pridate (pd01,pd02,pd03,pd04,pd05,pd06,pd07,pd08,pd09,pd10) " & _
                          "select " & CNULL(tm01) & ", " & CNULL(tm02) & ", " & CNULL(tm03) & ", " & CNULL(tm04) & ", pd05,pd06,pd07,pd08,pd09,pd10 from pridate where pd01='" & m_CaseNa239(1) & "' and pd02='" & m_CaseNa239(2) & "' and pd03='" & m_CaseNa239(3) & "' and pd04='" & m_CaseNa239(4) & "' "
             cnnConnection.Execute strSql
             'end 2021/01/11
             'Added by Lydia 2021/04/15 CFT英國脫歐委任代理之後續處理：收文英國延展及委任代理人新案，同時將代理人存入CP44。
             strExc(0) = "select np01,np15 from nextprogress where np07='710' and np15 like '%脫歐英國案代理人：%' " & _
                              "and ((np02='" & tm01 & "' and np03='" & tm02 & "' and np04='" & tm03 & "' and np05='" & tm04 & "') or (np02='" & m_CaseNa239(1) & "' and np03='" & m_CaseNa239(2) & "'  and np04='" & m_CaseNa239(3) & "'  and np05='" & m_CaseNa239(4) & "')) "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
                 strExc(1) = Mid("" & RsTemp.Fields("np15"), InStr(RsTemp.Fields("np15"), "脫歐英國案代理人：") + 9, 9)
                 If Left(strExc(1), 1) = "Y" Then
                     strSql = "Update CaseProgress set cp44=" & CNULL(strExc(1)) & " where cp09=" & CNULL(CP09)
                     cnnConnection.Execute strSql
                 End If
             End If
             'end 2021/04/15
             PUB_EUtoUK tm01, tm02, tm03, tm04, m_CaseNa239(1), m_CaseNa239(2), m_CaseNa239(3), m_CaseNa239(4), CP09, CP10 'Added by Morgan 2020/12/21 回覆單歸卷
        End If
        'end 2020/11/19
        
        'Added by Lydia 2021/03/05 CFT歐盟尚未註冊案轉換英國申請案收文控管：一併將歐盟關聯案之「商標圖樣」、「商品/服務類別及名稱」、「申請日」、「優先權資料」帶入新案號
        If tm01 = "CFT" And cp31 = "Y" And m_CaseNa239(1) <> "" And m_CaseNa239(2) <> "" And CP10 = "101" Then
             '建立歐盟案及英國案之關聯(相關卷號)
             strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(tm01) & ", " & CNULL(tm02) & ", " & CNULL(tm03) & ", " & CNULL(tm04) & ", " & CNULL(m_CaseNa239(1)) & ", " & CNULL(m_CaseNa239(2)) & ", " & CNULL(m_CaseNa239(3)) & ", " & CNULL(m_CaseNa239(4)) & " ) "
             cnnConnection.Execute strSql
             strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_CaseNa239(1)) & ", " & CNULL(m_CaseNa239(2)) & ", " & CNULL(m_CaseNa239(3)) & ", " & CNULL(m_CaseNa239(4)) & ", " & CNULL(tm01) & ", " & CNULL(tm02) & ", " & CNULL(tm03) & ", " & CNULL(tm04) & " ) "
             cnnConnection.Execute strSql
             '複製優先權資料
             strSql = "insert into pridate (pd01,pd02,pd03,pd04,pd05,pd06,pd07,pd08,pd09,pd10) " & _
                          "select " & CNULL(tm01) & ", " & CNULL(tm02) & ", " & CNULL(tm03) & ", " & CNULL(tm04) & ", pd05,pd06,pd07,pd08,pd09,pd10 from pridate where pd01='" & m_CaseNa239(1) & "' and pd02='" & m_CaseNa239(2) & "' and pd03='" & m_CaseNa239(3) & "' and pd04='" & m_CaseNa239(4) & "' "
             cnnConnection.Execute strSql
        End If
        'end 2021/03/05
        
        'Added by Lydia 2021/04/15 CFT英國脫歐委任代理之後續處理：收文英國延展及委任代理人新案，同時將代理人存入CP44。
                                                 '同一天接洽單之後收文的處理
        If txtSystem = "CFT" And cp31 <> "Y" And txtTrademark(5) = "201" And (txtTrademark(1) = "710" Or txtTrademark(1) = "102") And m_TM58 <> "" And InStr(m_TM58, "歐盟案案號：") > 0 Then
             strExc(0) = "select cp44 from caseprogress where cp01='" & tm01 & "' and cp02='" & tm02 & "' and cp03='" & tm03 & "' and cp04='" & tm04 & "' and cp05=" & DBDATE(txtTrademark(0)) & " and cp10 in ('102','710') and cp159=0 and cp31='Y' "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
                 If "" & RsTemp.Fields("cp44") <> "" Then
                     strSql = "Update CaseProgress set cp44=" & CNULL(RsTemp.Fields("cp44")) & " where cp09=" & CNULL(CP09)
                     cnnConnection.Execute strSql
                 End If
             End If
        End If
        'end 2021/04/15
        
        'Added by Lydia 2020/12/15 CFT緬甸重新申請案：建立關聯
        If txtSystem = "CFT" And txtTrademark(5) = "048" And txtTrademark(1) = "101" And txtCFTNa048.Visible = True And txtCFTNa048.Text <> "" Then
             strSql = txtCFTNa048.Text
             strExc(1) = "": strExc(2) = "": strExc(3) = "": strExc(4) = ""
             Call ChgCaseNo(strSql, strExc)
             strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(tm01) & ", " & CNULL(tm02) & ", " & CNULL(tm03) & ", " & CNULL(tm04) & ", " & CNULL(strExc(1)) & ", " & CNULL(strExc(2)) & ", " & CNULL(strExc(3)) & ", " & CNULL(strExc(4)) & " ) "
             cnnConnection.Execute strSql
             strSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(strExc(1)) & ", " & CNULL(strExc(2)) & ", " & CNULL(strExc(3)) & ", " & CNULL(strExc(4)) & ", " & CNULL(tm01) & ", " & CNULL(tm02) & ", " & CNULL(tm03) & ", " & CNULL(tm04) & " ) "
             cnnConnection.Execute strSql
             'Added by Lydia 2021/02/01 複製商品/服務類別及名稱、優先權資料
             strSql = "INSERT INTO TMGOODS(TG01,TG02,TG03,TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17) " & _
                         "SELECT '" & tm01 & "' as TG01, '" & tm02 & "' as TG02, '" & tm03 & "' as  TG03, '" & tm04 & "' as  TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17 " & _
                         "FROM TMGOODS WHERE TG01='" & strExc(1) & "'  AND TG02='" & strExc(2) & "' AND TG03='" & strExc(3) & "'  AND TG04='" & strExc(4) & "'  AND TG18 IS NULL "
             cnnConnection.Execute strSql
             strSql = "insert into pridate (pd01,pd02,pd03,pd04,pd05,pd06,pd07,pd08,pd09,pd10) " & _
                          "select " & CNULL(tm01) & ", " & CNULL(tm02) & ", " & CNULL(tm03) & ", " & CNULL(tm04) & ", pd05,pd06,pd07,pd08,pd09,pd10 from pridate where pd01='" & strExc(1) & "' and pd02='" & strExc(2) & "' and pd03='" & strExc(3) & "' and pd04='" & strExc(4) & "' "
             cnnConnection.Execute strSql
             'end 2021/02/01
        End If
        'end 2020/12/15
        
         strSql = "update customer set cu30=" + CNULL(cu30) + " where cu01=" + CNULL(Mid(tm23, 1, 8)) + " and cu02=" + CNULL(Mid(tm23, 9, 1))
         cnnConnection.Execute strSql
         '92.2.19 modify by sonia先抓NP06=NULL, 沒有資料才抓NU06<>'Y', 只有一筆才更新
         'If bolRt Then
         '  'Move By Cheng 2002/12/18
         '  adoquery.CursorLocation = adUseClient
         '  adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
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
         '   strSQL = "update nextprogress set np06='Y' where np02=" + CNULL(tm01) + " and np03=" + _
         '       CNULL(tm02) + " and np04=" + CNULL(tm03) + " and np05=" + CNULL(tm04) + _
         '       " and np07=" + CNULL(CP10) + " and (np06<>'Y' or np06 is null)"
         '   cnnConnection.Execute strSQL
         'End If
         adoquery.CursorLocation = adUseClient
         'add by nickc 2007/10/24 內商全期註冊費(717)時，抓第一期註冊費(715)
         'edit by nickc 2007/10/25 加入外商
         'If tm01 = "T" And CP10 = "717" Then
         'Modify By Sindy 2012/7/13
'         If (tm01 = "T" Or tm01 = "FCT") And CP10 = "717" Then
'               adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np06 is null and np07 =715 ", cnnConnection, adOpenStatic, adLockReadOnly
         If (tm01 = "T" Or tm01 = "FCT") And (CP10 = "717" Or CP10 = "715") Then
               adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np06 is null and np07 in('715','717') ", cnnConnection, adOpenStatic, adLockReadOnly
         Else
            'Add By Sindy 2012/7/4 馬德里使用宣誓
            If tm01 = "TF" And CP10 = "105" Then
               adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and substr(np03,1,5) = '" & Left(tm02, 5) & "' and np06 is null and np07 = '" & CP10 & "' and np08=" & IIf(DBDATE(txtTrademark(11)) = "", 0, DBDATE(txtTrademark(11))), cnnConnection, adOpenStatic, adLockReadOnly
            Else
            '2012/7/4 End
               adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
            End If
         End If
         If adoquery.RecordCount > 0 Then
            'Modify By Sindy 2012/7/4 馬德里使用宣誓有可能多筆
            'If adoquery.RecordCount = 1 Then
            If (adoquery.RecordCount = 1 Or (tm01 = "TF" And CP10 = "105" And adoquery.RecordCount >= 1)) Then
            '2012/7/4 End
               If IsNull(adoquery.Fields(0).Value) = False Then
                  '2011/6/16 add by sonia 異議答辯、評定答辯、廢止答辯要一並更新對造資料
                  If (CP10 = "602" Or CP10 = "604" Or CP10 = "606") Then
                     cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & CP09 & "'", intI
                  Else
                  '2011/6/16 end
                     cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & CP09 & "'"
                  End If  '2011/6/16 add by sonia
               End If
               'add by nick 2004/09/08
               If txtTrademark(1).Text <> "305" Then
                   'add by nickc 2007/10/24 內商全期註冊費(717)時，抓第一期註冊費(715)
                   'edit by nickc 2007/10/25 加入外商
                   'If tm01 = "T" And CP10 = "717" Then
                   'Modify By Sindy 2012/7/13
'                   If (tm01 = "T" Or tm01 = "FCT") And CP10 = "717" Then
'                       strSql = "update nextprogress set np06='Y' where np02=" + CNULL(tm01) + " and np03=" + _
'                          CNULL(tm02) + " and np04=" + CNULL(tm03) + " and np05=" + CNULL(tm04) + _
'                          " and np07=715 and np06 is null"
                   'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
                   If (tm01 = "T" Or tm01 = "FCT") And (CP10 = "717" Or CP10 = "715") Then
                       strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" + CNULL(tm01) + " and np03=" + _
                          CNULL(tm02) + " and np04=" + CNULL(tm03) + " and np05=" + CNULL(tm04) + _
                          " and np07 in('715','717') and np06 is null"
                       cnnConnection.Execute strSql
                   Else
                     'Add By Sindy 2012/7/4 馬德里使用宣誓
                     If tm01 = "TF" And CP10 = "105" Then
                        'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
                        strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" + CNULL(tm01) + " and substr(np03,1,5)=" + _
                           CNULL(Left(tm02, 5)) + _
                           " and np07=" + CNULL(CP10) + " and np06 is null and np08=" & IIf(DBDATE(txtTrademark(11)) = "", 0, DBDATE(txtTrademark(11))) + _
                           " and np02||np03||np04||np05 in(select tm01||tm02||tm03||tm04 from trademark where tm01=np02 and tm02=np03 and tm03=np04 and tm04=np05 and tm29 is null)"
                        cnnConnection.Execute strSql
                     Else
                     '2012/7/4 End
                        'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
                        strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" + CNULL(tm01) + " and np03=" + _
                           CNULL(tm02) + " and np04=" + CNULL(tm03) + " and np05=" + CNULL(tm04) + _
                           " and np07=" + CNULL(CP10) + " and np06 is null"
                        cnnConnection.Execute strSql
                     End If
                   End If
               End If
            End If
         Else
            adoquery.Close
            adoquery.CursorLocation = adUseClient
            'add by nickc 2007/10/24 內商全期註冊費(717)時，抓第一期註冊費(715)
            'edit by nickc 2007/10/25 加入外商
            'If tm01 = "T" And CP10 = "717" Then
            'Modify By Sindy 2012/7/13
'            If (tm01 = "T" Or tm01 = "FCT") And CP10 = "717" Then
'               adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np06 <>'Y' and np07 = 715 ", cnnConnection, adOpenStatic, adLockReadOnly
            If (tm01 = "T" Or tm01 = "FCT") And (CP10 = "717" Or CP10 = "715") Then
               adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np06 <>'Y' and np07 in('715','717') ", cnnConnection, adOpenStatic, adLockReadOnly
            Else
               'Add By Sindy 2012/7/4 馬德里使用宣誓
               If tm01 = "TF" And CP10 = "105" Then
                  adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and substr(np03,1,5) = '" & Left(tm02, 5) & "' and np06 <>'Y' and np07 = '" & CP10 & "' and np08=" & IIf(DBDATE(txtTrademark(11)) = "", 0, DBDATE(txtTrademark(11))), cnnConnection, adOpenStatic, adLockReadOnly
               Else
               '2012/7/4 End
                  adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np06 <>'Y' and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
               End If
            End If
            If adoquery.RecordCount > 0 Then
               'Modify By Sindy 2012/7/4 馬德里使用宣誓有可能多筆
               'If adoquery.RecordCount = 1 Then
               If (adoquery.RecordCount = 1 Or (tm01 = "TF" And CP10 = "105" And adoquery.RecordCount >= 1)) Then
               '2012/7/4 End
                  If IsNull(adoquery.Fields(0).Value) = False Then
                     '2011/6/16 add by sonia 異議答辯、評定答辯、廢止答辯要一並更新對造資料
                     If (CP10 = "602" Or CP10 = "604" Or CP10 = "606") Then
                        cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & CP09 & "'", intI
                     Else
                     '2011/6/16 end
                        cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & CP09 & "'"
                     End If  '2011/6/16 add by sonia
                  End If
                  'add by nick 2004/09/08
                  If txtTrademark(1).Text <> "305" Then
                       'add by nickc 2007/10/24 內商全期註冊費(717)時，抓第一期註冊費(715)
                       'edit by nickc 2007/10/25 加入外商
                       'If tm01 = "T" And CP10 = "717" Then
                       'Modify By Sindy 2012/7/13
'                       If (tm01 = "T" Or tm01 = "FCT") And CP10 = "717" Then
'                           strSql = "update nextprogress set np06='Y' where np02=" + CNULL(tm01) + " and np03=" + _
'                              CNULL(tm02) + " and np04=" + CNULL(tm03) + " and np05=" + CNULL(tm04) + _
'                              " and np07=715 and np06 <> 'Y'"
                       If (tm01 = "T" Or tm01 = "FCT") And (CP10 = "717" Or CP10 = "715") Then
                           'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
                           strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" + CNULL(tm01) + " and np03=" + _
                              CNULL(tm02) + " and np04=" + CNULL(tm03) + " and np05=" + CNULL(tm04) + _
                              " and np07 in('715','717') and np06 <> 'Y'"
                           cnnConnection.Execute strSql
                       Else
                           'Add By Sindy 2012/7/4 馬德里使用宣誓
                           If tm01 = "TF" And CP10 = "105" Then
                              'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
                              strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" + CNULL(tm01) + " and substr(np03,1,5)=" + _
                                 CNULL(Left(tm02, 5)) + _
                                 " and np07=" + CNULL(CP10) + " and np06 <> 'Y' and np08=" & IIf(DBDATE(txtTrademark(11)) = "", 0, DBDATE(txtTrademark(11))) + _
                                 " and np02||np03||np04||np05 in(select tm01||tm02||tm03||tm04 from trademark where tm01=np02 and tm02=np03 and tm03=np04 and tm04=np05 and tm29 is null)"
                              cnnConnection.Execute strSql
                           Else
                           '2012/7/4 End
                              'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(CP09) & "
                              strSql = "update nextprogress set np06='Y',np24=" & CNULL(CP09) & " where np02=" + CNULL(tm01) + " and np03=" + _
                                 CNULL(tm02) + " and np04=" + CNULL(tm03) + " and np05=" + CNULL(tm04) + _
                                 " and np07=" + CNULL(CP10) + " and np06 <> 'Y'"
                              cnnConnection.Execute strSql
                           End If
                       End If
                   End If
               End If
            End If
         End If
         adoquery.Close
         '92.2.19 END
         
         ' Add by Sindy 98/03/02
         '收文時若讓案號下一程序仍有F4103且是否續辦為NULL者,
         '更新下一程序F4103為收文智權人員
         If tm01 = "FCT" Then
            strSql = "update nextprogress set np10='" & txtTrademark(12) & "' " & _
               "where np02=" + CNULL(tm01) + " and np03=" + _
               CNULL(tm02) + " and np04=" + CNULL(tm03) + " and np05=" + CNULL(tm04) + _
               " and np10='F4103' and np06 is null"
            cnnConnection.Execute strSql
         End If
         ' 98/03/02 End
         
         'Added by Morgan 2021/6/22
         'T與FCT共同控管案件通知
         'Modified by Morgan 2023/5/4 改在PUB_2SysCaseInform內抓系統特殊設定
         'If (tm01 = "T" And (tm02 = "211948" Or tm02 = "211949") And tm03 = "0" And tm04 = "00") Or (tm01 = "FCT" And (tm02 = "047561" Or tm02 = "047562") And tm03 = "0" And tm04 = "00") Then
         If tm01 = "T" Or tm01 = "FCT" Then
         'end 2023/5/4
            PUB_2SysCaseInform tm01, tm02, tm03, tm04, CP09, 2
         End If
         'end 2021/6/22
      
         'edit by nickc 2007/02/06 不用 dll 了
         'If obj001.SetCaseProgressFee(tm01, tm10, CP10, CP09) = False Then bolError = True
         If Cls001SetCaseProgressFee(tm01, tm10, CP10, CP09) = False Then bolError = True
      Else
         bolError = True
      End If
   End If
   'Modify By Cheng 2002/12/18
   '更新動作往前移
   'adoquery.CursorLocation = adUseClient
   ''adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np07 = '" & cp10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
   'adoquery.Open "select np01 from nextprogress where np02 = '" & tm01 & "' and np03 = '" & tm02 & "' and np04 = '" & tm03 & "' and np05 = '" & tm04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
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
   If bolError = False Then
      'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'      Dim rtCnt As Integer
'      'Modify by Morgan 2010/12/9
'      'If txtTrademark(34) <> "" Then
'      '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " ", rtCnt
'      If txtTrademark(34) <> "" And txtTrademark(34) <> txtTrademark(34).Tag Then
'          cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
'      'end 2010/12/9
'          If rtCnt = 0 Then
'              cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from dual "
'          End If
'      End If
      'end 2018/08/22
      
      'Modified by Lydia 2022/08/22 改成共用模組
      'Call SaveFrame21(CP09) 'Add By Sindy 2012/5/8
      Call GetStrControl
      'Modified by Lydia 2022/09/29 傳入系統別,國家,案件性質 => tm01, tm10, CP10
      Call PUB_SaveByControl(CP09, m_strControl, tm01, tm10, CP10)
      'end 2022/08/22
   End If
   
   If bolError Then
      cnnConnection.RollbackTrans
      ShowMsg MsgText(9004)
      'add by nickc 2007/12/12
   IsSaveData = False
   Else
      cnnConnection.CommitTrans
      InsertTrademarkDatabase = True
      'add by nickc 2006/03/27
       If tm01 = "TF" Then
          txtTFCode(0) = Mid(tm02, 1, 5)
          txtTFCode(1) = Mid(tm02, 6, 1)
       Else
          txtCode(0) = tm02
       End If
   End If
   'edit by nickc 2007/02/06 不用 dll 了
   'Set objPublicData = Nothing
   Exit Function
ErrHand:
   'edit by nickc 2007/02/06 不用 dll 了
   'Set objPublicData = Nothing
   cnnConnection.RollbackTrans
   'edit by nickc 2006/03/07 解決 cp02=null 的問題
   'add by nickc 2005/08/25
   'If tm01 = "TF" Then
   '   txtTFCode(0) = ""
   '   txtTFCode(1) = ""
   'Else
   '   txtCode(0) = ""
   'End If
   ShowMsg MsgText(9004)
   'add by nickc 2007/12/12
   IsSaveData = False
End Function

'Add By Sindy 2012/5/8
Private Sub SaveFrame21(strCP09 As String)
   If Frame21.Visible = True Then
      '資料是否齊備
      'Memo by Lydia 2018/12/10 T台灣商申案=文件是否齊備
      If textEP06.Visible = True Then  'Added by Lydia 2018/12/10 +判斷顯示
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
      If textEP34.Visible = True Then 'Added by Lydia 2018/12/10 +判斷顯示
            strSql = "update engineerprogress set ep34='" & textEP34 & "' where ep02='" & strCP09 & "'"
            cnnConnection.Execute strSql
      End If
      '是否急件
      If textCP122.Visible = True Then 'Added by Lydia 2018/12/10 +判斷顯示
            strSql = "update caseprogress set cp122='" & textCP122 & "' where cp09='" & strCP09 & "'"
            cnnConnection.Execute strSql
      End If
      'Added by Lydia 2018/12/10 查名是否齊備
      If textCP143.Visible = True Then
            strSql = "update caseprogress set cp143='" & IIf(textCP143 = "Y", strSrvDate(1), IIf(textCP143 = "N", "0", "")) & "' where cp09='" & strCP09 & "'"
            cnnConnection.Execute strSql
      End If
   End If
End Sub

'從下一程序檔取回本所期限、法定期限
Private Function GetNextProgressDate(ByVal np02 As String, ByVal np03 As String, ByVal np04 As String, _
       ByVal np05 As String, ByVal NP07 As String, ByRef strDate1 As String, ByRef strDate2 As String, ByRef strNP15 As String, _
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
   'Add By Sindy 2012/7/4 馬德里使用宣誓
   If np02 = "TF" And NP07 = "105" Then
      strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and substr(np03,1,5)=" + _
                CNULL(Left(np03, 5)) + _
                " and np07 in(" + NP07 + ") and np06 is null order by np08 asc "
   Else
   '2012/7/4 End
      strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
                CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
                " and np07 in(" + NP07 + ") and np06 is null "
   End If
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection, adOpenStatic
   If rsRecordset.RecordCount > 0 Then
      rsRecordset.MoveFirst
      'Modify By Sindy 2012/7/4 馬德里使用宣誓有可能多筆
      'If rsRecordset.RecordCount = 1 Then
      If (rsRecordset.RecordCount = 1 Or (np02 = "TF" And NP07 = "105" And rsRecordset.RecordCount >= 1)) Then
      '2012/7/4 End
         strDate1 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
         strDate2 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
         strNP15 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
         'Add By Cheng 2001/12/17
         strNP10 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
         'Added by Lydia 2017/12/21 檢查是否有不續辦相同性質且未到期的期限，若有則提醒操作人員注意要輸入接洽單上填寫的期限
         If frm010001.mRole = "" Then 'Added by Lydia 2024/10/18 排除外專/外商自行收文
             strExc(1) = Pub_GetNPDoubleMsg(DBDATE(txtTrademark(0).Text), np02, np03, np04, np05, NP07)
             If strExc(1) <> "" Then MsgBox strExc(1), vbExclamation + vbOKOnly
         End If 'Added by Lydia 2024/10/18
         'end 2017/12/21
      End If
   Else
      '911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
      'Add By Sindy 2012/7/4 馬德里使用宣誓
      If np02 = "TF" And NP07 = "105" Then
         strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and substr(np03,1,5)=" + _
                   CNULL(Left(np03, 5)) + _
                  " and np07 in(" + NP07 + ") and np06 <>'Y' order by np08 asc "
      Else
      '2012/7/4 End
         strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
                   CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
                  " and np07 in(" + NP07 + ") and np06 <>'Y' "
      End If
      Set rsRecordset = New ADODB.Recordset
      rsRecordset.CursorLocation = adUseClient
      rsRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      'Modify By Sindy 2012/7/4 馬德里使用宣誓有可能多筆
      'If rsRecordset.RecordCount = 1 Then
      If (rsRecordset.RecordCount = 1 Or (np02 = "TF" And NP07 = "105" And rsRecordset.RecordCount >= 1)) Then
         rsRecordset.MoveFirst
      '2012/7/4 End
         strDate1 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
         strDate2 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
         strNP15 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
         '取得下一程序資料檔之智權人員代號
         strNP10 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
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
Dim objTxt 'Modify by Amy 2021/12/16 原:As TextBox
Dim ii As Integer
Dim Cancel As Boolean
'Add by Amy 2017/03/08
Dim strTmp(0) As String, strMCTF(0) As String, strMsg As String, strApply As String
Dim bolData As Boolean

   TxtValidate = False
   
   'Add by Amy 2013/07/19 lblCaseProperty顯示（無）不可以存檔
   If lblCaseProperty = "（無）" Then
      MsgBox "案件性質錯誤!!", vbExclamation
      Exit Function
   End If
   'end 2013/07/19
   
   For Each objTxt In Me.txtTrademark
      'Add By Cheng 2002/06/11
      If objTxt.Index = 21 And Me.fraPatition.Visible = False Then GoTo NextTxt
      
      If objTxt.Enabled = True Then
         If Not (objTxt.Index = 9 Or objTxt.Index = 24 Or objTxt.Index = 25 Or objTxt.Index = 26 Or objTxt.Index = 27) Then 'Added by Lydia 2024/02/16 因為在cmdOK_Click已呼叫CheckKeyin檢查
            Cancel = False
            txtTrademark_Validate objTxt.Index, Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If  'Added by Lydia 2024/02/16
      End If
NextTxt:
   Next
   
'cancel by sonia 2021/11/25 MCT案件不再檢查申請人與代理人MCT組別是否相同
'   'Add by Amy 2017/03/08 MCTF組別控制(有輸代理人且為MCTF,判斷申請人若與代理人的MCTF組別不同不可收文)
'   If Len(Trim(txtTrademark(10))) > 0 Then
'        bolData = GetCusORFagentData(ChangeCustomerL(txtTrademark(10)), "FA120", strMCTF())
'        If Left(strMCTF(0), 4) = "MCTF" Then
'             For ii = 0 To 4
'                 If ii = 0 Then
'                     strApply = txtTrademark(9)
'                 Else
'                     strApply = txtTrademark(ii + 23)
'                 End If
'                 If strApply = MsgText(601) Then Exit For
'                 bolData = GetCusORFagentData(ChangeCustomerL(strApply), "CU13", strTmp())
'                 If strMCTF(0) <> strTmp(0) And Left(strTmp(0), 4) = "MCTF" Then
'                     strMsg = strMsg & "申請人" & ii + 1 & "：" & strApply & " (" & strTmp(0) & ")" & "及"
'                 End If
'             Next ii
'             If strMsg <> MsgText(601) Then
'                 MsgBox Left(strMsg, Len(strMsg) - 1) & vbCrLf & "與代理人" & txtTrademark(10) & _
'                                "商標管控智權人員(" & strMCTF(0) & ")不同，請退回智權人員！"
'                 Exit Function
'             End If
'        End If
'   End If
'   'end 2017/03/08
   
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

'取得規費
Private Function GetOfficalFee(strCF01 As String, strCF02 As String, strCF03 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   GetOfficalFee = ""
   StrSQLa = "Select CF08 From CaseFee Where CF01='" & strCF01 & "' And CF02='" & strCF02 & "' And CF03='" & strCF03 & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       GetOfficalFee = Val("" & rsA.Fields(0).Value)
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
End Function

'add by nickc 2007/10/16 檢查星期四五收的文，期限是否為假日
Function ChkMyWeek(oDate As String) As Boolean
   ChkMyWeek = False
   If Weekday(ChangeWStringToWDateString(strSrvDate(1))) = 5 Or Weekday(ChangeWStringToWDateString(strSrvDate(1))) = 6 Then
       If ChkWorkDay(DBDATE(oDate)) = False Then
           If GetWorkDay(DBDATE(oDate), strSrvDate(1)) <= 2 Then
               ChkMyWeek = True
           End If
       End If
   End If
End Function

'Add By Sindy 2022/12/7 證書形式
Private Sub SetTM136()
   
   Label1(141).Visible = False
   txtTrademark(35).Visible = False
   Label28(1).Visible = False
   '台灣T,FCT案717.註冊費收文,下列欄位必須顯示出來
   If txtTrademark(5) = "000" Then
      'Modified by Morgan 2022/12/26
      'If (txtSystem = "T" Or txtSystem = "FCT") And Trim(txtTrademark(1)) = "717" Then
      If PUB_TWCertPty(txtSystem, txtTrademark(1), txtCode(0), txtCode(1), txtCode(2)) = True Then
      'end 2022/12/26
         Label1(141).Visible = True
         txtTrademark(35).Visible = True
         Label28(1).Visible = True
      End If
   End If
End Sub

'Modify By Sindy 2012/7/23 移出為獨立函數
Private Sub setFrame21()
   If txtTrademark(5).Text <> txtTrademark(5).Tag Then 'Added by Lydia 2018/12/10 判斷有修改才重設
        'Add By Sindy 2012/5/8
        '台灣商標Ｔ,FCT案若收文爭議案件性質時,開放Frame21欄位
        Frame21.Visible = False
        m_EP06 = "": textEP34.Enabled = True
        'Modifiec by Lydia 2018/12/10 T台灣案填寫接洽單管控文件及查名是否齊備
        'If (txtSystem = "T" Or txtSystem = "FCT") And txtTrademark(5) = "000" And InStr(TMdebate, txtTrademark(1)) > 0 Then
        Label39.Visible = False: textEP34.Visible = False '預設會稿、查名不顯示
        Label42.Visible = False: textCP143.Visible = False
        'Modified by Lydia 2022/07/15 T大陸案之齊備日管控
        'If ((txtSystem = "T" Or txtSystem = "FCT") And txtTrademark(5) = "000" And InStr(TMdebate, txtTrademark(1)) > 0) _
              Or (txtSystem = "T" And txtTrademark(5) = "000") Then
        'Modify By Sindy 2025/7/28 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
        If ((txtSystem = "T" Or txtSystem = "FCT") _
            And txtTrademark(5) = "000" _
            And InStr(TMdebate, txtTrademark(1)) > 0 _
            And Not (txtSystem = "FCT" And InStr(FCT_NotTMdebate, txtTrademark(1)) > 0)) _
           Or (txtSystem = "T" And (txtTrademark(5) = "000" Or txtTrademark(5) = "020")) Then
        'end 2018/12/10
           Frame21.Visible = True
           'Added by Lydia 2018/12/10 區分商爭和商申
           If txtSystem = "T" Then
             If InStr(TMdebate, txtTrademark(1)) > 0 Then   '商爭
                 Label39.Visible = True: textEP34.Visible = True
                 Label41.Caption = "資料是否齊備：       (Y/N)"
             Else
                 If txtTrademark(1) = 申請 Then  '商申
                    Label42.Visible = True: textCP143.Visible = True
                 End If
                 Label41.Caption = "文件是否齊備：       (Y/N)"
             End If
           ElseIf txtSystem = "FCT" Then
             If InStr(TMdebate, txtTrademark(1)) > 0 Then   '商爭
                 Label39.Visible = True: textEP34.Visible = True
                 Label41.Caption = "資料是否齊備：       (Y/N)"
             End If
           End If
           'end 2018/12/10
           If frm010001.intModifyKind = 0 Then
              If txtSystem = "T" And InStr(TMdebate, txtTrademark(1)) > 0 Then 'Added by Lydia 2018/12/10 區分商爭和商申
                 If Val(strSrvDate(1)) < Val(TMdebateStarDT) Then
                    textEP06.Text = "N"
                    textCP122.Text = "N"
                    textEP34.Text = "N"
                 End If
                 '案件性質為613補充答辯或612補充理由時，則只可不會稿
                 If txtTrademark(1) = "613" Or txtTrademark(1) = "612" Then
                    textEP34.Text = "N"
                    textEP34.Enabled = False
                 End If
              End If
           Else
              '讀取資料
              'Modified by Lydia 2018/12/10 +T案查名齊備日CP143
              strSql = "SELECT ep06,ep34,cp122,cp143 FROM engineerprogress,caseprogress WHERE cp09='" & txtRecieveCode & "' and cp09=ep02(+)"
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
                    '案件性質為613補充答辯或612補充理由時，則只可不會稿
                    If txtTrademark(1) = "613" Or txtTrademark(1) = "612" Then
                       If textEP34.Text = "N" Then 'Modify By Sindy 2013/3/12 +if
                          textEP34.Enabled = False
                       End If
                    End If
                 End If
                 If Not IsNull(RsTemp.Fields("cp122")) Then
                    textCP122.Text = RsTemp.Fields("cp122")
                 End If
                 'Added by Lydia 2018/12/10 T案查名齊備日
                 If Val("" & RsTemp.Fields("cp143")) > 0 Then
                    textCP143.Text = "Y"
                 Else
                    textCP143.Text = "N"
                 End If
                 'end 2018/12/10
              End If
           End If
           m_EP06 = textEP06
        End If
        '2012/5/8 End
               
  'Added by Lydia 2018/12/10
        txtTrademark(5).Tag = txtTrademark(5).Text
   End If
   'end 2018/12/10
End Sub
'Added by Lydia 2015/11/12 新增查名單對應
'Memo by Lydia 2016/04/27 改成直接在畫面輸入查名代號
Private Sub cmdTSMap_Click()
  'Added by Lydia 2015/11/11 要先輸入智權人員(預設查詢)
  If txtTrademark(12).Text = "" Or lblSales.Caption = "" Then
     MsgBox "請先輸入智權人員!", vbInformation
     Exit Sub
  End If
  
  bolOpen130 = True 'Added by Lydia 2016/03/28
  
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
    Tmpfrm090130.m_CP13 = txtTrademark(12).Text '智權人員(預設查詢)
    Tmpfrm090130.Show
    Tmpfrm090130.Caption = cmdTSMap.Caption
    Me.Hide
End Sub
'end 2015/11/12

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
    'end 2016/05/11
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
        '    t_LOSkind = PUB_GetLOSkind(txtSystem, txtTrademark(1), txtTrademark(5))
        'end 2020/06/10
        End If
    End If
    Set rsRd = Nothing
End Sub

'Added by Lydia 2020/12/15
Private Sub txtCFTNa048_GotFocus()
    TextInverse txtCFTNa048
End Sub

Private Sub txtCFTNa048_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCFTNa048_Validate(Cancel As Boolean)
   If txtCFTNa048.Tag <> txtCFTNa048.Text And txtCFTNa048.Text <> "" Then
        If Left(txtCFTNa048, 3) <> txtSystem Then
            MsgBox "請輸入" & txtSystem & "案！", vbCritical, "檢核資料"
            GoTo EXITSUB
        End If
        strExc(0) = Left(txtCFTNa048.Text & String(8, "0"), 12)
        Call ChgCaseNo(strExc(0), strExc)
        If Len("" & strExc(2)) <> 6 Then
            MsgBox "請輸入正確的緬甸舊案號！", vbCritical, "檢核資料"
            GoTo EXITSUB
        Else
            'Modified by Lydia 2021/02/01 +商品類別TM09
            strSql = "SELECT TM01,TM02,TM03,TM04,TM23,TM09 FROM TRADEMARK WHERE TM01='" & strExc(1) & "' AND TM02='" & strExc(2) & "' AND TM03='" & strExc(3) & "' AND TM04='" & strExc(4) & "' " & _
                        "AND TM11<20210101 AND TM15 IS NOT NULL AND TM10='048' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then
                MsgBox "請輸入正確的緬甸舊案號！", vbCritical, "檢核資料"
                GoTo EXITSUB
            ElseIf txtTrademark(9).Text <> "" Then
                If "" & RsTemp.Fields("TM23") <> ChangeCustomerL(txtTrademark(9).Text) Then
                    MsgBox "該案號不屬於申請人１的案件！", vbCritical, "檢核資料"
                    GoTo EXITSUB
                End If
            End If
            txtTrademark(4).Text = "" & RsTemp.Fields("TM09")  'Added by Lydia 2021/02/01 自動帶入商品類別
            txtCFTNa048.Text = strExc(1) & strExc(2) & strExc(3) & strExc(4)
        End If
   End If
   
   txtCFTNa048.Tag = txtCFTNa048.Text
   
   Exit Sub
   
EXITSUB:
   Cancel = True
   txtCFTNa048.SetFocus
   txtCFTNa048_GotFocus
End Sub

'Added by Lydia 2020/12/15 CFT緬甸重新申請案：檢查
Private Function CheckCFTna048() As Boolean
Dim bolTmp As Boolean

    If txtSystem = "CFT" And txtTrademark(5) = "048" And txtTrademark(1) = "101" And txtCFTNa048.Visible = True Then
        If txtCFTNa048 <> "" Then
            '重新檢查
            txtCFTNa048.Tag = ""
            Call txtCFTNa048_Validate(bolTmp)
            If bolTmp = True Then
                Exit Function
            End If
        Else
            strSql = "SELECT TM01,TM02,TM03,TM04,TM23 FROM TRADEMARK WHERE TM23=" & CNULL(ChangeCustomerL(txtTrademark(9).Text)) & _
                        "AND TM11<20210101 AND TM15 IS NOT NULL AND TM10='048' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
                If MsgBox("接洽單左上角是否有相關緬甸商標註冊案？", vbYesNo + vbInformation + vbDefaultButton1, "CFT緬甸重新申請案") = vbYes Then
                    txtCFTNa048.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    
    CheckCFTna048 = True
   
End Function

