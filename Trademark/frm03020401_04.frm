VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020401_04 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "«Dª§Ä³®×®Ö­ã¿é¤J"
   ClientHeight    =   6460
   ClientLeft      =   1690
   ClientTop       =   1860
   ClientWidth     =   9160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6460
   ScaleWidth      =   9160
   Begin VB.TextBox txtADate 
      Height          =   285
      Left            =   6000
      MaxLength       =   8
      TabIndex        =   95
      Top             =   2536
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm03020401_04.frx":0000
      Left            =   7350
      List            =   "frm03020401_04.frx":0016
      TabIndex        =   18
      Top             =   4860
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frm03020401_04.frx":0050
      Left            =   6120
      List            =   "frm03020401_04.frx":0052
      TabIndex        =   24
      Top             =   5745
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2430
      MaxLength       =   1
      TabIndex        =   23
      Top             =   5760
      Width           =   372
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5910
      MaxLength       =   8
      TabIndex        =   22
      Top             =   5460
      Width           =   2532
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1530
      MaxLength       =   1
      TabIndex        =   25
      Top             =   6060
      Width           =   492
   End
   Begin VB.TextBox textCP53 
      Height          =   285
      Left            =   6000
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2834
      Visible         =   0   'False
      Width           =   1308
   End
   Begin VB.TextBox textCP54 
      Height          =   285
      Left            =   7680
      MaxLength       =   8
      TabIndex        =   4
      Top             =   2834
      Visible         =   0   'False
      Width           =   1308
   End
   Begin VB.TextBox textMod 
      Height          =   285
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   21
      Top             =   5460
      Width           =   372
   End
   Begin VB.TextBox textPrtTrans 
      Height          =   285
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   17
      Top             =   4860
      Width           =   372
   End
   Begin VB.TextBox textDN 
      Height          =   285
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   19
      Top             =   5160
      Width           =   492
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "ÅÜ§ó¨Æ¶µ(R)"
      Height          =   400
      Left            =   4620
      TabIndex        =   27
      Top             =   15
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "¦^«eµe­±(&U)"
      Height          =   400
      Left            =   6840
      TabIndex        =   29
      Top             =   0
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5880
      TabIndex        =   28
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   30
      Top             =   0
      Width           =   912
   End
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   16
      Top             =   4860
      Width           =   732
   End
   Begin VB.TextBox textTMBM07_2 
      Height          =   285
      Left            =   7320
      MaxLength       =   2
      TabIndex        =   7
      Top             =   3132
      Width           =   732
   End
   Begin VB.TextBox textTMBM07_1 
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   6
      Top             =   3132
      Width           =   732
   End
   Begin VB.TextBox textTM14 
      Height          =   285
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3132
      Width           =   2532
   End
   Begin VB.TextBox textCP08 
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   2
      Top             =   2834
      Width           =   2532
   End
   Begin VB.TextBox textCP25 
      Enabled         =   0   'False
      Height          =   264
      Left            =   1788
      MaxLength       =   8
      TabIndex        =   1
      Top             =   96
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.TextBox textTM15 
      Height          =   285
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2536
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2238
      Width           =   2412
   End
   Begin VB.TextBox textCP45 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1940
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1940
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1642
      Width           =   3345
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1642
      Width           =   2532
   End
   Begin VB.TextBox textTM22S 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1344
      Width           =   1692
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1344
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   450
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   450
      Width           =   2532
   End
   Begin VB.TextBox textTM22 
      Height          =   285
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   9
      Top             =   3430
      Width           =   1092
   End
   Begin VB.TextBox textTM21 
      Height          =   285
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   8
      Top             =   3430
      Width           =   1092
   End
   Begin VB.TextBox textCP26 
      Height          =   285
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   10
      Top             =   3430
      Width           =   372
   End
   Begin VB.TextBox textCP14 
      Height          =   285
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   11
      Top             =   3728
      Width           =   732
   End
   Begin VB.TextBox textCP48 
      Height          =   285
      Left            =   5910
      MaxLength       =   8
      TabIndex        =   12
      Top             =   3728
      Width           =   2532
   End
   Begin VB.TextBox textTM17 
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   14
      Top             =   4026
      Width           =   372
   End
   Begin VB.TextBox textTM16S 
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   13
      Top             =   4026
      Width           =   405
   End
   Begin VB.Label lblADate 
      Caption         =   "­ì¨ç¤½§i¤é:"
      Height          =   252
      Left            =   4776
      TabIndex        =   94
      Top             =   2544
      Visible         =   0   'False
      Width           =   1068
   End
   Begin MSForms.TextBox textTM67 
      Height          =   285
      Left            =   5910
      TabIndex        =   26
      Top             =   6060
      Width           =   3195
      VariousPropertyBits=   -1476378597
      MaxLength       =   200
      Size            =   "5636;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPS 
      Height          =   525
      Left            =   1200
      TabIndex        =   15
      Top             =   4314
      Width           =   7815
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13785;926"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2370
      X2              =   2580
      Y1              =   3570
      Y2              =   3570
   End
   Begin MSForms.TextBox textCP35 
      Height          =   285
      Left            =   5910
      TabIndex        =   20
      Top             =   5160
      Width           =   2535
      VariousPropertyBits=   671105051
      MaxLength       =   32
      Size            =   "4471;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1200
      TabIndex        =   93
      Top             =   748
      Width           =   7875
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13891;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   1046
      Width           =   7755
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "13679;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5760
      TabIndex        =   91
      Top             =   2238
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_2 
      Height          =   285
      Left            =   1980
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   3728
      Width           =   1905
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3360;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.CheckBox Chk1 
      Height          =   255
      Left            =   2970
      TabIndex        =   89
      Top             =   5175
      Width           =   1500
      BackColor       =   -2147483633
      ForeColor       =   255
      DisplayStyle    =   4
      Size            =   "2646;450"
      Value           =   "0"
      Caption         =   "¼È¤£¦C¦L©w½Z"
      FontName        =   "·s²Ó©úÅé"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "©ñ±ó±M¥ÎÅv¡G"
      Height          =   180
      Index           =   0
      Left            =   4770
      TabIndex        =   88
      Top             =   6105
      Width           =   1080
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
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
      TabIndex        =   87
      Top             =   502
      Width           =   645
   End
   Begin VB.Label Label32 
      Caption         =   "©w½Z®×¥ó©Ê½è :"
      Height          =   255
      Left            =   5970
      TabIndex        =   86
      Top             =   4875
      Width           =   1455
   End
   Begin VB.Label Label31 
      Caption         =   "½Ð´Ú³æ¦Lªí¾÷ :"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4800
      TabIndex        =   85
      Top             =   5775
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label30 
      Caption         =   "(Y:¬O)"
      Height          =   255
      Left            =   2910
      TabIndex        =   84
      Top             =   5775
      Width           =   855
   End
   Begin VB.Label Label29 
      Caption         =   "¬O§_²£¥Íµù¥UÃÒ½Ð´Ú¸ê®Æ :"
      Height          =   255
      Left            =   120
      TabIndex        =   83
      Top             =   5775
      Width           =   2235
   End
   Begin VB.Label Label28 
      Caption         =   "ÃÒ®Ñ¤é´Á :"
      Height          =   255
      Left            =   4800
      TabIndex        =   82
      Top             =   5475
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "(Y:¤º³¡¦¬¤å§ó§ï)"
      Height          =   255
      Index           =   13
      Left            =   2130
      TabIndex        =   81
      Top             =   6075
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "¬O§_§ó§ï®Ö­ã¨ç : "
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   80
      Top             =   6075
      Width           =   1485
   End
   Begin VB.Label Label4 
      Caption         =   "½èÅv³]©w´Á¶¡ :"
      Height          =   255
      Index           =   0
      Left            =   4776
      TabIndex        =   79
      Top             =   2849
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "¡Ð"
      Height          =   180
      Index           =   1
      Left            =   7410
      TabIndex        =   78
      Top             =   2880
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label25 
      Caption         =   "¬O§_¬°ÃÒ®Ñ§ó§ï :"
      Height          =   255
      Left            =   120
      TabIndex        =   77
      Top             =   5475
      Width           =   1575
   End
   Begin VB.Label Label18 
      Caption         =   "(Y:¬O)"
      Height          =   255
      Left            =   2160
      TabIndex        =   76
      Top             =   5475
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "¼f¬d©e­û :"
      Height          =   255
      Left            =   4800
      TabIndex        =   75
      Top             =   5175
      Width           =   975
   End
   Begin VB.Label lbl4 
      Caption         =   "¬O§_¦C¦LÂ½Ä¶¨ç :"
      Height          =   255
      Left            =   2970
      TabIndex        =   74
      Top             =   4875
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "(N:¤£¦L)"
      Height          =   255
      Left            =   5010
      TabIndex        =   73
      Top             =   4875
      Width           =   855
   End
   Begin VB.Label Label36 
      Caption         =   "¬O§_¿é¤JD/N :"
      Height          =   255
      Left            =   120
      TabIndex        =   72
      Top             =   5175
      Width           =   1215
   End
   Begin VB.Label Label37 
      Caption         =   "(Y:¿é¤J)"
      Height          =   255
      Left            =   2040
      TabIndex        =   71
      Top             =   5175
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "(N:¤£¦L)"
      Height          =   255
      Left            =   2040
      TabIndex        =   70
      Top             =   4875
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "¦C¦L©w½Z :"
      Height          =   255
      Left            =   120
      TabIndex        =   69
      Top             =   4875
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "´Á"
      Height          =   255
      Left            =   8070
      TabIndex        =   68
      Top             =   3147
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "¨÷"
      Height          =   255
      Left            =   6840
      TabIndex        =   67
      Top             =   3147
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "¤½³ø¨÷´Á :"
      Height          =   255
      Left            =   4776
      TabIndex        =   66
      Top             =   3147
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "¤½§i¤é :"
      Height          =   180
      Left            =   120
      TabIndex        =   65
      Top             =   3184
      Width           =   990
   End
   Begin VB.Label Label8 
      Caption         =   "¾÷Ãö¤å¸¹ :"
      Height          =   255
      Left            =   120
      TabIndex        =   64
      Top             =   2874
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "®Ö­ã³qª¾¤é :"
      Enabled         =   0   'False
      Height          =   252
      Left            =   672
      TabIndex        =   63
      Top             =   108
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "¼f©w¸¹¼Æ :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   62
      Top             =   2571
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "´¼Åv¤H­û :"
      Height          =   255
      Index           =   11
      Left            =   4776
      TabIndex        =   61
      Top             =   2253
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "¨Ó¨ç¦¬¤å¤é :"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   60
      Top             =   2268
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "©¼©Ò®×¸¹ :"
      Height          =   255
      Index           =   9
      Left            =   4776
      TabIndex        =   59
      Top             =   1955
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "¥Ó½Ð°ê®a :"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   58
      Top             =   1965
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "°Ó«~Ãþ§O :"
      Height          =   255
      Index           =   7
      Left            =   4776
      TabIndex        =   57
      Top             =   1657
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "®×¥ó©Ê½è :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   56
      Top             =   1662
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "¥¿°Ó¼Ð±M¥Î´Á¤î¤é :"
      Height          =   255
      Index           =   5
      Left            =   4776
      TabIndex        =   55
      Top             =   1359
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "¦¬¤å¤é :"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   54
      Top             =   1359
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "°Ó¼ÐºØÃþ :"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   53
      Top             =   465
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "¥Ó½Ð¤H :"
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   1056
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "®×¥ó¦WºÙ :"
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   753
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "¥»©Ò®×¸¹ :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   50
      Top             =   450
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "±M¥Î´Á­­ :"
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   3445
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "(N:¤£ºâ)"
      Height          =   255
      Left            =   6840
      TabIndex        =   48
      Top             =   3445
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "¬O§_ºâ®×¥ó¼Æ :"
      Height          =   255
      Left            =   4776
      TabIndex        =   47
      Top             =   3445
      Width           =   1215
   End
   Begin VB.Label Label24 
      Caption         =   "©Ó¿ì¤H :"
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   3743
      Width           =   855
   End
   Begin VB.Label Label26 
      Caption         =   "©Ó¿ì´Á­­ :"
      Height          =   255
      Left            =   4776
      TabIndex        =   45
      Top             =   3743
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "¦C¦L³Æµù :"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   4314
      Width           =   975
   End
   Begin VB.Label Label20 
      Caption         =   "(Y / N)"
      Height          =   255
      Left            =   6840
      TabIndex        =   43
      Top             =   4041
      Width           =   615
   End
   Begin VB.Label Label19 
      Caption         =   "±M¥ÎÅv¬O§_¦s¦b :"
      Height          =   255
      Left            =   4776
      TabIndex        =   42
      Top             =   4041
      Width           =   1575
   End
   Begin VB.Label Label21 
      Caption         =   "(1:­ã , 2:»é)"
      Height          =   255
      Left            =   1950
      TabIndex        =   41
      Top             =   4041
      Width           =   1155
   End
   Begin VB.Label Label27 
      Caption         =   "®×¥ó¥Ø«e­ã»é :"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   4041
      Width           =   2295
   End
End
Attribute VB_Name = "frm03020401_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/13 §ï¦¨Form2.0 ; cmbTM05¡BtextTM23¡BtextCP13¡BtextCP14_2¡BtextCP35¡BtextPS¡BtextTM67(111/8/8)
'Memo By Sindy 2012/12/4 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/16 SQLDate¤wÀË¬d
'Memo By Sindy 2010/11/29 ­û¤u½s¸¹Äæ¤w­×§ï
'Memo By Sindy 2010/8/11 ¤é´ÁÄæ¤w­×§ï
'2005/7/19¾ã²z
Option Explicit

' ¥»©Ò®×¸¹
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' ¨Ó¨ç¦¬¤å¤é
Dim m_CP05 As String
' ­ìªk©w´Á­­
Dim m_CP07 As String
' ¦¬¤å¸¹
Dim m_CP09 As String
' ­ì®×¥ó©Ê½è
Dim m_CP10 As String
' ­ì·~°È°Ï
Dim m_CP12 As String
' ­ì´¼Åv¤H­û¥N¸¹
Dim m_CP13 As String
' ­ì©Ó¿ì¤H¥N¸¹
Dim m_CP14 As String
' ­ì±ÂÅv´Á¶¡(¨´)
Dim m_CP54 As String
' ­ì²¾Âà¥Ó½Ð¤H¥N¸¹
Dim m_CP56 As String
'Add By Sindy 2013/1/11
Dim m_CP89 As String
Dim m_CP90 As String
Dim m_CP91 As String
Dim m_CP92 As String
'2013/1/11 End
' °Ó¼ÐºØÃþ¥N½X
Dim m_TM08 As String
' °ê®a¥N½X
Dim m_TM10 As String
' ­ì±M¥Î´Á­­°_¤é
Dim m_TM21 As String
' ­ì±M¥Î´Á­­¤î¤é
Dim m_TM22 As String
' ­ì¥Ó½Ð¤H¥N¸¹
Dim m_TM23 As String
'Add By Sindy 2013/1/11
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
'2013/1/11 End
' ¥Ó½Ð°ê®aªº©µ®i¦~«×
Dim m_NA14 As Integer
' ³Q±ÂÅv¤H
Dim m_CP50 As String
' ²¾Âà¤H
Dim m_CP55 As String
' ¥¿°Ó¼Ð¸¹¼Æ
Dim m_TM27 As String
'Add By Cheng 2002/01/15
Dim m_intNumBegin As Integer
Dim m_intNumEnd As Integer
Dim m_strNumBegin As String
Dim m_strNumEnd As String
'Add By Cheng 2002/02/01
Dim m_strLastTextTM14 As String
Dim m_strLastTextTMBM07_1 As String
Dim m_strLastTextTMBM07_2 As String
Dim m_strLastTextTM16S As String
Dim m_strLastTextTM17 As String
'Add By Cheng 2002/12/11
'Dim m_blnClkChgButton As Boolean '¬O§_¦³«öÅÜ§ó¨Æ¶µ¶s
Public m_blnClkChgButton As Boolean '¬O§_¦³«öÅÜ§ó¨Æ¶µ¶s 'Modify By Sindy 2012/2/6 Dim->Public
'Add By Cheng 2003/03/11
Dim m_TM67 As String '©ñ±ó±M¥ÎÅv
'Add By Cheng 2003/07/14
Dim m_CP64 As String '¶i«×³Æµù
'Add By Cheng 2003/09/05
Dim m_strCP09 As String 'For ©w½Z
'Add By Cheng 2003/09/05
Dim m_blnPrintAddress As Boolean '¬O§_­n¦C¦L¦a±ø
Dim m_strSerialNo As String '½Ð´Ú³æ¸¹
Dim strPrint As String '°O¿ý¹w³]¦Lªí¾÷¦WºÙ
Dim prnPrint As Printer
'Add By Cheng 2003/11/19
Dim m_TM11 As String '¥Ó½Ð¤é
Dim m_blnPriDate As Boolean '§PÂ_¬O§_¦³Àu¥ýÅv
'Add By Cheng 2003/12/22
Dim m_strWithRegister As String '¬O§_ªþµù¥UÃÒ(©w½Z§PÂ_¨Ï¥Î, "Y" : ªþµù¥UÃÒ, ¨ä¥L : ¤£ªþµù¥UÃÒ)
'Add By Cheng 2004/01/16
Dim m_blnNewTrans As Boolean '¬O§_¥X·sÄ¶¤å
Dim m_TM14 As String '¤½§i¤é
Dim m_TM58 As String '®×¥ó³Æµù
'Add By Cheng 2004/04/13
Dim m_blnRestrictGoods As Boolean
'End
'Add by Morgan 2004/5/27
Dim m_CP27 As String 'µo¤å¤é
'add by nick 2004/08/20
Dim m_NickCp09 As String    '¤é¤å©w½Z¥Î­n§ì¦¬¤å¤é¬°¨Ó¨ç¤é
'add by nick 2004/10/28
Dim m_CP06 As String
'ADD BY NICK 2005/06/28
Dim Is716Have As Boolean
Dim StrSQLa As String
Dim m_TM122 As String       '2008/7/24 ADD BY SONIA FCTµù¥U¶O¦Û°Ê¥NÃº
Dim arrCP10
Dim strCP10Code As String
'Dim bChkChaEvent As Boolean 'Add By Sindy 2010/5/13
Dim m_TM118 As String 'Add By Sindy 2010/11/17
Dim bolChaEventNewCase As Boolean 'Add By Sindy 2012/2/1
Dim m_TM20 As String 'Add By Sindy 2012/8/7
Dim m_CP148 As String '¬O§_¬°¤@¥Ó½Ð®Ñ¦h¥ó Add By Sindy 2012/10/12
Dim ET01 As String, ET02 As String, ET03 As String, ET03_1 As String, ET03r As String
Dim m_CP28 As String 'Modify By Sindy 2012/11/08
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
Dim m_fa76 As String 'Add By Sindy 2013/12/20
Dim bolMod As Boolean 'Added by Lydia 2016/07/19 ¬O§_¦³ÅÜ§ó¨Æ¶µ
'Add By Sindy 2016/12/6
Dim m_strCE04 As String
Dim m_strCE23CE24CE25 As String
'2016/12/6 END
'Added by Morgan 2017/5/3 ¹q¤l¤½¤å
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/5/3
Dim m_NA86 As String 'Added by Sindy 2020/5/19 ¬O§_°±¤î¶l°È
Dim m_TM136 As String 'Added by Lydia 2023/02/24 µù¥UÃÒ§Î¦¡
Dim bolToFile As Boolean 'Added by Lydia 2023/06/05 (±qPrintLetter²¾¹L¨Ó)±N©w½Z¡BÂ½Ä¶¨ç©MÃÒ®Ñ¦s¤JFCT_WorkFlow; »P¿é¤Jµù¥UÃÒfrm03020404_03ªº³B²z¬Û¦P¡A­Y³W«h¦³ÅÜ§ó¡A½Ð¤@¨Ö­×§ï
Dim strFN03 As String  'Added by Lydia 2023/06/05 (±qPrintLetter²¾¹L¨Ó)ÃÒ®ÑÀÉ¦W
'Added by Lydia 2023/09/04 ­×§ï¤é¤å²Õ¤§®Ö­ã-§ó¥¿(©µ®i®Ö­ã¨ç)¤§©w½Z¤ÎÄ¶¤å:¦P®É²£¥Í¡u©µ®i¡v©w½Z+Ä¶¤å¡B¡u§ó¥¿¡vÄ¶¤å¡B¤U¸ü¡u©µ®i¡B§ó¥¿¡v©x¤è¨Ó¨ç
Dim strFilePath As String, strFN01 As String, strFN02 As String '(±qPrintLetter²¾¹L¨Ó)¦sÀÉ¸ô®|©M©w½ZÀÉ¦W
Dim m_CP43 As String, m_CP43pty As String '¬ÛÃö¦¬¤å¸¹©M®×¥ó©Ê½è
Dim strFN04 As String, strFN05 As String, ET03_ex As String '¥t¥~²£¥Í©w½Z
'end 2023/09/04

' ­ì¸ê®Æ¬O§_¦³¹ê»Úµ²ªG
Private Sub cmdCancel_Click()
   Unload Me
   frm03020401_03.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm03020401_03
   Unload frm03020401_02
   Unload frm03020401_01
   Unload Me
End Sub

' ´£¨Ñ¥~³¡µ{¦¡©I¥s¥Î¨Óµ²§ô¦¹¶µ§@·~
Public Sub OnAppExit()
   cmdExit_Click
End Sub

Private Sub cmdMod_Click()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
    'Add By Cheng 2002/12/11
    'Modify By Sindy 2012/2/6 Mark
'    m_blnClkChgButton = True
   
   bolMod = False 'Added by Lydia 2016/07/19
   
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   'edit by nickc 2005/08/04
   'rsTmp.Open StrSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then
      rsTmp.Close
      strMsg = "µLÅÜ§ó¨Æ¶µ°O¿ý"
      strTit = "¸ê®ÆÀË®Ö"
      'Modified by Lydia 2016/07/19 +§PÂ_
      If cmdMod.Visible = True Then
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      
      GoTo EXITSUB
   End If
   
   bolMod = True 'Added by Lydia 2016/07/19
   rsTmp.Close
   DisplayNextForm
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub cmdok_Click()
Dim strFilePath As String 'Added by Lydia 2020/03/09 ±½ºËÀÉªº¸ô®|
Dim rsA As New ADODB.Recordset

   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
      If TxtValidate = False Then Exit Sub
      'Added by Lydia 2020/03/09 §ó¥¿®Ö­ã(µù¥UÃÒ)­Y¯ÊÀÉ«h´£¿ô¤£¥i¿é¤J¡A¤£¯Ê«h¦Û°ÊÂk¤J¨÷©v°Ï¡C
      If frm03020401_03.GetSelectResult() = "1" Then
        If m_DocNo = "" Then 'Added by Lydia 2022/02/10  FCT¯È¥»¤½¤å¨Ó¨ç¡A¦P®É±N¤½¤å¨çFCT_OA_SCAN¶×¤J¨÷©v°Ï
            If PUB_FCTCheckPDF(m_TM01, m_TM02, m_TM03, m_TM04, "1001", m_CP09, strFilePath) = False Then
                 Exit Sub
            End If
        End If
      End If
      'end 2020/03/09
        
      ' ³]©w·Æ¹«´å¼Ð¬°µ¥«Ýª¬ºA
      Screen.MousePointer = vbHourglass
      ' Àx¦s¸ê®Æ
      'edit by  nick 2004/11/03
      'OnSaveData
     
      'Added by Lydia 2016/07/19 ©µ®i®Ö­ã¦b¦sÀÉ®É,ª½±µ±NÅÜ§ó¨Æ¶µ½T©w¥þ³¡®Ö­ã
      'Modified by Lydia 2017/07/28 +301ÅÜ§ó®Ö­ã,¤ñ·Ó©µ®i®Ö­ã¿ì²z
       If m_CP10 = "102" Or m_CP10 = "301" Then
          Call cmdMod_Click
          If bolMod Then '¦³ÅÜ§ó¨Æ¶µ
             If frm03020401_05.Get102_Approve = False Then
                 Screen.MousePointer = vbDefault: Exit Sub
             End If
          End If
       End If
       'end 2016/07/19
       
      If OnSaveData = False Then MsgBox "¦sÀÉ¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' ³]©w·Æ¹«´å¼Ð¬°¹w³]
      Screen.MousePointer = vbDefault
      
      'Add By Sindy 2023/3/9 T091286 ²¾Âà(501)©ÎÅÜ§ó(301)¥Ó½Ð¤H¦Û½ÐºM¦^¶·ÁÙ­ì¥Ó½Ð¤H
      '©ó¦Û½ÐºM¦^®Ö­ã¿é¤J®É¼u´£¿ô­×§ï¥Ó½Ð¤H¸ê®Æ
      If m_CP10 = "306" Then
         'Modified by Lydia 2023/09/04 §ï¥ÎÅÜ¼Æ
         'strSql = "Select CP09,CP10 From CaseProgress Where CP09=(Select CP43 From CaseProgress Where CP09='" & m_CP09 & "')"
         'rsA.CursorLocation = adUseClient
         'rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         'If rsA.RecordCount > 0 Then
         '   If rsA.Fields("CP10") = "501" Or rsA.Fields("CP10") = "301" Then
         '      strExc(10) = rsA.Fields("CP09")
         '      If rsA.Fields("CP10") = "301" Then
            If m_CP43pty = "501" Or m_CP43pty = "301" Then
         'end 2023/09/04
               strExc(10) = m_CP43
               If m_CP43pty = "301" Then
                  '¦³ÅÜ§ó¥Ó½Ð¤H
                  strSql = "Select CE01 From ChangeEvent Where CE01='" & strExc(10) & "'" & _
                           " AND CE04||CE05||CE06||CE07||CE08 is not null"
                  rsA.CursorLocation = adUseClient
                  rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount = 0 Then
                     strExc(10) = ""
                  End If
               End If
               If strExc(10) <> "" Then
                  MsgBox "½Ðª`·N¡A­×§ï¥Ó½Ð¤H¸ê®Æ¡I", vbCritical, "¦Û½ÐºM¦^®Ö­ã"
               End If
            End If
         'End If 'Mark by Lydia 2023/09/04
      End If
      '2023/3/9 END
      
      Unload frm03020401_03
      Unload frm03020401_02
      'Ken 91.04.09 -- Start
      If textDN = "Y" Then
        'Add By Cheng 2003/03/19
        '·s¼W¦a§}±ø¦Cªí¸ê®Æ
'edit by nick 2004/10/05 d/n ¤£¦L
'        pub_AddressListSN = pub_AddressListSN + 1
'        PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
         Screen.MousePointer = vbHourglass
         Frmacc21h0.Show
         mdiMain.ToolShow
         mdiMain.tool1_enabled
         Screen.MousePointer = vbDefault
         
         If m_DocNo = "" Then 'Added by Morgan 2017/5/3  ¹q¤l¤½¤å
            Set Frmacc21h0.frmlink = frm03020401_01
         End If 'Added by Morgan 2017/5/3  ¹q¤l¤½¤å
         
         'add by nick 2004/11/24
         Frmacc21h0.IsPrintAddress = False
      Else
         frm03020401_01.Show
      End If
        'Add By Cheng 2003/09/05
        '·s¼W¦a§}±ø¦Cªí¸ê®Æ
        'edit by nick 2004/11/17
        'If Me.textDN.Text = "" And m_blnPrintAddress = True Then
        'edit by nickc 2007/04/02 ªü½¬»¡©µ®i¤@©w½Ð´Ú¡A©Ò¥H¤£¥Î¦A¥t¥~¥X¦a§}±ø¡A¦³Ápµ¸³æ
        If m_blnPrintAddress = True Then
        '2010/6/11 modify by sonia ©µ®i¨ú®ø½Ð´Ú¬G­n¦L¦a§}±ø
        'If m_blnPrintAddress = True And m_CP10 <> "102" Then
            'Modify By Sindy 2025/10/2 ¨ú®ø¦a§}±ø
'            pub_AddressListSN = pub_AddressListSN + 1
'            'Modify By Sindy 2016/7/1 + , m_CP10
'            PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0", m_CP10
        End If
        
       'Added by Lydia 2020/03/09 FCT®×¿é¤Jµù¥UÃÒ©Î§ó¥¿®Ö­ã(µù¥UÃÒ)«e¡A¥ý±½ºËµù¥UÃÒ¦Ü©T©w¸ê®Æ§¨¡A¿éµù¥UÃÒ­Y¯ÊÀÉ«h´£¿ô¤£¥i¿é¤J¡A¤£¯Ê«h¦Û°ÊÂk¤Jµù¥UÃÒ¨º¹D¤§¨÷©v°Ï¡C
       If strFilePath <> "" Then
           If Pub_AutoSavePdf2_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_NickCp09, "1001", strFilePath) = False Then
               Exit Sub
           End If
       End If
       'end 2020/03/09
       
       'Added by Lydia 2023/06/05 ¤U¸ü¨÷©v°ÏªºÃÒ®ÑPDF
       If bolToFile = True Or strFilePath <> "" Then
          '«O¯d´ú¸Õ¥Î¡GFCT-46767
          'strSql = "select cpp14 From casepaperpdf where cpp01='CB2012458' " & _
                    "and instr(upper(cpp02),upper('." & IIf(m_TM136 = "1", "CERT", "1001") & ".PDF'))>0"
          If InStr("103,302", m_CP10) > 0 Then  'Added by Lydia 2023/05/03 ¦b¿é¤J¡u®Ö­ã-¸É´«µoÃÒ®Ñ103¡v¡B¡u®Ö­ã-§ó¥¿302¡v¡A¤ñ·Ó¡uµù¥UÃÒ¿é¤J1701¡vªº³W«h
             'Modified by Morgan 2025/3/28 +cpp19
             strSql = "select cpp14,cpp19 From casepaperpdf where cpp01='" & m_NickCp09 & "' and instr(upper(cpp02),upper('." & IIf(m_TM136 = "1", "CERT", "1001") & ".PDF'))>0"
          'Added by Lydia 2023/05/03 ¨ä¥L®Ö­ã
          Else
             'Modified by Morgan 2025/3/28 +cpp19
             strSql = "select cpp14,cpp19 From casepaperpdf where cpp01='" & m_NickCp09 & "' and instr(upper(cpp02),upper('." & "1001" & ".PDF'))>0"
          End If
          'end 2023/05/03
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strSql)
          If intI = 1 Then
             'Modified by Lydia 2023/09/04 ­×§ï¤é¤å²Õ¤§®Ö­ã-§ó¥¿(©µ®i®Ö­ã¨ç)¤§©w½Z¤ÎÄ¶¤å:¦P®É²£¥Í¡u©µ®i¡v©w½Z+Ä¶¤å¡B¡u§ó¥¿¡vÄ¶¤å¡B¤U¸ü¡u©µ®i¡B§ó¥¿¡v©x¤è¨Ó¨ç
             'If PUB_GetFtpFile("" & RsTemp.Fields("cpp14"), Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04) & "\" & strFN03, "Casepaperpdf") = True Then
             If ET03_ex <> "" And strFN05 <> "" Then
                strExc(1) = strFN05
             Else
                strExc(1) = strFN03
             End If
             'Modified by Morgan 2025/3/28 +cpp19
             If PUB_GetFtpFile("" & RsTemp.Fields("cpp14"), Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04) & "\" & strExc(1), "Casepaperpdf", , , "" & RsTemp.Fields("cpp19") <> "") = True Then
             'end 2023/09/04
             End If
          End If
          'Added by Lydia 2023/09/04 ­×§ï¤é¤å²Õ¤§®Ö­ã-§ó¥¿(©µ®i®Ö­ã¨ç)¤§©w½Z¤ÎÄ¶¤å:¦P®É²£¥Í¡u©µ®i¡v©w½Z+Ä¶¤å¡B¡u§ó¥¿¡vÄ¶¤å¡B¤U¸ü¡u©µ®i¡B§ó¥¿¡v©x¤è¨Ó¨ç
          If ET03_ex <> "" And strFN04 <> "" Then
             'Modified by Morgan 2025/3/28 +cpp19
             strSql = "select cpp14,cpp19 From casepaperpdf where cpp01='" & m_CP43 & "' and instr(upper(cpp02),upper('." & "1001" & ".PDF'))>0"
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strSql)
             If intI = 1 Then
                If PUB_GetFtpFile("" & RsTemp.Fields("cpp14"), Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04) & "\" & strFN04, "Casepaperpdf", , , "" & RsTemp.Fields("cpp19") <> "") = True Then
                End If
             End If
          End If
          'end 2023/09/04
       End If
       'end 2023/06/05
       
      'Ken 91.04.09 -- End
'      frm03020401_01.Show
      
      'Modified by Morgan 2017/5/3 ¹q¤l¤½¤å
      'Unload Me
      If m_DocNo <> "" Then
         cmdExit_Click
         frm02010412.GoNext
      Else
         Unload Me
      End If
      'end 2017/5/3
   End If
End Sub

'Added by Morgan 2022/1/11
Private Sub Form_Activate()
   Static bDone As Boolean
   
   If bDone = False Then
      '¹q¤l¤½¤å´å¼Ð¹w³]¦b¤½§i¤é--³¯ª÷½¬
      If m_DocWord <> "" And textTM14.Enabled Then
         textTM14.SetFocus
      End If
      bDone = True
   End If
End Sub

Private Sub Form_Load()
  
   ' ³]©w±±¨î¶µªº­I´ºÃC¦â
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   'textTM27.BackColor = &H8000000F     '2009/4/27 cancel by sonia
   textTM22S.BackColor = &H8000000F
   
   textCP05.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   'textCP09.BackColor = &H8000000F     '2009/4/27 cancel by sonia
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP45.BackColor = &H8000000F
   
   ' 90.08.29 modify (¼f©w¸¹Äæ¦ì§ï¬°¥uÅã¥Ü¤£¥i­×§ï)
   EnableTextBox textTM15, False
  
   MoveFormToCenter Me
   
   'Add By Cheng 2002/02/01
   '«O¯d¤W¤@¦¸¿é¤Jªº¸ê®Æ
   '2009/4/27 modify by soniaªü½¬»¡§ï¦^¥Á°ê¦~
   'edit by nickc 2008/03/12 §ï¦è¤¸¦~
   Me.textTM14.Text = "" & m_strLastTextTM14
   'Me.textTM14.Text = DBDATE("" & m_strLastTextTM14)
   Me.textTMBM07_1.Text = "" & m_strLastTextTMBM07_1
   Me.textTMBM07_2.Text = "" & m_strLastTextTMBM07_2
   'Modify By Cheng 2002/07/22
'   Me.textTM16S.Text = "" & m_strLastTextTM16S
'   Me.textTM17.Text = "" & m_strLastTextTM17
    'Add By Cheng 2002/12/11
'    m_blnClkChgButton = False
    
    PUB_SetPrinter Me.Name, Combo2, strPrint    'Modified by Morgan 2017/11/21 ³]©w¦Lªí¾÷§ï©I¥s¤½¥Î¨ç¼Æ,­ìµ{¦¡²¾°£
    
    'Add By Cheng 2003/02/27
    '¹w³]¤£¦C¦L¦a§}±ø
    m_blnPrintAddress = False
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   ' ²M°£·j´MªºKey
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' ¥»©Ò®×¸¹ Äæ¦ì1
      Case 0: m_TM01 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì2
      Case 1: m_TM02 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì3
      Case 2: m_TM03 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì4
      Case 3: m_TM04 = strData
      ' ¨Ó¨ç¦¬¤å¤é
      Case 4: m_CP05 = strData
      ' ¦¬¤å¸¹
      Case 5: m_CP09 = strData
             'add by nickc 2005/08/04
            strSql = "SELECT * FROM ChangeEvent " & _
                     "WHERE CE01 = '" & m_CP09 & "' "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <= 0 Then
               m_blnClkChgButton = True
            Else
               m_blnClkChgButton = False
            End If
            rsTmp.Close
   End Select
End Sub

' Åª¨ú°Ó¼Ð°ò¥»ÀÉ
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim strSub As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSub As ADODB.Recordset
   
   ' ¨ú±o°Ó¼Ð°ò¥»ÀÉªº¬ÛÃö¶µ¥Ø
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' ¥Ó½Ð°ê®a
      'Add By Cheng 2002/07/19
      m_TM10 = Empty
      m_NA14 = Empty
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         ' ¨ú±o°ê®aªº¦WºÙ
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
         ' ¨ú±o°ê®aªº©µ®i¦~«×
         m_NA14 = GetNationExtentYear(rsTmp.Fields("TM10"))
      End If
      
        'Add By Cheng 2003/11/19
        '¥Ó½Ð¤é
        m_TM11 = "" & rsTmp.Fields("TM11").Value
        '¤½§i¤é
        m_TM14 = "" & rsTmp.Fields("TM14").Value
        'add by nickc 2006/12/14
        '2009/4/27 modify by soniaªü½¬»¡§ï¦^¥Á°ê¦~
        'textTM14.Text = m_TM14
        'textTM14.Text = ChangeWStringToTString(m_TM14)
        If (m_CP10 = "101" Or m_CP10 = "308") And m_TM14 <> "" Then
           textTM14.Text = ChangeWStringToTString(m_TM14)
        End If
        'end 2015/5/26
        '®×¥ó³Æµù
        m_TM58 = "" & rsTmp.Fields("TM58").Value
        'End
      ' ¼f©w¸¹¼Æ
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' °Ó¼Ð¦WºÙ(¤¤)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' °Ó¼Ð¦WºÙ(­^)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' °Ó¼Ð¦WºÙ(¤é)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' Åã¥Ü°Ó¼Ð¦WºÙ
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' °Ó¼ÐºØÃþ
      'Add By Cheng 2002/07/19
      m_TM08 = Empty
      If IsNull(rsTmp.Fields("TM08")) = False Then
         m_TM08 = rsTmp.Fields("TM08")
         If m_TM10 < "010" Then
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
         Else
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
         End If
      End If
      ' °Ó«~Ãþ§O
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      
      'Add By Sindy 2012/8/7
      m_TM20 = Empty
      If IsNull(rsTmp.Fields("TM20")) = False Then
         m_TM20 = rsTmp.Fields("TM20")
      End If
      '2012/8/7 End
      
      ' ¥Ó½Ð¤H
      'Add By Cheng 2002/07/19
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      'Add By Sindy 2013/1/11
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = rsTmp.Fields("TM78")
      End If
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = rsTmp.Fields("TM79")
      End If
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = rsTmp.Fields("TM80")
      End If
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = rsTmp.Fields("TM81")
      End If
      '2013/1/11 End
  
      ' ¥¿°Ó¼Ð¸¹¼Æ
      'Add By Cheng 2002/07/19
      m_TM27 = Empty
      If IsNull(rsTmp.Fields("TM27")) = False Then
         m_TM27 = rsTmp.Fields("TM27")
         'textTM27 = rsTmp.Fields("TM27")    '2009/4/27 cancel by sonia
      End If
      'add by nickc 2006/05/29 ¥[¤J³¬¨÷´£¥Ü
      If IsNull(rsTmp.Fields("tm29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "¤w³¬¨÷"
      End If
      m_TM136 = "" & rsTmp.Fields("TM136") 'Added by Lydia 2023/02/24 µù¥UÃÒ§Î¦¡
      
      ' ¥¿°Ó¼Ð±M¥Î´Á¤î¤é
      Set rsSub = New ADODB.Recordset
      strSub = "SELECT * FROM TradeMark " & _
               "WHERE TM15 = '" & m_TM27 & "' AND " & _
                     "TM10 = '" & m_TM10 & "' "
      rsSub.CursorLocation = adUseClient
      rsSub.Open strSub, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
      If rsSub.RecordCount > 0 Then
         rsSub.MoveFirst
         If IsNull(rsSub.Fields("TM22")) = False Then
            'edit by nickc 2008/01/10 §ï¦¨¦è¤¸¦~
            'textTM22S = TAIWANDATE(rsSub.Fields("TM22"))
            textTM22S = DBDATE(rsSub.Fields("TM22"))
         End If
      End If
      rsSub.Close
      Set rsSub = Nothing
      'Modify By Cheng 2002/04/29
'      ' ¤½§i¤é
'      If IsNull(rsTmp.Fields("TM14")) = False Then
'         textTM14 = TAIWANDATE(rsTmp.Fields("TM14"))
'      End If
            
      'Add By Cheng 2002/07/22
      Me.textTM16S.Text = "" & rsTmp.Fields("TM16").Value
      
      'Modify By Cheng 2002/07/22
      'Åã¥Ü±M¥ÎÅv¬O§_¦s¦b
'      'Modify By Cheng 2002/07/11
'      '¤£­n±a¥X¸ê®Æ
'      ' ±M¥ÎÅv¬O§_¦s¦b
      If IsNull(rsTmp.Fields("TM17")) = False Then
         textTM17 = rsTmp.Fields("TM17")
      End If
      ' ±M¥Î´Á­­ (°_)
      'Add By Cheng 2002/07/19
      m_TM21 = Empty
      If IsNull(rsTmp.Fields("TM21")) = False Then
         m_TM21 = rsTmp.Fields("TM21")
         'edit by nickc 2008/01/10 §ï¦¨¦è¤¸¦~
         'textTM21 = TAIWANDATE(rsTmp.Fields("TM21"))
         textTM21 = DBDATE(rsTmp.Fields("TM21"))
      End If
      ' ±M¥Î´Á­­ (¤î)
      'Add By Cheng 2002/07/19
      m_TM22 = Empty
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_TM22 = rsTmp.Fields("TM22")
         'edit by  nickc 2008/01/10 §ï¦¨¦è¤¸¦~
         'textTM22 = TAIWANDATE(rsTmp.Fields("TM22"))
         textTM22 = DBDATE(rsTmp.Fields("TM22"))
      End If
        'Add By Cheng 2003/03/11
        '©ñ±ó±M¥ÎÅv
        m_TM67 = "" & rsTmp("TM67").Value
      'Add By Sindy 2010/01/05
      If IsNull(rsTmp.Fields("TM67")) = False Then
         textTM67 = "" & rsTmp("TM67").Value
      End If
      '2010/01/05 End
      
      'Add By Sindy 2010/11/17
      '¦P·N®Ñ°Ó¼Ð¸¹¼Æ
      m_TM118 = "" & rsTmp("TM118").Value
      '2010/11/17 End
      
      'Add By Sindy 2013/12/20
      m_fa76 = ""
      If IsNull(rsTmp.Fields("TM44")) = False Then
         Set rsSub = New ADODB.Recordset
         strSub = "SELECT fa76 FROM FAGENT " & _
                  "WHERE FA01 = '" & Mid(rsTmp.Fields("TM44"), 1, 8) & "' AND " & _
                        "FA02 = '" & Mid(rsTmp.Fields("TM44"), 9, 1) & "' "
         rsSub.CursorLocation = adUseClient
         rsSub.Open strSub, cnnConnection, adOpenStatic, adLockReadOnly
         If rsSub.RecordCount > 0 Then
            rsSub.MoveFirst
            If IsNull(rsSub.Fields("fa76")) = False Then
               m_fa76 = rsSub.Fields("fa76")
            End If
         End If
         rsSub.Close
         Set rsSub = Nothing
      End If
      '2013/12/20 END
      
      '2008/7/24 ADD BY SONIA ¨ÌTRADEMARK->FAGENT->CUSTOMER¶¶§Ç§ìFCTµù¥U¶O¦Û°Ê¥NÃº
      m_TM122 = ""
      'TRADEMARK
      If IsNull(rsTmp.Fields("TM122")) = False Then
         m_TM122 = rsTmp.Fields("TM122")
      Else
         'FAGENT
         If IsNull(rsTmp.Fields("TM44")) = False Then
            Set rsSub = New ADODB.Recordset
            strSub = "SELECT FA93 FROM FAGENT " & _
                     "WHERE FA01 = '" & Mid(rsTmp.Fields("TM44"), 1, 8) & "' AND " & _
                           "FA02 = '" & Mid(rsTmp.Fields("TM44"), 9, 1) & "' "
            rsSub.CursorLocation = adUseClient
            rsSub.Open strSub, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
            If rsSub.RecordCount > 0 Then
               rsSub.MoveFirst
               If IsNull(rsSub.Fields("FA93")) = False Then
                  m_TM122 = rsSub.Fields("FA93")
               End If
            End If
            rsSub.Close
            Set rsSub = Nothing
         End If
         'CUSTOMER
         If m_TM122 = "" Then
            Set rsSub = New ADODB.Recordset
            strSub = "SELECT * FROM CUSTOMER " & _
                     "WHERE CU01 = '" & Mid(rsTmp.Fields("TM23"), 1, 8) & "' AND " & _
                           "CU02 = '" & Mid(rsTmp.Fields("TM23"), 9, 1) & "' "
            rsSub.CursorLocation = adUseClient
            rsSub.Open strSub, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
            If rsSub.RecordCount > 0 Then
               rsSub.MoveFirst
               If IsNull(rsSub.Fields("CU128")) = False Then
                  m_TM122 = rsSub.Fields("CU128")
               End If
            End If
            rsSub.Close
            Set rsSub = Nothing
         End If
      End If
      '2008/7/24 END
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' Åª¨ú®×¥ó¶i«×ÀÉ
Private Sub QueryCaseProgress()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
'add by sonia 2019/4/30
Dim strSub As String
Dim rsSub As ADODB.Recordset
'end 2019/4/30
   
   ' ¨ú±o®×¥ó¶i«×ÀÉÀÉ®×¤¤Äæ¦ì
   'Modified by Lydia 2023/09/04 §ì¬ÛÃö¦¬¤å¸¹
   'strSql = "SELECT * FROM CaseProgress WHERE CP09 = '" & m_CP09 & "' "
   strSql = "SELECT C1.*,C2.CP10 as CP43pty FROM CaseProgress C1, CaseProgress C2 WHERE C1.CP09 = '" & m_CP09 & "' AND C1.CP43=C2.CP09(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'Add by Morgan 2004/5/27
      'µo¤å¤é
      m_CP27 = "" & rsTmp.Fields("CP27")
      
      'Add By Sindy 2012/11/08
      'µo¤å¦r¸¹
      m_CP28 = "" & rsTmp.Fields("CP28")
      
      ' ¦¬¤å¤é
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' ¾÷Ãö¤å¸¹
      If IsNull(rsTmp.Fields("CP08")) = False Then
         'Modify By Sindy 2012/5/31 Mark
         'textCP08 = rsTmp.Fields("CP08")
      End If
      '2009/4/27 cancel by sonia
      '' ¦¬¤å¸¹
      'If IsNull(rsTmp.Fields("CP09")) = False Then
      '   textCP09 = rsTmp.Fields("CP09")
      'End If
      '2009/4/27 end
      ' ®×¥ó©Ê½è
      'Add By Cheng 2002/07/19
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
        'Modify By Cheng 2003/01/16
        '¥[®×¥ó©Ê½è¦A±ÂÅv
'      'Add By Cheng 2002/06/13
'      '­Y®×¥ó©Ê½è¬°±ÂÅv
'      If m_CP10 = "502" Then
      If m_CP10 = "502" Or m_CP10 = "504" Then
         Me.Label4(0).Visible = True
         Me.Label4(1).Visible = True
            'Modify By Cheng 2003/01/16
            '¨Ì®×¥ó©Ê½è¼ÐÃD¤£¦P
'         Me.Label4(0).Caption = "±ÂÅv´Á¶¡¡G"
         Me.Label4(0).Caption = IIf(m_CP10 = "502", "±ÂÅv´Á¶¡¡G", "¦A±ÂÅv´Á¶¡¡G")
         Me.textCP53.Visible = True
         Me.textCP54.Visible = True
         '2009/4/27 modify by soniaªü½¬»¡§ï¦^¥Á°ê¦~
         'edit by nickc 2008/03/13
         Me.textCP53.MaxLength = 7
         Me.textCP54.MaxLength = 7
         Me.textCP53.Text = "" & TAIWANDATE("" & rsTmp.Fields("CP53"))
         Me.textCP54.Text = "" & TAIWANDATE("" & rsTmp.Fields("CP54"))
         'Me.textCP53.MaxLength = 8
         'Me.textCP54.MaxLength = 8
         'Me.textCP53.Text = "" & ("" & rsTmp.Fields("CP53"))
         'Me.textCP54.Text = "" & ("" & rsTmp.Fields("CP54"))
         '2009/4/27 end
      '­Y®×¥ó©Ê½è¬°³]©w½èÅv®É
      ElseIf m_CP10 = "506" Then
         Me.Label4(0).Visible = True
         Me.Label4(1).Visible = True
         Me.Label4(0).Caption = "½èÅv³]©w´Á¶¡¡G"
         Me.textCP53.Visible = True
         Me.textCP54.Visible = True
         '2009/4/27 modify by soniaªü½¬»¡§ï¦^¥Á°ê¦~
         'edit by nickc 2008/03/13
         Me.textCP53.MaxLength = 7
         Me.textCP54.MaxLength = 7
         Me.textCP53.Text = "" & TAIWANDATE("" & rsTmp.Fields("CP53"))
         Me.textCP54.Text = "" & TAIWANDATE("" & rsTmp.Fields("CP54"))
         'Me.textCP53.MaxLength = 8
         'Me.textCP54.MaxLength = 8
         'Me.textCP53.Text = "" & ("" & rsTmp.Fields("CP53"))
         'Me.textCP54.Text = "" & ("" & rsTmp.Fields("CP54"))
         '2009/4/27 end
      End If
      ' ·~°È°Ï
      m_CP12 = Empty
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' ´¼Åv¤H­û
      'Add By Cheng 2002/07/19
      m_CP13 = Empty
      'Modified by Lydia 2021/08/03 §ï¥ÑPUB_GetFCTSalesNo±a¥X©M²£¥ÍªºCÃþ¦¬¤å¤@­P
      'If IsNull(rsTmp.Fields("CP13")) = False Then
      '   m_CP13 = rsTmp.Fields("CP13")
      '   textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      'End If
      m_CP13 = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      textCP13 = GetStaffName(m_CP13)
      'end 2021/08/03
      
      '©Ó¿ì¤H
      m_CP14 = "" & rsTmp.Fields("CP14").Value
      'Added by Lydia 2023/09/04 ¬ÛÃö¦¬¤å¸¹©M®×¥ó©Ê½è
      m_CP43 = "" & rsTmp.Fields("CP43").Value
      m_CP43pty = "" & rsTmp.Fields("CP43pty").Value
      'end 2023/09/04
      
      ' ®Ö­ã³qª¾¤é 91.4.29 CANCEL
      'If IsNull(rsTmp.Fields("CP25")) = False Then
      '   textCP25 = TAIWANDATE(rsTmp.Fields("CP25"))
      'End If
      ' ©¼©Ò®×¸¹
      If IsNull(rsTmp.Fields("CP45")) = False Then
         textCP45 = rsTmp.Fields("CP45")
      End If
      ' ±ÂÅv´Á¶¡(¨´)
      'Add By Cheng 2002/07/19
      m_CP54 = Empty
      If IsNull(rsTmp.Fields("CP54")) = False Then
         m_CP54 = rsTmp.Fields("CP54")
      End If
      ' ³Q±ÂÅv¤H
      'Add By Cheng 2002/07/19
      m_CP50 = Empty
      If IsNull(rsTmp.Fields("CP50")) = False Then
         m_CP50 = rsTmp.Fields("CP50")
      End If
      ' ²¾Âà¤H
      'Add By Cheng 2002/07/19
      m_CP55 = Empty
      If IsNull(rsTmp.Fields("CP55")) = False Then
         m_CP55 = rsTmp.Fields("CP55")
      End If
      ' ²¾Âà¥Ó½Ð¤H¥N¸¹
      'Add By Cheng 2002/07/19
      m_CP56 = Empty
      If IsNull(rsTmp.Fields("CP56")) = False Then
         m_CP56 = rsTmp.Fields("CP56")
      End If
      'Add By Sindy 2013/1/11
      m_CP89 = Empty
      If IsNull(rsTmp.Fields("CP89")) = False Then
         m_CP89 = rsTmp.Fields("CP89")
      End If
      m_CP90 = Empty
      If IsNull(rsTmp.Fields("CP90")) = False Then
         m_CP90 = rsTmp.Fields("CP90")
      End If
      m_CP91 = Empty
      If IsNull(rsTmp.Fields("CP91")) = False Then
         m_CP91 = rsTmp.Fields("CP91")
      End If
      m_CP92 = Empty
      If IsNull(rsTmp.Fields("CP92")) = False Then
         m_CP92 = rsTmp.Fields("CP92")
      End If
      '2013/1/11 End
      '91.4.29 CANCEL
      ' ­Y¦¹¦¬¤å¸¹¤§¹ê»Úµ²ªG¬°1®É, «h±N­ã»é¤é¸m©ó®Ö­ã³qª¾¤éÄæ¦ì
      'If IsNull(rsTmp.Fields("CP24")) = False Then
      '   If rsTmp.Fields("CP24") = "1" Then
      '      If IsNull(rsTmp.Fields("CP25")) = False Then
      '         If IsEmptyText(rsTmp.Fields("CP25")) = False And rsTmp.Fields("CP25") <> "0" Then
      '            textCP25 = TAIWANDATE(rsTmp.Fields("CP25"))
      '         End If
      '      End If
      '   End If
      'End If
      ' ­Y®×¥ó©Ê½è¬°©µ®i®É, «h±N±ÂÅv´Á¶¡©ñ¤J±M¥Î´Á­­Äæ¦ì
      If m_CP10 = "102" Then
         If IsNull(rsTmp.Fields("CP53")) = False Then
            'edit by nickc 2008/01/10 §ï¦¨¦è¤¸¦~
            'textTM21 = TAIWANDATE(rsTmp.Fields("CP53"))
            textTM21 = DBDATE(rsTmp.Fields("CP53"))
         End If
         If IsNull(rsTmp.Fields("CP54")) = False Then
            'edit by nickc 2008/01/10 §ï¦¨¦è¤¸¦~
            'textTM22 = TAIWANDATE(rsTmp.Fields("CP54"))
            textTM22 = DBDATE(rsTmp.Fields("CP54"))
         End If
      End If
'      'Add By Sindy 2012/8/7 ÀË¬d®×¥ó³Æµù¸Ì¬O§_¦³"§ó§ïµù¥UÃÒ"¦r¼Ë,­Y¦³,¬O§_¬°ÃÒ®Ñ§ó§ï¹w³]¬°Y,¥X©w½Z®ÉIssueDate±aµoÃÒ¤é
      'modify by sonia 2019/4/30 §ï§PÂ_¬O§_¬°§ó¥¿¥B¨ä¬ÛÃöÁ`¦¬¤å¸¹¬°µù¥UÃÒFCT-038877
      'm_CP64 = ""
      'If IsNull(rsTmp.Fields("CP64")) = False Then
      '   m_CP64 = rsTmp.Fields("CP64")
      '   If InStr(rsTmp.Fields("CP64"), "§ó§ïµù¥UÃÒ") > 0 Then
      '      textMod.Text = "Y"
            'CANCEL BY SONIA 2015/6/22 ´ðûA»¡³£¤£­n±a,§ó§ï«áµoÃÒ¤é¤£·|©M­ì¨Ó¬Û¦PFCT-036102
            'Text1.Text = ChangeWStringToTString(m_TM20)
      '   End If
      'End If
      textMod.Text = ""
      'Modified by by Lydia 2023/09/04 §ï¥ÎÅÜ¼Æ
      'If m_CP10 = "302" Then
      '   Set rsSub = New ADODB.Recordset
      '   strSub = "SELECT * FROM CASEPROGRESS WHERE CP09='" & "" & rsTmp.Fields("CP43") & "' AND CP10='1701'"
      '   rsSub.CursorLocation = adUseClient
      '   rsSub.Open strSub, cnnConnection, adOpenStatic, adLockReadOnly
      '   If rsSub.RecordCount > 0 Then
      '      textMod.Text = "Y"
      '   End If
      '   rsSub.Close
      '   Set rsSub = Nothing
      'End If
      'end 2019/4/30
      If m_CP10 = "302" And m_CP43pty = "1701" Then
         textMod.Text = "Y"
      End If
      'end 2023/09/04
'      '2012/8/7 End
      'Add By Sindy 2012/10/12
      '¬O§_¬°¤@¥Ó½Ð®Ñ¦h¥ó
      m_CP148 = Empty
      If IsNull(rsTmp.Fields("CP148")) = False Then
         m_CP148 = rsTmp.Fields("CP148")
      End If
      'Modify By Sindy 2012/11/08 ¼W¥[ÀË¬d¦Pµo¤å¦r¸¹¬O§_¦³¦h¥ó
      If m_CP148 = "Y" Then
         If PUB_ChkIsOneAppMuchCase(m_CP28) = False Then
            m_CP148 = Empty
         End If
      End If
      '2012/10/12 End
   End If
   rsTmp.Close
   Set rsTmp = Nothing
    'Modify By Cheng 2003/12/16
    '¥Ó½Ð®Ö­ã©w½Z§ï¦b¦¹³B¥X
'   If m_CP10 = "101" Then textPrint = "N"
    'End
   ' 90.11.19 modify by sonia
   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   If textCP08 = "" Then
      textCP08 = "¡]" & strTmp & "¡^´¼°Ó¦r²Ä¸¹"
   End If
   '2010/12/22 ADD BY SONIA ¥xÆW®×®×¥ó©Ê½è¬°¥Ó½Ð¥B¬°§ïÅÜ­ì³B¤À®É, ²M°£°Ó¼Ð°ò¥»ÀÉªº¼f©w¸¹,§_«h¦bµoÃÒ«e·|¥H¬°¬O®Ö­ã¼f©w¸¹
   If m_CP10 = "101" And m_TM10 = "000" And frm03020401_03.GetSelectResult() = "2" Then
      textTM15 = ""
   End If
   '2010/12/22 END
   
   'Add By Cheng 2002/01/15
   m_strNumBegin = "°Ó"
   m_strNumEnd = "¦r"
   
   'Added by Morgan 2017/5/3 ¹q¤l¤½¤å
   If m_DocWord <> "" Then
      textCP08 = m_DocWord & "¦r²Ä" & PUB_GetEDocNo(m_DocNo) & "¸¹"
   ElseIf m_DocNo <> "" Then
      textCP08 = Replace(textCP08, "²Ä¸¹", "²Ä" & PUB_GetEDocNo(m_DocNo) & "¸¹")
   End If
   'end 2017/5/3
   
End Sub

Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   Dim strDay As String
   'Add By Cheng 2002/07/11
   Dim rsA As New ADODB.Recordset
   Dim StrSQLa As String
   Dim strFindCP43 As String, i As Integer
   
   ' ¨Ó¨ç¦¬¤å¤é
   textCP05S = m_CP05
   ' ¥»©Ò®×¸¹
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   ' Åª¨ú°Ó¼Ð°ò¥»ÀÉ
   QueryTradeMark
   ' Åª¨ú®×¥ó¶i«×ÀÉ
   QueryCaseProgress
   ' ¥H®×¥ó©Ê½è"®Ö­ã"©Î"§ïÅÜ­ì³B¤À"­pºâ©Ó¿ì´Á­­
''''edit by nickc 2007/10/11 §ï§ì¦³®É®Ä©Êªº
''''   strDay = Empty
   Select Case frm03020401_03.GetSelectResult
      Case "1":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1001")
            '2009/4/27 modify by soniaªü½¬»¡§ï¦^¥Á°ê¦~
            'edit by nickc 2008/03/13
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1001", DBDATE(m_CP05), , m_CP09))
            'textCP48 = DBDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1001", DBDATE(m_CP05), , m_CP09))
      Case "2":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1403")
            '2009/4/27 modify by soniaªü½¬»¡§ï¦^¥Á°ê¦~
            'edit by nickc 2008/03/13
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1403", DBDATE(m_CP05), , m_CP09))
            'textCP48 = DBDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1403", DBDATE(m_CP05), , m_CP09))
   End Select
''''   If IsEmptyText(strDay) = False Then
''''      ' 90.07.03 modify by louis (©Ó¿ì´Á­­¥H¹ê»Ú¤u§@¤Ñ¼Æ¨Ó­pºâ)
''''      'textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''      textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''   End If
   'Modify By Cheng 2002/04/29
'   ' ®×¥ó©Ê½è¬°¥Ó½Ð, ¥Ó½Ð°ê®a¬°¥xÆW®É, ¥H¼f©w¸¹¼Æ+°Ó¼ÐºØÃþ¥N¸¹§ì°Ó¼Ð¤½³øÀÉ, ±a¥X¨÷´Á
'   If m_CP10 = "101" And m_TM10 < "010" Then
'      strSQL = "SELECT * FROM TMBULLETIN " & _
'               "WHERE TMBM01 = '" & textTM15 & "' AND " & _
'                     "TMBM02 = '" & m_TM08 & "' "
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSQL, cnnConnection, adOpenDynamic
'      If rsTmp.RecordCount > 0 Then
'         rsTmp.MoveFirst
'         If IsNull(rsTmp.Fields("TMBM07")) = False Then
'            textTMBM07_1 = Mid(rsTmp.Fields("TMBM07"), 1, 2)
'            textTMBM07_2 = Mid(rsTmp.Fields("TMBM07"), 3, 3)
'         End If
'      End If
'      rsTmp.Close
'   End If
   ' ®×¥ó©Ê½è¬°©µ®i®É, ¤~¥i¿é¤J±M¥Î´Á­­
   'Modified by Lydia 2017/07/28 +301ÅÜ§ó®Ö­ã,¤ñ·Ó©µ®i®Ö­ã¿ì²z
   If m_CP10 = "102" Or m_CP10 = "301" Then
      textTM21.BackColor = &H80000005
      textTM22.BackColor = &H80000005
      textTM21.Locked = False
      textTM22.Locked = False
      textTM21.TabStop = True
      textTM22.TabStop = True
      cmdMod.Visible = False 'Added by Lydia 2016/07/19
   Else
      textTM21.BackColor = &H8000000F
      textTM22.BackColor = &H8000000F
      textTM21.Locked = True
      textTM22.Locked = True
      textTM21.TabStop = False
      textTM22.TabStop = False
      cmdMod.Visible = True 'Added by Lydia 2016/07/19
   End If
   'Add By Cheng 2002/07/11
   '­Y®×¥ó©Ê½è¬°"¥Ó½Ð"(101)®É
   'edit by nick 2004/12/23 ¤À³Î»P¥Ó½Ð°µ¬Û¦Pªº¨Æ±¡
   'If m_CP10 = "101" Then
   If m_CP10 = "101" Or m_CP10 = "308" Then
      'Modify By Cheng 2002/07/22
      '¬O§_§ó·s°ò¥»ÀÉ¥Ø«e­ã»é¹w³]¬°"1"
'      '¬O§_§ó·s°ò¥»ÀÉ¥Ø«e­ã»é¹w³]¬°"Y"
'      Me.textTM16S.Text = "Y"
      Me.textTM16S.Text = "1"
      'Åã¥Ü©Ó¿ì¤H¸ê®Æ
      'Modify By Sindy 2012/7/6
'      Me.textCP14.Text = m_CP14
'      Me.textCP14_2.Text = GetStaffName(Me.textCP14.Text)
      Me.textCP14.Text = strUserNum
      Me.textCP14_2.Text = strUserName
      '2012/7/6 End
   'Modify By Sindy 2012/7/6 ¯S©w®×¥ó©Ê½è®Ö­ã®É¹w³]¬°"¿é¤J¤§µ{§Ç¤H­û½s¸¹":½Ð±a¤U¦C®×¥ó©Ê½è¿é¤J®Ö­ã®É¤§©Ó¿ì¤H¬°¾Þ§@¤H­û
   '¥Ó½Ð (101), ©µ®i(102), ¸É´«µoÃÒ®Ñ(103), ÅÜ§ó(301), §ó¥¿(302), ¥Ó½Ð­^¤åÃÒ©ú(304)
   '¦Û½ÐºM¦^(306), ¦Û½Ð©ß±ó°Ó¼ÐÅv(307), ¤À³Î(308), ´îÁY°Ó«~(313), ²¾Âà(501), ±ÂÅv(502)
   '¼o¤î±ÂÅv(503), ¦A±ÂÅv(504), ¼o¤î¦A±ÂÅv(505), ³]©w½èÅv(506), ºM¾P³]©w½èÅv(507), °h¶O(725)
   ElseIf m_CP10 = "102" Or m_CP10 = "103" Or m_CP10 = "301" Or m_CP10 = "302" Or m_CP10 = "304" Or _
      m_CP10 = "306" Or m_CP10 = "307" Or m_CP10 = "313" Or m_CP10 = "501" Or m_CP10 = "502" Or _
      m_CP10 = "503" Or m_CP10 = "504" Or m_CP10 = "505" Or m_CP10 = "506" Or m_CP10 = "507" Or m_CP10 = "725" Then
      'Åã¥Ü©Ó¿ì¤H¸ê®Æ
      Me.textCP14.Text = strUserNum
      Me.textCP14_2.Text = strUserName
   '2012/7/6 End
   '¨ä¥L®×¥ó©Ê½è®É
   Else
      'Modify By Cheng 2002/07/22
'      Me.textTM16S.Text = "N"
'      strSQLA = "Select CP13 From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " AND CP10='101' "
'      rsA.CursorLocation = adUseClient
'      rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'         Me.textCP14.Text = "" & rsA.Fields(0).Value
'         Me.textCP14_2.Text = GetStaffName(Me.textCP14.Text)
'      End If
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
        'Add By Cheng 2003/10/08
        '¹w³]©Ó¿ì¤H
        Me.textCP14.Text = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
        Me.textCP14_2.Text = GetStaffName(Me.textCP14.Text)
   End If
   ' «DAÃþ¦¬¤å¨ä¹w³]¬°¤£¥iºâ®×¥ó¼Æ
   textCP26 = "N"
   Set rsTmp = Nothing
    'Add By Cheng 2003/10/01
    'Begin
    Select Case m_CP10
    Case "302" '§ó¥¿
        '¬O§_¬°ÃÒ®Ñ§ó§ï
        Me.Label25.Visible = True
        Me.textMod.Visible = True
        Me.textMod.Enabled = True
        Me.Label18.Visible = True
        '¬O§_²£¥Íµù¥UÃÒ½Ð´Ú¸ê®Æ
        Me.Label29.Visible = True
        Me.Text3.Visible = True
        Me.Text3.Enabled = True
        Me.Label30.Visible = True
        '©w½Z®×¥ó©Ê½è
        '92.10.24 CANCEL BY SONIA
        'Me.Label32.Visible = True
        'Me.Combo1.Visible = True
        'Me.Combo1.Enabled = True
        '92.10.24 END
        '¬O§_§ó§ïÃÒ®Ñ
        Me.Label1(12).Visible = False
        Me.Text2.Visible = False
        Me.Text2.Enabled = False
        Me.Label1(13).Visible = False
        '2011/7/12 add by sonia
        'Me.Label10.Caption = "­ì¨ç¤½§i¤é :" 'Mark by Lydia 2023/09/04 ¥t³]Äæ¦ì
        Me.textTM14.Text = ""
        textTMBM07_1.Enabled = False
        textTMBM07_2.Enabled = False
        '2011/7/12 end
    Case Else '¨ä¥L®×¥ó©Ê½è
        '¬O§_¬°ÃÒ®Ñ§ó§ï
        Me.Label25.Visible = False
        Me.textMod.Visible = False
        Me.textMod.Enabled = False
        Me.Label18.Visible = False
        '¬O§_²£¥Íµù¥UÃÒ½Ð´Ú¸ê®Æ
        Me.Label29.Visible = False
        Me.Text3.Visible = False
        Me.Text3.Enabled = False
        Me.Label30.Visible = False
        '©w½Z®×¥ó©Ê½è
        '92.10.24 CANCEL BY SONIA
        'Me.Label32.Visible = False
        'Me.Combo1.Visible = False
        'Me.Combo1.Enabled = False
        '92.10.24 END
        '¬O§_§ó§ïÃÒ®Ñ
        Me.Label1(12).Visible = True
        Me.Text2.Visible = True
        Me.Text2.Enabled = True
        Me.Label1(13).Visible = True
        '2011/7/12 add by sonia
        Me.Label10.Caption = "¤½§i¤é :"
        textTMBM07_1.Enabled = True
        textTMBM07_2.Enabled = True
        '2011/7/12 end
    End Select
    'End
    '92.10.24 ADD BY SONIA
    Me.Label32.Visible = True
    Me.Combo1.Visible = True
    Me.Combo1.Enabled = True
    '92.10.24 END
    'Add By Sindy 2011/11/4 FCT-016964¦]µo¥Í¤H¬°ÂI¿ï¿ù»~,¾É¦Ü²£¥Í©w½Z¬°¿ù»~ªº,¨t²Î¥ý¹w³]±a«DCÃþ¬ÛÃö¤å¸¹ªº®×¥ó©Ê½è
    If m_CP10 = "302" Then '§ó¥¿
      'Modified by Lydia 2023/09/04 §ï¥ÎÅÜ¼Æ
      'strSql = "select cp09,cp10,cp43 from caseprogress where cp09='" & m_CP09 & "'"
      'intI = 1
      'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      'If intI = 1 Then
      '   strFindCP43 = "" & RsTemp.Fields("cp43")
         strFindCP43 = m_CP43
      'end 2023/09/04
         Do While strFindCP43 <> ""
            'Modified by Lydia 2023/09/04
            'strSql = "select cp09,cp10,cp43 from caseprogress where cp09='" & strFindCP43 & "'"
            strSql = "select cp09,cp10,cp43,cpm03 from caseprogress,casepropertymap where cp09='" & strFindCP43 & "' and cp01=cpm01(+) and cp10=cpm02(+) "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            '«DCÃþªº¬ÛÃöÁ`¦¬¤å¸¹
            If Left(strFindCP43, 1) < "C" Then
               If intI = 1 Then
                  'Åª¨ú®×¥ó©Ê½è
                  'Modified by Lydia 2023/09/04 ©w½Z®×¥ó©Ê½è§ï¦¨±a¤J«DCÃþ¬ÛÃö¤å¸¹ªº®×¥ó©Ê½è¡A¤£¨Ï¥Î¯S©w®×¥ó©Ê½è²M³æ¡C
                  'For i = 0 To Combo1.ListCount - 1
                   '  If Trim(Left(Combo1.List(i), 4)) = RsTemp.Fields("cp10") Then
                   '     Combo1.ListIndex = i
                   '     Exit Do '§ä¨ì,Â÷¶}°j°é,µ{¦¡µ²§ô
                   '  End If
                  'Next i
                  Combo1.Clear
                  Combo1.AddItem RsTemp.Fields("cp10") & " " & RsTemp.Fields("cpm03")
                  Combo1.ListIndex = 0
                  Exit Do
                  'end 2023/09/04
               End If
               Exit Do 'µL¸ê®Æ,Â÷¶}°j°é,µ{¦¡µ²§ô
            Else
               If intI = 1 Then
                  strFindCP43 = "" & RsTemp.Fields("cp43")
               Else
                  strFindCP43 = ""
               End If
            End If
         Loop
      'End If 'Mark by Lydia 2023/09/04
    End If
    '2011/11/4 End
   '2012/4/25 add by sonia 92.11.28 ¥H«á¥Ó½Ðªº®×¥ó±Hµù¥UÃÒ®É¤£½Ð´Ú
   If DBDATE(Val(m_TM11)) >= 20031128 Then
        Text3.Locked = True
   End If
   
   'Added by Lydia 2023/09/04 ­×§ï¤é¤å²Õ¤§®Ö­ã-§ó¥¿(©µ®i®Ö­ã¨ç)¤§©w½Z¤ÎÄ¶¤å:
   txtADate = "": txtADate.Visible = False: lblADate.Visible = False: txtADate.Locked = False
   If m_CP10 = "302" And GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "3" Then
      'Åã¥Ü¬ÛÃö¦¬¤å¸¹ªº"­ì¨ç¤½§i¤é"¡A­Y"­ì¨ç¤½§i¤é"¦³±a¤J¤é´Á«hÂê©wÄæ¦ì¤£¥iÅÜ§ó¡A¨S¦³¤é´Á¶}©ñÅý¨Ï¥ÎªÌ¿é¤J¡A¤£¦^¦s¬ÛÃö¦¬¤å¸¹¤§"­ì¨ç¤½§i¤é"
      strSql = "select cp143 from caseprogress where cp09='" & m_CP43 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         txtADate.Visible = True: lblADate.Visible = True
         If "" & RsTemp.Fields("cp143") <> "" Then
            txtADate = TransDate("" & RsTemp.Fields("cp143"), 1)
            txtADate.Locked = True
         End If
      End If
   End If
   'end 2023/09/04
   
   'Add By Sindy 2013/1/11
   '­Y¸Óµ§²¾Âà©ÎÅý»Pªº¨üÅý¤H(5­Ó),»P°ò¥»ÀÉ¤£²Å®É,Åã¥Ü°T®§¥B¤£¥i¿é¤J®Ö­ã¨ç
   cmdok.Enabled = True
   If m_CP10 = "501" Then
      If m_TM23 <> m_CP56 Or m_TM78 <> m_CP89 Or m_TM79 <> m_CP90 Or m_TM80 <> m_CP91 Or m_TM81 <> m_CP92 Then
         MsgBox "¦¹®×°ò¥»ÀÉ¥Ó½Ð¤H»P¦¹µ{§Ç¨üÅý¤H¤£¦P¡A½Ð½T»{¸ê®Æ¡I"
         cmdok.Enabled = False
      End If
   End If
   '2013/1/11 End
   
   'Add by Sindy 2020/5/19 ¬O§_°±¤î¶l°È
   Call GetPrjPeopleNum6(m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04, "NA86", m_NA86)
   
   'Added by Lydia 2017/06/27 ©µ®i¡B²¾Âà¡BÅÜ§ó(102,501,301)®×¤§®Ö­ã¿é¤J¡A¼W¥[¡u¼È¤£¦C¦L©w½Z¡v
   If InStr("102,301,501", m_CP10) > 0 Then
      Chk1.Visible = True
      Chk1.Value = True '¹w³]¤Ä¿ï
   Else
      Chk1.Visible = False
      Chk1.Value = 0
   End If
   'end 2017/06/27
   
End Sub

Private Sub DisplayNextForm()
   frm03020401_05.SetData 0, m_TM01, True
   frm03020401_05.SetData 1, m_TM02, False
   frm03020401_05.SetData 2, m_TM03, False
   frm03020401_05.SetData 3, m_TM04, False
   frm03020401_05.SetData 4, m_CP09, False
   Me.Hide
   frm03020401_05.Show
   frm03020401_05.QueryData
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
Dim strSql As String
Dim strCP06 As String
Dim strCP07 As String
Dim strCP09 As String
Dim strCP10 As String
'Dim strCP12 As String
Dim strCP27 As String
Dim strNP07 As String
Dim strNP08 As String
Dim strNP09 As String
Dim strNP14 As String
Dim strNP15 As String
Dim strNP22 As String
'92.2.9 add by sonia
Dim m_Work20 As String
'Add By Cheng 2003/10/08
Dim strCP09BKind As String '·s¼WªºBÃþ¦¬¤å¸¹
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCP118 As String 'Add by Amy 2023/02/06 ¬O§_¹q¤l°e¥ó
           
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' ­Y«eµe­±©Ò¿ï¾Üªºµ²ªG¬°1®É, §ó·s­ì®×¥ó¶i«×¸ê®Æªº¹ê»Úµ²ªG¬°­ã¤Î­ã»é¤é
   If frm03020401_03.GetSelectResult() = "1" Then
      '91.4.29 MODIFY BY SONIA
      'strSQL = "UPDATE CaseProgress SET CP24 = '1', CP25 = " & DBDATE(textCP25) & " " & _
      '         "WHERE CP09 = '" & m_CP09 & "' AND " & _
      '               "(CP24 IS NULL OR CP24 = '' OR CP24 = ' ')"
      strSql = "UPDATE CaseProgress SET CP24 = '1', CP25 = " & DBDATE(m_CP05) & " " & _
               "WHERE CP09 = '" & m_CP09 & "' AND " & _
                     "(CP24 IS NULL OR CP24 = '' OR CP24 = ' ')"
      '91.4.29 END
      cnnConnection.Execute strSql
   End If
    'Modify By Cheng 2003/01/16
    '¥[®×©Ê½è½è¦A±ÂÅv(504)
'   'Add By Cheng 2002/06/14
'   If m_CP10 = "502" Or m_CP10 = "506" Then
   If m_CP10 = "502" Or m_CP10 = "504" Or m_CP10 = "506" Then
      strSql = "UPDATE CaseProgress SET CP53 = " & DBDATE(Me.textCP53.Text) & ", CP54 = " & DBDATE(Me.textCP54.Text) & " " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Modify By Cheng 2002/07/22
   '¨ú®ø§ó·s°Ó¼Ð°ò¥»ÀÉ¤§±M¥ÎÅv¬O§_¦s¦b
'   ' §ó·s°Ó¼Ð°ò¥»ÀÉ¤§±M¥ÎÅv¬O§_¦s¦b
'   If Not IsNull(textTM17) Then
'      strSQL = "UPDATE TradeMark SET TM17 = '" & textTM17 & "' " & _
'            "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                  "TM02 = '" & m_TM02 & "' AND " & _
'                  "TM03 = '" & m_TM03 & "' AND " & _
'                  "TM04 = '" & m_TM04 & "' "
'      cnnConnection.Execute strSQL
'   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' ­Y®×¥ó©Ê½è¬°©µ®i®É, §ó·s°Ó¼Ð°ò¥»ÀÉ¤§±M¥Î´Á­­Äæ¦ì
   If m_CP10 = "102" Then
      strSql = "UPDATE TradeMark SET TM21 = " & DBDATE(textTM21) & ", " & _
                                    "TM22 = " & DBDATE(textTM22) & " " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
      cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'add by nick 2004/09/14
   ' ­Y®×¥ó©Ê½è¬°307®É, §ó·s°Ó¼Ð°ò¥»ÀÉ¤§¬O§_³¬¨÷=Y
   If m_CP10 = "307" Then
      strSql = "UPDATE TradeMark SET TM29='Y' " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
      cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' ®×¥ó©Ê½è¬°¥Ó½Ð®É
   'edit by nick 2004/12/23 ¤À³Î»P¥Ó½Ð°µ¬Û¦Pªº¨Æ±¡
   'If m_CP10 = "101" Then
   If m_CP10 = "101" Or m_CP10 = "308" Then
      ' §ó·s¼f©w¸¹, ¤½§i¤é, ¼f©w¨Ó¨ç¤é(¨Ó¨ç¦¬¤å¤é)
      '93.7.2 MODIFY BY SONIA
      'strSQL = "UPDATE TradeMark SET TM15 = '" & textTM15 & "'," & _
      '                              "TM14 = " & DBDATE(textTM14) & ", " & _
      '                              "TM13 = " & DBDATE(m_CP05) & " " & _
      '         "WHERE TM01 = '" & m_TM01 & "' AND " & _
      '               "TM02 = '" & m_TM02 & "' AND " & _
      '               "TM03 = '" & m_TM03 & "' AND " & _
      '               "TM04 = '" & m_TM04 & "' "
      strSql = "UPDATE TradeMark SET TM15 = '" & textTM15 & "'," & _
                                    "TM14 = " & DBNullDate(textTM14) & ", " & _
                                    "TM13 = " & DBNullDate(m_CP05) & " " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "' "
      '93.7.2 END
      cnnConnection.Execute strSql
      'Modify By Cheng 2002/07/22
      '·í®×¥ó©Ê½è¬°°Ó¥Ó®É(101), §ó·s¥Ø«e­ã/»é¬°­ã¤Î¼f©w¨Ó¨ç¤é(®Ö­ã³qª¾¤é)¨â­ÓÄæ¦ì
'      ' ·í¨Ï¥ÎªÌ¿é¤J­n§ó·s°ò¥»ÀÉ¤§­ã/»é®É, §ó·s¥Ø«e­ã/»é¬°­ã¤Î¼f©w¨Ó¨ç¤é(®Ö­ã³qª¾¤é)¨â­ÓÄæ¦ì
'      If textTM16S = "Y" Then
      'edit by nick 2004/12/23 ¤À³Î»P¥Ó½Ð°µ¬Û¦Pªº¨Æ±¡
      'If m_CP10 = "101" Then
      If m_CP10 = "101" Or m_CP10 = "308" Then
         '91.4.29 MODIFY BY SONIA
         'strSQL = "UPDATE TradeMark SET TM16='1'," & _
         '                              "TM13=" & DBDATE(textCP25) & " " & _
         '         "WHERE TM01 = '" & m_TM01 & "' AND " & _
         '               "TM02 = '" & m_TM02 & "' AND " & _
         '               "TM03 = '" & m_TM03 & "' AND " & _
         '               "TM04 = '" & m_TM04 & "' "
         'Modify By Sindy 2010/01/05 ¼W¥[§ó·s¡u©ñ±ó±M¥ÎÅv¡vÄæ¦ì
'         strSQL = "UPDATE TradeMark SET TM16='1'," & _
'                                       "TM13=" & DBDATE(m_CP05) & " " & _
'                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                        "TM02 = '" & m_TM02 & "' AND " & _
'                        "TM03 = '" & m_TM03 & "' AND " & _
'                        "TM04 = '" & m_TM04 & "' "
         strSql = "UPDATE TradeMark SET TM16='1'," & _
                                       "TM13=" & DBDATE(m_CP05) & "," & _
                                       "TM67='" & ChgSQL(textTM67) & "' " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
         '91.4.29 END
         cnnConnection.Execute strSql
      End If
   End If
   '2005/8/2 MODIFY BY SONIA ³¯ª÷½¬­n¨D¯d¤U ©µ®i102¡BÅÜ§ó301
   '92.2.9 add by sonia ©µ®i102¡B¸Éµoµù¥UÃÒ103¡BÅÜ§ó301¡B²¾Âà201¡B±ÂÅv202¡B¦A±ÂÅv504¡B½èÅv506 ®Ö­ã®É­n¦V«È¤á½Ð´Ú
   If frm03020401_03.GetSelectResult() = "1" Then
      Select Case m_CP10
         '2005/8/2 MODIFY BY SONIA
         'Case "102", "103", "301", "501", "502", "504", "506"
         '2007/6/7 ¥[´îÁY°Ó«~313
         'Modify By Sindy 2010/01/27 301.ÅÜ§ó¤]­n¤WN
         'Case "102", "301", "313"
         '2010/6/11 MODIFY BY SONIA ªü½¬­n¨D¦Û2010/6/1°_©µ®i®Ö­ã¤]¤£½Ð´Ú
         'Case "102", "313"
         Case "313"
            m_Work20 = ""
         Case Else
            m_Work20 = "N"
      End Select
   End If
   '92.2.9 end
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '  ·s¼W¸ê®Æ¨ì®×¥ó¶i«×ÀÉ
   ' ¦¬¤å¸¹
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   m_NickCp09 = strCP09
   ' ®×¥ó©Ê½è¬°®Ö­ã©Î§ïÅÜ­ì³B¤À
   strCP10 = "1001"
   Select Case frm03020401_03.GetSelectResult
      Case "1", "2": '2006/11/1 MODIFY BY SONIA §ïÅÜ­ì³B¤À¤]­n±¾´Á­­
         Select Case frm03020401_03.GetSelectResult
            Case "1": strCP10 = "1001"
            Case "2": strCP10 = "1403"
         End Select
        'Add By Cheng 2003/11/19
        '­Y¬°°Ó¥Ó®×¥B¥»®×¥Ó½Ð¤é¬°921128(§t)¥H«áªÌ
        'edit by nick 2004/12/23 ¤À³Î»P¥Ó½Ð°µ¬Û¦Pªº¨Æ±¡
        'If m_CP10 = "101" Then
         If m_CP10 = "101" Or m_CP10 = "308" Then
            If Val(m_TM11) >= 20031128 Then
               'ªk©w´Á­­
               strCP07 = DBDATE(DateAdd("m", 2, ChangeWStringToWDateString(DBDATE(m_CP05))))
               '¥»©Ò´Á­­
               'edit by nick 2004/07/28 §ï¬°´î 4 ¤Ñ
               'strCP06 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strCP07))))
               'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
               If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                  strCP06 = PUB_GetOurDeadline(DBDATE(strCP07))
               Else
               '2014/10/6 END
                  strCP06 = DBDATE(DateAdd("d", -4, ChangeWStringToWDateString(DBDATE(strCP07))))
               End If
               strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
               'add by nick 2004/10/28 ¬ö¿ý·sªº¥»©Ò´Á­­
               m_CP06 = DBDATE(strCP06)
               m_CP07 = DBDATE(strCP07)   '2014/12/9 ADD BY SONIA ©w½Z§ï³qª¾ªk©w´Á­­(­ì¬°¥»©Ò´Á­­)
            End If
         End If
   End Select
   
    'Modify By Cheng 2003/04/07
    '´¼Åv¤H­û¦s³Ìªñ¦¬¤åAÃþ±µ¬¢°O¿ý³æªº´¼Åv¤H­û
'edit by nick 2004/08/03 ¥[¤J·í cp06 ©Î cp07 ¦³­È®É¡A­n¥[¤J cp06,cp07
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP35,CP43) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                    "'" & m_Work20 & "','" & textCP26 & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "') "
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,cp06,cp07,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP35,CP43) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & IIf(Trim(strCP06) = "", "NULL", strCP06) & "," & IIf(Trim(strCP07) = "", "NULL", strCP07) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                    "'" & m_Work20 & "','" & textCP26 & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "') "
   cnnConnection.Execute strSql
   
   'Add By Cheng 2003/11/19
   '­Y¬°°Ó¥Ó®×¥B¥»®×¥Ó½Ð¤é¬°921128(§t)¥H«áªÌ
   'edit by nick 2004/12/23 ¤À³Î»P¥Ó½Ð°µ¬Û¦Pªº¨Æ±¡
   'If m_CP10 = "101" Then
    If m_CP10 = "101" Or m_CP10 = "308" Then
       If Val(m_TM11) >= 20031128 Then
           '2005/7/19 MODIFY BY SONIA
           'strSQLA = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10 IN ('715','717') "
           StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10 IN ('715','717') AND CP57 IS NULL"
           '2005/7/19 END
           rsA.CursorLocation = adUseClient
           rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
           '­Y¦³¦¬¤å²Ä¤@´Áµù¥U¶O, §ó·s¶i«×ÀÉ
           If rsA.RecordCount > 0 Then
               StrSQLa = "Update CaseProgress Set CP06=" & strCP06 & ", CP07=" & strCP07 & " Where CP09='" & rsA("CP09").Value & "' "
               cnnConnection.Execute StrSQLa
           '­Y¥¼¦¬¤å²Ä¤@´Áµù¥U¶O, ·s¼W¤U¤@µ{§ÇÀÉ
           Else
               'Modify By Sindy 2012/6/27 °Ó¼Ð­×ªk
               If Val(DBDATE(m_CP05)) >= 20120701 Then
                  strNP07 = "717"
               Else
               '2012/6/27 End
                  strNP07 = "715"
               End If
               strNP22 = GetNextProgressNo()
               strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                               "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                               DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
               cnnConnection.Execute strSql
               ' ¦C¦L°ê¤º®×¥ó±µ¬¢¤Îµ²®×°O¿ý³æ
               'g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
               '·s¼W¦C¦L±µ¬¢µ²®×³æ¸ê®Æ
               pub_AddressListSN = pub_AddressListSN + 1
               PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
           End If
           'add by nick 2004/10/28 ¬ö¿ý·sªº¥»©Ò´Á­­
           If rsA.State <> adStateClosed Then rsA.Close
           Set rsA = Nothing
       End If
    'Added by Lydia 2023/09/04 «D101¥Ó½Ð©Î308¤À³Î¤§®Ö­ã¡A±N¿é¤J¤§¤½§i¤é°O¿ý¬°"­ì¨ç¤½§i¤éCP143"¡C
    ElseIf textTM14.Text <> "" Then
       strSql = "Update CaseProgress Set CP143=" & DBDATE(textTM14) & " Where CP09='" & strCP09 & "' "
       cnnConnection.Execute strSql
    'end 2023/09/04
    End If
   
   '92.11.20 ADD BY SONIA
   If strCP10 = "1403" Then
       strSql = "Update CaseProgress Set CP24='1' Where CP09='" & strCP09 & "' "
       cnnConnection.Execute strSql
   End If
   '92.11.20 END
    'Add By Cheng 2003/09/05
    '·s¼W¤º³¡¦¬¤å
    '2009/4/22 MODIFY BY SONIA §ó§ïµù¥UÃÒ->§ó§ï®Ö­ã¨ç
    If Me.Text2.Text <> "" Then
        strCP09BKind = AutoNo("B", 6)
        '2009/4/22 modify by sonia ¨ú®øµo¤å¤é, ¦]¬°°t¦Xµo¤å«Ç¹q¸£¤ÆÀ³©óªü½¬§Pµo®É¤~¤Wµo¤å¤é
        'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP32, CP43, CP64,CP20) " & _
                        "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & _
                        "'" & strCP09BKind & "','302','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
                        "'N'," & strSrvDate(1) & ",'N','" & strCP09 & "','§ó§ï®Ö­ã¨ç','N') "
        '2017/1/11 modify by sonia CP26§ï¬°­n­p¥ó
        'Modify by Amy 2023/02/06 +CP118 ¬O§_¹q¤l°e¥ó
        'Modify by Amy 2023/03/06 ­ì§PÂ_TM136='1'¤~³]BÃþ§ó¥¿¬°¹q¤l°e¥ó,§ï³£³]¹q¤l°e¥ó
        'strCP118 = IIf(Pub_GetField("TradeMark", "tm01||tm02||tm03||tm04='" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & "'", "TM136") = "1", "Y", "")
        strCP118 = "Y"
        'end 2023/03/06
        strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP32, CP43, CP64,CP20,CP118) " & _
                        "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & _
                        "'" & strCP09BKind & "','302','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
                        "'','N','" & strCP09 & "','§ó§ï®Ö­ã¨ç','N'," & CNULL(ChgSQL(strCP118)) & " ) "
        cnnConnection.Execute strSql
        'end 2023/02/06
        ' ­Y¦³¿é¤J©Ó¿ì¤H®É
        If IsEmptyText(textCP14) = False Then
           strSql = "UPDATE CaseProgress SET CP14 = '" & textCP14 & "' " & _
                    "WHERE CP09 = '" & strCP09BKind & "' "
           cnnConnection.Execute strSql
        End If
        '·s¼WªºCÃþ¨Ó¨ç©Ê½è¬°®Ö­ã, «h©Ó¿ì¤H¬°µ{§Ç¤H­û, µo¤å¤é¬°¨t²Î¤é
        If strCP10 = "1001" Then
           '2009/4/22 modify by sonia ¨ú®øµo¤å¤é, ¦]¬°°t¦Xµo¤å«Ç¹q¸£¤ÆÀ³©óªü½¬§Pµo®É¤~¤Wµo¤å¤é
           'strSQL = "UPDATE CaseProgress SET CP14 = '" & strUserNum & "',CP27= " & ServerDate & " " & _
                    "WHERE CP09 = '" & strCP09BKind & "' "
           strSql = "UPDATE CaseProgress SET CP14 = '" & strUserNum & "' " & _
                    "WHERE CP09 = '" & strCP09BKind & "' "
           cnnConnection.Execute strSql
        End If
    End If
   ' ­Y¦³¿é¤J©Ó¿ì¤H®É
   If IsEmptyText(textCP14) = False Then
      strSql = "UPDATE CaseProgress SET CP14 = '" & textCP14 & "' " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   ' ­Y¦³¿é¤J©Ó¿ì´Á­­®É
   If IsEmptyText(textCP48) = False Then
      strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(textCP48) & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
    'Add By Cheng 2002/12/18
    '·s¼WªºCÃþ¨Ó¨ç©Ê½è¬°®Ö­ã, «h©Ó¿ì¤H¬°µ{§Ç¤H­û, µo¤å¤é¬°¨t²Î¤é
   'If StrCP10 = "1001" Then    '2010/12/15 cancel by sonia ªü½¬»¡§ïÅÜ­ì³B¤À¤]­n¤Wµo¤å¤éFCT-029223
      strSql = "UPDATE CaseProgress SET CP14 = '" & strUserNum & "',CP27= " & ServerDate & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   'End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' §ó·s¤U¤@µ{§ÇÀÉ®×¥ó©Ê½è¬°¶Ê¼fªº¸ê®Æ
   strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP07 = " & "305"
   cnnConnection.Execute strSql
   
   'add by sonia 2017/6/8 ¥Ó½Ð®×®Ö­ã¦P®É±N¥Ó½Ð·N¨£®Ñ´Á­­¸Ñ°£FCT-038905
   If m_CP10 = "101" And strCP10 = "1001" Then
      strSql = "UPDATE NextProgress SET NP06 = 'N',NP15='°Ó¥Ó®×¤w®Ö­ã;'||NP15 " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = " & "202 AND NP06 IS NULL"
      cnnConnection.Execute strSql
   End If
   'end 2017/6/8
   
   '92.03.27 nick
   ' §ó·s¤U¤@µ{§ÇÀÉ®×¥ó©Ê½è¬°§ïÅÜ­ì³B¤Àªº¸ê®Æ
   If frm03020401_03.textResult.Text = "2" Then
        strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
                 "WHERE NP01 = '" & m_CP09 & "' AND " & _
                       "NP02 = '" & m_TM01 & "' AND " & _
                       "NP03 = '" & m_TM02 & "' AND " & _
                       "NP04 = '" & m_TM03 & "' AND " & _
                       "NP05 = '" & m_TM04 & "' AND " & _
                       "NP07 = " & "1403"
        cnnConnection.Execute strSql
   End If
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' ¨Ì®×¥ó©Ê½è¨Ó¨M©w¬O§_­n·s¼W¤@µ§¸ê®Æ¨ì¤U¤@µ{§ÇÀÉ
   Select Case m_CP10
      ' ©µ®i
      Case "102":
         'ªk©w´Á­­
         strNP09 = DBDATE(textTM22)
         '¥»©Ò´Á­­
        'Modify By Cheng 2003/09/02
'         strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP09)), Val(DBMONTH(strNP09)) - GetDelayTime(m_TM10), Val(DBDAY(strNP09)))))
         '2006/1/16 MODIFY BY SONIA
         'strNP08 = DBDATE(DateAdd("m", -GetDelayTime(m_TM10), ChangeWStringToWDateString(DBDATE(strNP09))))
         'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
         If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
            strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
         Else
         '2014/10/6 END
            strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
         End If
         strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
         '2006/1/16 END
         strNP22 = GetNextProgressNo()
        'Modify By Cheng 2003/04/07
        '´¼Åv¤H­û¦s³Ìªñ¦¬¤åAÃþ±µ¬¢°O¿ý³æªº´¼Åv¤H­û
        'Modify By Cheng 2003/09/05
'         strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                  "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & "102" & "," & _
'                          strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & "102" & "," & _
                          strNP08 & "," & strNP09 & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
         cnnConnection.Execute strSql
   End Select
   
    'Add By Cheng 2003/09/05
    '­Y³]©w²£¥Í½Ð´Ú¸ê®Æ
    If Me.Text3.Text = "Y" Then
        '³]©w­n¦C¦L¦a§}±ø
'        m_blnPrintAddress = True
       '·s¼W°ê¥~½Ð´Ú¸ê®Æ
       Dim strAgentNo As String '¥N²z¤H½s¸¹
       Dim strPrintCust  As String '¬O§_¦C¦L¥Ó½Ð¤H
       Dim dblUSRate As Double '¬üª÷¶×²v
       Dim strDisc As String '§é¦©
        Dim strA1K27 As String '¦C¦L¹ï¶H
        Dim strA1K28 As String '½Ð´Ú¹ï¶H
       
       '1:¥ý¥H"X"§ìACC1R0¤§°ê¥~½Ð´Ú³æªº¦Û°Ê½s¸¹, ¨Ã§ó·s¨ä¬y¤ô¸¹
       m_strSerialNo = AccAutoNo(MsgText(815), 5)
       AccSaveAutoNo MsgText(815), Right(m_strSerialNo, 5)
       '2:·s¼WACC1K0
'       strAgentNo = GetAgentNO
       strAgentNo = PUB_GetA1K03(m_TM01, m_TM02, m_TM03, m_TM04)
       strPrintCust = PUB_GetA1K04(m_TM01, m_TM02, m_TM03, m_TM04)
       'dblUSRate = GetUSRate
       
        strA1K27 = PUB_GetA1K27(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10)
        If strA1K27 = "" Then strA1K27 = strAgentNo
        strA1K28 = PUB_GetA1K28(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10)
        If strA1K28 = "" Then strA1K28 = strAgentNo
        
        'Added by Lydia 2014/12/15 ½Ð´Ú³æ½Ð§ï¬°¨Ì¥N²z¤H©Î«È¤áÀÉ³]©wªº½Ð´Ú¹ô§O
         Dim strA1K33 As String, strA1K18 As String
         'Modify By Sindy 2016/11/30
         'strA1K33 = PUB_GetInitCurrPrintType(m_TM01, strA1K28, strA1K18, dblUSRate)
         'Modified by Morgan 2018/4/27 +strA1K27
         strA1K33 = PUB_GetInitCurrPrintType(m_TM01, strA1K28, strA1K18, dblUSRate, m_TM02, m_TM03, m_TM04, strA1K27)
         '2016/11/30 END
       
       strDisc = 1 - (PUB_GetA1L07Disc(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, strSrvDate(2)) / 100)
        '§éÅý¤é´Á¦sNULL, §@¼o¤é´Á¦sNULL
        'Modify By Cheng 2004/01/07
        'A1K11­n¥ý¦©°£§é¦©«á¤~¦sÀÉ
        'Modify By Cheng 2004/04/26
        '¬üª÷¨ú¦Ü¾ã¼Æ¦ì(µL±ø¥ó±Ë¥h)
'       strSQL = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K17,A1K18,A1K19,A1K20,A1K21,A1K25,A1K26,A1K29,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04) " & _
'                "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",0,NULL,0," & dblUSRate & "," & 3500 - (3000 * Val(strDisc)) & ",NULL,'" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','','USD'," & Val(ACDate(ServerDate)) & "," & ServerTime & ",'" & strUserNum & "','','','',0," & IIf(dblUSRate = 0, 3500 - (3000 * Val(strDisc)), Format((3500 - (3000 * Val(strDisc))) / dblUSRate, "##0.00")) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "' )"
        'Added by Lydia 2014/12/15 ½Ð´Ú³æ½Ð§ï¬°¨Ì¥N²z¤H©Î«È¤áÀÉ³]©wªº½Ð´Ú¹ô§O
'       strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K17,A1K18,A1K19,A1K20,A1K21,A1K25,A1K26,A1K29,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04) " & _
                "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",0,NULL,0," & dblUSRate & "," & 3500 - (3000 * Val(strDisc)) & ",NULL,'" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','','USD'," & Val(ACDate(ServerDate)) & "," & ServerTime & ",'" & strUserNum & "','','','',0," & Fix(Val("" & IIf(dblUSRate = 0, 3500 - (3000 * Val(strDisc)), (3500 - (3000 * Val(strDisc))) / dblUSRate))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "' )"
        strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K17,A1K18,A1K19,A1K20,A1K21,A1K25,A1K26,A1K29,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K33) " & _
                "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",0,NULL,0," & dblUSRate & "," & 3500 - (3000 * Val(strDisc)) & ",NULL,'" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','','" & strA1K18 & "'," & Val(ACDate(ServerDate)) & "," & ServerTime & ",'" & strUserNum & "','','','',0," & Fix(Val("" & IIf(dblUSRate = 0, 3500 - (3000 * Val(strDisc)), (3500 - (3000 * Val(strDisc))) / dblUSRate))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "','" & strA1K33 & "')"
       
        'End
       cnnConnection.Execute strSql
       '3:·s¼W¨âµ§ACC1L0
'       strDisc = 1 - (PUB_GetA1L07Disc(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, strSrvDate(2)) / 100)
       strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
                "VALUES  ('" & m_strSerialNo & "','FCT',''," & 3000 * Val(strDisc) & ",'001','1701',3000," & Val(ACDate(ServerDate)) & "," & ServerTime & ",'" & strUserNum & "' )"
       cnnConnection.Execute strSql
       strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
                "VALUES  ('" & m_strSerialNo & "','FCT','', 0,'002','02',500," & Val(ACDate(ServerDate)) & "," & ServerTime & ",'" & strUserNum & "' )"
       cnnConnection.Execute strSql
       
       PUB_UpdateA1k08 m_strSerialNo 'Added by Morgan 2012/11/2 §ó·s½Ð´Ú³æ¥~¹ôª÷ÃB
       
       '4:·s¼WACC1W0
       strSql = "INSERT INTO ACC1W0 (A1W01,A1W02) " & _
                "VALUES  ('" & m_strSerialNo & "','" & strCP09 & "')"
       cnnConnection.Execute strSql
        'Modify By Cheng 2003/11/27
        '­Y§ó¥¿(302)®Ö­ã¥B²£¥Íµù¥UÃÒ½Ð´Ú¸ê®Æ
        If m_CP10 = "302" Then
           '5:§ó·s§ó¥¿¬ÛÃöÁ`¦¬¤å¸¹(µù¥UÃÒ)
           'Modified by Lydia 2023/09/04 §ï¥ÎÅÜ¼Æ
           'strSql = "UPDATE CASEPROGRESS SET CP60='" & m_strSerialNo & "' WHERE CP09=(Select CP43 From CaseProgress Where CP09='" & m_CP09 & "' )"
           strSql = "UPDATE CASEPROGRESS SET CP60='" & m_strSerialNo & "' WHERE CP09='" & m_CP43 & "'"
        Else
           '5:§ó·s·s¼WªºCÃþ¦¬¤å¸¹
           strSql = "UPDATE CASEPROGRESS SET CP60='" & m_strSerialNo & "' WHERE CP09='" & strCP09 & "'"
        End If
       cnnConnection.Execute strSql
       
        'Moved By Cheng 2004/05/12
'       '6:¦C¦L·s¼Wªº½Ð´Ú¸ê®Æ
'       ProcessPrint
        'End
        
        PUB_PointAutoassign m_strSerialNo, True 'Add by Morgan 2010/4/21 ¦Û°Ê¤À°tÂI¼Æ
    End If
   '2006/6/1 ADD BY SONIA ­ì¥¼¼¶¼g¦¹¬q§ó·s
   ' ­Y®×¥ó©Ê½è¬°²¾Âà®É, §ó·s°Ó¼Ð°ò¥»ÀÉ¤§¨÷©v©Ê½è
   If m_CP10 = "501" Then
      strSql = "UPDATE TradeMark SET TM28 = '1' " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
      cnnConnection.Execute strSql
   End If
   '2006/6/1 END
   
   Dim SeekMonTM01 As String
   Dim SeekMonTM02 As String
   Dim SeekMonTM03 As String
   Dim SeekMonTM04 As String
   'ADD BY nickc 2006/09/27 ­Y¬OBÃþ¥Ó½Ð®×¡A«h¥Nªí¬O¤À³Î²£¥Í¡A­nÀË¬d¤À³Îªº¬ÛÃö¤l®×¬O§_¦³­ã»é¡A­Y¥þ³£¦³¡A«h±N¥À®×¤W³¬¨÷
   If Mid(m_CP09, 1, 1) = "B" And m_CP10 = "101" Then
       Set rsA = New ADODB.Recordset
       If rsA.State = 1 Then rsA.Close
       strSql = "select * from divisioncase where dc01='" & m_TM01 & "' and dc02='" & m_TM02 & "' and dc03='" & m_TM03 & "' and dc04='" & m_TM04 & "' "
       rsA.CursorLocation = adUseClient
       rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If rsA.RecordCount <> 0 Then
            SeekMonTM01 = CheckStr(rsA.Fields("dc05"))
            SeekMonTM02 = CheckStr(rsA.Fields("dc06"))
            SeekMonTM03 = CheckStr(rsA.Fields("dc07"))
            SeekMonTM04 = CheckStr(rsA.Fields("dc08"))
            Set rsA = New ADODB.Recordset
            If rsA.State = 1 Then rsA.Close
            strSql = "select * from divisioncase,trademark where dc05='" & SeekMonTM01 & "' and dc06='" & SeekMonTM02 & "' and dc07='" & SeekMonTM03 & "' and dc08='" & SeekMonTM04 & "' and dc01=tm01(+) and dc02=tm02(+) and dc03=tm03(+) and dc04=tm04(+) and (tm16 is null or tm16='') "
            rsA.CursorLocation = adUseClient
            rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount = 0 Then
                strSql = "update trademark set tm29='Y',tm30=to_number(to_char(sysdate,'YYYYMMDD')),tm31='87' where tm01='" & SeekMonTM01 & "' and tm02='" & SeekMonTM02 & "' and tm03='" & SeekMonTM03 & "' and tm04='" & SeekMonTM04 & "' and (tm29 is null or tm29='') "
                cnnConnection.Execute strSql
            End If
       End If
   End If
   
   'Added by Morgan 2017/5/3 ¹q¤l¤½¤å
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
   End If
   'end 2017/5/3
   
 '911107 nick transation
  cnnConnection.CommitTrans
  
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '­Y³]©w²£¥Í½Ð´Ú¸ê®Æ
    If Me.Text3.Text = "Y" Then
        '6:¦C¦L·s¼Wªº½Ð´Ú¸ê®Æ
        ProcessPrint
        'Added by Lydia 2016/11/17 ¥H½Ð´Ú¹ï¶HÀË¬d¬O§_¦s¦b©ó°ê¥~©T©w±H¶Ê´Ú³æ¥N²z¤HÀÉ(ACC225)¥B¤U¦¸±Hµo¤é´Á¡Ö¨t²Î¤é¡A­Y¦s¦b«hÅã¥Ü°T®§´£¿ô¾Þ§@¤H­û
        If m_strSerialNo <> "" And strA1K28 <> "" Then
           If PUB_ChkAcc225MsgList(m_strSerialNo, strA1K28, m_TM01, m_TM02, m_TM03, m_TM04) Then
           End If
        End If
        'end 2016/11/17
    End If
    
   ' ¦C¦L©w½Z
   If textPrint <> "N" Then
        '2009/4/22 ¥ÑPrintLetter²¾¹L¨Ó
        If Me.Combo1.Text <> "" Then
            arrCP10 = Split(Me.Combo1.Text, " ")
            strCP10Code = arrCP10(0)
            '2009/4/22 modify by sonia §ï§ì¸ÓÂI¿ï¦¬¤å¸¹¤§¨Ó¨ç¬ÛÃöÁ`¦¬¤åªº­ì¬ÛÃöÁ`¦¬¤å¸¹
            'm_strCP09 = m_TM01 & m_TM02 & m_TM03 & m_TM04 & "&" & strCP10Code
            StrSQLa = "Select * From CaseProgress Where CP09 =(SELECT CP43 FROM CASEPROGRESS WHERE CP09=(Select CP43 From CaseProgress Where CP09='" & m_CP09 & "' ))"
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               m_strCP09 = "" & rsA("CP09").Value
            Else
                strCP10Code = ""
                m_strCP09 = ""
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        Else
            strCP10Code = ""
            m_strCP09 = ""
        End If
        '2009/4/22 end
        'Add By Cheng 2003/12/22
        '2009/4/22 modify by sonia
        'Select Case m_CP10
        Select Case IIf(strCP10Code <> "", strCP10Code, m_CP10)
        '2009/4/22 end
        Case "102", "301", "501", "502", "313" '©µ®i, ÅÜ§ó, ²¾Âà, ±ÂÅv 2007/6/7 ¥[´îÁY°Ó«~313
            '2009/8/24 MODIFY BY SONIA ªü½¬»¡¤£¥²¦A°Ý¤F,³£¤£ªþµù¥UÃÒ
            'm_strWithRegister = UCase(InputBox("¬O§_ªþµù¥UÃÒ???" & vbCrLf & vbCrLf & "Y : ªþµù¥UÃÒ(¨Ï¥ÎÂÂ©w½Z¤ÎÄ¶¤å)" & vbCrLf & "N : ¤£ªþµù¥UÃÒ(¨Ï¥Î·s©w½Z¤ÎÄ¶¤å)", , "N"))
            m_strWithRegister = "N"
            '2009/8/24 END
        Case Else
            m_strWithRegister = "Y"
        End Select
        'End
        PrintLetter
        
        'Added by Lydia 2017/06/27 ©µ®i¡B²¾Âà¡BÅÜ§ó(102,501,301)®×¤§®Ö­ã¿é¤J¡A¤Ä¿ï¡u¼È¤£¦C¦L©w½Z¡v®É¡A±N©w½Z¤é´Á§ï¬°99999999
        strExc(1) = IIf(strCP10Code <> "", strCP10Code, m_CP10)
        'Mark by Lydia 2023/08/01 ¨ú®øºÞ¨î: ¦]¬°²{¦bFCT©Ò¦³©w½Z(¶Ê©µ®i°£¥~)¦b²£¥Í©ó©w½Z§@·~ºûÅ@¦P®É¡A·|¥t±N©w½ZÀx¦s©óFCT-workflow
                                            '©Ò¥Hµ{§Ç¤H­û³£¦bFCT -workflow°µ­×§ï©Î¦C¦Lªº°Ê§@, ¤£·|¨C¥ó³£±q©w½Z§@·~ºûÅ@¦C¦L©w½Z¤F
        'If InStr("102,301,501", strExc(1)) > 0 And Chk1.Visible = True And Chk1.Value = True And (ET03 <> "" Or ET03_1 <> "" Or ET03r <> "") Then
        '   '¦]¬°¨Ò¥~Äæ¦ìªºET07¬OTrigger¼g¤J,©Ò¥H¦sÀÉ«áÅÜ§ó©w½Z¤é´Á
        '   cnnConnection.BeginTrans
        '      strExc(2) = ""
        '      If ET03 <> "" Then strExc(2) = strExc(2) & IIf(strExc(2) <> "", ",", "") & CNULL(ET03)
        '      If ET03_1 <> "" Then strExc(2) = strExc(2) & IIf(strExc(2) <> "", ",", "") & CNULL(ET03_1)
        '      If ET03r <> "" Then strExc(2) = strExc(2) & IIf(strExc(2) <> "", ",", "") & CNULL(ET03r)
        '      'Modified by Lydia 2017/08/23 §PÂ_©w½Z®É¶¡¤£¥i­«½Æ
        '      'strSql = "update letterdemand set ld02=99999999 where ld04='" & m_CP09 & "' and ld01='" & strUserNum & "' and ld02=" & strSrvDate(1) & " and ld10='03' and ld11 in (" & strExc(2) & ") "
        '      'cnnConnection.Execute strSql, intI
        '      StrSQLa = "Select ld04,ld01,ld02,ld03,ld10,ld11 From letterdemand where ld04='" & m_CP09 & "' and ld01='" & strUserNum & "' and ld02=" & strSrvDate(1) & " and ld10='03' and ld11 in (" & strExc(2) & ") "
        '      rsA.CursorLocation = adUseClient
        '      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        '      If rsA.RecordCount > 0 Then
        '         rsA.MoveFirst
        '         Do While Not rsA.EOF
        '            strExc(3) = PUB_GetUniqeLD03(strUserNum, "99999999", Format(ServerTime, "000000"))
        '            strSql = "update letterdemand set ld02=99999999,ld03=" & Val(strExc(3)) & " where ld04='" & rsA.Fields("ld04") & "' and ld01='" & rsA.Fields("ld01") & "' and ld02=" & rsA.Fields("ld02") & " and ld03=" & rsA.Fields("ld03") & " and ld10='" & rsA.Fields("ld10") & "' and ld11 ='" & rsA.Fields("ld11") & "' "
        '            cnnConnection.Execute strSql, intI
        '            rsA.MoveNext
        '         Loop
         '     End If
         '     'end 2017/08/23
         '     strSql = "update exceptcondition set et07=99999999 where et02='" & m_CP09 & "' and et04='" & strUserNum & "' and et07=" & strSrvDate(1) & " and et01='03' and et03 in (" & strExc(2) & ") "
         '     cnnConnection.Execute strSql, intI
         '  cnnConnection.CommitTrans
        'End If
        ''end 2017/06/27
        'end 2023/08/01
        
         m_blnPrintAddress = True
   End If

    '911107 nick transation
     Exit Function
CheckingErr:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)
    'edit by nick 2004/11/03
    OnSaveData = False

End Function

Private Sub Form_Unload(Cancel As Integer)
    '­Y¦Lªí¾÷ÅÜ°Ê, «h§ó·s¦C¦L³]©w
    If Me.Combo2.Text <> Me.Combo2.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
    End If
   'Add By Cheng 2002/07/19
   Set frm03020401_04 = Nothing
End Sub

Private Sub Text1_GotFocus()
    TextInverse Me.Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
    If Me.Text1.Text <> "" Then
      '2009/4/27 modify by soniaªü½¬»¡§ï¦^¥Á°ê¦~
      'edit by nickc 2008/03/12 §ï¦è¤¸¦~
      If CheckIsTaiwanDate(Me.Text1.Text) = False Then
      'If CheckIsDate(Text1, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½TªºÃÒ®Ñ¤é´Á"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1_GotFocus
      End If
    End If
    If Cancel = True Then TextInverse Me.Text1
End Sub

Private Sub Text2_GotFocus()
    TextInverse Me.Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 89 Then
        KeyAscii = 0
    End If
    '2009/4/22 add by sonia ¿ï¾Ü§ó¥¿®Ö­ã¨ç®É¤£¦L©w½Z
    If KeyAscii = 89 Then
       textPrint = "N"
    End If
    '2009/4/22 end
End Sub

Private Sub Text3_Change()
    'Add By Cheng 2003/12/02
    If Me.Text3.Text = "Y" Then
        Me.Label31.Visible = True
        Me.Label31.Enabled = True
        Me.Combo2.Visible = True
        Me.Combo2.Enabled = True
    Else
        Me.Label31.Visible = False
        Me.Label31.Enabled = False
        Me.Combo2.Visible = False
        Me.Combo2.Enabled = False
    End If
    'End
End Sub

Private Sub Text3_GotFocus()
    TextInverse Me.Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 89 Then
        KeyAscii = 0
    End If
End Sub

Private Sub textCP08_LostFocus()
On Error GoTo ErrorHandler

'Add By Cheng 2002/01/15
If Len(Me.textCP08.Text) > 0 Then
   m_intNumBegin = InStr(Me.textCP08.Text, m_strNumBegin)
   m_intNumEnd = InStr(Me.textCP08.Text, m_strNumEnd)
Else
   m_intNumBegin = 0
   m_intNumEnd = 0
End If
If m_intNumBegin < m_intNumEnd Then
   Me.textCP35.Text = Mid(Me.textCP08.Text, m_intNumBegin + 1, (m_intNumEnd - m_intNumBegin - 1))
End If

Exit Sub

ErrorHandler:
   m_intNumBegin = 0
   m_intNumEnd = 0
End Sub

'Add By Sindy 2010/11/29
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' ©Ó¿ì¤H
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
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "©Ó¿ì¤H¥N¸¹¤£¦s¦b"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP14_GotFocus
      End If
   End If
End Sub

Private Sub textCP25_GotFocus()
InverseTextBox textCP25
End Sub

Private Sub textCP25_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

If CheckIsDate(Me.textCP25, False) = False Then
   Cancel = True
   strTit = "¸ê®ÆÀË®Ö"
   strMsg = "½Ð¿é¤J¥¿½Tªº¤é´Á"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Me.textCP25.SetFocus
   textCP25_GotFocus
   Exit Sub
End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    'Add By Cheng 2003/09/23
    'Begin
    If KeyAscii <> 8 And KeyAscii <> 78 Then
        KeyAscii = 0
    End If
    'End
End Sub

' ¼f¬d©e­û
Private Sub textCP35_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP35, 32) = False Then
      Cancel = True
      strTit = "¸ê®ÆÀË®Ö"
      strMsg = "¼f¬d©e­û¸ê®Æ¤º®e¤Óªø"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP35_GotFocus
   End If
End Sub

' ©Ó¿ì¤H´Á­­
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   ' ©Ó¿ì´Á­­ªº¤é´ÁÀ³¬°¨Ó¨ç¦¬¤å¤é¥[¤W¤u§@¤Ñ¼Æ
   ' ¤u§@¤Ñ¼Æ¥Ñ¨t²Î§O+°ê®a¥N½X+®×¥ó©Ê½è(®Ö­ã)·j´M®×¥ó¦¬¶Oªíªº¤u§@¤Ñ¼Æ
   ' ­Y¦³­È¤~°µÀË¬d
   If IsEmptyText(textCP48) = False Then
      ' ÀË¬d¬O§_¬°¥Á°ê¤é´Á
      '2009/4/27 modify by soniaªü½¬»¡§ï¦^¥Á°ê¦~
      'edit by nickc 2008/03/12 §ï¦è¤¸¦~
      If CheckIsTaiwanDate(textCP48, False) = False Then
      'If CheckIsDate(textCP48, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½Tªº©Ó¿ì´Á­­"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
      End If
   End If
End Sub

Private Sub textCP53_GotFocus()
InverseTextBox textCP53
End Sub

Private Sub textCP53_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

' ÀË®Ö¬O§_¬°¥Á°ê¤é´Á
'2009/4/27 modify by soniaªü½¬»¡§ï¦^¥Á°ê¦~
'edit by nickc 2008/03/12 §ï¦è¤¸¦~
If CheckIsTaiwanDate(Me.textCP53, False) = False Then
'If CheckIsDate(Me.textCP53, False) = False Then
   Cancel = True
   strTit = "¸ê®ÆÀË®Ö"
   strMsg = "½Ð¿é¤J¥¿½Tªº¤é´Á"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Me.textCP53.SetFocus
   textCP53_GotFocus
   Exit Sub
End If
'edit by nickc 2008/01/10 ±M¥Î¨ä§ï¦è¤¸¦~¡A©Ò¥H­n­×¥¿
'If Val(Me.textCP53.Text) < Val(Me.textTM21.Text) Or Val(Me.textCP53.Text) > Val(Me.textTM22.Text) Then
If Val(DBDATE(Me.textCP53.Text)) < Val(Me.textTM21.Text) Or Val(DBDATE(Me.textCP53.Text)) > Val(Me.textTM22.Text) Then
   Cancel = True
   strTit = "¸ê®ÆÀË®Ö"
   'edit by nickc 2008/01/10
   'strMsg = Replace(Me.Label4(0).Caption, "¡G", "") & "»P±M¥Î´Á¶¡¤£²Å, ¬O§_­«·s¿é¤J???" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "±M¥Î´Á¶¡¡G" & Me.textTM21.Text & "¡Ð" & Me.textTM22.Text & Chr(10) & Chr(13) & Me.Label4(0).Caption & Me.textCP53.Text & "¡Ð" & Me.textCP54.Text
   strMsg = Replace(Me.Label4(0).Caption, "¡G", "") & "»P±M¥Î´Á¶¡¤£²Å, ¬O§_­«·s¿é¤J???" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "±M¥Î´Á¶¡¡G" & TAIWANDATE(Me.textTM21.Text) & "¡Ð" & TAIWANDATE(Me.textTM22.Text) & Chr(10) & Chr(13) & Me.Label4(0).Caption & Me.textCP53.Text & "¡Ð" & Me.textCP54.Text
   nResponse = MsgBox(strMsg, vbYesNo, strTit)
   If nResponse = vbYes Then
      Me.textCP53.SetFocus
      textCP53_GotFocus
      Exit Sub
   End If
   Cancel = False
End If

End Sub

Private Sub textCP54_GotFocus()
InverseTextBox textCP54
End Sub

Private Sub textCP54_lostfocus()
If Me.textCP53.Visible And Me.textCP54.Visible Then
   If Val(Me.textCP53.Text) > Val(Me.textCP54.Text) Then
      MsgBox Replace(Me.Label4(0).Caption, "¡G", "") & "¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
      Me.textCP53.SetFocus
      textCP53_GotFocus
   End If
End If
End Sub

Private Sub textCP54_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

' ÀË®Ö¬O§_¬°¥Á°ê¤é´Á
'2009/4/27 modify by soniaªü½¬»¡§ï¦^¥Á°ê¦~
'edit by nickc 2008/03/12 §ï¦è¤¸¦~
If CheckIsTaiwanDate(Me.textCP54, False) = False Then
'If CheckIsDate(Me.textCP54, False) = False Then
   Cancel = True
   strTit = "¸ê®ÆÀË®Ö"
   strMsg = "½Ð¿é¤J¥¿½Tªº¤é´Á"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Me.textCP54.SetFocus
   textCP54_GotFocus
   Exit Sub
End If
'edit by nickc 2008/01/10 ±M¥Î´Á§ï¦è¤¸¦~¡A¬G­×¥¿
'If Val(Me.textCP54.Text) < Val(Me.textTM21.Text) Or Val(Me.textCP54.Text) > Val(Me.textTM22.Text) Then
If Val(DBDATE(Me.textCP54.Text)) < Val(Me.textTM21.Text) Or Val(DBDATE(Me.textCP54.Text)) > Val(Me.textTM22.Text) Then
   Cancel = True
   strTit = "¸ê®ÆÀË®Ö"
   'edit by nickc 2008/01/10
   'strMsg = Replace(Me.Label4(0).Caption, "¡G", "") & "»P±M¥Î´Á¶¡¤£²Å, ¬O§_­«·s¿é¤J???" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "±M¥Î´Á¶¡¡G" & Me.textTM21.Text & "¡Ð" & Me.textTM22.Text & Chr(10) & Chr(13) & Me.Label4(0).Caption & Me.textCP53.Text & "¡Ð" & Me.textCP54.Text
   strMsg = Replace(Me.Label4(0).Caption, "¡G", "") & "»P±M¥Î´Á¶¡¤£²Å, ¬O§_­«·s¿é¤J???" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "±M¥Î´Á¶¡¡G" & TAIWANDATE(Me.textTM21.Text) & "¡Ð" & TAIWANDATE(Me.textTM22.Text) & Chr(10) & Chr(13) & Me.Label4(0).Caption & Me.textCP53.Text & "¡Ð" & Me.textCP54.Text
   nResponse = MsgBox(strMsg, vbYesNo, strTit)
   If nResponse = vbYes Then
      Me.textCP54.SetFocus
      textCP54_GotFocus
      Exit Sub
   End If
   Cancel = False
End If

End Sub

Private Sub textDN_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    'Add By Cheng 2003/09/23
    'Begin
    If KeyAscii <> 8 And KeyAscii <> 89 Then
        KeyAscii = 0
    End If
    'End
End Sub

Private Sub textMod_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    'Add By Cheng 2003/09/23
    'Begin
    If KeyAscii <> 8 And KeyAscii <> 89 Then
        KeyAscii = 0
    End If
    'End
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    'Add By Cheng 2003/09/23
    'Begin
    If KeyAscii <> 8 And KeyAscii <> 78 Then
        KeyAscii = 0
    End If
    'End
End Sub

Private Sub textPrtTrans_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    'Add By Cheng 2003/09/23
    'Begin
    If KeyAscii <> 8 And KeyAscii <> 78 Then
        KeyAscii = 0
    End If
    'End
End Sub

' ¦C¦L³Æµù
Private Sub textPS_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If CheckLengthIsOK(textPS, 2000) = False Then
      Cancel = True
      strTit = "¸ê®ÆÀË®Ö"
      strMsg = "¦C¦L³Æµù¸ê®Æ¤º®eªø«×¤Óªø"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textPS_GotFocus
   End If
End Sub

Private Sub textTM14_Change()
m_strLastTextTM14 = Me.textTM14.Text
End Sub

' ¤½§i¤é
Private Sub textTM14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM14) = False Then
      ' ÀË¬d¬O§_¬°¥Á°ê¦~
      '2009/4/27 modify by soniaªü½¬»¡§ï¦^¥Á°ê¦~
      'edit by nickc 2008/03/12 §ï¦è¤¸¦~
      If CheckIsTaiwanDate(textTM14, False) = False Then
      'If CheckIsDate(textTM14, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½Tªº¤½§i¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14_GotFocus
      End If
      ' ¤½§i¤é¤£¥i¶W¹L¨t²Î¤é
      'If Val(DBDATE(textTM14)) > Val(DBDATE(SystemDate())) Then
      '   Cancel = True
      '   strTit = "¸ê®ÆÀË®Ö"
      '   strMsg = "¤½§i¤é¤£¥i¶W¹L¨t²Î¤é"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textTM14_GotFocus
      'End If
   End If
End Sub

' ¼f©w¸¹¼Æ
Private Sub textTM15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   Cancel = False
            
   If IsEmptyText(textTM15) = False Then
      'Add By Sindy 2010/9/1
      'ÀË¬d¼f©w¸¹©Ò¿é¤Jªºªø«×¬O§_¥¿½T
      If bolNewAppNoFormat Then
         'Add By Sindy 2017/5/17 + strRetrunText
         If PUB_ChkTm12Tm15Length("2", textTM15, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10, , , strRetrunText) = False Then
            Cancel = True
            textTM15_GotFocus
            Exit Sub
         'Add By Sindy 2017/5/17
         Else
            textTM15 = strRetrunText
         '2017/5/17 END
         End If
      Else
         If IsNumeric(Mid(textTM15, 1, 8)) = False Then
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "½Ð¿é¤J¥¿½Tªº¼f©w¸¹¼Æ"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM15_GotFocus
         End If
      End If
   End If
End Sub

Private Sub textTM16S_Change()
'Modify By Cheng 2002/07/22
'm_strLastTextTM16S = Me.textTM16S.Text
End Sub

Private Sub textTM16S_KeyPress(KeyAscii As Integer)
   'Modify By Cheng 2002/07/22
'   KeyAscii = UpperCase(KeyAscii)
End Sub

' ±M¥Î´Á­­°_¤é
Private Sub textTM21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strCorrDate As String
   Dim strDate As String
   
   Cancel = False
   ' ­ì±M¥Î´Á­­¤î¤é
   If IsEmptyText(m_TM22) = True Then
      GoTo EXITSUB
   End If
   ' ¥¼¿é¤J±M¥Î´Á­­°_¤é
   If IsEmptyText(textTM21) = True Then
      GoTo EXITSUB
   End If
   ' ®×¥ó©Ê½è«D©µ®i
   If m_CP10 <> "102" Then
      GoTo EXITSUB
   End If
   
   ' ÀË®Ö¬O§_¬°¥Á°ê¤é´Á
   'edit by nickc 2007/11/30 ªü½¬»¡§ï¦¨¸òÃÒ®Ñ¤W¤@¼Ë¦è¤¸¦~
   'If CheckIsTaiwanDate(textTM21, False) = False Then
   If CheckIsDate(textTM21, False) = False Then
      Cancel = True
      strTit = "¸ê®ÆÀË®Ö"
      strMsg = "½Ð¿é¤J¥¿½Tªº¤é´Á"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM21_GotFocus
   End If
    'Modify By Cheng 2003/09/02
'   strCorrDate = ChangeWDateStringToWString(Format(DateSerial(Val(Left(m_TM22, 4)), Val(Mid(m_TM22, 5, 2)), Right(m_TM22, 2) + 1)))
   strCorrDate = ChangeWDateStringToWString(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22))))
   strDate = textTM21
    'Modify By Cheng 2003/09/02
'   strDate = ChangeWDateStringToWString(Format(DateSerial(Val(Left(strDate, 4)), Val(Mid(strDate, 5, 2)), Right(strDate, 2) + 1)))
   strDate = ChangeWDateStringToWString(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(strDate))))
   If Val(DBDATE(textTM21)) <> Val(DBDATE(m_TM21)) Then
      Cancel = True
      strTit = "¸ê®ÆÀË®Ö"
      strMsg = "±M¥Î´Á­­°_¤é¥²¶·¬°­ì±M¥Î´Á­­°_¤é"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM21_GotFocus
   End If
   
EXITSUB:
End Sub

' ±M¥Î´Á­­¤î¤é
Private Sub textTM22_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strCorrDate As String
   Dim strDate As String
   Cancel = False
   
   ' ­ì±M¥Î´Á­­¤î¤é
   If IsEmptyText(m_TM22) = True Then
      GoTo EXITSUB
   End If
   ' ¥¼¿é¤J±M¥Î´Á­­°_¤é
   If IsEmptyText(textTM22) = True Then
      GoTo EXITSUB
   End If
   ' ®×¥ó©Ê½è«D©µ®i
   If m_CP10 <> "102" Then
      GoTo EXITSUB
   End If
   
   ' ÀË®Ö¬O§_¬°¥Á°ê¤é´Á
   'edit by nickc 2007/11/30 ªü½¬»¡§ï¦¨¸òÃÒ®Ñ¤W¤@¼Ë¦è¤¸¦~
   'If CheckIsTaiwanDate(textTM22, False) = False Then
   If CheckIsDate(textTM22, False) = False Then
      Cancel = True
      strTit = "¸ê®ÆÀË®Ö"
      strMsg = "½Ð¿é¤J¥¿½Tªº¤é´Á"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM22_GotFocus
      GoTo EXITSUB
   End If
   
   strDate = DBDATE(textTM22)
   
   Select Case m_TM08
      Case "1", "4", "7", "8":
            'Modify By Cheng 2003/09/02
'         strCorrDate = ChangeWDateStringToWString(Format(DateSerial(Val(Left(m_TM22, 4)) + Val(m_NA14), Val(Mid(m_TM22, 5, 2)), Right(m_TM22, 2))))
          'Modified by Lydia 2019/11/13  §ï¥Î¦@¥Î¼Ò²ÕÀË¬d2/29, ¨Ã¥B¦]À³°Ó¼Ð®×ªººâªk,¤£§ìNA85ª½±µ³]¡u­pºâ°Ó¼Ð±M¥Î´Á¬O§_´î1¤Ñ¡v=N
         'strCorrDate = ChangeWDateStringToWString(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
         'Modify By Sindy 2022/3/7 + m_TM10 : ©µ®i«á¤§±M¥Î´Á­­¦~«×­Õ¦³2¤ë29¤é®É¡A±M¥Î´Á­­¤î¤éÀ³¬°2¤ë29¤é¡A¦Ó«D¥H¥[10¦~¤§¤è¦¡­pºâ¬°2¤ë28¤é
         strCorrDate = PUB_GetEndDate(DBDATE(m_TM22), Val(m_NA14), "N", m_TM10)
      Case Else:
         strCorrDate = textTM22S
   End Select
   '91.12.8 MODIFY BY SONIA
   'If Val(strDate) <> Val(strCorrDate) Then
   If Val(DBDATE(strDate)) <> Val(DBDATE(strCorrDate)) Then
   '91.12.8 END
      Cancel = True
      strTit = "¸ê®ÆÀË®Ö"
      strMsg = "±M¥Î´Á­­¤î¤é¤£¥¿½T"
        'Modify By Cheng 2002/12/23
        '«ö½T©w¤´¥iÄ~Äò§@·~
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      nResponse = MsgBox(strMsg, vbOKCancel, strTit)
      If nResponse = vbCancel Then
        textTM22_GotFocus
      Else
        Cancel = False
      End If
   End If
EXITSUB:
End Sub

Private Function CheckDataValid() As Boolean
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   CheckDataValid = False
   
   ' 90.08.29 modify (¤£»ÝÀË¬d¼f©w¸¹Äæ¦ì¬O§_¿é¤J)
   ' ¼f©w¸¹¼Æ¤£¥i¬°ªÅ¥Õ
   'If IsEmptyText(textTM15) = True Then
   '   strTit = "¸ê®ÆÀË®Ö"
   '   strMsg = "¼f©w¸¹¤£¥i¬°ªÅ¥Õ"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   textTM15.SetFocus
   '   GoTo EXITSUB
   'End If
   ' ®Ö­ã³qª¾¤é¤£¥i¬°ªÅ¥Õ
   '91.4.29 MODIFY BY SONIA ¨ú®ø
   'If IsEmptyText(textCP25) = True Then
   '   strTit = "¸ê®ÆÀË®Ö"
   '   strMsg = "®Ö­ã³qª¾¤é¤£¥i¬°ªÅ¥Õ"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   textCP25.SetFocus
   '   GoTo EXITSUB
   'End If
    'Add By Cheng 2002/12/11
    '­Y®×¥ó©Ê½è¬°ÅÜ§ó(301)
'edit by nickc 2005/08/04
'    If m_CP10 = "301" Then
        'Modified by Lydia 2016/07/19 +§PÂ_
        'If m_blnClkChgButton = False Then
        If m_blnClkChgButton = False And Me.cmdMod.Visible = True Then
            MsgBox "½Ð¿é¤JÅÜ§ó¨Æ¶µ!!!", vbExclamation + vbOKOnly
            Me.cmdMod.SetFocus
            GoTo EXITSUB
        End If
'    End If
   '93.7.2 cancel by sonia
   '' ¤½§i¤é
   'If IsEmptyText(textTM14) = True Then
   '   strTit = "¸ê®ÆÀË®Ö"
   '   strMsg = "¤½§i¤é¤£¥i¬°ªÅ¥Õ"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   textTM14.SetFocus
   '   GoTo EXITSUB
   'End If
   '93.7.2 end
   ' ±M¥Î´Á­­¤Î¤½§i¤é
   If m_CP10 = "102" Then
      'add by sonia 2017/8/14
      If IsEmptyText(textTM14) = True Then
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "®×¥ó©Ê½è¬°©µ®i, ¤½§i¤é¤£¥i¬°ªÅ¥Õ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14.SetFocus
         GoTo EXITSUB
      End If
      'end 2017/8/14
      If IsEmptyText(textTM21) = True Or IsEmptyText(textTM22) = True Then
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "®×¥ó©Ê½è¬°©µ®i, ±M¥Î´Á­­¤£¥i¬°ªÅ¥Õ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM21.SetFocus
         GoTo EXITSUB
      End If
      If Val(textTM21) > Val(textTM22) Then
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "±M¥Î´Á­­ªº°_¤é¤£¥i¶W¹L¨´¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM21.SetFocus
         GoTo EXITSUB
      End If
   End If
   'Modify By Cheng 2002/07/22
'   ' ±M¥ÎÅv¬O§_¦s¦b
'   If textCP10 <> "101" And IsEmptyText(textTM17) = True Then
'      strTit = "¸ê®ÆÀË®Ö"
'      strMsg = "±M¥ÎÅv¬O§_¦s¦b¤£¥i¬°ªÅ¥Õ"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM17.SetFocus
'      GoTo EXITSUB
'   End If
   'Modify By Cheng 2002/07/22
'   ' ¬O§_§ó·s°ò¥»ÀÉ¥Ø«e­ã»é
'   If IsEmptyText(textTM16S) = True Then
'      strTit = "¸ê®ÆÀË®Ö"
'      strMsg = "¬O§_§ó·s°ò¥»ÀÉ¥Ø«e­ã»é¤£¥i¬°ªÅ¥Õ"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM16S.SetFocus
'      GoTo EXITSUB
'   End If
   
   ' ©Ó¿ì´Á­­¦³©w¸q¤u§@¤Ñ¼Æ®É¤£¥i¬°ªÅ¥Õ
   If IsEmpty(textCP48) = True Then
      Set rsTmp = New ADODB.Recordset
      ' ©Ó¿ì´Á­­ªº¤é´ÁÀ³¬°¨Ó¨ç¦¬¤å¤é¥[¤W¤u§@¤Ñ¼Æ
      ' ¤u§@¤Ñ¼Æ¥Ñ¨t²Î§O+°ê®a¥N½X+®×¥ó©Ê½è(®Ö­ã)·j´M®×¥ó¦¬¶Oªíªº¤u§@¤Ñ¼Æ
      ' ­Y¦³­È¤~°µÀË¬d
      strSql = "SELECT * FROM CaseFee " & _
               "WHERE CF01 = '" & m_TM01 & "' AND " & _
                     "CF02 = '" & m_TM10 & "' AND " & _
                     "CF03 = '1001' AND " & _
                     "CF04 <> NULL "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
      If rsTmp.RecordCount > 0 Then
         rsTmp.Close
         Set rsTmp = Nothing
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "©Ó¿ì´Á­­¦³©w¸q¤u§@¤Ñ¼Æ®É¤£¥i¬°ªÅ¥Õ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.SetFocus
         GoTo EXITSUB
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
   'Add By Cheng 2003/11/27
    '­Y¬°§ó¥¿(302)®Ö­ã¥B²£¥Íµù¥UÃÒ½Ð´Ú¸ê®Æ
    If m_CP10 = "302" And Me.Text3.Text = "Y" Then
        StrSQLa = "Select * From CaseProgress Where CP09 =(Select CP43 From CaseProgress Where CP09='" & m_CP09 & "' )"
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            If "" & rsA("CP60").Value <> "" Then
                strTit = "¸ê®ÆÀË®Ö"
                strMsg = "¦¹µ§¬ÛÃöªºµù¥UÃÒ¸ê®Æ¤w½Ð´Ú!!!"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                Me.Text3.SetFocus
                Text3_GotFocus
                GoTo EXITSUB
            End If
        Else
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¦¹µ§µL¬ÛÃöµù¥UÃÒ¸ê®Æ!!!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Me.Text3.SetFocus
            Text3_GotFocus
            GoTo EXITSUB
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
    'End
   
   'Add By Sindy 2014/9/9
   If m_CP10 = "103" Then '¸Éµoµù¥UÃÒ
      If IsEmptyText(Text1) = True Then
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "®×¥ó©Ê½è¬°¸Éµoµù¥UÃÒ, ÃÒ®Ñ¤é´Á¤£¥i¬°ªÅ¥Õ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1.SetFocus
         GoTo EXITSUB
      End If
   End If
   '2014/9/9 END
      
   'Added by Lydia 2017/09/19
   'Modified by Morgan 2022/6/17 ®Ö­ã«áªºÅÜ§ó¤~­n¿é(²¾Âà¥u·|¬O­ã«á)--ªü½¬
   'If (m_CP10 = "501" Or m_CP10 = "301") And frm03020401_03.GetSelectResult() = "1" Then
   If textTM16S = "1" And (m_CP10 = "501" Or m_CP10 = "301") And frm03020401_03.GetSelectResult() = "1" Then
   'end 2022/6/17
   
      If IsEmptyText(textTM14) = True Then
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "®×¥ó©Ê½è¬°" & Trim(textCP10) & ", ¤½§i¤é¤£¥i¬°ªÅ¥Õ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14.SetFocus
         GoTo EXITSUB
      End If
   End If
   'end 2017/09/19
   
    'Added by Lydia 2021/09/13 ÀË¬dµe­±ªº TextBox, ComboBox ¬O§_§t¦³Unicode¤å¦r
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         GoTo EXITSUB
    End If

   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPS_GotFocus()
   InverseTextBox textPS
End Sub

Private Sub textPrint_GotFocus()
    InverseTextBox textPrint
End Sub

Private Sub textDN_GotFocus()
    InverseTextBox textDN
End Sub

Private Sub textPrtTrans_GotFocus()
    InverseTextBox textPrtTrans
End Sub

Private Sub textMod_GotFocus()
   InverseTextBox textMod
End Sub

Private Sub textTMBM07_1_Change()
m_strLastTextTMBM07_1 = Me.textTMBM07_1.Text
End Sub

Private Sub textTMBM07_1_GotFocus()
   InverseTextBox textTMBM07_1
End Sub

Private Sub textTMBM07_2_Change()
m_strLastTextTMBM07_2 = Me.textTMBM07_2.Text
End Sub

Private Sub textTMBM07_2_GotFocus()
   InverseTextBox textTMBM07_2
End Sub

Private Sub textTM14_GotFocus()
   InverseTextBox textTM14
End Sub

Private Sub textTM15_GotFocus()
   InverseTextBox textTM15
End Sub

'Add By Sindy 2010/01/05
Private Sub textTM67_GotFocus()
   InverseTextBox textTM67
End Sub

Private Sub textTM21_GotFocus()
   InverseTextBox textTM21
End Sub

Private Sub textTM22_GotFocus()
   InverseTextBox textTM22
End Sub

Private Sub textCP08_GotFocus()
   'Modify By Cheng 2002/04/22
   '±N´å¼Ð°±¦b"¦r"ªº«e­±
'   InverseTextBox textCP08
Dim intPos As Integer
With Me.textCP08
   If Len("" & .Text) > 0 Then
      intPos = InStr("" & .Text, "¦r")
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

Private Sub textCP35_GotFocus()
   InverseTextBox textCP35
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

' ¨ú±o«È¤áÀÉªº­^¤å¦WºÙ(¤¤¶¡¥HªÅ¥Õ°µ¬°¶¡¹j)
Private Function GetCustomerEngName(ByVal strCU01 As String) As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   GetCustomerEngName = Empty
   strTemp = Empty
   
   strSql = "SELECT * FROM CUSTOMER " & _
            "WHERE CU01 = '" & strCU01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CU05")) = False Then
         strTemp = rsTmp.Fields("CU05")
      End If
      If IsNull(rsTmp.Fields("CU88")) = False Then
         If IsEmptyText(strTemp) = False Then: strTemp = strTemp & " "
         strTemp = strTemp & rsTmp.Fields("CU88")
      End If
      If IsNull(rsTmp.Fields("CU89")) = False Then
         If IsEmptyText(strTemp) = False Then: strTemp = strTemp & " "
         strTemp = strTemp & rsTmp.Fields("CU89")
      End If
      If IsNull(rsTmp.Fields("CU90")) = False Then
         If IsEmptyText(strTemp) = False Then: strTemp = strTemp & " "
         strTemp = strTemp & rsTmp.Fields("CU90")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   GetCustomerEngName = strTemp
End Function

' ÀË¬dÅÜ§ó¨Æ¶µÀÉªº¥Ó½Ð¤H¬O§_®Ö­ã
Private Function IsCE09Approve(ByVal strCE01 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   IsCE09Approve = False
   
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & strCE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CE09")) = False Then
         If rsTmp.Fields("CE09") = "1" Then
            IsCE09Approve = True
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' ¨ú±o·s¥Ó½Ð¤H
Private Function GetNewTM23() As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   GetNewTM23 = Empty
   
   strSql = "SELECT * FROM Trademark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("TM23")) = False Then
         GetNewTM23 = rsTmp.Fields("TM23")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' ¦C¦L©w½Z«e±N¨Ò¥~Äæ¦ì¥[¤J¨ì¦C¦L©w½Z¨Ò¥~Äæ¦ìÀÉ®×¤¤
Private Sub InsExpField()
Dim strSql As String
Dim strTemp As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strChgEvent As String
Dim intFee As Long 'Add By Sindy 2010/8/25
Dim intTotFee As Long 'Add By Sindy 2014/8/22
Dim strText13 As String, strText14 As String 'Add By Sindy 2014/3/31
Dim bolEType0513 As Boolean 'Add By Sindy 2015/8/3
Dim intRow As Integer, intCnt As Integer
Dim strTemp09 As String, strTemp38 As String
Dim strET03 As String 'Add By Sindy 2023/7/19
Dim strDisc As String '§é¦©
   
   bolEType0513 = False 'Add By Sindy 2015/8/3
   'Add  By Cheng 2003/01/23
   '§PÂ_¬O§_¦³Àu¥ýÅv¸ê®Æ
   StrSQLa = "Select Count(*) From PriDate Where PD01='" & m_TM01 & "' And PD02='" & m_TM02 & "' And PD03='" & m_TM03 & "' And PD04='" & m_TM04 & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.Fields(0).Value > 0 Then
       m_blnPriDate = True
   Else
       m_blnPriDate = False
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   'add by nickc 2006/10/25 ¥[¤J¶O¥Î¨ÌÃþ§O¼ÆÅÜ°Ê
   Dim tmpVarTm09 As Variant
   Dim tmpTm09Cnt As Integer
   Dim tmpTm09CntS As Variant
   tmpVarTm09 = Split(textTM09, ",")
   tmpTm09CntS = 0
   For tmpTm09Cnt = 0 To UBound(tmpVarTm09)
       If Trim(tmpVarTm09(tmpTm09Cnt)) <> "" Then
           tmpTm09CntS = tmpTm09CntS + 1
       End If
   Next tmpTm09Cnt
   
   ' ®×¥ó©Ê½è
   Select Case IIf(strCP10Code <> "", strCP10Code, m_CP10)
        'Modify By Cheng 2003/12/16
        '¥Ó½Ð®Ö­ã©w½Z§ï¦b¦¹¥X, ­ì¦bFC¤½§i³qª¾¨ç¥X
      'edit by nick 2004/12/23 ¤À³Î»P¥Ó½Ð°µ¬Û¦Pªº¨Æ±¡
      'Case "101": ' ¥Ó½Ð
      Case "101", "308": ' ¥Ó½Ð
         
         If IIf(strCP10Code <> "", strCP10Code, m_CP10) = "308" Then 'Add By Sindy 2011/8/10 ªü½¬¥u¯d308­n¥X¶Ç¯u«Ê­±©w½Z
            EndLetter "03", m_CP09, "98", strUserNum
            'Add By Sindy 2010/01/14 FCTµù¥U¶O¦Û°Ê¥NÃº
            If m_TM122 = "Y" Then
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & m_CP09 & "','" & "98" & "','" & strUserNum & _
                        "','¶Ç¯u­¶¼Æ','2')"
               cnnConnection.Execute strSql
            Else
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & m_CP09 & "','" & "98" & "','" & strUserNum & _
                        "','¶Ç¯u­¶¼Æ','4')"
               cnnConnection.Execute strSql
            End If
            '2010/01/14 End
         End If
         
         'Add By Sindy 2025/3/5 ¥Ó½Ð®×®Ö­ã®É,§ìÃºµù¥U¶Oªº§é¦©
         If IIf(strCP10Code <> "", strCP10Code, m_CP10) = "101" Then
            strDisc = PUB_GetA1L07Disc(m_TM01, m_TM02, m_TM03, m_TM04, "717", strSrvDate(2))
            If strDisc = 100 Then strDisc = ""
         End If
         
         ' ©w½Z»y¤å
         Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
            ' ¤¤¤å
            Case "1":
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "03", m_CP09, "01", strUserNum
               ' ¨÷¼Æ
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                        "','¨÷¼Æ','" & textTMBM07_1 & "')"
               cnnConnection.Execute strSql
               ' ´Á¼Æ
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                        "','´Á¼Æ','" & textTMBM07_2 & "')"
               cnnConnection.Execute strSql
               ' ¦C¦L³Æµù
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                        "','¦C¦L³Æµù','" & ChgSQL(textPS) & "')"
               cnnConnection.Execute strSql
            ' ­^¤å
            Case "2":
'2014/12/9 CANCEL BY SONIA
'                '­Y¥Ó½Ð¤é¤p©ó921128
'                If Val(m_TM11) < 20031128 Then
'                     ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                     EndLetter "03", m_CP09, "99", strUserNum
'                     'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                     If bolEmail = True And bolPlusPaper = False Then
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "99" & "','" & strUserNum & _
'                                 "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of an Official Notice of Acceptance.')"
'                        cnnConnection.Execute strSql
'                     Else '¶l¥ó
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "99" & "','" & strUserNum & _
'                                 "','¨Ò¥~¤º¤å','A copy of an Official Notice of Acceptance will be mailed to you with the confirmation copy of this letter for your records.')"
'                        cnnConnection.Execute strSql
'                     End If
'                     '2012/11/27 End
'                '­Y¥Ó½Ð¤é¤j©óµ¥©ó921128
'                Else
'2014/12/9 END
                   '2008/11/13 ADD BY SONIA FCTµù¥U¶O¦Û°Ê¥NÃº
                   If m_TM122 = "Y" Then
                     'Modify By Sindy 2010/01/05
                     If Trim(m_TM67) = "" And Trim(textTM67) <> "" Then
                        EndLetter "03", m_CP09, "14", strUserNum
                        '2014/12/9 MODIFY BY SONIA §ï³qª¾ªk©w´Á­­
                        'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & _
                                 "','¥»©Ò´Á­­','" & m_CP06 & "')"
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & _
                                 "','ªk©w´Á­­','" & m_CP07 & "')"
                        '2014/12/9 END
                        cnnConnection.Execute strSql
                        
                        'Modify By Sindy 2022/6/13 Mark
'                        'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                        If bolEmail = True And bolPlusPaper = False Then
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & _
'                                    "','¨Ò¥~¤º¤å','A scanned copy of the Official Notice of Approval is attached for your records.')"
'                           cnnConnection.Execute strSql
'                        Else '¶l¥ó
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & _
'                                    "','¨Ò¥~¤º¤å','A copy of the Official Notice of Approval will be mailed to you with the confirmation copy of this letter.')"
'                           cnnConnection.Execute strSql
'                        End If
'                        '2012/11/27 End
                     Else
                        'Modify By Sindy 2024/8/2
                        If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, "101", ET03, , "03") = True Then
                           EndLetter "03", m_CP09, ET03, strUserNum
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
                                    "','ªk©w´Á­­','" & m_CP07 & "')"
                           cnnConnection.Execute strSql
                        Else
                           ET03 = "10" 'Add By Sindy 2024/8/7
                        '2024/8/2 END
                           EndLetter "03", m_CP09, ET03, strUserNum
                           '2014/12/9 MODIFY BY SONIA §ï³qª¾ªk©w´Á­­
                           'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & _
                                    "','¥»©Ò´Á­­','" & m_CP06 & "')"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
                                    "','ªk©w´Á­­','" & m_CP07 & "')"
                           '2014/12/9 END
                           cnnConnection.Execute strSql
                        End If
                        'Modify By Sindy 2022/6/13 Mark
'                        'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                        If bolEmail = True And bolPlusPaper = False Then
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & _
'                                    "','¨Ò¥~¤º¤å','A scanned copy of the Official Notice of Approval is attached for your records.')"
'                           cnnConnection.Execute strSql
'                        Else '¶l¥ó
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & _
'                                    "','¨Ò¥~¤º¤å','A copy of the Official Notice of Approval will be mailed to you with the confirmation copy of this letter.')"
'                           cnnConnection.Execute strSql
'                        End If
'                        '2012/11/27 End
                     End If
                   Else
                   '2008/11/13 END
                      'Modify By Sindy 2010/01/05
                      If Trim(m_TM67) = "" And Trim(textTM67) <> "" Then
                           strET03 = "13"
                           ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                           EndLetter "03", m_CP09, strET03, strUserNum
                           '2014/12/9 MODIFY BY SONIA §ï³qª¾ªk©w´Á­­
                           'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
                                    "','¥»©Ò´Á­­','" & m_CP06 & "')"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','ªk©w´Á­­','" & m_CP07 & "')"
                           '2014/12/9 END
                           cnnConnection.Execute strSql
                           
                           'Modify By Sindy 2022/6/13 Mark
'                           'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                           If bolEmail = True And bolPlusPaper = False Then
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å','A scanned copy of the Official Notice of Approval is attached for your records.')"
'                              cnnConnection.Execute strSql
'                              'Remove by Lydia 2018/03/22 ¨ú®ø
'                              'Mark by Lydia 2018/03/28 ¤À³Î¥ý¤£§ï
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å2','Enclosed herewith please also find a return sheet for payment of registration fee for your use.')"
'                              cnnConnection.Execute strSql
'                           Else '¶l¥ó
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å','A copy of the Official Notice of Approval will be mailed to you with the confirmation copy of this letter.')"
'                              cnnConnection.Execute strSql
'                              'Remove by Lydia 2018/03/22 ¨ú®ø
'                              'Mark by Lydia 2018/03/28 ¤À³Î¥ý¤£§ï
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å2','Enclosed herewith please find a return sheet for payment of registration fee for your use.')"
'                              cnnConnection.Execute strSql
'                           End If
'                           '2012/11/27 End
                      Else
                        'Modify By Sindy 2012/6/27 °Ó¼Ð­×ªk
'2014/12/9 CANCEL BY SONIA
'                        If Val(DBDATE(m_CP05)) >= 20120701 Then
                           strET03 = "17"
                           ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                           EndLetter "03", m_CP09, strET03, strUserNum
                           '2014/12/9 MODIFY BY SONIA §ï³qª¾ªk©w´Á­­
                           'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & _
                                    "','¥»©Ò´Á­­','" & m_CP06 & "')"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','ªk©w´Á­­','" & m_CP07 & "')"
                           '2014/12/9 END
                           cnnConnection.Execute strSql
                           
                           'Modify By Sindy 2022/6/13 Mark
'                           'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                           If bolEmail = True And bolPlusPaper = False Then
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å','A scanned copy of the Official Notice of Approval is attached for your records.')"
'                              cnnConnection.Execute strSql
'                              'Remove by Lydia 2018/03/22 ¨ú®ø
'                              'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å2','Enclosed herewith please also find a return sheet for payment of registration fee for your use.')"
'                              'cnnConnection.Execute strSql
'                           Else '¶l¥ó
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å','A copy of the Official Notice of Approval will be mailed to you with the confirmation copy of this letter.')"
'                              cnnConnection.Execute strSql
'                              'Remove by Lydia 2018/03/22 ¨ú®ø
'                              'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å2','Enclosed herewith please find a return sheet for payment of registration fee for your use.')"
'                              'cnnConnection.Execute strSql
'                           End If
'                           '2012/11/27 End
                           
'2014/12/9 CANCEL BY SONIA
'                        Else
'                        '2012/6/27 End
'                           ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                           EndLetter "03", m_CP09, "06", strUserNum
'                             'edit by nick 2004/10/28 §ï¦¨¥Î¥»©Ò´Á­­
'         '                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         '                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
'         '                             "','¨ä¥L¤½§i¤é','" & DBDATE(m_CP05) & "')"
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
'                                    "','¥»©Ò´Á­­','" & m_CP06 & "')"
'                           cnnConnection.Execute strSql
'                           'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                           If bolEmail = True And bolPlusPaper = False Then
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å','A scanned copy of the Official Notice of Approval is attached for your records.')"
'                              cnnConnection.Execute strSql
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å2','Enclosed herewith please also find a return sheet for payment of registration fee for your use.')"
'                              cnnConnection.Execute strSql
'                           Else '¶l¥ó
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å','A copy of the Official Notice of Approval will be mailed to you with the confirmation copy of this letter.')"
'                              cnnConnection.Execute strSql
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å2','Enclosed herewith please find a return sheet for payment of registration fee for your use.')"
'                              cnnConnection.Execute strSql
'                           End If
'                           '2012/11/27 End
'                        End If
'2014/12/9 END
                      End If
                        'Add By Sindy 2025/3/6
                        If strDisc <> "" Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "','§é¦©',' x " & strDisc & "¢H')"
                           cnnConnection.Execute strSql
                        End If
                        '2025/3/6 END
                        'Added by Lydia 2018/03/22 ¬O§_¬°¤@®×¤@Ãþ§O
                        If tmpTm09CntS = 1 Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "','¤@®×¤@Ãþ§O','¡ð')"
                           cnnConnection.Execute strSql
                        'Modify By Sindy 2025/3/6
                           '¥[¤@Ãþ§O:¤£Åã¥Ü¤º®e
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "','¥[¤@Ãþ§O¦³§é¦©','¡ð')"
                           cnnConnection.Execute strSql
                        ElseIf tmpTm09CntS > 1 Then
                           If strDisc <> "" Then
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "','¥[¤@Ãþ§O¦³§é¦©','¡@¡@¡@¡@¡@¡@¡@¡@¡@ NT$3,000" & IIf(strDisc <> "", " x " & strDisc & "¢H", "") & " for each additional class')"
                              cnnConnection.Execute strSql
                           End If
                        '2025/3/6 END
                        End If
                        'end 2018/03/22
                        
                      'Modify By Sindy 2012/6/27 °Ó¼Ð­×ªk
'2014/12/9 CANCEL BY SONIA
'                      If Val(DBDATE(m_CP05)) >= 20120701 Then
 'Remove by Lydia 2018/03/22 ¨ú®ø¦^ÂÐ³æ
 'Mark by Lydia 2018/03/28 ¤À³Î¥ý¤£§ï
                       If IIf(strCP10Code <> "", strCP10Code, m_CP10) = "308" Then 'Added by Lydia 2018/03/28 ¤À³Î¥ý¤£§ï
                           EndLetter "03", m_CP09, "18", strUserNum
                           '2014/12/9 MODIFY BY SONIA §ï³qª¾ªk©w´Á­­
                           'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "18" & "','" & strUserNum & _
                                    "','¥»©Ò´Á­­','" & m_CP06 & "')"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "18" & "','" & strUserNum & _
                                    "','ªk©w´Á­­','" & m_CP07 & "')"
                           '2014/12/9 END
                           cnnConnection.Execute strSql
                       End If
'end 2018/03/22
'end 2018/03/28
'2014/12/9 CANCEL BY SONIA
'                      Else
'                      '2012/6/27 End
'                        'add by nickc 2007/02/16 ¥[¦h¥Ó½Ð¤H®É¡A¤W­z©w½Z¶W¹L 4000¡A©Ò¥H©î¦¨ 2 ­Ó
'                        EndLetter "03", m_CP09, "09", strUserNum
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & _
'                                 "','¥»©Ò´Á­­','" & m_CP06 & "')"
'                        cnnConnection.Execute strSql
'                      End If
'2014/12/9 END

                   End If
'                End If   '2014/12/9 CANCEL BY SONIA
            ' ¤é¤å
            Case "3":
'2014/12/9 CANCEL BY SONIA
               'Add By Sindy 2019/7/22 ªü½¬»¡¤À³Î¨S¤é¤å©w½Z,¤£­n¥X©w½Z ex:FCT-43164
               If m_CP10 = "101" Then
               '2019/7/22 END
'                '­Y¥Ó½Ð¤é¤p©ó921128
'                If Val(DBDATE(m_TM11)) < 20031128 Then
'                    ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                    EndLetter "03", m_CP09, "05", strUserNum
'                'edit by nick 2004/08/03
'                Else
                   '2008/11/13 ADD BY SONIA FCTµù¥U¶O¦Û°Ê¥NÃº
                   If m_TM122 = "Y" Then
                   Else
                   '2008/11/13 END
                     'add by nick 2005/01/26 ¦]¬°¤À³Îªº¤é¤å©w½Z¦³¶O¥Î¸ò¬üª÷­nÅÜ°Ê
                     'If IIf(strCP10Code <> "", strCP10Code, m_CP10) = "308" Then
                     Dim oRate As Double   '¶×²v
                     Dim o71706 As Double, o71706New As Double '¶O¥Î
                     Dim o71708 As Double  '³W¶O
                     Dim o71707 As Double  'Âø¶O Add By Sindy 2012/3/22
'                     Dim o71606 As Double
'                     Dim o71608 As Double
'                     Dim o71506 As Double
'                     Dim o71507 As Double  'Add By Sindy 2012/3/22
'                     Dim o71508 As Double
                     Dim oFaFee As Double, oFaFeeNew As Double 'Add By Sindy 2025/3/6
                     CheckOC3
                     strSql = "select * from usxrate where USXR01 in (select max(USXR01) from usxrate where USXR01<=to_number(to_char(sysdate, 'YYYYMMDD'))) "
                     AdoRecordSet3.CursorLocation = adUseClient
                     AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If AdoRecordSet3.RecordCount <> 0 Then
                         oRate = AdoRecordSet3.Fields("USXR02").Value
                     End If
                     CheckOC3
'                        strSql = "select * from casefee where cf01='" & m_TM01 & "' and cf02='" & m_TM10 & "' and cf03 in ('715','716','717') order by cf03 "
'                        AdoRecordSet3.CursorLocation = adUseClient
'                        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                        If AdoRecordSet3.RecordCount <> 0 Then
'                            AdoRecordSet3.MoveFirst
'                            Do While Not AdoRecordSet3.EOF
'                                Select Case AdoRecordSet3.Fields("cf03").Value
'                                Case "715"
'                                    o71508 = AdoRecordSet3.Fields("cf08").Value
'                                    o71506 = AdoRecordSet3.Fields("cf06").Value
'                                Case "716"
'                                    o71608 = AdoRecordSet3.Fields("cf08").Value
'                                    o71606 = AdoRecordSet3.Fields("cf06").Value
'                                Case "717"
'                                    o71708 = AdoRecordSet3.Fields("cf08").Value
'                                    o71706 = AdoRecordSet3.Fields("cf06").Value
'                                Case Else
'                                End Select
'                                AdoRecordSet3.MoveNext
'                            Loop
'                        End If
'                        CheckOC3
'                     'Modify By Sindy 2011/5/30
'                     o71508 = 1000
'                     o71507 = 700 'Add By Sindy 2012/3/22
'                     o71506 = 8000
'                     o71608 = 1500
'                     o71606 = 5000
                     o71708 = 2500 '³W¶O
                     o71707 = 700 'Âø¶O Add By Sindy 2012/3/22
                     'Add By Sindy 2013/12/20 ¶O¥Î
                     If m_fa76 = "A" Then 'A.¥N²z¤H«ß®v¨Æ°È©Ò
                        o71706 = 6000
                     Else
                     '2013/12/20 END
                        o71706 = 7000
                     End If
                     oFaFee = 3000 'Add By Sindy 2025/3/6
'                     '2011/5/30 End
                     'Modify By Sindy 2012/6/26 °Ó¼Ð­×ªk
'2014/12/9 CANCEL BY SONIA
'                     If Val(DBDATE(m_CP05)) >= 20120701 Then
                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                        EndLetter "03", m_CP09, "15", strUserNum
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
                                 "','¿ú1','" & Format(o71708, "###,###,##0") & "')"
                        cnnConnection.Execute strSql
                        'intFee = o71708 / oRate 'Modify By Sindy 2010/8/25 o71708 \ oRate
                        'intFee = o71708 \ oRate 'Modify By Sindy 2014/8/22 ¤p¼Æ¦ì¥þ³¡±Ë¥h
                        intFee = Int(o71708 / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
                                 "','¿ú2','" & Format(intFee, "###,###,##0") & "')"
                        cnnConnection.Execute strSql
                        'Modify By Sindy 2012/3/22 old:'" & Format((o71706 - o71708), "###,###,##0") & "'
                        'Modify By Sindy 2012/7/18
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                 "','¿ú3','7,000 + ’Ê¶ONT$700 = NT$7,700')"
                        'Add By Sindy 2013/12/20
                        'Modified by Morgan 2022/12/15 "Âø"§ïUnicode
'                        If m_fa76 = "A" Then 'A.¥N²z¤H«ß®v¨Æ°È©Ò
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                    "','¿ú3','6,000 + " & PUB_GetUniText(Me.Name, "Âø") & "¶ONT$700" & vbCrLf & "                            = NT$6,700')"
'                        Else
'                        '2013/12/20 END
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                    "','¿ú3','7,000 + " & PUB_GetUniText(Me.Name, "Âø") & "¶ONT$700" & vbCrLf & "                            = NT$7,700')"
'                        End If
                        '2012/7/18 End
                        strExc(1) = Format(o71706, "###,###,##0") & IIf(strDisc <> "", " x " & strDisc & "¢H", "") & " + " & PUB_GetUniText(Me.Name, "Âø") & "¶ONT$" & o71707
                        If strDisc <> "" Then
                           o71706New = (o71706 * strDisc / 100) + o71707
                        Else
                           o71706New = o71706 + o71707
                        End If
                        strExc(1) = strExc(1) & vbCrLf & "                            = NT$" & Format(o71706New, "###,###,##0")
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
                                 "','¿ú3','" & strExc(1) & "')"
                        cnnConnection.Execute strSql
                        'Modify By Sindy 2012/3/22
                        'intFee = ((o71706 - o71708) / oRate) 'Modify By Sindy 2010/8/25 ((o71706 - o71708) \ oRate)
                        'intFee = ((o71706 - o71708 + o71707) / oRate)
                        'intFee = ((o71706 - o71708 + o71707) \ oRate) 'Modify By Sindy 2014/8/22 ¤p¼Æ¦ì¥þ³¡±Ë¥h
                        intFee = Int(o71706New / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
                                 "','¿ú4','" & Format(intFee, "###,###,##0") & "')"
                        cnnConnection.Execute strSql
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                 "','¿ú5','" & Format(o71508, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        intFee = o71508 / oRate 'Modify By Sindy 2010/8/25 o71508 \ oRate
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                 "','¿ú6','" & Format((intFee + IIf(o71508 Mod oRate <> 0, 1, 0)), "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        'Modify By Sindy 2012/3/22 old:'" & Format((o71506 - o71508), "###,###,##0") & "'
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                 "','¿ú7','7,000 + ’Ê¶ONT$700 = NT$7,700')"
'                        cnnConnection.Execute strSql
'                        'Modify By Sindy 2012/3/22
'                        'intFee = ((o71506 - o71508) / oRate) 'Modify By Sindy 2010/8/25 ((o71506 - o71508) \ oRate)
'                        intFee = ((o71506 - o71508 + o71507) / oRate)
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                 "','¿ú8','" & Format(intFee, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                 "','¿ú9','" & Format(o71608, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        intFee = o71608 / oRate 'Modify By Sindy 2010/8/25 o71608 \ oRate
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                 "','¿ú10','" & Format((intFee + IIf(o71608 Mod oRate <> 0, 1, 0)), "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                 "','¿ú11','" & Format((o71606 - o71608), "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        intFee = ((o71606 - o71608) / oRate) 'Modify By Sindy 2010/8/25 ((o71606 - o71608) \ oRate)
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                 "','¿ú12','" & Format(intFee, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/2/1 1000§ï3000
                        'Add By Sindy 2013/12/20
                        If m_fa76 = "A" Then 'A.¥N²z¤H«ß®v¨Æ°È©Ò
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                    "','¿ú13','3,000 x 50%" & vbCrLf & "                                              = NT$1,500')"
                           'Modify By Sindy 2015/12/3
                           'strText13 = "3,000 x 50%" & vbCrLf & "                                              = NT$1,500" 'Add By Sindy 2014/3/31
                           'Modify By Sindy 2025/3/6
                           If strDisc = "" Then strDisc = 50
                           oFaFeeNew = (oFaFee * strDisc / 100)
                           strText13 = Format(oFaFeeNew, "###,###,##0") '"1,500" 'Add By Sindy 2014/3/31
                           '2025/3/6 END
                           '2015/12/3 END
'                           cnnConnection.Execute strSql
                           'intFee = 1500 / oRate
                           'intFee = 1500 \ oRate 'Modify By Sindy 2014/8/22 ¤p¼Æ¦ì¥þ³¡±Ë¥h
                           intFee = Int(oFaFeeNew / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                    "','¿ú14','" & Format(intFee, "###,###,##0") & "')"
                           strText14 = Format(intFee, "###,###,##0") 'Add By Sindy 2014/3/31
'                           cnnConnection.Execute strSql
                           
                        Else
                        '2013/12/20 END
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                    "','¿ú13','" & Format(3000, "###,###,##0") & "')"
                           If strDisc <> "" Then
                              oFaFeeNew = (oFaFee * strDisc / 100)
                           Else
                              oFaFeeNew = oFaFee
                           End If
                           strText13 = Format(oFaFeeNew, "###,###,##0") 'Add By Sindy 2014/3/31
'                           cnnConnection.Execute strSql
                           'Modify By Sindy 2010/8/25 1000 \ oRate
                           'Modify By Sindy 2011/2/1 1000§ï3000
                           'intFee = 3000 / oRate
                           'intFee = 3000 \ oRate 'Modify By Sindy 2014/8/22 ¤p¼Æ¦ì¥þ³¡±Ë¥h
                           intFee = Int(oFaFeeNew / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                    "','¿ú14','" & Format(intFee, "###,###,##0") & "')"
                           strText14 = Format(intFee, "###,###,##0")  'Add By Sindy 2014/3/31
'                           cnnConnection.Execute strSql

'                           'Add By Sindy 2014/8/22
'                           If tmpTm09CntS > 1 Then
'                              '10200=(o71706New + o71708)
'                              '5500=(oFaFeeNew + o71708)
'                              intTotFee = (o71706New + o71708) + ((oFaFeeNew + o71708) * (tmpTm09CntS - 1))
'                           Else
'                              intTotFee = (o71706New + o71708)
'                           End If
'                           '2014/8/22 END
                        End If
                        'Add By Sindy 2014/8/22
                        If tmpTm09CntS > 1 Then
                           '9200=(o71706New + o71708)
                           '4000=(oFaFeeNew + o71708)
                           intTotFee = (o71706New + o71708) + ((oFaFeeNew + o71708) * (tmpTm09CntS - 1))
                        Else
                           intTotFee = (o71706New + o71708)
                        End If
                        '2014/8/22 END
                        
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
                                 "','TOTAmtNT','" & Format(intTotFee, "###,###,##0") & "')"
                        cnnConnection.Execute strSql
                        intFee = Int(intTotFee / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
                                 "','TOTAmtUS','" & Format(intFee, "###,###,##0") & "')"
                        cnnConnection.Execute strSql
                        'Add By Sindy 2012/7/18 ¤@®×¦hÃþ§O
                        If tmpTm09CntS > 1 Then
                           'Modify By Sindy 2014/3/31
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
'                                    "','¿ú13©M¿ú14','¡¹2üP¤À¥Ø¥H­°ÇR«YÇr’U©ÒÇU¤â“g®ÆÇV1üP¤ÀÇRþ÷þà¡GNT$<¿ú13> (US$<¿ú14>)')"
                           'Modified by Morgan 2023/3/15
                           'strExc(1) = "¡¹2üP¤À¥Ø¥H­°ÇR«YÇr’U©ÒÇU¤â“g®ÆÇV1üP¤ÀÇRþ÷þà¡GNT$" & strText13 & " (US$" & strText14 & ")"
                           strExc(1) = PUB_GetUniText(Me.Name, "¿ú13©M¿ú14")
                           If strDisc = "" Then
                              strExc(1) = strExc(1) & strText13 & " (US$" & strText14 & ")"
                           Else
                              strExc(1) = strExc(1) & Format(oFaFee, "###,###,##0") & IIf(strDisc <> "", " x " & strDisc & "¢H", "")
                              strExc(1) = strExc(1) & vbCrLf & "                            = NT$" & strText13 & " (US$" & strText14 & ")"
                           End If
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
                                    "','¿ú13©M¿ú14','" & strExc(1) & "')"
                           cnnConnection.Execute strSql
                        End If
                        '2012/7/18 End
                        
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "15" & "','" & strUserNum & _
                                 "','¨ä¥L¤½§i¤é','" & DBDATE(m_CP05) & "')"
                        cnnConnection.Execute strSql
                        
                        'Modify By Sindy 2021/6/28 ¨ó§U­×§ïFCT¤é¤å²Õ¤§¡u¥Ó½Ð(°Ó¥Ó)¡v®Ö­ã©w½Z¡G§R°£¡uFAXªð«H¥Î¯È¡v
'                        'add by nick 2004/10/15 ©î¦¨2 ±i¡A¦]¬°¤£¦P¯È±i
'                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                        EndLetter "03", m_CP09, "16", strUserNum
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "16" & "','" & strUserNum & _
'                                 "','¨ä¥L¤½§i¤é','" & DBDATE(m_CP05) & "')"
'                        cnnConnection.Execute strSql
                        '2021/6/28 END
                        
'2014/12/9 CANCEL BY SONIA
'                     Else
'                     '2012/6/26 End
'                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                        EndLetter "03", m_CP09, "07", strUserNum
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú1','" & Format(o71708, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        'intFee = o71708 / oRate 'Modify By Sindy 2010/8/25 o71708 \ oRate
'                        'intFee = o71708 \ oRate 'Modify By Sindy 2014/8/22 ¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        intFee = Int(o71708 / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú2','" & Format(intFee, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        'Modify By Sindy 2012/3/22 old:'" & Format((o71706 - o71708), "###,###,##0") & "'
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú3','7,000 + ’Ê¶ONT$700 = NT$7,700')"
'                        cnnConnection.Execute strSql
'                        'Modify By Sindy 2012/3/22
'                        'intFee = ((o71706 - o71708) / oRate) 'Modify By Sindy 2010/8/25 ((o71706 - o71708) \ oRate)
'                        'intFee = ((o71706 - o71708 + o71707) / oRate)
'                        'intFee = ((o71706 - o71708 + o71707) \ oRate) 'Modify By Sindy 2014/8/22 ¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        intFee = Int((o71706 - o71708 + o71707) / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú4','" & Format(intFee, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú5','" & Format(o71508, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        'intFee = o71508 / oRate 'Modify By Sindy 2010/8/25 o71508 \ oRate
'                        'intFee = o71508 \ oRate 'Modify By Sindy 2014/8/22 ¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        intFee = Int(o71508 / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú6','" & Format(intFee, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        'Modify By Sindy 2012/3/22 old:'" & Format((o71506 - o71508), "###,###,##0") & "'
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú7','7,000 + ’Ê¶ONT$700 = NT$7,700')"
'                        cnnConnection.Execute strSql
'                        'Modify By Sindy 2012/3/22
'                        'intFee = ((o71506 - o71508) / oRate) 'Modify By Sindy 2010/8/25 ((o71506 - o71508) \ oRate)
'                        'intFee = ((o71506 - o71508 + o71507) / oRate)
'                        'intFee = ((o71506 - o71508 + o71507) \ oRate) 'Modify By Sindy 2014/8/22 ¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        intFee = Int((o71506 - o71508 + o71507) / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú8','" & Format(intFee, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú9','" & Format(o71608, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        'intFee = o71608 / oRate 'Modify By Sindy 2010/8/25 o71608 \ oRate
'                        'intFee = o71608 \ oRate 'Modify By Sindy 2014/8/22 ¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        intFee = Int(o71608 / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú10','" & Format(intFee, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú11','" & Format((o71606 - o71608), "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        'intFee = ((o71606 - o71608) / oRate) 'Modify By Sindy 2010/8/25 ((o71606 - o71608) \ oRate)
'                        'intFee = ((o71606 - o71608) \ oRate) 'Modify By Sindy 2014/8/22 ¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        intFee = Int((o71606 - o71608) / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú12','" & Format(intFee, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        'Modify By Sindy 2011/2/1 1000§ï3000
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú13','" & Format(3000, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        'Modify By Sindy 2010/8/25 1000 \ oRate
'                        'Modify By Sindy 2011/2/1 1000§ï3000
'                        'intFee = 3000 / oRate
'                        'intFee = 3000 \ oRate 'Modify By Sindy 2014/8/22 ¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        intFee = Int(3000 / oRate) 'Modify By Sindy 2014/9/16 °£ªk,¤p¼Æ¦ì¥þ³¡±Ë¥h
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¿ú14','" & Format(intFee, "###,###,##0") & "')"
'                        cnnConnection.Execute strSql
'                        'End If
'                        'add end
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
'                                 "','¨ä¥L¤½§i¤é','" & DBDATE(m_CP05) & "')"
'                        cnnConnection.Execute strSql
'                        'add by nick 2004/10/15 ©î¦¨2 ±i¡A¦]¬°¤£¦P¯È±i
'                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                        EndLetter "03", m_CP09, "08", strUserNum
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & _
'                                 "','¨ä¥L¤½§i¤é','" & DBDATE(m_CP05) & "')"
'                        cnnConnection.Execute strSql
'                     End If
'                   End If
'2014/12/9 CANCEL BY SONIA
                   
                    '2008/11/13 ±NÄ¶¤å¦Û08FAXªð«H¥Î¯È¿W¥ß¥X¨Ó12,¦Û°Ê¥NÃº¤]­n¦L
                    EndLetter "03", m_CP09, "12", strUserNum
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & _
                             "','¶O¥Î1','" & Format(Trim(tmpTm09CntS * 2500), "###,###") & "')"
                    cnnConnection.Execute strSql
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & _
                             "','¶O¥Î2','" & Format(Trim(tmpTm09CntS * 1000), "###,###") & "')"
                    cnnConnection.Execute strSql
                    'add by nickc 2005/11/22 ¤é¤å¥[¤J©ñ±ó±M¥ÎÅv may ¸ò ªü½¬
                    'Modify by Morgan 2008/5/28 +ChgSQL ¦]¬°¤º®e·|¦³³æ¤Þ¸¹¡A FCT-26349
                    'If m_TM67 <> "" Then
                    If Trim(textTM67) <> "" Then
                        'Modify By Sindy 2022/10/12 ˆü¥e“¸Çy¦³ §ï¬° °Ó¼Ð“¸Çy¥D±i
                        'Modified by Morgan 2023/3/15
                        'strExc(1) = "°Ó¼Ð¨£¥»ÇRÆèÇr¡u" & ChgSQL(textTM67) & "¡vÇU°Ó¼Ð“¸Çy¥D±iþêÇQÆê¡C"
                        strExc(1) = PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv1") & ChgSQL(textTM67) & PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv2")
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & _
                                 "','©ñ±ó±M¥ÎÅv','" & strExc(1) & "')"
                        cnnConnection.Execute strSql
                    End If
                    'Add By Sindy 2010/11/17
                    If m_TM118 <> "" Then
                        'Modified by Morgan 2023/3/15
                        'strExc(1) = "°Ó¼Ðªk²Ä30’f²Ä1¶µ²Ä10†AÇU³W©wÇR°òþøþà¡Bµn“÷°Ó¼Ð²Ä" & ChgSQL(m_TM118) & "†AÇU°Ó¼Ð“¸ªÌÇU¦P·NÇRÇoÇqµn“÷Çy³\¥iþìÇr¡C"
                        strExc(1) = PUB_GetUniText(Me.Name, "¦P·N®Ñ°Ó¼Ð¸¹¼Æ1") & ChgSQL(m_TM118) & PUB_GetUniText(Me.Name, "¦P·N®Ñ°Ó¼Ð¸¹¼Æ2")
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & _
                                 "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & strExc(1) & "')"
                        cnnConnection.Execute strSql
                    End If
                    '2010/11/17 End
                    
                    'Add By Sindy 2011/6/15
                    'Àu¥ýÅv¸ê®Æ
                    strExc(0) = "select pd05,pd07,na03,pd06,pd10 from pridate,nation " & _
                                "where pd01='" & m_TM01 & "' and pd02='" & m_TM02 & "' and pd03='" & m_TM03 & "' and pd04='" & m_TM04 & "' " & _
                                "and pd07=na01 "
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                    strExc(1) = ""
'                    strExc(2) = ""
'                    If intI = 1 Then
'                        strExc(1) = "" & RsTemp.Fields("pd05")
'                        If strExc(1) <> "" Then strExc(1) = Left(strExc(1), 4) & "¦~" & Val(Mid(strExc(1), 5, 2)) & "¤ë" & Val(Right(strExc(1), 2)) & "¤é"
'                        strExc(2) = "" & RsTemp.Fields("na03")
'                        If strExc(2) = "¤é¥»" Then strExc(2) = strExc(2) & "üÂ"
'                    End If
'                    If strExc(1) <> "" Then
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & _
'                                 "','¥D±iÀu¥ýÅv','Àu¥ý“¸ûÐ¥Í¤é¤ÎÇZþðÇUÀu¥ý“¸¥D±iüÂ¡G" & strExc(1) & "¡@" & strExc(2) & "')"
'                        cnnConnection.Execute strSql
'                    End If
                    'Modify By Sindy 2017/8/11 ­ì¥u§ì³æµ§§ï¬°¦hµ§
                    intRow = 0: strTemp = ""
                    If intI = 1 Then
                        RsTemp.MoveFirst
                        'Add By Sindy 2018/4/9 Àu¥ýÅv³æµ§
                        If RsTemp.RecordCount = 1 Then
                           strExc(1) = "" & RsTemp.Fields("pd05")
                           If strExc(1) <> "" Then strExc(1) = Left(strExc(1), 4) & "¦~" & Mid(strExc(1), 5, 2) & "¤ë" & Right(strExc(1), 2) & "¤é"
                           strExc(2) = "" & RsTemp.Fields("na03")
                           'Modified by Morgan 2023/3/15
                           'If strExc(2) = "¤é¥»" Then strExc(2) = strExc(2) & "üÂ"
                           If strExc(2) = "¤é¥»" Then strExc(2) = strExc(2) & PUB_GetUniText(Me.Name, "°ê")
                           'end 2023/3/15
                           strTemp = strExc(1) & " " & strExc(2)
                        Else
                        '2018/4/9 END
                           Do While Not RsTemp.EOF
                              intRow = intRow + 1
                              strExc(1) = "" & RsTemp.Fields("pd05")
                              If strExc(1) <> "" Then strExc(1) = Left(strExc(1), 4) & "¦~" & Mid(strExc(1), 5, 2) & "¤ë" & Right(strExc(1), 2) & "¤é"
                              strExc(2) = "" & RsTemp.Fields("na03")
                              'Modified by Morgan 2023/3/15
                              'If strExc(2) = "¤é¥»" Then strExc(2) = strExc(2) & "üÂ"
                              If strExc(2) = "¤é¥»" Then strExc(2) = strExc(2) & PUB_GetUniText(Me.Name, "°ê")
                              'end 2023/3/15
                              strExc(3) = "" & RsTemp.Fields("pd10")
   '                           If strExc(3) <> "" And InStr(strExc(3), "(") > 0 Then
   '                              strExc(3) = Mid(strExc(3), InStr(strExc(3), "(Cl.") + 4)
   '                              strExc(3) = Left(strExc(3), InStr(strExc(3), ")") - 1)
   '                           End If
                              'Modified by Morgan 2023/3/15
                              'strTemp = strTemp & vbCrLf & "¡@¡@¡@¡@¡@" & intRow & "." & strExc(1) & "¡@" & strExc(2) & vbCrLf & "¡@¡@¡@¡@¡@¡@Àu¥ý“¸ÇU°Ó«~¡G²Ä" & Trim(strExc(3)) & "Ãþ"
                              strTemp = strTemp & vbCrLf & "¡@¡@¡@¡@¡@" & intRow & "." & strExc(1) & "¡@" & strExc(2) & vbCrLf & "¡@¡@¡@¡@¡@¡@" & PUB_GetUniText(Me.Name, "Àu¥ýÅvªº°Ó«~") & "¡G²Ä" & Trim(strExc(3)) & "Ãþ"
                              'end 2023/3/15
                              RsTemp.MoveNext
                           Loop
                        End If
                    End If
                    If strTemp <> "" Then
                        'Modified by Morgan 2023/3/15
                        'strExc(1) = "Àu¥ý“¸ûÐ¥Í¤é¤ÎÇZþðÇUÀu¥ý“¸¥D±iüÂ¡G" & strTemp
                        strExc(1) = PUB_GetUniText(Me.Name, "¥D±iÀu¥ýÅv") & strTemp
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & _
                                 "','¥D±iÀu¥ýÅv','" & strExc(1) & "')"
                        cnnConnection.Execute strSql
                    End If
                    '2017/8/11 END
                    '2011/6/15 END
                    
                    'Add By Sindy 2013/1/22
                    '¥Ó½Ð¤é¦b1010630(§t0630)«e¥Ó½Ð¤§®×¥ó¬°13¡F1010701¥H«á¥Ó½Ð¤§®×¥ó¬°19
                    If DBDATE(m_TM11) <= 20120630 Then
                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                "VALUES ('" & "03" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & _
                                "','±ø´Ú','13')"
                       cnnConnection.Execute strSql
                    Else
                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                "VALUES ('" & "03" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & _
                                "','±ø´Ú','19')"
                       cnnConnection.Execute strSql
                    End If
                    '2013/1/22 End
                   End If
'                End If  '2014/12/9 CANCEL BY SONIA
               End If
         End Select
      ' ©µ®i
      Case "102":
         ' ©w½Z»y¤å
         Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
            ' ­^¤å
            Case "2":
               'Modify By Sindy 2010/5/13 ¦³ÅÜ§ó¥Ó½Ð¤H
               'If bChkChaEvent = True Then
               If m_strCE04 <> "" Then
                     'Modify By Sindy 2012/2/1 ¥Ñ©µ®i±µ¶i¨Ó¤§·s®×¥X¤£¦P©w½Z(­^Ä¶¤å¤£ÅÜ)
                     If bolChaEventNewCase = True Then
                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                        EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "17", strUserNum
                        ' ¨ä¥L¤½§i¤é
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "17" & "','" & strUserNum & _
                                 "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                        cnnConnection.Execute strSql
'                        'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                        If bolEmail = True And bolPlusPaper = False Then
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "17" & "','" & strUserNum & _
'                                    "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the Notice of Approval and its translation indicating the goods/services renewed. The originals will be sent to you via registered mail.')"
'                           cnnConnection.Execute strSql
'                        Else '¶l¥ó
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "17" & "','" & strUserNum & _
'                                    "','¨Ò¥~¤º¤å','A copy of the Notice of Approval and its translation indicating the goods/services renewed will be mailed to you with the confirmation copy of this letter for your records.')"
'                           cnnConnection.Execute strSql
'                        End If
'                        '2012/11/27 End
                     Else
                     '2012/2/1 End
                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                        EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "14", strUserNum
                        ' ¨ä¥L¤½§i¤é
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "14" & "','" & strUserNum & _
                                 "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                        cnnConnection.Execute strSql
'                        'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                        If bolEmail = True And bolPlusPaper = False Then
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "14" & "','" & strUserNum & _
'                                    "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the official notice and its translation indicating the goods/services renewed. The originals will be sent to you via registered mail.')"
'                           cnnConnection.Execute strSql
'                        Else '¶l¥ó
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "14" & "','" & strUserNum & _
'                                    "','¨Ò¥~¤º¤å','A copy of the official notice and its translation indicating the goods/services renewed will be mailed to you with the confirmation copy of this letter for your records.')"
'                           cnnConnection.Execute strSql
'                        End If
'                        '2012/11/27 End
                     End If
                     ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                     EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "15", strUserNum
                     ' ¨ä¥L¤½§i¤é
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "15" & "','" & strUserNum & _
                              "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                     cnnConnection.Execute strSql
               '2010/5/13 End
               Else
'                     '­Yªþµù¥UÃÒ(ÂÂ©w½Z)
'                     If m_strWithRegister <> "N" Then
'                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                        EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "06", strUserNum
'                        'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                        If bolEmail = True And bolPlusPaper = False Then
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "06" & "','" & strUserNum & _
'                                    "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the Notice of Approval with its English translation. The originals will be sent to you via registered mail.')"
'                           cnnConnection.Execute strSql
'                        Else '¶l¥ó
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "06" & "','" & strUserNum & _
'                                    "','¨Ò¥~¤º¤å','The original Registration Certificate with English translation and official notice indicating the goods/services renewed will be mailed to you with the confirmation of this letter for your records.')"
'                           cnnConnection.Execute strSql
'                        End If
'                        '2012/11/27 End
'                        ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                        If textPrtTrans <> "N" Then
'                           ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                             'Modify By Cheng 2003/03/12
'         '                  EndLetter "03", m_CP09, IIf(m_TM08 = "2" Or m_TM08 = "5", "08", "07"), strUserNum
'                           EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), IIf(m_TM08 = "2", "08", IIf(m_TM08 = "5", "09", "07")), strUserNum
'                           'Add By Cheng 2003/03/11
'                           ' ©ñ±ó±M¥ÎÅv
''                           If IsEmptyText(m_TM67) = False Then
'                           If IsEmptyText(Trim(textTM67)) = False Then
'                              ' Áp¦X°Ó¼Ð
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & IIf(m_TM08 = "2", "08", IIf(m_TM08 = "5", "09", "07")) & "','" & strUserNum & _
'                                       "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & ChgSQL(textTM67) & "')"
'                              cnnConnection.Execute strSql
'                           End If
'                        End If
'                     '­Y¤£ªþµù¥UÃÒ(·s©w½Z)
'                     Else
'                        '2011/9/7 ADD BY SONIA ¥Ñ©µ®i±µ¶i¨Ó¤§·s®×¥X¤£¦P©w½Z(­^Ä¶¤å¤£ÅÜ)
'                        StrSQLa = "Select C1.cp05,C2.cp09,C3.cp09,C3.cp05 From CaseProgress C1,CaseProgress C2,CaseProgress C3 Where C1.cp09='" & m_CP09 & "' " & _
'                                  "AND C1.cp01=C2.cp01(+) and C1.cp02=C2.cp02(+) and C1.cp03=C2.cp03(+) and C1.cp04=C2.cp04(+) and '101'=C2.cp10(+) " & _
'                                  "AND C1.cp01=C3.cp01(+) and C1.cp02=C3.cp02(+) and C1.cp03=C3.cp03(+) and C1.cp04=C3.cp04(+) and '102'=C3.cp10(+) " & _
'                                  "order by c3.cp05"
'                        rsA.CursorLocation = adUseClient
'                        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'                        If rsA.RecordCount > 0 Then
'                           If "" & rsA.Fields(1) = "" And Val(rsA.Fields(3)) = Val(rsA.Fields(0)) Then
                           'Modify By Sindy 2012/2/1
                           If bolChaEventNewCase = True Then
                              ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                              EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "16", strUserNum
                              ' ¨ä¥L¤½§i¤é
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "16" & "','" & strUserNum & _
                                       "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                              cnnConnection.Execute strSql
'                              'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                              If bolEmail = True And bolPlusPaper = False Then
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "16" & "','" & strUserNum & _
'                                          "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the Notice of Approval and its translation indicating the goods/services renewed. The originals will be sent to you via registered mail.')"
'                                 cnnConnection.Execute strSql
'                              Else '¶l¥ó
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "16" & "','" & strUserNum & _
'                                          "','¨Ò¥~¤º¤å','A copy of the Notice of Approval and its translation indicating the goods/services renewed will be mailed to you with the confirmation copy of this letter for your records.')"
'                                 cnnConnection.Execute strSql
'                              End If
'                              '2012/11/27 End

                              'Add By Sindy 2013/12/30
                              If m_TM08 = "7" Then 'ÃÒ©ú¼Ð³¹
                                 ' °Ó¼ÐºØÃþ¤º¤å¤@
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "16" & "','" & strUserNum & _
                                          "','°Ó¼ÐºØÃþ¤º¤å¤@','')"
                                 cnnConnection.Execute strSql
                                 'Modify By Sindy 2022/6/13 Mark
'                                 ' °Ó¼ÐºØÃþ¤º¤å¤G
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "16" & "','" & strUserNum & _
'                                          "','°Ó¼ÐºØÃþ¤º¤å¤G','contents of certification')"
'                                 cnnConnection.Execute strSql
                              Else
                                 ' °Ó¼ÐºØÃþ¤º¤å¤@
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "16" & "','" & strUserNum & _
                                          "','°Ó¼ÐºØÃþ¤º¤å¤@',' indicating the goods/services renewed')"
                                 cnnConnection.Execute strSql
                                 'Modify By Sindy 2022/6/13 Mark
'                                 ' °Ó¼ÐºØÃþ¤º¤å¤G
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "16" & "','" & strUserNum & _
'                                          "','°Ó¼ÐºØÃþ¤º¤å¤G','specification of goods')"
'                                 cnnConnection.Execute strSql
                              End If
                              '2013/12/30 END
                           '2012/2/1 End
                           Else
                           '2011/9/7 END
                              'Modify By Sindy 2012/6/27 °Ó¼Ð­×ªk
                              If Val(strSrvDate(1)) >= 20120701 Then
                                 ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                                 EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "18", strUserNum
                                 ' ¨ä¥L¤½§i¤é
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "18" & "','" & strUserNum & _
                                          "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                                 cnnConnection.Execute strSql
'                                 'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                                 If bolEmail = True And bolPlusPaper = False Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "18" & "','" & strUserNum & _
'                                             "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the Notice of Approval and its translation indicating the goods/services renewed. The originals will be sent to you via registered mail.')"
'                                    cnnConnection.Execute strSql
'                                 Else '¶l¥ó
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "18" & "','" & strUserNum & _
'                                             "','¨Ò¥~¤º¤å','A copy of the Notice of Approval and its translation indicating the goods/services renewed will be mailed to you with the confirmation copy of this letter for your records.')"
'                                    cnnConnection.Execute strSql
'                                 End If
'                                 '2012/11/27 End
                              Else
                              '2012/6/27 End
                                 ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                                 EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "10", strUserNum
                                 ' ¨ä¥L¤½§i¤é
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "10" & "','" & strUserNum & _
                                          "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                                 cnnConnection.Execute strSql
'                                 'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                                 If bolEmail = True And bolPlusPaper = False Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "10" & "','" & strUserNum & _
'                                             "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the Notice of Approval and its translation indicating the goods/services renewed. The originals will be sent to you via registered mail.')"
'                                    cnnConnection.Execute strSql
'                                 Else '¶l¥ó
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "10" & "','" & strUserNum & _
'                                             "','¨Ò¥~¤º¤å','A copy of the Notice of Approval and its translation indicating the goods/services renewed will be mailed to you with the confirmation copy of this letter for your records.')"
'                                    cnnConnection.Execute strSql
'                                 End If
'                                 '2012/11/27 End
                              End If
                           End If
'                        End If
                        ' ¬O§_¦C¦LÂ½Ä¶¨ç
                        If textPrtTrans <> "N" Then
                           ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                             'Modify By Cheng 2003/03/12
         '                  EndLetter "03", m_CP09, IIf(m_TM08 = "2" Or m_TM08 = "5", "08", "07"), strUserNum
                           EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "11", strUserNum
                           'Add By Cheng 2003/03/11
                           ' ©ñ±ó±M¥ÎÅv
'                           If IsEmptyText(m_TM67) = False Then
                           If IsEmptyText(Trim(textTM67)) = False Then
                              ' Áp¦X°Ó¼Ð
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','11','" & strUserNum & _
                                       "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & ChgSQL(textTM67) & "')"
                              cnnConnection.Execute strSql
                           End If
                           ' ¨ä¥L¤½§i¤é
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "11" & "','" & strUserNum & _
                                    "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                           cnnConnection.Execute strSql
                           'Add By Sindy 2013/12/30
                           If m_TM08 = "7" Then 'ÃÒ©ú¼Ð³¹
                              ' °Ó¼ÐºØÃþ¤º¤å¤@
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "11" & "','" & strUserNum & _
                                       "','°Ó¼ÐºØÃþ¤º¤å¤@','contents of certification')"
                              cnnConnection.Execute strSql
                           Else
                              ' °Ó¼ÐºØÃþ¤º¤å¤@
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "11" & "','" & strUserNum & _
                                       "','°Ó¼ÐºØÃþ¤º¤å¤@','specification of good/services')"
                              cnnConnection.Execute strSql
                           End If
                           '2013/12/30 END
                        End If
'                     End If
                End If
            ' ¤é¤å
            Case "3":
'                '­Yªþµù¥UÃÒ(ÂÂ©w½Z)
'                If m_strWithRegister <> "N" Then
'                    ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                    'Modify By Cheng 2002/12/18
'                    '               EndLetter "03", m_CP09, "08", strUserNum
'                    EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "09", strUserNum
'                    ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                    If textPrtTrans <> "N" Then
'                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                        EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "09", strUserNum
'                        ' Áp¦X°Ó¼Ð
'                        If IsEmptyText(m_TM27) = False Then
'                            ' Áp¦X°Ó¼Ð
'                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "09" & "','" & strUserNum & _
'                                        "','Áp¦X°Ó¼Ð','" & "¨Ì¦s ¥¿°Ó¼Ð µn¿ýµf¸¹ : (" & m_TM27 & ")" & "')"
'                            cnnConnection.Execute strSql
'                        End If
'                    End If
'                '­Y¤£ªþµù¥UÃÒ(·s©w½Z)
'                Else
                 'Add By Sindy 2016/12/16 ÀË¬d¬O§_¦³ÅÜ§ó¨Æ¶µ
                If m_strCE04 <> "" Or m_strCE23CE24CE25 <> "" Then
                   ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                    EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "19", strUserNum
                    ' ÅÜ§ó¨Æ¶µ
                    If m_strCE04 <> "" And m_strCE23CE24CE25 <> "" Then
                       'Modified by Morgan 2023/3/15
                       'strExc(1) = "¡]°Ó¼Ð“¸ªÌÇUªí¥Ü¤ÎÇZ¦í©ÒŒi§ó³\¥iþß§tÇeÇsþòÇiÇU¡^"
                       'Added by Lydia 2023/09/04 ®Ö­ã-§ó¥¿(©µ®i®Ö­ã¨ç)¡AÅÜ§ó¨Æ¶µ²Î¤@¥Î¦P¤@ºØ´y­z¡A¨ä¾l¥Ñ©Ó¿ì¤H­û¤H¤u­×§ï
                       If txtADate.Visible = True And txtADate <> "" Then
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "19" & "','" & strUserNum & _
                                    "','¨ä¥Lªþ¥ó','¡B­q¥¿³qª¾®Ñ')"
                          cnnConnection.Execute strSql
                          strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ7")
                       Else
                       'end 2023/09/04
                          strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ1")
                       End If 'Added by Lydia 2023/09/04
                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "19" & "','" & strUserNum & _
                                "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                       cnnConnection.Execute strSql
                    ElseIf m_strCE04 <> "" Then
                       'Modified by Morgan 2023/3/15
                       'strExc(1) = "¡]°Ó¼Ð“¸ªÌÇUªí¥ÜŒi§ó³\¥iþß§tÇeÇsþòÇiÇU¡^"
                       'Added by Lydia 2023/09/04 ®Ö­ã-§ó¥¿(©µ®i®Ö­ã¨ç)¡AÅÜ§ó¨Æ¶µ²Î¤@¥Î¦P¤@ºØ´y­z¡A¨ä¾l¥Ñ©Ó¿ì¤H­û¤H¤u­×§ï
                       If txtADate.Visible = True And txtADate <> "" Then
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "19" & "','" & strUserNum & _
                                    "','¨ä¥Lªþ¥ó','¡B­q¥¿³qª¾®Ñ')"
                          cnnConnection.Execute strSql
                          strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ7")
                       Else
                       'end 2023/09/04
                          strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ2")
                       End If 'Added by Lydia 2023/09/04
                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "19" & "','" & strUserNum & _
                                "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                       cnnConnection.Execute strSql
                    ElseIf m_strCE23CE24CE25 <> "" Then
                       'Modified by Morgan 2023/3/15
                       'strExc(1) = "¡]°Ó¼Ð“¸ªÌÇU¦í©ÒŒi§ó³\¥iþß§tÇeÇsþòÇiÇU¡^"
                       'Added by Lydia 2023/09/04 ®Ö­ã-§ó¥¿(©µ®i®Ö­ã¨ç)¡AÅÜ§ó¨Æ¶µ²Î¤@¥Î¦P¤@ºØ´y­z¡A¨ä¾l¥Ñ©Ó¿ì¤H­û¤H¤u­×§ï
                       If txtADate.Visible = True And txtADate <> "" Then
                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "19" & "','" & strUserNum & _
                                    "','¨ä¥Lªþ¥ó','¡B­q¥¿³qª¾®Ñ')"
                          cnnConnection.Execute strSql
                          strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ7")
                       Else
                       'end 2023/09/04
                          strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ3")
                       End If 'Added by Lydia 2023/09/04
                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "19" & "','" & strUserNum & _
                                "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                       cnnConnection.Execute strSql
                    End If
                    ' ¬O§_¦C¦LÂ½Ä¶¨ç
                    If textPrtTrans <> "N" Then
                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                        EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "20", strUserNum
                        'Added by Lydia 2023/09/04 ­×§ï¤é¤å²Õ¤§®Ö­ã-§ó¥¿(©µ®i®Ö­ã¨ç)¤§©w½Z¤ÎÄ¶¤å:©µ®iÄ¶¤å¥Î­ì¨ç¤½§i¤é
                        If txtADate.Visible = True And txtADate <> "" Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                    "','­ì¨ç¤½§i¤é','" & DBDATE(Me.txtADate.Text) & "')"
                           cnnConnection.Execute strSql
                           '¥t¥~²£¥Í©w½Z
                           EndLetter "03", m_CP09, ET03_ex, strUserNum
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & ET03_ex & "','" & strUserNum & _
                                    "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                           cnnConnection.Execute strSql
                        Else
                        'end 2023/09/04
                        ' ¨ä¥L¤½§i¤é
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                    "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                           cnnConnection.Execute strSql
                        End If 'Addded by Lydia 2023/09/04
                        ' ÅÜ§ó¨Æ¶µ
                        If m_strCE04 <> "" And m_strCE23CE24CE25 <> "" Then
                           'Modified by Morgan 2023/3/15
                           'strExc(1) = "°Ó¼Ð“¸ªÌÇUªí¥Ü¤ÎÇZ¦í©ÒŒi§óµn“÷¥Ó½ÐÇyÇi»{ÇhÇr¡C"
                           strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ4")
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                    "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                           cnnConnection.Execute strSql
                        ElseIf m_strCE04 <> "" Then
                           'Modified by Morgan 2023/3/15
                           'strExc(1) = "°Ó¼Ð“¸ªÌÇUªí¥ÜŒi§óµn“÷¥Ó½ÐÇyÇi»{ÇhÇr¡C"
                           strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ5")
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                    "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                           cnnConnection.Execute strSql
                        ElseIf m_strCE23CE24CE25 <> "" Then
                           'Modified by Morgan 2023/3/15
                           'strExc(1) = "°Ó¼Ð“¸ªÌÇU¦í©ÒŒi§óµn“÷¥Ó½ÐÇyÇi»{ÇhÇr¡C"
                           strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ6")
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                    "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                           cnnConnection.Execute strSql
                        End If
' ÅÜ§ó¨Æ¶µSubject
                        'Modify By Sindy 2017/1/9 + ChgSQL
                        If m_strCE04 <> "" And m_strCE23CE24CE25 <> "" Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                    "','ÅÜ§ó¨Æ¶µSubject','" & ChgSQL("the change of the Registrant's name and address is also recorded.") & "')"
                           cnnConnection.Execute strSql
                        ElseIf m_strCE04 <> "" Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                    "','ÅÜ§ó¨Æ¶µSubject','" & ChgSQL("the change of the Registrant's name is also recorded.") & "')"
                           cnnConnection.Execute strSql
                        ElseIf m_strCE23CE24CE25 <> "" Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                    "','ÅÜ§ó¨Æ¶µSubject','" & ChgSQL("the change of the Registrant's address is also recorded.") & "')"
                           cnnConnection.Execute strSql
                        End If
                        '2017/1/9 ENDection.Execute strSql
                    End If
                Else
                '2016/12/16 END
                    ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                    EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "12", strUserNum
                    ' ¬O§_¦C¦LÂ½Ä¶¨ç
                    If textPrtTrans <> "N" Then
                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                        EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "13", strUserNum
                        'Added by Lydia 2023/09/04 ­×§ï¤é¤å²Õ¤§®Ö­ã-§ó¥¿(©µ®i®Ö­ã¨ç)¤§©w½Z¤ÎÄ¶¤å:©µ®iÄ¶¤å¥Î­ì¨ç¤½§i¤é
                        If txtADate.Visible = True And txtADate <> "" Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "13" & "','" & strUserNum & _
                                    "','­ì¨ç¤½§i¤é','" & DBDATE(Me.txtADate.Text) & "')"
                           cnnConnection.Execute strSql
                           '¥t¥~²£¥Í©w½Z
                           EndLetter "03", m_CP09, ET03_ex, strUserNum
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & ET03_ex & "','" & strUserNum & _
                                    "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                           cnnConnection.Execute strSql
                        Else
                        'end 2023/09/04
                        ' ¨ä¥L¤½§i¤é
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "13" & "','" & strUserNum & _
                                    "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                           cnnConnection.Execute strSql
                        End If 'Added by Lydia 2023/09/04
                    End If
                End If
         End Select
      ' ²¾Âà
      Case "501":
         ' ©w½Z»y¤å
         Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
            ' ¤¤¤å
            Case "1":
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "10", strUserNum
               ' ¨÷¼Æ
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "10" & "','" & strUserNum & _
                        "','¨÷¼Æ','" & textTMBM07_1 & "')"
               cnnConnection.Execute strSql
               ' ´Á¼Æ
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "10" & "','" & strUserNum & _
                        "','´Á¼Æ','" & textTMBM07_2 & "')"
               cnnConnection.Execute strSql
               ' ¦C¦L³Æµù
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "10" & "','" & strUserNum & _
                        "','¦C¦L³Æµù','" & ChgSQL(textPS) & "')"
               cnnConnection.Execute strSql
               'Add By Cheng 2002/06/14
               ' ¨ä¥L¤½§i¤é
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "10" & "','" & strUserNum & _
                        "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
               cnnConnection.Execute strSql
            ' ­^¤å
            Case "2":
                'Modify By Sindy 2012/10/12 Mark¤w¤£°Ï¤À¤F
'                '­Yªþµù¥UÃÒ(ÂÂ©w½Z)
'                If m_strWithRegister <> "N" Then
'                   ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                   EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "11", strUserNum
'                   ' ²¾Âà¤H
'                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                            "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "11" & "','" & strUserNum & _
'                            "','²¾Âà¤H','" & GetCustomerEngName(m_CP55) & "')"
'                   cnnConnection.Execute strSql
'                   ' ²¾Âà¥Ó½Ð¤H
'                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                            "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "11" & "','" & strUserNum & _
'                            "','²¾Âà¥Ó½Ð¤H','" & GetCustomerEngName(m_CP56) & "')"
'                   cnnConnection.Execute strSql
'                    '92.2.18 ADD BY SONIA
'                    ' ¨ä¥L¤½§i¤é
'                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "11" & "','" & strUserNum & _
'                             "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
'                    cnnConnection.Execute strSql
'                   ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                   If textPrtTrans <> "N" Then
'                      ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                      EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "12", strUserNum
'                      ' ²¾Âà¤H
'                      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "12" & "','" & strUserNum & _
'                               "','²¾Âà¤H','" & GetCustomerEngName(m_CP55) & "')"
'                      cnnConnection.Execute strSql
'                      ' ²¾Âà¥Ó½Ð¤H
'                      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "12" & "','" & strUserNum & _
'                               "','²¾Âà¥Ó½Ð¤H','" & GetCustomerEngName(m_CP56) & "')"
'                      cnnConnection.Execute strSql
'                      'Add By Cheng 2003/03/13
'                      ' ©ñ±ó±M¥ÎÅv
''                      If IsEmptyText(m_TM67) = False Then
'                      If IsEmptyText(Trim(textTM67)) = False Then
'                         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                  "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "12" & "','" & strUserNum & _
'                                  "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & ChgSQL(textTM67) & "')"
'                         cnnConnection.Execute strSql
'                      End If
'                      ' ¥¿°Ó¼Ð¸¹¼Æ
'                      If IsEmptyText(m_TM27) = False Then
'                         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                  "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "12" & "','" & strUserNum & _
'                                  "','¥¿°Ó¼Ð¸¹¼Æ','" & "Its Principal " & IIf(m_TM08 >= "4" And m_TM08 <= "6", "Service Mark", "Trademark") & " No. : " & m_TM27 & "')"
'                         cnnConnection.Execute strSql
'                      End If
'                        'Modify By Cheng 2003/03/13
'    '                  '92.2.18 ADD BY SONIA
'    '                  ' ¨ä¥L¤½§i¤é
'    '                  strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'    '                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & _
'    '                           "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
'    '                  cnnConnection.Execute strSQL
'    '                  '92.2.18 END
'                   End If
'                '­Y¤£ªþµù¥UÃÒ(·s©w½Z)
'                Else
                  ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                  EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), ET03, strUserNum
                  ' ²¾Âà¤H
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                           "','²¾Âà¤H','" & GetCustomerEngName(m_CP55) & "')"
                  cnnConnection.Execute strSql
                  ' ²¾Âà¥Ó½Ð¤H
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                           "','²¾Âà¥Ó½Ð¤H','" & GetCustomerEngName(m_CP56) & "')"
                  cnnConnection.Execute strSql
                  '92.2.18 ADD BY SONIA
                  ' ¨ä¥L¤½§i¤é
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                           "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                  cnnConnection.Execute strSql
                  'add by nickc 2008/05/08 ­Y¬O²¾Âà®Ö­ã¡AÀË¬d¥b¦~«e«áµo¤åªº©µ®i¡A§ì¨ä±ÂÅv´Á¶¡¤î¤é¡A­YµL¡AÁÙ¬O§ì°ò¥»ÀÉªº±M¥Î´Á¤î¤é
                  Dim m_tmpday As String
                  Dim m_rs As New ADODB.Recordset
                  m_tmpday = ""
                  Set m_rs = New ADODB.Recordset
                  strSql = "select cp54 from caseprogress where cp09 in (select max(cp09) from caseprogress where cp10='102' and cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp27>='" & DBDATE(DateAdd("m", -6, ChangeWStringToWDateString(DBDATE(m_CP27)))) & "' and cp27<='" & DBDATE(DateAdd("m", 6, ChangeWStringToWDateString(DBDATE(m_CP27)))) & "' )"
                  If m_rs.State = 1 Then m_rs.Close
                  m_rs.CursorLocation = adUseClient
                  m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If Not m_rs.EOF And Not m_rs.BOF Then
                      m_tmpday = "" & m_rs.Fields("cp54")
                  End If
                  If m_tmpday = "" Then
                      m_tmpday = textTM22
                  End If
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                           "','¨Ò¥~±M¥Î´Á¶¡','" & DBDATE(m_tmpday) & "')"
                  cnnConnection.Execute strSql
                  '­Y¥Ó½Ð¤é¤p©óµ¥©ó930324
                  'Modify by Morgan 2004/5/27
                  '§ï§ìµo¤å¤é
                  'If Val(DBDATE(m_TM11)) <= 20040324 Then
                  If Val(m_CP27) <= 20040324 Then
                      ' ½Ðµ²²M½Ð´Ú³æ
                      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                               "','½Ðµ²²M½Ð´Ú³æ','" & " Our final debit note is also enclosed for your kind settlement." & "')"
                      cnnConnection.Execute strSql
                  End If
                  'Add By Sindy 2012/10/17
                  If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                     '¤@®×¦h¥ó²M³æ
                     'Modify By Sindy 2012/11/08 +m_CP28
                     strTemp = PUB_GetFCTAppendix(m_TM01, m_TM02, m_TM03, m_TM04, "501", m_CP27, "03", m_CP28, IIf(m_strCP09 <> "", m_strCP09, m_CP09), ET03)
                     'Modify By Sindy 2013/5/2 µ{¦¡²¾¨ìPUB_GetFCTAppendix
'                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                               "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
'                               "','¤@®×¦h¥ó²M³æ','" & ChgSQL(strTemp) & "')"
'                     cnnConnection.Execute strSql
                  End If
                  '2012/10/17 End
                  
                  If ET03 = "13" Then
                     'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                     If bolEmail = True And bolPlusPaper = False Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                                 "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the official notice and its translation for your reference. The originals will be sent to you via registered mail.')"
                        cnnConnection.Execute strSql
                     Else '¶l¥ó
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                                 "','¨Ò¥~¤º¤å','A copy of the official notice and its translation will be mailed to you with the confirmation copy of this letter for your records.')"
                        cnnConnection.Execute strSql
                     End If
                     '2012/11/27 End
                  ElseIf ET03 = "17" Then
                     'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                     If bolEmail = True And bolPlusPaper = False Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                                 "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the Notice of Approval and its translation for your reference. The originals will be sent to you via registered mail.')"
                        cnnConnection.Execute strSql
                     Else '¶l¥ó
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                                 "','¨Ò¥~¤º¤å','A copy of the Notice of Approval and its translation will be mailed to you with the confirmation copy of this letter for your records.')"
                        cnnConnection.Execute strSql
                     End If
                     '2012/11/27 End
                  End If
                  'Add By Sindy 2018/6/28
                  '¦³©µ®i¥¼®Ö­ã¤£¦L:©µ®i¤wµo¤å¥B¥¼®Ö­ã,¥BµL306¦Û½ÐºM¦^¤§¬ÛÃöÁ`¦¬¤å¸¹¬°¸Ó©µ®i.¨Ò:FCT-011076,FCT-026892
                  strSql = "select c2.cp09 from caseprogress c2" & _
                           " where c2.cp01='" & m_TM01 & "' and c2.cp02='" & m_TM02 & "' and c2.cp03='" & m_TM03 & "' and c2.cp04='" & m_TM04 & "'" & _
                           " and c2.cp10='102' and c2.cp27>0 and c2.cp159=0 and (c2.cp24 is null or c2.cp24='2')" & _
                           " and not exists(select c1.cp09 from caseprogress c1 where c1.cp01='" & m_TM01 & "' and c1.cp02='" & m_TM02 & "' and c1.cp03='" & m_TM03 & "' and c1.cp04='" & m_TM04 & "' and c1.cp10='306' and c1.cp43=c2.cp09)"
                  If m_rs.State = 1 Then m_rs.Close
                  m_rs.CursorLocation = adUseClient
                  m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If m_rs.RecordCount = 0 Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                              "','µL©µ®i¥¼®Ö­ã­n¦L','¡ð')"
                     cnnConnection.Execute strSql
                  End If
                  '2018/6/28 END
                  
                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
                  If textPrtTrans <> "N" Then
                     ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                     EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), ET03_1, strUserNum
                     ' ²¾Âà¤H
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                              "','²¾Âà¤H','" & GetCustomerEngName(m_CP55) & "')"
                     cnnConnection.Execute strSql
                     ' ²¾Âà¥Ó½Ð¤H
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                              "','²¾Âà¥Ó½Ð¤H','" & GetCustomerEngName(m_CP56) & "')"
                     cnnConnection.Execute strSql
                     'Add By Cheng 2003/03/13
                     ' ©ñ±ó±M¥ÎÅv
'                     If IsEmptyText(m_TM67) = False Then
                     If IsEmptyText(Trim(textTM67)) = False Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                 "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & ChgSQL(textTM67) & "')"
                        cnnConnection.Execute strSql
                     End If
                     ' ¥¿°Ó¼Ð¸¹¼Æ
                     If IsEmptyText(m_TM27) = False Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                 "','¥¿°Ó¼Ð¸¹¼Æ','" & "Its Principal " & IIf(m_TM08 >= "4" And m_TM08 <= "6", "Service Mark", "Trademark") & " No. : " & m_TM27 & "')"
                        cnnConnection.Execute strSql
                     End If
                     '92.2.18 ADD BY SONIA
                     ' ¨ä¥L¤½§i¤é
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                              "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                     cnnConnection.Execute strSql
                     'Add By Sindy 2012/10/12
                     If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                        strTemp = ""
                        CheckOC3
                        'Modify By Sindy 2012/11/08 +m_CP28
                        strSql = PUB_GetOneAppMuchCaseSql(m_TM01, m_TM02, m_TM03, m_TM04, "501", m_CP27, m_CP28)
                        AdoRecordSet3.CursorLocation = adUseClient
                        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                        If AdoRecordSet3.RecordCount <> 0 Then
                           AdoRecordSet3.MoveFirst
                           Do While Not AdoRecordSet3.EOF
                              strTemp = strTemp & "¡B" & "" & AdoRecordSet3.Fields("tm15").Value
                              AdoRecordSet3.MoveNext
                           Loop
                           If strTemp <> "" Then strTemp = Mid(strTemp, 2, Len(strTemp))
                        End If
                        CheckOC3
                        ' ©Ò¦³²¾Âà¤§µù¥U¸¹¼Æ
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                 "','©Ò¦³²¾Âà¤§µù¥U¸¹¼Æ','" & strTemp & "')"
                        cnnConnection.Execute strSql
                     End If
                     '2012/10/12 End
                  End If
'                End If
            '  ¤é¤å
            Case "3":
                ' ¬O§_¦C¦LÂ½Ä¶¨ç
                If textPrtTrans <> "N" Then
                  'Add By Sindy 2018/11/22
                  If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                     EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), ET03_1, strUserNum
                     '¤@¤å¦h®×²M³æšd
                     strTemp = PUB_GetFCTAppendix_JP(m_TM01, m_TM02, m_TM03, m_TM04, "501", m_CP27, "03", m_CP28, IIf(m_strCP09 <> "", m_strCP09, m_CP09), ET03_1, intCnt)
                     ' ¤@®×¦h¥ó¥ó¼Æ
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                              "','¤@®×¦h¥ó¥ó¼Æ','" & intCnt & "')"
                     cnnConnection.Execute strSql
                  Else
                  '2018/11/22 END
                     ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                     EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "15", strUserNum
                     ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                     EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), ET03_1, strUserNum
                  End If
                  ' ¨ä¥L¤½§i¤é
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                  "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                  cnnConnection.Execute strSql
                End If
         End Select
      ' ±ÂÅv
      Case "502":
         ' ©w½Z»y¤å
         Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
            ' ¤¤¤å
            Case "1":
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "13", strUserNum
               ' ¨÷¼Æ
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "13" & "','" & strUserNum & _
                        "','¨÷¼Æ','" & textTMBM07_1 & "')"
               cnnConnection.Execute strSql
               ' ´Á¼Æ
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "13" & "','" & strUserNum & _
                        "','´Á¼Æ','" & textTMBM07_2 & "')"
               cnnConnection.Execute strSql
               ' ¦C¦L³Æµù
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "13" & "','" & strUserNum & _
                        "','¦C¦L³Æµù','" & ChgSQL(textPS) & "')"
               cnnConnection.Execute strSql
               'Add By Cheng 2002/06/14
               ' ¨ä¥L¤½§i¤é
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "13" & "','" & strUserNum & _
                        "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
               cnnConnection.Execute strSql
            ' ­^¤å
            Case "2":
                'Modify By Sindy 2012/10/12 Mark¤w¤£°Ï¤À¤F
'                '­Yªþµù¥UÃÒ(ÂÂ©w½Z)
'                If m_strWithRegister <> "N" Then
'                    ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                    EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "14", strUserNum
'                    ' ±ÂÅv¤H
'                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "14" & "','" & strUserNum & _
'                             "','±ÂÅv¤H','" & GetCustomerEngName(m_TM23) & "')"
'                    cnnConnection.Execute strSql
'                    ' ³Q±ÂÅv¤H
'                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "14" & "','" & strUserNum & _
'                             "','³Q±ÂÅv¤H','" & GetCustomerEngName(m_CP50) & "')"
'                    cnnConnection.Execute strSql
'                    ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                    If textPrtTrans <> "N" Then
'                       ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                       EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "15", strUserNum
'                       ' ±ÂÅv¤H
'                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "15" & "','" & strUserNum & _
'                                "','±ÂÅv¤H','" & GetCustomerEngName(m_TM23) & "')"
'                       cnnConnection.Execute strSql
'                       ' ³Q±ÂÅv¤H
'                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "15" & "','" & strUserNum & _
'                                "','³Q±ÂÅv¤H','" & GetCustomerEngName(m_CP50) & "')"
'                       cnnConnection.Execute strSql
'                       'Add By Cheng 2003/03/13
'                       ' ©ñ±ó±M¥ÎÅv
''                       If IsEmptyText(m_TM67) = False Then
'                       If IsEmptyText(Trim(textTM67)) = False Then
'                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "15" & "','" & strUserNum & _
'                                   "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & ChgSQL(textTM67) & "')"
'                          cnnConnection.Execute strSql
'                       End If
'                       ' ¥¿°Ó¼Ð¸¹¼Æ
'                       If IsEmptyText(m_TM27) = False Then
'                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "15" & "','" & strUserNum & _
'                                   "','¥¿°Ó¼Ð¸¹¼Æ','" & "Its Principal " & IIf(m_TM08 >= "4" And m_TM08 <= "6", "Service Mark", "Trademark") & " No. : " & m_TM27 & "')"
'                          cnnConnection.Execute strSql
'                       End If
'                    End If
'                '­Y¤£ªþµù¥UÃÒ(·s©w½Z)
'                Else
                    ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                    EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), ET03, strUserNum
                    'Add By Cheng 2002/06/14
                    ' ¨ä¥L¤½§i¤é
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                             "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                    cnnConnection.Execute strSql
                    'End
                    '­Y¥Ó½Ð¤é¤p©óµ¥©ó930324
                    'Modify by Morgan 2004/5/27
                    '§ï§ìµo¤å¤é
                    'If Val(DBDATE(m_TM11)) <= 20040324 Then
                    If Val(m_CP27) <= 20040324 Then
                        ' ½Ðµ²²M½Ð´Ú³æ
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                                 "','½Ðµ²²M½Ð´Ú³æ','" & "Enclosed please find our final debit note for your kind settlement." & vbCrLf & "')"
                        cnnConnection.Execute strSql
                    End If
                    'Add By Sindy 2012/10/17
                    If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                        '¤@®×¦h¥ó²M³æ
                        'Modify By Sindy 2012/11/08 +m_CP28
                        strTemp = PUB_GetFCTAppendix(m_TM01, m_TM02, m_TM03, m_TM04, "502", m_CP27, "03", m_CP28, IIf(m_strCP09 <> "", m_strCP09, m_CP09), ET03)
                        'Modify By Sindy 2013/5/2 µ{¦¡²¾¨ìPUB_GetFCTAppendix
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                  "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
'                                  "','¤@®×¦h¥ó²M³æ','" & ChgSQL(strTemp) & "')"
'                        cnnConnection.Execute strSql
                    End If
                    '2012/10/17 End
                    
                    If ET03 = "16" Then
                        'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                        If bolEmail = True And bolPlusPaper = False Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                                    "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the official notice from the IPO and its translation for your reference. The originals will be sent to you via registered mail.')"
                           cnnConnection.Execute strSql
                        Else '¶l¥ó
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                                    "','¨Ò¥~¤º¤å','A copy of the official notice from the IPO and its translation will be mailed to you with the confirmation copy of this letter for your records.')"
                           cnnConnection.Execute strSql
                        End If
                        '2012/11/27 End
                    ElseIf ET03 = "18" Then
                        'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                        If bolEmail = True And bolPlusPaper = False Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                                    "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the Notice of Approval from the IPO and its translation for your reference. The originals will be sent to you via registered mail.')"
                           cnnConnection.Execute strSql
                        Else '¶l¥ó
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03 & "','" & strUserNum & _
                                    "','¨Ò¥~¤º¤å','A copy of the Notice of Approval from the IPO and its translation will be mailed to you with the confirmation copy of this letter for your records.')"
                           cnnConnection.Execute strSql
                        End If
                        '2012/11/27 End
                    End If
                    
                    ' ¬O§_¦C¦LÂ½Ä¶¨ç
                    If textPrtTrans <> "N" Then
                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                        EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), ET03_1, strUserNum
                        'Add By Cheng 2002/06/14
                        ' ¨ä¥L¤½§i¤é
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                 "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                        cnnConnection.Execute strSql
                        'End
                        'Add By Sindy 2012/10/12
                        If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                           intRow = 0: strTemp = ""
                           CheckOC3
                           'Modify By Sindy 2012/11/08 +m_CP28
                           strSql = PUB_GetOneAppMuchCaseSql(m_TM01, m_TM02, m_TM03, m_TM04, "502", m_CP27, m_CP28)
                           AdoRecordSet3.CursorLocation = adUseClient
                           AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                           If AdoRecordSet3.RecordCount <> 0 Then
                              AdoRecordSet3.MoveFirst
                              Do While Not AdoRecordSet3.EOF
                                 intRow = intRow + 1
                                 strTemp = strTemp & intRow & ") Reg. No. " & "" & AdoRecordSet3.Fields("tm15").Value & vbCrLf
                                 strTemp = strTemp & Mid("      ", 1, Len(intRow & ") ") - 1) & "Goods/Services:" & "|?TMGoods:" & AdoRecordSet3.Fields("cp01").Value & "-" & AdoRecordSet3.Fields("cp02").Value & "-" & AdoRecordSet3.Fields("cp03").Value & "-" & AdoRecordSet3.Fields("cp04").Value & "-­^¤å?|" & vbCrLf
                                 AdoRecordSet3.MoveNext
                              Loop
                           End If
                           CheckOC3
                           ' ©Ò¦³±ÂÅv¤§µù¥U¸¹¼Æ¤Î°Ó«~
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                    "','©Ò¦³±ÂÅv¤§µù¥U¸¹¼Æ¤Î°Ó«~','" & ChgSQL(strTemp) & "')"
                           cnnConnection.Execute strSql
                        End If
                        '2012/10/12 End
                    End If
'                End If
         End Select
      ' ÅÜ§ó 2007/6/7 ¥[´îÁY°Ó«~313
      Case "301", "313":
        'Modify By Cheng 2004/04/27
        '­YÅÜ§ó¨Æ¶µÀÉªº¥Ó½Ð¤H¬O§_®Ö­ã¥Bªþµù¥UÃÒ, ©Î¤£ªþµù¥UÃÒ, ©Î¤£ªþµù¥UÃÒ¥B´îÁY°Ó«~
'         If IsCE09Approve(IIf(m_strCP09 <> "", m_strCP09, m_CP09)) = True Or m_blnRestrictGoods = True Then
         If (IsCE09Approve(IIf(m_strCP09 <> "", m_strCP09, m_CP09)) = True And m_strWithRegister <> "N") Or m_strWithRegister = "N" Or (m_strWithRegister = "N" And m_blnRestrictGoods = True) Then
            ' ©w½Z»y¤å
            Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
               ' ¤¤¤å
               Case "1":
                  ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                  EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "16", strUserNum
                  ' ¨÷¼Æ
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "16" & "','" & strUserNum & _
                           "','¨÷¼Æ','" & textTMBM07_1 & "')"
                  cnnConnection.Execute strSql
                  ' ´Á¼Æ
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "16" & "','" & strUserNum & _
                           "','´Á¼Æ','" & textTMBM07_2 & "')"
                  cnnConnection.Execute strSql
                  ' ¦C¦L³Æµù
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "16" & "','" & strUserNum & _
                           "','¦C¦L³Æµù','" & ChgSQL(textPS) & "')"
                  cnnConnection.Execute strSql
                  'Add By Cheng 2002/06/14
                  ' ¨ä¥L¤½§i¤é
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "16" & "','" & strUserNum & _
                           "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                  cnnConnection.Execute strSql
               ' ­^¤å
               Case "2":
                    '­Yªþµù¥UÃÒ(ÂÂ©w½Z)
                    If m_strWithRegister <> "N" Then
                           ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                           EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "17", strUserNum
                           ' ÅÜ§ó«e¥Ó½Ð¤H
                             'Modify By Cheng 2003/07/14
         '                  strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         '                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & _
         '                           "','ÅÜ§ó«e¥Ó½Ð¤H','" & m_TM23 & "')"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "17" & "','" & strUserNum & _
                                    "','ÅÜ§ó«e¥Ó½Ð¤H','" & GetCustomerName(GetOldTM23(m_CP09), "1") & "')"
                           cnnConnection.Execute strSql
                           ' ÅÜ§ó«á¥Ó½Ð¤H
                             'Modify By Cheng 2003/07/14
         '                  strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         '                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & _
         '                           "','ÅÜ§ó«á¥Ó½Ð¤H','" & GetNewTM23() & "')"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "17" & "','" & strUserNum & _
                                    "','ÅÜ§ó«á¥Ó½Ð¤H','" & GetCustomerName(GetNewTM23()) & "')"
                           cnnConnection.Execute strSql
                           'Add By Cheng 2003/07/14
                           ' ¨ä¥L¤½§i¤é
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "17" & "','" & strUserNum & _
                                    "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                           cnnConnection.Execute strSql
                           'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                           If bolEmail = True And bolPlusPaper = False Then
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "17" & "','" & strUserNum & _
                                       "','¨Ò¥~¤º¤å','Enclosed you will find a scanned copy of the Notice of Approval as well as its English translation for your reference. The originals will be sent to you via registered mail.')"
                              cnnConnection.Execute strSql
                           Else '¶l¥ó
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "17" & "','" & strUserNum & _
                                       "','¨Ò¥~¤º¤å','Enclosed you will find the original Registration Certificate, on the reverse side of which the change has been endorsed with an official stamp, as well as its English translation for your reference. Our debit note is also enclosed for your kind settlement.')"
                              cnnConnection.Execute strSql
                           End If
                           '2012/11/27 End
                          ' ¬O§_¦C¦LÂ½Ä¶¨ç
                          If textPrtTrans <> "N" Then
                             ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                             EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "18", strUserNum
                             ' ÅÜ§ó«e¥Ó½Ð¤H
                            'Modify By Cheng 2003/07/14
        '                     strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        '                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & "18" & "','" & strUserNum & _
        '                              "','ÅÜ§ó«e¥Ó½Ð¤H','" & m_TM23 & "')"
                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                      "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "18" & "','" & strUserNum & _
                                      "','ÅÜ§ó«e¥Ó½Ð¤H','" & GetCustomerName(GetOldTM23(m_CP09), "1") & "')"
                             cnnConnection.Execute strSql
                             ' ÅÜ§ó«á¥Ó½Ð¤H
                            'Modify By Cheng 2003/07/14
        '                     strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        '                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & "18" & "','" & strUserNum & _
        '                              "','ÅÜ§ó«á¥Ó½Ð¤H','" & GetNewTM23() & "')"
                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                      "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "18" & "','" & strUserNum & _
                                      "','ÅÜ§ó«á¥Ó½Ð¤H','" & GetCustomerName(GetNewTM23(), "1") & "')"
                             cnnConnection.Execute strSql
                          'Add By Cheng 2003/03/13
                          ' ©ñ±ó±M¥ÎÅv
'                          If IsEmptyText(m_TM67) = False Then
                          If IsEmptyText(Trim(textTM67)) = False Then
                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                      "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "18" & "','" & strUserNum & _
                                      "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & ChgSQL(textTM67) & "')"
                             cnnConnection.Execute strSql
                          End If
                          ' ¥¿°Ó¼Ð¸¹¼Æ
                          If IsEmptyText(m_TM27) = False Then
                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                      "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "18" & "','" & strUserNum & _
                                      "','¥¿°Ó¼Ð¸¹¼Æ','" & "Its Principal " & IIf(m_TM08 >= "4" And m_TM08 <= "6", "Service Mark", "Trademark") & " No. : " & m_TM27 & "')"
                             cnnConnection.Execute strSql
                          End If
                             '92.2.18 ADD BY SONIA
                             ' ¨ä¥L¤½§i¤é
                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                      "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "18" & "','" & strUserNum & _
                                      "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                             cnnConnection.Execute strSql
                             '92.2.18 END
                          End If
                    '­Y¤£ªþµù¥UÃÒ(·s©w½Z)
                    Else
                        '­YµL´îÁY°Ó«~
                        If m_blnRestrictGoods = False Then
                           'Add By Sindy 2012/11/14
                           If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                              ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                              EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "23", strUserNum
                              ' ¨ä¥L¤½§i¤é
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "23" & "','" & strUserNum & _
                                       "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                              cnnConnection.Execute strSql
                              'ÅÜ§ó¨Æ¶µ¤º®e
                              strTemp = ""
                              '¥Ó½Ð¤HÅÜ§ó
                              strChgEvent = GetCustEngName(ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE09", "CE04"))
                              If strChgEvent <> "" Then strTemp = "name"
                              '¥Ó½Ð¦a§}ÅÜ§ó
                              strChgEvent = ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE38", "'changed'")
                              If strChgEvent <> "" Then
                                 If strTemp = "name" Then
                                    strTemp = "name and address"
                                 Else
                                    strTemp = "address"
                                 End If
                              End If
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "23" & "','" & strUserNum & _
                                       "','ÅÜ§ó¨Æ¶µ¤º®e','" & ChgSQL(strTemp) & "')"
                              cnnConnection.Execute strSql
                              '¤@®×¦h¥ó²M³æ
                              strTemp = PUB_GetFCTAppendix(m_TM01, m_TM02, m_TM03, m_TM04, "301", m_CP27, "03", m_CP28, IIf(m_strCP09 <> "", m_strCP09, m_CP09), "23")
                              'Modify By Sindy 2013/5/2 µ{¦¡²¾¨ìPUB_GetFCTAppendix
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "23" & "','" & strUserNum & _
'                                       "','¤@®×¦h¥ó²M³æ','" & ChgSQL(strTemp) & "')"
'                              cnnConnection.Execute strSql
                              If Val(m_CP27) <= 20040324 Then
                                  ' ½Ðµ²²M½Ð´Ú³æ
                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                           "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "23" & "','" & strUserNum & _
                                           "','½Ðµ²²M½Ð´Ú³æ','" & "Enclosed please find our final debit note for your kind settlement." & vbCrLf & "')"
                                  cnnConnection.Execute strSql
                              End If
                              'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                              If bolEmail = True And bolPlusPaper = False Then
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "23" & "','" & strUserNum & _
                                          "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the Notice of Approval from the IPO and its translation for your reference. The originals will be sent to you via registered mail.')"
                                 cnnConnection.Execute strSql
                              Else '¶l¥ó
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "23" & "','" & strUserNum & _
                                          "','¨Ò¥~¤º¤å','A copy of the Notice of Approval from the IPO and its translation will be mailed to you with the confirmation copy of this letter for your records.')"
                                 cnnConnection.Execute strSql
                              End If
                              '2012/11/27 End
                              ' ¬O§_¦C¦LÂ½Ä¶¨ç
                              If textPrtTrans <> "N" Then
                                 ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                                 EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "24", strUserNum
                                 ' ¨ä¥L¤½§i¤é
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "24" & "','" & strUserNum & _
                                          "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                                 cnnConnection.Execute strSql
                                 '¥Ó½Ð¤HÅÜ§ó
                                 strChgEvent = GetCustEngName(ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE09", "CE04"))
                                 If strChgEvent <> "" Then
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "24" & "','" & strUserNum & _
                                             "','¥Ó½Ð¤HÅÜ§ó','" & ChgSQL("Registrant's Name: " & strChgEvent) & "')"
                                    cnnConnection.Execute strSql
                                 End If
                                 '¥Ó½Ð¤H¤¤Ä¶¤åÅÜ§ó
                                 strChgEvent = ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE22", "CE17")
                                 If strChgEvent <> "" Then
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "24" & "','" & strUserNum & _
                                             "','¥Ó½Ð¤H¤¤Ä¶¤åÅÜ§ó','" & ChgSQL("Chinese characters of Registrant's Name: " & strChgEvent) & "')"
                                    cnnConnection.Execute strSql
                                 End If
                                 '¥Ó½Ð¦a§}ÅÜ§ó
                                 strChgEvent = ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE38", "'changed'")
                                 If strChgEvent <> "" Then
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "24" & "','" & strUserNum & _
                                             "','¥Ó½Ð¦a§}ÅÜ§ó','" & ChgSQL("Registrant's Address: " & strChgEvent) & "')"
                                    cnnConnection.Execute strSql
                                 End If
                                 '¥Nªí¤HÅÜ§ó
                                 strChgEvent = ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE16", "'changed'")
                                 If strChgEvent <> "" Then
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "24" & "','" & strUserNum & _
                                             "','¥Nªí¤HÅÜ§ó','" & ChgSQL("Name of Registrant's Representative: " & strChgEvent) & "')"
                                    cnnConnection.Execute strSql
                                 End If
                                 '¥N²z¤HÅÜ§ó
                                 strChgEvent = ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE56", "'changed'")
                                 m_MySt(1) = m_TM01: m_MySt(2) = m_TM02: m_MySt(3) = m_TM03: m_MySt(4) = m_TM04 'Add By Sindy 2014/4/23
                                 If strChgEvent <> "" Then
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "24" & "','" & strUserNum & _
                                             "','¥N²z¤HÅÜ§ó','Attorneys'' names: " & ExceptFieldData("¥X¦W¥N²z¤H/­^") & "')"
                                    cnnConnection.Execute strSql
                                 End If
                                 'Modify By Sindy 2012/11/15
                                 strTemp = ""
                                 CheckOC3
                                 strSql = PUB_GetOneAppMuchCaseSql(m_TM01, m_TM02, m_TM03, m_TM04, "301", m_CP27, m_CP28)
                                 AdoRecordSet3.CursorLocation = adUseClient
                                 AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                                 If AdoRecordSet3.RecordCount <> 0 Then
                                    AdoRecordSet3.MoveFirst
                                    Do While Not AdoRecordSet3.EOF
                                       strTemp = strTemp & "¡B" & "" & AdoRecordSet3.Fields("tm15").Value
                                       AdoRecordSet3.MoveNext
                                    Loop
                                    If strTemp <> "" Then strTemp = Mid(strTemp, 2, Len(strTemp))
                                 End If
                                 CheckOC3
                                 '©Ò¦³ÅÜ§ó¤§µù¥U¸¹¼Æ
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                          "','©Ò¦³ÅÜ§ó¤§µù¥U¸¹¼Æ','" & strTemp & "')"
                                 cnnConnection.Execute strSql
                                 '2012/11/15 End
                              End If
                           Else
                           '2012/11/14 End
                              ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                              EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "19", strUserNum
                              ' ¨ä¥L¤½§i¤é
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "19" & "','" & strUserNum & _
                                       "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                              cnnConnection.Execute strSql
                              'Add By Sindy 2013/6/5
                              'ÅÜ§ó¨Æ¶µ¤º®e
                              strTemp = ""
                              '¥Ó½Ð¤HÅÜ§ó
                              strChgEvent = GetCustEngName(ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE09", "CE04"))
                              If strChgEvent <> "" Then strTemp = "name"
                              '¥Ó½Ð¦a§}ÅÜ§ó
                              strChgEvent = ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE38", "'changed'")
                              If strChgEvent <> "" Then
                                 If strTemp = "name" Then
                                    strTemp = "name and address"
                                 Else
                                    strTemp = "address"
                                 End If
                              End If
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "19" & "','" & strUserNum & _
                                       "','ÅÜ§ó¨Æ¶µ¤º®e','" & ChgSQL(strTemp) & "')"
                              cnnConnection.Execute strSql
                              '2013/6/5 END
                              'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                              If bolEmail = True And bolPlusPaper = False Then
                                 'Modify By Sindy 2013/7/4
                                 'Old:Enclosed you will find the scanned copy of the Notice of Approval as well as its English translation for your reference. The originals will be sent to you via registered mail. Our debit note is also enclosed for your kind settlement.
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "19" & "','" & strUserNum & _
                                          "','¨Ò¥~¤º¤å','Enclosed you will find the scanned copy of the Notice of Approval as well as its English translation for your reference. The originals will be sent to you via registered mail.')"
                                 cnnConnection.Execute strSql
                              Else '¶l¥ó
                                 'Modify By Sindy 2013/3/29
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "19" & "','" & strUserNum & _
                                          "','¨Ò¥~¤º¤å','A copy of the Notice of Approval from the IPO and its translation will be mailed to you with the confirmation copy of this letter for your records.')"
                                 cnnConnection.Execute strSql
                              End If
                              '2012/11/27 End
                              '­Y¥Ó½Ð¤é¤p©óµ¥©ó930324
                              'Modify by Morgan 2004/5/27
                              '§ï§ìµo¤å¤é
                              'If Val(DBDATE(m_TM11)) <= 20040324 Then
                              If Val(m_CP27) <= 20040324 Then
                                  ' ½Ðµ²²M½Ð´Ú³æ
                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                           "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "19" & "','" & strUserNum & _
                                           "','½Ðµ²²M½Ð´Ú³æ','" & "Enclosed please find our final debit note for your kind settlement." & vbCrLf & "')"
                                  cnnConnection.Execute strSql
                              End If
                              ' ¬O§_¦C¦LÂ½Ä¶¨ç
                              If textPrtTrans <> "N" Then
                                 ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                                 EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "20", strUserNum
                                 ' ¨ä¥L¤½§i¤é
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                          "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                                 cnnConnection.Execute strSql
                                 '¥Ó½Ð¤HÅÜ§ó
                                 strChgEvent = GetCustEngName(ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE09", "CE04"))
                                 If strChgEvent <> "" Then
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                             "','¥Ó½Ð¤HÅÜ§ó','" & ChgSQL("Registrant's Name: " & strChgEvent) & "')"
                                    cnnConnection.Execute strSql
                                 End If
                                 '¥Ó½Ð¤H¤¤Ä¶¤åÅÜ§ó
                                 strChgEvent = ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE22", "CE17")
                                 If strChgEvent <> "" Then
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                             "','¥Ó½Ð¤H¤¤Ä¶¤åÅÜ§ó','" & ChgSQL("Chinese characters of Registrant's Name: " & strChgEvent) & "')"
                                    cnnConnection.Execute strSql
                                 End If
                                 '¥Ó½Ð¦a§}ÅÜ§ó
                                 strChgEvent = ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE38", "'changed'")
                                 If strChgEvent <> "" Then
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                             "','¥Ó½Ð¦a§}ÅÜ§ó','" & ChgSQL("Registrant's Address: " & strChgEvent) & "')"
                                    cnnConnection.Execute strSql
                                 End If
                                 '¥Nªí¤HÅÜ§ó
                                 strChgEvent = ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE16", "'changed'")
                                 If strChgEvent <> "" Then
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                             "','¥Nªí¤HÅÜ§ó','" & ChgSQL("Name of Registrant's Representative: " & strChgEvent) & "')"
                                    cnnConnection.Execute strSql
                                 End If
                                 '2009/4/17 ADD BY SONIA
                                 '¥N²z¤HÅÜ§ó
                                 strChgEvent = ChkChangeEvent(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE56", "'changed'")
                                 m_MySt(1) = m_TM01: m_MySt(2) = m_TM02: m_MySt(3) = m_TM03: m_MySt(4) = m_TM04 'Add By Sindy 2014/4/23
                                 If strChgEvent <> "" Then
                                    'Modify By Sindy 2010/6/1
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
'                                             "','¥N²z¤HÅÜ§ó','Attorneys'' names: Henry Chi-heng Guei, Fred C.T. Yen')"
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "20" & "','" & strUserNum & _
                                             "','¥N²z¤HÅÜ§ó','Attorneys'' names: " & ExceptFieldData("¥X¦W¥N²z¤H/­^") & "')"
                                    cnnConnection.Execute strSql
                                 End If
                                 '2009/4/17 END
                              End If
                           End If
                        '­Y¦³´îÁY°Ó«~
                        Else
                           ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                           EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "21", strUserNum
                           ' ¨ä¥L¤½§i¤é
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "21" & "','" & strUserNum & _
                                    "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                           cnnConnection.Execute strSql
                           
                           'Modify By Sindy 2022/6/13 Mark
'                           'Add By Sindy 2012/11/27 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
'                           If bolEmail = True And bolPlusPaper = False Then
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "21" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å','Enclosed herewith please find a scanned copy of the official notice from the IPO and its translation for your reference. The originals will be sent to you via registered mail.')"
'                              cnnConnection.Execute strSql
'                           Else '¶l¥ó
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "21" & "','" & strUserNum & _
'                                       "','¨Ò¥~¤º¤å','A copy of the official notice from the IPO and its translation will be mailed to you with the confirmation copy of this letter for your records.')"
'                              cnnConnection.Execute strSql
'                           End If
'                           '2012/11/27 End
                           
                           '­Y¥Ó½Ð¤é¤p©óµ¥©ó930324
                           'Modify by Morgan 2004/5/27
                           '§ï§ìµo¤å¤é
                           'If Val(DBDATE(m_TM11)) <= 20040324 Then
                           If Val(m_CP27) <= 20040324 Then
                               ' ½Ðµ²²M½Ð´Ú³æ
                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "21" & "','" & strUserNum & _
                                        "','½Ðµ²²M½Ð´Ú³æ','" & "Enclosed please find our final debit note for your kind settlement." & vbCrLf & "')"
                               cnnConnection.Execute strSql
                           End If
                           ' ¬O§_¦C¦LÂ½Ä¶¨ç
                           If textPrtTrans <> "N" Then
                               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                               EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "22", strUserNum
                               ' ¨ä¥L¤½§i¤é
                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & "22" & "','" & strUserNum & _
                                        "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                               cnnConnection.Execute strSql
                           End If
                        End If
                    End If
               ' ¤é¤å
               Case "3":
                  If Trim(textTM15.Text) = "" Then 'µù¥U«eÅÜ§ó
                  Else 'µù¥U«áÅÜ§ó
                     'ÀË¬dÅÜ§ó¨Æ¶µ
                     strSql = "select * from changeevent where ce01='" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "' "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        If "" & RsTemp.Fields("ce04") <> "" Or _
                           "" & RsTemp.Fields("ce05") <> "" Or _
                           "" & RsTemp.Fields("ce06") <> "" Or _
                           "" & RsTemp.Fields("ce07") <> "" Or _
                           "" & RsTemp.Fields("ce08") <> "" Then
                           strTemp09 = "Y"
                        End If
                        If "" & RsTemp.Fields("ce23") <> "" Or _
                           "" & RsTemp.Fields("ce24") <> "" Or _
                           "" & RsTemp.Fields("ce25") <> "" Or _
                           "" & RsTemp.Fields("ce26") <> "" Or _
                           "" & RsTemp.Fields("ce27") <> "" Or _
                           "" & RsTemp.Fields("ce28") <> "" Or _
                           "" & RsTemp.Fields("ce29") <> "" Or _
                           "" & RsTemp.Fields("ce30") <> "" Or _
                           "" & RsTemp.Fields("ce31") <> "" Or _
                           "" & RsTemp.Fields("ce32") <> "" Or _
                           "" & RsTemp.Fields("ce33") <> "" Or _
                           "" & RsTemp.Fields("ce34") <> "" Or _
                           "" & RsTemp.Fields("ce35") <> "" Or _
                           "" & RsTemp.Fields("ce36") <> "" Or _
                           "" & RsTemp.Fields("ce37") <> "" Then
                           strTemp38 = "Y"
                        End If
                        If strTemp09 = "Y" And strTemp38 = "Y" Then
                           'Modified by Morgan 2023/3/15
                           'strTemp = "¡]°Ó¼Ð“¸ªÌÇU¦W†ï¤ÎÇZ¦í©ÒŒi§ó¡^"
                           'Modified by Morgan 2024/4/2
                           'strTemp = PUB_GetUniText(Me.Name, "ÅÜ§ó1")
                           strTemp = PUB_GetUniText(Me.Name, "ÅÜ§ó°Ó¼Ð¥Ó½Ð¤H¤Î¦a§}")
                        ElseIf strTemp09 = "Y" Then
                           'Modified by Morgan 2023/3/15
                           'strTemp = "¡]°Ó¼Ð“¸ªÌÇU¦W†ïŒi§ó¡^"
                           'Modified by Morgan 2024/4/2
                           'strTemp = PUB_GetUniText(Me.Name, "ÅÜ§ó2")
                           strTemp = PUB_GetUniText(Me.Name, "ÅÜ§ó°Ó¼Ð¥Ó½Ð¤H")
                        ElseIf strTemp38 = "Y" Then
                           'Modified by Morgan 2023/3/15
                           'strTemp = "¡]°Ó¼Ð“¸ªÌÇU¦í©ÒŒi§ó¡^"
                           'Modified by Morgan 2024/4/2
                           'strTemp = PUB_GetUniText(Me.Name, "ÅÜ§ó3")
                           strTemp = PUB_GetUniText(Me.Name, "ÅÜ§ó°Ó¼Ð¥Ó½Ð¦a§}")
                        End If
                     End If
                     
                     ' ¬O§_¦C¦LÂ½Ä¶¨ç
                     If textPrtTrans <> "N" Then
                        EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), ET03_1, strUserNum
                        'Add By Sindy 2018/11/22
                        If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                           '¤@¤å¦h®×²M³æšd
                           strTemp = PUB_GetFCTAppendix_JP(m_TM01, m_TM02, m_TM03, m_TM04, "301", m_CP27, "03", m_CP28, IIf(m_strCP09 <> "", m_strCP09, m_CP09), ET03_1, intCnt)
                           ' ¤@®×¦h¥ó¥ó¼Æ
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                    "','¤@®×¦h¥ó¥ó¼Æ','" & intCnt & "')"
                           cnnConnection.Execute strSql
                        Else
                           EndLetter "03", IIf(m_strCP09 <> "", m_strCP09, m_CP09), "27", strUserNum
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('03','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','27','" & strUserNum & _
                                    "','ÅÜ§ó¨Æ¶µ','" & strTemp & "')"
                           cnnConnection.Execute strSql
                        End If
                        
                        'Add By Sindy 2018/12/3 Åª¨úÅÜ§óÀÉ
                        StrSQLa = "Select ce01,ce09,ce38,ce16,ce56 From changeevent Where ce01='" & m_CP09 & "'"
                        rsA.CursorLocation = adUseClient
                        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                        If rsA.Fields(0).Value > 0 Then
                           If "" & rsA.Fields("ce09") = "1" Then
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                       "','ÅÜ§ó¥Ó½Ð¤H¦WºÙ','¡ð')"
                              cnnConnection.Execute strSql
                           End If
                           If "" & rsA.Fields("ce38") = "1" Then
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                       "','ÅÜ§ó¥Ó½Ð¤H¦í©Ò','¡ð')"
                              cnnConnection.Execute strSql
                           End If
                           If "" & rsA.Fields("ce16") = "1" Then
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                       "','ÅÜ§ó¥Nªí¤H','¡ð')"
                              cnnConnection.Execute strSql
                           End If
                           If "" & rsA.Fields("ce56") = "1" Then
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                       "','ÅÜ§ó¥X¦W¥N²z¤H','¡ð')"
                              cnnConnection.Execute strSql
                           End If
                        End If
                        If rsA.State <> adStateClosed Then rsA.Close
                        Set rsA = Nothing
                        '2018/12/3 END
                        ' ¨ä¥L¤½§i¤é
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "','" & ET03_1 & "','" & strUserNum & _
                                 "','¨ä¥L¤½§i¤é','" & DBDATE(Me.textTM14.Text) & "')"
                        cnnConnection.Execute strSql
                        '2018/11/22 END
                     End If
                  End If
            End Select
         End If
         
      'Add By Cheng 2003/09/05
      '§ó¥¿
      Case "302":
        '­Y¬OÃÒ®Ñ§ó§ï
        If Me.textMod.Text <> "" Then
            m_strCP09 = m_TM01 & m_TM02 & m_TM03 & m_TM04 & "&1701"
             ' ©w½Z»y¤å
             Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
                ' ¤¤¤å
                Case "1":
                   '2005/8/26 MODIFY BY SONIA
                   ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                   'EndLetter "05", m_strCP09, "01", strUserNum
                   EndLetter "05", m_strCP09, "21", strUserNum
                   '2005/8/26 END
                
                ' ­^¤å
                Case "2":
                  'Modify By Sindy 2022/8/25
                  If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, "1701", strET03, , "05") = True Then
                     EndLetter "05", m_strCP09, strET03, strUserNum
                  Else
                  '2022/8/25 END
                     'Add by Sindy 2020/4/24 ¬O§_°±¤î¶l°È
                     If m_NA86 = "Y" Then
                        strET03 = "23"
                        EndLetter "05", m_strCP09, strET03, strUserNum
                     Else
                     '2020/4/24 END
                       'edit by nick 2004/09/24
   '                    If Query716717_cp Then
'                           'Modify By Sindy 2012/6/27 °Ó¼Ð­×ªk
'                           If Val(strSrvDate(1)) >= 20120701 Then
                              strET03 = "22"
                              EndLetter "05", m_strCP09, strET03, strUserNum
'                              Else
'                              '2012/6/27 End
'                                 strET03 = "19"
'                                 EndLetter "05", m_strCP09, strET03, strUserNum
'                           End If
   '                    Else
   '                        EndLetter "05", m_strCP09, "18", strUserNum
   '                    End If
                     End If
                  End If
                  'Add By Sindy 2015/6/23
                  If m_TM08 = "7" Then 'ÃÒ©ú¼Ð³¹
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "05" & "','" & m_strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','°Ó¼ÐºØÃþ','Certification Mark')"
                     cnnConnection.Execute strSql
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "05" & "','" & m_strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','Class','')"
                     cnnConnection.Execute strSql
                  Else
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "05" & "','" & m_strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','°Ó¼ÐºØÃþ','Trademark')"
                     cnnConnection.Execute strSql
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "05" & "','" & m_strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','Class','Class(es) : " & textTM09 & "')"
                     cnnConnection.Execute strSql
                  End If
                  '2015/6/23 ENd
                 'edit by nick 2004/10/07
                 If textPrtTrans <> "N" Then
                    ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                    EndLetter "05", m_strCP09, "13", strUserNum
                     'Add By Sindy 2015/6/23
                     If m_TM08 = "7" Then 'ÃÒ©ú¼Ð³¹
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                 "','°Ó¼ÐºØÃþ','CERTIFICATION MARK')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                 "','Class','')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                 "','ªA°È¶µ¥Ø','Contents of Certification : ')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                 "','Trademark','')"
                        cnnConnection.Execute strSql
                     Else
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                 "','°Ó¼ÐºØÃþ','TRADEMARK')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                 "','Class','Class(es) : " & textTM09 & "')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                 "','ªA°È¶µ¥Ø','Specification of Goods/Services :')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                 "','Trademark','Trademark ')"
                        cnnConnection.Execute strSql
                     End If
                     '2015/6/23 END
                     '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
                     If Me.Text1.Text <> "" Then
                         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                  "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
                         cnnConnection.Execute strSql
                     End If
                     '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
                     If m_TM67 <> "" Then
                         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                  "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed¡G" & ChgSQL(m_TM67) & "')"
                         cnnConnection.Execute strSql
                     End If
                     '¨Ò¥~Äæ¦ì--ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù
                     '                           If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹") > 0 Then
                     If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹") > 0 Or InStr(m_TM58, "­ì¬°Áp¦XªA°È¼Ð³¹") > 0 Then
                         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                  "','ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù','(Service Mark of prior Trademark Law)')"
                         cnnConnection.Execute strSql
                     End If
                     'add by nickc 2007/03/08 ¥[¤J¦P·N®Ñ°Ó¼Ð¸¹¼Æ
                     If m_TM118 <> "" Then
                         'Modify By Sindy 2012/11/06 23-I-13=>30-I-10
                         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                                  "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & vbCrLf & "*In accordance with the proviso of Article 30-I-10 of the Trademark Law, this mark is granted registration with consent from the proprietor(s) of Reg. No(s). " & ChgSQL(m_TM118) & "') "
                         cnnConnection.Execute strSql
                     End If
                  End If
                    
                ' ¤é¤å
                Case "3":
                    '­Y±M¥Î´Á°_¤é¤p©ó921201(¥ÎÂÂ©w½Z)
                    'edit by nickc 2005/06/28 §ï¦¨¸òÃÒ®Ñ³W«h¬Û¦P
'                        '­Y¥Ó½Ð¤é¤p©ó921128(¥ÎÂÂ©w½Z)
'                        If Val(DBDATE(m_TM11)) < 20031128 Then
'                            ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                            EndLetter "05", m_strCP09, "14", strUserNum
''Removed by Morgan 2023/3/15 ©w½Z¨S¥Î¨ì
''                            ' Áp¦X°Ó¼Ð
''                            If IsEmptyText(m_TM27) = False Then
''                               ' Áp¦X°Ó¼Ð
''                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                        "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "14" & "','" & strUserNum & _
''                                        "','Áp¦X°Ó¼Ð','" & "¨Ì¦s ¥¿°Ó¼Ð µn¿ýµf¸¹ : (" & m_TM27 & ")" & "')"
''                               cnnConnection.Execute strSql
''                            End If
''                            ' ¬O§_¦C¦LÂ½Ä¶¨ç
''                            If textPrtTrans <> "N" Then
''                               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
''                               EndLetter "05", m_strCP09, "15", strUserNum
''                               ' Áp¦X°Ó¼Ð
''                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                        "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "15" & "','" & strUserNum & _
''                                        "','Áp¦X°Ó¼Ð','" & "¨Ì¦s ¥¿°Ó¼Ð µn¿ýµf¸¹ : (" & m_TM27 & ")" & "')"
''                               cnnConnection.Execute strSql
''                               ' °Ó«~°Ï¤À
''                               If m_TM08 = "4" Then
''                                  ' °Ó«~°Ï¤À
''                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "15" & "','" & strUserNum & _
''                                           "','°Ó«~°Ï¤À','" & "ªA°È°Ï¤À" & "')"
''                                  cnnConnection.Execute strSql
''                               Else
''                                  ' °Ó«~°Ï¤À
''                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "15" & "','" & strUserNum & _
''                                           "','°Ó«~°Ï¤À','" & "°Ó«~°Ï¤À" & "')"
''                                  cnnConnection.Execute strSql
''                               End If
''                               ' «ü©w°Ó«~
''                               If m_TM08 = "4" Then
''                                  ' «ü©w°Ó«~
''                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "15" & "','" & strUserNum & _
''                                           "','«ü©w°Ó«~','" & "«ü©w§Ð°È" & "')"
''                                  cnnConnection.Execute strSql
''                               Else
''                                  ' «ü©w°Ó«~
''                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "15" & "','" & strUserNum & _
''                                           "','«ü©w°Ó«~','" & "«ü©w°Ó«~" & "')"
''                                  cnnConnection.Execute strSql
''                               End If
''                            End If
''end 2023/3/15
'                        '­Y¥Ó½Ð¤é¤j©óµ¥©ó921128(¥Î·s©w½Z)
'                        Else
                            ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                            If Is716Have = False Then
                                EndLetter "05", m_strCP09, "17", strUserNum
                            Else
                                EndLetter "05", m_strCP09, "16", strUserNum
                            End If
                            ' Áp¦X°Ó¼Ð
                            If IsEmptyText(m_TM27) = False Then
                               ' Áp¦X°Ó¼Ð
                               'Removed by Morgan 2023/3/15 ©w½Z¨S¥Î¨ì
                               'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               '         "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "16" & "','" & strUserNum & _
                               '         "','Áp¦X°Ó¼Ð','" & "¨Ì¦s ¥¿°Ó¼Ð µn¿ýµf¸¹ : (" & m_TM27 & ")" & "')"
                               'cnnConnection.Execute strSql
                            End If
                            ' ¬O§_¦C¦LÂ½Ä¶¨ç
                            If textPrtTrans <> "N" Then
                               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                               'edit by nick 2004/08/17 ¦]¬°¸­©ö¶³»¡­×ªk«e«áªºÄ¶¤å¬Ò¬Û¦P
                               'EndLetter "05", strCP09, "17", strUserNum
                               EndLetter "05", m_strCP09, "15", strUserNum
                                'Add By Cheng 2003/02/19
                                '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
'                                If m_TM67 <> "" Then
                                If Trim(textTM67) <> "" Then
                                    'edit by nick 2004/08/17 ¦]¬°¸­©ö¶³»¡­×ªk«e«áªºÄ¶¤å¬Ò¬Û¦P
                                    'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "17" & "','" & strUserNum & _
                                             "','©ñ±ó±M¥ÎÅv','°Ó¼Ð¨£¥»ÇRÆèÇr¡u" & ChgSQL(m_TM67) & "¡vÇUˆü¥e“¸Çy¦³þêÇQÆê¡C')"
                                    'Modify By Sindy 2022/10/12 ˆü¥e“¸Çy¦³ §ï¬° °Ó¼Ð“¸Çy¥D±i
                                    'Modified by Morgan 2023/3/15
                                    'strExc(1) = "°Ó¼Ð¨£¥»ÇRÆèÇr¡u" & ChgSQL(textTM67) & "¡vÇU°Ó¼Ð“¸Çy¥D±iþêÇQÆê¡C"
                                    strExc(1) = PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv1") & ChgSQL(textTM67) & PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv2")
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "15" & "','" & strUserNum & _
                                             "','©ñ±ó±M¥ÎÅv','" & strExc(1) & "')"
                                    cnnConnection.Execute strSql
                                End If
                                'Add By Sindy 2010/11/17
                                If m_TM118 <> "" Then
                                    'Modified by Morgan 2023/3/15
                                    'strExc(1) = "°Ó¼Ðªk²Ä30’f²Ä1¶µ²Ä10†AÇU³W©wÇR°òþøþà¡Bµn“÷°Ó¼Ð²Ä" & ChgSQL(m_TM118) & "†AÇU°Ó¼Ð“¸ªÌÇU¦P·NÇRÇoÇqµn“÷Çy³\¥iþìÇr¡C"
                                    strExc(1) = PUB_GetUniText(Me.Name, "¦P·N®Ñ°Ó¼Ð¸¹¼Æ1") & ChgSQL(m_TM118) & PUB_GetUniText(Me.Name, "¦P·N®Ñ°Ó¼Ð¸¹¼Æ2")
                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                             "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "15" & "','" & strUserNum & _
                                             "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & strExc(1) & "')"
                                    cnnConnection.Execute strSql
                                End If
                                '2010/11/17 End
                            End If
'                        End If
             End Select
        End If
        
      'Add By Sindy 2014/9/9
      Case "103": '¸Éµoµù¥UÃÒ
         m_strCP09 = m_TM01 & m_TM02 & m_TM03 & m_TM04 & "&1701"
         ' ©w½Z»y¤å
         Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
            ' ­^¤å
            Case "2":
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "05", m_strCP09, "13", strUserNum
               'Add By Sindy 2015/8/3
               If m_TM08 = "7" Then 'ÃÒ©ú¼Ð³¹
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                           "','°Ó¼ÐºØÃþ','CERTIFICATION MARK')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                           "','Class','')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                           "','ªA°È¶µ¥Ø','Contents of Certification : ')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                           "','Trademark','')"
                  cnnConnection.Execute strSql
               Else
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                           "','°Ó¼ÐºØÃþ','TRADEMARK')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                           "','Class','Class(es) : " & textTM09 & "')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                           "','ªA°È¶µ¥Ø','Specification of Goods/Services :')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                           "','Trademark','Trademark ')"
                  cnnConnection.Execute strSql
               End If
               '2015/8/3 END
               '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
               If Me.Text1.Text <> "" Then
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                            "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
                   cnnConnection.Execute strSql
               End If
               '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
               If Trim(textTM67) <> "" Then
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                            "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed¡G" & ChgSQL(textTM67) & "')"
                   cnnConnection.Execute strSql
               End If
               '¨Ò¥~Äæ¦ì--ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù
               If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹") > 0 Or InStr(m_TM58, "­ì¬°Áp¦XªA°È¼Ð³¹") > 0 Then
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                            "','ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù','(Service Mark of prior Trademark Law)')"
                   cnnConnection.Execute strSql
               End If
               '¨Ò¥~Äæ¦ì--¦P·N®Ñ°Ó¼Ð¸¹¼Æ
               If m_TM118 <> "" Then
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "13" & "','" & strUserNum & _
                           "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & vbCrLf & "*In accordance with the proviso of Article 30-I-10 of the Trademark Law, this mark is granted registration with consent from the proprietor(s) of Reg. No(s). " & ChgSQL(m_TM118) & ".') "
                  cnnConnection.Execute strSql
               End If
            ' ¤é¤å
            Case "3":
               EndLetter "05", m_strCP09, "24", strUserNum 'Add By Sindy 2020/12/17 ¸Éµoµù¥UÃÒ©w½Z
               EndLetter "05", m_strCP09, "15", strUserNum
               If Trim(textTM67) <> "" Then
                  'Modify By Sindy 2022/10/12 ˆü¥e“¸Çy¦³ §ï¬° °Ó¼Ð“¸Çy¥D±i
                  'Modified by Morgan 2023/3/15
                  'strExc(1) = "°Ó¼Ð¨£¥»ÇRÆèÇr¡u" & ChgSQL(textTM67) & "¡vÇU°Ó¼Ð“¸Çy¥D±iþêÇQÆê¡C"
                  strExc(1) = PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv1") & ChgSQL(textTM67) & PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv2")
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "15" & "','" & strUserNum & _
                           "','©ñ±ó±M¥ÎÅv','" & strExc(1) & "')"
                  cnnConnection.Execute strSql
               End If
               '¨Ò¥~Äæ¦ì--¦P·N®Ñ°Ó¼Ð¸¹¼Æ
               If m_TM118 <> "" Then
                  'Modified by Morgan 2023/3/15
                  'strExc(1) = "°Ó¼Ðªk²Ä30’f²Ä1¶µ²Ä10†AÇU³W©wÇR°òþøþà¡Bµn“÷°Ó¼Ð²Ä" & ChgSQL(m_TM118) & "†AÇU°Ó¼Ð“¸ªÌÇU¦P·NÇRÇoÇqµn“÷Çy³\¥iþìÇr¡C"
                  strExc(1) = PUB_GetUniText(Me.Name, "¦P·N®Ñ°Ó¼Ð¸¹¼Æ1") & ChgSQL(m_TM118) & PUB_GetUniText(Me.Name, "¦P·N®Ñ°Ó¼Ð¸¹¼Æ2")
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "05" & "','" & m_strCP09 & "','" & "15" & "','" & strUserNum & _
                           "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & strExc(1) & "')"
                  cnnConnection.Execute strSql
               End If
         End Select
      '2014/9/9 END
   End Select
End Sub

Private Sub PrintLetter()
'add by nickc 2005/06/28
Dim rsA As New ADODB.Recordset
'Add by Morgan 2008/6/12
Dim stCP10 As String, stContent As String
'Added by Lydia 2023/03/08
Dim stLang As String '©w½Z»y¤å
Dim m_strCP10 As String 'Added by Lydia 2023/05/03 ¨Ó¨ç©Ê½è

On Error GoTo ErrHnd
   
'   'Add By Sindy 2010/5/13 ÀË¬d¬O§_¦³ÅÜ§ó¥Ó½Ð¤H
'   bChkChaEvent = False
'   strSql = "SELECT * FROM ChangeEvent WHERE CE01='" & IIf(m_strCP09 <> "", m_strCP09, m_CP09) & "' "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      If Trim("" & RsTemp.Fields("CE04")) <> "" Then bChkChaEvent = True
'   End If
'   '2010/5/13 End
   
   'Add By Sindy 2016/12/6 ÀË¬d¬O§_¦³ÅÜ§ó¨Æ¶µ
   'ÅÜ§ó¥Ó½Ð¤H:m_strCE04
   'ÅÜ§ó¦a§}:m_strCE23CE24CE25
   If PUB_FCTchkChangeEventData(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE04", m_strCE04) = True Then
      Call PUB_FCTchkChangeEventData(IIf(m_strCP09 <> "", m_strCP09, m_CP09), "CE23||CE24||CE25", m_strCE23CE24CE25)
   End If
   '2016/12/6 END
   
   bolToFile = True 'Added by Lydia 2023/06/05
   
   stCP10 = IIf(strCP10Code <> "", strCP10Code, m_CP10)
   Select Case stCP10
      Case "301", "313" 'ÅÜ§ó 2007/6/7 ¥[´îÁY°Ó«~313
         m_blnRestrictGoods = RestrictGoods(IIf(m_strCP09 <> "", m_strCP09, m_CP09))
   End Select
   
   'Modify By Sindy 2012/2/1 ¦]¦¹¬qµ{¦¡¦³«Ü¦h¦a¤è³£»Ý­n§PÂ_¨ì,©Ò¥H´£¨ì³Ì«e­±¤@¦¸ÀË¬d
   bolChaEventNewCase = False
   '2011/9/7 ADD BY SONIA ¥Ñ©µ®i±µ¶i¨Ó¤§·s®×¥X¤£¦P©w½Z(­^Ä¶¤å¤£ÅÜ)
   'Modify By Sindy 2015/3/18 ¥[¤À³Î®×
   '¡@¡@"AND C1.cp01=C2.cp01(+) and C1.cp02=C2.cp02(+) and C1.cp03=C2.cp03(+) and C1.cp04=C2.cp04(+) and '101'=C2.cp10(+) " ==>
   '¡@¡@"AND C1.cp01=C2.cp01(+) and C1.cp02=C2.cp02(+) and C1.cp03=C2.cp03(+) and C1.cp04=C2.cp04(+) and instr('101,308',C2.cp10)>0 "
   'Modify By Sindy 2015/3/27 ex.FCT-27670,FCT-27672
'   StrSQLa = "Select C1.cp05,C2.cp09,C3.cp09,C3.cp05 From CaseProgress C1,CaseProgress C2,CaseProgress C3 Where C1.cp09='" & m_CP09 & "' " & _
'             "AND C1.cp01=C2.cp01(+) and C1.cp02=C2.cp02(+) and C1.cp03=C2.cp03(+) and C1.cp04=C2.cp04(+) and instr('101,308',C2.cp10)>0 " & _
'             "AND C1.cp01=C3.cp01(+) and C1.cp02=C3.cp02(+) and C1.cp03=C3.cp03(+) and C1.cp04=C3.cp04(+) and '102'=C3.cp10(+) " & _
'             "order by c3.cp05"
   '¦³101.¥Ó½Ð308.¤À³Î¥B¬°AÃþ¦¬¤å=¥¿±`·s¥Ó½Ð®×,¤Ï¤§«h¬°¤¤¶¡±µ¶i¨Ó
   'modify by sonia 2016/10/20 ¤£§PÂ_AÃþ¦¬¤å§ï§PÂ_CP05<>19221111(FCT-039304-T¥Ó½Ð¬°74/12/6¤§BÃþ)
   'StrSQLa = "select cp09 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "'" & _
             " and cp10 in(101,308) and substr(cp09,1,1)='A'"
   StrSQLa = "select cp09 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "'" & _
             " and cp10 in(101,308) and cp05<>19221111"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   'If rsA.RecordCount > 0 Then
   If rsA.RecordCount = 0 Then
      'If "" & rsA.Fields(1) = "" And Val("" & rsA.Fields(3)) = Val("" & rsA.Fields(0)) Then
         bolChaEventNewCase = True '¤¤¶¡±µ¶i¨Ó
      'End If
   End If
   '2012/2/1 End
   
   'Add By Sindy 2012/11/23 ±q¤U­±µ{¦¡©¹¤WMove¦Ü¦¹
   bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, stCP10 = "102", , bolPlusPaper) 'ÀË¬d¬O§_¥HE-Mail³qª¾
   '2012/11/23 End
    
   ET01 = "03"
   ET02 = m_CP09
   stLang = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) 'Added by Lydia 2023/03/08 §ï¦¨ÅÜ¼Æ
   
   ' ®×¥ó©Ê½è
   Select Case stCP10
      'Modify By Cheng 2003/12/16
      '¥Ó½Ð®Ö­ãªº©w½Z§ï¦b¦¹³B¥X, ­ì¦bFC¤½§i³qª¾¨ç¥X
      ' ¥Ó½Ð
      'edit by nick 2004/12/23 ¤À³Î»P¥Ó½Ð°µ¬Û¦Pªº¨Æ±¡
      'Case "101":
      Case "101", "308":
         ' ©w½Z»y¤å
         'Modified by Lydia 2023/03/08 §ï¦¨ÅÜ¼Æ
         Select Case stLang
            ' ¤¤¤å
            Case "1":
               ET03 = "01"
            ' ­^¤å
            Case "2":
'2014/12/9 CANCEL BY SONIA
'                '­Y¥Ó½Ð¤é¤p©ó921128
'                If Val(m_TM11) < 20031128 Then
'                   ET03 = "99"
'                '­Y¥Ó½Ð¤é¤j©óµ¥©ó921128
'                Else
                   '2008/7/24 ADD BY SONIA FCTµù¥U¶O¦Û°Ê¥NÃº
                   If m_TM122 = "Y" Then
                     'Modify By Sindy 2010/01/05
                     If Trim(m_TM67) = "" And Trim(textTM67) <> "" Then
                        ET03 = "14"
                     Else
                        'Modify By Sindy 2024/8/2
                        If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, "101", ET03, , "03") = True Then
                        Else
                        '2024/8/2 END
                           ET03 = "10"
                        End If
                     End If
                   Else
                   '2008/7/24 END
                        '93.6.23 ADD BY SONIA
                        'Modify By Sindy 2010/01/05
                        'Remove by Lydia 2018/03/22 ¨ú®ø"©ñ±ó±M¥ÎÅv"©w½Z
                        'Modified by by Lydia 2018/03/28  ¤À³Î¥ý¤£§ï + And stCP10 = "308"
                        If Trim(m_TM67) = "" And Trim(textTM67) <> "" And stCP10 = "308" Then
                           ET03 = "13"
                        Else
                           '2014/12/9 MODIFY BY SONIA
                           'Modify By Sindy 2012/6/27 °Ó¼Ð­×ªk
                           'If Val(DBDATE(m_CP05)) >= 20120701 Then
                           '   ET03 = "17"
                           'Else
                           ''2012/6/27 End
                           '   ET03 = "06"
                           'End If
                           ET03 = "17"
                           '2014/12/9 END
                        End If 'Remove by Lydia  2018/03/22
                        'add by nickc 2007/02/16 ¥[¦h¥Ó½Ð¤H®É¡A¤W­z©w½Z¶W¹L 4000 ¡A©Ò¥H©î¦¨ 2 ­Ó
                        '¦^ÂÐ³æ
                        'Modify By Sindy 2012/6/27 °Ó¼Ð­×ªk
                        '2014/12/9 MODIFY BY SONIA
                        'If Val(DBDATE(m_CP05)) >= 20120701 Then
                        '   ET03r = "18"
                        'Else
                        ''2012/6/27 End
                        '   ET03r = "09"
                        'End If
                        'Modified by Lydia 2018/03/28 ¤£¦L¦^ÂÐ³æ (¥ý±Æ°£¤À³Î)
                        'ET03r = "18"
                        If stCP10 = "308" Then ET03r = "18"
                        '2014/12/9 END
                        '93.6.23 END
                   End If
'                End If  '2014/12/9 CANCEL BY SONIA
            
            ' ¤é¤å
            Case "3":
               'Add By Sindy 2019/7/22 ªü½¬»¡¤À³Î¨S¤é¤å©w½Z,¤£­n¥X©w½Z ex:FCT-43164
               If stCP10 = "101" Then
               '2019/7/22 END
'2014/12/9 CANCEL BY SONIA
'                '­Y¥Ó½Ð¤é¤p©ó921128
'                If Val(DBDATE(m_TM11)) < 20031128 Then
'                    ET03 = "05"
'                'edit by nick 2004/08/03 ¥[¤J¤é¤å©w½Z
'                Else
                   '2008/7/24 ADD BY SONIA FCTµù¥U¶O¦Û°Ê¥NÃº
                   If m_TM122 = "Y" Then
                     ET03 = "11"
                   Else
                   '2008/7/24 END
                     'Modify By Sindy 2012/6/26 °Ó¼Ð­×ªk
                     '2014/12/9 MODIFY BY SONIA
                     'If Val(DBDATE(m_CP05)) >= 20120701 Then
                     '   ET03 = "15"
                     '   '¦^ÂÐ³æ
                     '   ET03r = "16"
                     'Else
                     ''2012/6/26 End
                     '   ET03 = "07"
                     '   'add by nick 2004/10/15 ©î¦¨2 ±i¡A¦]¬°¤£¦P¯È±i
                     '   '¦^ÂÐ³æ
                     '   ET03r = "08"
                     'End If
                     ET03 = "15"
                     
                     'Modify By Sindy 2021/6/28 ¨ó§U­×§ïFCT¤é¤å²Õ¤§¡u¥Ó½Ð(°Ó¥Ó)¡v®Ö­ã©w½Z¡G§R°£¡uFAXªð«H¥Î¯È¡v
'                     '¦^ÂÐ³æ
'                      ET03r = "16"
                      '2021/6/28 END
                      '2014/12/9 END
                   End If
                   '2008/11/13 add by sonia ±NÄ¶¤å¿W¥ß¥X¨Ó¦L
                   ET03_1 = "12"
                   '2008/11/13 end
'                End If    '2014/12/9 CANCEL BY SONIA
               End If
         End Select
      ' ©µ®i
      Case "102":
         ET02 = IIf(m_strCP09 <> "", m_strCP09, m_CP09)
         ' ©w½Z»y¤å
         'Modified by Lydia 2023/03/08 §ï¦¨ÅÜ¼Æ
         Select Case stLang
            ' ­^¤å
            Case "2":
               'Modify By Sindy 2010/5/13 ¦³ÅÜ§ó¥Ó½Ð¤H
               'If bChkChaEvent = True Then
               If m_strCE04 <> "" Then
                  'Modify By Sindy 2012/2/1 ¥Ñ©µ®i±µ¶i¨Ó¤§·s®×¥X¤£¦P©w½Z(­^Ä¶¤å¤£ÅÜ)
                  If bolChaEventNewCase = True Then
                     ET03 = "17"
                  Else
                  '2012/2/1 End
                     ET03 = "14"
                  End If
                  ET03_1 = "15"
               '2010/5/13 End
               Else
                  '­Yªþµù¥UÃÒ(ÂÂ©w½Z)
'                  If m_strWithRegister <> "N" Then
'                     ET03 = "06"
'                     ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                     If textPrtTrans <> "N" Then
'                        If m_TM08 = "2" Then
'                           ET03_1 = "08"
'                        ElseIf m_TM08 = "5" Then
'                           ET03_1 = "09"
'                        Else
'                           ET03_1 = "07"
'                        End If
'                      End If
'                  '­Y¤£ªþµù¥UÃÒ(·s©w½Z)
'                  Else
                     '2011/9/7 ADD BY SONIA ¥Ñ©µ®i±µ¶i¨Ó¤§·s®×¥X¤£¦P©w½Z(­^Ä¶¤å¤£ÅÜ)
'                     StrSQLa = "Select C1.cp05,C2.cp09,C3.cp09,C3.cp05 From CaseProgress C1,CaseProgress C2,CaseProgress C3 Where C1.cp09='" & m_CP09 & "' " & _
'                               "AND C1.cp01=C2.cp01(+) and C1.cp02=C2.cp02(+) and C1.cp03=C2.cp03(+) and C1.cp04=C2.cp04(+) and '101'=C2.cp10(+) " & _
'                               "AND C1.cp01=C3.cp01(+) and C1.cp02=C3.cp02(+) and C1.cp03=C3.cp03(+) and C1.cp04=C3.cp04(+) and '102'=C3.cp10(+) " & _
'                               "order by c3.cp05"
'                     rsA.CursorLocation = adUseClient
'                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'                     If rsA.RecordCount > 0 Then
'                        If "" & rsA.Fields(1) = "" And Val(rsA.Fields(3)) = Val(rsA.Fields(0)) Then
                     'Modify By Sindy 2012/2/1
                     If bolChaEventNewCase = True Then
                        ET03 = "16"
                     Else
                     '2012/2/1 End
                        'Modify By Sindy 2012/6/27 °Ó¼Ð­×ªk
                        If Val(strSrvDate(1)) >= 20120701 Then
                           ET03 = "18"
                        Else
                        '2012/6/27 End
                           ET03 = "10"
                        End If
                     End If
'                        End If
'                     End If
'                     '2011/9/7 END
                     ' ¬O§_¦C¦LÂ½Ä¶¨ç
                     If textPrtTrans <> "N" Then
                        ET03_1 = "11"
                     End If
'                  End If
               End If
            ' ¤é¤å
            Case "3":
'2009/8/24 CANCEL BY SONIA ¤é¤åµLÂÂ©w½Z
'                '­Yªþµù¥UÃÒ(ÂÂ©w½Z)
'                If m_strWithRegister <> "N" Then
'                  ET03 = "08"
'                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                  If textPrtTrans <> "N" Then
'                     ET03_1 = "09"
'                  End If
'                '­Y¤£ªþµù¥UÃÒ(·s©w½Z)
'                Else
'2009/8/24 END
               'Add By Sindy 2016/12/16 ÀË¬d¬O§_¦³ÅÜ§ó¨Æ¶µ
               If m_strCE04 <> "" Or m_strCE23CE24CE25 <> "" Then
                  ET03 = "19"
                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
                  If textPrtTrans <> "N" Then
                     ET03_1 = "20"
                     'Added by Lydia 2023/09/04 ¥t¥~²£¥Í©w½Z
                     If txtADate.Visible = True And txtADate <> "" Then
                        ET03_ex = "28"
                     End If
                     'end 2023/09/04
                  End If
               Else
               '2016/12/16 END
                  ET03 = "12"
                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
                  If textPrtTrans <> "N" Then
                     ET03_1 = "13"
                     'Added by Lydia 2023/09/04 ¥t¥~²£¥Í©w½Z
                     If txtADate.Visible = True And txtADate <> "" Then
                        ET03_ex = "28"
                     End If
                     'end 2023/09/04
                  End If
               End If
         End Select
      ' ²¾Âà
      Case "501":
         ET02 = IIf(m_strCP09 <> "", m_strCP09, m_CP09)
         ' ©w½Z»y¤å
         'Modified by Lydia 2023/03/08 §ï¦¨ÅÜ¼Æ
         Select Case stLang
            ' ¤¤¤å
            Case "1":
               ET03 = "10"
            ' ­^¤å
            Case "2":
                'Modify By Sindy 2012/10/12 Mark¤w¤£°Ï¤À¤F
'                '­Yªþµù¥UÃÒ(ÂÂ©w½Z)
'                If m_strWithRegister <> "N" Then
'                  ET03 = "11"
'                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                  If textPrtTrans <> "N" Then
'                     ET03_1 = "12"
'                  End If
'                '­Y¤£ªþµù¥UÃÒ(·s©w½Z)
'                Else
                  'Add By Sindy 2012/10/12
                  If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                     ET03 = "17"
                     ' ¬O§_¦C¦LÂ½Ä¶¨ç
                     If textPrtTrans <> "N" Then
                        ET03_1 = "18"
                     End If
                  Else
                  '2012/10/12 End
                     ET03 = "13"
                     ' ¬O§_¦C¦LÂ½Ä¶¨ç
                     If textPrtTrans <> "N" Then
                        ET03_1 = "14"
                     End If
                  End If
'                End If
            ' ¤é¤å
            Case "3":
                ' ¬O§_¦C¦LÂ½Ä¶¨ç
                If textPrtTrans <> "N" Then
                  'Add By Sindy 2018/11/22
                  If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                     ET03_1 = "19"
                  Else
                  '2018/11/22 END
                     ET03 = "15"
                     ET03_1 = "16"
                  End If
                End If
         End Select
      ' ±ÂÅv
      Case "502":
         ET02 = IIf(m_strCP09 <> "", m_strCP09, m_CP09)
         ' ©w½Z»y¤å
         'Modified by Lydia 2023/03/08 §ï¦¨ÅÜ¼Æ
         Select Case stLang
            ' ¤¤¤å
            Case "1":
               ET03 = "13"
            ' ­^¤å
            Case "2":
                'Modify By Sindy 2012/10/12 Mark¤w¤£°Ï¤À¤F
'                '­Yªþµù¥UÃÒ(ÂÂ©w½Z)
'                If m_strWithRegister <> "N" Then
'                  ET03 = "14"
'                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                  If textPrtTrans <> "N" Then
'                     ET03_1 = "15"
'                  End If
'                '­Y¤£ªþµù¥UÃÒ(·s©w½Z)
'                Else
                  'Add By Sindy 2012/10/12
                  If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                     ET03 = "18"
                     If textPrtTrans <> "N" Then
                        ET03_1 = "19"
                     End If
                  Else
                  '2012/10/12 End
                     ET03 = "16"
                     If textPrtTrans <> "N" Then
                        ET03_1 = "17"
                     End If
                  End If
'                End If
         End Select
      ' ÅÜ§ó 2007/6/7 ¥[´îÁY°Ó«~313
      Case "301", "313":
         ET02 = IIf(m_strCP09 <> "", m_strCP09, m_CP09)
        '­YÅÜ§ó¨Æ¶µÀÉªº¥Ó½Ð¤H¬O§_®Ö­ã¥Bªþµù¥UÃÒ, ©Î¤£ªþµù¥UÃÒ, ©Î¤£ªþµù¥UÃÒ¥B´îÁY°Ó«~
         If (IsCE09Approve(ET02) = True And m_strWithRegister <> "N") Or m_strWithRegister = "N" Or (m_strWithRegister = "N" Or m_blnRestrictGoods = True) Then
            ' ©w½Z»y¤å
            'Modified by Lydia 2023/03/08 §ï¦¨ÅÜ¼Æ
            Select Case stLang
               ' ¤¤¤å
               Case "1":
                  ET03 = "16"
               ' ­^¤å
               Case "2":
                  '­Yªþµù¥UÃÒ(ÂÂ©w½Z)
                  If m_strWithRegister <> "N" Then
                      ET03 = "17"
                      ' ¬O§_¦C¦LÂ½Ä¶¨ç
                      If textPrtTrans <> "N" Then
                         ET03_1 = "18"
                      End If
                  '­Y¤£ªþµù¥UÃÒ(·s©w½Z)
                  Else
                      '­YµL´îÁY°Ó«~
                      If m_blnRestrictGoods = False Then
                         'Add By Sindy 2012/11/14
                         If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                            ET03 = "23"
                            ' ¬O§_¦C¦LÂ½Ä¶¨ç
                            If textPrtTrans <> "N" Then
                               ET03_1 = "24"
                            End If
                         Else
                         '2012/11/14 End
                            ET03 = "19"
                            ' ¬O§_¦C¦LÂ½Ä¶¨ç
                            If textPrtTrans <> "N" Then
                               ET03_1 = "20"
                            End If
                         End If
                      '­Y¦³´îÁY°Ó«~
                      Else
                          ET03 = "21"
                          ' ¬O§_¦C¦LÂ½Ä¶¨ç
                          If textPrtTrans <> "N" Then
                              ET03_1 = "22"
                          End If
                      End If
                  End If
               ' ¤é¤å
               Case "3":
                  If Trim(textTM15.Text) = "" Then 'µù¥U«eÅÜ§ó
                  Else 'µù¥U«áÅÜ§ó
                     ET03 = ""
                     ' ¬O§_¦C¦LÂ½Ä¶¨ç
                     If textPrtTrans <> "N" Then
                        'Add By Sindy 2018/11/22
                        If m_CP148 = "Y" Then '¤@¥Ó½Ð®Ñ¦h¥ó
                           ET03_1 = "25"
                        Else
                        '2018/11/22 END
                        'Add By Sindy 2019/3/27
                           ET03 = "27"
                           ET03_1 = "26"
                        '2019/3/27 END
                        End If
                     End If
                  End If
            End Select
         End If
        'Add By Cheng 2003/09/05
      ' §ó¥¿
      Case "302":
        '­Y¬OÃÒ®Ñ§ó§ï
        If Me.textMod.Text <> "" Then
            m_strCP09 = m_TM01 & m_TM02 & m_TM03 & m_TM04 & "&1701"
            ET02 = m_strCP09
            ET01 = "05"
            bolToFile = True 'Added by Lydia 2023/03/08 ±N©w½Z¡BÂ½Ä¶¨ç©MÃÒ®Ñ¦s¤JFCT_WorkFlow
             ' ©w½Z»y¤å
             'Modified by Lydia 2023/03/08 §ï¦¨ÅÜ¼Æ
             Select Case stLang
                ' ¤¤¤å
                Case "1":
                    ET03 = "21"
                ' ­^¤å
                Case "2":
                     'Modify By Sindy 2022/8/25
                     If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, "1701", ET03, , "05") = True Then
                     Else
                     '2022/8/25 END
                        'Add by Sindy 2020/4/24 ¬O§_°±¤î¶l°È
                        If m_NA86 = "Y" Then
                           ET03 = "23"
                        Else
                        '2020/4/24 END
                           ET03 = "22"
                        End If
                     End If
                     '³]©w­n¦C¦L¦a§}±ø
                     m_blnPrintAddress = True
                     ' ¬O§_¦C¦LÂ½Ä¶¨ç
                     If textPrtTrans <> "N" Then
                        ET03_1 = "13"
                     End If
                     '2022/8/25 END
                     
               ' ¤é¤å
                Case "3":
                    'edit by nickc 2005/06/28 §ï³W«h¸òÃÒ®Ñ¦P
                        '­Y¥Ó½Ð¤é¤p©ó921128(¥ÎÂÂ©w½Z)
                        If Val(DBDATE(m_TM11)) < 20031128 Then
                            ET03 = "14"
                            ' ¬O§_¦C¦LÂ½Ä¶¨ç
                            If textPrtTrans <> "N" Then
                                ET03_1 = "15"
                            End If
                        '­Y¥Ó½Ð¤é¤j©óµ¥©ó921128(¥Î·s©w½Z)
                        Else
                            'add by nickc 2005/06/28
                            Is716Have = True
                              'If (DBDATE(textTM21) >= 20031128) Or (DBDATE(textTM14) <= 20030901 And DBDATE(textTM21) < 20031128 And Trim(textTM14) <> "") Then
                              If (Val(DBDATE(textTM21)) >= Val(20031128)) Or (Val(DBDATE(textTM14)) <= Val(20030901) And Val(DBDATE(textTM21)) < Val(20031128) And Trim(textTM14) <> "") Then
                                   'add by nick 2004/08/17
                                   '¥ýÀË¬d¬O§_¦³ 717
                                    Set rsA = New ADODB.Recordset 'Add By Sindy 2012/3/2
                                    StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='717' and cp05 is not null and cp57 is null "
                                    rsA.CursorLocation = adUseClient
                                    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                                    If rsA.RecordCount > 0 Then
                                    Else
                                       Set rsA = New ADODB.Recordset
                                       StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='716' "
                                       rsA.CursorLocation = adUseClient
                                       rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                                       '­Y¦³¦¬¤å²Ä¤G´Áµù¥U¶O
                                       If rsA.RecordCount > 0 Then
                                       '­Y¥¼¦¬¤å²Ä¤G´Áµù¥U¶O
                                       Else
                                               Is716Have = False
                                       End If
                                   End If
                                   If rsA.State <> adStateClosed Then rsA.Close
                                   Set rsA = Nothing
                              End If
                            ' ¦C¦L©w½Z
                            'edit by nick 2004/08/17
                            If Is716Have = False Then
                                ET03 = "17"
                            Else
                                ET03 = "16"
                            End If
                            ' ¬O§_¦C¦LÂ½Ä¶¨ç
                            If textPrtTrans <> "N" Then
                                ET03_1 = "15"
                            End If
                        End If
             End Select
        End If
        
      'Add By Sindy 2014/9/9
      Case "103": '¸Éµoµù¥UÃÒ
         m_strCP09 = m_TM01 & m_TM02 & m_TM03 & m_TM04 & "&1701"
         ET02 = m_strCP09
         ET01 = "05"
         bolToFile = True 'Added by Lydia 2023/03/08 ±N©w½Z¡BÂ½Ä¶¨ç©MÃÒ®Ñ¦s¤JFCT_WorkFlow
         ' ©w½Z»y¤å
         'Modified by Lydia 2023/03/08 §ï¦¨ÅÜ¼Æ
         Select Case stLang
            ' ­^¤å
            Case "2":
               ET03 = ""
               'Â½Ä¶¨ç
               ET03_1 = "13"
            ' ¤é¤å
            Case "3":
               ET03 = "24" 'Add By Sindy 2020/12/17 ¸Éµoµù¥UÃÒ©w½Z
               'Â½Ä¶¨ç
               ET03_1 = "15"
         End Select
   End Select
   'Modify By Sindy 2012/10/12 ­ì¦b¤WÀYµ{¦¡¬q¸Ì,²¾¦Ü¦¹³B
   ' ¥ý©I¥s©w½Zµ{¦¡ªº²M°£­ì©w½Z¸ê®Æªº¨ç¦¡¥h²M°£¤§«e´Ý¯d¦b¨Ò¥~Äæ¦ìÀÉ¤¤ªº¸ê®Æ
   InsExpField
   '2012/10/12 End
   
   'Modify By Sindy 2014/9/9
   'If ET03 <> "" Then
   If ET03 <> "" Or ET03_1 <> "" Or ET03r <> "" Then
   '2014/9/9 END
      'Added by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW¡G¤£¥ÎºÞ´¼¼z§½¨Ó¨ç¡u®×¥Ñ¡v¡A°w¹ï©Ò¦³®Ö­ã¨Ã¥B¦³©w½Z´N©ñFCT_WORKFLOW
      If frm03020401_03.GetSelectResult = "1" Then
         m_strCP10 = "1001"
         bolToFile = True
      ElseIf frm03020401_03.GetSelectResult = "2" Then
         m_strCP10 = "1403"
      End If
      'end 2023/05/03
      'Added by Lydia 2023/03/08
      If bolToFile = True Then
         'Modified by Lydia 2023/05/03 §ï¦¨¦@¥Î¼Ò²Õ
         strFilePath = Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04)
         'Modified by Lydia 2023/09/04 ¥t¥~²£¥Í©w½Z
         'If Pub_GetFCTeFileName(strFilePath, m_TM01, m_TM02, m_TM03, m_TM04, m_strCP10, m_CP10, strFN01, strFN02, strFN03) = False Then
         strExc(1) = m_CP10 & IIf(stLang = "3" And stCP10 = "102" And txtADate <> "", stCP10, "")
         If Pub_GetFCTeFileName(strFilePath, m_TM01, m_TM02, m_TM03, m_TM04, m_strCP10, strExc(1), strFN01, strFN02, strFN03, strFN04, strFN05) = False Then
         'end 2023/09/04
            Exit Sub
         End If
         'end 2023/05/03
      End If
      'end 2023/03/08
      'Add by Morgan 2008/6/12
'      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, stCP10 = "102", , bolPlusPaper)
      If bolEmail Then
         'Add by Morgan 2009/10/20 +§PÂ_¬O§_EMail¦P®É±H¯È¥»
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'end 2009/10/20
         m_blnPrintAddress = False
         
         If ET03_1 <> "" Then
            '²£¥Í¯È¥»
            'Modified by Lydia 2023/03/08 False, strUserNum =>§ï§PÂ_ IIf(bolToFile = True, True, False)
            NowPrint ET02, ET01, ET03, IIf(bolToFile = True, True, False), strUserNum, , , , , iCopy
            'Added by Lydia 2023/03/08 'Memo by Lydia 2024/11/14 ³qª¾¨ç(*.LTR,*.®ÑÂ²)
            If bolToFile = True Then '©w½Z
                'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                'Call WordToFile(strFilePath & "\" & strFN01)
                If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                    Sleep 100
                End If
                'end 2023/05/03
            End If
            'end 2023/03/08
            
            '2008/11/13 modify by sonia
            'NowPrint ET02, ET01, ET03, False, strUserNum, , , True, stContent, , , , True
            'NowPrint ET02, ET01, ET03_1, False, strUserNum, , stContent, , , , , True, True
            If ET03r = "" Then
               '²£¥Í¹q¤lÀÉ
               NowPrint ET02, ET01, ET03, False, strUserNum, , , True, stContent, , , , True
               'Modified by Lydia 2019/04/01 FCT-38643¦]¬°¬Oµù¥U«áÅÜ§ó¨S¦³®Ö­ã©w½Z¡A©Ò¥H¹w³]¤£¦s©w½Z
               'NowPrint ET02, ET01, ET03_1, False, strUserNum, , stContent, , , , , True, True
               NowPrint ET02, ET01, ET03_1, False, strUserNum, , stContent, , , , , True, True, False 'Memo by Lydia 2024/11/14 ­ì¥»ÀÉ®×¦WºÙ:®×¸¹_¤é´Á(³qª¾¨ç+Ä¶¤å)=ET03+ET03_1
            Else
               'Modify by Morgan 2011/9/27 ¦^ÂÐ³æ¥u­n¦L1¥÷(°Ñ¦Ò¤U­±«D¹q¤lÀÉµ{¦¡)
               NowPrint ET02, ET01, ET03r, False, strUserNum, , , , , 1
               '²£¥Í¹q¤lÀÉ
               'Modified by Morgan 2020/7/24 ¶¶§Ç¦³»~,¹q¤lÀÉ¯ÊÄ¶¤å
               'NowPrint ET02, ET01, ET03, False, strUserNum, , , True, stContent, , , , True
               'NowPrint ET02, ET01, ET03r, False, strUserNum, , stContent, , , , , True, True
               ''Modified by Lydia 2019/04/01 FCT-38643¦]¬°¬Oµù¥U«áÅÜ§ó¨S¦³®Ö­ã©w½Z¡A©Ò¥H¹w³]¤£¦s©w½Z
               ''NowPrint ET02, ET01, ET03_1, False, strUserNum, , stContent, True, stContent, , , , True
               'NowPrint ET02, ET01, ET03_1, False, strUserNum, , stContent, True, stContent, , , , True, False
               NowPrint ET02, ET01, ET03, False, strUserNum, , , True, stContent
               NowPrint ET02, ET01, ET03r, False, strUserNum, , stContent, True, stContent
               NowPrint ET02, ET01, ET03_1, False, strUserNum, , stContent, , , , , True, True
               'end 2020/7/24
            End If
            '2008/11/13 end
            
            '²£¥Í¯È¥»
            'Modified by Lydia 2023/03/08 False, strUserNum =>§ï§PÂ_ IIf(bolToFile = True, True, False)
            NowPrint ET02, ET01, ET03_1, IIf(bolToFile = True, True, False), strUserNum, , , , , iCopy
            'Added by Lydia 2023/03/08 'Memo by Lydia 2024/11/14 Ä¶¤å(*.TRANS,*.Ä¶¤å)
            If bolToFile = True Then 'Â½Ä¶
               'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                If PUB_PrintWord2File(g_WordAp, strFilePath, strFN02) = True Then
                    Sleep 100
                End If
                'end 2023/05/03
            End If
            'end 2023/03/08
            'Added by Lydia 2023/09/04 ¥t¥~²£¥Í©w½Z
            If ET03_ex <> "" Then
               NowPrint m_CP09, ET01, ET03_ex, IIf(bolToFile = True, True, False), strUserNum, , , , , iCopy
                If bolToFile = True Then
                   If PUB_PrintWord2File(g_WordAp, strFilePath, strFN03) = True Then
                       Sleep 100
                   End If
               End If
            End If
         ElseIf ET03r <> "" Then
            'Modified by Lydia 2023/03/08 False, strUserNum =>§ï§PÂ_ IIf(bolToFile = True, True, False)
            NowPrint ET02, ET01, ET03, IIf(bolToFile = True, True, False), strUserNum, , , , , iCopy
            'Added by Lydia 2023/03/08
            If bolToFile = True Then '©w½Z
                'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                    Sleep 100
                End If
                'end 2023/05/03
            End If
            'end 2023/03/08
            'Modify by Morgan 2011/9/27 ¦^ÂÐ³æ¥u­n¦L1¥÷(°Ñ¦Ò¤U­±«D¹q¤lÀÉµ{¦¡)
            NowPrint ET02, ET01, ET03r, False, strUserNum, , , , , 1
            '²£¥Í¹q¤lÀÉ
            NowPrint ET02, ET01, ET03, False, strUserNum, , , True, stContent, , , , True
            NowPrint ET02, ET01, ET03r, False, strUserNum, , stContent, , , , , True, True
            'end 2023/03/08
         Else
            'Add By Sindy 2018/11/22 µù¥UÅÜ§ó(¤@¤å¦h®×)¥u¦³Ä¶¤å¨S¦³©w½Z
            If ET03 <> "" Then
            '2018/11/22 END
               'Modified by Lydia 2023/03/08 False, strUserNum =>§ï§PÂ_ IIf(bolToFile = True, True, False)
               NowPrint ET02, ET01, ET03, IIf(bolToFile = True, True, False), strUserNum, , , , , iCopy, , True, True
               'Added by Lydia 2023/03/08
               If bolToFile = True Then '©w½Z
                  'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                  If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                      Sleep 100
                  End If
                  'end 2023/05/03
               End If
               'end 2023/03/08
            End If
         End If
         MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(m_TM01) & " ]¡I"
         
      Else
      'end 2008/6/12
         '³]©w­n¦C¦L¦a§}±ø
         m_blnPrintAddress = True
         'Add By Sindy 2010/01/14
         'Add By Sindy 2011/8/10 ªü½¬¥u¯d308­n¥X¶Ç¯u«Ê­±©w½Z
         'If stCP10 = "101" Or stCP10 = "308" Then
         If stCP10 = "308" Then
            '¥[­^¤å¶Ç¯u«Ê­±
            NowPrint m_CP09, "03", "98", False, strUserNum, , , , , 1
         End If
         '2010/01/14 End
         
         'Add By Sindy 2018/11/22 µù¥UÅÜ§ó(¤@¤å¦h®×)¥u¦³Ä¶¤å¨S¦³©w½Z
         If ET03 <> "" Then
         '2018/11/22 END
            'Add By Sindy 2010/7/28 ªü½¬»¡­n2¥÷§ï¬°1¥÷
            If (stCP10 = "101" Or stCP10 = "308") And ET03 = "10" Then
               NowPrint ET02, ET01, ET03, False, strUserNum, 0, , , , 1
            '2010/7/28 End
            Else
               'Modified by Lydia 2023/03/08 False, strUserNum =>§ï§PÂ_ IIf(bolToFile = True, True, False)
               NowPrint ET02, ET01, ET03, IIf(bolToFile = True, True, False), strUserNum, 0
               'Added by Lydia 2023/03/08
               If bolToFile = True Then '©w½Z
                   'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                   If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                      Sleep 100
                   End If
                   'end 2023/05/03
               End If
               'end 2023/03/08
            End If
         End If
         
         '¦^ÂÐ³æ¥u­n¦L1¥÷
         If ET03r <> "" Then
            NowPrint ET02, ET01, ET03r, False, strUserNum, 0, , , , 1
         End If
         
         If ET03_1 <> "" Then
            'Modified by Lydia 2023/03/08 False, strUserNum =>§ï§PÂ_ IIf(bolToFile = True, True, False)
            NowPrint ET02, ET01, ET03_1, IIf(bolToFile = True, True, False), strUserNum, 0
            'Added by Lydia 2023/03/08
            If bolToFile = True Then 'Â½Ä¶
               'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
               If PUB_PrintWord2File(g_WordAp, strFilePath, strFN02) = True Then
                  Sleep 100
               End If
               'end 2023/05/03
            End If
            'end 2023/03/08
         End If
      End If
   End If
   
   'Added by Lydia 2023/03/08 ¤U¸ü¨÷©v°ÏªºÃÒ®ÑPDF
   'Mark by Lydia 2023/06/05 ¹q¤l©Î¯È¥»ÃÒ®Ñ²Î¤@¦b³Ì«á¤U¸ü¨÷©v°ÏªºÃÒ®ÑPDF
   'If bolToFile = True Then
   '   '«O¯d´ú¸Õ¥Î¡GFCT-46767
   '   'strSql = "select cpp14 From casepaperpdf where cpp01='CB2012458' " & _
   '               "and instr(upper(cpp02),upper('." & IIf(m_TM136 = "1", "CERT", "1001") & ".PDF'))>0"
   '   If InStr("103,302", m_CP10) > 0 Then  'Added by Lydia 2023/05/03 ¦b¿é¤J¡u®Ö­ã-¸É´«µoÃÒ®Ñ103¡v¡B¡u®Ö­ã-§ó¥¿302¡v¡A¤ñ·Ó¡uµù¥UÃÒ¿é¤J1701¡vªº³W«h
   '      strSql = "select cpp14 From casepaperpdf where cpp01='" & m_NickCp09 & "' and instr(upper(cpp02),upper('." & IIf(m_TM136 = "1", "CERT", m_strCP10) & ".PDF'))>0"
  '    'Added by Lydia 2023/05/03 ¨ä¥L®Ö­ã
   '   Else
   '      strSql = "select cpp14 From casepaperpdf where cpp01='" & m_NickCp09 & "' and instr(upper(cpp02),upper('." & m_strCP10 & ".PDF'))>0"
   '   End If
   '   'end 2023/05/03
   '   intI = 1
   '   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   '   If intI = 1 Then
   '      If PUB_GetFtpFile("" & RsTemp.Fields("cpp14"), strFilePath & "\" & strFN03, "Casepaperpdf") = True Then
   '      End If
   '   End If
   'End If
   'end 2023/03/08
   'end 2023/06/05
   
   Exit Sub
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Sub

'Add By Cheng 2002/02/01
'«O¯d¤W¤@¦¸¿é¤Jªº¸ê®Æ
Public Sub SetLastData()
Me.textTM14.Text = "" & m_strLastTextTM14
Me.textTMBM07_1.Text = "" & m_strLastTextTMBM07_1
Me.textTMBM07_2.Text = "" & m_strLastTextTMBM07_2
'Modify By Cheng 2002/07/22
'Me.textTM16S.Text = "" & m_strLastTextTM16S
'Me.textTM17.Text = "" & m_strLastTextTM17
End Sub

'Add By Cheng 2002/02/01
'²MªÅ¤W¤@¦¸¿é¤Jªº¸ê®Æ
Public Sub ClearLastData()
m_strLastTextTM14 = Empty
m_strLastTextTMBM07_1 = Empty
m_strLastTextTMBM07_2 = Empty
'Modify By Cheng 2002/07/22
'm_strLastTextTM16S = Empty
'm_strLastTextTM17 = Empty
End Sub

'Add By Cheng 2002/06/05
Private Function GetDelayTime(strTM10 As String) As Integer
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

StrSQLa = "Select NA15 From Nation Where NA01='" & strTM10 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   GetDelayTime = Val("0" & rsA.Fields(0).Value)
Else
   GetDelayTime = 0
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   'Add By Sindy 2010/12/24
   'Modified by Morgan 2022/6/17
   'If Me.textTM15.Enabled = True And Me.textTM15.Visible = True Then
   If Me.textTM15.Enabled = True And Me.textTM15.Locked = False And Me.textTM15.Visible = True Then
   'end 2022/6/17
      Cancel = False
      textTM15_Validate Cancel
      If Cancel = True Then
         textTM15.SetFocus
         Exit Function
      End If
   End If
   
   If Me.textCP53.Enabled = True And Me.textCP53.Visible = True Then
      Cancel = False
      textCP53_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP54.Enabled = True And Me.textCP54.Visible = True Then
      Cancel = False
      textCP54_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      If Me.textCP53.Visible And Me.textCP54.Visible Then
         If Val(Me.textCP53.Text) > Val(Me.textCP54.Text) Then
            MsgBox "¤é´Á°Ï¶¡¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
            Me.textCP53.SetFocus
            textCP53_GotFocus
            Exit Function
         End If
      End If
   End If
   If Me.Text1.Enabled = True And Me.Text1.Visible = True Then
      'MODIFY BY SONIA 2015/6/22 ´ðûA»¡ÃÒ®Ñ¤é´Á¤£­n±a,§ó§ï«áµoÃÒ¤é¤£·|©M­ì¨Ó¬Û¦P,¦ý¤£¥iªÅ¥ÕFCT-036102
      'Text1_Validate Cancel
      If Text1 = "" Then
         'Modify By Sindy 2015/6/25
         'modify by sonia 2019/5/2
         'If m_CP64 <> "" And _
            (InStr(m_CP64, "§ó§ïµù¥UÃÒ") > 0 Or InStr(m_CP64, "¸Éµoµù¥UÃÒ") > 0) Then
         If (m_CP64 <> "" And (InStr(m_CP64, "§ó§ïµù¥UÃÒ") > 0 Or InStr(m_CP64, "¸Éµoµù¥UÃÒ") > 0)) Or textMod.Text = "Y" Then
         'end 2019/5/2
         '2015/6/25 END
            MsgBox "ÃÒ®Ñ¤é´Á¤£¥iªÅ¥Õ!!!", vbExclamation + vbOKOnly
            Text1.SetFocus
            Text1_GotFocus
            Exit Function
         End If
      Else
         Cancel = False
         Text1_Validate Cancel
         If Cancel = True Then
            Text1.SetFocus
            Text1_GotFocus
            Exit Function
         End If
      End If
      'END 2015/6/22
   End If
   
   'Add By Sindy 2022/5/5
   If Val(textTM14) > 0 Then
      Cancel = False
      textTM14_Validate Cancel
      If Cancel = True Then
         textTM14.SetFocus
         textTM14_GotFocus
         Exit Function
      End If
   End If
   Cancel = False
   textCP14_Validate Cancel
   If Cancel = True Then
      textCP14.SetFocus
      textCP14_GotFocus
      Exit Function
   End If
   Cancel = False
   textCP48_Validate Cancel
   If Cancel = True Then
      textCP48.SetFocus
      textCP48_GotFocus
      Exit Function
   End If
   '2022/5/5 END
   
   'Added by Lydia 2023/09/04 ­×§ï¤é¤å²Õ¤§®Ö­ã-§ó¥¿(©µ®i®Ö­ã¨ç)¤§©w½Z¤ÎÄ¶¤å:©w½Z®×¥ó©Ê½è¬°¡u©µ®i¡v®É¡AÀË¬d"­ì¨ç¤½§i¤é"¤£¥i¬°ªÅ¥Õ
   If txtADate.Visible = True And Trim(txtADate) = "" And Trim(Left(Combo1, 4)) = "102" Then
      MsgBox "½Ð¿é¤J­ì¨ç¤½§i¤é¡I", vbExclamation
      txtADate.SetFocus
      txtADate_GotFocus
      Exit Function
   End If
   'end 2023/09/04
   
   TxtValidate = True
End Function

Private Function GetOldTM23(strCP09 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim intPos As Integer

GetOldTM23 = ""
StrSQLa = "Select CP64 From Caseprogress Where CP09='" & strCP09 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    intPos = InStr("" & rsA.Fields(0).Value, "X")
    If intPos > 0 Then
        GetOldTM23 = "" & Mid("" & rsA.Fields(0).Value, intPos, 9)
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2002/09/05
Private Function GetUSRate() As Double
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetUSRate = 0
'Modify By Cheng 2002/12/13
'À³¥H¥Á°ê¦~§ì³Ì±µªñ¨t²Î¤éªº¸ê®Æ
'strSQLA = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & ServerDate & " AND ROWNUM = 1 ORDER BY USXR01 "
StrSQLa = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & (ServerDate - 19110000) & " ORDER BY USXR01 DESC "
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    'Modify By Cheng 2002/12/13
'   GetUSRate = rsA.Fields(0).Value
   GetUSRate = CDbl(rsA.Fields(0).Value)
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2002/09/05
Private Sub ProcessPrint()
Screen.MousePointer = vbHourglass

Load Frmacc2480
Frmacc2480.Text1.Text = m_strSerialNo
Frmacc2480.Text2.Text = m_strSerialNo
Frmacc2480.Combo1.Text = Me.Combo2.Text
Frmacc2480.Command2_Click: DoEvents
Unload Frmacc2480
Screen.MousePointer = vbDefault
End Sub

'Add By Cheng 2004/04/01
'ÀË¬dÅÜ§ó¨Æ¶µÀÉ¬O§_¦³¤W­ã
Private Function ChkChangeEvent(strCE01 As String, strColName As String, strColName1 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select " & strColName & "," & strColName1 & " From ChangeEvent Where CE01='" & strCE01 & "' And " & strColName & "='1' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    ChkChangeEvent = "" & rsA.Fields(1).Value
Else
    ChkChangeEvent = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'¨ú±o«È¤á­^¤å¦WºÙ
Private Function GetCustEngName(strCU0102 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

If strCU0102 = "" Then GetCustEngName = "": Exit Function
StrSQLa = "Select CU05||Decode(CU88, Null, '', ' '||CU88)||Decode(CU89, Null, '', ' '||CU89)||Decode(CU90, Null, '', ' '||CU90) From Customer Where CU01='" & Mid(strCU0102, 1, 8) & "' And CU02='" & Mid(strCU0102, 9, 1) & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCustEngName = "" & rsA.Fields(0).Value
Else
    GetCustEngName = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Function RestrictGoods(strCE01 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

RestrictGoods = False
StrSQLa = "Select * From ChangeEvent Where CE01='" & strCE01 & "' And CE46 Is Not Null "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    RestrictGoods = True
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'add by nick 2005/06/28 §PÂ_¦³µL²Ä¤G´Á©Î¬O¥þ´Áªº
' Åª¨ú®×¥ó¶i«×ÀÉ
Private Function Query716717_cp() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' ¨ú±o®×¥ó¶i«×ÀÉÀÉ®×¤¤Äæ¦ì
   strSql = "SELECT count(*) FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' and cp10 in ('716','717') and cp27 is not null "
            
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
   If rsTmp.Fields(0).Value > 0 Then
        Query716717_cp = True
   Else
        Query716717_cp = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Added by Lydia 2023/09/04
Private Sub txtADate_GotFocus()
   TextInverse txtADate
End Sub
'Added by Lydia 2023/09/04
Private Sub txtADate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(txtADate) = False Then
      ' ÀË¬d¬O§_¬°¥Á°ê¦~
      If CheckIsTaiwanDate(txtADate, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½Tªº¤½§i¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtADate_GotFocus
      End If
      ' ­ì¨ç¤½§i¤é¤£¥i¶W¹L¨t²Î¤é
      'Modified by Lydia 2024/03/28 ±Æ°£¦Û°Ê±a¤J
      If txtADate.Locked = False And DBDATE(txtADate) > strSrvDate(1) Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "­ì¨ç¤½§i¤é¤£¥i¶W¹L¨t²Î¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtADate_GotFocus
      End If
   End If
End Sub
