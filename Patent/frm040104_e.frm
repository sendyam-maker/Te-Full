VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040104_e 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專發文-減免退費"
   ClientHeight    =   6180
   ClientLeft      =   12
   ClientTop       =   732
   ClientWidth     =   9276
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9276
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   900
      MaxLength       =   4
      TabIndex        =   10
      Top             =   3510
      Width           =   540
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   1530
      MaxLength       =   1
      TabIndex        =   15
      Top             =   4140
      Width           =   255
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   5310
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1320
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   5310
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1590
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   900
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1590
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   5
      Left            =   900
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1860
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   900
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1320
      Width           =   240
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   4
      Left            =   5976
      TabIndex        =   21
      Top             =   4395
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   3
      Left            =   2616
      MaxLength       =   1
      TabIndex        =   20
      Top             =   4395
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   6000
      TabIndex        =   18
      Top             =   4140
      Width           =   720
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   4740
      MaxLength       =   1
      TabIndex        =   17
      Top             =   4125
      Width           =   360
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   3930
      MaxLength       =   1
      TabIndex        =   16
      Top             =   4125
      Width           =   360
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   2
      Left            =   5550
      MaxLength       =   9
      TabIndex        =   22
      Top             =   1920
      Width           =   1032
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   0
      Left            =   1545
      MaxLength       =   1
      TabIndex        =   11
      Top             =   3870
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   1
      Left            =   6384
      MaxLength       =   1
      TabIndex        =   13
      Top             =   3870
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   2
      Left            =   5130
      MaxLength       =   1
      TabIndex        =   12
      Top             =   3870
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   1
      Left            =   900
      MaxLength       =   7
      TabIndex        =   8
      Top             =   3210
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   37
      Top             =   750
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   36
      Top             =   750
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   35
      Top             =   750
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   960
      MaxLength       =   3
      TabIndex        =   34
      Top             =   750
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   435
      Index           =   3
      Left            =   3396
      TabIndex        =   28
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   435
      Index           =   1
      Left            =   6672
      TabIndex        =   31
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   5844
      TabIndex        =   30
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "繳費紀錄(&Y)"
      CausesValidation=   0   'False
      Height          =   435
      Index           =   2
      Left            =   7896
      TabIndex        =   32
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "變更事項(&R)"
      Height          =   435
      Index           =   4
      Left            =   4635
      Style           =   1  '圖片外觀
      TabIndex        =   29
      Top             =   15
      Width           =   1200
   End
   Begin VB.CheckBox Check1 
      Caption         =   "減免退費移至次年"
      Enabled         =   0   'False
      Height          =   195
      Left            =   105
      TabIndex        =   19
      Top             =   4440
      Width           =   1812
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1485
      Left            =   60
      TabIndex        =   61
      Top             =   4650
      Width           =   9060
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   0
         Left            =   975
         TabIndex        =   23
         Top             =   720
         Width           =   495
         VariousPropertyBits=   671107099
         MaxLength       =   2
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   1
         Left            =   1935
         TabIndex        =   24
         Top             =   720
         Width           =   495
         VariousPropertyBits=   671107099
         MaxLength       =   2
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   2
         Left            =   5190
         TabIndex        =   25
         Top             =   720
         Width           =   255
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   270
         Index           =   3
         Left            =   7770
         TabIndex        =   26
         Top             =   720
         Width           =   1095
         VariousPropertyBits=   671107097
         MaxLength       =   8
         Size            =   "1931;476"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   375
         Index           =   4
         Left            =   975
         TabIndex        =   27
         Top             =   1050
         Width           =   7905
         VariousPropertyBits=   -1467987941
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13944;661"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   2
         Left            =   4620
         TabIndex        =   78
         Top             =   480
         Width           =   1365
         VariousPropertyBits=   27
         Caption         =   "Label2"
         Size            =   "2408;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "業務人員:"
         Height          =   180
         Index           =   4
         Left            =   3690
         TabIndex        =   77
         Top             =   480
         Width           =   765
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   16
         Left            =   4620
         TabIndex        =   73
         Top             =   240
         Width           =   1365
         VariousPropertyBits=   27
         Caption         =   "Label2"
         Size            =   "2408;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限:"
         Height          =   180
         Index           =   23
         Left            =   3690
         TabIndex        =   72
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限:"
         Height          =   180
         Index           =   24
         Left            =   6750
         TabIndex        =   71
         Top             =   240
         Width           =   765
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   15
         Left            =   7680
         TabIndex        =   70
         Top             =   240
         Width           =   1215
         VariousPropertyBits=   27
         Caption         =   "Label2"
         Size            =   "2143;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人:"
         Height          =   180
         Index           =   22
         Left            =   135
         TabIndex        =   69
         Top             =   510
         Width           =   585
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   14
         Left            =   975
         TabIndex        =   68
         Top             =   510
         Width           =   1485
         VariousPropertyBits=   27
         Caption         =   "Label2"
         Size            =   "2619;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   8
         Left            =   975
         TabIndex        =   67
         Top             =   240
         Width           =   1500
         VariousPropertyBits=   27
         Caption         =   "Label2"
         Size            =   "2646;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收文號:"
         Height          =   180
         Index           =   21
         Left            =   135
         TabIndex        =   66
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "費用是否要雙倍:              (Y:雙倍)"
         Height          =   180
         Index           =   26
         Left            =   3690
         TabIndex        =   65
         Top             =   765
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "繳納第:                    至                  年年費"
         Height          =   180
         Index           =   25
         Left            =   135
         TabIndex        =   64
         Top             =   765
         Width           =   3285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "進度備註:"
         Height          =   180
         Index           =   29
         Left            =   135
         TabIndex        =   63
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "下次繳費日:                            "
         Height          =   180
         Index           =   28
         Left            =   6750
         TabIndex        =   62
         Top             =   765
         Width           =   2205
      End
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   864
      Left            =   7704
      TabIndex        =   14
      Top             =   3840
      Width           =   1500
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;1524"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   570
      Index           =   5
      Left            =   3060
      TabIndex        =   9
      Top             =   3210
      Width           =   6045
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "10663;1005"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   300
      Index           =   0
      Left            =   1245
      TabIndex        =   0
      Top             =   2295
      Width           =   7830
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "13811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   300
      Index           =   1
      Left            =   1245
      TabIndex        =   1
      Top             =   2580
      Width           =   7830
      VariousPropertyBits=   671107099
      MaxLength       =   250
      Size            =   "13811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   300
      Index           =   2
      Left            =   1245
      TabIndex        =   2
      Top             =   2850
      Width           =   7830
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "13811;529"
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
      Left            =   120
      TabIndex        =   89
      Top             =   3555
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Index           =   16
      Left            =   2220
      TabIndex        =   88
      Top             =   3255
      Width           =   765
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否列印通知函：        (N:不印)"
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   87
      Top             =   4185
      Width           =   2445
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人:"
      Height          =   180
      Left            =   6696
      TabIndex        =   86
      Top             =   3912
      Width           =   948
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收據號碼："
      Height          =   180
      Index           =   31
      Left            =   4992
      TabIndex        =   85
      Top             =   4440
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收據：　　(1.附件 2.遺失 3.已作帳)"
      Height          =   180
      Index           =   5
      Left            =   2028
      TabIndex        =   84
      Top             =   4440
      Width           =   2772
   End
   Begin VB.Label lblCP81 
      AutoSize        =   -1  'True
      Caption         =   "lblCP81"
      Height          =   180
      Left            =   8595
      TabIndex        =   83
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblCP81C 
      AutoSize        =   -1  'True
      Caption         =   "本案最新減免狀態："
      Height          =   180
      Left            =   6885
      TabIndex        =   82
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "減免第　　　 至　　　 年年費共　　　　　元"
      Height          =   180
      Index           =   30
      Left            =   3300
      TabIndex        =   81
      Top             =   4170
      Width           =   3690
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   11
      Left            =   6690
      TabIndex        =   80
      Top             =   1965
      Width           =   2475
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4366;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年費通知人:"
      Height          =   180
      Index           =   27
      Left            =   4575
      TabIndex        =   79
      Top             =   1965
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容：　　(Y:Word)"
      Height          =   180
      Index           =   20
      Left            =   3300
      TabIndex        =   76
      Top             =   3915
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否列印申請書：　　(N:不印)"
      Height          =   180
      Index           =   19
      Left            =   105
      TabIndex        =   75
      Top             =   3915
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Index           =   18
      Left            =   120
      TabIndex        =   74
      Top             =   3255
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(日):"
      Height          =   180
      Index           =   15
      Left            =   120
      TabIndex        =   60
      Top             =   2895
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(英):"
      Height          =   180
      Index           =   14
      Left            =   120
      TabIndex        =   59
      Top             =   2625
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(中):"
      Height          =   180
      Index           =   13
      Left            =   120
      TabIndex        =   58
      Top             =   2340
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   9120
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   120
      X2              =   9120
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Index           =   1
      Left            =   4560
      TabIndex        =   57
      Top             =   510
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   12
      Left            =   5415
      TabIndex        =   56
      Top             =   510
      Width           =   1365
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2408;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   10
      Left            =   7515
      TabIndex        =   55
      Top             =   510
      Width           =   1575
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2778;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   960
      TabIndex        =   54
      Top             =   1065
      Width           =   3285
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5794;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "巳繳年費:"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   53
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Index           =   12
      Left            =   120
      TabIndex        =   52
      Top             =   1905
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Index           =   11
      Left            =   4560
      TabIndex        =   51
      Top             =   1635
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   50
      Top             =   1635
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Index           =   9
      Left            =   4560
      TabIndex        =   49
      Top             =   1365
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Index           =   8
      Left            =   120
      TabIndex        =   48
      Top             =   1365
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   47
      Top             =   750
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   7
      Left            =   4560
      TabIndex        =   46
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   45
      Top             =   510
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   0
      Left            =   960
      TabIndex        =   44
      Top             =   510
      Width           =   480
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "847;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   5415
      TabIndex        =   43
      Top             =   1065
      Width           =   1365
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2408;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   3
      Left            =   1230
      TabIndex        =   42
      Top             =   1365
      Width           =   3285
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5794;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   4
      Left            =   5640
      TabIndex        =   41
      Top             =   1365
      Width           =   3555
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "6271;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   5
      Left            =   1230
      TabIndex        =   40
      Top             =   1635
      Width           =   3285
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5794;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   6
      Left            =   5640
      TabIndex        =   39
      Top             =   1635
      Width           =   3555
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "6271;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   7
      Left            =   1230
      TabIndex        =   38
      Top             =   1905
      Width           =   3285
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5794;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Index           =   2
      Left            =   6855
      TabIndex        =   33
      Top             =   510
      Width           =   585
   End
End
Attribute VB_Name = "frm040104_e"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (Text4,Text5,lstNameAgent,Label2)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'整理 by Morgan 2005/7/15
Option Explicit
'Modify by Morgan 2005/7/15 改用動態陣列
'Dim pa(1 To T_PA) As String, cp(1 To T_CP) As String
Dim pa() As String, cp() As String

Dim intWhere As Integer
Dim strCaseFee(1 To 2) As String 'strCaseFee(1) 國家檔中繳費年度，strCaseFee(2) 國家檔中起算日
Dim blnOverDate As Boolean '下次繳費日是否超過專用期止日
Dim m_strOfficalFee  As String
Dim m_strServiceFee  As String
Dim m_lngDiscount As Long  '預繳年費減免金額
Dim m_strNP09 As String '舊法定期限(抓工作日) 西元年
Dim m_strNP09_1 As String '新法定期限 西元年
Dim m_blnClkChgEvnBtn As Boolean '是否按下變更事項按鈕
Dim m_EndDate As String

'減免退費收文號
Dim m_919CP09 As String
'年費收文號
Dim m_605CP09 As String
Dim m_lngFeeDiscount As Long  '待繳年費減免金額
'Add by Morgan 2004/7/22
Dim m_DiscType As String   '減免身分
Dim m_bolActive As Boolean 'Active事件是否已觸發
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/23 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_t109 As Boolean 'Added by Morgan 2021/4/16
Dim m_bolFMP As Boolean 'Added by Lydia 2023/06/20 是否為FMP案
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/20 是否為寰華

Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)

   Dim strTxt(1 To 5) As String, strTmp As String
   Dim ii As Integer
   
   ii = 1
   If Text3(0).Text = Text3(1).Text Then
      strTmp = Text3(0).Text
   Else
      strTmp = Text3(0).Text & "年至第" & Text3(1).Text
   End If
   EndLetter ET01, ET02, ET03, strUserNum
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','減免年次','" & strTmp & "')"
   
    ii = ii + 1
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','退費金額','" & Format(m_lngDiscount) & "')"
       
    ii = ii + 1
    Select Case Text2(3)
      Case "1"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','收據號碼','" & Text2(4).Text & "')"
      Case "2"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','理由','，惟因遺失　鈞局第 " & Text2(4).Text & " 號收據正本')"
      Case "3"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','理由','，惟因　鈞局第 " & Text2(4).Text & " 號收據已作帳不易取回')"
   End Select
       
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(ii, strTxt) Then
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Sub
Private Function Process(Index As Integer) As Boolean
   
   Dim stET03 As String, stET01 As String
   
   '重新檢查欄位有效性
   If TxtValidate = True Then
   
      'Add by Morgan 2009/4/28
      If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text2(1)) = False Then
         Exit Function
      End If
      If m_CP123s = "Y" Then
      'end 2009/4/28
         'Add by Morgan 2009/3/23 設定是否算發文室案件
         'modify by sonia 2014/6/23 加傳發文規費, P-108903
         If ModifyDispatch(cp(9), m_CP09s, m_CP123s, 0, Text2(1)) = False Then
             Exit Function
         End If
      End If
      'Add by Amy 2014/10/14 P台灣案發文控制
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
        If pa(1) = "P" And cp(9) < "C" And pa(9) = 台灣國家代號 Then
            'Modify by Amy 2014/11/27 取消ChkOneDayHasCP27判斷,接洽單改檢查,因考慮可能同時發文其他案件性質情形
            '類一定要有接洽單才可發文
            If m_919CP09 < "B" Then
                'If PUB_CheckPDF2(cp(9), 0, True, strExc(0)) = False And ChkOneDayHasCP27(pa(1), pa(2), pa(3), pa(4), cp(5) + 19110000) = False Then
                If PUB_CheckPDF3(Text1(0), Text1(1), Text1(2), Text1(3)) = False Then
                    Exit Function
                End If
            End If
          
            'AB類申請書確認檢查,符合條件才可發文
            'Modified by Morgan 2015/3/17
            'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And Text8(0) = "N" And PUB_CheckPDF2(m_919CP09, 1, True, strExc(0)) = False Then
            If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And Text8(0) = "N" Then
               If PUB_CheckPDF2(m_919CP09, 1, True, strExc(0)) = False Then
            'end 2015/3/17
                  MsgBox "無申請書PDF檔 ,不可發文!", vbInformation
                  Exit Function
               End If 'Added by Morgan 2015/3/17
            End If
            'end 2014/11/27
            
         'Added by Morgan 2016/6/29 非臺灣案電子化
         ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
            If m_919CP09 < "B" And Left(cp(12), 1) <> "F" Then
                If PUB_CheckPDF3(Text1(0), Text1(1), Text1(2), Text1(3)) = False Then
                    Exit Function
                End If
            End If
         'end 2016/6/29
         
        End If
      End If
      'end 2014/10/14
      
      If FormSave = False Then
         MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
      Else
         Process = True
         
         'Add by Morgan 2007/6/14
         If pa(9) = "000" Then
            PUB_ReAsignInform pa(1), pa(2), pa(3), pa(4), m_605CP09
         End If
         
         '2012/7/23 add by sonia 檢查計算出的規費與進度檔的規費不同,仍繼續發文時發mail給智權人員
         If pa(9) = "000" Then
            If Val(m_strOfficalFee) <> Val(cp(17)) Then
               '2013/7/2 modify by sonia 改用共用module
               PUB_ChkOfficialFee cp(9), m_strOfficalFee
            End If
         End If
         '2012/7/23 end
         
         '列印申請書
         If Text8(0) <> "N" Then
            If Check1.Value = 1 Then
               PrintLetter IIf(Text8(2) = "Y", True, False)
            Else
               stET01 = "01": stET03 = "01"
               If Text2(3) = "1" Then stET03 = "02"
               StartLetter stET01, m_919CP09, stET03
               'Modify by Amy 2014/08/25 +傳strLetterRecNo 申請書修改從frm1105_1改
               NowPrint m_919CP09, stET01, stET03, IIf(Text8(2) = "Y", True, False), strUserNum, 0, , , , , , , , , , , , m_919CP09
               If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And Text8(2) = "Y" Then
                    frm1105_1.m_RecNo = m_919CP09
                    'Modify By Sindy 2022/5/11 流水號要足6碼
                    frm1105_1.m_PdfName = Text1(0) & Text1(1) & IIf(Text1(2) & Text1(3) = "000", "", "-" & Text1(2) & "-" & Text1(3)) & "." & cp(10) & ".DATA.PDF"
                    frm1105_1.Show
               End If
               'end 2014/08/25
            End If
         End If
         
         'Added by Morgan 2021/5/6
         If Text8(3) <> "N" Then
            If Check1.Value = 1 Then
               StartLetter2 "02", "30"
               NowPrint m_605CP09, "02", "30", False, strUserNum, 0, , , , , , , , , , , , m_605CP09
            End If
         End If
         'end 2021/5/6
      End If
   End If
   
End Function
'列印專利年費減免申請書
Private Sub PrintLetter(ByVal bolEdit As Boolean)

   Dim iAppCnt As Integer, ii As Integer
   Dim stLetter(1 To 20) As String
   Dim stData(1 To 5, 0 To 5) As String
   Dim iCopys As Integer
   Dim stSales As String
   Dim bolName As Boolean
   
   stLetter(1) = pa(11) '申請案號
   stLetter(2) = pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)  '本所案號
   stLetter(3) = ChangeTStringToTDateString(Text2(1).Text)  '申請日期
   stLetter(4) = pa(8)  '專利種類
   stLetter(5) = pa(22) '證書號
   'Modify by Morgan 2004/7/22
   'stLetter(6) = Text2(0)  '減免身分
   stLetter(6) = m_DiscType  '減免身分
   '減免年次
   If Check1.Value = 1 Then
      If Val(Text5(1)) <= 6 Then
         stLetter(7) = Text3(0) & " 年至第 " & Text5(1)
      Else
         stLetter(7) = Text3(0) & " 年至第 6"
      End If
   Else
      stLetter(7) = Text3(0) & " 年至第 " & Text3(1)
   End If
   stLetter(8) = Format(m_lngDiscount + m_lngFeeDiscount, "###,###") '減免金額
   
   iAppCnt = 1
   For ii = 27 To 30
      If pa(ii) <> "" Then
         iAppCnt = iAppCnt + 1
      End If
   Next
   stLetter(9) = Format(iAppCnt) '申請人數
   Erase stData
   For ii = 1 To iAppCnt
      Call PUB_GetAppData(pa(25 + ii), stData, ii)
      stData(ii, 4) = Label2(2 + ii)
      stData(ii, 5) = pa(30 + ii)
   Next
   stLetter(10) = m_strOfficalFee   '規費
   stLetter(11) = Text5(0) '年費起
   stLetter(12) = Text5(1) '年費迄
   stLetter(13) = ChangeTStringToTDateString(Text5(3).Text) '下次期限
   iCopys = PUB_GetCopys(pa, stSales)
   stLetter(14) = PUB_GetStaffST15(stSales, "2")   '業務區
   stLetter(15) = GetStaffName(stSales)   '智權人員
   stLetter(16) = Format(iCopys)   '份數
   stLetter(17) = pa(47) '分所號
   stLetter(18) = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         stLetter(18) = stLetter(18) & "、" & lstNameAgent.List(ii)
      End If
   Next
   stLetter(18) = Mid(stLetter(18), 2)
   '是否出名
   If Text8(1) <> "N" Then
      bolName = True
   Else
      bolName = False
   End If
   stLetter(19) = m_919CP09 'Add by Morgan 2006/3/9
   stLetter(20) = Text5(2) 'Add by Morgan 2009/3/19
   PUB_PrintDiscForm stLetter, stData, bolName, bolEdit
   
End Sub

Private Sub Check1_Click()
   If Check1.Enabled = True Then
      If Check1.Value = 1 Then
         Frame1.Enabled = True
      Else
         Frame1.Enabled = False
      End If
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0, 3 '確定'同時發文
         'Modify by Morgan 2010/2/10 改呼叫函數方式以便鎖定按鍵
         cmdOK(Index).Enabled = False
         If Not Process(Index) Then
            cmdOK(Index).Enabled = True
         Else
            If Index = 0 Then
'               '若有未發文資料顯示警告
'               PUB_GetCPunIssueDatas "" & Me.Text1(0).Text & "-" & Me.Text1(1).Text & "-" & IIf(Len("" & Me.Text1(2).Text) <= 0, "0", Me.Text1(2).Text) & "-" & IIf(Len("" & Me.Text1(3).Text) <= 0, "00", Me.Text1(3).Text)
'               frm040104_1.Show
'               frm040104_1.Clear
               '若有未發文資料顯示警告
               If PUB_GetCPunIssueDatas("" & Me.Text1(0).Text & "-" & Me.Text1(1).Text & "-" & IIf(Len("" & Me.Text1(2).Text) <= 0, "0", Me.Text1(2).Text) & "-" & IIf(Len("" & Me.Text1(3).Text) <= 0, "00", Me.Text1(3).Text)) Then
                  'Add By Sindy 2013/5/20
                  If frm040104_1.bolIsEMPFlow = True Then
                     frm090202_4.QueryData
                  End If
                  '2013/5/20 End
                  frm040104_1.Show
                  frm040104_1.Command1_Click
               Else
                  'Add By Sindy 2013/5/20
                  If frm040104_1.bolIsEMPFlow = True Then
                     Unload frm040104_1
                     frm090202_4.Show
                     frm090202_4.QueryData
                  Else
                  '2013/5/20 End
                     frm040104_1.Show
                     ' 90.07.11 modify by louis (回第一個畫面清除)
                     frm040104_1.Clear
                  End If
               End If
            Else
               'Add By Sindy 2013/5/20
               If frm040104_1.bolIsEMPFlow = True Then
                  frm090202_4.QueryData
               End If
               '2013/5/20 End
               frm040104_1.Show
               frm040104_1.ReQuery
            End If
            Unload Me
         End If
      Case 1
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            frm040104_1.Show
         End If
         Unload Me
      Case 2
         Me.Hide
         Set frm060104_b.oParent = Me 'Add by Morgan 2011/10/5
         frm060104_b.LoadMe pa(1), pa(2), pa(3), pa(4), 5
         m_blnClkChgEvnBtn = True
      Case 4
         Me.Hide
         Set frm06010303_1.oParent = Me 'Add by Morgan 2011/10/5
         frm06010303_1.LoadMe cp(9), pa(1), pa(2), pa(3), pa(4), 47
         frm06010303_1.Text41 = "N"
         frm06010303_1.Caption = "內專發文-變更事項"
         m_blnClkChgEvnBtn = True
   End Select
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
End Sub

Private Function FormSave() As Boolean

   Dim ii As Integer, iMax As Long
   Dim bInTrans As Boolean
   
On Error GoTo ErrHnd
   '年費通知人
   If IsEmptyText(Text2(2)) = False Then
      Text2(2) = Text2(2) & String(9 - Len(Text2(2)), "0")
   End If
   '案件名稱
   pa(5) = CNULL(ChgSQL(Text4(0)))
   pa(6) = CNULL(ChgSQL(Text4(1)))
   pa(7) = CNULL(ChgSQL(Text4(2)))
   pa(76) = CNULL(ChgSQL(Text2(2)))
   
   '繳年費資料
   If Check1.Value = 1 Then
      For ii = Val(Text5(0)) To Val(Text5(1))
         pa(72) = pa(72) & "," & Format(ii)
         pa(73) = pa(73) & "," & DBDATE(Text2(1))
         If ii = Val(Text5(0)) Then
            pa(74) = pa(74) & "," & Trim(Text5(2))
         Else
            pa(74) = pa(74) & ","
         End If
      Next
      
   End If
   
   cnnConnection.BeginTrans
   bInTrans = True
   
   'Added by Morgan 2013/6/7 自 lstNameAgent_Validate 移來,否則若觸發 Form_Activate 事件會跑 ReadPatent 導致 cp(110) 被清除
   cp(110) = ""
   If lstNameAgent.Visible = True Then
      For ii = 0 To lstNameAgent.ListCount - 1
         If lstNameAgent.Selected(ii) = True Then
            'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
            'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
            'Modified by Morgan 2021/12/15f Forms2.0 改用模組
            'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
            cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         End If
      Next
      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   End If
   'end 2013/6/7
   
   'Add by Amy 2014/09/09 P台灣案電子化
   If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
      '沒有客戶函
      If pa(9) = 台灣國家代號 Then
         'Modify by Amy 2015/02/13 原:判斷同一天 沒有其他有規費的發文(若與年費一起發文,cp(10)=605)
           '1.    電子送件且規費>0,有收據
           '2.非電子送件且經發文室要計件,有回執
         'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'         strExc(1) = PUB_GetLetterJudge(pa(1), "919")
'         If cp(118) = "Y" Then
'            If Val(m_strOfficalFee) > 0 Then
'                PUB_AddLetterProgress m_919CP09, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'            End If
'         Else
'            If Left(m_CP123s, 1) = "Y" Then
'                PUB_AddLetterProgress m_919CP09, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'            End If
'         End If
'         'end 2015/02/13
         'end 2015/03/06
       
        '申請書
        If Text8(0) <> "N" And ExistCheck("AppForm", "AF01", m_919CP09, "", False) = False Then
           '新增申請書轉檔記錄
           PUB_AddAppForm m_919CP09
        End If
      End If
   End If
   'end 2014/09/09
   
   '更新進度檔 919
   'Modify by Morgan 2005/7/15 加 cp110
   'Modified by Lydia 2021/05/25 +CP113工作時數
   'Modified by Lydia 2023/06/20 +CP14
   strSql = "UPDATE CASEPROGRESS SET CP27=" & CNULL(TransDate(Text2(1), 2)) & ",CP22=" & CNULL(Text8(1)) & ",CP81='Y'" & _
      ",cp110=" & CNULL(cp(110)) & ",cp64=" & CNULL(ChgSQL(Text5(5))) & " ,cp113=" & CNULL(txtCP113, True) & ",CP14=" & CNULL(cp(14)) & _
      " WHERE CP09='" & m_919CP09 & "'"
      
   cnnConnection.Execute strSql
  
   'Add by Morgan 2004/7/22
   '設定客戶減免身分
   If pa(9) = "000" Then
      For ii = 1 To 5
         If txtAD(ii).Enabled = True Then
            '身分有變更才要做
            If txtAD(ii).Tag <> txtAD(ii).Text Then
               '不可減免
               If txtAD(ii).Text = "N" Then
                  strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "N")
               '自然人
               'Modify by Morgan 2004/7/29
               '學校也不用證明
               'ElseIf txtAD(ii).Text = "1" Then
               ElseIf (txtAD(ii).Text = "1" Or txtAD(ii).Text = "2") Then
                  strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "Y", txtAD(ii).Text)
               '公司
               Else
                  '原來沒有減免資料或不可減免
                  If txtAD(ii).Tag = "" Or txtAD(ii).Tag = "N" Then
                     strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "Y", txtAD(ii).Text, pa(1), pa(2), pa(3), pa(4))
                  '修改減免身分別
                  Else
                     strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "Y", txtAD(ii).Text)
                  End If
               End If
               cnnConnection.Execute strSql
            End If
         End If
      Next
   End If

   If Check1.Value = 1 Then
      '更新進度檔 605
      'Modify by morgan 2004/8/11 加 cp84
      'Modify by Morgan 2005/7/15 加 cp110
      'Modified by Lydia 2023/06/20 +CP14
      strSql = "UPDATE CASEPROGRESS SET CP27=" & CNULL(TransDate(Text2(1), 2)) & ",CP22=" & CNULL(Text8(1)) & ",CP81='Y', cp84=" & Format(Val(m_strOfficalFee)) & _
         ",CP14=" & CNULL(cp(14)) & ",cp110=" & CNULL(cp(110)) & ",cp64=" & CNULL(ChgSQL(Text5(4))) & ",CP118=NULL WHERE CP09='" & m_605CP09 & "'"
      
      cnnConnection.Execute strSql
      
      '若有下次繳費日時, 要新增下一程序檔
      If Text5(3) <> "" Then
         iMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了 iMax = objPublicData.GetNextProgressNo
         strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
              "VALUES ('" & m_605CP09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & _
              "','" & PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & DBDATE(Text5(3)) & _
              "," & m_strNP09_1 & "," & iMax & ")"
        
        cnnConnection.Execute strSql
      End If
   End If
   
   '更新基本檔
   strSql = "UPDATE PATENT SET PA05=" & pa(5) & ",PA06=" & pa(6) & _
      ",PA07=" & pa(7) & ",PA76=" & pa(76) & _
      ",PA72='" & pa(72) & "',PA73='" & pa(73) & "'," & _
      "PA74='" & pa(74) & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      
   cnnConnection.Execute strSql
   
   'Add by Morgan 2009/3/23
   PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130
   'Add by Amy 2015/02/13 更新收據/回執設定
   'Modify by Amy 2015/03/06 +發文日參數
   PUB_UpdateLP19 cp(1), cp(2), cp(3), cp(4), m_CP09s, m_CP123s, Text2(1)
   
   'Added by Morgan 2021/4/16 更新109修法減免退費紀錄
   If m_t109 Then
      strSql = "update t109 set t14=" & DBDATE(Text2(1)) & " where t01='" & pa(1) & "' and t02='" & pa(2) & "' and t03='" & pa(3) & "' and t04='" & pa(4) & "' and t14=0"
      cnnConnection.Execute strSql, intI
   End If
   'end 2021/4/16
   
   '*沒出客戶通知函
   If Text8(3) = "N" Then
      '可能會重新發文
      cnnConnection.Execute "delete LetterProgress where lp01='" & m_919CP09 & "'", intI
      If Check1.Value = 1 Then
         cnnConnection.Execute "delete LetterProgress where lp01='" & m_605CP09 & "'", intI
      End If
   Else
      strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10))
      If Left(m_CP123s, 1) = "Y" Then
         If Check1.Value = 1 Then
            PUB_AddLetterProgress m_605CP09, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
         Else
            PUB_AddLetterProgress m_919CP09, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
         End If
      Else
         If Check1.Value = 1 Then
            PUB_AddLetterProgress m_605CP09, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
         Else
            PUB_AddLetterProgress m_919CP09, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
         End If
      End If
   End If
        
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   If bInTrans Then cnnConnection.RollbackTrans
   
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If

End Function

Private Sub Form_Activate()
   '若有按下變更事項按鈕, 則重新讀取資料
   If m_blnClkChgEvnBtn = True Then
      cp(9) = m_919CP09
      ReadPatent
       m_bolActive = False
      Label2(0) = m_919CP09
      m_blnClkChgEvnBtn = False
   End If
   
   'Add by Morgan 2004/7/22
   If m_bolActive = True Then Exit Sub
   m_bolActive = True
   
   If pa(9) = "000" Then
      Dim i As Integer
      For i = 1 To 5
         If txtAD(i).Enabled = True And txtAD(i).Text = "" Then
            txtAD(i).SetFocus
            Exit Sub
         End If
      Next
   End If
   
   'Modify by Morgan 2005/7/15
   '若沒有客戶減免身分需輸入則游標預設在發文日
   'If Text8(1).Enabled = True Then Text8(1).SetFocus
   'Modify by Morgan 2006/1/19 --玲玲
   'If Text2(1).Enabled = True Then Text2(1).SetFocus
   If Text2(3).Enabled = True Then Text2(3).SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   '本所案號
   With frm040104_1
      Text1(0) = .Text1
      Text1(1) = .Text2
      Text1(2) = .Text3
      Text1(3) = .Text4
      m_919CP09 = .Tag
   End With
   
   'Add by Morgan 2005/7/15
   ReDim pa(TF_PA)
   ReDim cp(TF_CP)
      
   cp(9) = m_919CP09
   ReadPatent
   
   cp(110) = "" '要清空,否則若重新發文會殘留前次發文資料,當新案有改出名人而本程序未改選將會造成不一致 Added by Morgan 2012/9/7
   
   'Add by Morgan 2005/7/15
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   Text8(1).Visible = False
   lstNameAgent.Clear
   If pa(9) = "000" Then
      PUB_SetOurAgent lstNameAgent, pa(), cp(110), , True  'Modified by Morgan 2021/12/15 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2005/7/14 END
   
   Label2(0).Caption = m_919CP09
   m_blnClkChgEvnBtn = False
   
   'Added by Morgan 2017/1/11
   '專利處人員操作時年費通知人欄位鎖住以避免不小心改到(目前只有外專人員會設定)
   If Left(Pub_StrUserSt03, 1) = "P" Then
      Text2(2).Locked = True
   End If
   'end 2017/1/11
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Set frm040104_e = Nothing 'Removed by Morgan 2021/12/15 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub ReadPatent()
Dim oEle
Dim i As Integer
Dim strTmp(0 To 5) As String, varTmp As Variant, strTmp1(0 To 5) As String
'Add by Morgan 2004/7/22
Dim strAD10 As String, strCU15 As String
Dim m_Fee As String         '銷帳服務費 2012/8/1 add by sonia
Dim m_Official As String    '銷帳規費   2012/8/1 add by sonia
  
   For Each oEle In Label2
      oEle.Caption = ""
   Next
   
   Set oEle = Nothing
   
   For Each oEle In Text5
      oEle.Text = ""
   Next
   
   pa(1) = Text1(0): pa(2) = Text1(1): pa(3) = Text1(2): pa(4) = Text1(3)
         
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      For i = 26 To 30
         If pa(i) <> "" Then ChgType i
      Next
      
      If pa(9) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(pA(9), strExc(0)) Then Label2(12) = strExc(0)
         If ClsPDGetNation(pa(9), strExc(0)) Then Label2(12) = strExc(0)
      End If
      
      Text4(0) = pa(5): Text4(1) = pa(6): Text4(2) = pa(7)
      Label2(9) = pa(72)
      
      If pa(76) <> "" Then Text2(2) = pa(76): ChgType 6

      strTmp1(0) = cp(9)
      
      For i = 1 To 4
         strTmp1(i) = pa(i)
      Next
      If GetMoneyDate(pa(8), pa(9), strTmp1, strCaseFee(1), strCaseFee(2)) = True Then
         
      End If
   End If
   
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(13) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(13), strExc(0)) Then Label2(1) = strExc(0)
         If ClsPDGetStaff(cp(13), strExc(0)) Then Label2(1) = strExc(0)
      End If
      If cp(27) = "" Then
         Text2(1) = strSrvDate(2)
      Else
         Text2(1) = cp(27)
      End If
      'Added by Lydia 2023/06/20 判斷FCP案,寰華案
      If Left(cp(12), 1) = "F" And pa(9) <> "000" Then
         m_bolFMP = True
      Else
         m_bolFMP = False
      End If
      m_bolFMP2 = False
      If m_bolFMP = True Then
         m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4))
         '寰華案:承辦人為外專程序時,改為操作人員
         If m_bolFMP2 = True Then
            cp(14) = GetFCPUser(cp(14))
         End If
      End If
      'end 2023/06/20
      If cp(14) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(14), strExc(0)) Then
         If ClsPDGetStaff(cp(14), strExc(0)) Then
               Label2(10) = strExc(0)
         End If
      End If
      Text5(5) = cp(64)
      Text8(1) = cp(22)
   End If
   
   '2012/8/1 add by sonia 若有銷帳則要扣除銷帳規費
   If Val(cp(77)) > 0 Then
      If GetCP77Detail(cp(9), m_Fee, m_Official) = True Then
         cp(17) = cp(17) - m_Official
      End If
   End If
   '2012/8/1 end
   
   'Add by Morgan 2004/6/23
   '設定申請人年費減免身分
   lblCP81C.Visible = True
   lblCP81.Visible = True
   lblCP81.Caption = PUB_GetCP81(pa)
   
   'Add by Morgan 2004/7/21
   '減免身分
   If pa(9) = "000" Then
      For i = 1 To 5
         txtAD(i).Enabled = False
         txtAD(i).Tag = ""
         txtAD(i).Text = ""
         If pa(25 + i) <> "" Then
            txtAD(i).Text = PUB_GetAD03(pa(25 + i), pa(9), strAD10, strCU15)
            txtAD(i).Tag = txtAD(i).Text
            '個人只可設定自然人(1)
            If strCU15 = "0" Then
               txtAD(i).Text = "1"
            'Added by Morgan 2014/7/15 學校也預設--玲玲
            ElseIf strCU15 = "2" Then
               txtAD(i).Text = "2"
            'end 2014/7/15
            '公司
            Else
               If txtAD(i).Text = "Y" Then
                  txtAD(i).Text = strAD10
                  txtAD(i).Tag = txtAD(i).Text
               End If
               txtAD(i).Enabled = True
            End If
         End If
      Next
   End If
   
   '設定減免資料
   SetDiscData pa(14), pa(72), pa(25)
   
   If Get605No(pa, m_605CP09) = True Then
      'Modify by Morgan 2005/7/15
      'Erase cp
      ReDim cp(TF_CP)
      
      m_strNP09 = ""
      cp(9) = m_605CP09
      Label2(8).Caption = cp(9)
      If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
         If cp(13) <> "" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetStaff(cp(13), strExc(0)) Then Label2(2) = strExc(0)
            If ClsPDGetStaff(cp(13), strExc(0)) Then Label2(2) = strExc(0)
         End If
         Label2(16) = cp(6): Label2(15) = cp(7)
         '取得下一程序的法定期限
         If GetNPLimit(pa, m_strNP09) = True Then
            '若法定期限為假日時, 抓大於法定期限最近的工作天
            If m_strNP09 <> "" Then
               m_strNP09 = DBDATE(PUB_GetLawDay(DBDATE(m_strNP09)))
            End If
         End If
         'Added by Lydia 2023/06/20 寰華案:承辦人為外專程序時,改為操作人員
         If m_bolFMP2 = True Then
            cp(14) = GetFCPUser(cp(14))
         End If
         'end 2023/06/20
         If cp(14) <> "" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetStaff(cp(14), strExc(0)) Then Label2(14) = strExc(0)
            If ClsPDGetStaff(cp(14), strExc(0)) Then Label2(14) = strExc(0)
         End If
         Text5(4) = cp(64)
         'Added by Morgan 2021/5/6
         Text5(0) = cp(53)
         Text5(1) = cp(54)
         If Text5(1) <> "" Then
            Text5_Validate 1, False
         End If
         'end 2021/5/6
      End If
      Check1.Enabled = True
      Check1.Value = 1
      Check1_Click
   Else
      Check1.Enabled = False
      Check1.Value = 0
   End If
    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25
End Sub
'取得下一程序的法定期限
Private Function GetNPLimit(ByRef pa() As String, ByRef stNP09 As String) As Boolean
On Error GoTo ErrHnd

   strSql = "Select NP09 From NextProgress Where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " And Np07='605' Order By NP09 Desc "
      
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      stNP09 = "" & adoRecordset("NP09").Value
      GetNPLimit = True
   End If
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
End Function

Private Function Get605No(ByRef pa() As String, ByRef stCP09 As String) As Boolean

On Error GoTo ErrHnd

   strSql = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
      " AND CP10='605' AND CP27 IS NULL AND CP57 IS NULL"
      
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      stCP09 = adoRecordset.Fields(0)
      Get605No = True
   End If
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
End Function

Private Function ChgType(iSitu As Integer) As Boolean

   Dim i As Integer, bolChk As Boolean, varTmp As Variant
   Dim nPos As Integer
   Dim nCurrPos As Integer
   Dim aryCaseFee As Variant
   Dim aryCurrFee As Variant
   Dim bFind As Boolean
 
   ChgType = False
   Select Case iSitu
      Case 0:
         If IsEmptyText(Text5(0)) = False Then
            aryCaseFee = Split(strCaseFee(2), ",")
            aryCurrFee = Split(Label2(9), ",")
            ' 找尋已繳年度串列中空白的位置
            For nPos = 0 To UBound(aryCurrFee)
               If IsEmptyText(aryCurrFee(nPos)) = True Then
                  Exit For
               End If
            Next nPos
            If nPos > UBound(aryCaseFee) Then
               MsgBox "無繳年費年度，請查明後再輸入 !", vbCritical
            Else
               If Text5(0) <> aryCaseFee(nPos) Then
                  MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
               Else
                  ChgType = True
               End If
            End If
            Erase aryCurrFee
            Erase aryCaseFee
         Else
            MsgBox "起始繳費年度不可空白！", vbCritical
         End If
         
      Case 1:
         If IsEmptyText(Text5(1)) = False Then
            aryCaseFee = Split(strCaseFee(2), ",")
            bFind = False
            ' 找尋繳費年度迄在繳費年度串列中的位置(是否存在?)
            For nCurrPos = 0 To UBound(aryCaseFee)
               If Text5(1) = aryCaseFee(nCurrPos) Then
                  bFind = True
                  Exit For
               End If
            Next nCurrPos
            ' 數入的年度不在繳費年度串列中
            If bFind = False Then
               MsgBox "繳費年度迄輸入錯誤，請查明後再輸入 !", vbCritical
            Else
               ' 找尋繳年度起在繳費年度串列中的位置
               bFind = False
               For nPos = 0 To UBound(aryCaseFee)
                  If Text5(0) = aryCaseFee(nPos) Then
                     bFind = True
                     Exit For
                  End If
               Next nPos
               ' 繳費年度起及迄的範圍不正確
               If nPos > nCurrPos Then
                  MsgBox "繳費年度範圍輸入錯誤，請查明後再輸入 !", vbCritical
               Else
                  ChgType = True
               End If
            End If
            Erase aryCaseFee
            
            ' 計算下次繳年費日期
            Call CompNextFeeDate(pa, cp, m_strNP09_1, blnOverDate, Text2(1), Text5(1))
            If m_strNP09_1 <> "" Then
               'Added by Lydia 2025/10/29
               If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                  Text5(3) = TransDate(PUB_GetPOurDeadline(m_strNP09_1, pa(9)), 1)
               Else
               'end 2025/10/29
                  'Added by Morgan 2014/10/28
                  If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                     Text5(3) = TransDate(PUB_GetOurDeadline(m_strNP09_1), 1)
                  Else
                  'end 2014/10/28
                     Text5(3) = TAIWANDATE(PUB_GetWorkDay1(CompDate(2, -2, m_strNP09_1), True))
                  End If 'Added by Morgan 2014/10/28
               End If 'Added by Lydia 2025/10/29
            Else
               Text5(3) = ""
            End If
         Else
            MsgBox "繳費年度迄年不可空白！", vbCritical
         End If
      Case 2
        '若法定期限為假日則用大於法定期限最近的工作日與發文日比較
         If DBDATE(Text2(1)) > DBDATE(m_strNP09) Then
            If Text5(iSitu) <> "Y" Then
               MsgBox "發文日大於法定期限則此欄必須為 Y !", vbCritical
               Text5(iSitu) = "Y"
            End If
            ChgType = True
         Else
            If Text5(iSitu) = "Y" Then
               MsgBox "費用是否雙倍錯誤 !", vbCritical
               Text5(iSitu) = "Y"
            Else
               ChgType = True
            End If
         End If
         
      Case 3
         '若下次繳費日不超過專用期止日時才要檢查
         If blnOverDate = False Then
            If IsEmptyText(Text5(3)) = False Then
               If CheckIsTaiwanDate(Text5(3), False) = False Then
                  ChgType = False
                  MsgBox "請輸入正確的下次繳費日日期", vbOKOnly + vbCritical, "檢核資料"
               Else
                  ChgType = True
               End If
            Else
               ChgType = False
               MsgBox "請輸入正確的下次繳費日日期", vbOKOnly + vbCritical, "檢核資料"
            End If
         Else
            Text5(iSitu) = ""
            ChgType = True
         End If
         
      '申請人中文
      Case 26, 27, 28, 29, 30
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(pa(iSitu), strExc(0)) Then
         If ClsLawLawGetName(pa(iSitu), strExc(0)) Then
            Label2(iSitu - 23) = strExc(0)
            ChgType = True
         End If
         
      Case Else
         ChgType = True
         
   End Select
End Function

Private Sub Text2_GotFocus(Index As Integer)
   TextInverse Text2(Index)
   If Index <> 4 Then
      'edit by nickc 2007/07/11 切換輸入法改用API
      'Text2(Index).IMEMode = 2
      CloseIme
   End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 2
         KeyAscii = UpperCase(KeyAscii)
      Case 3   '附件
         If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      
   End Select
End Sub

Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1
         If Text2(Index) = "" Then
            MsgBox "發文日不可空白 !", vbCritical
            Cancel = True
         Else
            'Modify by Morgan 2003/6/24 暫改可大於系統日但不可大於 930709,並取消930709控制
            'If Not ChkDate(Text2(1)) Or Val(Text2(1)) > Val(strSrvDate(2)) Then
            '   MsgBox "發文日期不正確或發文日大於系統日，請重新輸入 !", vbCritical
            If Not ChkDate(Text2(1)) Or DBDATE(Val(Text2(1))) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
               MsgBox "發文日期不正確或發文日大於系統日下一個工作日，請重新輸入 !", vbCritical
            '2011/12/8 END
               Cancel = True
            End If
         End If
      Case 2
         If Text2(Index) <> "" Then
            'edit by nickc 2007/02/05 不用 dll 了
            'If objLawDll.LawGetName(Text2(Index), strExc(0)) Then
            If ClsLawLawGetName(Text2(Index), strExc(0)) Then
               Label2(11) = strExc(0)
            Else
               Cancel = True
            End If
         End If
         'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
         If Cancel = False Then
            If PUB_CheckStatus(Text2(Index).Text) = False Then Cancel = True
         End If
      Case 3
         If Check1.Value <> 1 Then
            If Text2(Index) = "" Then
               MsgBox "收據不可空白 !", vbCritical
               Cancel = True
            End If
         End If
      Case 4
         If Check1.Value <> 1 Then
            If Text2(Index) = "" Then
               MsgBox "收據號碼不可空白 !", vbCritical
               Cancel = True
            End If
         End If
         
   End Select
End Sub

Private Sub Text3_Change(Index As Integer)
   If Index = 0 Or Index = 1 Then
      If Val(Text3(0)) > 0 And Val(Text3(1)) > 0 Then
         m_lngDiscount = GetDiscount(Val(Text3(0)), Val(Text3(1)))
         Text3(2) = Format(m_lngDiscount, "###,###")
      End If
   End If
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Or Index = 1 Then
      If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub Text4_GotFocus(Index As Integer)
   TextInverse Text4(Index)
End Sub

Private Sub Text4_Validate(Index As Integer, Cancel As Boolean)
   Dim iMax As Integer
   If Index = 1 Then
      iMax = 180
   Else
      iMax = 160
   End If
   
   If CheckLengthIsOK(Text4(Index), iMax) = False Then
      Cancel = True
   End If
   If Cancel = True Then TextInverse Text4(Index)
End Sub

Private Sub Text5_GotFocus(Index As Integer)
   TextInverse Text5(Index)
   If Index = 4 Then
      'edit by nickc 2007/07/11 切換輸入法改用API
      'Text5(Index).IMEMode = 1
      OpenIme
   Else
      'edit by nickc 2007/07/11 切換輸入法改用API
      'Text5(Index).IMEMode = 2
      CloseIme
   End If
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 2
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
         
   End Select
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)

   Select Case Index
      Case 0, 1, 2, 3   '年費起訖
         Cancel = Not ChgType(Index)
      Case 4   '進度備註
         If CheckLengthIsOK(Text5(Index), 2000) = False Then
            Cancel = True
         End If
   End Select
   If Cancel = True Then TextInverse Text5(Index)
End Sub

Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 2   '修改申請書
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      
      Case Else
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

'計算下次繳年費日期
Private Function CompNextFeeDate(ByRef pa() As String, ByRef cp() As String, ByRef stNP09 As String, ByRef bolOver As Boolean, ByVal stCP27 As String, ByVal strPayEnd As String) As Boolean
   
   Dim strDate As String, strTmp1(0 To 5) As String, i As Integer
   
On Error GoTo ErrHnd
   
   
   '預設下次繳年費日期不超過專用期止日
   bolOver = False
   Select Case pa(8)
      Case "1":
         strSql = "SELECT NA06 FROM NATION " & _
            "WHERE NA01 = '" & pa(9) & "' "
      Case "2":
         strSql = "SELECT NA08 FROM NATION " & _
            "WHERE NA01 = '" & pa(9) & "' "
      Case "3":
         strSql = "SELECT NA10 FROM NATION " & _
            "WHERE NA01 = '" & pa(9) & "' "
   End Select
   
   If IsEmptyText(strSql) = False Then
      CheckOC
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         If IsNull(adoRecordset.Fields(0)) = False Then
            Select Case adoRecordset.Fields(0)
               Case 1: strDate = DBDATE(cp(5))
               Case 2: strDate = DBDATE(pa(10))
               Case 3: strDate = DBDATE(stCP27)
               Case 4: strDate = DBDATE(cp(25))
               Case 5: strDate = DBDATE(pa(14))
               Case 6: strDate = DBDATE(pa(27))
               Case 7: strDate = DBDATE(pa(12))
            End Select
            
            ' 依日期再+繳年費迄
            If IsEmptyText(strDate) = False Then
               strDate = DBDATE(DateAdd("yyyy", Val(strPayEnd), ChangeWStringToWDateString(DBYEAR(strDate) & DBMONTH(strDate) & DBDAY(strDate))))
               strDate = DBDATE(DateSerial(Val(DBYEAR(strDate)), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) - 1))
               '若有專用期止日
               If pa(25) <> "" Then
                  '若專用期止日小於等於下次繳費日
                  If Val(DBDATE(pa(25))) <= Val(strDate) Then
                     stNP09 = ""
                     '下次繳年費日期超過專用期止日
                     bolOver = True
                  '若專用期止日大於下次繳費日
                  Else
                     stNP09 = strDate
                     '下次繳年費日期不超過專用期止日
                     bolOver = False
                  End If
               '若無專用期止日
               Else
                  strTmp1(0) = cp(9)
                  For i = 1 To 4
                     strTmp1(i) = pa(i)
                  Next
                  If GetMoneyDate(Val(pa(8)) + 10, pa(9), strTmp1, strCaseFee(1), strCaseFee(2), m_EndDate) Then   '抓專用期起止日
                      If GetMoneyDate(pa(8), pa(9), strTmp1, strCaseFee(1), strCaseFee(2)) = True Then
                      End If
                  End If
                  '若專用期止日小於等於下次繳費日
                  If Val(DBDATE(m_EndDate)) <= Val(strDate) Then
                      Text5(3) = ""
                      '下次繳年費日期超過專用期止日
                      bolOver = True
                  '若專用期止日大於下次繳費日
                  Else
                      stNP09 = strDate
                      '下次繳年費日期不超過專用期止日
                      bolOver = False
                  End If
               End If
            End If
         End If
      End If
   End If
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
   
End Function

Private Function TxtValidate() As Boolean
   'add by nickc 2008/05/01
   If IsDebt(pa(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        Exit Function
   End If
   
   Dim objTxt As Object
   Dim Cancel As Boolean
   Dim ii As Integer, i As Integer
   'Add by Morgan 2004/7/20
   Dim stAppNo As String   '未設定減免身分客戶代碼
   
   'Added by Morgan 2021/12/15 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/15
   
   '919 必要欄位
   For Each objTxt In Text2
      If objTxt.Enabled = True Then
         Cancel = False
         Text2_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Me.Text2(objTxt.Index).SetFocus
            Text2_GotFocus objTxt.Index
            Exit Function
         End If
      End If
   Next
   
   '案件名稱
   TxtValidate = False
   For Each objTxt In Text4
      If objTxt.Enabled = True Then
         Cancel = False
         Text4_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Me.Text4(objTxt.Index).SetFocus
            Text4_GotFocus objTxt.Index
            Exit Function
         End If
      End If
   Next
      'Add by Morgan 2004/7/22
      If pa(9) = "000" Then
         m_DiscType = ""
         For i = 1 To 5
            m_DiscType = m_DiscType & txtAD(i).Text
            If txtAD(i).Enabled = True Then
               If txtAD(i).Text = "" Then
                  MsgBox "申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】減免身分不可空白", vbInformation
                  txtAD(i).SetFocus
                  txtAD_GotFocus i
                  Exit Function
               ElseIf txtAD(i).Text = "N" Then
                  MsgBox "申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】減免身分不可為【N】", vbInformation
                  txtAD(i).SetFocus
                  txtAD_GotFocus i
                  Exit Function
               '公司可減免
               'Modify by Morgan 2004/7/29
               '學校不需證明
               'ElseIf (txtAD(i).Text = "2" Or txtAD(i).Text = "3") Then
               ElseIf (txtAD(i).Text = "2") Then
                  '變更
                  If (txtAD(i).Tag <> "2" And txtAD(i).Tag <> "") Then
                     If MsgBox("確定要變更申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】減免身分為【學校】？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                        txtAD(i).SetFocus
                        txtAD_GotFocus i
                        Exit Function
                     End If
                  End If
               ElseIf (txtAD(i).Text = "3") Then
                  '新增或變更
                  If (txtAD(i).Text <> "3") Then
                     If MsgBox("申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】的減免身分將設定為【中小企業】，確定有【證明文件】存放於本卷？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                        txtAD(i).SetFocus
                        txtAD_GotFocus i
                        Exit Function
                     End If
                  End If
               End If
            End If
         Next
       End If
   
   '年費檢查
   If Check1.Value = 1 Then
      'Added by Morgan 2021/4/16
      If cp(14) = "" Then
         MsgBox "本案年費尚未分案，不可發文！", vbCritical
         Exit Function
      End If
      'end 20214/16
      
      '605 必要欄位
      For Each objTxt In Text5
         If objTxt.Enabled = True Then
            Cancel = False
            Text5_Validate objTxt.Index, Cancel
            If Cancel = True Then
               Me.Text5(objTxt.Index).SetFocus
               Text5_GotFocus objTxt.Index
               Exit Function
            End If
         End If
      Next
   
      '檢查計算出的規費與進度檔的規費是否相同
      If ChkPatentYearFee(pa(9), pa(8), "Y00000001", cp(10), Me.Text5(0).Text, Me.Text5(1).Text, IIf(Me.Text5(2).Text = "Y", True, False)) = False Then Exit Function
      '若下次繳費日不超過專用期止日才要檢查
      If blnOverDate = False Then
          If IsEmptyText(Text5(3)) = True Then
             MsgBox "請輸入下次繳費日", vbOKOnly + vbCritical, "檢核資料"
             If Me.Text5(3).Enabled Then Text5(3).SetFocus
             Exit Function
          End If
      End If
      
      If PUB_ChkNP605(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text) Then
          MsgBox "本案下一程序有<年費>期限，不可發文!!!", vbExclamation + vbOKOnly
          Exit Function
      End If
   End If
   
   'Add by Morgan 2005/7/15
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   '2005/7/14 END
   
   'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
   If Pub_ChkACS112isNull(pa(1), pa(2), pa(3), pa(4), txtCP113) = True Then
         txtCP113.SetFocus
         txtCP113_GotFocus
         Exit Function
   End If
   'end 2021/05/25
   
   TxtValidate = True
End Function

'計算相關費用
Private Function ChkPatentYearFee( _
    strYF01 As String, strYF02 As String, strYF03 As String, _
    strYF04 As String, strYF05From As String, strYF05To As String, blnDouble As Boolean) As Boolean
'strYF01  申請國家
'strYF02  專利種類
'strYF03  代理人
'strYF04  案件性質
'strYF05From  起始年度
'strYF05To  終止年度
'blnDouble  規費是否雙倍

Dim ii As Integer, iYear As Integer

   ChkPatentYearFee = False
   m_strOfficalFee = 0
   m_strServiceFee = 0
   m_lngFeeDiscount = 0
   ii = 1
   '取得案件性質為年費的相關費用
   strSql = "Select * From PatentYearFee Where YF01='" & strYF01 & "' AND YF02='" & strYF02 & "' AND YF03='" & strYF03 & "' AND YF04='" & strYF04 & "' AND YF05>=" & Val(strYF05From) & " AND YF05<=" & Val(strYF05To) & " Order By YF05 "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   While Not adoRecordset.EOF
      iYear = Val(adoRecordset.Fields("YF05").Value)
      m_strOfficalFee = Val(m_strOfficalFee) + Val(adoRecordset.Fields("YF07").Value)
      If iYear >= 1 And iYear <= 3 Then
         m_lngFeeDiscount = m_lngFeeDiscount + 800
         m_strOfficalFee = m_strOfficalFee - 800
      ElseIf iYear >= 4 And iYear <= 6 Then
         m_lngFeeDiscount = m_lngFeeDiscount + 1200
         m_strOfficalFee = m_strOfficalFee - 1200
      End If
      
      '起始那年年費是否雙倍
      If blnDouble = True And ii = 1 Then m_strOfficalFee = Val(m_strOfficalFee) * 2
      m_strServiceFee = Val(m_strServiceFee) + Val(adoRecordset.Fields("YF06").Value)
      adoRecordset.MoveNext
      'Add By Cheng 2003/01/02
      ii = ii + 1
   Wend
   m_strOfficalFee = m_strOfficalFee - m_lngDiscount
    '若不等
   If "" & cp(17) <> m_strOfficalFee Then
      If MsgBox("計算出的規費( " & Format(m_strOfficalFee, "#,##0") & " )與目前進度檔的規費( " & Format(cp(17), "#,##0") & " )不同，是否要繼續作業???", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
         ChkPatentYearFee = True
      Else
          ChkPatentYearFee = False
      End If
   '若相等
   Else
     ChkPatentYearFee = True
   End If
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
   
End Function

Private Sub SetDiscData(ByRef stPA14 As String, ByRef stPA72 As String, ByRef stPA25 As String)

   Dim stYear1 As String, stYear2 As String
   
   'Added by Morgan 2021/4/16
   '檢查109修法減免退費紀錄
   m_t109 = False
   strExc(0) = "select * from t109 where t01='" & pa(1) & "' and t02='" & pa(2) & "' and t03='" & pa(3) & "' and t04='" & pa(4) & "' and t14=0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_t109 = True
      stYear1 = "" & RsTemp("t09")
      stYear2 = "" & RsTemp("t10")
   Else
   'end 2021/4/16
   
      '減免起年：未使用過的起始年度
      'Modify by Morgan 2005/6/17 公告日為7/1者則93年不可減免,依繳費期限判斷是否適用減免規定
      'stYear1 = Format((940700 - Val(stPA14)) \ 10000 + 1)
      stYear1 = Format((940701 - Val(stPA14)) \ 10000 + 1)
      '已繳迄年
      stYear2 = Right(Trim(stPA72), 2)
      If Left(stYear2, 1) = "," Then stYear2 = Mid(stYear2, 2)
   
      '減免迄年最多到第6年
      If Val(stYear2) > 6 Then stYear2 = "6"
      
      'Add by Morgan 2010/8/4
      '930701以後公告者為錯繳退費起始年固定為1
      If Val(stYear1) < 1 Then
         stYear1 = "1"
         Text3(0).Enabled = True
         Text3(1).Enabled = True
      End If
      'end 2010/8/4
      
   End If 'Added by Morgan 2021/4/16
   
   Text3(0) = stYear1: Text3(1) = stYear2
   
End Sub

Private Function GetDiscount(iStartYr As Integer, iEndYr As Integer) As Long
   Dim ii As Integer, lngDiscount As Long
   
   lngDiscount = 0
   For ii = Val(iStartYr) To Val(iEndYr)
      '第1-3年減免800
      If ii < 4 Then
         lngDiscount = lngDiscount + 800
      '第4-6年減免1200
      ElseIf ii <= 6 Then
         lngDiscount = lngDiscount + 1200
      End If
   Next ii
   GetDiscount = lngDiscount
End Function

'Add by Morgan 2004/7/22
Private Sub txtAD_GotFocus(Index As Integer)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtAD(Index).IMEMode = 2
   CloseIme
   TextInverse txtAD(Index)
End Sub
'Add by Morgan 2004/7/22
'只有公司可輸入 2,3
Private Sub txtAD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2014/7/15 學校改預設且不可改
   'If Not (KeyAscii = 8 Or KeyAscii = 50 Or KeyAscii = 51 Or KeyAscii = 78) Then
   If Not (KeyAscii = 8 Or KeyAscii = 51 Or KeyAscii = 78) Then
      KeyAscii = 0
   End If
End Sub
'Add by Morgan 2005/7/15
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   cp(110) = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         'Modified by Morgan 2021/12/15f Forms2.0 改用模組
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   If bolCheck = True Then
      Text8(1) = ""
   Else
      Text8(1) = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

'Added by Morgan 2021/5/6
Private Sub StartLetter2(ByVal ET01 As String, ByVal ET03 As String)
    Dim strTxt(1 To 2) As String, strTmp As String
    Dim ii As Integer
    
    ii = 0
    EndLetter ET01, m_605CP09, ET03, strUserNum
    
    
     If Text5(0).Text = Text5(1).Text Then
        strTmp = "第 " & Text5(0) & " 年"
     Else
        strTmp = "第 " & Text5(0) & " 至 " & Text5(1) & " 年"
     End If
     
     ii = ii + 1
     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
     "VALUES ('" & ET01 & "','" & m_605CP09 & "','" & ET03 & "','" & strUserNum & _
     "','第幾年至幾年費','" & strTmp & "年費')"
     
     If Text5(3) <> "" Then
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & m_605CP09 & "','" & ET03 & "','" & strUserNum & _
        "','下次繳年費日','" & DBDATE(Text5(3)) & "')"
     End If
     
     If Not ClsLawExecSQL(ii, strTxt) Then
        MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
     End If
   
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/05/25
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
