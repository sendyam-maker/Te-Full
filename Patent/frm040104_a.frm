VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040104_a 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專發文-年費、維持費、延展費、年費移作次年"
   ClientHeight    =   5676
   ClientLeft      =   12
   ClientTop       =   732
   ClientWidth     =   9276
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5676
   ScaleWidth      =   9276
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   5850
      MaxLength       =   4
      TabIndex        =   15
      Top             =   3384
      Width           =   540
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   285
      Left            =   3000
      TabIndex        =   79
      Top             =   4950
      Width           =   6075
      Begin VB.TextBox txtPayToday 
         Height          =   264
         Left            =   5295
         MaxLength       =   1
         TabIndex        =   82
         Top             =   3
         Width           =   255
      End
      Begin VB.TextBox txtCP118 
         Height          =   264
         Left            =   1860
         MaxLength       =   1
         TabIndex        =   80
         Top             =   3
         Width           =   255
      End
      Begin VB.Label lblPayToday 
         AutoSize        =   -1  'True
         Caption         =   "電子送件是否當日扣款:         (Y/N)"
         Height          =   180
         Left            =   3360
         TabIndex        =   83
         Top             =   45
         Width           =   2655
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "是否電子送件:         (Y:是)"
         Height          =   180
         Index           =   2
         Left            =   690
         TabIndex        =   81
         Top             =   45
         Width           =   1995
      End
   End
   Begin VB.TextBox textCP84 
      Height          =   285
      Left            =   3180
      TabIndex        =   14
      Top             =   3375
      Width           =   1092
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   855
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1290
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   5
      Left            =   855
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1830
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   855
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1560
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   5310
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1560
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   5310
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1290
      Width           =   240
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   9
      Left            =   5190
      MaxLength       =   1
      TabIndex        =   17
      Top             =   3675
      Width           =   255
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1140
      TabIndex        =   18
      Top             =   4020
      Width           =   1395
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   3
      Left            =   5910
      MaxLength       =   1
      TabIndex        =   20
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   16
      Top             =   3675
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   1
      Left            =   6312
      MaxLength       =   1
      TabIndex        =   11
      Top             =   3060
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   0
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   19
      Top             =   4335
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   8
      Left            =   1140
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   23
      Top             =   5250
      Width           =   7992
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   270
      Index           =   7
      Left            =   1140
      MaxLength       =   8
      TabIndex        =   22
      Top             =   4935
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   6
      Left            =   1140
      MaxLength       =   9
      TabIndex        =   21
      Top             =   4620
      Width           =   1032
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   4
      Left            =   960
      MaxLength       =   8
      TabIndex        =   13
      Top             =   3375
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   2
      Left            =   4860
      MaxLength       =   1
      TabIndex        =   10
      Top             =   3060
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   1
      Left            =   1830
      MaxLength       =   2
      TabIndex        =   9
      Top             =   3060
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   0
      Left            =   960
      MaxLength       =   2
      TabIndex        =   8
      Top             =   3060
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   39
      Top             =   750
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   38
      Top             =   750
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   37
      Top             =   750
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   36
      Top             =   750
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   3
      Left            =   3396
      TabIndex        =   24
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6672
      TabIndex        =   27
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5844
      TabIndex        =   26
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "繳費紀錄(&Y)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7896
      TabIndex        =   28
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "變更事項(&R)"
      Height          =   400
      Index           =   4
      Left            =   4620
      Style           =   1  '圖片外觀
      TabIndex        =   25
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   1800
      Left            =   7596
      TabIndex        =   12
      Top             =   3072
      Width           =   1500
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;3175"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   0
      Left            =   1245
      TabIndex        =   0
      Top             =   2235
      Width           =   7830
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "13811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   1
      Left            =   1245
      TabIndex        =   1
      Top             =   2475
      Width           =   7830
      VariousPropertyBits=   671107099
      MaxLength       =   250
      Size            =   "13811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   2
      Left            =   1245
      TabIndex        =   2
      Top             =   2715
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
      Left            =   4950
      TabIndex        =   84
      Top             =   3435
      Width           =   765
   End
   Begin VB.Label lblCP84 
      AutoSize        =   -1  'True
      Caption         =   "發文規費:"
      Height          =   180
      Left            =   2280
      TabIndex        =   78
      Top             =   3420
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "第                    至                  年年費"
      Height          =   180
      Left            =   495
      TabIndex        =   77
      Top             =   3105
      Width           =   2610
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人:"
      Height          =   180
      Left            =   6612
      TabIndex        =   76
      Top             =   3108
      Width           =   948
   End
   Begin VB.Label lblCP81C 
      AutoSize        =   -1  'True
      Caption         =   "本案最新減免狀態："
      Height          =   180
      Left            =   6975
      TabIndex        =   75
      Top             =   1065
      Width           =   1665
   End
   Begin VB.Label lblCP81 
      AutoSize        =   -1  'True
      Caption         =   "lblCP81"
      Height          =   180
      Left            =   8685
      TabIndex        =   74
      Top             =   1065
      Width           =   480
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否修改通知函內容:              (Y:Word)"
      Height          =   180
      Left            =   3360
      TabIndex        =   73
      Top             =   3720
      Width           =   2985
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(日):"
      Height          =   180
      Left            =   120
      TabIndex        =   72
      Top             =   2745
      Width           =   1065
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(英):"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   71
      Top             =   2505
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(中):"
      Height          =   180
      Left            =   120
      TabIndex        =   70
      Top             =   2265
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   9120
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   120
      X2              =   9120
      Y1              =   2130
      Y2              =   2130
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   14
      Left            =   2640
      TabIndex        =   69
      Top             =   4065
      Width           =   4455
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "7858;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "代理人:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   68
      Top             =   4065
      Width           =   585
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "是否修改指示信或申請書內容:              (Y:Word)"
      Height          =   180
      Left            =   3360
      TabIndex        =   67
      Top             =   4380
      Width           =   3705
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否列印指示信或申請書:        (N:不印)"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   66
      Top             =   4380
      Width           =   3030
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否列印通知函:       (N:不印)"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   65
      Top             =   3720
      Width           =   2265
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   13
      Left            =   7830
      TabIndex        =   64
      Top             =   795
      Width           =   1395
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2461;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   1
      Left            =   6990
      TabIndex        =   63
      Top             =   795
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Index           =   1
      Left            =   4560
      TabIndex        =   62
      Top             =   510
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   12
      Left            =   5415
      TabIndex        =   61
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
      Index           =   11
      Left            =   2280
      TabIndex        =   60
      Top             =   4665
      Width           =   4815
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "8493;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   10
      Left            =   7710
      TabIndex        =   59
      Top             =   510
      Width           =   1395
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2461;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   960
      TabIndex        =   58
      Top             =   1065
      Width           =   3495
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "6165;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "巳繳年費:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   57
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Left            =   120
      TabIndex        =   56
      Top             =   1875
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Left            =   4560
      TabIndex        =   55
      Top             =   1605
      Width           =   675
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Left            =   120
      TabIndex        =   54
      Top             =   1605
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Left            =   4560
      TabIndex        =   53
      Top             =   1335
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Left            =   120
      TabIndex        =   52
      Top             =   1335
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   0
      Left            =   4560
      TabIndex        =   51
      Top             =   795
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   50
      Top             =   795
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   0
      Left            =   4560
      TabIndex        =   49
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   120
      TabIndex        =   48
      Top             =   510
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   0
      Left            =   960
      TabIndex        =   47
      Top             =   510
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3440;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   5400
      TabIndex        =   46
      Top             =   1065
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
      Index           =   2
      Left            =   5400
      TabIndex        =   45
      Top             =   795
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
      Index           =   3
      Left            =   1185
      TabIndex        =   44
      Top             =   1335
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
      Left            =   5625
      TabIndex        =   43
      Top             =   1335
      Width           =   3585
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "6324;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   5
      Left            =   1185
      TabIndex        =   42
      Top             =   1605
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
      Left            =   5625
      TabIndex        =   41
      Top             =   1605
      Width           =   3585
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "6324;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   7
      Left            =   1185
      TabIndex        =   40
      Top             =   1875
      Width           =   3285
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5794;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "下次繳費日:                            "
      Height          =   180
      Left            =   120
      TabIndex        =   35
      Top             =   4950
      Width           =   2205
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   120
      TabIndex        =   34
      Top             =   3420
      Width           =   585
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   6990
      TabIndex        =   33
      Top             =   510
      Width           =   585
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "年費通知人:"
      Height          =   180
      Left            =   120
      TabIndex        =   32
      Top             =   4620
      Width           =   945
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   120
      TabIndex        =   31
      Top             =   5280
      Width           =   765
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "繳納"
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   3105
      Width           =   360
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "費用是否要雙倍:              (Y:雙倍)"
      Height          =   180
      Left            =   3390
      TabIndex        =   29
      Top             =   3105
      Width           =   2580
   End
End
Attribute VB_Name = "frm040104_a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (Text15,lstNameAgent,Label2)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'整理 by Morgan 2005/7/15
Option Explicit
Dim strReceiveNo As String
'Modify by Morgan 2005/7/15 改用動態陣列
'Dim pa(1 To T_PA) As String, cp(1 To T_CP) As String
Dim pa() As String, cp() As String

Dim intWhere As Integer
Dim m_NP09 As String      ' 大陸案原法定期限
Dim m_NEXTNP08 As String  ' 下次繳年費本所期限
Dim strCaseFee(1 To 2) As String 'strCaseFee(1) 國家檔中繳費年度，strCaseFee(2) 國家檔中起算日
'Add By Cheng 2003/01/14
Dim blnOverDate As Boolean '下次繳費日是否超過專用期止日
'Add By Cheng 2003/04/15
Dim m_strOfficalFee  As String
Dim m_strServiceFee  As String
Dim m_strPoints  As String
'Add By Cheng 2003/04/24
Dim m_strNP09 As String
'Add By Cheng 2003/09/01
Dim m_strNP09_1 As String
Dim m_strNP09_2 As String '原始法限 Added by Morgan 2016/4/8

'Add By Cheng 2003/10/06
Dim m_blnClkChgEvnBtn As Boolean '是否按下變更事項按鈕
'92.10.7 ADD BY SONIA
Dim m_EndDate As String
'Add by Morgan 2004/7/1
Dim m_lngFeeDiscount As Long  '待繳年費減免金額
'Add by Morgan 2004/7/22
Dim m_DiscType As String   '減免身分
Dim m_bolActive As Boolean 'Active事件是否已觸發
Dim m_bolConfirm As Boolean '無專用期確認
Dim m_bolNeedReasign As Boolean '是否需要重新委任
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/23 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_bolFMP As Boolean 'Add by Morgan 2010/1/7
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/20 是否為寰華
'Add by Morgan 2010/1/20
Dim m_lngRefund As Long '未退預繳年費金額
Dim m_strFromYear As String, m_strToYear As String '預繳年費起迄
'Add by Morgan 2011/12/15
Public m_bolBeCalled As Boolean
Public m_CP01 As String
Public m_CP02 As String
Public m_CP03 As String
Public m_CP04 As String
Public m_CP09 As String
Dim m_Subject As String
Dim m_bolAutoMail As Boolean 'Added by Morgan 2012/3/22
Dim m_bol2ndY As Boolean 'Added by Morgan 2012/10/3 次年是否為補繳
Dim m_bolChgAgent As Boolean 'Added by Morgan 2012/11/15 自動更換代理人
Dim m_AD1516(5, 3) As String 'Added by Morgan 2013/3/25 中小企業減免資格
Dim m_UseNewForm As Boolean 'Added by Morgan 2013/3/25 使用新申請書
Dim m_str414CP09 As String 'Added by Morgan 2103/6/25 回復原狀收文號
Dim m_strFeeDate As String 'Added by Morgan 2020/2/6 大陸年費逾期繳費期限
Dim stCP12 As String, stCP13 As String

'Add by Morgan 2010/7/27
Private Sub StartLetter1(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(1 To 50) As String, strTmp As String, strTmp2 As String
   Dim ii As Integer
   Dim iAppCnt As Integer
   Dim intJ As Integer

   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   '專利種類
   '勾選1~3
   For intI = 1 To 3
      If pa(8) = Format(intI) Then
         strTmp = "■ "
      Else
         strTmp = "□ "
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','勾選" & Format(intI) & "','" & strTmp & "')"
   Next
   
   '減免身分
   '勾選4~6
   For intI = 1 To 3
      If InStr(m_DiscType, Format(intI)) > 0 Then
         strTmp = "■ "
      Else
         strTmp = "□ "
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','勾選" & Format(intI + 3) & "','" & strTmp & "')"
   Next
   
   'Added by Morgan 2013/3/25
   '中小企業可減免
   If InStr(m_DiscType, "3") > 0 And m_UseNewForm = True Then
      For intI = 1 To 5
         If pa(25 + intI) <> "" Then
            If txtAD(intI) = "1" Then
               strTmp = "■ 自然人"
            ElseIf txtAD(intI) = "2" Then
               strTmp = "■ 學校"
            ElseIf txtAD(intI) = "3" Then
               strTmp = "■ 中小企業"
            End If
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','減免身分" & intI & "','" & strTmp & "')"
            If txtAD(intI) = "3" Then
               'Modify By Sindy 2020/7/27
               'For intJ = 1 To 4
               For intJ = 5 To 6
               '2020/7/27 END
                  If Val(m_AD1516(intI, 1)) = intJ Then
                     strTmp = "■ "
                     strTmp2 = Format(m_AD1516(intI, 2), "#,###")
                  Else
                     strTmp = "□ "
                     strTmp2 = "　　　"
                  End If
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','勾選" & intI & intJ & "','" & strTmp & "')"
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','數值" & intI & intJ & "','" & strTmp2 & "')"
               Next
            End If
         End If
      Next
   End If
   'end 2013/3/25

   '減免年次
   If Val(Text5(1)) > 6 Then
      strTmp = Text5(0) & " 年至第 6"
   Else
      strTmp = Text5(0) & " 年至第 " & Text5(1)
   End If
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','減免年次','" & strTmp & "')"
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','減免金額','" & m_lngFeeDiscount & "')"
   
   '申請人數
   iAppCnt = 1
   For intI = 27 To 30
      If pa(intI) <> "" Then
         iAppCnt = iAppCnt + 1
      End If
   Next
   strTmp = Format(iAppCnt)
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人數','" & strTmp & "')"

   If Text5(0).Text = Text5(1).Text Then
      strTmp = "第 " & Text5(0) & " 年"
   Else
      strTmp = "第 " & Text5(0) & " 至 " & Text5(1) & " 年"
   End If
      
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','第幾年至幾年','" & strTmp & "')"
   'Modified by Lydia 2025/06/12 整理定稿：調整為法定期限m_NEXTNP08=>Text5(7)---杜協理
   If Text5(7) <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','下次繳年費日','" & DBDATE(Text5(7)) & "')"
   End If
   'end 2025/06/12
    ii = ii + 1
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
       "','規費','" & m_strOfficalFee & "')"

   'Add by Morgan 2011/8/5
   If Text5(2) = "Y" Then
      strExc(1) = "逾期要印"
      'Added by Morgan 2013/6/25
      If m_str414CP09 <> "" Then
         strExc(1) = "有回復原狀要印"
      End If
      'end 2013/6/25
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','" & strExc(1) & "','♀')"
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(1 To 25) As String, strTmp As String
   'Add By Cheng 2003/01/13
   Dim ii As Integer

   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   'Modified by Morgan 2012/11/19 +維持費 606
   'Modify by Amy 2018/03/20 +年費移作次年 612
   If cp(10) = "605" Or cp(10) = "606" Or cp(10) = "612" Then
      'Modify by Morgan 2009/2/18 香港短期和外觀改輸入第?次(P-75730)
      If pa(9) = "013" And pa(8) <> "1" Then
         strTmp = "續期費" '不必帶次數--玲玲
      Else
      'end 2009/2/18
         If Text5(0).Text = Text5(1).Text Then
            strTmp = "第 " & Text5(0) & " 年"
         Else
            strTmp = "第 " & Text5(0) & " 至 " & Text5(1) & " 年"
         End If
         
         If cp(10) = "606" Then
               strTmp = strTmp & "維持費"
         Else
            strTmp = strTmp & "年費"
         End If
         
      End If
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','第幾年至幾年費','" & strTmp & "')"
         
      'Added by Morgan 2020/2/5
      '大陸發文日過法限定稿要帶繳費期限(發文日後的最近一個法限日) Ex:P112461 發文日:2018.3.2 原法限:2018.2.7 繳費期限:2018.3.7--玲玲
      If m_strFeeDate <> "" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','繳費期限','" & m_strFeeDate & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','逾期才印','♀')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','逾期不印','♀')"
      End If
      'end 2020/2/5
         
      'Add by Morgan 2010/12/17
      'Modified by Morgan 2012/3/22
      'If m_bolBeCalled Then
      If m_bolBeCalled Or m_bolAutoMail Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','自動發文','♀')"
      End If
   End If
   
   If m_NP09 <> "" Then
   ii = ii + 1
   '92.2.5 modify by sonia
   'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '   "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
   '   "','年費法定期限','" & Label2(13).Caption & "')"
   ' 大陸案原法定期限
   '93.3.5 MODIFY BY SONIA 法定未逾期時印法定期限, 逾期才印原法定加半年
   'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '   "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
   '   "','年費法定期限','" & m_NP09 & "')"
      If m_NP09 < DBDATE(Text5(4)) Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','年費法定期限','" & DBDATE(Label2(13).Caption) & "')"
      Else
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','年費法定期限','" & m_NP09 & "')"
      End If
   End If
   '93.3.5 END
   '92.2.5 end
    'Add By Cheng 2003/01/10
    'Modified by Lydia 2025/06/12 整理定稿：調整為法定期限m_NEXTNP08=>Text5(7)---杜協理
   If Text5(7) <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','下次繳年費日','" & DBDATE(Text5(7)) & "')"
   End If
   'end 2025/06/12
   'Add by Amy 2018/03/20
   If cp(10) = "612" Then m_strOfficalFee = Val(textCP84)
    ii = ii + 1
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
       "','規費','" & m_strOfficalFee & "')"
      
   'Add by Morgan 2007/6/23
   If m_bolNeedReasign = True Then
      strExc(0) = GetAgentName
      If strExc(0) <> "" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','同時辦理事項','本案同時辦理重新委任「" & strExc(0) & "」為專利代理人，檢附委任書。')"
      End If
   End If
   
   If m_lngRefund > 0 Then
      If m_strFromYear = m_strToYear Then
         strExc(1) = m_strFromYear
      Else
         strExc(1) = m_strFromYear & "年至第" & m_strToYear
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','預繳起迄年','" & strExc(1) & "')"
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','可退金額','" & m_lngRefund & "')"
   End If
   
   'Added by Morgan 2012/11/15
   If m_bolChgAgent = True Then
      ii = ii + 1
       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
          "','有改代理人才印','♀')"
    End If
    'end 2012/11/15
    
   'Add by Morgan 2016/11/3
   If m_str414CP09 <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','有回復原狀要印','♀')"
   End If
   'end 2016/11/3
       
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(ii, strTxt) Then
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Add by Amy 2014/08/29 通知函
Private Sub StartLetter2(ByVal ET01 As String, ByVal ET03 As String)
    Dim strTxt(1 To 2) As String, strTmp As String
    Dim ii As Integer
    
    ii = 0
    EndLetter ET01, strReceiveNo, ET03, strUserNum
    'Modify by Amy 2018/03/20 +年費移作次年 612
    If cp(10) = "605" Or cp(10) = "612" Then
         If Text5(0).Text = Text5(1).Text Then
            strTmp = "第 " & Text5(0) & " 年"
         Else
            strTmp = "第 " & Text5(0) & " 至 " & Text5(1) & " 年"
         End If
         
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','第幾年至幾年費','" & strTmp & "年費')"
         'Modified by Lydia 2025/06/12 整理定稿：調整為法定期限m_NEXTNP08=>Text5(7)---杜協理
         If Text5(7) <> "" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','下次繳年費日','" & DBDATE(Text5(7)) & "')"
         End If
         'end 2025/06/12
         If Not ClsLawExecSQL(ii, strTxt) Then
            MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
         End If
    End If
        
    
End Sub
'end 2014/08/29

Private Function GetAgentName() As String
   Dim ii As Integer, strName As String
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         strName = strName & "、" & lstNameAgent.List(ii)
      End If
   Next
   If strName <> "" Then
      strName = Mid(strName, 2)
   End If
   GetAgentName = strName
End Function

Public Function Process(Index As Integer) As Boolean

   Dim strTmp As String, bolChk As Boolean, i As Integer
   Dim strContent As String, strMailTo As String 'Add by Morgan 2010/12/16
   Dim strPath As String, strFile As String 'Added by Morgan 2016/3/30
   'Added by Lydia 2020/04/07
   Dim strFilePath As String '記錄智慧局收文文號
   Dim bolUp As Boolean '是否需要上傳檔案到卷宗區
   Dim strNewCP64 As String '保留進度備註

   For i = 0 To 8
      If Not ChgType(i) Then
         If Me.Text5(i).Enabled Then Text5(i).SetFocus
         Exit Function
      End If
   Next
      'Modify By Cheng 2003/01/14
      '若下次繳費日不超過專用期止日才要檢查
      If blnOverDate = False Then
          ' 90.07.10 modify by louis
          If IsEmptyText(Text5(7)) = True Then
             MsgBox "請輸入下次繳費日", vbOKOnly + vbCritical, "檢核資料"
             If Me.Text5(7).Enabled Then Text5(7).SetFocus
             Exit Function
          End If
      End If
   'Add By Cheng 2002/03/08
   '檢查輸入資料的完整性
   If CheckDataIntegrity = False Then Exit Function
   'Add By Cheng 2002/05/22
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Function
   
   'Add by Morgan 2006/9/19 檢查下次繳費期限不可小於系統日--請作單
   If IsEmptyText(Text5(7)) = False Then
      If TransDate(Text5(7), 2) < strSrvDate(1) Then
         MsgBox "本案繳費年度有誤，請確認！", vbInformation
         Exit Function
      End If
   End If
   
   strNewCP64 = Text5(8).Text  'Added by Lydia 2020/04/07 保留進度備註
   'Added by Lydia 2020/04/08 電子送件要在發文前，先產生申請書；所以發文不用印
   If txtCP118 = "Y" Then Text8(0) = "N"

   'Modify By Cheng 2003/04/16
   '若申請國家為台灣
   If pa(9) = 台灣國家代號 Then
      'Add By Cheng 2003/04/15
      '檢查計算出的規費與進度檔的規費是否相同
      If cp(10) = 年費 Then
         'Modify by Morgan 2004/6/25
         '若發文日>=930701 且 系統日<930701 則用新法規費
         If Val(Text5(4).Text) >= 930701 And Val(strSrvDate(2)) < 930701 Then
            If ChkPatentYearFee(pa(9), pa(8), "Y00000002", cp(10), Me.Text5(0).Text, Me.Text5(1).Text, IIf(Me.Text5(2).Text = "Y", True, False)) = False Then Exit Function
         Else
            If ChkPatentYearFee(pa(9), pa(8), "Y00000001", cp(10), Me.Text5(0).Text, Me.Text5(1).Text, IIf(Me.Text5(2).Text = "Y", True, False)) = False Then Exit Function
         End If
      
      'Added by Morgan 2020/11/11
      ElseIf cp(10) = "612" Then
         m_strOfficalFee = Val(textCP84)
      'end 2020/11/11
      End If
   End If
   
   'Add by Morgan 2009/3/23 設定是否算發文室案件
   If pa(9) = 台灣國家代號 Then
      'Added by Lydia 2020/04/07 年費發文開放可電子送件
      m_CP09s = "": m_CP123s = ""
      If Frame1.Visible = True And txtCP118 = "Y" Then
             '電子送件也要記錄主管機關
             If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text5(4), , True) = False Then
                Exit Function
             End If
    
             strExc(0) = InputBox("請輸入智慧局收文文號!!")
             If strExc(0) = "" Then
                Exit Function
             Else
                strFilePath = strExc(0)  '記錄智慧局收文文號
                strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & Text5(8).Text  '保留進度備註
             End If
      Else
      'end 2020/04/07
            'Add by Morgan 2009/4/28
            If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text5(4)) = False Then
               Exit Function
            End If
            If m_CP123s = "Y" Then
            'end 2009/4/28
               'modify by sonia 2014/6/23 加傳發文規費, P-108903
               If ModifyDispatch(cp(9), m_CP09s, m_CP123s, m_strOfficalFee, Text5(4)) = False Then
                   Exit Function
               End If
            End If
      End If 'Added by Lydia 2020/04/07
      
      'Add by Amy 2014/10/14 P台灣案發文控制
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
        'Modified by Morgan 2016/6/29 不限制P--玲玲確認
        'If pa(1) = "P" And cp(9) < "C" Then
        If cp(9) < "C" Then
        'end 2016/6/29
            If cp(9) < "B" Then
                'A類一定要有接洽單才可發文
                'Modify by Amy 2014/11/27 取消ChkOneDayHasCP27判斷,接洽單改檢查,因考慮可能同時發文其他案件性質情形
                'If PUB_CheckPDF2(cp(9), 0, True, strExc(0)) = False And ChkOneDayHasCP27(pa(1), pa(2), pa(3), pa(4), cp(5) + 19110000) = False Then
                If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
                    Exit Function
                End If
            End If
            'AB類申請書確認檢查,符合條件才可發文
            'Modified by Morgan 2015/3/17
            'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And Text8(0) = "N" And PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            'Modifiedb by Lydia 2020/04/08 排除電子送件
            'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And Text8(0) = "N" Then
            If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And Text8(0) = "N" And txtCP118 <> "Y" Then
               If PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            'end 2015/3/17
                  MsgBox "無申請書PDF檔 ,不可發文!", vbInformation
                  Exit Function
               End If 'Added by Morgan 2015/3/17
            End If
        End If
      End If
      'end 2014/10/14
      
   'Added by Morgan 2016/6/29 非臺灣案電子化
   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      If cp(9) < "B" And Left(cp(12), 1) <> "F" Then
          If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
              Exit Function
          End If
      End If
   'end 2016/6/29
   End If
   
   'Added by Morgan 2012/3/22
   m_bolAutoMail = False
   'Modified by Morgan 2014/6/4 + 改與整批發文相同判斷非台灣
   'If pa(9) = "020" And Not m_bolBeCalled Then
   'Modified by Lydia 2015/06/01 外專人員不彈訊息也不發email
   If Left(Pub_StrUserSt03, 1) <> "F" Then
        If pa(9) <> "000" And Not m_bolBeCalled And Text5(3) <> "Y" Then
        'end 2014/6/4
           'Modified by Morgan 2016/5/20 指示信電子化,不直接寄送則開啟編輯畫面
           If MsgBox("是否直接發 E-Mail 給代理人??" & vbCrLf & vbCrLf & "(選否將開啟編輯畫面)", vbYesNo + vbDefaultButton2) = vbYes Then
              m_bolAutoMail = True
           Else
              Text5(3) = "Y"
           End If
        End If
   'end 2012/3/22
   End If
   
    'Added by Lydia 2020/04/07 檢查是否有電子送件的檔案
    bolUp = False
    If Frame1.Visible = True And txtCP118.Text = "Y" And strFilePath <> "" Then
        strExc(1) = cp(82)
        If Val(cp(82)) > 0 Then
            If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                 strExc(1) = ""
            End If
        End If
        If Val(strExc(1)) = 0 Then
           '先判斷是否上傳檔案
           If Pub_AutoEsetToCppByP(True, pa(1), pa(2), pa(3), pa(4), pa(8), "", cp(10), strFilePath, Text5(4).Text) = False Then
                 Exit Function
           End If
           bolUp = True
        End If
    End If
    Text5(8).Text = strNewCP64 '檢查完畢，更新備註欄位
    'end 2020/04/07
   
   If FormSave = False Then
      MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
      Exit Function
   End If
   
   Process = True
   
   'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
   PUB_CheckEMail Combo2
   PUB_CheckEMail pa(75), pa(144)
   If pa(145) <> "" Then
      PUB_CheckEMail pa(75), pa(145)
   End If
   'end 2008/2/20
   
   If Text5(9) = "Y" Then
      bolChk = True
   Else
      bolChk = False
   End If
   'Add by Morgan 2007/6/14
   m_bolNeedReasign = False
   If pa(9) = "000" Then
      m_bolNeedReasign = PUB_IsLatestAgent(pa(1), pa(2), pa(3), pa(4), strReceiveNo)
      If m_bolNeedReasign = True Then
         PUB_ReAsignInform pa(1), pa(2), pa(3), pa(4), strReceiveNo, True
      End If
   End If
   
   '2012/7/23 add by sonia 檢查計算出的規費與進度檔的規費不同,仍繼續發文時發mail給智權人員
   If pa(9) = "000" Then
      If Val(m_strOfficalFee) <> Val(cp(17)) Then
         '2013/7/2 modify by sonia 改用共用module
         'Modified by Lydia 2020/04/07 +是否電子送件
         'PUB_ChkOfficialFee cp(9), m_strOfficalFee
         PUB_ChkOfficialFee cp(9), m_strOfficalFee, IIf(txtCP118 = "Y", "A", "")
      End If
   End If
   '2012/7/23 end
         
   If Text8(2) <> "N" Then '通知函
      If pa(9) = 台灣國家代號 Then '台灣 00
        'Modify by Amy 2014/09/16
        'Modify by Amy 2018/03/20 移作次年612使用年費605定稿
         If cp(10) = "605" Or cp(10) = "612" Then
            'Modify by Amy 2014/09/19 +if 大對台定稿
            If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
                strTmp = "31"
            Else
                strTmp = "30"
            End If
            'end 2014/09/19
            StartLetter2 "02", strTmp
         Else
            strTmp = "00"
            StartLetter "02", strTmp
         End If
         'end 2014/09/16
      ElseIf pa(9) <> 台灣國家代號 Then
         strTmp = "01"
         StartLetter "02", strTmp
      End If
      'StartLetter "02", strTmp 'Modify by Amy 2014/09/09 往上搬
      'Modify by Amy 2014/08/29 +傳strLetterRecNo
      NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
   End If
   
   If Text5(3) = "Y" Then
      bolChk = True
   Else
      bolChk = False
   End If
   If Text8(0) <> "N" Then '指示信/申請書
      If pa(9) = 台灣國家代號 Then '台灣 07
         'Modify By Cheng 2003/01/13
         '預設未逾期的處理狀況
         strTmp = "07"
         
         If cp(10) <> "612" Then 'Added by Morgan 2020/11/11 移作次年目前只有一種申請書(人工調整內容)
            
   'Modify by Morgan 2008/10/8 改判斷是否雙倍
   '               '92.2.5 MODIFY BY SONIA
   '               '若有法定期限
   '               'If Me.Label2(13).Caption <> "" Then
   '               '   '若發文日大於法定期限(逾期)
   '               '   If Val(DBDATE(Text5(4).Text)) > Val(DBDATE(Me.Label2(13).Caption)) Then
   '               '      strTmp = "08"
   '               '   End If
   '               'End If
   '               strExc(0) = "SELECT MAX(NP09) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
   '                  " AND NP06 IS NOT NULL AND NP07='" & cp(10) & "' And NP09 Is Not Null "
   '               intI = 1
   '               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   '               If RsTemp.Fields(0) <> "" Then
   '                  '若發文日大於原法定期限(逾期)
   '                    'Modify By Cheng 2003/09/01
   '                    '若法定期限為假日則用大於法定期限最近的工作日與發文日比較
   ''                  If Val(DBDATE(Text5(4).Text)) > Val(DBDATE(rsTemp.Fields(0))) Then
   '                  If DBDATE(Text5(4).Text) > IIf(DBDATE(RsTemp.Fields(0)) >= DBDATE(m_strNP09_1), DBDATE(RsTemp.Fields(0)), DBDATE(m_strNP09_1)) Then
   '                     strTmp = "08"
   '                  End If
   '               End If
   '               '92.2.5 END
            If Text5(2) = "Y" Then
               strTmp = "08"
            End If
   'end 2008/10/8
            
            'Add by Morgan 2010/1/21
            '有預繳年費
            If m_lngRefund > 0 Then
               strTmp = "09"
               StartLetter "01", strTmp
            Else
            'end 2010/1/21
               'Modify by Morgan 2004/7/1
               '若符合年費減免則印新申請書
               If cp(81) = "Y" And Val(Text5(0).Text) < 7 Then
                  'Modify by Morgan 2010/7/27 取消套印
                  'PrintLetter bolChk
                  
                  'Modified by Morgan 2013/3/25 若有中小企業則跑新申請書(要印減免資格)
                  'strTmp = "10"
                  If InStr(m_DiscType, "3") > 0 And m_UseNewForm = True Then
                     '多人申請
                     If pa(27) <> "" Then
                        strTmp = "12"
                     Else
                        strTmp = "11"
                     End If
                  Else
                     strTmp = "10"
                  End If
                  'end 2013/3/25
                  
                  StartLetter1 "01", strTmp
                  'end 2010/7/27
               Else
                  StartLetter "01", strTmp
               End If
            End If
            
         End If 'Added by Morgan 2020/11/11
      
         'Modify by Amy 2014/08/29 +傳strLetterRecNo
         NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
         'Add by Amy 2014/09/09 申請書修改改開1105_1 for P台灣案電子化
         If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And Text8(0) <> "N" And Text5(3) = "Y" Then
                frm1105_1.m_RecNo = strReceiveNo
                'Modify By Sindy 2022/5/11 流水號要足6碼
                frm1105_1.m_PdfName = Text1 & Text2 & IIf(Text3 & Text4 = "000", "", "-" & Text3 & "-" & Text4) & "." & cp(10) & ".DATA.PDF"
                frm1105_1.Show
         End If
         'end 2014/09/09
         
         'Modify by Morgan 2010/12/16
         If m_bolBeCalled Then
            PUB_PrintLetter strReceiveNo, True
         End If
         
      Else
         'Modify by Morgan 2008/3/18 +澳門
         'If pa(9) = "020" Then
         If pa(9) = "020" Or pa(9) = "044" Then
            strTmp = "36" '大陸未曾結案 36
            If m_bolChgAgent = False Then 'Added by Morgan 2012/11/15
               '92.3.14 MODIFY BY SONIA
               'strExc(0) = "SELECT MAX(NP09) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
               '   " AND NP06 IS NOT NULL AND NP11 IS NOT NULL AND NP07='" & cp(10) & "'"
                'Modify By Cheng 2003/04/24
   '               strExc(0) = "SELECT MAX(NP09) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
   '                  " AND NP06 IS NOT NULL AND NP11 IS NOT NULL AND NP07='" & cp(10) & "' AND NP09<=" & cp(7)
               'Modify by Morgan 2004/3/11
               '改用相關總收文號判斷是否曾結案(解除期限)
               'strExc(0) = "SELECT MAX(NP09) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                  " AND NP06 IS NOT NULL AND NP11 IS NOT NULL AND NP07='" & cp(10) & "' AND NP09<=" & DBDATE(m_strNP09)
               '92.3.14 END
               'Modify by Morgan 2007/12/24
               '要判斷有相關總收文號才要作否則語法會錯
               'Modified by Morgan 2016/4/7 不續辦管制半年的期限相關收文號是不續辦那道,與原期限不同都抓不到資料,改判斷有一年內的不續辦期限(Ex. P-91131)
               'If cp(43) <> "" Then
                  'strExc(0) = "SELECT MAX(NP09) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                     " AND NP06 IS NOT NULL AND NP11 IS NOT NULL AND NP07='" & cp(10) & "' AND NP09<=" & DBDATE(m_strNP09) & _
                     " AND NP01='" & cp(43) & "'"
                  'Modified by Morgan 2016/4/8 要用原法定期限m_strNP09_2判斷,不可抓下一工作天m_strNP09_1(Ex. P-95175)
                  strExc(0) = "SELECT MAX(NP09) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                     " AND NP06 IS NOT NULL AND NP11 IS NOT NULL AND NP07='" & cp(10) & "' AND NP09<=" & DBDATE(m_strNP09) & " and NP09>=" & Val(m_strNP09_2)
               'end 2016/4/7
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If RsTemp.Fields(0) > 0 Then '大陸曾結案 37  cp(7)
                     strTmp = "37"
                     m_NP09 = DBDATE(RsTemp.Fields(0))
                  End If
               'End If
            End If
         'Add by Morgan 2007/9/11 香港指示信
         ElseIf pa(9) = "013" Then
            strTmp = "40"
         End If
         StartLetter "02", strTmp
         
         'Modify by Morgan 2010/12/16
         'NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, 0
         'Modified by Morgan 2012/3/22
         'If m_bolBeCalled Then
         'Modified by Morgan 2016/3/30
         '指示信電子化
         'If m_bolBeCalled Or m_bolAutoMail Then
         '   strMailTo = PUB_GetCFAgentEMail(strReceiveNo)
         '   'Modify by Amy 2014/08/29 +傳strLetterRecNo
         '   NowPrint strReceiveNo, "02", strTmp, False, strUserNum, 0, , True, strContent, 1, , , , , True, , , strReceiveNo
         '   'Modified by Morgan 2014/5/19 mail 內容 word 格式未轉換(Ex 底線...),改加轉tag函數並用HTML方式寄發
         '   'PUB_SendMail strUserNum, strMailTo, "", m_Subject, strContent, , , , , , , "patent"
         '   PUB_SendMail strUserNum, strMailTo, "", m_Subject, ChgHTMLFormat(strContent), , , True, , , , "patent"
         'Else
         '   'Modify by Amy 2014/08/29 +傳strLetterRecNo
         '   NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
         'End If
         If m_bolBeCalled Or m_bolAutoMail Then
            strPath = "": strFile = ""
            '指示信轉pdf檔
            If PUB_PrintDocAsPdf(strReceiveNo, "02", strTmp, strReceiveNo, strPath, strFile) Then
               'FMP案還是要印紙本
               'Removed by Morgan 2016/5/20 不印指示信,只印寄件備份就好--玲玲
               'If m_bolFMP Then
               '   PUB_PrintPDF strPath & "\" & strFile
               'End If
               'end 2016/5/20
               
               '上傳卷宗區
               If PUB_UploadOrderLetter(strPath & "\" & strFile, strReceiveNo) Then
                  'EMail
                  PUB_SendOrderLetterP strReceiveNo, m_Subject, True
               End If
            End If
         Else
            If Left(Pub_StrUserSt03, 1) = "F" Then
               NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum
            Else
               NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, , , , , , , , , , , , , strReceiveNo
               If bolChk Then
                  frm1105_1.m_RecNo = strReceiveNo
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(Text1, Text2, Text3, Text4) & "." & cp(10) & ".DATA.PDF"
                  frm1105_1.m_Subject = m_Subject
                  frm1105_1.Show
               End If
            End If
         End If
         'end 2016/3/30
         
      End If
   End If
      
   'Added by Lydia 2020/04/07 是否可以上傳檔案,前面已判斷
   If bolUp = True Then
      If Pub_AutoEsetToCppByP(False, pa(1), pa(2), pa(3), pa(4), pa(8), cp(9), cp(10), strFilePath, Text5(4).Text) = False Then
           Exit Function
      End If
   End If
   'end 2020/04/07
End Function

Private Sub cmdok_Click(Index As Integer)
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   Select Case Index
      'Modify by Morgan 2004/9/9 將同時發文併入
      Case 0, 3 '確定
         'Modify by Morgan 2010/2/10 改呼叫函數方式以便鎖定按鍵
         cmdOK(Index).Enabled = False
         If Not Process(Index) Then
            cmdOK(Index).Enabled = True
         Else
            'Modify by Morgan 2004/9/9 將同時發文合併
            If Index = 3 Then
               'Add By Sindy 2013/5/20
               If frm040104_1.bolIsEMPFlow = True Then
                  frm090202_4.QueryData
               End If
               '2013/5/20 End
               frm040104_1.Show
               frm040104_1.ReQuery
            Else
'               'Add By Cheng 2002/04/30
'               '若有未發文資料顯示警告
'               PUB_GetCPunIssueDatas "" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)
'               frm040104_1.Show
'               ' 90.07.11 modify by louis (回第一個畫面清除)
'               frm040104_1.Clear
               '若有未發文資料顯示警告
               If PUB_GetCPunIssueDatas("" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)) Then
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
         Set frm060104_b.oParent = Me 'Add by Morgan 2011/10/5
         frm060104_b.LoadMe pa(1), pa(2), pa(3), pa(4), 4
         Me.Hide
      Case 4
         Me.Hide
         Set frm06010303_1.oParent = Me 'Add by Morgan 2011/10/5
         frm06010303_1.LoadMe strReceiveNo, pa(1), pa(2), pa(3), pa(4), 44
         frm06010303_1.Text41 = "N"
         frm06010303_1.Caption = "內專發文-變更事項"
        m_blnClkChgEvnBtn = True
   End Select
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
End Sub

Private Function FormSave() As Boolean
   Dim i As Integer
   Dim strTmp(0) As String, iMax As Long, strTxt(1 To 20) As String
   Dim strFLD As String
   Dim nMaxNo As String
   Dim nPos As Integer
   Dim aryCurr As Variant
   Dim aryAll As Variant
   Dim aryDate As Variant
   Dim nPosBegin As Integer
   Dim nPosEnd As Integer
   Dim nDot As Integer
   Dim ii As Integer
   Dim stCP118 As String, stCP152 As String 'Added by Lydia 2020/04/07 電子送件和自動扣款日
   Dim stNP23 As String 'Added by Lydia 2025/10/29
   
On Error GoTo ErrHnd
'Add By Cheng 2002/11/06
FormSave = True
cnnConnection.BeginTrans

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

   ' 90.07.18 modify by louis
   If IsEmptyText(Text5(6)) = False Then
      Text5(6) = Text5(6) & String(9 - Len(Text5(6)), "0")
   End If
   
   ' 90.07.25 modify by louis
   aryAll = Split(strCaseFee(2), ",")
   aryCurr = Split(pa(72), ",")
      
   'Add by Morgan 2009/2/18 香港短期和外觀改輸入第?次(P-75730)
   If pa(9) = "013" And pa(8) <> "1" Then
      nPosBegin = Val(Text5(0)) - 1
      nPosEnd = Val(Text5(1)) - 1
   Else
   'end 2009/2/18
   
      ' 找尋繳年費起始點位置
      nPosBegin = 0
      For nPos = 0 To UBound(aryAll)
         If aryAll(nPos) = Text5(0) Then
            nPosBegin = nPos
            Exit For
         End If
      Next nPos
      ' 找尋繳年費終止點位置
      nPosEnd = 0
      For nPos = 0 To UBound(aryAll)
         If aryAll(nPos) = Text5(1) Then
            nPosEnd = nPos
            Exit For
         End If
      Next nPos
      
   End If
   ' 組繳年費年度字串
   strFLD = Empty
   For nPos = 0 To nPosEnd
      If nPos > 0 Then: strFLD = strFLD & ","
      strFLD = strFLD & aryAll(nPos)
   Next nPos
   
   ' 計算逗號的總數(幾格)
   nDot = 0
   For nPos = 1 To Len(pa(72))
      If Mid(pa(72), nPos, 1) = "," Then nDot = nDot + 1
   Next nPos
   
   If nDot > nPosEnd Then
      strFLD = strFLD & String(nDot - nPosEnd, ",")
   End If
   pa(72) = strFLD
   
   ' 重新計算共有幾欄
   nDot = 0
   For nPos = 1 To Len(pa(72))
      If Mid(pa(72), nPos, 1) = "," Then nDot = nDot + 1
   Next nPos
   
   ' 繳年費日期
   ReDim aryCurr(nDot)
   'Modify by Morgan 2007/3/8 只繳過一年的(沒有",")也要考慮否則第一次的資料會被清掉
   'If InStr(pa(73), ",") > 0 Then
      aryDate = Split(pa(73), ",")
      ' 拷貝原資料
      For nPos = 0 To UBound(aryDate)
         If IsEmptyText(aryDate(nPos)) = False Then
            If nDot > 0 Then
               aryCurr(nPos) = aryDate(nPos)
            End If
         End If
      Next nPos
   'End If
   ' 填入新資料
   For nPos = nPosBegin To nPosEnd
      aryCurr(nPos) = DBDATE(Text5(4))
   Next nPos
   ' 讀取新資料
   strFLD = Empty
   For nPos = 0 To UBound(aryCurr)
      If nPos > 0 Then: strFLD = strFLD & ","
      strFLD = strFLD & aryCurr(nPos)
   Next nPos
   pa(73) = strFLD
   
   '費用是否雙倍
   ReDim aryCurr(nDot)
   'Modify by Morgan 2007/3/8 只繳過一年的(沒有",")也要考慮否則第一次的資料會被清掉
   'If InStr(pa(74), ",") > 0 Then
      Dim aryFee As Variant
      aryFee = Split(pa(74), ",")
      ' 拷貝原資料
      For nPos = 0 To UBound(aryFee)
        If IsEmptyText(aryFee(nPos)) = False Then
           If nDot > 0 Then
              aryCurr(nPos) = aryFee(nPos)
           End If
        End If
      Next nPos
   'End If
   
   ' 填入新資料
   For nPos = nPosBegin To nPosEnd
      'Modify by Morgan 2007/6/11 起始年上Y就好,Ex.P84273
      'If Text5(2) = "Y" Then
      If Text5(2) = "Y" And nPos = nPosBegin Then
      'end 2007/6/11
         aryCurr(nPos) = "Y"
      'Added by Morgan 2012/10/3 次年補繳也要紀錄
      ElseIf m_bol2ndY = True And nPos = nPosBegin + 1 Then
         aryCurr(nPos) = "Y"
      'end 2012/10/3
      Else
         aryCurr(nPos) = Empty
      End If
   Next nPos
   ' 讀取新資料
   strFLD = Empty
   For nPos = 0 To UBound(aryCurr)
      If nPos > 0 Then: strFLD = strFLD & ","
      strFLD = strFLD & aryCurr(nPos)
   Next nPos
   pa(74) = strFLD

   strTmp(0) = ""
   strTxt(1) = "UPDATE PATENT SET PA05=" & CNULL(Text15(0)) & ",PA06=" & CNULL(ChgSQL(Text15(1))) & _
      ",PA07=" & CNULL(Text15(2)) & ",PA76=" & CNULL(ChgSQL(Text5(6))) & "," & _
      "PA72='" & pa(72) & "',PA73='" & pa(73) & "'," & _
      "PA74='" & pa(74) & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   cnnConnection.Execute strTxt(1)
           
   If Combo2 <> "" Then
      'Modify by Morgan 2008/2/22
      'cp(44) = ChangeCustomerL(Combo2)
      intI = InStr(Combo2, "-")
      If intI > 0 Then
         cp(44) = Left(Combo2, intI - 1)
         cp(116) = Mid(Combo2, intI + 1)
      Else
         cp(44) = Combo2
         cp(116) = ""
      End If
      cp(44) = ChangeCustomerL(cp(44))
      'end 2008/2/22
      
      cp(45) = "" 'Added by Morgan 2012/11/15 一率先清除，避免重新發文殘留
      
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.GetCaseThatCode(cp) Then cp(45) = ""
      If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   Else
      cp(44) = ""
      cp(116) = ""
      cp(45) = ""
   End If
   cp(53) = Text5(0)
   cp(54) = Text5(1)
   'Add by Morgan 2010/8/17
   If pa(9) = "013" And pa(8) <> "1" Then
      cp(53) = TransHKData(Text5(0), True)
      cp(54) = cp(53)
   End If
   'end 2010/8/17
   
   'Added by Lydia 2020/04/07 電子送件和自動扣款日
   If Frame1.Visible = True Then
        stCP118 = txtCP118
        stCP152 = ""
        If txtCP118 = "Y" Then
           If txtPayToday <> "" Then
              stCP118 = "A"
              If txtPayToday = "Y" Then
                 stCP152 = CompWorkDay(2, DBDATE(Text5(4)))
              Else
                 stCP152 = CompWorkDay(3, DBDATE(Text5(4)))
              End If
           End If
        End If
   End If
   'end 2020/04/07
   
   ' 91.03.25 modify by louis (單引號)
   'Modify by morgan 2004/6/23 加 cp81
   'Modify by morgan 2004/8/11 加 cp84
   'Modify by Morgan 2005/7/15 加 cp110
   'Modify by Morgan 2010/3/9 +CP53,CP54
   'Modified by Morgan 2013/3/22 -CP14
   'Modified by Lydia 2020/04/07+CP118,CP152 電子送件和自動扣款日
   'Modified by Lydia 2021/05/25 +CP113工作時數
   'Modified by Lydia 2023/06/20 +CP14
   strTxt(2) = "UPDATE CASEPROGRESS SET CP27=" & CNULL(TransDate(Text5(4), 2)) & ",CP14=" & CNULL(cp(14)) & _
      ",CP44=" & CNULL(cp(44)) & ",CP116=" & CNULL(cp(116)) & _
      ",CP45=" & CNULL(ChgSQL(cp(45))) & ",cp64=" & CNULL(ChgSQL(Text5(8))) & _
      ",CP22=" & CNULL(Text8(1)) & ",cp81=" & CNULL(cp(81)) & ", cp84=" & Format(Val(m_strOfficalFee)) & _
      ",cp110=" & CNULL(cp(110)) & ",cp53='" & cp(53) & "',cp54='" & cp(54) & "'" & _
      ",cp118=" & CNULL(stCP118) & ", cp152=" & CNULL(stCP152) & " ,cp113=" & CNULL(txtCP113, True) & _
      " WHERE CP09='" & strReceiveNo & "'"
   cnnConnection.Execute strTxt(2), intI

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
   
    'Modify By Cheng 2003/01/14
    '若有下次繳費日時, 才要新增下一程序檔
    m_NEXTNP08 = ""
    If Me.Text5(7).Text <> "" Then
        CompNextFeeDate
         '計算本所期限
         If pa(9) = 台灣國家代號 Then
             'Added by Morgan 2014/10/28
             If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                strTmp(0) = PUB_GetOurDeadline(Text5(7))
             Else
             'end 2014/10/28
                strTmp(0) = CompDate(2, -2, TransDate(Text5(7).Text, 2))
             End If 'Added by Morgan 2014/10/28
         Else
             'Modify by Morgan 2010/1/7 FMP案所限改法限-10天
             'Modified by Morgan 2018/10/3 非FMP也改10天
             'If m_bolFMP Then
                strTmp(0) = CompDate(2, -10, TransDate(Text5(7).Text, 2))
             'Else
             '   strTmp(0) = CompDate(1, -1, TransDate(Text5(7).Text, 2))
             '   strTmp(0) = CompDate(2, -5, strTmp(0))
             'End If
             'end 2018/10/3
         End If
        strExc(2) = 年費
        
        'Added by Morgan 2012/10/23
        '香港維持費
        If pa(9) = "013" And cp(10) = 維持費 Then
            strExc(2) = 維持費
        End If
        'Added by Lydia 2025/10/29
        stNP23 = ""
        If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
            strTmp(0) = PUB_GetPOurDeadline(DBDATE(Text5(7)), pa(9), stNP23, pa(1), strExc(2))
        End If
        'end 2025/10/29
        iMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了 iMax = objPublicData.GetNextProgressNo
        '重抓智權人員
        '若本所期限非工作天則抓最近的工作天
        'Modified by Lydia 2025/10/29 +NP23
        strTxt(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22,NP23) " & _
           "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
           "','" & pa(4) & "','" & stCP13 & "'," & strExc(2) & "," & PUB_GetWorkDay1(strTmp(0), True) & "," & _
           TransDate(Text5(7), 2) & "," & iMax & "," & CNULL(stNP23, True) & ")"
         'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(3), intI
        m_NEXTNP08 = strTmp(0)
        '   iMax = iMax + 1
        iMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了 iMax = objPublicData.GetNextProgressNo
    End If
   
'Modify by Morgan 2011/8/19 改都要管制收達且管制比照其他案件性質模式--玲玲
'   If m_bolBeCalled Then 'Add by Morgan 2010/12/17 自動發文設定3天的收達管制(原來收費表沒設管制天數)
'      i = 3
'      strExc(0) = "SELECT CF23 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & cp(10) & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If Not IsNull(RsTemp.Fields(0)) Or RsTemp.Fields(0) <> 0 Then
'           iMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了 iMax = objPublicData.GetNextProgressNo
'            i = 4
'               'Modify By Cheng 2003/12/08
'               '若本所期限非工作天則抓最近的工作天
'            strTxt(i) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'               "NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & _
'               "','" & pa(3) & "','" & pa(4) & "'," & 收達 & "," & _
'               PUB_GetWorkDay1(CompDate(2, RsTemp.Fields(0), TransDate(Text5(4), 2)), True) & "," & _
'               CompDate(2, RsTemp.Fields(0), TransDate(Text5(4), 2)) & ",'" & _
'               strUserNum & "'," & iMax & ")"
'           cnnConnection.Execute strTxt(i)
'         End If
'      End If
'   End If
   If pa(9) <> "000" Then PUB_SetArriveDate strReceiveNo
'end 2011/8/19
   
   'Add by Amy 2014/09/09 for 台灣案電子化
    If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And pa(9) = 台灣國家代號 Then
        cnnConnection.Execute "delete LetterProgress where lp01='" & strReceiveNo & "'", intI 'Added by Morgan 2016/2/26 可能會重新發文
        '*沒出客戶通知函
        If Text8(2) = "N" Then
            'Modify by Amy 2015/02/13 原:判斷同一天 沒有其他有規費的發文
              '1.    電子送件且規費>0,有收據
              '2.非電子送件且經發文室要計件,有回執
            'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'            strExc(1) = PUB_GetLetterJudge(pa(1), cp(10))
'            If cp(118) = "Y" Then
'                If Val(m_strOfficalFee) > 0 Then
'                    PUB_AddLetterProgress strReceiveNo, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                End If
'            Else
'                If Left(m_CP123s, 1) = "Y" Then
'                    PUB_AddLetterProgress strReceiveNo, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                End If
'            End If
            
        '*有出客戶通知函
        Else
            'Modified by Morgan 2018/8/1
            'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , , pa(1), pa(2), pa(3), pa(4))
            strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10))
            'Modify by Amy 2015/02/13 修改、整理判斷條件
            'PUB_AddLetterProgress strReceiveNo, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
              '1.　電子送件有規費的有收據；無規費的無回執
              '2.非電子送件要計件的有回執；不計件的無回執
            'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'            If cp(118) = "Y" Then
'                If Val(m_strOfficalFee) > 0 Then
'                    PUB_AddLetterProgress strReceiveNo, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
'                Else
'                    PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
'                End If
'            Else
                If Left(m_CP123s, 1) = "Y" Then
                    PUB_AddLetterProgress strReceiveNo, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
                Else
                    PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
                End If
'            End If
            'end 2015/03/06
        End If
        '*有申請書
        If Text8(0) <> "N" Then
            If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
                 '新增申請書轉檔記錄
                 PUB_AddAppForm strReceiveNo
            End If
        End If
        
    'Added by Morgan 2016/3/25
    '指示信電子化
    ElseIf pa(9) <> 台灣國家代號 And Left(Pub_StrUserSt03, 1) <> "F" Then
      If Text8(0) <> "N" Then
         'Modified by Morgan 2016/12/13 若已有指示信表示有異常(Ex.P100930)
         'If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
         If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) Then
            Err.Raise 999, , "指示信記錄(Appform)已存在，若為之前有誤操作請通知電腦中心刪除該筆記錄！"
         Else
         'end 2016/12/13
            m_strFeeDate = "" 'Added by Morgan 2020/2/6
            strExc(2) = "" 'Added by Morgan 2020/2/6
            'Added by Morgan 2016/5/19 主旨要寫到Appfrom自StartLetter移來
            If pa(9) = "013" And pa(8) <> "1" Then
               strExc(1) = "續期費" '不必帶次數--玲玲
            Else
               If Text5(0).Text = Text5(1).Text Then
                  strExc(1) = "第 " & Text5(0) & " 年"
               Else
                  strExc(1) = "第 " & Text5(0) & " 至 " & Text5(1) & " 年"
               End If
               
               If cp(10) = "606" Then
                     strExc(1) = strExc(1) & "維持費"
               Else
                  strExc(1) = strExc(1) & "年費"
               End If
               'Added by Morgan 2020/2/5
               '大陸發文日過法限定稿要帶繳費期限(發文日後的最近一個法限日) Ex:P112461 發文日:2018.3.2 原法限:2018.2.7 繳費期限:2018.3.7--玲玲
               If pa(9) = "020" Then
                  strExc(3) = DBDATE(Text5(4)) '發文日
                  strExc(4) = DBDATE(m_strNP09_2) '原法限
                  If strExc(4) < strExc(3) Then
                     m_strFeeDate = strExc(4)
                     Do While (m_strFeeDate < strExc(3))
                        m_strFeeDate = AddMonth(m_strFeeDate, 1)
                     Loop
                     strExc(2) = "於" & TranslateKeyWord(incCNV_CHINESE_CUN1, m_strFeeDate, "") & "前"
                  End If
               End If
               'end 2020/2/5
            End If
            'Modified by Morgan 2020/2/6 +strExc(2)
            m_Subject = "請" & strExc(2) & "代為繳納" & strExc(1) & " Y/R:" & cp(45) & "; O/R:" & Text1 & "-" & Text2 & IIf(Text3 & Text4 = "000", "", "-" & Text3 & "-" & Text4)
            'end 2016/5/19
            'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
            strExc(2) = PUB_GetLetterJudgeNew("2", pa(1), cp(10), pa(9))
            PUB_AddAppForm strReceiveNo, True, strExc(2), m_Subject '不轉檔,自行判發
         End If
      End If
    'end 2016/3/25
    
      'Added by Morgan 2016/5/26 非臺灣案電子化
      '客戶通知函
      If 內專全面電子化啟用日 <= Val(strSrvDate(1)) Then
         If Text8(2) <> "N" Then
            'Modified by Morgan 2018/8/1
            'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , pa(9), pa(1), pa(2), pa(3), pa(4))
            strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10), pa(9), , , m_bolFMP)
            PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75)
         End If
      End If
      'end 2016/5/26
    End If
    'end 2014/09/09
    
   'Add by Morgan 2009/3/23
   If pa(9) = 台灣國家代號 Then
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130
      'Add by Amy 2015/02/13 更新收據/回執設定
      'Modify by Amy 2015/03/06 +發文日參數
      PUB_UpdateLP19 cp(1), cp(2), cp(3), cp(4), m_CP09s, m_CP123s, Text5(4)
      
      'Add by Morgan 2010/1/20
      If m_lngRefund > 0 Then
         'Modified by Morgan 2021/5/6
         'strSql = "update t99 set t14=" & DBDATE(Text5(4)) & " where t01='" & pa(1) & "' and t02='" & pa(2) & "' and t03='" & pa(3) & "' and t04='" & pa(4) & "'"
         'cnnConnection.Execute strSql, intI
         ''Add by Morgan 2011/6/17
         'strSql = "update t100 set t14=" & DBDATE(Text5(4)) & " where t01='" & pa(1) & "' and t02='" & pa(2) & "' and t03='" & pa(3) & "' and t04='" & pa(4) & "'"
         'cnnConnection.Execute strSql, intI
         
         strSql = "update t109 set t14=" & DBDATE(Text5(4)) & " where t01='" & pa(1) & "' and t02='" & pa(2) & "' and t03='" & pa(3) & "' and t04='" & pa(4) & "' and t14=0"
         cnnConnection.Execute strSql, intI
         
         If PUB_ChkCPExist(pa, 減免退費, 1, strExc(1)) = True Then
            strSql = "UPDATE CASEPROGRESS SET CP27=" & DBDATE(Text5(4)) & ",CP110=" & CNULL(cp(110)) & " WHERE CP09='" & strExc(1) & "'"
            cnnConnection.Execute strSql
         End If
         'end 2021/5/6
         
      'Added by Morgan 2021/6/8
      ElseIf cp(81) = "N" Then
         strSql = "update t109 set t14=19221111 where t01='" & pa(1) & "' and t02='" & pa(2) & "' and t03='" & pa(3) & "' and t04='" & pa(4) & "' and t14=0"
         cnnConnection.Execute strSql, intI
         
      End If
   End If
   
   'Added by Morgan 2013/6/25
   If m_str414CP09 <> "" Then
      strSql = "UPDATE CASEPROGRESS SET CP27=" & DBDATE(Text5(4)) & ",CP110=" & CNULL(cp(110)) & " WHERE CP09='" & m_str414CP09 & "'"
      cnnConnection.Execute strSql
   End If
   'end 2013/6/25
   
   cnnConnection.CommitTrans
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   FormSave = False
   
   MsgBox Err.Description, vbCritical 'Added by Morgan 2016/12/13
End Function

Private Sub Combo2_Click()
   Combo2_Validate False
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Dim strNo As String, iPos As Integer
   If Combo2.Text = "" Then
      If pa(9) <> 台灣國家代號 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Cancel = True
         Exit Sub
      End If
      
   ElseIf Not ChgType(12) Then
      Cancel = True
      
   Else
      strNo = Combo2.Text
      
      'Add by Morgan 2008/2/22 加聯絡人判斷
      iPos = InStr(strNo, "-")
      If iPos > 0 Then
         strNo = Left(strNo, iPos - 1)
      End If
      'end 2008/2/22
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If PUB_CheckStatus(strNo) = False Then Cancel = True
   End If
End Sub

Private Sub Form_Activate()
    'Add By Cheng 2003/10/06
    '若有按下變更事項按鈕, 則重新讀取資料
    If m_blnClkChgEvnBtn = True Then
        ReadPatent
        'Add by Morgan 2004/7/22
        m_bolActive = False
        Label2(0) = strReceiveNo
        m_blnClkChgEvnBtn = False
'Removed by Morgan 2016/8/19
'    Else
'        ReadPatent
'        Label2(0) = strReceiveNo
'end 2016/8/19
    End If
    
   'Add by Morgan 2004/7/22
   '若沒有客戶減免身分需輸入則游標預設在繳費年度
   If m_bolActive = True Then Exit Sub
   m_bolActive = True
   If pa(9) = "000" Then
      
      'Add by Morgan 2005/7/4
      '檢查是否可減免退費且未收文
      If PUB_GetCaseDiscStat(pa(1) & pa(2) & pa(3) & pa(4)) = "Y" Then
         If PUB_CheckYearFeeReturn(pa) = True Then
            AddNewCP
            frm040104_1.Show
            frm040104_1.Command1_Click
            Unload Me
            Exit Sub
         End If
      End If
      '2005/7/4 end
   
      Dim i As Integer
      For i = 1 To 5
         If txtAD(i).Enabled = True And txtAD(i).Text = "" Then
            txtAD(i).SetFocus
            Exit Sub
         End If
      Next
   End If
   If Text5(0).Enabled = True Then Text5(0).SetFocus

End Sub

Private Sub Form_Initialize()
   'Add by Morgan 2005/7/15
   ReDim pa(TF_PA)
   ReDim cp(TF_CP)
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   
   'Add by Morgan 2010/12/15
   If m_bolBeCalled Then
      Text1 = m_CP01
      Text2 = m_CP02
      Text3 = m_CP03
      Text4 = m_CP04
      strReceiveNo = m_CP09
   Else
   'end 2010/12/15
      
      With frm040104_1
         Text1 = .Text1
         Text2 = .Text2
         Text3 = .Text3
         Text4 = .Text4
         strReceiveNo = .Tag
      End With
   End If
   
   ReadPatent
   Text8(2) = "N" 'Add by Amy 2014/09/03由ReadPatent搬過來
   
   'Add by Morgan 2005/7/15
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   Text8(1).Visible = False
   lstNameAgent.Clear
   'Add by Amy 2018/03/20 +發文規費欄
   lblCP84.Enabled = False
   textCP84.Enabled = False
   If pa(9) = "000" Then
      'Added by Morgan 2012/9/7 不要傳出名代理人,否則若重新發文會殘留前次發文資料,當新案有改出名人而本程序未改選將會造成不一致
      'PUB_SetOurAgent lstNameAgent, pa(), cp(110)
      'Modified by Morgan 2020/3/20 +cp10
      PUB_SetOurAgent lstNameAgent, pa(), , cp(10), True   'Modified by Morgan 2021/12/15 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
      'Add by Amy 2018/03/20 年費移作次年,顯示發文規費欄,預設修改指示信或申請書
      If cp(10) = "612" Then
         lblCP84.Enabled = True
        textCP84.Enabled = True
        Text8(0) = ""
        Text5(3) = "Y"
        Text5(3).Enabled = False
      End If
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
      'Add by Amy 2018/03/20 年費移作次年,預設修改指示信或申請書
      If cp(10) = "612" Then
        Text5(3) = "Y"
        Text5(3).Enabled = False
      End If
   End If
   '2005/7/14 END
   
   Label2(0) = strReceiveNo
   m_blnClkChgEvnBtn = False
   
   'Add by Morgan 2007/8/10 延展費
   '2008/11/11 modify by sonia
   'If cp(10) = "607" Then
   If pa(9) = "013" And pa(8) <> "1" Then
      Text5(1).Visible = False
      'Modify by Morgan 2009/2/18 香港都收年費605所以寫死
      'Label11 = "第                    次" & GetCaseTypeName("P", cp(10), 1)
      Label11 = "第                    次續期費"
   End If
   'end 2007/8/10
   
   'Added by Morgan 2012/11/19
   If cp(10) = "606" Then
      Label11 = "第                    至                  年維持費"
   End If
   'end 2012/11/19
   
   'Modify by Amy 2014/09/16 年費定稿上線 是否列印通知函開放可輸
   Text8(2) = ""
    Text8(2).Enabled = True
   'end 2014/09/16
   
   'Added by Morgan 2021/1/27 從 Formsave 移來以便共用
   'Mark by Lydia 2023/06/20 改在ReadPatent
   'stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
   'stCP12 = GetSalesArea(stCP13)
   'end 2021/1/27
   'end 2023/06/20
   
   'Modified by Morgan 2012/8/20 自 Activate 事件移來,因為批次發文不會觸發
   'Add by Morgan 2009/11/13
   'FMP案領證/年費發文不出定稿，收到收據後才要
   'Modified by Morgan 2021/1/27
   'If Left(cp(12), 1) = "F" And pa(9) <> "000" Then
   If Left(stCP12, 1) = "F" And pa(9) <> "000" Then
   'end 2021/1/27
      'm_bolFMP = True 'Mark by Lydia 2023/06/20 改在ReadPatent
      Text8(2) = "N"
      Text8(2).Enabled = False
   End If
   '2010/4/8 ADD BY SONIA 大陸案B類收文預設不印通知函
   If cp(9) > "B" And pa(9) = "020" Then
      Text8(2) = "N"
   End If
   '2010/4/8 END
   'end 2012/08/20
   
   'Add by Amy 2014/09/09
   'Modified by Morgan 2016/6/22 非臺灣案電子化
   'If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
   If Left(Pub_StrUserSt03, 1) <> "F" Then
   'end 2016/6/22
        '通函不可修改
        Text5(9).Enabled = False
   End If
   'end 2014/09/09
   
   'Added by Morgan 2017/1/11
   '專利處人員操作時年費通知人欄位鎖住以避免不小心改到(目前只有外專人員會設定)
   If Left(Pub_StrUserSt03, 1) = "P" Then
      Text5(6).Locked = True
   End If
   'end 2017/1/11
   
   Frame1.BackColor = &H8000000F 'Added by Lydia 2020/04/07
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Set frm040104_a = Nothing 'Removed by Morgan 2021/12/15 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, txt As TextBox, i As Integer
Dim strTmp(0 To 5) As String, varTmp As Variant, strTmp1(0 To 5) As String
'Add By Cheng 2003/04/24
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Add by Morgan 2004/7/22
Dim strAD10 As String, strCU15 As String
Dim m_Fee As String         '銷帳服務費 2012/8/1 add by sonia
Dim m_Official As String    '銷帳規費   2012/8/1 add by sonia
  
   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
         
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      For i = 26 To 30
         If pa(i) <> "" Then ChgType i
      Next
      
      'Text8(2).Text = "N" 'Modify by Amy 2014/09/03 搬至FormLoad 避免通知函設空又觸發Active導致通知函會再被設為N
      If pa(9) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(pA(9), strExc(0)) Then Label2(12) = strExc(0)
         If ClsPDGetNation(pa(9), strExc(0)) Then Label2(12) = strExc(0)
      End If
      
      Text15(0) = pa(5)
      Text15(1) = pa(6)
      Text15(2) = pa(7)

      Label2(9) = pa(72)
      If pa(76) <> "" Then Text5(6) = pa(76): ChgType 6

      strTmp1(0) = strReceiveNo
      
      For i = 1 To 4
         strTmp1(i) = pa(i)
      Next
      If GetMoneyDate(pa(8), pa(9), strTmp1, strCaseFee(1), strCaseFee(2)) = True Then
         'Modify by Morgan 2004/12/14 舊法新型專用期12年
         If pa(9) = "000" And pa(8) = "2" And Val(pa(14)) > 0 And Val(pa(14)) < 930701 Then
            strCaseFee(2) = "1,2,3,4,5,6,7,8,9,10,11,12"
         End If
      End If
   End If
   
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(13) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(13), strExc(0)) Then Label2(1) = strExc(0)
         If ClsPDGetStaff(cp(13), strExc(0)) Then Label2(1) = strExc(0)
      End If
      Label2(2) = cp(6)
      Label2(13) = cp(7)
      If cp(27) = "" Then
         Text5(4) = strSrvDate(2)
      Else
         Text5(4) = cp(27)
      End If
      
      Text5(4).Tag = Text5(4).Text 'Added by Lydia 2020/04/07
      
      'Added by Lydia 2023/06/20
      stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
      stCP12 = GetSalesArea(stCP13)
      If Left(stCP12, 1) = "F" And pa(9) <> "000" Then '判斷FMP案
         m_bolFMP = True
      Else
         m_bolFMP = False
      End If
      '寰華案
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
         'Modified by Morgan 2013/3/22
         'Text5(5) = cp(14): ChgType (5)
         If ClsPDGetStaff(cp(14), strExc(0)) Then Label2(10) = strExc(0)
         'end 2013/3/22
      End If
      
      Text5(8) = cp(64)
      Text8(1) = cp(22)
      'Modify by Morgan 2008/10/16 +若進度檔已有代理人則預設
      'Modified by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設 => cp(9), pa(9), pa(26)
      AddAgent Combo2, cp, , cp(44), cp(116), cp(9), pa(9), pa(26)
         
      'Added by Morgan 2012/11/15
      m_bolChgAgent = False
      If strSrvDate(1) >= "20121201" Then
         If Left(Combo2, 6) = "Y43343" Then
            Combo2 = "Y53374"
            Combo2_Validate False
            m_bolChgAgent = True
         End If
      End If
      'end 2012/11/15
   
      'Added by Morgan 2013/4/16
      'modify by sonia 2015/5/5 再加北京瑞思Y52862
      'modify by sonia 2015/11/27 雅娟通知取消北京瑞思Y52862
      'If Left(Combo2, 6) = "Y37580" Or Left(Combo2, 6) = "Y52862" Then
      'modify by sonia 2017/10/5 玲玲通知+Y52871北京市立方律師事務所
      'If Left(Combo2, 6) = "Y37580" Then
      'modify by sonia 2018/4/10 +Y52459
      'If Left(Combo2, 6) = "Y37580" Or Left(Combo2, 6) = "Y52871" Then
      'modify by sonia 2019/3/28 北京瑞思Y52862再加回
      'modify by sonia 2019/9/4 +Y50515
      'Modified by Morgan 2020/5/11 +Y51350--潘韻丞
      If Left(Combo2, 6) = "Y37580" Or Left(Combo2, 6) = "Y52871" Or Left(Combo2, 6) = "Y52459" Or Left(Combo2, 6) = "Y52862" Or Left(Combo2, 6) = "Y50515" Or Left(Combo2, 6) = "Y51350" Then
         Combo2 = "Y53374"
         Combo2_Validate False
         m_bolChgAgent = True
      End If
      'end 2013/4/16
      
      'Added by Morgan 2023/10/13
      'Modified by Morgan 2023/10/13 大陸年費一律都給寰華--郭
      'Modified by Morgan 2023/10/20 X0822505 開平威技電器有限公司、X0822506 開平威寶精密電機有限公司 除外--潘韻丞
      'Modified by Morgan 2024/3/26 +X41570080 亞浩電子,X41570060 訊強電子,X41570070 訊豪電子 除外--郭
      'Modified by Morgan 2024/5/28 +X00497100康舒電子(東莞)有限公司 除外--潘韻丞
      If pa(9) = "020" And Combo2 <> "Y53374" And InStr("X0822505,X0822506,X4157008,X4157006,X4157007,X0049710", Left(pa(26), 8)) = 0 Then
         If Combo2 <> "" Then m_bolChgAgent = True
         Combo2 = "Y53374"
         Combo2_Validate False
      End If
      'end 2023/10/13
      
      'Add by Morgan 2010/8/17 香港案
      If pa(9) = "013" And pa(8) <> "1" Then
         If cp(53) <> "" Then
            Text5(0) = TransHKData(cp(53))
            Text5(1) = Text5(0)
         End If
      Else
      'end 2010/8/17
      
         'Add by Morgan 2010/1/21 年費收文有輸入起迄時帶出
         Text5(0) = cp(53)
         Text5(1) = cp(54)
      End If
   End If
   
   '2012/8/1 add by sonia 若有銷帳則要扣除銷帳規費
   If Val(cp(77)) > 0 Then
      If GetCP77Detail(cp(9), m_Fee, m_Official) = True Then
         cp(17) = cp(17) - m_Official
      End If
   End If
   '2012/8/1 end
      
    'Add By Cheng 2003/04/24
    '取得下一程序的法定期限
    StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " And Np07='" & cp(10) & "' Order By NP09 Desc "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        m_strNP09 = "" & ChangeWStringToTString(rsA("NP09").Value)
    Else
        m_strNP09 = "" & cp(7)
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing

'Modify by Morgan 2008/10/8 因NP的年費期限可能會延期故需重算原期限
'    'Add By Cheng 2003/09/01
'    m_strNP09_1 = m_strNP09
   m_strNP09_1 = PUB_GetNextFeeDate(pa)
'END 2008/10/8

   m_strNP09_2 = m_strNP09_1 'Added by Morgan 2016/4/8
    
    '若法定期限為假日時, 抓大於法定期限最近的工作天
    If m_strNP09_1 <> "" Then
        m_strNP09_1 = DBDATE(PUB_GetLawDay(DBDATE(m_strNP09_1)))
    End If
   
   'Add by Morgan 2004/6/23
   '台灣可設定申請人年費減免身分
   If pa(9) = "000" Then
      lblCP81C.Visible = True
      lblCP81.Visible = True
      lblCP81.Caption = PUB_GetCP81(pa)
      
      If cp(10) = "612" Then textCP84.Tag = cp(17) 'Add by Amy 2018/03/20 移作次年需判斷收/發文年費
      
      'Add by Morgan 2004/7/21
      '減免身分
      For i = 1 To 5
         txtAD(i).Enabled = False
         txtAD(i).Tag = ""
         txtAD(i).Text = ""
         If pa(25 + i) <> "" Then
            txtAD(i).Text = PUB_GetAD03(pa(25 + i), pa(9), strAD10, strCU15)
            txtAD(i).Tag = txtAD(i).Text
            'Added by Morgan 2016/8/19 設定優先,因可能會有特例(目前美國有)--玲玲確認
            If txtAD(i).Text = "N" Then
               If strCU15 <> "0" And strCU15 <> "2" Then
                  txtAD(i).Enabled = True
               End If
            Else
            'end 2016/8/19
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
            End If 'Added by Morgan 2016/8/19
         End If
      Next

   Else
      lblCP81.Visible = False
      lblCP81C.Visible = False
   End If
   
   'Added by Lydia 2020/04/07 是否電子送件
   txtCP118 = ""
   Frame1.Visible = False
   If pa(9) = 台灣國家代號 And cp(10) = 年費 And m_bolBeCalled = False Then '限年費發文,並且排除整批發文
        If cp(118) <> "" Then txtCP118 = "Y"
        txtPayToday = "" '自動扣款
        Frame1.Visible = True
        If txtCP118 = "Y" Then
            If Val(ServerTime) <= 153000 Then '自動扣款(A)若超過3點半發文則須人工輸入是否當日扣款
               txtPayToday = "Y"
            End If
        End If
   End If
   'end 2020/04/07
   
    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25
End Sub

Private Function ChgType(iSitu As Integer) As Boolean
 Dim i As Integer, bolChk As Boolean, varTmp As Variant
 Dim strTempName As String
 Dim nPos As Integer
 Dim nCurrPos As Integer
 Dim aryCaseFee As Variant
 Dim aryCurrFee As Variant
 Dim bFind As Boolean
 Dim strTmp1(0 To 5) As String
 Dim strEffDate As String
 Dim iStart As String    '2008/11/7 ADD BY SONIA
 
   ChgType = False
   Select Case iSitu
      Case 0:
         If IsEmptyText(Text5(0)) = False Then
            aryCaseFee = Split(strCaseFee(2), ",")
            aryCurrFee = Split(Label2(9), ",")
            If nPos > UBound(aryCaseFee) Then
               MsgBox "無繳年費年度，請查明後再輸入 !", vbCritical
            '2008/11/11 modify by sonia
            'Add by Morgan 2007/8/10 區分年費與延展費
            'ElseIf cp(10) = "607" Then
            '   If Text5(0) <> UBound(aryCurrFee) + 2 Then
            ElseIf pa(9) = "013" And pa(8) <> "1" Then
               'Modify by Morgan 2009/2/18 香港短期和外觀改輸入第?次(P-75730)
               'If Label2(9) <> "" Then
               '   For i = 0 To UBound(aryCaseFee)
               '      If Format(aryCaseFee(i)) = aryCurrFee(UBound(aryCurrFee)) Then
               '         iStart = i
               '         Exit For
               '      End If
               '   Next
               'Else
               '   iStart = -1
               'End If
               'If iStart = UBound(aryCaseFee) Then
               '   MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
               'ElseIf Text5(0) <> aryCaseFee(iStart + 1) Then
               '2008/11/11 end
               '   MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
               iStart = 0
               '沒有繳費記錄時為第1次
               If Label2(9) = "" Then
                  iStart = 1
               Else
                  For i = 0 To UBound(aryCaseFee)
                     If Val(aryCaseFee(i)) > Val(aryCurrFee(UBound(aryCurrFee))) Then
                        '陣列從0開始,次數要加1
                        iStart = i + 1
                        Exit For
                     End If
                  Next
               End If
               If iStart = 0 Then
                  MsgBox "無待繳交續期費，請查明後再輸入 !", vbCritical
               ElseIf Val(Text5(0)) <> iStart Then
                  MsgBox "繳費次錯誤，請查明後再輸入 !", vbCritical
               'end 2009/2/18
               Else
                  Text5(1) = Text5(0)
                  ChgType = True
                  ' 計算下次繳年費日期
                  CompNextFeeDate
                  If pa(9) > "010" Then
                     If Text5(7) <> "" Then
                        Text5(7) = TAIWANDATE(DBDATE(DateAdd("d", -5, DateAdd("m", -1, ChangeWStringToWDateString(DBYEAR(Text5(7)) & DBMONTH(Text5(7)) & DBDAY(Text5(7)))))))
                     End If
                  Else
                     If Text5(7) <> "" Then
                        Text5(7) = TAIWANDATE(DBDATE(DateSerial(Val(DBYEAR(Text5(7))), Val(DBMONTH(Text5(7))), Val(DBDAY(Text5(7))) - 2)))
                     End If
                  End If
               End If
               
            Else
               ' 找尋已繳年度串列中空白的位置
               For nPos = 0 To UBound(aryCurrFee)
                  If IsEmptyText(aryCurrFee(nPos)) = True Then
                     Exit For
                  End If
               Next nPos
               If Text5(0) <> aryCaseFee(nPos) Then
                  MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
               Else
                  ChgType = True
               End If
            End If
            Erase aryCurrFee
            Erase aryCaseFee
         End If
      Case 1:
         'Add by Morgan 2007/8/10 延展費
         If Text5(1).Visible = False Then
            ChgType = True
         'end 2007/8/10
         ElseIf IsEmptyText(Text5(1)) = False Then
            Set aryCaseFee = Nothing
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
                  'Add by Morgan 2006/4/25
                  '判斷繳費年度迄-1是否有超過專用期(多繳)
                  If nCurrPos > 0 Then
                     strEffDate = TransDate(CompDate(0, aryCaseFee(nCurrPos - 1), strCaseFee(1)), 1)
                     'Add by Morgan 2006/7/25 若無專用期時提醒
                     If pa(25) = "" Then
                        If m_bolConfirm = False Then
                           If cp(10) <> "606" Then 'Added by Morgan 2012/11/19 維持費會在發證前還不會有專用期
                              If MsgBox("本案尚無專用期，是否繼續發文?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                                 Exit Function
                              End If
                           End If
                           
                           strTmp1(0) = strReceiveNo
                           For i = 1 To 4
                              strTmp1(i) = pa(i)
                           Next
                           If GetMoneyDate(Val(pa(8)) + 10, pa(9), strTmp1, "", "", m_EndDate) = False Then
                              MsgBox "無法計算出專用期止日!", vbCritical
                              Exit Function
                           Else
                              If Val(DBDATE(strEffDate)) > Val(m_EndDate) Then
                                 MsgBox "繳費年度迄超出應繳範圍，請查明後再輸入 !", vbCritical
                                 Exit Function
                              Else
                                 m_bolConfirm = True
                              End If
                           End If
                        End If
                     Else
                     'end 2006/7/25
                        If Val(strEffDate) > Val(pa(25)) Then
                           MsgBox "繳費年度迄超出應繳範圍，請查明後再輸入 !", vbCritical
                           Exit Function
                        End If
                     End If
                  End If
                  '2006/4/25 END
                  ChgType = True
               End If
            End If
            
            ' 計算下次繳年費日期
            CompNextFeeDate
            If pa(9) > "010" Then
               If Text5(7) <> "" Then
                  Text5(7) = TAIWANDATE(DBDATE(DateAdd("d", -5, DateAdd("m", -1, ChangeWStringToWDateString(DBYEAR(Text5(7)) & DBMONTH(Text5(7)) & DBDAY(Text5(7)))))))
               End If
            Else
               If Text5(7) <> "" Then
                  Text5(7) = TAIWANDATE(DBDATE(DateSerial(Val(DBYEAR(Text5(7))), Val(DBMONTH(Text5(7))), Val(DBDAY(Text5(7))) - 2)))
               End If
            End If
         End If

      Case 2
         '若法定期限為假日則用大於法定期限最近的工作日與發文日比較
         'Modify by Morgan 2008/10/8 需用原年費期限判斷
         'If DBDATE(Text5(4)) > IIf(DBDATE(m_strNP09) >= DBDATE(m_strNP09_1), DBDATE(m_strNP09), DBDATE(m_strNP09_1)) And pa(9) = 台灣國家代號 Then
         If DBDATE(Text5(4)) > m_strNP09_1 And pa(9) = 台灣國家代號 Then
            If Text5(iSitu) <> "Y" Then
               If Not m_bolBeCalled Then 'Add by Morgan 2010/12/15
                  MsgBox "發文日大於法定期限則[費用是否要雙倍]欄必須為 Y !", vbCritical
               End If
               Text5(iSitu) = "Y"
            End If
            ChgType = True
         Else
            If Text5(iSitu) = "Y" Then
               MsgBox "[費用是否要雙倍]欄錯誤 !", vbCritical
            Else
               ChgType = True
            End If
         End If
      Case 4
         If Text5(iSitu) <> "" Then
            'Modify by Morgan 暫改可大於系統日但不可大於 930709
            'If Not ChkDate(Text5(4)) Or Val(Text5(4)) > Val(strSrvDate(2)) Then
            '2011/12/8 MODIFY BY SONIA 發文日可輸系統日的下一個工作日,並取消930709控制
            'If Not ChkDate(Text5(4)) Or (Val(Text5(4)) > Val(strSrvDate(2)) And Val(Text5(4)) > 930709) Then
            '   MsgBox "發文日期不正確或發文日大於系統日，請重新輸入 !", vbCritical
            If Not ChkDate(Text5(4)) Or DBDATE(Val(Text5(4))) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
               MsgBox "發文日期不正確或發文日大於系統日下一個工作日，請重新輸入 !", vbCritical
            '2011/12/8 END
            Else
               ChgType = True
            End If
         Else
            MsgBox "發文日不可空白 !", vbCritical
         End If
      Case 26, 27, 28, 29, 30
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(pa(iSitu), strTempName) Then
         If ClsLawLawGetName(pa(iSitu), strTempName) Then
            Label2(iSitu - 23) = strTempName
            ChgType = True
         End If
         
'Removed by Morgan 2013/3/22 已經沒有再用
'      Case 5
'         If Text5(5) <> "" Then
'            'edit by nickc 2007/02/02 不用 dll 了
'            'If objPublicData.GetStaff(Text5(5), strTempName) Then
'            If ClsPDGetStaff(Text5(5), strTempName) Then
'               Label2(10) = strTempName
'               ChgType = True
'            End If
'         Else
'            ChgType = True
'         End If

      Case 6
         If Text5(6) <> "" Then
            'edit by nickc 2007/02/05 不用 dll 了
            'If objLawDll.LawGetName(Text5(6), strTempName) Then
            If ClsLawLawGetName(Text5(6), strTempName) Then
               Label2(11) = strTempName
               ChgType = True
            End If
         Else
            ChgType = True
         End If
      Case 7
            '若下次繳費日不超過專用期止日時才要檢查
            If blnOverDate = False Then
                If IsEmptyText(Text5(7)) = False Then
                   If CheckIsTaiwanDate(Text5(7), False) = False Then
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
                ChgType = True
            End If
      Case 12 '代理人
         strExc(1) = Combo2.Text
         'Add by Morgan 2008/2/22 加判斷是否為聯絡人
         If InStr(strExc(1), "-") > 0 Then
            If ClsPDGetContact(strExc(1), strTempName) Then
               Combo2 = strExc(1)
               Label2(14) = strTempName
               ChgType = True
            End If
         
         '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
         ElseIf PUB_GetAgentName(pa(1), strExc(1), strTempName) = True Then
            Combo2.Text = strExc(1)
            Label2(14).Caption = strTempName
            ChgType = True
         Else
            Label2(14).Caption = ""
         End If
      Case Else
         ChgType = True
   End Select
End Function

Private Sub Text15_GotFocus(Index As Integer)
    Select Case Index
    Case 0
        Me.Text15(Index).SelStart = 0
        Me.Text15(Index).SelLength = 0
    Case Else
        TextInverse Text15(Index)
    End Select
End Sub

Private Sub Text5_GotFocus(Index As Integer)
  TextInverse Text5(Index)
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      'Add by Morgan 2009/1/10 控制不可輸入非數字
      Case 0, 1
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
         
      Case 2, 3, 9
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
         
      Case 5, 6
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
   Cancel = Not ChgType(Index)
   'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
   'Moddified by Lydia 2020/04/07
   'If Cancel = False And Index = 6 Then
   '   If PUB_CheckStatus(Text5(Index).Text) = False Then Cancel = True
   'End If
   Select Case Index
        Case 4
            If Text5(Index).Tag <> Text5(Index).Text Then
                '當發文日有改時,電子送件案要人工輸入是否當日扣款
                If pa(9) = "000" And Frame1.Visible = True And txtCP118 = "Y" Then
                   txtPayToday.Text = ""
                End If
                Text5(Index).Tag = Text5(Index).Text
            End If
        Case 6
            '檢查客戶/代理人是否不再使用
            If Cancel = False Then
                If PUB_CheckStatus(Text5(Index).Text) = False Then Cancel = True
            End If
   End Select
   'end 2020/04/07
   
   If Cancel = True Then TextInverse Text5(Index)
End Sub

Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

' 90.07.10 modify by louis (計算下次繳年費日期)
Private Sub CompNextFeeDate()
   Dim strPA09 As String
   Dim strPA08 As String
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strDate As String, strTmp1(0 To 5) As String, i As Integer
   Dim aryCaseFee As Variant
   Dim iStart As String    '2008/11/7 ADD BY SONIA
   
   'Add By Cheng 2003/01/14
   '預設下次繳年費日期不超過專用期止日
   blnOverDate = False
   strPA08 = pa(8)
   strPA09 = pa(9)
   Select Case strPA08
      Case "1":
         strSql = "SELECT NA06 FROM NATION " & _
            "WHERE NA01 = '" & strPA09 & "' "
      Case "2":
         strSql = "SELECT NA08 FROM NATION " & _
            "WHERE NA01 = '" & strPA09 & "' "
      Case "3":
         strSql = "SELECT NA10 FROM NATION " & _
            "WHERE NA01 = '" & strPA09 & "' "
   End Select
   
   If IsEmptyText(strSql) = False Then
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then
            Select Case rsTmp.Fields(0)
               Case 1: strDate = DBDATE(cp(5))
               Case 2: strDate = DBDATE(pa(10))
               Case 3: strDate = DBDATE(Text5(4))
               Case 4: strDate = DBDATE(cp(25))
               Case 5: strDate = DBDATE(pa(14))
               Case 6: strDate = DBDATE(pa(21))  '2009/6/24 modify by sonia 原為pa(27)
               Case 7: strDate = DBDATE(pa(12))
            End Select
            
            ' 依日期再+繳年費迄
            If IsEmptyText(strDate) = False Then
               'Add by Morgan 2007/8/10 延展費
               '2008/11/6 modify by sonia 以P-074053,085139測試
               'If cp(10) = "607" Then
               '   aryCaseFee = Split(strCaseFee(2), ",")
               '   If UBound(aryCaseFee) + 1 > Val(Text5(0)) Then
               '      strDate = CompDate(0, aryCaseFee(Val(Text5(0))), strDate)
               If strPA09 = "013" And strPA08 <> "1" Then
                  aryCaseFee = Split(strCaseFee(2), ",")
                  'Modify by Morgan 2009/2/18 香港短期和外觀改輸入第?次(P-75730)
                  'If aryCaseFee(UBound(aryCaseFee)) > Val(Text5(1)) Then
                  '   For i = 0 To UBound(aryCaseFee)
                  '      If Format(aryCaseFee(i)) = Val(Text5(1)) Then
                  '         iStart = i
                  '         Exit For
                  '      End If
                  '   Next
                  '   strDate = CompDate(0, aryCaseFee(iStart + 1), strDate)
                  iStart = Val(Text5(1))
                  If UBound(aryCaseFee) + 1 > iStart Then
                     strDate = CompDate(0, aryCaseFee(iStart), strDate)
                  'end 2009/2/18
               '2008/11/6 end
                  Else
                     Text5(7) = ""
                     blnOverDate = True
                     rsTmp.Close
                     Set rsTmp = Nothing
                     Exit Sub
                  End If
               Else
               'end 2007/8/10
                  strDate = DBDATE(DateAdd("yyyy", Val(Text5(1)), ChangeWStringToWDateString(DBYEAR(strDate) & DBMONTH(strDate) & DBDAY(strDate))))
               End If
               ' 依申請國家減天數
               If pa(9) < "010" Then
                    '只要減一天
                  strDate = DBDATE(DateSerial(Val(DBYEAR(strDate)), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) - 1))
               'Modify by Morgan 2007/9/6 非台灣案都不必減一天--郭,玲玲
               'Else
               '   'Modify by Morgan 2006/3/10 大陸不必減一天
               '   'strDate = DBDATE(DateSerial(Val(DBYEAR(strDate)), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) - 1))
               '   If pa(9) <> "020" Then
               '      strDate = DBDATE(DateSerial(Val(DBYEAR(strDate)), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) - 1))
               '   End If
               '   '2006/3/10 end
               'end 2007/9/6
               End If
                '若有專用期止日
                If pa(25) <> "" Then
                    '若專用期止日小於等於下次繳費日
                    If Val(DBDATE(pa(25))) <= Val(strDate) Then
                        Text5(7) = ""
                        '下次繳年費日期超過專用期止日
                        blnOverDate = True
                    '若專用期止日大於下次繳費日
                    Else
                        Text5(7) = TAIWANDATE(strDate)
                        '下次繳年費日期不超過專用期止日
                        blnOverDate = False
                        
                        'Added by Morgan 2024/10/8
                        '大陸發明案增加檢查是否有補償期,若有則需以原專用期止日判斷
                        If pa(9) = "020" And pa(8) = "1" Then
                           intI = 0: strExc(1) = ""
                           If PUB_GetCNExtDays(pa(), pa(25), intI, , strExc(1)) = True Then
                              If intI > 0 Then
                                 If Val(strExc(1)) <= Val(strDate) Then
                                    Text5(7) = ""
                                    '下次繳年費日期超過專用期止日
                                    blnOverDate = True
                                 End If
                              End If
                           End If
                        End If
                        'end 2024/10/8
                        
                    End If
                '若無專用期止日
                Else
                    strTmp1(0) = strReceiveNo
                    For i = 1 To 4
                       strTmp1(i) = pa(i)
                    Next
                    If GetMoneyDate(Val(pa(8)) + 10, pa(9), strTmp1, strCaseFee(1), strCaseFee(2), m_EndDate) Then   '抓專用期起止日
                        If GetMoneyDate(pa(8), pa(9), strTmp1, strCaseFee(1), strCaseFee(2)) = True Then
                           'Modify by Morgan 2004/12/14 舊法新型專用期12年
                           If pa(9) = "000" And pa(8) = "2" And Val(pa(14)) > 0 And Val(pa(14)) < 930701 Then
                              strCaseFee(2) = "1,2,3,4,5,6,7,8,9,10,11,12"
                           End If
                        End If
                    End If
                    '若專用期止日小於等於下次繳費日
                    If Val(DBDATE(m_EndDate)) <= Val(strDate) Then
                        Text5(7) = ""
                        '下次繳年費日期超過專用期止日
                        blnOverDate = True
                    '若專用期止日大於下次繳費日
                    Else
                        Text5(7) = TAIWANDATE(strDate)
                        '下次繳年費日期不超過專用期止日
                        blnOverDate = False
                    End If
                End If
            End If
         End If
      End If
     
      rsTmp.Close
      Set rsTmp = Nothing
   End If
End Sub

'Add By Cheng 2002/03/08
Private Function CheckDataIntegrity() As Boolean
Dim Cancel As Boolean
   'add by nickc 2008/05/01
   If IsDebt(pa(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        GoTo IntegrityOrNot
   End If
Cancel = False

'檢查代理人欄位
Combo2_Validate Cancel
If Cancel = True Then
   Me.Combo2.SetFocus
   GoTo IntegrityOrNot
End If

CheckDataIntegrity = True
Exit Function

IntegrityOrNot:
CheckDataIntegrity = False
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
   Dim objTxt As Object
   Dim ii As Integer, i As Integer
   Dim Cancel As Boolean
   'Add by Morgan 2004/7/20
   Dim stAppNo As String   '未設定減免身分客戶代碼

   TxtValidate = False
   
   'Added by Morgan 2021/12/15 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/15
   
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
   
   'Add by Amy 2018/03/20
    If textCP84.Enabled = True Then
       If textCP84 = MsgText(601) Then
         MsgBox "請輸入發文規費！"
         Me.textCP84.SetFocus
         textCP84_GotFocus
         Exit Function
       End If
       Cancel = False
       textCP84_Validate Cancel
       If Cancel = True Then
          Me.textCP84.SetFocus
          textCP84_GotFocus
          Exit Function
       End If
    End If
  
      'Add by Morgan 2004/7/22
      If pa(9) = "000" Then
         m_DiscType = ""
         For i = 1 To 5
            m_DiscType = m_DiscType & txtAD(i).Text
            If txtAD(i).Enabled = True Then
               If txtAD(i).Text = "" Then
                  MsgBox "申請人減免身分不可空白", vbInformation
                  txtAD(i).SetFocus
                  txtAD_GotFocus i
                  Exit Function
               '公司可減免
               'Modify by Morgan 2004/7/29
               '學校不需證明
               'ElseIf (txtAD(i).Text = "2" Or txtAD(i).Text = "3") Then
               '學校
               ElseIf (txtAD(i).Text = "2") Then
                  '變更
                  If (txtAD(i).Tag <> "2" And txtAD(i).Tag <> "") Then
                     If MsgBox("確定要變更申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】減免身分為【學校】？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                        txtAD(i).SetFocus
                        txtAD_GotFocus i
                        Exit Function
                     End If
                  End If
               '公司
               ElseIf (txtAD(i).Text = "3") Then
                  '新增或變更
                  If (txtAD(i).Tag <> "3") Then
                     If MsgBox("申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】的減免身分將設定為【中小企業】，確定有【證明文件】存放於本卷？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                        txtAD(i).SetFocus
                        txtAD_GotFocus i
                        Exit Function
                     End If
                  End If
               '不可減免
               ElseIf (txtAD(i).Text = "N") Then
                  '身分變更
                  If (txtAD(i).Tag <> "N" And txtAD(i).Tag <> "") Then
                     If MsgBox("確定要變更申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】減免身分為【不可減免】？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                        txtAD(i).SetFocus
                        txtAD_GotFocus i
                        Exit Function
                     End If
                  End If
               End If
            End If
         Next
         If InStr(m_DiscType, "N") > 0 Then
            cp(81) = "N"
         Else
            cp(81) = "Y"
         End If
      End If
      
   'Add by Morgan 2004/9/14
   If Combo2.Enabled = True Then
      Cancel = False
      Combo2_Validate Cancel
      If Cancel = True Then
         Combo2.SetFocus
         Exit Function
      End If
   End If
   '2004/9/14 end

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
   
   'Added by Lydia 2020/04/07
   If Frame1.Visible = True And Me.txtCP118 = "Y" Then
      If Me.Text8(1) = "N" Then
          MsgBox "電子送件不可不出名！", vbCritical
          Exit Function
      End If
      If Me.txtPayToday = "" Then
          MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
          Me.txtPayToday.SetFocus
          Exit Function
      End If
   End If
   'end 2020/04/07
    
   'Added by Morgan 2012/11/15
   If strSrvDate(1) >= "20121201" Then
      If Left(Combo2, 6) = "Y43343" Then
         MsgBox "Y43343 北京中原華和不再代繳本所年費，請更換代理人！", vbExclamation
         Exit Function
      End If
   End If
   'end 2012/11/15
   
   'Added by Morgan 2013/4/16
   If Left(Combo2, 6) = "Y37580" Then
      MsgBox "Y37580 天津三元不再代繳本所年費，請更換代理人！", vbExclamation
      Exit Function
   End If
   'end 2013/4/16
   'add by sonia 2018/4/10
   If Left(Combo2, 6) = "Y52871" Then
      MsgBox "Y52871 北京市立方不再代繳本所年費，請更換代理人！", vbExclamation
      Exit Function
   End If
   If Left(Combo2, 6) = "Y52459" Then
      MsgBox "Y52459 萬慧達知識產權不再代繳本所年費，請更換代理人！", vbExclamation
      Exit Function
   End If
   'end 2018/4/10
'modify by sonia 2015/11/27 雅娟通知取消北京瑞思Y52862
   'add by sonia 2015/5/5 加北京瑞思Y52862
   'add by sonia 2019/3/28 再加回北京瑞思Y52862
   If Left(Combo2, 6) = "Y52862" Then
      MsgBox "Y52862 北京瑞思不再代繳本所年費，請更換代理人！", vbExclamation
      Exit Function
   End If
   'end 2015/5/5
'end 2015/11/27
   
   'Added by Morgan 2013/3/22
   m_UseNewForm = False
   Erase m_AD1516
   If DBDATE(cp(7)) >= "20130801" Then
      If pa(9) = "000" And cp(81) = "Y" And Val(Text5(0)) < 7 Then
         For ii = 1 To 5
            If txtAD(ii) = "3" Then
               strExc(0) = "select ad15,ad16 from applicantdiscount where ad01='" & Left(pa(25 + ii) & "00", 8) & "' and ad02='000' and ad15 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  m_AD1516(ii, 1) = "" & RsTemp("ad15")
                  m_AD1516(ii, 2) = "" & RsTemp("ad16")
               Else
                  If MsgBox("申請人【" & pa(25 + ii) & " " & Label2(2 + ii) & " 】之減免身分為中小企業但尚未勾選減免資格，是否確定要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                     Exit For
                  Else
                     Exit Function
                  End If
               End If
            End If
         Next
         If ii > 5 Then
            m_UseNewForm = True
         End If
      End If
   End If
   'end 2013/3/22

   'Added by Morgan 2013/6/25
   m_str414CP09 = ""
   If pa(9) = "000" Then
      If PUB_ChkCPExist(pa, "414", 1, m_str414CP09, strExc(1)) = True Then
         intI = MsgBox("系統將同時發文回復原狀？", vbOKCancel + vbDefaultButton2 + vbInformation)
         If intI = vbOK Then
            If strExc(1) = "" Then
               MsgBox "回復原狀尚未分案！", vbExclamation
               Exit Function
            End If
         Else
            Exit Function
         End If
      End If
   End If
   'end 2013/6/25
   
   'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
   If Pub_ChkACS112isNull(pa(1), pa(2), pa(3), pa(4), txtCP113) = True Then
         txtCP113.SetFocus
         txtCP113_GotFocus
         Exit Function
   End If
   'end 2021/05/25
   
   TxtValidate = True
End Function

'Add By Cheng 2002/12/31
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

Dim stAddMsg As String 'Add by Morgan 2010/1/21

PUB_GetPatentYearFee strYF01, strYF02, strYF03, strYF04, strYF05From, strYF05To, blnDouble, cp(81), pa(14), Text5(4), m_strOfficalFee, m_strServiceFee, m_lngFeeDiscount, m_bol2ndY

'Removed by Morgan 2012/10/12 改呼叫共用函數
'Dim rsA As New ADODB.Recordset
'Dim StrSQLa As String
'Dim ii As Integer
'Dim iYear As Integer 'Add by Morgan 2004/6/21
'Dim lngOfficalFee As Long
'
'   ChkPatentYearFee = False
'   m_strOfficalFee = 0
'   m_strServiceFee = 0
'   m_lngFeeDiscount = 0
'   m_strPoints = 0
'   ii = 1
'   '取得案件性質為年費的相關費用
'   StrSQLa = "Select * From PatentYearFee Where YF01='" & strYF01 & "' AND YF02='" & strYF02 & "' AND YF03='" & strYF03 & "' AND YF04='" & strYF04 & "' AND YF05>=" & Val(strYF05From) & " AND YF05<=" & Val(strYF05To) & " Order By YF05 "
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'   While Not rsA.EOF
'      lngOfficalFee = Val(rsA.Fields("YF07").Value)
'      'Add by Morgan 2004/6/23
'      '年費減免
'      If cp(81) = "Y" Then
'         iYear = Val(rsA.Fields("YF05").Value)
'         If iYear >= 1 And iYear <= 3 Then
'            m_lngFeeDiscount = m_lngFeeDiscount + 800
'            lngOfficalFee = lngOfficalFee - 800
'         ElseIf iYear >= 4 And iYear <= 6 Then
'            m_lngFeeDiscount = m_lngFeeDiscount + 1200
'            lngOfficalFee = lngOfficalFee - 1200
'         End If
'      End If
'
'      '起始那年年費是否雙倍
'      If blnDouble = True And ii = 1 Then lngOfficalFee = lngOfficalFee * 2
'
'      m_strOfficalFee = Val(m_strOfficalFee) + lngOfficalFee
'      m_strServiceFee = Val(m_strServiceFee) + Val(rsA.Fields("YF06").Value)
'      rsA.MoveNext
'      ii = ii + 1
'   Wend
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
   
   
   m_strPoints = Val(m_strServiceFee) / 1000
    '若不等
    
   'Add by Morgan 2010/1/21
   '台灣年費或退費發文檢查是否有預繳年費可退
   If PUB_ChkRefund(pa, m_lngRefund, m_strFromYear, m_strToYear) Then
      If m_strOfficalFee - m_lngRefund > 0 Then
         m_strOfficalFee = m_strOfficalFee - m_lngRefund
         stAddMsg = ",已扣除可退之預繳年費 " & Format(m_lngRefund, DDollar)
      Else
         MsgBox "本案有可退預繳年費共 " & Format(m_lngRefund, DDollar) & " 元" & _
            "，現欲繳年費計 " & Format(m_strOfficalFee, DDollar) & " 元，故本次無需繳交規費請智權人員改收文【退費】！"
         ChkPatentYearFee = False
         Exit Function
      End If
   End If
   'end 2010/1/21
   
   If "" & cp(17) <> m_strOfficalFee Then
      If MsgBox("計算出的規費( " & Format(m_strOfficalFee, "#,##0") & stAddMsg & " )與目前進度檔的規費( " & Format(cp(17), "#,##0") & " )不同，是否要繼續作業???", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
          ChkPatentYearFee = True
      Else
          ChkPatentYearFee = False
      End If
   '若相等
   Else
      ChkPatentYearFee = True
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
   stLetter(3) = ChangeTStringToTDateString(Text5(4).Text)  '申請日期
   stLetter(4) = pa(8)  '專利種類
   stLetter(5) = pa(22) '證書號
   stLetter(6) = m_DiscType  '減免身分
   '減免年次
   If Val(Text5(1)) <= 6 Then
      stLetter(7) = Text5(0) & " 年至第 " & Text5(1)
   Else
      stLetter(7) = Text5(0) & " 年至第 6"
   End If
   stLetter(8) = Format(m_lngFeeDiscount, "###,###")  '減免金額
   
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
   'Modify by Morgan 2005/9/29 要抓本所期限
   'stLetter(13) = ChangeTStringToTDateString(Text5(7).Text) '下次期限
   stLetter(13) = ChangeTStringToTDateString(TransDate(PUB_GetWorkDay1(m_NEXTNP08, True), 1)) '下次期限
   
   iCopys = PUB_GetCopys(pa, stSales)
   stLetter(14) = PUB_GetStaffST15(stSales, "2")   '業務區
   stLetter(15) = GetStaffName(stSales)   '智權人員
   stLetter(16) = Format(iCopys)   '份數
   stLetter(17) = pa(47) '分所號
   stLetter(18) = GetAgentName
   
   '是否出名
   If Text8(1) <> "N" Then
      bolName = True
   Else
      bolName = False
   End If
   stLetter(19) = strReceiveNo 'Add by Morgan 2006/3/9
   stLetter(20) = Text5(2) 'Add by Morgan 2009/2/26
   PUB_PrintDiscForm stLetter, stData, bolName, bolEdit, m_bolNeedReasign
   
End Sub

'Add by Amy 2018/03/20
Private Sub textCP84_GotFocus()
    TextInverse textCP84
End Sub

Private Sub textCP84_KeyPress(KeyAscii As Integer)
    '只能輸倒退及數字鍵
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub textCP84_Validate(Cancel As Boolean)
    If textCP84.Enabled = False Then Exit Sub
    If Trim(textCP84) = MsgText(601) Then Exit Sub
    
    If Val(textCP84) <> Val(cp(17)) And Val(textCP84.Text) <> Val(textCP84.Tag) Then
        If MsgBox("發文規費【" & textCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            textCP84.Tag = textCP84.Text
        Else
            textCP84_GotFocus
            Cancel = True
        End If
    End If
End Sub
'end 2018/03/20

'Add by Morgan 2004/7/22
Private Sub txtAD_GotFocus(Index As Integer)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtAD(Index).IMEMode = 2
   CloseIme
   TextInverse txtAD(Index)
End Sub
'Add by Morgan 2004/7/22
'只有公司可輸入 2,3,N
Private Sub txtAD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2014/7/15 學校改預設且不可改
   'If Not (KeyAscii = 8 Or KeyAscii = 50 Or KeyAscii = 51 Or KeyAscii = 78) Then
   If Not (KeyAscii = 8 Or KeyAscii = 51 Or KeyAscii = 78) Then
      KeyAscii = 0
   End If
End Sub
'Add by Morgan 2005/5/18 新增減免退費內部收文
Private Sub AddNewCP()
   Dim Ncp() As String
   ReDim Ncp(TF_CP) As String
   Ncp(1) = cp(1)
   Ncp(2) = cp(2)
   Ncp(3) = cp(3)
   Ncp(4) = cp(4)
   Ncp(5) = strSrvDate(2)
   'Modify by Morgan 2011/2/24 修正百年收文號問題
   'Ncp(9) = "B" & Left(strSrvDate(2), 2)
   Ncp(9) = "B" & CompAutoNumberYear(GetTaiwanThisYear)
   
   Ncp(10) = "919" '減免退費
   Ncp(11) = "90"
   Ncp(12) = cp(12)
   Ncp(13) = cp(13)
   Ncp(14) = cp(14)
   Ncp(16) = 0
   Ncp(17) = 0
   Ncp(18) = 0
   Ncp(20) = "N"
   Ncp(26) = "N"
   Ncp(32) = "N"
   Ncp(86) = "Y" 'Add by Morgan 2005/7/4
   'edit by nickc 2007/02/02 不用 dll 了
   'If Not objPublicData.SaveNewCaseProgressDatabase("B", Ncp, intWhere) Then
   If Not ClsPDSaveNewCaseProgressDatabase("B", Ncp, intWhere) Then
      MsgBox "無法產生 ""減免退費"" 內部收文！", vbCritical
   End If
End Sub

'Add by Morgan 2005/7/15
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         bolCheck = True
      End If
   Next
   
   If bolCheck = True Then
      Text8(1) = ""
   Else
      Text8(1) = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

'Add by Morgan 2010/8/17
'年度次數互轉 bolReverse:False ->次數,True ->年度
Private Function TransHKData(strInput As String, Optional bolReverse As Boolean) As String
   Dim varTmp
   Dim ii As Integer
   varTmp = Split(strCaseFee(2), ",")
   If bolReverse Then
      TransHKData = varTmp(strInput - 1)
   Else
      For ii = LBound(varTmp) To UBound(varTmp)
         If strInput = varTmp(ii) Then
            TransHKData = ii + 1
            Exit For
         End If
      Next
   End If
End Function

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

'Added by Lydia 2020/04/07
Private Sub txtCP118_GotFocus()
   TextInverse txtCP118
End Sub

Private Sub txtCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP118_Change()
    'Added by Lydia 2020/04/08 電子送件要在發文前，先產生申請書；所以發文不用印
    If txtCP118 = "Y" Then
        Text8(0) = "N"
    End If
End Sub

Private Sub txtPayToday_GotFocus()
   TextInverse txtPayToday
End Sub

Private Sub txtPayToday_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
