VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040104_f 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專發文-技術報告/第三人申請技術報告"
   ClientHeight    =   5784
   ClientLeft      =   672
   ClientTop       =   996
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5784
   ScaleWidth      =   8580
   Begin TabDlg.SSTab SSTab1 
      Height          =   3555
      Left            =   45
      TabIndex        =   63
      Top             =   2160
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   6265
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "一般"
      TabPicture(0)   =   "frm040104_f.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCaseFees"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblItemCount"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCaseFee"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label32"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblNameAgent"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(13)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(14)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(15)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(18)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(19)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(20)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblCP84"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblPayToday"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label20(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblCP113(18)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text5"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lstNameAgent"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text4(2)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text4(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text4(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtItemCount"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtChkRltDate"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtLetter(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtLetter(4)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtCP27"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtLetter(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtCP22"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtLetter(0)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtLetter(2)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtCP84"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtPayToday"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtCP118"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtCP113"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).ControlCount=   37
      TabCaption(1)   =   "相關事項"
      TabPicture(1)   =   "frm040104_f.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check1(3)"
      Tab(1).Control(1)=   "Check1(2)"
      Tab(1).Control(2)=   "Check1(1)"
      Tab(1).Control(3)=   "Check1(0)"
      Tab(1).Control(4)=   "Check1(4)"
      Tab(1).Control(5)=   "Check1(5)"
      Tab(1).Control(6)=   "Check1(6)"
      Tab(1).Control(7)=   "Check1(7)"
      Tab(1).Control(8)=   "Check1(8)"
      Tab(1).Control(9)=   "Check1(9)"
      Tab(1).ControlCount=   10
      Begin VB.TextBox txtCP113 
         Height          =   270
         Left            =   4860
         MaxLength       =   4
         TabIndex        =   28
         Top             =   2607
         Width           =   540
      End
      Begin VB.TextBox txtCP118 
         Height          =   270
         Left            =   4740
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1470
         Width           =   255
      End
      Begin VB.TextBox txtPayToday 
         Height          =   270
         Left            =   2025
         MaxLength       =   1
         TabIndex        =   27
         Top             =   2610
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "非專利權人僅委任代理人辦理申請新型專利技術報告事宜"
         Height          =   210
         Index           =   9
         Left            =   -74775
         TabIndex        =   42
         Top             =   3120
         Width           =   6180
      End
      Begin VB.CheckBox Check1 
         Caption         =   "委任代理人辦理本案有關專利之事項"
         Height          =   210
         Index           =   8
         Left            =   -74280
         TabIndex        =   41
         Top             =   2820
         Width           =   3435
      End
      Begin VB.CheckBox Check1 
         Caption         =   "僅委任代理人辦理申請新型專利技術報告事宜"
         Height          =   210
         Index           =   7
         Left            =   -74055
         TabIndex        =   40
         Top             =   2520
         Width           =   4650
      End
      Begin VB.CheckBox Check1 
         Caption         =   "委任代理人辦理事項聲明"
         Height          =   210
         Index           =   6
         Left            =   -74775
         TabIndex        =   39
         Top             =   2250
         Width           =   2580
      End
      Begin VB.TextBox txtCP84 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4185
         TabIndex        =   15
         Top             =   1185
         Width           =   1092
      End
      Begin VB.TextBox txtLetter 
         Height          =   270
         Index           =   2
         Left            =   2070
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1425
         Width           =   255
      End
      Begin VB.TextBox txtLetter 
         Height          =   270
         Index           =   0
         Left            =   1530
         MaxLength       =   1
         TabIndex        =   21
         Top             =   2025
         Width           =   255
      End
      Begin VB.TextBox txtCP22 
         Height          =   270
         Left            =   5655
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1185
         Width           =   255
      End
      Begin VB.TextBox txtLetter 
         Height          =   270
         Index           =   1
         Left            =   4860
         MaxLength       =   1
         TabIndex        =   22
         Top             =   2025
         Width           =   255
      End
      Begin VB.TextBox txtCP27 
         Height          =   270
         Left            =   870
         MaxLength       =   7
         TabIndex        =   13
         Top             =   1185
         Width           =   1095
      End
      Begin VB.TextBox txtLetter 
         Height          =   270
         Index           =   4
         Left            =   4860
         MaxLength       =   1
         TabIndex        =   26
         Top             =   2295
         Width           =   255
      End
      Begin VB.TextBox txtLetter 
         Height          =   270
         Index           =   3
         Left            =   1530
         MaxLength       =   1
         TabIndex        =   25
         Top             =   2295
         Width           =   255
      End
      Begin VB.TextBox txtChkRltDate 
         Height          =   270
         Left            =   7140
         MaxLength       =   8
         TabIndex        =   23
         Top             =   2055
         Width           =   975
      End
      Begin VB.TextBox txtItemCount 
         Height          =   270
         Left            =   2655
         MaxLength       =   4
         TabIndex        =   14
         Top             =   1185
         Width           =   345
      End
      Begin VB.CheckBox Check1 
         Caption         =   $"frm040104_f.frx":0038
         Height          =   375
         Index           =   5
         Left            =   -74325
         TabIndex        =   38
         Top             =   1830
         Width           =   7620
      End
      Begin VB.CheckBox Check1 
         Caption         =   "非專利權人申請技術報告，主張本案涉及專利侵權爭議之情事"
         Height          =   210
         Index           =   4
         Left            =   -74775
         TabIndex        =   37
         Top             =   1590
         Width           =   6990
      End
      Begin VB.CheckBox Check1 
         Caption         =   "新型專利業經公告。"
         Height          =   210
         Index           =   0
         Left            =   -74775
         TabIndex        =   33
         Top             =   450
         Value           =   1  '核取
         Width           =   2085
      End
      Begin VB.CheckBox Check1 
         Caption         =   "專利權已當然消滅。"
         Height          =   210
         Index           =   1
         Left            =   -74775
         TabIndex        =   34
         Top             =   720
         Width           =   2130
      End
      Begin VB.CheckBox Check1 
         Caption         =   "專利權人申請技術報告，主張本案涉有非專利權人為商業上實施之情事"
         Height          =   210
         Index           =   2
         Left            =   -74775
         TabIndex        =   35
         Top             =   960
         Width           =   6585
      End
      Begin VB.CheckBox Check1 
         Caption         =   "應檢附有關證明文件：為專利權人對商業上實施之非專利權人之書面通知、廣告目錄或其他商業上實施事實之書面資料。"
         Height          =   645
         Index           =   3
         Left            =   -74325
         TabIndex        =   36
         Top             =   1050
         Width           =   7485
      End
      Begin MSForms.TextBox Text4 
         Height          =   300
         Index           =   0
         Left            =   1215
         TabIndex        =   10
         Top             =   300
         Width           =   7200
         VariousPropertyBits=   671107099
         MaxLength       =   160
         Size            =   "12700;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text4 
         Height          =   300
         Index           =   1
         Left            =   1215
         TabIndex        =   11
         Top             =   600
         Width           =   7200
         VariousPropertyBits=   671107099
         MaxLength       =   250
         Size            =   "12700;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text4 
         Height          =   300
         Index           =   2
         Left            =   1215
         TabIndex        =   12
         Top             =   900
         Width           =   7200
         VariousPropertyBits=   671107099
         MaxLength       =   160
         Size            =   "12700;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   540
         Left            =   6900
         TabIndex        =   17
         Top             =   1188
         Width           =   1500
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;952"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Left            =   1530
         TabIndex        =   20
         Top             =   1725
         Width           =   6870
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "12118;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCP113 
         AutoSize        =   -1  'True
         Caption         =   "工作時數："
         Height          =   180
         Index           =   18
         Left            =   3900
         TabIndex        =   81
         Top             =   2655
         Width           =   900
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "是否電子送件:         (Y:是)"
         Height          =   180
         Index           =   0
         Left            =   3570
         TabIndex        =   80
         Top             =   1500
         Width           =   1995
      End
      Begin VB.Label lblPayToday 
         AutoSize        =   -1  'True
         Caption         =   "電子送件是否當日扣款:         (Y/N)"
         Height          =   180
         Left            =   90
         TabIndex        =   79
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label lblCP84 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Left            =   3285
         TabIndex        =   77
         Top             =   1230
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否申請人為專利權人：　　(N:否)"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   76
         Top             =   1485
         Width           =   2805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否修改申請書內容：　　(Y:Word)"
         Height          =   180
         Index           =   20
         Left            =   3030
         TabIndex        =   75
         Top             =   2070
         Width           =   2850
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否列印申請書：　　(N:不印)"
         Height          =   180
         Index           =   19
         Left            =   90
         TabIndex        =   74
         Top             =   2070
         Width           =   2445
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發文日："
         Height          =   180
         Index           =   18
         Left            =   90
         TabIndex        =   73
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(日):"
         Height          =   180
         Index           =   15
         Left            =   90
         TabIndex        =   72
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(英):"
         Height          =   180
         Index           =   14
         Left            =   90
         TabIndex        =   71
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(中):"
         Height          =   180
         Index           =   13
         Left            =   90
         TabIndex        =   70
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否列印通知函：        (N:不印)"
         Height          =   180
         Index           =   5
         Left            =   90
         TabIndex        =   69
         Top             =   2340
         Width           =   2445
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否修改通知函內容：　　(Y:Word)"
         Height          =   180
         Index           =   6
         Left            =   3015
         TabIndex        =   68
         Top             =   2340
         Width           =   2850
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人:"
         Height          =   180
         Left            =   5925
         TabIndex        =   67
         Top             =   1230
         Width           =   945
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱(中)："
         Height          =   180
         Left            =   90
         TabIndex        =   66
         Top             =   1770
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "催審期限:"
         Height          =   180
         Left            =   6345
         TabIndex        =   65
         Top             =   2070
         Width           =   765
      End
      Begin VB.Label lblCaseFee 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         Caption         =   "@"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   8130
         TabIndex        =   24
         Tag             =   "Y"
         Top             =   1995
         Width           =   255
      End
      Begin VB.Label lblItemCount 
         AutoSize        =   -1  'True
         Caption         =   "項數："
         Height          =   180
         Left            =   2115
         TabIndex        =   64
         Top             =   1245
         Width           =   540
      End
      Begin VB.Label lblCaseFees 
         BackColor       =   &H80000010&
         Height          =   255
         Left            =   8175
         TabIndex        =   78
         Top             =   2055
         Width           =   255
      End
   End
   Begin VB.TextBox txtPA 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   930
      MaxLength       =   3
      TabIndex        =   31
      Top             =   825
      Width           =   495
   End
   Begin VB.TextBox txtPA 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   1410
      MaxLength       =   6
      TabIndex        =   30
      Top             =   825
      Width           =   855
   End
   Begin VB.TextBox txtPA 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   2250
      MaxLength       =   1
      TabIndex        =   29
      Top             =   825
      Width           =   255
   End
   Begin VB.TextBox txtPA 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   2490
      MaxLength       =   2
      TabIndex        =   9
      Top             =   825
      Width           =   375
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   915
      MaxLength       =   1
      TabIndex        =   0
      Top             =   1140
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   5
      Left            =   915
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1680
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   915
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1410
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   5010
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1410
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   5010
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1140
      Width           =   240
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   3
      Left            =   4368
      TabIndex        =   8
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6396
      TabIndex        =   6
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5580
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7620
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   10
      Left            =   5130
      TabIndex        =   62
      Top             =   1725
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4075;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   4260
      TabIndex        =   61
      Top             =   1725
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Index           =   4
      Left            =   6600
      TabIndex        =   60
      Top             =   870
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   8
      Left            =   1245
      TabIndex        =   59
      Top             =   1725
      Width           =   2835
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5001;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   7
      Left            =   5340
      TabIndex        =   58
      Top             =   1455
      Width           =   3165
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5583;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   6
      Left            =   1245
      TabIndex        =   57
      Top             =   1455
      Width           =   2835
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5001;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   5
      Left            =   5340
      TabIndex        =   56
      Top             =   1185
      Width           =   3165
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5583;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   4
      Left            =   1245
      TabIndex        =   55
      Top             =   1185
      Width           =   2835
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5001;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   2
      Left            =   5115
      TabIndex        =   54
      Top             =   870
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
      Index           =   0
      Left            =   930
      TabIndex        =   53
      Top             =   600
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3440;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   52
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   7
      Left            =   4260
      TabIndex        =   51
      Top             =   870
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   3
      Left            =   135
      TabIndex        =   50
      Top             =   870
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Index           =   8
      Left            =   135
      TabIndex        =   49
      Top             =   1185
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Index           =   9
      Left            =   4260
      TabIndex        =   48
      Top             =   1185
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Index           =   10
      Left            =   135
      TabIndex        =   47
      Top             =   1455
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Index           =   11
      Left            =   4260
      TabIndex        =   46
      Top             =   1455
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Index           =   12
      Left            =   135
      TabIndex        =   45
      Top             =   1725
      Width           =   675
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   3
      Left            =   7305
      TabIndex        =   44
      Top             =   870
      Width           =   1185
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2090;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   5115
      TabIndex        =   43
      Top             =   600
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
      Caption         =   "申請國家:"
      Height          =   180
      Index           =   1
      Left            =   4260
      TabIndex        =   32
      Top             =   600
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   0
      X2              =   8600
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   8600
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frm040104_f"
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

Dim intWhere As Integer
'Modify by Morgan 2005/7/15 改用動態陣列
'Dim pa(1 To T_PA) As String, cp(1 To T_CP) As String
Dim pa() As String, cp() As String
Dim m_bolActive As Boolean
Dim i As Integer
Dim m_DiscType As String '減免身分
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/23 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_lngOverItemFeeDiff As Long 'Add By Sindy 2014/5/27 超項費差額(參考Frm040104_3的CheckOfficialFee)
Dim m_lngRecOverItemFee As Long 'Add By Sindy 2014/5/27 已收文超項費
Dim m_lngOverItemFee As Long 'Add By Sindy 2014/6/18 超項費
Dim skMail() As SeekMails 'Add By Sindy 2014/5/27
Dim m_CP64Add As String 'Added by Morgan 2020/10/27
Dim m_bolFMP As Boolean 'Added by Lydia 2023/06/20 是否為FMP案
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/20 是否為寰華

Private Sub cmdOK_Click(Index As Integer)
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
'               PUB_GetCPunIssueDatas "" & txtPA(1).Text & "-" & txtPA(2).Text & "-" & IIf(Len("" & txtPA(3).Text) <= 0, "0", txtPA(3).Text) & "-" & IIf(Len("" & txtPA(4).Text) <= 0, "00", txtPA(4).Text)
'               frm040104_1.Show
'               frm040104_1.Clear
               '若有未發文資料顯示警告
               If PUB_GetCPunIssueDatas("" & txtPA(1).Text & "-" & txtPA(2).Text & "-" & IIf(Len("" & txtPA(3).Text) <= 0, "0", txtPA(3).Text) & "-" & IIf(Len("" & txtPA(4).Text) <= 0, "00", txtPA(4).Text)) Then
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
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            Unload frm040104_1
         End If
         Unload Me
   End Select
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
End Sub

Private Function Process(Index As Integer) As Boolean
   
   Dim stET03 As String, stET01 As String
   'Added by Morgan 2020/10/27
   Dim strFilePath As String '記錄智慧局收文文號
   Dim bolUp As Boolean '是否需要上傳檔案到卷宗區
   'end 2020/10/27
   
   '重新檢查欄位有效性
   If TxtValidate = True Then
   
      'Added by Morgan 2020/10/27 電子送件
      m_CP64Add = ""
      If txtCP118 = "Y" And pa(9) = "000" Then
         txtLetter(0) = "N"
         m_CP123s = ""
         '主管機關
         If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, txtCP27, , True) = False Then
            Exit Function
         End If
        
         strExc(0) = InputBox("請輸入智慧局收文文號!!")
         If strExc(0) = "" Then
            Exit Function
         Else
            strFilePath = strExc(0)  '記錄智慧局收文文號
            m_CP64Add = "智慧局收文文號:" & strExc(0) & ";" '進度備註
         End If
         
         '檢查是否有電子送件的檔案
         bolUp = False
         If strFilePath <> "" Then
            strExc(1) = cp(82) '有發文時間表示重新發文
            If Val(cp(82)) > 0 Then
               If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                  strExc(1) = ""
               End If
            End If
            If Val(strExc(1)) = 0 Then
               If Pub_AutoEsetToCppByP(True, pa(1), pa(2), pa(3), pa(4), pa(8), "", cp(10), strFilePath, txtCP27) = False Then
                  Exit Function
               End If
               bolUp = True
            End If
         End If
         
      Else
      'end 2020/10/27
   
         'Add by Morgan 2009/4/28
         If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, txtCP27) = False Then
            Exit Function
         End If
         If m_CP123s = "Y" Then
         'end 2009/4/28
            'Add by Morgan 2009/3/23 設定是否算發文室案件
            'modify by sonia 2014/6/23 加傳發文規費, P-108903
            If ModifyDispatch(cp(9), m_CP09s, m_CP123s, txtCP84, txtCP27) = False Then
                Exit Function
            End If
         End If
         
      'Added by Morgan 2020/10/27
         'AB類申請書確認檢查,符合條件才可發文
         If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And txtLetter(0) = "N" Then
            If PUB_CheckPDF2(cp(9), 1, True) = False Then
               MsgBox "無申請書PDF檔 ,不可發文!", vbInformation
               Exit Function
            End If
         End If
            
      End If
      'Added by Morgan 2020/10/27
      
      'Add by Amy 2014/10/14 P台灣案發文控制
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
        If pa(1) = "P" And cp(9) < "C" And pa(9) = 台灣國家代號 Then
            If cp(9) < "B" Then
                'A類一定要有接洽單才可發文
                'Modify by Amy 2014/11/27 取消ChkOneDayHasCP27判斷,接洽單改檢查,因考慮可能同時發文其他案件性質情形
                'If PUB_CheckPDF2(cp(9), 0, True, strExc(0)) = False And ChkOneDayHasCP27(pa(1), pa(2), pa(3), pa(4), cp(5) + 19110000) = False Then
                If PUB_CheckPDF3(txtPA(1), txtPA(2), txtPA(3), txtPA(4)) = False Then
                    Exit Function
                End If
            End If
            'AB類申請書確認檢查,符合條件才可發文
            'Modified by Morgan 2015/3/17
            'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And txtLetter(0) = "N" And PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            'Removed by Morgan 2020/10/27 移到上面非電子送件內
            'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And txtLetter(0) = "N" Then
            '   If PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            ''end 2015/3/17
            '      MsgBox "無申請書PDF檔 ,不可發文!", vbInformation
            '      Exit Function
            '   End If 'Added by Morgan 2015/3/17
            'End If
            'end 2020/10/27
        End If
      End If
      'end 2014/10/14
         
      If FormSave = False Then
         MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
      Else
         Process = True
         
         'Add by Morgan 2007/6/14
         If pa(9) = "000" Then
            PUB_ReAsignInform pa(1), pa(2), pa(3), pa(4), cp(9)
         End If
         
         '2012/7/23 add by sonia
         '台灣案發文規費與收文規費不符時,mail給智權人員
         If txtCP84.Enabled = True And pa(9) = "000" And Val(Me.txtCP84.Text) <> Val(cp(17)) Then
            '2013/7/2 modify by sonia 改用共用module
            'Modified by Morgan 2020/10/27 +電子送件自動扣款參數
            PUB_ChkOfficialFee cp(9), Me.txtCP84.Text, IIf(txtCP118 = "Y", "A", "")
         End If
         '2012/7/23 end
         
         BatchMail 'Add By Sindy 2014/5/27
         
         '列印申請書
         If txtLetter(0) <> "N" Then
            stET01 = "01": stET03 = "00"
            If StartLetter(cp(9), stET01, stET03) Then
               'Modify by Amy 2014/08/26 +傳strLetterRecNo 及台灣案申請書修改改開1105_1
               NowPrint cp(9), stET01, stET03, IIf(txtLetter(1) = "Y", True, False), strUserNum, 0, , , , , , , , , , , , cp(9)
               If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And txtLetter(0) <> "N" And txtLetter(1) = "Y" Then
                    frm1105_1.m_RecNo = cp(9)
                    'Modify By Sindy 2022/5/11 流水號要足6碼
                    frm1105_1.m_PdfName = txtPA(1) & txtPA(2) & IIf(txtPA(3) & txtPA(4) = "000", "", "-" & txtPA(3) & "-" & txtPA(4)) & "." & cp(10) & ".DATA.PDF"
                    frm1105_1.Show
                End If
                'end 2014/08/26
                
               '直接列印申請書
               If txtLetter(1) = "" Then PrintLetter cp(9)
            End If
         End If
         '列印通知函
         If txtLetter(3) <> "N" Then
            stET01 = "02": stET03 = "00"
            If StartLetter1(cp(9), stET01, stET03) Then
               'Modify by Amy 2014/08/26 +傳strLetterRecNo
               NowPrint cp(9), stET01, stET03, IIf(txtLetter(4) = "Y", True, False), strUserNum, 0, , , , , , , , , , , , cp(9)
            End If
         End If
         
         'Added by Morgan 2020/10/27
         '上傳檔案
         If bolUp = True Then
            If Pub_AutoEsetToCppByP(False, pa(1), pa(2), pa(3), pa(4), pa(8), cp(9), cp(10), strFilePath, txtCP27) = False Then
                 Exit Function
            End If
         End If
         'end 2020/10/27
      End If
   End If
   
End Function
'列印定稿
Private Sub PrintLetter(ByVal strReceiveNo As String)

   Dim stLD03 As String
   
On Error GoTo ErrHnd

   strSql = "select LD03 from letterdemand where ld01='" & strUserNum & "' and ld02=" & strSrvDate(1) & " and ld04='" & strReceiveNo & "' order by ld03 desc"
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         stLD03 = "" & .Fields("LD03")
         PrinterLetterDB strUserNum, "", strSrvDate(1), stLD03
         strSql = "Update letterdemand Set LD16='*' Where ld01='" & strUserNum & "' AND ld02=" & strSrvDate(1) & " and ld03=" & stLD03
         cnnConnection.Execute strSql
      End If
   End With
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
   
End Sub
'Add by Morgan 2006/6/2
Private Function StartLetter1(ByVal strReceiveNo As String, ByVal ET01 As String, ByVal ET03 As String) As Boolean
   
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   strExc(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "select '" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','來函主管機關',CF10 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & cp(10) & "'"
   'edit by nickc 2007/02/05 不用 dll 了
   'If objLawDll.ExecSQL(1, strExc) Then
   If ClsLawExecSQL(1, strExc) Then
       StartLetter1 = True
   Else
       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Function

Private Function StartLetter(ByVal strReceiveNo As String, ByVal ET01 As String, ByVal ET03 As String) As Boolean

   Dim strTxt(1 To 30) As String, strTmp As String, strTmp1 As String
   Dim iAppCnt As Integer
   Dim stAppData(1 To 1, 0 To 3) As String
   Dim ii As Integer, iLen As Integer
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   '1 發文日
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','發文日','" & ChangeTStringToTDateString(ChangeWStringToTString(cp(27))) & "')"
      
   '2 申請人數
   iAppCnt = 1
   For i = 27 To 30
      If pa(i) <> "" Then
         iAppCnt = iAppCnt + 1
      End If
   Next
   
   strTmp = Format(iAppCnt) '申請人數
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','申請人數','" & strTmp & "')"
    
   '3~12 國籍
   For i = 1 To iAppCnt
      Erase stAppData
      Call PUB_GetAppData(pa(25 + i), stAppData, 1)
      strTmp = IIf(Val(stAppData(1, 2)) < 10, "中華民國", stAppData(1, 3))
      strTmp1 = Label2(3 + i) & "　ID：" & stAppData(1, 0)
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','申請人" & Format(i) & "的國籍','" & strTmp & "')"
         
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','申請人" & Format(i) & "的名稱','" & strTmp1 & "')"
   Next
  
    
   '12~13 是否申請人為專利權人
   If txtLetter(2) = "N" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','勾選2','■ ')"
         
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','勾選1','□ ')"
   
   Else
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','勾選1','■ ')"
         
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','勾選2','□ ')"
   End If
   
   'Modified by Morgan 2013/1/10 +7~12
   '14~23 相關事項 勾選3~12
   For i = 0 To 9
      If Check1(i).Value = 1 Then
         strTmp = "■ "
      Else
         strTmp = "□ "
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','勾選" & Format(i + 3) & "','" & strTmp & "')"
   Next
   
   'Added by Morgan 2013/1/10
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','項數','" & Val(txtItemCount) & "')"
   'end 2013/1/10
         
   '發文規費
   strTmp = Format(txtCP84)
   iLen = Len(strTmp)
   strTmp1 = ""
   For i = 1 To iLen
      If Left(strTmp, 1) <> 0 Then
         strTmp1 = strTmp1 & ShowNumberWord(Left(strTmp, 1))
         If Len(strTmp) = 5 Then
            strTmp1 = strTmp1 & "萬"
         ElseIf Len(strTmp) = 4 Then
            strTmp1 = strTmp1 & "千"
         ElseIf Len(strTmp) = 3 And Right(strTmp, 1) <> " 零" Then
            strTmp1 = strTmp1 & "百"
         ElseIf Len(strTmp) = 2 And Right(strTmp, 1) <> " 零" Then
            strTmp1 = strTmp1 & "拾"
         End If
      End If
      strTmp = Mid(strTmp, 2)
   Next
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','發文規費','" & strTmp1 & "')"
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If objLawDll.ExecSQL(ii, strTxt) Then
   If ClsLawExecSQL(ii, strTxt) Then
      StartLetter = True
   Else
       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Function

Private Function TxtValidate() As Boolean
   Dim oText As Object, Cancel As Boolean
   Dim bYesName As Boolean '是否有案件名稱
   
   'Added by Morgan 2021/12/15 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/15
   
   'add by nickc 2008/05/01
   If IsDebt(pa(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        Exit Function
   End If
   '案件名稱
   bYesName = False
   For Each oText In Text4
      If oText.Enabled = True Then
         Cancel = False
         Text4_Validate oText.Index, Cancel
         If Cancel = True Then
            Text4(oText.Index).SetFocus
            Text4_GotFocus oText.Index
            Exit Function
         End If
      End If
      If oText.Text <> "" Then
         bYesName = True
      End If
   Next
   
   'Add by Morgan 2007/12/14
   If bYesName = False Then
      MsgBox "案件名稱不可同時空白 !", vbCritical
      Text4(0).SetFocus
      Text4_GotFocus 0
      Exit Function
   End If
   'end 2007/12/14
   
   If pa(9) = "000" Then
      m_DiscType = ""
      For i = 1 To 5
         m_DiscType = m_DiscType & txtAD(i).Text
         If txtAD(i).Enabled = True Then
            If txtAD(i).Text = "" Then
               'Add by Morgan 2007/8/30 第三人申請技術報告不用
               If cp(10) <> "807" Then
                  MsgBox "申請人【" & pa(25 + i) & "-" & Label2(i + 3) & "】減免身分不可空白", vbInformation
                  txtAD(i).SetFocus
                  txtAD_GotFocus i
                  Exit Function
               End If
            '公司可減免
            'Modify by Morgan 2004/7/29
            '學校不需證明
            'ElseIf (txtAD(i).Text = "2" Or txtAD(i).Text = "3") Then
            '學校
            ElseIf (txtAD(i).Text = "2") Then
               '變更
               If (txtAD(i).Tag <> "2" And txtAD(i).Tag <> "") Then
                  If MsgBox("確定要變更申請人【" & pa(25 + i) & "-" & Label2(i + 3) & "】減免身分為【學校】？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                     txtAD(i).SetFocus
                     txtAD_GotFocus i
                     Exit Function
                  End If
               End If
            '公司
            ElseIf (txtAD(i).Text = "3") Then
               '新增或變更
               If (txtAD(i).Tag <> "3") Then
                  If MsgBox("申請人【" & pa(25 + i) & "-" & Label2(i + 3) & "】的減免身分將設定為【中小企業】，確定有【證明文件】存放於本卷？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                     txtAD(i).SetFocus
                     txtAD_GotFocus i
                     Exit Function
                  End If
               End If
            '不可減免
            ElseIf (txtAD(i).Text = "N") Then
               '身分變更
               If (txtAD(i).Tag <> "N" And txtAD(i).Tag <> "") Then
                  If MsgBox("確定要變更申請人【" & pa(25 + i) & "-" & Label2(i + 3) & "】減免身分為【不可減免】？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                     txtAD(i).SetFocus
                     txtAD_GotFocus i
                     Exit Function
                  End If
               End If
            End If
         End If
      Next
      If m_DiscType <> "" Then
         If InStr(m_DiscType, "N") > 0 Then
            cp(81) = "N"
         Else
            cp(81) = "Y"
         End If
      End If
      
      'Added by Morgan 2013/1/3
      If txtItemCount = "" Then
         MsgBox "請輸入項數！", vbInformation
         txtItemCount.SetFocus
         Exit Function
      End If
      
      'Add By Sindy 2014/5/27
      m_lngOverItemFeeDiff = 0
      If pa(9) = "000" Then '台灣案
         If Not CheckOfficialFee Then
            Exit Function
         End If
      End If
      '2014/5/27 END
      
      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！" & vbCrLf & vbCrLf & "( 10項以內規費 $5000, 每超出1項加收 $600 )", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            txtCP84.Tag = txtCP84.Text
         Else
            txtItemCount.SetFocus
            Exit Function
         End If
      End If
      'end 2013/1/3
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
   
   'Added by Morgan 2020/10/23
   If txtCP118 = "Y" And txtCP22 = "N" And pa(9) = "000" Then
      MsgBox "電子送件不可不出名！", vbCritical
      Exit Function
   End If
   If txtCP118 = "Y" And Val(txtCP84) > 0 And pa(9) = "000" Then
      If txtPayToday = "" Then
         MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
         txtPayToday.SetFocus
         Exit Function
      End If
   End If
   'end 2020/10/23
   
   
   'Add by Morgan 2007/8/30
   If cp(10) = "807" Then
      If Text5 = "" Then
         MsgBox "對造名稱不可空白 !", vbCritical
         Text5.SetFocus
         Exit Function
      End If
   End If
   
   cp(40) = Text5
   'end 2007/8/30
   
   '案件名稱
   pa(5) = CNULL(ChgSQL(Text4(0)))
   pa(6) = CNULL(ChgSQL(Text4(1)))
   pa(7) = CNULL(ChgSQL(Text4(2)))
   '發文日
   cp(27) = TransDate(txtCP27, 2)
   '是否出名
   cp(22) = txtCP22
   
   'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
   If Pub_ChkACS112isNull(pa(1), pa(2), pa(3), pa(4), txtCP113) = True Then
         txtCP113.SetFocus
         txtCP113_GotFocus
         Exit Function
   End If
   'end 2021/05/25
      
   'Added by Morgan 2024/1/24
   If pa(9) = "020" And txtCP118 = "" Then
      If MsgBox("請確認本案是否為紙本送件？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         txtCP118.SetFocus
         Exit Function
      End If
   End If
   'end 2024/1/24

   TxtValidate = True
End Function

'Add By Sindy 2014/5/27
'檢查規費
Private Function CheckOfficialFee() As Boolean
   Dim lngRecFee As Long, lngFeeDif As Long
   
   CheckOfficialFee = True
   m_lngRecOverItemFee = 0
   
   strExc(0) = "select cp10,sum(nvl(cp17,0)-nvl(cp77,0)) as Fee from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='939' and cp27 is null and cp57 is null group by cp10"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Do While Not RsTemp.EOF
         '已收超項費
         If RsTemp.Fields("cp10") = "939" Then m_lngRecOverItemFee = m_lngRecOverItemFee + Val("" & RsTemp.Fields("Fee"))
         RsTemp.MoveNext
      Loop
   End If
   '已收規費
   lngRecFee = Val("" & cp(17)) + m_lngRecOverItemFee
   '不足規費
   lngFeeDif = Val(txtCP84) - lngRecFee
   
   If lngFeeDif > 0 Then
      m_lngOverItemFeeDiff = lngFeeDif
      If MsgBox("已收文規費(含未發文超項費)共【" & Format(lngRecFee, DAmount) & "】元，尚缺【" & Format(lngFeeDif, DAmount) & "】元，存檔時會自動內部收文超項費補足差額，是否確定要繼續？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
         CheckOfficialFee = False
      End If
   'Add By Sindy 2014/6/18
   ElseIf lngFeeDif < 0 Then
      MsgBox "已收文規費(含未發文超項費)共【" & Format(lngRecFee, DAmount) & "】元，超收【" & Format(-1 * lngFeeDif, DAmount) & "】元！", vbExclamation
   '2014/6/18 END
   End If
End Function

'Add By Sindy 2014/5/27 批次發Mail
Private Sub BatchMail()
   Dim i As Integer
   For i = 1 To UBound(skMail)
        PUB_SendMail skMail(i).fiSender, skMail(i).fiReceiver, skMail(i).fiRecriverNo, skMail(i).fiSubject, skMail(i).fiContent
   Next i
   ReDim skMail(0) As SeekMails
End Sub

Private Sub Form_Activate()
   '若沒有客戶減免身分需輸入則游標預設在是否出名
   If m_bolActive = True Then Exit Sub
   m_bolActive = True
   
   If pa(9) = "000" Then
      For i = 1 To 5
         If txtAD(i).Enabled = True And txtAD(i).Text = "" Then
            txtAD(i).SetFocus
            Exit Sub
         End If
      Next
   End If
   
   'Modified by Morgan 2013/1/3
   'If txtCP84.Enabled = True Then txtCP84.SetFocus
   'Modify by Amy 2018/10/05 原:txtItemCount.Visible 因前畫面有強制視窗,Form 為Enabled,故會當掉
   If txtItemCount.Enabled = True Then
      txtItemCount.SetFocus
   ElseIf txtCP84.Enabled = True Then
      txtCP84.SetFocus
   End If
   'end 2013/1/3
End Sub

Private Sub Form_Initialize()
    'Add by Morgan 2005/7/15
    ReDim pa(TF_PA)
    ReDim cp(TF_CP)
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內

    
   '本所案號
   With frm040104_1
      pa(1) = .Text1: txtPA(1) = pa(1)
      pa(2) = .Text2: txtPA(2) = pa(2)
      pa(3) = .Text3: txtPA(3) = pa(3)
      pa(4) = .Text4: txtPA(4) = pa(4)
      cp(9) = .Tag
   End With
   ReadPatent
   
   cp(110) = "" '要清空,否則若重新發文會殘留前次發文資料,當新案有改出名人而本程序未改選將會造成不一致 Added by Morgan 2012/9/7
   
   'Add by Morgan 2005/7/15
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   txtCP22.Visible = False
   lstNameAgent.Clear
   If pa(9) = "000" Then
      PUB_SetOurAgent lstNameAgent, pa(), cp(110), , True 'Modified by Morgan 2021/12/15 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2005/7/14 END
   
   'Add by Morgan 2007/8/30
   Text5 = cp(40)
   txtLetter(2).Enabled = False
   If cp(10) = "421" Then
      txtLetter(2) = ""
      Text5.Enabled = False
   Else
      txtLetter(2) = "N"
   End If
   'end 2007/8/30
   'Add by Amy 2014/08/26 台灣案客戶函不可修改 for 電子化
   If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And pa(9) = 台灣國家代號 Then
        txtLetter(4).Enabled = False
   End If
   SSTab1.Tab = 0
End Sub

Private Sub ReadPatent()
Dim strAD10 As String, strCU15 As String, oLabel As Object
Dim m_Fee As String         '銷帳服務費 2012/8/1 add by sonia
Dim m_Official As String    '銷帳規費   2012/8/1 add by sonia
   
   '清除舊內容
   For Each oLabel In Label2
      oLabel.Caption = ""
   Next

   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      '申請人
      For i = 26 To 30
         If pa(i) <> "" Then ChgType i
      Next
      
      '申請國家
      If pa(9) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(pA(9), strExc(0)) Then Label2(1) = strExc(0)
         If ClsPDGetNation(pa(9), strExc(0)) Then Label2(1) = strExc(0)
      End If
      
      '專利名稱
      Text4(0) = pa(5): Text4(1) = pa(6): Text4(2) = pa(7)
            
   End If
   
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
   
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      '收文號
      Label2(0) = cp(9)
      
      '智權人員
      If cp(13) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(13), strExc(0)) Then Label2(2) = strExc(0)
         If ClsPDGetStaff(cp(13), strExc(0)) Then Label2(2) = strExc(0)
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
      '承辦人
      If cp(14) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(14), strExc(0)) Then
         If ClsPDGetStaff(cp(14), strExc(0)) Then
               Label2(3) = strExc(0)
         End If
      End If
      txtCP22 = cp(22)
      '2012/8/1 add by sonia 若有銷帳則要扣除銷帳規費
      If Val(cp(77)) > 0 Then
         If GetCP77Detail(cp(9), m_Fee, m_Official) = True Then
            cp(17) = cp(17) - m_Official
         End If
      End If
      '2012/8/1 end
      txtCP84.Tag = cp(17)
   End If
   
   '發文日
   txtCP27 = strSrvDate(2)
   '案件性質
   ClsPDGetCaseProperty pa(1), cp(10), strExc(1), IIf(pa(9) = "000", False, True)
   Label2(10).Caption = strExc(1)
   
   'Add by Morgan 2009/8/17
   If txtCP27 <> "" Then
      PUB_SetChkResultDate pa(1), pa(9), cp(10), txtCP27, txtChkRltDate, cp, pa(8)
      txtCP27.Tag = txtCP27
   End If
   
   'Added by Morgan 2010/10/27
   txtCP118 = ""
   txtPayToday = ""
   If pa(9) = "000" Then
      txtCP118 = "Y" '預設電子送件
      '電子送件一律自動扣款(A)若超過3點半發文則須人工輸入是否當日扣款
      If txtCP118 = "Y" Then
         If Val(ServerTime) <= 153000 Then
            txtPayToday = "Y"
         End If
      End If
   Else
      'Added by Morgan 2024/1/14 大陸案預設電子送件--郭
      If pa(9) = "020" Then
         'Modified by Morgan 2024/1/25 有設定大陸P案要公文正本者預設紙本送件--郭
         'Modified by Morgan 2024/1/30 改分案預設--郭
         'If PUB_GetCustomerValue(pa(26), "CU182") = "Y" Then
         '   txtCP118 = ""
         'Else
         '   txtCP118 = "Y"
         'End If
         If cp(118) <> "" Then txtCP118 = "Y"
         'end 2024/1/30
      Else
      'end 2024/1/24
         txtCP118.Enabled = False
      End If
      txtPayToday.Enabled = False
   End If
   'end 2020/10/27
   
    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25
End Sub
Private Function ChgType(iSitu As Integer) As Boolean

   ChgType = False
   Select Case iSitu
         
      '申請人中文
      Case 26, 27, 28, 29, 30
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(pa(iSitu), strExc(0)) Then
         If ClsLawLawGetName(pa(iSitu), strExc(0)) Then
            Label2(iSitu - 22) = strExc(0)
            ChgType = True
         End If
         
      Case Else
         ChgType = True
         
   End Select
   
End Function


Private Sub Form_Unload(Cancel As Integer)
   'Set frm040104_f = Nothing 'Removed by Morgan 2021/12/15 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub Text4_GotFocus(Index As Integer)
   Select Case Index
      Case 0, 2
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text4(Index).IMEMode = 1
         OpenIme
      Case 1
         'edit by nickc 2007/07/11 切換輸入法改用API
         'Text4(Index).IMEMode = 2
         CloseIme
   End Select
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
   If Cancel = True Then Text4_GotFocus Index
End Sub

Private Sub txtAD_GotFocus(Index As Integer)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtAD(Index).IMEMode = 2
   CloseIme
   TextInverse txtAD(Index)
End Sub

'只有公司可輸入 2,3
Private Sub txtAD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2014/7/15 學校改預設且不可改
   'If Not (KeyAscii = 8 Or KeyAscii = 50 Or KeyAscii = 51 Or KeyAscii = 78) Then
   If Not (KeyAscii = 8 Or KeyAscii = 51 Or KeyAscii = 78) Then
      KeyAscii = 0
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

Private Sub txtCP118_GotFocus()
   TextInverse txtCP118
   CloseIme
End Sub

Private Sub txtCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
'Add by Morgan 2009/8/17
Private Sub txtCP27_Validate(Cancel As Boolean)
   If txtCP27 <> "" Then
      '2011/12/8 MODIFY BY SONIA 發文日可輸系統日的下一個工作日
      'If ChkDate(txtCP27) = False Then
      If Not ChkDate(txtCP27) Or DBDATE(Val(txtCP27)) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
         MsgBox "發文日期不正確或發文日大於系統日下一個工作日，請重新輸入 !", vbCritical
      '2011/12/8 END
         Cancel = True
      Else
         If txtCP27.Tag <> txtCP27 Then
            PUB_SetChkResultDate pa(1), pa(9), cp(10), txtCP27, txtChkRltDate, cp, pa(8)
            txtCP27.Tag = txtCP27
            
            'Added by Morgan 2020/10/27
            '當發文日有改時,電子送件案要人工輸入是否當日扣款
            If txtCP118 = "Y" Then
               txtPayToday.Text = ""
            End If
            'end 2020/10/27
         End If
      End If
   End If
End Sub

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP84_Validate(Cancel As Boolean)
   If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
      If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
         txtCP84.Tag = txtCP84.Text
      Else
         Cancel = True
      End If
   End If
End Sub

Private Sub txtItemCount_GotFocus()
   TextInverse txtItemCount
   CloseIme
End Sub

Private Sub txtItemCount_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
   End If
End Sub

'Added by Morgan 2013/1/3
Private Sub txtItemCount_Validate(Cancel As Boolean)
   If pa(9) = "000" Then
      If txtItemCount = "" Then
         MsgBox "請輸入項數！", vbInformation
         Cancel = True
      Else
         'Modify By Sindy 2014/6/18 + m_lngOverItemFee 超項費
         'txtCP84 = PUB_GetReportFee(pa(1), pa(9), cp(10), Val(txtItemCount))
         txtCP84 = PUB_GetReportFee(pa(1), pa(9), cp(10), Val(txtItemCount), m_lngOverItemFee)
      End If
   End If
End Sub

Private Sub txtLetter_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      '是否列印申請書 N, 是否申請人為專利權人 N
      Case 0, 2, 3
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      '是否列印申請書 Y
      Case 1, 4
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Function FormSave() As Boolean
   Dim ii As Integer
   Dim stCP12 As String, stCP13 As String '最新收文智權人員
   Dim stCP118 As String, stCP152 As String 'Added by Morgan 2020/10/27
   Dim strTo As String 'Add by Amy 2024/05/15
   
On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
   'Add By Sindy 2014/5/27
   ReDim skMail(0) As SeekMails
   stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
   stCP12 = GetSalesArea(stCP13)
   '2014/5/27 END
   
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
   
   '更新基本檔
   strSql = "UPDATE PATENT SET PA05=" & pa(5) & ",PA06=" & pa(6) & ",PA07=" & pa(7) & _
      " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   cnnConnection.Execute strSql


   'Added by Morgan 2020/10/27
   stCP118 = txtCP118
   stCP152 = ""
   If txtCP118 = "Y" And Val(txtCP84) > 0 And pa(9) = "000" Then
      If txtPayToday <> "" Then
         stCP118 = "A"
         If txtPayToday = "Y" Then
            stCP152 = CompWorkDay(2, DBDATE(txtCP27))
         Else
            stCP152 = CompWorkDay(3, DBDATE(txtCP27))
         End If
      End If
   End If
   'end 2020/10/27
   
   '更新進度檔
   'Modify by Morgan 2005/7/15 加 cp110
   'Modify by Morgan 2007/8/30 加 cp40
   'Modified by Morgan 2013/1/3 +cp136
   'Modified by Morgan 2020/10/27 +CP64,CP118,CP120,CP152
   'Modified by Lydia 2021/05/25 +CP113工作時數
   'Modified by Lydia 2023/06/20 +CP14
   'Modified by Morgan 2024/1/24 CP120加判斷台灣案
   strSql = "UPDATE CASEPROGRESS SET CP27=" & cp(27) & ",CP22=" & CNULL(cp(22)) & ",CP40=" & CNULL(ChgSQL(cp(40))) & ",CP64='" & ChgSQL(m_CP64Add) & "'||CP64" & _
   ",CP81=" & CNULL(cp(81)) & ",cp84=" & Format(Val(txtCP84.Text)) & ",cp110=" & CNULL(cp(110)) & ",CP136=" & Val(txtItemCount) & _
   ",CP118='" & stCP118 & "',CP120='" & IIf(stCP118 <> "" And pa(9) = "000", "Y", "") & "',CP152='" & stCP152 & "'" & _
   ", cp113=" & CNULL(txtCP113, True) & ",CP14=" & CNULL(cp(14)) & _
   " WHERE CP09='" & cp(9) & "'"
   cnnConnection.Execute strSql, intI
   
   'Add By Sindy 2014/5/27 超過10項請收文939超項費,每一項600元
   'Add By Sindy 2014/6/18 更新已收文的超項費資料
   If m_lngOverItemFee > 0 And m_lngRecOverItemFee > 0 Then
      strSql = "Update caseprogress set cp27=" & cp(27) & ",cp84=0" & ",cp22=" & CNULL(cp(22)) & ",cp110=" & CNULL(cp(110)) & _
         " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='939' and cp27||cp57 is null"
      cnnConnection.Execute strSql, intI
   End If
   '2014/6/18 END
   If m_lngOverItemFeeDiff > 0 Then
      strExc(1) = AutoNo("B", 6) 'B類總收文號
      strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp16,cp17,cp18,cp26,cp27,cp43,cp84,cp22,cp110) values " & _
               " ('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & DBDATE(txtCP27) & _
               ",'" & strExc(1) & "','939','" & stCP12 & "','" & stCP13 & "'," & CNULL(cp(14)) & "," & m_lngOverItemFeeDiff & "," & m_lngOverItemFeeDiff & ",0,'N'" & _
               "," & strSrvDate(1) & ",'" & cp(9) & "',0," & CNULL(txtCP22) & "," & CNULL(cp(110)) & ")"
      cnnConnection.Execute strSql, intI
      '文雄要求同時上完稿日及會稿完成日,雅娟也同意
      strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & ",EP09=" & strSrvDate(1) & ",EP08=" & strSrvDate(1) & " WHERE EP02='" & strExc(1) & "'"
      cnnConnection.Execute strSql, intI
   End If
   '2014/5/27 END
   
   '設定客戶減免身分
   If pa(9) = "000" Then
      For i = 1 To 5
         If txtAD(i).Enabled = True Then
            '身分有變更才要做
            If txtAD(i).Tag <> txtAD(i).Text Then
               '不可減免
               If txtAD(i).Text = "N" Then
                  strSql = PUB_GetADSQL(pa(25 + i), pa(9), "N")
               '自然人
               '學校也不用證明
               ElseIf (txtAD(i).Text = "1" Or txtAD(i).Text = "2") Then
                  strSql = PUB_GetADSQL(pa(25 + i), pa(9), "Y", txtAD(i).Text)
               '公司
               Else
                  '原來沒有減免資料或不可減免
                  If txtAD(i).Tag = "" Or txtAD(i).Tag = "N" Then
                     strSql = PUB_GetADSQL(pa(25 + i), pa(9), "Y", txtAD(i).Text, pa(1), pa(2), pa(3), pa(4))
                  '修改減免身分別
                  Else
                     strSql = PUB_GetADSQL(pa(25 + i), pa(9), "Y", txtAD(i).Text)
                  End If
               End If
               cnnConnection.Execute strSql
            End If
         End If
      Next
   End If
   
    'Add by Amy 2014/09/09 for 台灣案電子化
    If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And pa(9) = 台灣國家代號 Then
        cnnConnection.Execute "delete LetterProgress where lp01='" & cp(9) & "'", intI 'Added by Morgan 2016/2/26 可能會重新發文
        '*沒出客戶通知函
        If txtLetter(3) = "N" Then
            'Modify by Amy 2015/02/13 原:判斷同一天 沒有其他有規費的發文
              '1.    電子送件且規費>0,有收據
              '2.非電子送件且經發文室要計件,有回執
            'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'            strExc(1) = PUB_GetLetterJudge(pa(1), cp(10))
'            If cp(118) = "Y" Then
'                If Val(txtCP84) > 0 Then
'                    PUB_AddLetterProgress cp(9), 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                End If
'            Else
'                If Left(m_CP123s, 1) = "Y" Then
'                    PUB_AddLetterProgress cp(9), 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                End If
'            End If
'            'end 2015/02/13
          
        '*有出客戶通知函
        Else
            'Modified by Morgan 2018/8/1
            'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , , pa(1), pa(2), pa(3), pa(4))
            strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10))
            'Modify by Amy 2015/02/13 修改、整理判斷條件
            'PUB_AddLetterProgress cp(9), 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
              '1.　電子送件有規費的有收據；無規費的無回執
              '2.非電子送件要計件的有回執；不計件的無回執
            'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
            'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10)) 'Removed by Morgan 2016/6/7 程式重複
'            If cp(118) = "Y" Then
'                If Val(txtCP84) > 0 Then
'                    PUB_AddLetterProgress cp(9), 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
'                Else
'                    PUB_AddLetterProgress cp(9), 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
'                End If
'            Else
                If Left(m_CP123s, 1) = "Y" Then
                    PUB_AddLetterProgress cp(9), 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
                Else
                    PUB_AddLetterProgress cp(9), 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
                End If
'            End If
'            'end 2015/02/13
            'end 2015/03/06
        End If
        '*有申請書
        If txtLetter(0) <> "N" Then
            If ExistCheck("AppForm", "AF01", cp(9), "", False) = False Then
                 '新增申請書轉檔記錄
                 PUB_AddAppForm cp(9)
            End If
        End If
    End If
    'end 2014/09/09
      
   'Add by Morgan 2009/3/23
   PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130
   'Add by Amy 2015/02/13 更新收據/回執設定
   'Modify by Amy 2015/03/06 +發文日參數
   PUB_UpdateLP19 cp(1), cp(2), cp(3), cp(4), m_CP09s, m_CP123s, txtCP27
   
   'Add by Morgan 2009/8/17
   If txtChkRltDate <> "" Then
      PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9), cp(10), cp(43)
   End If
   
   cnnConnection.CommitTrans
   FormSave = True
   
   'Add By Sindy 2014/5/27 補收文超項費發MAIL通知智權人員及財務處
   If m_lngOverItemFeeDiff > 0 Then
      'Modify by Amy 2024/05/15 財務2個特殊設定拆成3個
      If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
         strTo = Pub_GetSpecMan("財務處應收處理人員")
      Else
         strTo = Pub_GetSpecMan("財務處總帳人員")
      End If
      ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
      skMail(UBound(skMail)).fiSender = strUserNum
      skMail(UBound(skMail)).fiReceiver = stCP13 & ";" & strTo
      'end 2024/05/15
      skMail(UBound(skMail)).fiRecriverNo = ""
      skMail(UBound(skMail)).fiSubject = Me.txtPA(1).Text & "-" & Me.txtPA(2).Text & "-" & Me.txtPA(3).Text & "-" & Me.txtPA(4).Text & "的" & Label2(10) & "已發文，尚有"
      If m_lngOverItemFeeDiff > 0 Then skMail(UBound(skMail)).fiSubject = skMail(UBound(skMail)).fiSubject & " 超項費 "
      skMail(UBound(skMail)).fiSubject = skMail(UBound(skMail)).fiSubject & "需向申請人收取, 系統已做內部收文財務處將開立收據！謝謝！"
      skMail(UBound(skMail)).fiContent = "本所案號：" & Me.txtPA(1).Text & "-" & Me.txtPA(2).Text & "-" & Me.txtPA(3).Text & "-" & Me.txtPA(4).Text & vbCrLf & "案件名稱：" & Text4(0) & vbCrLf & vbCrLf
      If m_lngOverItemFeeDiff > 0 Then skMail(UBound(skMail)).fiContent = skMail(UBound(skMail)).fiContent & "超項費：" & Format(m_lngOverItemFeeDiff, "###,##0") & vbCrLf & vbCrLf
      skMail(UBound(skMail)).fiContent = skMail(UBound(skMail)).fiContent & "注意：１．請財務處開立收據交智權人員" & vbCrLf & "　　　２．請智權人員不要再填接洽單"
   End If
   '2014/5/27 END
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If

End Function
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
      txtCP22 = ""
   Else
      txtCP22 = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub
'Add by Morgan 2009/8/17
Private Sub lblCaseFee_Click()
   frm12040102_2.txtCF(1) = cp(1)
   frm12040102_2.txtCF(2) = pa(9)
   frm12040102_2.txtCF(3) = cp(10)
   frm12040102_2.Show vbModal
   If Val(txtCP27) > 0 Then
      PUB_SetChkResultDate pa(1), pa(9), cp(10), txtCP27, txtChkRltDate, cp, pa(8)
   End If
End Sub

Private Sub lblCaseFee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblCaseFee, lblCaseFees
End Sub

Private Sub lblCaseFee_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblCaseFee, lblCaseFees
End Sub

Private Sub txtChkRltDate_Validate(Cancel As Boolean)
   If txtChkRltDate <> "" Then
      If ChkDate(txtChkRltDate) = False Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtPayToday_GotFocus()
   TextInverse txtPayToday
   CloseIme
End Sub

Private Sub txtPayToday_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
