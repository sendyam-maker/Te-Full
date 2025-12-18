VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050102_5 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文（實審、答辯、修正、提供前案資料、選取、讓與、繼承）"
   ClientHeight    =   5640
   ClientLeft      =   456
   ClientTop       =   996
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8520
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   3510
      MaxLength       =   4
      TabIndex        =   10
      Top             =   2790
      Width           =   540
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "微個體"
      Height          =   255
      Index           =   2
      Left            =   2550
      TabIndex        =   18
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CheckBox chkChoose 
      Caption         =   "一併提供前案資料"
      Height          =   180
      Index           =   1
      Left            =   6030
      TabIndex        =   12
      Top             =   2820
      Width           =   2115
   End
   Begin VB.CheckBox chkChoose 
      Caption         =   "一併繳指定費"
      Height          =   180
      Index           =   2
      Left            =   4320
      TabIndex        =   11
      Top             =   2820
      Width           =   1395
   End
   Begin VB.TextBox txtChkRltDate 
      Height          =   270
      Left            =   930
      MaxLength       =   8
      TabIndex        =   8
      Top             =   2790
      Width           =   975
   End
   Begin VB.ListBox lstItems 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      ItemData        =   "frm050102_5.frx":0000
      Left            =   5760
      List            =   "frm050102_5.frx":0007
      Sorted          =   -1  'True
      Style           =   1  '項目包含核取方塊
      TabIndex        =   19
      Top             =   4350
      Width           =   2610
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1185
      Left            =   90
      TabIndex        =   62
      Top             =   3120
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   2096
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "受讓人1"
      TabPicture(0)   =   "frm050102_5.frx":0015
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblAppName(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAppAddr(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAppNew(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAD(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "受讓人2"
      TabPicture(1)   =   "frm050102_5.frx":0031
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtAD(2)"
      Tab(1).Control(1)=   "txtAppNew(2)"
      Tab(1).Control(2)=   "txtAppAddr(2)"
      Tab(1).Control(3)=   "lblAppName(2)"
      Tab(1).Control(4)=   "Label2(1)"
      Tab(1).Control(5)=   "Label13(1)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "受讓人3"
      TabPicture(2)   =   "frm050102_5.frx":004D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtAD(3)"
      Tab(2).Control(1)=   "txtAppNew(3)"
      Tab(2).Control(2)=   "txtAppAddr(3)"
      Tab(2).Control(3)=   "lblAppName(3)"
      Tab(2).Control(4)=   "Label13(4)"
      Tab(2).Control(5)=   "Label2(4)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "受讓人4"
      TabPicture(3)   =   "frm050102_5.frx":0069
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtAD(4)"
      Tab(3).Control(1)=   "txtAppNew(4)"
      Tab(3).Control(2)=   "txtAppAddr(4)"
      Tab(3).Control(3)=   "lblAppName(4)"
      Tab(3).Control(4)=   "Label13(3)"
      Tab(3).Control(5)=   "Label2(5)"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "受讓人5"
      TabPicture(4)   =   "frm050102_5.frx":0085
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtAD(5)"
      Tab(4).Control(1)=   "txtAppNew(5)"
      Tab(4).Control(2)=   "txtAppAddr(5)"
      Tab(4).Control(3)=   "lblAppName(5)"
      Tab(4).Control(4)=   "Label13(2)"
      Tab(4).Control(5)=   "Label2(6)"
      Tab(4).ControlCount=   6
      Begin VB.TextBox txtAD 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   5
         Left            =   -74100
         MaxLength       =   1
         TabIndex        =   40
         Top             =   450
         Width           =   240
      End
      Begin VB.TextBox txtAD 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   4
         Left            =   -74100
         MaxLength       =   1
         TabIndex        =   36
         Top             =   450
         Width           =   240
      End
      Begin VB.TextBox txtAD 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   3
         Left            =   -74100
         MaxLength       =   1
         TabIndex        =   30
         Top             =   450
         Width           =   240
      End
      Begin VB.TextBox txtAD 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   2
         Left            =   -74100
         MaxLength       =   1
         TabIndex        =   25
         Top             =   450
         Width           =   240
      End
      Begin VB.TextBox txtAD 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   1
         Left            =   900
         MaxLength       =   1
         TabIndex        =   13
         Top             =   450
         Width           =   240
      End
      Begin VB.TextBox txtAppNew 
         Height          =   270
         Index           =   5
         Left            =   -73875
         MaxLength       =   9
         TabIndex        =   41
         Top             =   450
         Width           =   1092
      End
      Begin VB.TextBox txtAppNew 
         Height          =   270
         Index           =   4
         Left            =   -73875
         MaxLength       =   9
         TabIndex        =   37
         Top             =   450
         Width           =   1092
      End
      Begin VB.TextBox txtAppNew 
         Height          =   270
         Index           =   3
         Left            =   -73875
         MaxLength       =   9
         TabIndex        =   31
         Top             =   450
         Width           =   1092
      End
      Begin VB.TextBox txtAppNew 
         Height          =   270
         Index           =   2
         Left            =   -73875
         MaxLength       =   9
         TabIndex        =   26
         Top             =   450
         Width           =   1092
      End
      Begin VB.TextBox txtAppNew 
         Height          =   270
         Index           =   1
         Left            =   1125
         MaxLength       =   9
         TabIndex        =   14
         Top             =   450
         Width           =   1092
      End
      Begin MSForms.TextBox txtAppAddr 
         Height          =   300
         Index           =   5
         Left            =   -74100
         TabIndex        =   42
         Top             =   750
         Width           =   7215
         VariousPropertyBits=   671107099
         Size            =   "12726;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppAddr 
         Height          =   300
         Index           =   4
         Left            =   -74100
         TabIndex        =   38
         Top             =   750
         Width           =   7215
         VariousPropertyBits=   671107099
         Size            =   "12726;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppAddr 
         Height          =   300
         Index           =   3
         Left            =   -74100
         TabIndex        =   32
         Top             =   750
         Width           =   7215
         VariousPropertyBits=   671107099
         Size            =   "12726;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppAddr 
         Height          =   300
         Index           =   2
         Left            =   -74100
         TabIndex        =   27
         Top             =   750
         Width           =   7215
         VariousPropertyBits=   671107099
         Size            =   "12726;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppAddr 
         Height          =   300
         Index           =   1
         Left            =   900
         TabIndex        =   15
         Top             =   750
         Width           =   7215
         VariousPropertyBits=   671107099
         Size            =   "12726;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblAppName 
         Height          =   255
         Index           =   5
         Left            =   -72705
         TabIndex        =   77
         Top             =   450
         Width           =   5385
         VariousPropertyBits=   27
         Caption         =   "lblAppName"
         Size            =   "9499;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblAppName 
         Height          =   255
         Index           =   4
         Left            =   -72705
         TabIndex        =   76
         Top             =   450
         Width           =   5250
         VariousPropertyBits=   27
         Caption         =   "lblAppName"
         Size            =   "9260;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblAppName 
         Height          =   255
         Index           =   3
         Left            =   -72705
         TabIndex        =   75
         Top             =   450
         Width           =   5520
         VariousPropertyBits=   27
         Caption         =   "lblAppName"
         Size            =   "9737;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblAppName 
         Height          =   255
         Index           =   2
         Left            =   -72705
         TabIndex        =   74
         Top             =   450
         Width           =   5520
         VariousPropertyBits=   27
         Caption         =   "lblAppName"
         Size            =   "9737;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "編號："
         Height          =   180
         Index           =   4
         Left            =   -74685
         TabIndex        =   73
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "編號："
         Height          =   180
         Index           =   3
         Left            =   -74685
         TabIndex        =   72
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "編號："
         Height          =   180
         Index           =   2
         Left            =   -74685
         TabIndex        =   71
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "地址："
         Height          =   180
         Index           =   6
         Left            =   -74685
         TabIndex        =   70
         Top             =   810
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "地址："
         Height          =   180
         Index           =   5
         Left            =   -74685
         TabIndex        =   69
         Top             =   810
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "地址："
         Height          =   180
         Index           =   4
         Left            =   -74685
         TabIndex        =   68
         Top             =   810
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "地址："
         Height          =   180
         Index           =   1
         Left            =   -74685
         TabIndex        =   67
         Top             =   810
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "編號："
         Height          =   180
         Index           =   1
         Left            =   -74685
         TabIndex        =   66
         Top             =   480
         Width           =   540
      End
      Begin MSForms.Label lblAppName 
         Height          =   255
         Index           =   1
         Left            =   2295
         TabIndex        =   65
         Top             =   450
         Width           =   5565
         VariousPropertyBits=   27
         Caption         =   "lblAppName"
         Size            =   "9816;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "編號："
         Height          =   180
         Index           =   0
         Left            =   315
         TabIndex        =   64
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "地址："
         Height          =   180
         Index           =   3
         Left            =   315
         TabIndex        =   63
         Top             =   810
         Width           =   540
      End
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "小個體"
      Height          =   255
      Index           =   1
      Left            =   1230
      TabIndex        =   17
      Top             =   4560
      Width           =   1275
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "大個體"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   1065
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   5040
      TabIndex        =   1
      Top             =   1545
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7644
      TabIndex        =   24
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5592
      TabIndex        =   22
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6420
      TabIndex        =   23
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   405
      Index           =   3
      Left            =   4368
      TabIndex        =   21
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   555
      Index           =   8
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   8250
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14552;979"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   10
      Left            =   6120
      TabIndex        =   5
      Top             =   2160
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   9
      Left            =   6120
      TabIndex        =   3
      Top             =   1860
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   4
      Left            =   720
      TabIndex        =   6
      Top             =   2460
      Width           =   1695
      VariousPropertyBits=   671107099
      Size            =   "2990;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   1860
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   5
      Left            =   5760
      TabIndex        =   7
      Top             =   2460
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
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
      Left            =   2610
      TabIndex        =   81
      Top             =   2835
      Width           =   900
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
      Left            =   1920
      TabIndex        =   9
      Tag             =   "Y"
      Top             =   2730
      Width           =   255
   End
   Begin VB.Label lblChkRltDate 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   120
      TabIndex        =   79
      Top             =   2805
      Width           =   765
   End
   Begin VB.Label Label9 
      Caption         =   "IDS揭露項目："
      Height          =   180
      Left            =   4530
      TabIndex        =   78
      Top             =   4380
      Width           =   1170
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "（美、加、法國案）"
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   61
      Top             =   4350
      Width           =   1620
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "是否修改通知函內容：        （Y：Word）"
      Height          =   180
      Left            =   4320
      TabIndex        =   60
      Top             =   2205
      Width           =   3225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "是否修改指示信內容：        （Y:Word）"
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   59
      Top             =   1905
      Width           =   3090
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "進度備註："
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   58
      Top             =   4845
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "隨函附                                         資料"
      Height          =   180
      Left            =   120
      TabIndex        =   57
      Top             =   2505
      Width           =   3405
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "是否列印通知函：        （N：不印）"
      Height          =   180
      Left            =   120
      TabIndex        =   56
      Top             =   2205
      Width           =   2820
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Left            =   120
      TabIndex        =   55
      Top             =   1605
      Width           =   720
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   4320
      TabIndex        =   54
      Top             =   1605
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "是否列印指示信：        （N:不印）"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   53
      Top             =   1905
      Width           =   2685
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "有無前案資料：            （Y/N）"
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   52
      Top             =   2505
      Width           =   2490
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   6300
      TabIndex        =   51
      Top             =   1560
      Width           =   2175
      VariousPropertyBits=   27
      Size            =   "3836;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   5880
      TabIndex        =   44
      Top             =   570
      Width           =   2535
   End
   Begin MSForms.Label lblSalesName 
      Height          =   255
      Left            =   6000
      TabIndex        =   43
      Top             =   900
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   39
      Top             =   1230
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   35
      Top             =   900
      Width           =   615
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   34
      Top             =   570
      Width           =   495
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   33
      Top             =   1230
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   29
      Top             =   900
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   28
      Top             =   570
      Width           =   3135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   50
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   49
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   48
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   47
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      Height          =   180
      Left            =   120
      TabIndex        =   46
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   45
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label lblCaseFees 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   1965
      TabIndex        =   80
      Top             =   2790
      Width           =   255
   End
End
Attribute VB_Name = "frm050102_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/6 改成Form2.0 (txtCaseField,lblSalesName...)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
'92.1.18 add by sonia
Dim old_Entity As String   '原大小個體
Dim new_Entity As String   '新大小個體
'Add By Cheng 2003/09/16
Dim strCountry As String '存放EPC指定國家
'Add by Morgan 2006/3/23
Dim m_str222MailCP14 As String '告建議性處分承辦人
Dim m_str222MailCP09 As String '告建議性處分收文號
Dim m_bolActive As Boolean 'Active事件是否已觸發
Dim skMail() As SeekMails     '2010/1/22 ADD BY SONIA
'Add By Sindy 2018/1/8
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2018/1/8 END
Dim m_strAF01 As String, m_strLD18 As String 'Added by Morgan 2018/8/22
Dim m_bolEngLetter As Boolean, m_strSubject As String 'Added by Morgan 2018/9/4
Dim m_strCP81 As String, m_strJpMemo As String 'Added by Morgan 2019/4/23
Dim m_strNoDisc(5) As String 'Added by Morgan 2019/6/19 是否設定不可減免
Dim strCP09List As String  'Added by Morgan 2025/8/4 子案總收文號

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(1 To 5) As String, strTmp As String, intStep As Integer, i As Integer
   EndLetter ET01, cp(9), ET03, strUserNum
   intStep = 1
   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
      "','隨函附資料','" & txtCaseField(4) & "')"
   intStep = intStep + 1
   '92.1.18 add by sonia
   'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
   If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
      If OptChoose(0).Value = True Then
         strTmp = "(Large Entity)"
      ElseIf OptChoose(1).Value = True Then
         strTmp = "(Small Entity)"
      'Added by Morgan 2013/5/7
      ElseIf OptChoose(2).Value = True Then
         strTmp = "(Micro Entity)"
      'end 2013/5/7
      End If
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','大小個體','" & strTmp & "')"
      intStep = intStep + 1
   End If
   '92.1.18 end
   
   'Add by Morgan 2011/4/14
   'EPC實審416發文若有勾選一併提指定費時要帶入信函內
   If field(9) = "221" And cp(10) = "416" Then
      If chkChoose(2).Value = vbChecked Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','同時辦理事項','及繳交指定費')"
         intStep = intStep + 1
      End If
   
   ElseIf field(9) = "018" And cp(10) = "416" Then
      If chkChoose(1).Value = vbChecked Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','同時辦理事項','及提供前案資料')"
         intStep = intStep + 1
      End If
   End If
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(intStep - 1, strTxt) Then
   If Not ClsLawExecSQL(intStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)

   Dim stLetter As String 'Add by Morgan 2004/9/27
   Dim i As Integer, strTmp As String
   'Add By Cheng 2002/07/30
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim ii As Integer, arrAF01() As String 'Added by Morgan 2025/8/4

   Select Case Index
      Case 0, 3 '確定, 同時發文
         'Add by Morgan 2009/6/1
         If PUB_ChkFileNP(cp(9)) Then
            MsgBox "下一程序已有提申或收達期限，不可發文！"
            Exit Sub
         End If
         
         '92.1.18 add by sonia
         'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
         If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
            'Modified by Morgan 2013/5/7 +微個體
            If Not OptChoose(0).Value And Not OptChoose(1).Value And Not OptChoose(2).Value Then
               'Modified by Morgan 2023/3/24
               If OptChoose(2).Enabled = True Then
                  MsgBox "請選擇" & OptChoose(0).Caption & "、" & OptChoose(1).Caption & "或" & OptChoose(2).Caption & "資料 !", vbCritical
               Else
                  MsgBox "請選擇" & OptChoose(0).Caption & "或" & OptChoose(1).Caption & "資料 !", vbCritical
               End If
               Exit Sub
            End If
         End If
         '92.1.18 end
         
         'Add by Morgan 2007/9/3
         If cp(10) = "214" Then
            strExc(1) = "": strExc(2) = ""
            For intI = 0 To lstItems.ListCount - 1
               If lstItems.Selected(intI) = True Then
                  If strExc(1) <> "" Then
                     strExc(2) = strExc(2) & "、" & Left(strExc(1), Len(strExc(1)) - 3)
                  End If
                  strExc(1) = lstItems.List(intI)
               End If
            Next
            If strExc(1) = "" Then
               MsgBox "IDS發文必須勾選揭露項目！", vbExclamation
               Exit Sub
            End If
            If strExc(2) = "" Then
               strExc(2) = "揭露項目：" & Replace(strExc(1), " ", "")
            Else
               If strExc(1) = "其他資料" Then
                  strExc(2) = "揭露項目：" & Mid(strExc(2), 2) & "前案及" & strExc(1)
               Else
                  strExc(2) = "揭露項目：" & Mid(strExc(2), 2) & "及" & Left(strExc(1), Len(strExc(1)) - 3) & "前案資料"
               End If
            End If
            txtCaseField(8) = strExc(2) & ";" & txtCaseField(8)
         End If
         'end 2007/9/3
         
         Screen.MousePointer = vbHourglass
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
         
         If SaveDatabase Then
         
            'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
            PUB_CheckEMail cp(44), cp(116)
            PUB_CheckEMail field(75), field(144)
            If field(145) <> "" Then
               PUB_CheckEMail field(75), field(145)
            End If
            'end 2008/2/20
               
            'Add by Morgan 2006/3/23 告建議性處分上齊備日通知承辦人
            If m_str222MailCP14 <> Empty Then
               'Modified by Morgan 2016/3/3 +126 期末拋棄
               'Modified by Lydia 2016/08/26 +438 再考量試行計畫(AFCP2.0)
               'PUB_SendMail strUserNum, m_str222MailCP14, m_str222MailCP09, field(1) & "-" & field(2) & "-" & field(3) & "-" & field(4) & "文件齊備通知！", "", IIf(cp(10) = "126", "期末拋棄", "答辯") & "發文逾准駁通知日(" & Format(TransDate(field(20), 1), "###/##/##") & ")+5個月，系統自動上文件齊備日！"
               PUB_SendMail strUserNum, m_str222MailCP14, m_str222MailCP09, field(1) & "-" & field(2) & "-" & field(3) & "-" & field(4) & "文件齊備通知！", "", IIf(cp(10) = "126", "期末拋棄", IIf(cp(10) = "438", "再考量試行計畫", "答辯")) & "發文逾准駁通知日(" & Format(TransDate(field(20), 1), "###/##/##") & ")+5個月，系統自動上文件齊備日！"
            End If
            
            BatchMail '2010/1/21 ADD BY SONIA 發E-Mail給承辦人
            
            '指示信
            If txtCaseField(2) <> "N" Then
               'Added by Morgan 2025/8/4
               If strCP09List <> "" Then
                  arrAF01() = Split(strCP09List, ",")
                  For ii = 0 To UBound(arrAF01)
                     NowPrint arrAF01(ii), "01", "30", False, strUserNum, 0, , , , , , , , , , , , arrAF01(ii)
                  Next
                  MsgBox "子案指示信請至待處理區作業！", vbInformation
               Else
               'end 2025/8/4
                  
                  '93.3.6 modify by sonia  日本實審指示信不同
                  'StartLetter "01", "30"
                  ''Modify By Cheng 2002/07/30
   '                 ' NowPrint cp(9), "01", "00", False, strUserNum, 0
                  'NowPrint cp(9), "01", "30", IIf(Me.txtCaseField(9).Text = "Y", True, False), strUserNum, 0
                  strTmp = "30"
                  If field(9) = "011" And (cp(10) = 實體審查 Or cp(10) = "427") Then
                     strTmp = "31"
                  End If
                  'Add by Morgan 2005/3/17 加日本讓渡
                  If field(9) = "011" And cp(10) = "701" Then
                     strTmp = "31"
                  End If
   
                  StartLetter "01", strTmp
                  'Modify By Cheng 2002/07/30
   '                  NowPrint cp(9), "01", "00", False, strUserNum, 0
   
                  'Modify by Morgan 2004/9/27
                  '實審加印傳真封面
                  'NowPrint cp(9), "01", strTmp, IIf(Me.txtCaseField(9).Text = "Y", True, False), strUserNum, 0
                  'Removed by Morgan 2018/10/22 取消傳真封面--慧汶
                  'If cp(10) = 實體審查 Then
                  '   If txtCaseField(9).Text = "Y" Then
                  '      NowPrint cp(9), "01", "89", False, strUserNum, , , True, stLetter, , , , , , , , , m_strAF01
                  '   Else
                  '      NowPrint cp(9), "01", "89", False, strUserNum, , , , , , , , , , , , , m_strAF01
                  '   End If
                  '   If m_strAF01 <> "" Then Sleep 1000 '等1秒以確保letterdemand不會發生dupe錯誤 Added by Morgan 2018/8/20
                  'End If
                  'end 2018/10/22
                  NowPrint cp(9), "01", strTmp, IIf(txtCaseField(9).Text = "Y", True, False), strUserNum, 0, stLetter, , , , , , , , , , , m_strAF01
                  '2004/9/27 end
               
                  'Added by Morgan 2018/8/22 CFP電子化
                  If txtCaseField(9).Text = "Y" And m_strAF01 <> "" Then
                     frm1105_1.m_RecNo = m_strAF01
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".DATA.PDF"
                     frm1105_1.Show
                     If txtCaseField(10).Text = "Y" Then
                        MsgBox "指示信編輯中，客戶函請至定稿維護修改！", vbExclamation
                        txtCaseField(10).Text = ""
                     End If
                  End If
                  'end 2018/8/22
                  
               End If 'Added by Morgan 2025/8/4
               
'Removed by Morgan 2012/3/7 不必詢問,需要時程序自行列印--甄妮
'               'Add By Cheng 2002/07/30
'               StrSQLa = "Select FA04,FA05||' '||FA63||' '||FA64||' '||FA65,FA06,FA01||FA02 From CASEPROGRESS, STAFF, FAGENT WHERE (CP10='416' OR CP10='427') AND CP14=ST01(+) AND ST03='P12' AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP09='" & cp(9) & "'"
'               rsA.CursorLocation = adUseClient
'               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'               If rsA.RecordCount > 0 Then
'                 If MsgBox("代理人名稱(中)：" & rsA.Fields(0).Value & Chr(10) & Chr(13) & _
'                           "　　　　　(英)：" & rsA.Fields(1).Value & Chr(10) & Chr(13) & _
'                           "　　　　　(日)：" & rsA.Fields(2).Value & Chr(10) & Chr(13) & Chr(10) & Chr(13) & _
'                           "是否列印代理人小信封？", vbExclamation + vbYesNo) = vbYes Then
'                    '列印地址條
'                    'Modify by Morgan 2006/10/17 改Call公用函數
'                    'PrintCase "" & rsA.Fields(3).Value
'                    PUB_PrintCase "" & rsA.Fields(3).Value
'                 End If
'               End If
'               If rsA.State <> adStateClosed Then rsA.Close
'               Set rsA = Nothing

            'Added by Morgan 2018/9/6+有工程師的指示信
            ElseIf m_bolEngLetter Then
               PUB_SendOrderLetterP m_strAF01, m_strSubject
            'end 2018/9/6
            
            End If
            
            '通知函
            If txtCaseField(3) <> "N" Then
               strTmp = "00"
               If cp(10) = 提供前案資料 And txtCaseField(5) = "N" Then
                  strTmp = "01"   ' 無前案
               End If
               'Added by Lydia 2016/08/26 再考量試行計畫(AFCP2.0)
               If cp(10) = "438" Then
                  strTmp = "32"
               End If
                              
               StartLetter "01", strTmp
               'Modify By Cheng 2002/07/30
'                  NowPrint cp(9), "01", strTmp, False, strUserNum, 0
               NowPrint cp(9), "01", strTmp, IIf(Me.txtCaseField(10).Text = "Y", True, False), strUserNum, 0, , , , , , , , , , , , m_strLD18
               
               'Added by Morgan 2018/8/22 CFP電子化
               If txtCaseField(10).Text = "Y" And m_strLD18 <> "" Then
                  frm1105_1.m_RecNo = m_strLD18
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".CUS.PDF"
                  frm1105_1.Show
               End If
               'end 2018/8/22
            End If
            bolLeave = True
            intLeaveKind = 1
            'Add By Cheng 2002/04/30
            '若有未發文資料顯示警告
            PUB_GetCPunIssueDatas "" & Me.lblCaseField(1).Caption
             'Add By Cheng 2003/11/27
             ' 發文回前畫面時
             Select Case Index
                Case 0:
                   ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
                  'Add By Sindy 2013/5/28
                  If frm050102_1.bolIsEMPFlow = True Then
                     intLeaveKind = 0
                     'Unload frm050102_1
                     frm090202_4.Show
                     frm090202_4.QueryData
                  '2013/5/28 End
                  'Add By Sindy 2018/1/8
                  ElseIf Me.m_strIR01 <> "" Then
                     intLeaveKind = 0
                     'Modify By Sindy 2022/5/20
                     'frm04010519.GoNext
                     Forms(0).Tmpfrm04010519.GoNext
                     Set Forms(0).Tmpfrm04010519 = Nothing
                     '2022/5/20 END
                  '2018/1/8 END
                  Else
                     frm050102_1.Show
                     frm050102_1.Clear
                  End If
                Case 3:
                     '若尚有未發文資料
                     If PUB_ChkUnissueDatas(Me.lblCaseField(1).Caption) = True Then
                         ' 90.07.12 modify by louis (回發文主畫面並重新查詢)
                        'Add By Sindy 2013/5/28
                        If frm050102_1.bolIsEMPFlow = True Then
                           frm090202_4.QueryData
                        'End If
                        '2013/5/28 End
                        'Add By Sindy 2018/1/8
                        ElseIf Me.m_strIR01 <> "" Then
                           'intLeaveKind = 0
                           'Modify By Sindy 2022/5/20
                           'frm04010519.GoNext
                           Forms(0).Tmpfrm04010519.GoNext
                           Set Forms(0).Tmpfrm04010519 = Nothing
                           '2022/5/20 END
                        '2018/1/8 END
                        End If
                        frm050102_1.Show
                        frm050102_1.ReQuery
                     '若無未發文資料
                     Else
                         ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
                        'Add By Sindy 2013/5/28
                        If frm050102_1.bolIsEMPFlow = True Then
                           intLeaveKind = 0
                           'Unload frm050102_1
                           frm090202_4.Show
                           frm090202_4.QueryData
                        '2013/5/28 End
                        'Add By Sindy 2018/1/8
                        ElseIf Me.m_strIR01 <> "" Then
                           intLeaveKind = 0
                           'Modify By Sindy 2022/5/20
                           'frm04010519.GoNext
                           Forms(0).Tmpfrm04010519.GoNext
                           Set Forms(0).Tmpfrm04010519 = Nothing
                           '2022/5/20 END
                        '2018/1/8 END
                        Else
                           frm050102_1.Show
                           frm050102_1.Clear
                        End If
                     End If
             End Select
             'End
            Unload Me

         '911202 nick
         Else
             MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
         End If
         Screen.MousePointer = vbDefault
      Case 1, 2
         'Add By Sindy 2013/5/28
         If frm050102_1.bolIsEMPFlow = True Then
            intLeaveKind = 0
            'Unload frm050102_1
            frm090202_4.Show
            frm090202_4.QueryData
         '2013/5/28 End
         'Add By Sindy 2018/1/8
         ElseIf Me.m_strIR01 <> "" Then
            intLeaveKind = 0
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         '2018/1/8 END
         Else
            If Index = 2 Then
               intLeaveKind = 0
            Else
               intLeaveKind = 1
            End If
         End If
         Unload Me
   End Select
    'Modify By Cheng 2003/11/27
    '本段程式往上移
'   ' 發文回前畫面時
'   Select Case Index
'      Case 0:
'         ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
'         frm050102_1.Clear
'      Case 3:
'         ' 90.07.12 modify by louis (回發文主畫面並重新查詢)
'         frm050102_1.ReQuery
'   End Select
    'End
End Sub

Private Function SaveDatabase() As Boolean
Dim strTxt(1 To 10) As String, iStep As Integer, iIdx As Integer
'Add By Cheng 2003/03/04
Dim StrSQLa As String
Dim strPromoteDate As String '2010/1/14 add by sonia
Dim bolLstInnerCase As Boolean '2010/1/21 是否最後一個國內案
Dim ii As Integer
Dim strLetterJudge As String '指示信判發人 Added by Morgan 2018/8/22

'911106 nick transation
SaveDatabase = True
On Error GoTo CheckingErr
cnnConnection.BeginTrans
   
   'Add by Morgan 2005/11/22 更新讓與相關資料
   'Modify by Morgan 2007/7/27 加"繼承"
   'If cp(10) = 讓與 Then
   If cp(10) = 讓與 Or cp(10) = 繼承 Then
   
      'Added by Morgan 2013/6/5
      '設定受讓人減免身分
      For ii = 1 To 5
         If txtAD(ii).Enabled = True Then
            '身分有變更才要做
            If txtAD(ii).Tag <> txtAD(ii).Text Then
               'Modified by Morgan 2020/3/26
               'strSql = PUB_GetADSQL(field(25 + ii), field(9), txtAD(ii).Text)
               strSql = PUB_GetADSQL(Me.txtAppNew(ii), field(9), txtAD(ii).Text)
               'end 2020/3/26
               cnnConnection.Execute strSql
            End If
         End If
      Next
      'end 2013/6/5
   
      field(157) = "N" 'Add by Morgan 2010/6/17
      
      '若原申請人與讓與申請人不同時, 才更新讓與人(避免讓與人欄位被蓋掉)
      '讓與申請人1
      cp(56) = ChangeCustomerL(txtAppNew(1))
      If ChangeCustomerL(field(26)) <> cp(56) Then
         '讓與人1
         cp(55) = ChangeCustomerL(field(26))
         '新申請人1
         field(26) = cp(56)
         '新申請人1地址
         'edit by nickc 2007/02/02 不用 dll 了
         'Call objPublicData.GetCustomerNameAndAddress(cp(56), strExc(0), strExc(1), strExc(2), strExc(3))
         Call ClsPDGetCustomerNameAndAddress(cp(56), strExc(0), strExc(1), strExc(2), strExc(3))
         cp(56) = ChangeCustomerL(cp(56)) 'Add by Morgan 2006/10/30
         field(31) = Trim(strExc(1))
         field(36) = Trim(strExc(2))
         field(41) = Trim(strExc(3))
      'Add by Morgan 2006/3/23
      '讓與人與受讓人相同時判斷CP有資料時才不更新
      Else
         '讓與人1
         If cp(55) = "" Then cp(55) = ChangeCustomerL(field(26))
      '2006/3/23 end
      End If
      '受讓人2~5
      For iIdx = 2 To 5
         '讓與申請人
         cp(87 + iIdx) = ChangeCustomerL(txtAppNew(iIdx))
         '讓與人(原申請人與讓與申請人不同時, 才更新)
         If ChangeCustomerL(field(25 + iIdx)) <> cp(87 + iIdx) Then
            '讓與人
            cp(91 + iIdx) = ChangeCustomerL(field(25 + iIdx))
            '新申請人
            field(25 + iIdx) = cp(87 + iIdx)
            '新申請人地址
            If cp(87 + iIdx) <> "" Then
               'edit by nickc 2007/02/02 不用 dll 了
               'Call objPublicData.GetCustomerNameAndAddress(cp(87 + iIdx), strExc(0), strExc(1), strExc(2), strExc(3))
               Call ClsPDGetCustomerNameAndAddress(cp(87 + iIdx), strExc(0), strExc(1), strExc(2), strExc(3))
               cp(87 + iIdx) = ChangeCustomerL(cp(87 + iIdx)) 'Add by Morgan 2006/10/30
               field(30 + iIdx) = Trim(strExc(1))
               field(35 + iIdx) = Trim(strExc(2))
               field(40 + iIdx) = Trim(strExc(3))
            Else
               field(30 + iIdx) = Empty
               field(35 + iIdx) = Empty
               field(40 + iIdx) = Empty
            End If
         'Add by Morgan 2006/3/23
         '讓與人與受讓人相同時判斷CP有資料時才不更新
         Else
            '讓與人
            If cp(91 + iIdx) = "" Then cp(91 + iIdx) = ChangeCustomerL(field(25 + iIdx))
         End If
      Next
   End If
   '2005/11/22 end

   cp(27) = txtCaseField(0)

   'Modify by Morgan 2008/2/20
   'cp(44) = Combo1
   intI = InStr(Combo1, "-")
   If intI > 0 Then
      cp(44) = Left(Combo1, intI - 1)
      cp(116) = Mid(Combo1, intI + 1)
   Else
      cp(44) = Combo1
      cp(116) = ""
   End If
   'end 2008/2/20
   cp(44) = ChangeCustomerL(cp(44))
   
   'Remove by Morgan 2005/11/22 移到上面程式做
'    'Modify By Cheng 2003/02/20
'    '受讓人(讓與申請人)
''   cp(55) = txtCaseField(6)
'   cp(56) = ChangeCustomerL(txtCaseField(6))
'    'Add By Cheng 2003/03/04
'    If Me.txtCaseField(6).Text <> "" Then
'        '若受讓人與原申請人不同時
'        If ChangeCustomerL(Me.txtCaseField(6).Text) <> ChangeCustomerL(field(26)) Then
'            '讓與人
'            cp(55) = field(26)
'        End If
'    End If
    '2005/11/22 end
    
   '92.1.18 modify by sonia
   'cp(64) = txtCaseField(8)
   'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
   'Modified by Morgan 2023/3/24 條件同 ReadAllData
   'If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
   If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
      If strSrvDate(1) >= PA179啟用日 Then
         If OptChoose(0).Value = True Then
            new_Entity = OptChoose(0).Caption
         ElseIf OptChoose(1).Value = True Then
            new_Entity = OptChoose(1).Caption
         ElseIf OptChoose(2).Value = True Then
            new_Entity = OptChoose(2).Caption
         End If
      Else
   'end 2023/3/24
   
         If OptChoose(0).Value = True Then
            new_Entity = "大個體"
         ElseIf OptChoose(1).Value = True Then
            new_Entity = "小個體"
         'Added by Morgan 2013/3/20
         ElseIf OptChoose(2).Value = True Then
            new_Entity = "微個體"
         'end 2013/3/20
         End If
         
      End If 'Added by Morgan 2023/3/24
      
      If old_Entity <> new_Entity And old_Entity <> "" Then  '改大小個體時
         If txtCaseField(8) = "" Then
            cp(64) = "原大小個體為" & old_Entity
         Else
            cp(64) = "原大小個體為" & old_Entity & "，" & Me.txtCaseField(8).Text
         End If
      Else
         cp(64) = txtCaseField(8)
      End If
   Else
      cp(64) = txtCaseField(8)
   End If
   '92.1.18 end
   
   If m_strJpMemo <> "" Then cp(64) = m_strJpMemo & ";" & cp(64) 'Added by Morgan 2019/4/29 日本實審可減免備註加減免身分
   
   cp(81) = m_strCP81 'Added by Morgan 2019/4/23
   cp(113) = txtCP113 'Added by Lydia 2021/05/25 工作時數
   
   'Modified by Morgan 2012/2/15 改呼叫共用函數
   'strExc(0) = "select cp45 from caseprogress where cp01=" + CNULL(cp(1)) + _
   '   " and cp02=" + CNULL(cp(2)) + " and cp03=" + CNULL(cp(3)) + _
   '   " and cp04=" + CNULL(cp(4)) + " and cp44=" + CNULL(cp(44)) + " order by cp27 desc"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   'cp(45) = ""
   'If intI = 1 And Not IsNull(RsTemp.Fields("CP45")) Then cp(45) = RsTemp.Fields("CP45")
   If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   'end 2012/2/15
   
   strTxt(1) = GetCPSQL(cp())
   
   '911106 nick transation
   'SaveDatabase = objLawDll.ExecSQL(1, strTxt)
   cnnConnection.Execute strTxt(1)
   
   '92.1.18 add by sonia 改大小個體時
   'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
   'Modified by Morgan 2023/3/24 條件同 ReadAllData
   'If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
   If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
      If strSrvDate(1) >= PA179啟用日 Then
         If OptChoose(0).Value = True Then
            field(179) = "1"
         ElseIf OptChoose(1).Value = True Then
            field(179) = "2"
         ElseIf OptChoose(2).Value = True Then
            field(179) = "3"
         End If
      Else
   'end 2023/3/24
   
         If old_Entity <> new_Entity Then
            If InStr(1, field(91), old_Entity, 1) > 0 Then
               field(91) = Replace(field(91), old_Entity, new_Entity, InStr(1, field(91), old_Entity, 1), , 1)
            Else
               If field(91) = "" Then
                  field(91) = new_Entity
               Else
                  field(91) = new_Entity & "，" & field(91)
               End If
            End If
         End If
         
      End If 'Added by Morgan 2023/3/24
   End If
   strTxt(2) = GetPASQL(field())
   
   cnnConnection.Execute strTxt(2)
   '92.1.18 end
   
   '2010/3/22 ADD BY SONIA 讓與發文時,更新下一程序非專業部掌控之案件性質(未續辦)的智權人員
   '2013/4/18 MODIFY BY SONIA 改智權人員
   'pub_ChgSalesTargetIsNp cp(1), cp(2), cp(3), cp(4), cp(13) CFP-025472
   'modify by sonia 2019/12/23 從上面移下來,要基本檔申請人已改才有用
   pub_ChgSalesTargetIsNp cp(1), cp(2), cp(3), cp(4), PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
   
   '2009/8/21 add by sonia EPC母案讓與子案也要改申請人
   If field(9) = EPC指定國家 And cp(10) = 讓與 Then
      strTxt(3) = "UPDATE PATENT SET PA26=(SELECT PA26 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA27=(SELECT PA27 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA28=(SELECT PA28 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA29=(SELECT PA29 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA30=(SELECT PA30 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA31=(SELECT PA31 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA32=(SELECT PA32 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA33=(SELECT PA33 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA34=(SELECT PA34 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA35=(SELECT PA35 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA36=(SELECT PA36 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA37=(SELECT PA37 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA38=(SELECT PA38 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA39=(SELECT PA39 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA40=(SELECT PA40 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA41=(SELECT PA41 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA42=(SELECT PA42 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA43=(SELECT PA43 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA44=(SELECT PA44 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA45=(SELECT PA45 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'), " & _
                                    "PA91=(SELECT PA91 FROM PATENT WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "') " & _
                  "WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04<>'00'"
      cnnConnection.Execute strTxt(3)
   End If
   '2009/8/21 END
   
   'Remove by Morgan 2005/11/22 移到上面程式做
'    'Add By Cheng 2003/03/04
'    '更新基本檔的申請人
'    If Me.txtCaseField(6).Text <> "" Then
'        '若受讓人與原申請人不同時
'        '92.8.13 modify by sonia
'        'If ChangeCustomerL(Me.txtCaseField(6).Text) <> ChangeCustomerL(field(26)) Then
'
'
'        If cp(10) = "701" Then
'        '92.8.13 END
'            'Modiry By Cheng 2003/03/07
'            '更新基本檔申請人及相關資料
'            '92.8.13 modify by sonia
'            'strSQLA = "Update Patent Set PA26='" & ChangeCustomerL(Me.txtCaseField(6).Text) & "' " & _
'            '                " ,PA31='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.txtCaseField(6).Text), "1")) & "' " & _
'            '                " ,PA36='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.txtCaseField(6).Text), "2")) & "' " & _
'            '                " ,PA41='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.txtCaseField(6).Text), "3")) & "' " & _
'            '                " Where " & ChgPatent(field(1) & field(2) & field(3) & field(4))
'            StrSQLa = "Update Patent Set PA26='" & ChangeCustomerL(Me.txtCaseField(6).Text) & "' " & _
'                            " ,PA31='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.txtCaseField(6).Text), "1")) & "' " & _
'                            " ,PA36='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.txtCaseField(6).Text), "2")) & "' " & _
'                            " ,PA41='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.txtCaseField(6).Text), "3")) & "' " & _
'                            " ,PA27=NULL,PA28=NULL,PA29=NULL,PA30=NULL,PA32=NULL,PA33=NULL,PA34=NULL,PA35=NULL,PA37=NULL" & _
'                            " ,PA38=NULL,PA39=NULL,PA40=NULL,PA42=NULL,PA43=NULL,PA44=NULL,PA45=NULL" & _
'                            " Where " & ChgPatent(field(1) & field(2) & field(3) & field(4))
'            '92.8.13 end
'            cnnConnection.Execute StrSQLa
'        End If
'    End If

    'Add By Cheng 2003/09/16
    '若有ECP指定國家, 則新增案件進度檔資料
    If field(9) = EPC指定國家 And strCountry <> "" Then
      'Added by Morgan 2025/7/24
      'EPC進入國家階段的讓渡子案也要掛承辦人、產生指示信(抓子案指定國註冊費代理人)並管制收達、提申及催審--玫音
      If cp(10) = "701" And field(21) <> "" Then '以發證日判斷是否進入國家階段(同期限通知管制表)
         Dim strAgentList As String, pa04 As String
         Dim arrPA04() As String, ArrCP09() As String, varTmp1() As String, arrCP(4) As String
         Dim strSubject As String

         varTmp1 = Split(strCountry, ",")
         ReDim arrPA04(UBound(varTmp1))

         For ii = 0 To UBound(varTmp1)
            strExc(0) = "select pa04,cp44,cp116 from patent,caseprogress" & _
               " where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & cp(3) & "' and pa09='" & varTmp1(ii) & "'" & _
               " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10(+) in ('224','249')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               arrPA04(ii) = RsTemp("pa04")
               strAgentList = strAgentList & "," & RsTemp("cp44")
               If Not IsNull(RsTemp("cp116")) Then
                  strAgentList = strAgentList & "-" & RsTemp("cp116")
               End If
            Else
               GoTo CheckingErr
            End If
         Next
         strAgentList = Mid(strAgentList, 2)

         If PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strCountry, strAgentList, strCP09List, cp(10)) Then
            ArrCP09 = Split(strCP09List, ",")
            For ii = 0 To UBound(varTmp1)
               strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), cp(10), Format(varTmp1(ii)))
               strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), arrPA04(ii), cp(10), field(11), , Format(varTmp1(ii)))
               PUB_AddAppForm ArrCP09(ii), True, strLetterJudge, strSubject

               arrCP(1) = cp(1): arrCP(2) = cp(2): arrCP(3) = cp(3): arrCP(4) = arrPA04(ii)

               '催審期限
               If txtChkRltDate <> "" Then
                  PUB_UpdateChkResultDate txtChkRltDate, arrCP, ArrCP09(ii), , , Format(varTmp1(ii))
               End If
               '提申期限
               PUB_SetApplyDate cp(1), cp(2), cp(3), arrPA04(ii), cp(7), ArrCP09(ii), cp(10), txtCaseField(0), Format(varTmp1(ii))
               '收達
               PUB_SetArriveDate ArrCP09(ii)
            Next
         Else
            GoTo CheckingErr
         End If
      Else
      'end 2025/7/24
      
         'Modify by Morgan 2006/12/25
         'If Not objPublicData.SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strCountry) Then
         If Not PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strCountry) Then
             GoTo CheckingErr
         End If
         
      End If 'Added by Morgan 2025/7/24
    End If
   ' 90.12.05 modify by louis 若案件國家收費表存在代理人收達天數則新增一筆收達的下一程序檔
   'Modify by Morgan 2007/4/13 判斷發文日非111111的才要 -->郭雅娟
   'If SaveDatabase = True Then
   If SaveDatabase = True And txtCaseField(0) <> "111111" Then
'Modify by Morgan 2009/11/11 發文收達期限管控改呼叫公用函式
      PUB_SetArriveDate cp(9)
'end 2009/11/11
   End If
   
   'Add by Morgan 2006/3/23 美國答辯發文時,若相關收文為"最終核駁"時自動產生內部收文"告建議性處分"
   'Modified by Morgan 2016/3/3 +126 期末拋棄
   m_str222MailCP14 = Empty: m_str222MailCP09 = Empty
   'Modified by Lydia 2016/08/26 +438 再考量試行計畫(AFCP2.0)
   If field(9) = 美國國家代號 And (cp(10) = 答辯 Or cp(10) = "126" Or cp(10) = "438") Then
      strExc(0) = "select 1 from caseprogress where cp09='" & cp(43) & "' and cp10='1006'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      '法定期限=准駁通知日+6個月,本所期限=法定期限-14,承辦人=粘竺儒84012
      If intI = 1 Then
         '收文號
         strExc(0) = AutoNo("B", 6)
         '法定期限
         strExc(1) = CompDate(1, 6, field(20))
         '本所期限
         strExc(2) = CompDate(2, -14, strExc(1))
         strExc(2) = PUB_GetWorkDay1(strExc(2), True)
         '承辦人
         strExc(3) = "89026"   'modify by sonia 2016/6/13 承辦人=粘竺儒84012改張偉城89026
         strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
            "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43) VALUES " & _
            "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & _
            strSrvDate(1) & "," & strExc(2) & "," & strExc(1) & _
            ",'" & strExc(0) & "','222','90','" & cp(12) & "','" & cp(13) & "'" & _
            ",'" & strExc(3) & "','N','N','N','" & cp(9) & "') "
         cnnConnection.Execute strSql
         '若發文日>=准駁通知日+5個月,文件齊備日上系統日
         strExc(4) = CompDate(1, 5, field(20))
         If TransDate(txtCaseField(0), 2) >= strExc(4) Then
            strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & strExc(0) & "'"
            cnnConnection.Execute strSql
            If PUB_IfSetCP48() Then  'Add by Morgan 2010/10/6
            
               '2010/1/14 add by sonia 更新承辦期限
               strPromoteDate = Pub_GetHandleDay(cp(1), field(9), "222")
               If strPromoteDate <> "" Then
                  strSql = "Update CaseProgress Set CP48=" & CNULL(strPromoteDate) & " Where CP09='" & strExc(0) & "' "
                  cnnConnection.Execute strSql
               End If
               '2010/1/14 end
               
            End If 'Add by Morgna 2010/10/6
            
            m_str222MailCP14 = strExc(3)
            m_str222MailCP09 = strExc(0)
         End If
      End If
   End If
   '2006/3/23 end
   
   '2010/1/20 add by sonia A類修正(203,204)發文時,國外案未發文之新申請程序重新更新齊備日並計算承辦期限並發Mail通知工程師,不管有沒有齊備日,自frm050102_3複製過來
   ReDim skMail(0) As SeekMails
   If Left(cp(9), 1) = "A" And (cp(10) = "203" Or cp(10) = "204") Then
      strExc(0) = "SELECT CM01,CM02,CM03,CM04,PA09,NVL(PA05,NVL(PA06,PA07)) PA05 FROM CASEMAP,PATENT WHERE " & ChgCaseMap(cp(1) & cp(2) & cp(3) & cp(4), 0, 1) & " AND PA01(+)=CM01 AND PA02(+)=CM02 AND PA03(+)=CM03 AND PA04(+)=CM04 AND CM10='0' AND PA57 IS NULL"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
            Do While Not .EOF
               strSql = "Select EP02,CP06,CP14,CP48,EP06,EP09,EP28,EP07,EP08,EP33 From CaseProgress,EngineerProgress WHERE " & ChgCaseprogress("" & .Fields(0) & .Fields(1) & .Fields(2) & .Fields(3)) & _
                        " AND CP10 in (" & CaseMapOut & ") And CP27 Is Null AND  CP57 IS NULL and EP02=CP09 "
               CheckOC
               With adoRecordset
                  .CursorLocation = adUseClient
                  .Open strSql, cnnConnection
                  If .RecordCount > 0 Then
                     '若國外案有其他國內案未發文時不做
                     bolLstInnerCase = True
                     strExc(0) = "select cp01,cp02,cp03,cp04 from casemap,caseprogress" & _
                        " where cm10='0' and cm01='" & RsTemp.Fields("cm01") & "' and cm02='" & RsTemp.Fields("cm02") & "' and cm03='" & RsTemp.Fields("cm03") & "' and cm04='" & RsTemp.Fields("cm04") & "'" & _
                        " and not (cm05='" & cp(1) & "' and cm06='" & cp(2) & "' and cm07='" & cp(3) & "' and cm08='" & cp(4) & "')" & _
                        " and cp01(+)=cm05 and cp02(+)=cm06 and cp03(+)=cm07 and cp04(+)=cm08 AND CP10 IN (" & CaseMapIn & ") and cp27 is null and cp57 is null and rownum<2"
                     intI = 1
                     Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        bolLstInnerCase = False
                     End If
                     If bolLstInnerCase = True Then
                        strSql = "Update ENGINEERPROGRESS Set EP06=" & strSrvDate(1) & ",EP09=NULL,EP07=NULL,EP28=NULL,EP33=NULL,EP08=NULL Where EP02='" & adoRecordset.Fields(0).Value & "' "
                        cnnConnection.Execute strSql, intI
                        ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
                        skMail(UBound(skMail)).fiSender = strUserNum
                        skMail(UBound(skMail)).fiReceiver = adoRecordset.Fields("CP14").Value
                        skMail(UBound(skMail)).fiRecriverNo = ""
                        If "" & adoRecordset.Fields("EP06").Value = "" Then
                           skMail(UBound(skMail)).fiSubject = "因關聯基礎案" & lblCaseField(1) & "的修正發文，上未齊備關聯案的文件齊備日及承辦期限！"
                           skMail(UBound(skMail)).fiContent = "未齊備關聯案：" & vbCrLf & vbCrLf & "本所案號：" & RsTemp("cm01") & "-" & RsTemp("cm02") & IIf(RsTemp("cm03") & RsTemp("cm04") = "000", "", "-" & RsTemp("cm03") & "-" & RsTemp("cm04")) & vbCrLf & "總收文號：" & adoRecordset.Fields(0).Value & vbCrLf & "案件名稱：" & RsTemp("PA05") & vbCrLf & "文件齊備日：" & ChangeTStringToTDateString(strSrvDate(2))
                        Else
                           skMail(UBound(skMail)).fiSubject = "因關聯基礎案" & lblCaseField(1) & "的修正發文，更新未發文關聯案之文件齊備日為系統日並重新計算承辦期限！"
                           skMail(UBound(skMail)).fiContent = "未發文關聯案原承辦進度資料如下：" & vbCrLf & vbCrLf & "本所案號：" & RsTemp("cm01") & "-" & RsTemp("cm02") & IIf(RsTemp("cm03") & RsTemp("cm04") = "000", "", "-" & RsTemp("cm03") & "-" & RsTemp("cm04")) & vbCrLf & "總收文號：" & adoRecordset.Fields(0).Value & vbCrLf & "案件名稱：" & RsTemp("PA05") & vbCrLf & _
                                                              "文件齊備日：" & ChangeWStringToTDateString("" & adoRecordset.Fields("EP06").Value) & vbCrLf & _
                                                              "承辦　期限：" & ChangeWStringToTDateString("" & adoRecordset.Fields("CP48").Value) & vbCrLf & _
                                                              "完　稿　日：" & ChangeWStringToTDateString("" & adoRecordset.Fields("EP09").Value) & vbCrLf & _
                                                              "會　稿　日：" & ChangeWStringToTDateString("" & adoRecordset.Fields("EP07").Value) & vbCrLf & _
                                                              "預定會稿日：" & ChangeWStringToTDateString("" & adoRecordset.Fields("EP28").Value) & vbCrLf & _
                                                              "英文核完日：" & ChangeWStringToTDateString("" & adoRecordset.Fields("EP33").Value) & vbCrLf & _
                                                              "會稿完成日：" & ChangeWStringToTDateString("" & adoRecordset.Fields("EP08").Value)
                        End If
                        
                        If PUB_IfSetCP48() Then  'Add by Morgan 2010/10/6
                       
                           strSql = "Select cp01,cp10,pa09 From CaseProgress, Patent Where CP09='" & adoRecordset.Fields(0).Value & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
                           CheckOC2
                           With adoRecordset1
                              .CursorLocation = adUseClient
                              .Open strSql, cnnConnection
                              If .RecordCount > 0 Then
                                 strPromoteDate = Pub_GetHandleDay(.Fields("cp01"), .Fields("pa09"), .Fields("cp10"), , "" & adoRecordset.Fields(1))
                                 If strPromoteDate <> "" Then
                                    strSql = "Update CaseProgress Set CP48=" & CNULL(strPromoteDate) & " Where CP09='" & adoRecordset.Fields(0).Value & "' "
                                    cnnConnection.Execute strSql, intI
                                 End If
                              End If
                           End With
                           
                        End If 'Add by Morgna 2010/10/6
                     End If
                  End If
               End With
               .MoveNext
            Loop
         End With
      End If
   End If
   '2010/1/20 end
   
   'Add by Morgan 2009/5/14
   '提申管制
   'Modified by Morgan 2015/8/7 改呼叫共用
   PUB_SetApplyDate cp(1), cp(2), cp(3), cp(4), cp(7), cp(9), cp(10), txtCaseField(0), field(9)
   'end 2015/8/7
   
   'Add by Morgan 2009/8/18
   If txtChkRltDate <> "" Then
      'Modify by Morgan 2010/11/8 核准後的實審發文掛本程序的催審期限
      If field(16) = "1" And cp(10) = "416" Then
         'Modified by Lydia 2016/10/13 + pa09
         'PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9)
         PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9), , , field(9)
      Else
         'Modified by Lydia 2016/10/13 + pa09
         PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9), cp(10), cp(43), field(9)
      End If
   End If
   
   'Add By Sindy 2015/8/3 發文時,若工程師各項日期未輸入者,自動更新為發文日
   Call PUB_UpdEmpDate(cp(9), cp(1), cp(10), DBDATE(cp(27)))
   
   'Add by Sindy 2018/1/8
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050102_1"
   End If
   '2018/1/8 END
   
   'Added by Morgan 2018/8/22 CFP電子化
   If strSrvDate(1) >= CFP指示信電子化啟用日 Then
      'Modified by Morgan 2018/9/6 +有工程師的指示信
      'Modified by Moragn 2025/8/4 子案有指示信時母案就不必出
      If strCP09List = "" And (txtCaseField(2) <> "N" Or m_bolEngLetter = True) Then
         If m_bolEngLetter Then
            strLetterJudge = strUserNum
         Else
            strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), cp(10), field(9))
         End If
         
         m_strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), cp(4), cp(10), field(11), cp(45), field(9))
         PUB_AddAppForm cp(9), True, strLetterJudge, m_strSubject
         m_strAF01 = cp(9)
      End If
   End If
   If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
      If txtCaseField(3) <> "N" Then
         strLetterJudge = PUB_GetLetterJudgeNew("1", field(1), cp(10), field(9))
         PUB_AddLetterProgress cp(9), 0, True, strLetterJudge, False, field(26), cp(10), field(75)
         m_strLD18 = cp(9)
      End If
   End If
   'end 2018/8/22
   
   cnnConnection.CommitTrans
     Exit Function
CheckingErr:
    SaveDatabase = False
     cnnConnection.RollbackTrans
End Function


Private Sub ReadAllData()
Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String, strTemp1 As String, j As Integer
Dim adoRecord As Object, strSameName As String
'Add By Cheng 2003/02/20
Dim strCusTemp As String '受讓人代號

On Error GoTo ErrHnd
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'Modify by Morgan 2005/11/22
'If objPublicData.ReadAllData(frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.Row, 5), cp(), field(), intCaseKind, intPWhere) Then
ReDim cp(TF_CP) As String
cp(9) = frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.row, 5)
If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then

    lblCaseField(0) = cp(9)
    lblCaseField(1) = cp(1) + " - " + cp(2) + _
    IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
    IIf(cp(4) = "00", "", " - " + cp(4))
    lblCaseField(2) = TransDate(cp(6), 1)
    lblCaseField(4) = cp(13)
    lblCaseField(5) = TransDate(cp(7), 1)
    lblCaseField(3) = field(8)
   
   '92.1.18 add by sonia
   'Modify by Morgan 2006/9/20 加法國
   'Modified by Morgan 2023/3/24
   'If field(9) = "101" Or field(9) = "102" Or field(9) = "203" Then
   '  'Added by Morgan 2013/5/13
   '  If field(9) = "101" Then
   '     optChoose(2).Enabled = True
   '  Else
   '     optChoose(2).Enabled = False
   '  End If
   '  'end 2013/5/13
   PUB_SetEntityOpt field(1), field(9), field(8), OptChoose
   If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
      If strSrvDate(1) >= PA179啟用日 Then
         If field(179) = "1" Then
            OptChoose(0).Value = True
            old_Entity = OptChoose(0).Caption
         ElseIf field(179) = "2" Then
            OptChoose(1).Value = True
            old_Entity = OptChoose(1).Caption
         ElseIf field(179) = "3" Then
            OptChoose(2).Value = True
            old_Entity = OptChoose(2).Caption
         Else
            old_Entity = ""
         End If
      Else
   'end 2023/3/24
      
        If InStr(1, field(91), "大個體", 1) > 0 Then
           OptChoose(0).Value = True
           old_Entity = "大個體"
        ElseIf InStr(1, field(91), "小個體", 1) > 0 Then
           OptChoose(1).Value = True
           old_Entity = "小個體"
        'Added by Morgan 2013/5/7
        ElseIf InStr(1, field(91), "微個體", 1) > 0 Then
           OptChoose(2).Value = True
           old_Entity = "微個體"
        'end 2013/5/7
        Else
           old_Entity = ""
        End If
        
      End If 'Added by Morgan 2023/3/24
      
      'Added by Morgan 2024/12/10 個體別順序會因國家有所不同,且客戶設定目前只設定是否可減免,原預設規則只適用於1,2選項為大小個體時
      If OptChoose(0).Caption = "大個體" And OptChoose(1).Caption = "小個體" Then
      'end 2024/12/10
      
         'Add by Morgan 2004/9/24
         'Modified by Morgan 2013/5/7 +微個體
         If OptChoose(0).Value = False And OptChoose(1).Value = False And OptChoose(2).Value = False Then
            Dim stAD03 As String
            For i = 1 To 5
               'Modify by Morgan 2004/12/17 讓渡要抓受讓人
               'Modify by Morgan 2007/7/27 加"繼承"
               'If cp(10) = 讓與 Then
               If cp(10) = 讓與 Or cp(10) = 繼承 Then
                  If field(25 + i) <> "" Then
                     If i = 1 Then
                        stAD03 = PUB_GetAD03(cp(56), field(9))
                     Else
                        stAD03 = PUB_GetAD03(cp(87 + i), field(9))
                     End If
                     If stAD03 = "N" Then
                        OptChoose(0).Value = True
                        Exit For
                     '只要有未設定減免身分的公司申請人則不預設大小個體
                     ElseIf stAD03 = "" Then
                        Exit For
                     End If
                  End If
               Else
                  If field(25 + i) <> "" Then
                     stAD03 = PUB_GetAD03(field(25 + i), field(9))
                     If stAD03 = "N" Then
                        OptChoose(0).Value = True
                        Exit For
                     '只要有未設定減免身分的公司申請人則不預設大小個體
                     ElseIf stAD03 = "" Then
                        Exit For
                     End If
                  End If
               End If
               
            Next
            '若五個申請人檢查完都不是大個體則為小個體
            If OptChoose(2).Enabled = False Then 'Added by Morgan 2013/5/7 不可選微個體時才預設
               If OptChoose(0).Value = False And i = 6 Then OptChoose(1).Value = True
            End If
         End If
         
      End If 'Added by Morgan 2024/12/10
       
    End If
    '92.1.18 end
   
   'Modify By Cheng 2002/08/19
'   objPublicData.GetCasePreAgent cp(), strTemp
'   If strTemp <> "" Then
'      txtCaseField(1) = strTemp
'      CheckKeyIn 1
'   Else
'      txtCaseField(1) = ""
'   End If
   Set adoRecord = CreateObject("ADODB.Recordset")
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.SelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
   '2007/4/23 MODIFY BY SONIA 加發文日降冪排序
   'If ClsPDSelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
   'Modify by Morgan 2008/2/18 加聯絡人
   'Added by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設
   If cp(31) = "Y" Then
      AddAgent Combo1, cp, , , , cp(9), field(9), field(26)
      If Combo1 <> "" Then CheckKeyIn 1
      
   Else '非新案照原本
        If ClsPDSelectTable("select cp44||decode(cp116,null,null,'-'||cp116) from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "' and cp09<'C' and cp44 is not null order by cp27 desc", adoRecord) Then
        '2007/4/23 END
           Do While adoRecord.EOF = False
              If IsNull(adoRecord.Fields(0).Value) = False Then
                 If strSameName <> adoRecord.Fields(0).Value Then
                    Combo1.AddItem adoRecord.Fields(0).Value
                    strSameName = adoRecord.Fields(0).Value
                 End If
              End If
              adoRecord.MoveNext
           Loop
           Combo1 = Combo1.List(0)
        End If
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetCasePreAgent(cp(), strTemp) Then
        
      'Added by Morgan 2023/10/30 已有設定時不必再重新設定(IDS分案會先設,且抓預設代理人時也會剔除)
      If cp(44) <> "" Then
         Combo1 = cp(44) & IIf(cp(116) <> "", "-" & cp(116), "")
         CheckKeyIn 1
      Else
      'end 2023/10/30
           
         If ClsPDGetCasePreAgent(cp(), strTemp) Then
            Combo1 = strTemp
            CheckKeyIn 1
         End If
         
      End If 'Added by Morgan 2023/10/30
      
   End If
   'end 2016/10/27
   
   '例外
   txtCaseField(9) = "Y"
   txtCaseField(2) = "N"
   txtCaseField(8) = cp(64)
   'Modify by Morgan 2007/7/27 加繼承
   'If cp(10) <> 實體審查 And cp(10) <> 讓與 And cp(10) <> "427" Then
   If cp(10) <> 實體審查 And cp(10) <> 讓與 And cp(10) <> "427" And cp(10) <> 繼承 Then
      txtCaseField(2).Enabled = False
   Else
      txtCaseField(2).Enabled = True
      txtCaseField(2) = ""
   End If
   
   'Added by Morgan 2017/9/12
   '實體審查發文時若當日有新案發文時預設不出定稿及指示信--玫音
   If cp(10) = 實體審查 Then
      strExc(0) = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
         " and cp57 is null AND CP31='Y' and cp27=" & strSrvDate(1)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtCaseField(2) = "N"
         txtCaseField(3) = "N"
      End If
   End If
   'end 2017/9/12
   
   'Add by Morgan 2011/4/15
   chkChoose(1).Enabled = False
   chkChoose(2).Enabled = False
   'EPC實審發文時若有覆檢索報告218當日或未發文時預設不出指示信及通知函
   If field(9) = "221" And cp(10) = 實體審查 Then
      strExc(0) = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
         " and cp57 is null AND CP10='218' and (cp27 is null or cp27=" & strSrvDate(1) & ")"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtCaseField(2) = "N"
         txtCaseField(3) = "N"
      End If
      
      'EPC實審發文時若有指定費當日或未發文時預設一併送件
      chkChoose(2).Enabled = True
      strExc(0) = "select distinct cp10 from caseprogress where " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
         " and cp57 is null and cp10='215' and (cp27 is null or cp27=" & strSrvDate(1) & ")" & _
         " order by 1"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         chkChoose(2).Value = vbChecked
      End If
   End If
   
   '美國讓渡發文時若有申請程序(101,103,118,307)當日或未發文時預設不出指示信及通知函
   If field(9) = "101" And cp(10) = 讓與 Then
      strExc(0) = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
         " and cp57 is null AND CP10 IN ('101','103','113','307') and (cp27 is null or cp27=" & strSrvDate(1) & ")"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtCaseField(2) = "N"
         txtCaseField(3) = "N"
      End If
   End If
   
   '馬來西亞
   If field(9) = "018" Then
      '提供前案資料發文時若實審當日或未發文時預設不出指示信及通知函
      If cp(10) = "207" Then
         strExc(0) = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
            " and cp57 is null AND CP10='416' and (cp27 is null or cp27=" & strSrvDate(1) & ")"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            txtCaseField(2) = "N"
            txtCaseField(3) = "N"
         End If
      End If
      '實審發文時若有提供前案資料當日或未發文時預設一併送件
      If cp(10) = "416" Then
         chkChoose(1).Enabled = True
         strExc(0) = "select distinct cp10 from caseprogress where " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
            " and cp57 is null and cp10='207' and (cp27 is null or cp27=" & strSrvDate(1) & ")" & _
            " order by 1"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            chkChoose(1).Value = vbChecked
         End If
      End If
   End If
   'end 2011/4/14
   
    'edit by nickc 2007/02/02 不用 dll 了
    'If objPublicData.GetCaseProperty(cp(1), cp(10), strTemp) Then
    If ClsPDGetCaseProperty(cp(1), cp(10), strTemp) Then
       txtCaseField(4) = strTemp
    Else
       GoTo err1
    End If
    
    'Modify by Morgan 2005/11/22 讓與申請人1~5
'    'Add By Cheng 2003/02/20
'    '顯示受讓人資料
'    Me.txtCaseField(6).Text = cp(56)
'    strCusTemp = Me.txtCaseField(6).Text
'    'Modify By Cheng 2003/02/24
'    '若有受讓人資料, 則顯示名稱
'    If Me.txtCaseField(6).Text <> "" Then
'        If objPublicData.GETCUSTOMER(strCusTemp, strTemp, , strTemp1) Then
'           lblBeGive = strTemp
'           txtCaseField(7) = strTemp1
'        End If
'    End If

   Me.lblAppName(1).Caption = "": Me.txtAppAddr(1) = "": Me.txtAppAddr(1).Locked = True
   Me.txtAppNew(1).Text = cp(56)
   strCusTemp = Me.txtAppNew(1).Text
   If strCusTemp <> "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCustomer(strCusTemp, strTemp, , strTemp1) Then
      If ClsPDGetCustomer(strCusTemp, strTemp, , strTemp1) Then
         Me.lblAppName(1).Caption = strTemp
         Me.txtAppAddr(1) = strTemp1
      End If
   End If
    
   For j = 2 To 5
      Me.lblAppName(j).Caption = "": Me.txtAppAddr(j) = "": Me.txtAppAddr(j).Locked = True
      Me.txtAppNew(j).Text = cp(87 + j)
      strCusTemp = Me.txtAppNew(j).Text
      If strCusTemp <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCustomer(strCusTemp, strTemp, , strTemp1) Then
         If ClsPDGetCustomer(strCusTemp, strTemp, , strTemp1) Then
            '名稱
            Me.lblAppName(j).Caption = strTemp
            '地址
            Me.txtAppAddr(j) = strTemp1
         End If
      End If
   Next
   '2005/11/22 end
   
   'Added by Morgan 2013/6/5
   For j = 1 To 5
      SetAD j
      txtAppNew(j).Tag = txtAppNew(j) 'Added by Morgan 2020/3/26
   Next
   
   'Add By Cheng 2003/09/16
   '讀取ECP指定國家
   'edit by nickc 2007/02/02 不用 dll 了
   'If field(9) = EPC指定國家 Then objPublicData.ReadCountry intCaseKind, cp(), strCountry, True, False
   If field(9) = EPC指定國家 Then ClsPDReadCountry intCaseKind, cp(), strCountry, True, False
    
   'Add by Morgan 2007/9/3
   If cp(10) = "214" Then
      lstItems.Visible = True
      Label9.Visible = True
      txtCaseField(8).Width = 5460
      
      strExc(1) = "select cr05,cr06,cr07,cr08 from caserelation" & _
         " where cr01='" & cp(1) & "' and cr02='" & cp(2) & "' and cr03='" & cp(3) & "' and cr04='" & cp(4) & "'"
      'Add by Morgan 2007/9/20 加國內外案
      '國內案
      strExc(1) = strExc(1) & " UNION SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
         " WHERE CM01='" & cp(1) & "' AND CM02='" & cp(2) & "' AND CM03='" & cp(3) & "' AND CM04='" & cp(4) & "'"
      '國內案的其他國外案
      strExc(1) = strExc(1) & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
         "(SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
         " WHERE CM01='" & cp(1) & "' AND CM02='" & cp(2) & "' AND CM03='" & cp(3) & "' AND CM04='" & cp(4) & "')"
      '國外案
      strExc(1) = strExc(1) & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP" & _
         " WHERE CM05='" & cp(1) & "' AND CM06='" & cp(2) & "' AND CM07='" & cp(3) & "' AND CM08='" & cp(4) & "'"
      '國外案的其他國外案
      strExc(1) = strExc(1) & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
         "(SELECT CM01,CM02,CM03,CM04 FROM CASEMAP" & _
         " WHERE CM05='" & cp(1) & "' AND CM06='" & cp(2) & "' AND CM07='" & cp(3) & "' AND CM08='" & cp(4) & "')"
      'end 2007/9/20
      
      strExc(0) = "select distinct na03 from (" & strExc(1) & ") X,patent,nation" & _
         " where not(cr05='" & cp(1) & "' and cr06='" & cp(2) & "' and cr07='" & cp(3) & "' and cr08='" & cp(4) & "')" & _
         " and pa01(+)=cr05 and pa02(+)=cr06 and pa03(+)=cr07 and pa04(+)=cr08 and na01(+)=pa09 order by na03 desc"
         
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Do While Not RsTemp.EOF
            lstItems.AddItem "" & RsTemp.Fields(0) & " 前案", 0
            RsTemp.MoveNext
         Loop
      End If
   Else
      lstItems.Visible = False
      Label9.Visible = False
   End If
   'end 2007/9/3
   
   'Add by Morgan 2009/8/18
   If txtCaseField(0).Tag <> txtCaseField(0) Then
      PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8), field(16)
      txtCaseField(0).Tag = txtCaseField(0)
   End If
   
    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25
   
Else
err1:
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If
ErrHnd:
ErrorMsg
Screen.MousePointer = varSaveCursor
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   Dim strNo As String, iPos As Integer
   If Combo1.Text <> "" Then
      If CheckKeyIn(1) = -1 Then
         Cancel = True
      End If
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If Cancel = False Then
         strNo = Combo1.Text
         'Add by Morgan 2008/2/20 加聯絡人判斷
         iPos = InStr(Combo1.Text, "-")
         If iPos > 0 Then
            strNo = Left(Combo1.Text, iPos - 1)
         End If
         'end 2008/2/20
         
         If PUB_CheckStatus(strNo) = False Then
            Cancel = True
         'Added by Morgan 2012/3/7 發文都要顯示代理人備註--甄妮
         Else
            strExc(0) = "select FA29 from Fagent where " & ChgFagent(strNo) & " and FA29 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "" & RsTemp(0), vbExclamation, "代理人備註"
            End If
         'end 2012/3/7
         End If
      End If
      
      If Cancel Then Combo1.SetFocus
   End If
End Sub



Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String

Select Case Index
   Case 3
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
      If ClsPDGetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
         lblTrademarkKind = strTemp
      End If
   Case 4
      '91.12.3 MODIFY BY SONIA
      'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetStaffN(lblCaseField(Index), strTemp) Then
      If ClsPDGetStaffN(lblCaseField(Index), strTemp) Then
      '91.12.3 END
         lblSalesName = strTemp
      Else
         lblSalesName = ""
      End If
End Select
End Sub
Private Sub Form_Activate()
   If m_bolActive = True Then Exit Sub
   m_bolActive = True
   
   txtCaseField(0) = strSrvDate(2)
   ReadAllData
   'Add by Morgan 2007/7/27
   If cp(10) = 繼承 Then
      For intI = 5 To 1 Step -1
         SSTab1.Tab = intI - 1
         SSTab1.Caption = "繼承人" & intI
      Next
   End If
   'end 2007/7/27
         
   'Add by Morgan 2009/6/1
   If PUB_ChkFileNP(cp(9)) Then MsgBox "下一程序已有提申或收達期限，若為重新發文時需要先刪除後才可作業！"
   
   'Add by Amy 2013/05/27 微個體案件辦理後續辦理讓與發文時跳訊息提醒程序
   'Modify by Amy 2015/01/15 +if 判斷此道為讓與才做
   If cp(10) = "701" Then
     If field(9) = "101" And OptChoose(2).Value Then
       'Modify by Amy 2015/01/15 +CP27>0
       strExc(0) = "Select NVL(CP27,0) From CaseProgress Where CP01='" & cp(1) & "' And CP02='" & cp(2) & "' And CP03='" & cp(3) & _
                        "' And instr('" & NewCasePtyList & "',CP10)>0 And CP27 >0"
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
           If Val(cp(5)) > Val(RsTemp.Fields(0).Value) Then MsgBox "請確認受讓人的個體狀態！"
       End If
     End If
   End If
   'end 2013/05/27
   
   'Added by Morgan 2023/9/5
   If Left(cp(9), 1) = "B" Then
      'Added by Morgan 2023/9/11
      If txtCaseField(3) = "N" Then
         strExc(0) = PUB_AskBKindLetter(cp(1), cp(9), cp(10), 1)
      Else
      'end 2023/9/11
         strExc(0) = PUB_AskBKindLetter(cp(1), cp(9), cp(10))
         If txtCaseField(3) <> strExc(0) Then
            txtCaseField(3) = strExc(0)
         End If
      End If
   End If
   'end 2023/9/5
   
   txtCaseField(0).SetFocus
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   
   'Add By Sindy 2018/1/8
   m_strIR01 = frm050102_1.m_strIR01
   m_strIR02 = frm050102_1.m_strIR02
   m_strIR03 = frm050102_1.m_strIR03
   m_strIR04 = frm050102_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2018/1/8 END
   
   SSTab1.Tab = 0 'Added by Lydia 2021/05/25
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2009/8/18
   If intLeaveKind = 1 Then
      frm050102_1.Show
   ElseIf intLeaveKind = 0 Then
     Unload frm050102_1
   End If
   ShowEditForm 'Added by Morgan 2018/8/22
   
   'Set frm050102_5 = Nothing'Removed by Morgan 2021/12/10 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub txtAD_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtAD(Index)
End Sub

Private Sub txtAD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Not (KeyAscii = 8 Or KeyAscii = 89 Or KeyAscii = 78) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtAD_Validate(Index As Integer, Cancel As Boolean)
   If txtAD(Index) = "" Then
      MsgBox "請設定減免身分(Y/N)！"
      Cancel = True
   End If
End Sub

Private Sub txtAppNew_Change(Index As Integer)
   Me.lblAppName(Index).Caption = ""
   Me.txtAppAddr(Index).Text = ""
End Sub

Private Sub txtAppNew_GotFocus(Index As Integer)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'Me.txtAppNew(Index).IMEMode = 2
   CloseIme
   TextInverse Me.txtAppNew(Index)
End Sub

Private Sub txtAppNew_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtAppNew_Validate(Index As Integer, Cancel As Boolean)
   If txtAppNew(Index).Tag <> txtAppNew(Index) Then SetAD Index 'Added by Morgan 2020/3/26
     
   If CheckKeyIn(Index + 10) = -1 Then
      Cancel = True
   'Added by Morgan 2013/6/5
   'Removed by Morgan 2020/3/26
   'Else
   '   SetAD Index
   'end 2020/3/26
   'end 2013/6/5
   End If
   txtAppNew(Index).Tag = txtAppNew(Index) 'Added by Morgan 2020/3/26
End Sub

Private Sub txtCaseField_Change(Index As Integer)
Select Case Index
             Case 1
                         lblAgent = ""
             'Remove by Morgan 2005/11/22
'             Case 6
'                         lblBeGive = ""
'                         txtCaseField(7) = ""
End Select
End Sub
Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
             'Modify by Morgan 2005/11/22
             'Case 1, 2, 3, 4, 5, 6
             Case 1, 2, 3, 4, 5
                       KeyAscii = UpperCase(KeyAscii)
            
            'Add By Cheng 2002/07/30
            Case 9, 10
               KeyAscii = UpperCase(KeyAscii)
               If KeyAscii <> 8 And KeyAscii <> 89 Then
                  KeyAscii = 0
               End If
End Select
End Sub
Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = -1 Then
   Cancel = True
End If

'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
If Cancel = False And Index = 6 Then
   If PUB_CheckStatus(txtCaseField(Index).Text) = False Then Cancel = True
End If

If Cancel Then txtCaseField_GotFocus (Index)
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, strCusTemp As String, j As Integer

CheckKeyIn = -1
Select Case intIndex
             Case 0 '發文日
                        If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           CheckKeyIn = 1
                           'Add by Morgan 2009/8/18
                           If txtCaseField(0).Tag <> txtCaseField(0) Then
                              PUB_SetChkResultDate field(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8), field(16)
                              txtCaseField(0).Tag = txtCaseField(0)
                           End If
                        End If
             Case 1 '代理人
                        lblAgent = ""
                        If Combo1.Text = "" Then
                           MsgBox "代理人欄不可空白!!!", vbExclamation
                        Else
                           strCusTemp = Combo1
                           'Add by Morgan 2008/2/14 加判斷是否為聯絡人
                           If InStr(strCusTemp, "-") > 0 Then
                              If ClsPDGetContact(strCusTemp, strTemp) Then
                                 Combo1 = strCusTemp
                                 lblAgent.Caption = strTemp
                                 CheckKeyIn = 1
                              End If
                           
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetAgent(strCusTemp, strTemp) Then
                           ElseIf ClsPDGetAgent(strCusTemp, strTemp) Then
                              Combo1 = strCusTemp
                              lblAgent.Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        End If
             Case 2, 3 '是否列印指示信, 是否列印通知函
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 5 '有無前案資料
                        If cp(10) = 提供前案資料 Then
                           If txtCaseField(intIndex) = "N" Or txtCaseField(intIndex) = "Y" Then
                              CheckKeyIn = 1
                           Else
                              ShowMsg MsgText(9188)
                           End If
                        Else
                           If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Or txtCaseField(intIndex) = "Y" Then
                              CheckKeyIn = 1
                           Else
                              ShowMsg MsgText(9177)
                           End If
                        End If
             
             'Modify by Morgan 2005/11/22
'             Case 6 '受讓人名稱
'                        If txtCaseField(intIndex) = "" Then
'                           CheckKeyIn = 1
'                        Else
'                           strCusTemp = txtCaseField(intIndex)
'                           If objPublicData.GETCUSTOMER(strCusTemp, strTemp, , strTemp1) Then
'                              txtCaseField(intIndex) = strCusTemp
'                              lblBeGive = strTemp
'                              txtCaseField(7) = strTemp1
'                              CheckKeyIn = 1
'                              'Add by Morgan 2004/12/17
'                              If PUB_GetAD03(txtCaseField(intIndex), field(9)) = "N" Then
'                                 optChoose(0).Value = True
'                              Else
'                                 optChoose(1).Value = True
'                              End If
'                           End If
'                        End If
             
             Case 11, 12, 13, 14, 15
                  strCusTemp = txtAppNew((intIndex - 10))
                  If strCusTemp = "" Then
                     CheckKeyIn = 1
                  Else
                     'edit by nickc 2007/02/02 不用 dll 了
                     'If objPublicData.GetCustomer(strCusTemp, strTemp, , strTemp1) Then
                     If ClsPDGetCustomer(strCusTemp, strTemp, , strTemp1) Then
                        txtAppNew((intIndex - 10)) = strCusTemp
                        lblAppName(intIndex - 10) = strTemp
                        txtAppAddr(intIndex - 10) = strTemp1
                        CheckKeyIn = 1
                        '大小個體
                        'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
                        If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
                           'txtAppNew((intIndex - 10)).Tag = PUB_GetAD03(txtAppNew((intIndex - 10)), field(9)) 'Removed by Morgan 2020/3/26
                           For j = 1 To 5
                              'Mofified by Morgan 2020/3/26
                              'If txtAppNew((intIndex - 10)).Tag = "N" Then
                              If txtAD((intIndex - 10)) = "N" Then
                              'end 2020/3/26
                                 OptChoose(0).Value = True
                                 Exit For
                              End If
                           Next
                           If j = 6 Then
                              If OptChoose(2).Enabled = False Then 'Added by Morgan 2013/5/7 不可選微個體時才預設
                                 OptChoose(1).Value = True
                              End If
                           End If
                        End If
                     End If
                  End If
                  
             Case Else
                        CheckKeyIn = 1
End Select
End Function

Private Sub txtCaseField_GotFocus(Index As Integer)
   TextInverse txtCaseField(Index)
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

   'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   'add by nickc 2008/05/01
   If IsDebt(field(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        Exit Function
   End If
For Each objTxt In Me.txtCaseField
   If objTxt.Enabled = True Then
      Cancel = False
      txtCaseField_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Add by Morgan 2004/9/14
If Combo1.Enabled = True Then
   If Combo1.Text = "" Then
      MsgBox "代理人欄不可空白!!!", vbExclamation
      Exit Function
   End If
   Cancel = False
   Combo1_Validate Cancel
   If Cancel = True Then
      Combo1.SetFocus
      Exit Function
   End If
End If

'Add by Morgan 2005/11/22
For Each objTxt In Me.txtAppNew
   If objTxt.Enabled = True Then
      Cancel = False
      txtAppNew_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Add by Morgan 2009/9/18
If cp(10) = "701" Or cp(10) = "703" Then
   If txtAppNew(1) = "" Then
      SSTab1.Tab = 0
      MsgBox SSTab1.Caption & "欄位不可空白！"
      txtAppNew(1).SetFocus
      Exit Function
   End If
   
   'Added by Morgan 2013/6/5
   '檢查受讓人減免身分要設定
   For ii = 1 To 5
      If txtAD(ii).Enabled = True Then
         Cancel = False
         txtAD_Validate ii, Cancel
         If Cancel = True Then
            txtAD(ii).SetFocus
            Exit Function
         End If
      End If
   Next
   'end 2013/6/5
End If

'Added by Morgan 2018/9/6
'若系統不出指示信時判斷是否有工程師的指示信要寄送
m_bolEngLetter = False
If txtCaseField(2) = "N" Then
   If PUB_EngLtrChk(cp(9), txtCaseField(0).Text, m_bolEngLetter) = False Then
      Exit Function
   End If
End If
'end 2018/9/6

'Added by Morgan 2018/9/12 CFP電子化-接洽單檢查
If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
   If cp(9) < "B" And Left(cp(12), 1) <> "F" Then
      If PUB_CheckPDF3(cp(1), cp(2), cp(3), cp(4)) = False Then
         Exit Function
      End If
   End If
End If
'end 2018/9/12

'Added by Morgan 2019/4/23
'日本實審發文要設定減免身分
m_strCP81 = ""
m_strJpMemo = ""
If field(9) = "011" And cp(10) = "416" Then
   Dim stAD10 As String, stAD15 As String
   For ii = 1 To 5
      If field(25 + ii) <> "" Then
         strExc(1) = PUB_GetAD03(field(25 + ii), "011", stAD10, , stAD15)
         m_strJpMemo = m_strJpMemo & PUB_GetJpDiscountDesc(stAD10, stAD15) & ";"
         If strExc(1) = "" Then
            'Modified by Morgan 2019/6/19 改詢問是否不可減免,若是則系統自動設定--禧佩
            'MsgBox "申請人 " & field(25 + ii) & " " & GetCustomerNameAndState(field(25 + ii)) & " 尚未設定減免身分不可發文！", vbCritical, "日本實審發文減免身分檢查"
            'Exit Function
            If MsgBox("申請人【" & field(25 + ii) & " " & GetCustomerNameAndState(field(25 + ii)) & "】尚未設定減免身分！" & vbCrLf & vbCrLf & "是否要設定為【不可減免】？", vbYesNo + vbDefaultButton2 + vbExclamation, "日本實審發文減免身分檢查") = vbYes Then
               PUB_SetNoDisc field(25 + ii), field(9)
               m_strCP81 = "N"
            Else
               Exit Function
            End If
            'end 2019/6/19
         ElseIf m_strCP81 <> "N" Then
            m_strCP81 = strExc(1)
         End If
      End If
   Next
   If m_strCP81 <> "Y" Then m_strJpMemo = ""
End If
'end 2019/4/23

'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
If Pub_ChkACS112isNull(field(1), field(2), field(3), field(4), txtCP113) = True Then
    txtCP113.SetFocus
    txtCP113_GotFocus
    Exit Function
End If
'end 2021/05/25

TxtValidate = True
End Function

'Add by Morgan 2009/8/18
Private Sub lblCaseFee_Click()
   strExc(1) = cp(10)
   
   'Modify by Morgan 2009/9/10 實審要抓申請程序的審查天數
   'Modify by Morgan 2010/11/8
   'If cp(10) = "416" Then
   If cp(10) = "416" And field(16) <> "1" Then
      strExc(0) = "SELECT cp10 FROM caseprogress" & _
         " WHERE cp01='" & cp(1) & "' and cp02='" & cp(2) & "'" & _
         " and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'" & _
         " and cp10 in ('101','102','103','301','302','303','307')" & _
         " and cp27>19221111 order by cp27 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(1) = RsTemp.Fields(0)
      Else
         MsgBox "無法讀取已發文申請程序！"
         Exit Sub
      End If
   End If
   'end 2009/9/10
   
   frm12040102_2.txtCF(1) = cp(1)
   frm12040102_2.txtCF(2) = field(9)
   frm12040102_2.txtCF(3) = strExc(1)
   frm12040102_2.Show vbModal
   If Val(txtCaseField(0)) > 0 Then
      PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8), field(16)
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

'2010/1/21 ADD BY SONIA 批次發Mail
Private Sub BatchMail()
   Dim i As Integer
   For i = 1 To UBound(skMail)
        PUB_SendMail skMail(i).fiSender, skMail(i).fiReceiver, skMail(i).fiRecriverNo, skMail(i).fiSubject, skMail(i).fiContent
   Next i
   ReDim skMail(0) As SeekMails
End Sub
'2010/1/21 END

'Added by Morgan 2013/6/5
Private Sub SetAD(ByVal i As Integer)
   txtAD(i).Enabled = False
   txtAD(i).Tag = ""
   txtAD(i).Text = ""
   If cp(10) = 讓與 Or cp(10) = 繼承 Then
      'Modified by Morgan 2020/3/26
      'If field(i + 25) <> "" And (field(9) = "101" Or field(9) = "102" Or field(9) = "203") Then
      '   txtAD(i).Text = PUB_GetAD03(field(i + 25), field(9))
      '   txtAD(i).Tag = txtAD(i).Text
      '   txtAD(i).Enabled = True
      'End If
      If txtAppNew(i) <> "" And (field(9) = "101" Or field(9) = "102" Or field(9) = "203") Then
         txtAD(i).Text = PUB_GetAD03(txtAppNew(i), field(9))
         txtAD(i).Tag = txtAD(i).Text
         txtAD(i).Enabled = True
      End If
      'end 2020/3/26
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
