VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm180102 
   BorderStyle     =   1  '單線固定
   Caption         =   "表單維護作業"
   ClientHeight    =   6190
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6190
   ScaleWidth      =   8950
   Tag             =   "加班資料"
   Begin VB.Frame FrameNote 
      Height          =   610
      Left            =   4890
      TabIndex        =   84
      Top             =   690
      Width           =   4030
      Begin VB.TextBox txtB1008_14 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   86
         TabStop         =   0   'False
         Text            =   "@可補休：剩餘 3.5 天"
         Top             =   390
         Width           =   3930
      End
      Begin VB.TextBox txtB1008_2 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   85
         TabStop         =   0   'False
         Text            =   "@特別假：7天  已休3天"
         Top             =   120
         Width           =   3930
      End
   End
   Begin VB.TextBox txtB1002_01Note 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   68
      Text            =   "frm180102.frx":0000
      Top             =   5490
      Width           =   4665
   End
   Begin VB.TextBox TxtOverTimeNote 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4950
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   79
      Text            =   "frm180102.frx":0061
      Top             =   5520
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Frame Frame02 
      BorderStyle     =   0  '沒有框線
      Height          =   705
      Left            =   90
      TabIndex        =   51
      Top             =   4350
      Visible         =   0   'False
      Width           =   3825
      Begin VB.TextBox txtB101213 
         Enabled         =   0   'False
         Height          =   315
         Left            =   810
         MaxLength       =   4
         TabIndex        =   24
         Top             =   360
         Width           =   705
      End
      Begin VB.TextBox txtB1030 
         Height          =   315
         Left            =   810
         MaxLength       =   4
         TabIndex        =   23
         Top             =   30
         Width           =   705
      End
      Begin VB.Label Label31 
         Caption         =   "（請自行扣除休息時間）"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   1560
         TabIndex        =   76
         Top             =   90
         Width           =   1995
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "假日-共                     時"
         Height          =   180
         Left            =   60
         TabIndex        =   53
         Top             =   420
         Width           =   1725
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "實際時數        時"
         Height          =   180
         Left            =   60
         TabIndex        =   52
         Top             =   90
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   645
      Left            =   60
      TabIndex        =   70
      Top             =   2100
      Visible         =   0   'False
      Width           =   2895
      Begin VB.ComboBox cboSTime 
         Height          =   300
         ItemData        =   "frm180102.frx":0081
         Left            =   1680
         List            =   "frm180102.frx":0083
         Style           =   2  '單純下拉式
         TabIndex        =   72
         Top             =   0
         Width           =   1005
      End
      Begin VB.ComboBox cboETime 
         Height          =   300
         ItemData        =   "frm180102.frx":0085
         Left            =   1680
         List            =   "frm180102.frx":0087
         Style           =   2  '單純下拉式
         TabIndex        =   71
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "起日上班時段："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   390
         TabIndex        =   74
         Top             =   30
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "迄日下班時段："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   390
         TabIndex        =   73
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Frame Frame03 
      BorderStyle     =   0  '沒有框線
      Height          =   945
      Left            =   5670
      TabIndex        =   47
      Top             =   720
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txtB1002_03Note 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   30
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   75
         Text            =   "frm180102.frx":0089
         Top             =   60
         Width           =   2925
      End
      Begin VB.TextBox txtB1014 
         Height          =   315
         Left            =   540
         MaxLength       =   1
         TabIndex        =   2
         Top             =   300
         Width           =   225
      End
      Begin MSForms.TextBox txtB1015 
         Height          =   285
         Left            =   540
         TabIndex        =   3
         Top             =   630
         Width           =   2535
         VariousPropertyBits=   679495707
         ScrollBars      =   3
         Size            =   "4471;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "差程："
         Height          =   180
         Left            =   30
         TabIndex        =   50
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "(1:長程 2:短程 3:大陸 4:國外)"
         Height          =   180
         Left            =   780
         TabIndex        =   49
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "地點："
         Height          =   180
         Left            =   30
         TabIndex        =   48
         Top             =   660
         Width           =   540
      End
   End
   Begin VB.CheckBox Chk1Day 
      Caption         =   "非整日"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtB1005_1 
      Height          =   300
      Left            =   960
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1004 
      Height          =   300
      Left            =   1410
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1290
      Width           =   945
   End
   Begin VB.TextBox txtB1005_2 
      Height          =   300
      Left            =   1770
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1007_2 
      Height          =   300
      Left            =   4170
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1007_1 
      Height          =   300
      Left            =   3300
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1001 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   120
      Width           =   945
   End
   Begin VB.TextBox txtB1018 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   420
      Width           =   1845
   End
   Begin VB.TextBox txtB1006 
      Height          =   300
      Left            =   3810
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1290
      Width           =   945
   End
   Begin VB.ComboBox CboB1002 
      Height          =   300
      ItemData        =   "frm180102.frx":00AB
      Left            =   960
      List            =   "frm180102.frx":00AD
      TabIndex        =   0
      Top             =   690
      Width           =   1695
   End
   Begin VB.TextBox txtB1003 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   420
      Width           =   645
   End
   Begin VB.ComboBox CboB1008 
      Height          =   300
      ItemData        =   "frm180102.frx":00AF
      Left            =   3300
      List            =   "frm180102.frx":00B1
      TabIndex        =   1
      Top             =   690
      Width           =   1515
   End
   Begin VB.CommandButton cmdagainSend 
      Caption         =   "重送(&R)"
      Height          =   360
      Left            =   7290
      TabIndex        =   28
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "送出(&O)"
      Height          =   360
      Left            =   6450
      TabIndex        =   27
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "刪除(&D)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5610
      TabIndex        =   26
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改(&M)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4770
      TabIndex        =   25
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8130
      TabIndex        =   29
      Top             =   30
      Width           =   800
   End
   Begin VB.Frame Frame01 
      BorderStyle     =   0  '沒有框線
      Height          =   495
      Left            =   3090
      TabIndex        =   54
      Top             =   2070
      Width           =   1965
      Begin VB.TextBox txtB1010 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1110
         MaxLength       =   5
         TabIndex        =   12
         Top             =   30
         Width           =   525
      End
      Begin VB.TextBox txtB1009 
         Enabled         =   0   'False
         Height          =   315
         Left            =   270
         MaxLength       =   3
         TabIndex        =   11
         Top             =   30
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "共             日              時"
         Height          =   180
         Left            =   30
         TabIndex        =   55
         Top             =   90
         Width           =   1890
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm180102.frx":00B3
      Height          =   1485
      Left            =   4710
      TabIndex        =   30
      Top             =   3960
      Width           =   4215
      _ExtentX        =   7426
      _ExtentY        =   2628
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
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
   Begin MSForms.Label Label26 
      Height          =   195
      Left            =   150
      TabIndex        =   83
      Top             =   5970
      Width           =   7905
      VariousPropertyBits=   27
      Caption         =   "CREATE :                                                    UPDATE : "
      Size            =   "13944;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBoss 
      Height          =   285
      Index           =   4
      Left            =   3750
      TabIndex        =   21
      Top             =   3630
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBoss 
      Height          =   285
      Index           =   2
      Left            =   6510
      TabIndex        =   19
      Top             =   3330
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBoss 
      Height          =   285
      Index           =   1
      Left            =   3750
      TabIndex        =   18
      Top             =   3330
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBoss 
      Height          =   285
      Index           =   0
      Left            =   1020
      TabIndex        =   17
      Top             =   3330
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBoss 
      Height          =   285
      Index           =   3
      Left            =   1020
      TabIndex        =   20
      Top             =   3630
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboEmp 
      Height          =   285
      Index           =   2
      Left            =   6510
      TabIndex        =   15
      Top             =   2520
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboEmp 
      Height          =   285
      Index           =   1
      Left            =   6510
      TabIndex        =   14
      Top             =   2220
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboEmp 
      Height          =   285
      Index           =   0
      Left            =   6510
      TabIndex        =   13
      Top             =   1920
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1207 
      Height          =   1935
      Left            =   990
      TabIndex        =   22
      Top             =   3960
      Width           =   3675
      VariousPropertyBits=   -1466939361
      ScrollBars      =   3
      Size            =   "6482;3413"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1011 
      Height          =   555
      Left            =   1020
      TabIndex        =   16
      Top             =   2790
      Width           =   7875
      VariousPropertyBits=   -1466939365
      MaxLength       =   120
      ScrollBars      =   3
      Size            =   "13891;979"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1003_2 
      Height          =   285
      Left            =   1650
      TabIndex        =   82
      Top             =   420
      Width           =   1605
      VariousPropertyBits=   679495711
      ScrollBars      =   3
      Size            =   "2831;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblEndW 
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   4800
      TabIndex        =   80
      Top             =   1350
      Width           =   825
   End
   Begin VB.Label LblStarW 
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   2400
      TabIndex        =   81
      Top             =   1350
      Width           =   825
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "　　　交人事處。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   180
      TabIndex        =   78
      Top             =   2580
      Width           =   2730
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "注意：教育召集請將教召令影本"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   180
      TabIndex        =   77
      Top             =   2370
      Width           =   2730
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "(止日的分：只接受10,20,30,40,50,00之值)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   4980
      TabIndex        =   69
      Top             =   1710
      Width           =   3225
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "審核主管5："
      Height          =   180
      Left            =   2760
      TabIndex        =   67
      Top             =   3660
      Width           =   990
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "(未滿半小時以半小時計)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3150
      TabIndex        =   66
      Top             =   2610
      Width           =   1920
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人(3)："
      Height          =   180
      Left            =   5190
      TabIndex        =   65
      Top             =   2550
      Width           =   1290
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "簽核意見："
      Height          =   180
      Left            =   30
      TabIndex        =   64
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "表單編號："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   63
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人(2)："
      Height          =   180
      Left            =   5190
      TabIndex        =   62
      Top             =   2250
      Width           =   1290
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "目前表單狀態："
      Height          =   180
      Left            =   3555
      TabIndex        =   61
      Top             =   420
      Width           =   1260
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   3150
      X2              =   4950
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2850
      TabIndex        =   60
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "時"
      Height          =   180
      Left            =   3930
      TabIndex        =   59
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      Height          =   180
      Index           =   2
      Left            =   3390
      TabIndex        =   58
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "            迄"
      Height          =   180
      Left            =   3570
      TabIndex        =   57
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "分"
      Height          =   180
      Left            =   4770
      TabIndex        =   56
      Top             =   1710
      Width           =   180
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   930
      X2              =   4950
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "表單類別："
      Height          =   180
      Left            =   30
      TabIndex        =   46
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   45
      Top             =   420
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "假別："
      Height          =   180
      Left            =   2730
      TabIndex        =   44
      Top             =   750
      Width           =   540
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   960
      X2              =   2760
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "分"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2400
      TabIndex        =   43
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "時間：                 起                    "
      Height          =   180
      Left            =   390
      TabIndex        =   42
      Top             =   1020
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   17
      Left            =   990
      TabIndex        =   41
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "時"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1560
      TabIndex        =   40
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人(1)："
      Height          =   180
      Left            =   5190
      TabIndex        =   39
      Top             =   1980
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "事由："
      Height          =   180
      Left            =   390
      TabIndex        =   38
      Top             =   2850
      Width           =   540
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "審核主管1："
      Height          =   180
      Left            =   30
      TabIndex        =   37
      Top             =   3390
      Width           =   990
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "審核主管2："
      Height          =   180
      Left            =   2760
      TabIndex        =   36
      Top             =   3390
      Width           =   990
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "審核主管3："
      Height          =   180
      Left            =   5490
      TabIndex        =   35
      Top             =   3390
      Width           =   990
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "審核主管4："
      Height          =   180
      Left            =   30
      TabIndex        =   34
      Top             =   3660
      Width           =   990
   End
End
Attribute VB_Name = "frm180102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/8/8
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
Dim m_B1003 As String
Public m_B1017 As String
Dim m_B1018 As String
Dim m_B1019 As String
Dim m_B1023 As String
Dim m_ABS001_1 As String
Dim m_ABS001_2 As String
Dim m_ABS001_3 As String
Dim m_cboEmp(3) As String
Dim m_cboBoss(5) As String
Dim i As Integer, j As Integer, k As Integer, h As Integer
Dim strSubject As String, strContent As String
Dim m_B1004 As String, m_B1005_1 As String, m_B1005_2 As String, m_B1006 As String
Dim m_B1007_1 As String, m_B1007_2 As String, m_B1028 As String, m_B1029 As String
Dim m_B1014 As String
Dim m_B1008 As String
Dim bolChk As Boolean
Dim dblPrevRow As Double
Dim m_BossNum As Integer
Dim dHour As Double
Dim strSysDtBef1M As String '當月前1個月
Dim strSysDtAft3M As String '系統日後3個月 Add By Sindy 2014/10/7
Dim m_ST13 As String 'Add By Sindy 2013/3/20
Dim m_PrevForm As Form '前一畫面 'Add By Sindy 2013/7/12
Dim bolSetCboEmp As Boolean 'Add By Sindy 2018/8/3
Dim bolSetCboEmpSir As Boolean 'Add By Sindy 2018/8/3
Dim dblDay As Double


'Add By Sindy 2013/7/12
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("簽核人員", "身份", "日期", "時間", "簽核結果", "B1104")
   arrGridHeadWidth = Array(1050, 600, 800, 800, 800, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Public Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   m_EditMode = 0
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   '出缺勤電子簽核主檔
   'Modify By Sindy 2016/12/27 +,B1030
   strSql = "Select B1001,B1002,B1003,B1004,substr(ltrim(to_char('0000'||to_char(B1005),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),3,2) B1005,B1006,substr(ltrim(to_char('0000'||to_char(B1007),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),3,2) B1007,B1008||' '||AC03 B1008,B1009,B1010,B1011,B1012,B1013,B1014,B1015,B1016,B1017," & B1018CName & " B1018,B1019,B1020,B1021,B1022,B1023,B1024,B1025,B1026,B1027,substr(ltrim(to_char('0000'||to_char(B1028),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1028),'0000')),3,2) B1028,substr(ltrim(to_char('0000'||to_char(B1029),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1029),'0000')),3,2) B1029,B1030 " & _
            "From ABS010, allcode " & _
            "Where ac01(+)='04' and B1008=ac02(+) " & _
            "and B1001='" & Me.txtB1001 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   m_B1003 = "": m_B1017 = "": m_B1018 = "": m_B1019 = ""
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("B1001")) Then txtB1001 = rsTmp.Fields("B1001")
      If Not IsNull(rsTmp.Fields("B1002")) Then CboB1002 = GetB1002Value(rsTmp.Fields("B1002"))
      Call CboB1002_Click
      If Not IsNull(rsTmp.Fields("B1003")) Then txtB1003 = rsTmp.Fields("B1003"): m_B1003 = rsTmp.Fields("B1003"): txtB1003_2 = GetPrjSalesNM(rsTmp.Fields("B1003"))
      If Not IsNull(rsTmp.Fields("B1004")) Then txtB1004 = ChangeWStringToTString(rsTmp.Fields("B1004"))
      If Not IsNull(rsTmp.Fields("B1005")) Then txtB1005_1 = Left(rsTmp.Fields("B1005"), 2): txtB1005_2 = Right(rsTmp.Fields("B1005"), 2)
      If Not IsNull(rsTmp.Fields("B1006")) Then txtB1006 = ChangeWStringToTString(rsTmp.Fields("B1006"))
      If Not IsNull(rsTmp.Fields("B1007")) Then txtB1007_1 = Left(rsTmp.Fields("B1007"), 2): txtB1007_2 = Right(rsTmp.Fields("B1007"), 2)
      If Not IsNull(rsTmp.Fields("B1008")) Then CboB1008 = Trim(rsTmp.Fields("B1008"))
      If Not IsNull(rsTmp.Fields("B1009")) Then txtB1009 = rsTmp.Fields("B1009")
      If Not IsNull(rsTmp.Fields("B1010")) Then txtB1010 = rsTmp.Fields("B1010")
      If Not IsNull(rsTmp.Fields("B1011")) Then txtB1011 = rsTmp.Fields("B1011")
      
      'Add By Sindy 2021/8/13
      SetB102829Combo cboSTime, 1, txtB1004, txtB1003
      SetB102829Combo cboETime, 2, txtB1004, txtB1003
      '2021/8/13 END
      
      'Add By Sindy 2020/8/14 顯示星期幾
      If Val(txtB1004) > 0 Then
         LblStarW.Caption = "(" & GetWeekDay(CDate(Format(DBDATE(txtB1004), "####/##/##"))) & ")"
      End If
      If Val(txtB1006) > 0 Then
         LblEndW.Caption = "(" & GetWeekDay(CDate(Format(DBDATE(txtB1006), "####/##/##"))) & ")"
      End If
      '2020/8/14 END
      
      'Add By Sindy 2016/12/26
      If Not IsNull(rsTmp.Fields("B1012")) Then
         Label16.Caption = "平日-共                     時"
         txtB101213.Text = rsTmp.Fields("B1012")
      ElseIf Not IsNull(rsTmp.Fields("B1013")) Then
         Label16.Caption = "假日-共                     時"
         txtB101213.Text = rsTmp.Fields("B1013")
      End If
      If Not IsNull(rsTmp.Fields("B1030")) Then
         txtB1030 = rsTmp.Fields("B1030")
      Else
         txtB1030 = txtB101213
      End If
      '2016/12/26 END
      If Not IsNull(rsTmp.Fields("B1014")) Then txtB1014 = rsTmp.Fields("B1014")
      If Not IsNull(rsTmp.Fields("B1015")) Then txtB1015 = rsTmp.Fields("B1015")
      'If Not IsNull(rsTmp.Fields("B1016")) Then m_B1016 = rsTmp.Fields("B1016")
      If Not IsNull(rsTmp.Fields("B1017")) Then m_B1017 = rsTmp.Fields("B1017")
      If Not IsNull(rsTmp.Fields("B1018")) Then txtB1018 = rsTmp.Fields("B1018"): Call GetB1018CodeOrCName(m_B1018, rsTmp.Fields("B1018"))
      If Not IsNull(rsTmp.Fields("B1028")) And rsTmp.Fields("B1028") <> "00:00" Then
         For i = 0 To cboSTime.ListCount - 1
            If cboSTime.List(i) = Format(Format(rsTmp.Fields("B1028"), "hhmm"), "00:00") Then
               cboSTime.ListIndex = i
               Exit For
            End If
         Next i
'         Label1(3).Visible = True
'         cboSTime.Visible = True
         Frame1.Visible = True
         Chk1Day.Value = 1
      Else
'         Label1(3).Visible = False
'         cboSTime.Visible = False
         Frame1.Visible = False
         Chk1Day.Value = 0
      End If
      If Not IsNull(rsTmp.Fields("B1029")) And rsTmp.Fields("B1029") <> "00:00" Then
         For i = 0 To cboETime.ListCount - 1
            If cboETime.List(i) = Format(Format(rsTmp.Fields("B1029"), "hhmm"), "00:00") Then
               cboETime.ListIndex = i
               Exit For
            End If
         Next i
'         Label1(4).Visible = True
'         cboETime.Visible = True
         Frame1.Visible = True
         Chk1Day.Value = 1
      Else
'         Label1(4).Visible = False
'         cboETime.Visible = False
         Frame1.Visible = False
         Chk1Day.Value = 0
      End If
      'If (cboSTime.Visible = True And cboETime.Visible = True) Or rsTmp.Fields("B1002") = 表單類別_加班 Or Left(rsTmp.Fields("B1008"), 2) = "08" Then
      'Modify By Sindy 2012/4/13
      'If rsTmp.Fields("B1002") = 表單類別_加班 Or rsTmp.Fields("B1002") = 表單類別_出差 Then
      If rsTmp.Fields("B1002") = 表單類別_加班 Then
         Chk1Day.Value = 1
      End If
      
      If IsNull(rsTmp.Fields("B1019")) Then
         FrameNote.Visible = True 'Modify By Sindy 2024/11/5
'         txtB1008_2.Visible = True
'         'txtB1008_2 = GetCurrSpecRestDay(Trim(txtB1003))
      Else
         m_B1019 = rsTmp.Fields("B1019")
         FrameNote.Visible = False 'Modify By Sindy 2024/11/5
'         txtB1008_2.Visible = False
      End If
      
      'Add By Sindy 2015/11/19
      '記錄計算完畢當時的日期及時間,方便比對是否有需要重新計算
      m_B1004 = Val(txtB1004)
      m_B1005_1 = Val(txtB1005_1)
      m_B1005_2 = Val(txtB1005_2)
      m_B1006 = Val(txtB1006)
      m_B1007_1 = Val(txtB1007_1)
      m_B1007_2 = Val(txtB1007_2)
      m_B1014 = txtB1014
      m_B1028 = Val(Format(cboSTime.Text, "hhmm"))
      m_B1029 = Val(Format(cboETime.Text, "hhmm"))
      m_B1008 = Left(CboB1008, 2)
      '2015/11/19 END
      Call UpdateCUID(rsTmp)
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   rsTmp.Close
   
   Call SetCtrlReadOnly(False, False)
   'Add By Sindy 2017/1/10
   SetABS001_1Combo txtB1003
   SetABS001_2Combo txtB1003
   '2017/1/10 END
   
   '先清空欄位值
   Call ClearFieldCbo
   
   m_BossNum = 0
   If Trim(txtB1001) <> "" Then
      '表單流程備註檔
      SetABS012TextBox txtB1207, txtB1001
      '表單簽核檔
      strSql = "SELECT ST02||nvl(B1108,'') 簽核人員," & B1102CName & " 身份,sqldateT(B1105) 日期,sqltime6(B1106) 時間," & B1107CName & " 簽核結果,B1104 FROM ABS011,Staff WHERE B1101='" & txtB1001 & "' and B1104=ST01(+) order by B1102,B1103 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         Set GRD1.Recordset = rsTmp
         For i = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(i, 1) = "職代" Then
               For j = 0 To cboEmp.UBound
                  If cboEmp(j).Text = "" Then
                     cboEmp(j).Text = SetCboStaffName(GRD1.TextMatrix(i, 5)): m_cboEmp(j) = GRD1.TextMatrix(i, 5)
                     If GRD1.TextMatrix(i, 2) <> "" Then cboEmp(j).Enabled = False
                     Exit For
                  End If
               Next j
            ElseIf GRD1.TextMatrix(i, 1) = "審核主管" Then
               m_BossNum = m_BossNum + 1
               For j = 0 To CboBoss.UBound
                  If CboBoss(j).Text = "" Then
                     CboBoss(j).Text = SetCboStaffName(GRD1.TextMatrix(i, 5)): m_cboBoss(j) = GRD1.TextMatrix(i, 5)
                     If GRD1.TextMatrix(i, 2) <> "" Then CboBoss(j).Enabled = False
                     Exit For
                  End If
               Next j
            End If
         Next i
      End If
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
'   If rsTmp.RecordCount > 0 Then
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   
   '人事處已簽收,簽核完畢不可再異動資料
   If m_B1019 <> "" Then
      cmdModify.Enabled = False
      cmdDel.Enabled = False
      cmdSend.Enabled = False
      cmdagainSend.Enabled = False
   Else
      '尚未送簽核
      If m_B1017 = "" Then
         cmdModify.Enabled = False
         'Modify By Sindy 2013/5/10 Mark
'         If m_B1018 = 主管代填 Then
            cmdDel.Enabled = True
'         Else
'            cmdDel.Enabled = False
'         End If
         cmdSend.Enabled = True
         cmdagainSend.Enabled = False
         Call SetCtrlReadOnly(True, False)
         m_EditMode = 1 '新增
      Else
         '下一處理人員=自己
         If m_B1017 = strUserNum And m_B1003 = strUserNum Then
            cmdModify.Enabled = True
            cmdDel.Enabled = True
         '下一處理人員<>自己
         Else
            cmdModify.Enabled = False
            cmdDel.Enabled = False
         End If
         cmdSend.Enabled = False
         If m_B1018 = 送人事處簽收 Then
            cmdagainSend.Enabled = False
         Else
            'Modify By Sindy 2011/10/11
            'cmdagainSend.Enabled = True
            cmdagainSend.Enabled = False
         End If
      End If
   End If
   
   '檢查人事系統裡是否已有表單編號
   If ChkPerSysB1001Exist(txtB1001, txtB1003) = True Then
      cmdDel.Enabled = False 'Modify By Sindy 2014/10/30
      'Modify By Sindy 2011/12/6
      '尚未送簽核
      If m_B1017 = "" Then
         CboB1002.Enabled = False
         CboB1008.Enabled = False
         Chk1Day.Enabled = False
         txtB1004.Enabled = False
         txtB1005_1.Enabled = False
         txtB1005_2.Enabled = False
         txtB1006.Enabled = False
         txtB1007_1.Enabled = False
         txtB1007_2.Enabled = False
      Else
      '2011/12/6 End
         cmdModify.Enabled = False
         'cmdDel.Enabled = False 'Modify By Sindy 2014/10/30 Mark
         Call SetCtrlReadOnly(False, False)
         m_EditMode = 0
      End If
   End If
   
   If Left(Trim(CboB1002), 2) <> 表單類別_請假 Then FrameNote.Visible = False 'Modify By Sindy 2024/11/5 txtB1008_2.Visible = False
   
   'Modify By Sindy 2011/10/11 鎖住職代及審核主管
   For i = 0 To cboEmp.UBound
      cboEmp(i).Enabled = False
   Next i
   For i = 0 To CboBoss.UBound
      CboBoss(i).Enabled = False
   Next i
   
   'Add By Sindy 2013/5/10
   If m_B1018 = 主管代填 Then
      Call GetCountDayHour(False)
      'Add By Sindy 2022/5/5
      If bolSetCboEmp = False Then '無設定職代
         For i = 0 To cboEmp.UBound
            cboEmp(i).Enabled = True
         Next i
      End If
      If bolSetCboEmpSir = False Then '無設定審核主管
         For i = 0 To CboBoss.UBound
            CboBoss(i).Enabled = True
         Next i
      End If
      '2022/5/5 END
      'Add By Sindy 2025/9/5 純顯示打勾,不執行Chk1Day_Click裡面的程式段
      If Val(txtB1010) > 0 Then
         Chk1Day.Tag = "1"
         Chk1Day.Value = 1 '非整日
      End If
      '2025/9/5 END
   End If
   '2013/5/10 End

EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub ClearFieldCbo()
   For i = 0 To cboEmp.UBound
      m_cboEmp(i) = Empty
      cboEmp(i) = Empty
      cboEmp(i).Enabled = True
   Next i
   For i = 0 To CboBoss.UBound
      m_cboBoss(i) = Empty
      CboBoss(i) = Empty
      CboBoss(i).Enabled = True
   Next i
End Sub

Private Sub ClearField()
   txtB1001 = Empty
   CboB1002 = Empty
   txtB1003 = strUserNum
   txtB1003_2 = strUserName
   txtB1004 = Empty
   txtB1006 = Empty
'   Chk1Day.Value = 0 '整日
'   Call Chk1Day_Click
   CboB1008 = Empty
   'Add By Sindy 2024/11/5
   FrameNote.Visible = False
   txtB1008_2 = Empty
   txtB1008_14 = Empty
   '2024/11/5 END
   txtB1009 = Empty
   txtB1010 = Empty
   txtB1011 = Empty
   txtB1030 = Empty
   txtB101213 = Empty
   txtB1014 = Empty
   txtB1015 = Empty
   txtB1018 = Empty
   txtB1207 = Empty
   GRD1.Clear
   SetGrd
   Call ClearFieldCbo
   LblStarW.Caption = Empty 'Add By Sindy 2020/8/14
   LblEndW.Caption = Empty 'Add By Sindy 2020/8/14
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean, bolModify As Boolean)
   CboB1002.Enabled = bEnable
   CboB1008.Enabled = bEnable
   txtB1004.Enabled = bEnable
   txtB1005_1.Enabled = bEnable
   txtB1005_2.Enabled = bEnable
   txtB1006.Enabled = bEnable
   txtB1007_1.Enabled = bEnable
   txtB1007_2.Enabled = bEnable
'   txtB1009.Enabled = bEnable
'   txtB1010.Enabled = bEnable
   txtB1011.Enabled = bEnable
   txtB1030.Enabled = bEnable
   txtB1014.Enabled = bEnable
   txtB1015.Enabled = bEnable
   Chk1Day.Enabled = bEnable
   cboSTime.Enabled = bEnable
   cboETime.Enabled = bEnable
   'Modify By Sindy 2011/10/11
'   If bolModify = False Then
'      For i = 0 To CboEmp.UBound
'         CboEmp(i).Enabled = True
'      Next i
'      For i = 0 To CboBoss.UBound
'         CboBoss(i).Enabled = True
'      Next i
'   End If
   If bEnable = True Then
      Call Chk1Day_Click
   End If
End Sub

'Add By Sindy 2017/11/3
Private Sub cboETime_Click()
   If cboETime.ListIndex >= 0 Then cboSTime.ListIndex = cboETime.ListIndex
End Sub
Private Sub cboSTime_Click()
   If cboSTime.ListIndex >= 0 Then cboETime.ListIndex = cboSTime.ListIndex
End Sub
'2017/11/3 END

Private Sub Chk1Day_Click()
Dim strStarWorkTime As String, strEndWorkTime As String
   
   'Modify By Sindy 2025/9/5 + Or Chk1Day.Tag = CStr(Chk1Day.Value)
   If (m_EditMode <> 1 And m_EditMode <> 2) Or Chk1Day.Tag = CStr(Chk1Day.Value) Then
      Exit Sub
   End If
   Call Pub_GetSpecWorkHour(txtB1003, IIf(txtB1004 = "", strSrvDate(1), txtB1004), strStarWorkTime, strEndWorkTime) 'Add By Sindy 2014/2/13
   
   Me.Enabled = False
   Chk1Day.Tag = CStr(Chk1Day.Value) 'Add By Sindy 2025/9/5
   '預設值
   '特殊時段上班的人員
'   If txtB1003 = "99029" Then
   If PUB_bWkSpec = True Then
      If (txtB1005_1 = "" And txtB1005_2 = "" And txtB1007_1 = "" And txtB1007_2 = "") Or _
         (txtB1005_1 = "00" And txtB1005_2 = "00" And txtB1007_1 = "00" And txtB1007_2 = "00") Then
         'Modify By Sindy 2012/10/9 英文顧問Iain每天工作時間6小時(中午不休息)：上午 11:30 ~ 下午 17:30
         txtB1005_1 = Left(strStarWorkTime, 2)
         txtB1005_2 = Right(strStarWorkTime, 2)
         txtB1007_1 = Left(strEndWorkTime, 2)
         txtB1007_2 = Right(strEndWorkTime, 2)
      End If
'   'Add By Sindy 2012/7/9
'   '尤春彬
'   ElseIf txtB1003 = "84043" Then
'      If (txtB1005_1 = "" And txtB1005_2 = "" And txtB1007_1 = "" And txtB1007_2 = "") Or _
'         (txtB1005_1 = "00" And txtB1005_2 = "00" And txtB1007_1 = "00" And txtB1007_2 = "00") Then
'         txtB1005_1 = "08"
'         txtB1005_2 = "00"
'         txtB1007_1 = "12"
'         txtB1007_2 = "00"
'      End If
'   '2012/7/9 End
   Else
      If (txtB1005_1 = "" And txtB1005_2 = "" And txtB1007_1 = "" And txtB1007_2 = "") Or _
         (txtB1005_1 = "00" And txtB1005_2 = "00" And txtB1007_1 = "00" And txtB1007_2 = "00") Then
         txtB1005_1 = "08"
         txtB1005_2 = "00"
         txtB1007_1 = "17"
         txtB1007_2 = "00"
      End If
   End If
   txtB1005_1.Enabled = False
   txtB1005_2.Enabled = False
   txtB1007_1.Enabled = False
   txtB1007_2.Enabled = False
'   Label1(3).Visible = False
'   cboSTime.Visible = False
'   Label1(4).Visible = False
'   cboETime.Visible = False
   Frame1.Visible = False
   
   '特別假時一定要請整日
   'Modify By Sindy 2017/11/2 + And DBDATE(txtB1006) < "20180101"
   'Modify By Sindy 2017/12/28 + And strSrvDate(1) < "20180101"
   If CboB1008.Visible = True And Left(CboB1008, 2) = "08" And strSrvDate(1) < "20180101" And DBDATE(txtB1006) < "20180101" Then
      Chk1Day.Value = 0
      Chk1Day.Enabled = False
      '固定
'      If txtB1003 = "99029" Then '伊恩
      If PUB_bWkSpec = True Then
         'Modify By Sindy 2012/10/9 英文顧問Iain每天工作時間6小時(中午不休息)：上午 11:30 ~ 下午 17:30
         txtB1005_1 = Left(strStarWorkTime, 2)
         txtB1005_2 = Right(strStarWorkTime, 2)
         txtB1007_1 = Left(strEndWorkTime, 2)
         txtB1007_2 = Right(strEndWorkTime, 2)
'      'Add By Sindy 2012/7/9
'      ElseIf txtB1003 = "84043" Then '尤春彬
'         txtB1005_1 = "08"
'         txtB1005_2 = "00"
'         txtB1007_1 = "12"
'         txtB1007_2 = "00"
'      '2012/7/9 End
      Else
         txtB1005_1 = "08"
         txtB1005_2 = "00"
         txtB1007_1 = "17"
         txtB1007_2 = "00"
      End If
   Else
      Chk1Day.Enabled = True
      If Chk1Day.Value = 1 Then '非整日
         txtB1005_1.Enabled = True
         txtB1005_2.Enabled = True
         txtB1007_1.Enabled = True
         txtB1007_2.Enabled = True
         '非加班,非同一天
         'Modify By Sindy 2011/11/16
'         If Left(CboB1002, 2) <> 表單類別_加班 And _
'            txtB1004 <> "" And txtB1006 <> "" And _
'            txtB1004 <> txtB1006 Then
         'Modify By Sindy 2012/4/13
         'If Left(CboB1002, 2) = 表單類別_請假 Then
         If Left(CboB1002, 2) = 表單類別_請假 Or Left(CboB1002, 2) = 表單類別_出差 Then
'            Label1(3).Visible = True
'            cboSTime.Visible = True
'            Label1(4).Visible = True
'            cboETime.Visible = True
            Frame1.Visible = True
         End If
         
         If m_EditMode <> 2 Then
            'If txtB1003 <> "99029" And txtB1003 <> "84043" Then '伊恩,尤春彬
            If PUB_bWkSpec = False Then '非特殊上班時段人員
               'Add By Sindy 2011/12/7 因有可能同一天,填2張不同時段的表單
               txtB1005_1 = "00"
               txtB1005_2 = "00"
               txtB1007_1 = "00"
               txtB1007_2 = "00"
               '2011/12/7 End
            End If
         End If
      Else '整日
'         If txtB1003 <> "99029" And txtB1003 <> "84043" Then '伊恩,尤春彬
         If PUB_bWkSpec = False Then '非特殊上班時段人員
            txtB1005_1 = "08"
            txtB1005_2 = "00"
            txtB1007_1 = "17"
            txtB1007_2 = "00"
         End If
      End If
   End If
   
   'Add By Sindy 2011/12/7 因有可能同一天,填2張不同時段的表單
   If Left(CboB1002, 2) = 表單類別_加班 And m_EditMode <> 2 Then
      txtB1005_1 = "00"
      txtB1005_2 = "00"
      txtB1007_1 = "00"
      txtB1007_2 = "00"
   End If
   '2011/12/7 End
   
   'Add By Sindy 2014/2/13 特殊人員開放可以自行輸入時分及不計算時數
   If PUB_bWkSpec = True And Left(CboB1002, 2) <> 表單類別_加班 Then
      txtB1005_1.Enabled = True
      txtB1005_2.Enabled = True
      txtB1007_1.Enabled = True
      txtB1007_2.Enabled = True
      txtB1009.Enabled = True
      txtB1010.Enabled = True
      txtB1009.BackColor = &H80000005
      txtB1010.BackColor = &H80000005
      If Val(txtB1009) = 0 Then txtB1009 = 0
      If Val(txtB1010) = 0 Then txtB1010 = 0
   Else
      txtB1009.Enabled = False
      txtB1010.Enabled = False
      txtB1009.BackColor = &H8000000F
      txtB1010.BackColor = &H8000000F
   End If
   '2014/2/13 END
   
   Me.Enabled = True
End Sub

Private Function AutoCount() As Boolean
'Dim dblDay As Double
Dim m_Day As Integer, m_Hour As Double
   
   AutoCount = False
   
   If PUB_bWkSpec = False Or (PUB_bWkSpec = True And Left(CboB1002, 2) = 表單類別_加班) Then  'Add By Sindy 2014/2/13
      '計算時數
      If Left(CboB1002, 2) = 表單類別_加班 Then
         Call CountHour
         'If Val(IIf(txtB1012 = "", 0, txtB1012)) > 0 Or Val(IIf(txtB1013 = "", 0, txtB1013)) > 0 Then
         If Val(txtB101213) > 0 Then
            '加班不會超過1天故*0.1
'            If Val(IIf(txtB1012 = "", 0, txtB1012)) > 0 Then dblDay = Val(txtB1012) * 0.1
'            If Val(IIf(txtB1013 = "", 0, txtB1013)) > 0 Then dblDay = Val(txtB1013) * 0.1
            dblDay = Val(txtB101213) * 0.1
         Else
            Exit Function
         End If
      'Add By Sindy 2012/3/14
      ElseIf Left(CboB1002, 2) = 表單類別_出差 Then
         Call PUB_CountHour_Busi_Trip(txtB1004, Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), txtB1006, Format(txtB1007_1, "00") & Format(txtB1007_2, "00"), m_Day, m_Hour)
         If m_Day > 0 Then
            txtB1009 = m_Day
         Else
            txtB1009 = "0"
         End If
         If m_Hour > 0 Then
            txtB1010 = m_Hour
         Else
            txtB1010 = "0"
         End If
         If Val(IIf(txtB1009 = "", 0, txtB1009)) > 0 Or Val(IIf(txtB1010 = "", 0, txtB1010)) > 0 Then
            dblDay = Val(txtB1009) + (Val(txtB1010) * 0.1)
         Else
            Exit Function
         End If
      Else
         Call CountDayHour
         If Left(CboB1002, 2) = 表單類別_請假 Then
            'Modify By Sindy 2015/2/10 人事已先行作業時,則不需要再檢查特別假天數
            If CboB1008.Enabled = True Then
            '2015/2/10 END
               If ChkSA06_08(txtB1009, txtB1010, txtB1003, txtB1004, txtB1005_1, txtB1005_2, txtB1006, txtB1007_1, txtB1007_2, CboB1008, 0, , txtB1001) = False Then
                  CboB1008.Text = ""
                  CboB1008.SetFocus
                  Exit Function
               End If
               'Add By Sindy 2014/12/31 +健檢假
               If ChkSA06_23(txtB1009, txtB1010, txtB1003, txtB1004, txtB1005_1, txtB1005_2, txtB1006, txtB1007_1, txtB1007_2, CboB1008, 0, 0) = False Then
                  CboB1008.Text = ""
                  CboB1008.SetFocus
                  Exit Function
               End If
               'Add By Sindy 2024/12/10 檢查可補休
               If ChkSA06_14(txtB1009, txtB1010, txtB1003, txtB1004, txtB1005_1, txtB1005_2, txtB1006, txtB1007_1, txtB1007_2, CboB1008, 0, , txtB1001) = False Then
                  CboB1008.Text = ""
                  CboB1008.SetFocus
                  Exit Function
               End If
            End If
         End If
         If Val(IIf(txtB1009 = "", 0, txtB1009)) > 0 Or Val(IIf(txtB1010 = "", 0, txtB1010)) > 0 Then
            dblDay = Val(txtB1009) + (Val(txtB1010) * 0.1)
         Else
            Exit Function
         End If
      End If
   End If
   AutoCount = True
   
   '記錄計算完,當時的日期及時間,方便比對是否有需要重新計算
   m_B1004 = Val(txtB1004)
   m_B1005_1 = Val(txtB1005_1)
   m_B1005_2 = Val(txtB1005_2)
   If txtB1006.Visible = True Then
      m_B1006 = Val(txtB1006)
   Else
      m_B1006 = ""
   End If
   m_B1007_1 = Val(txtB1007_1)
   m_B1007_2 = Val(txtB1007_2)
   If Frame03.Visible = True Then
      m_B1014 = txtB1014
   Else
      m_B1014 = ""
   End If
   m_B1028 = ""
   m_B1029 = ""
   If Frame1.Visible = True Then
      'If cboSTime.Visible = True And cboSTime.Text <> "" Then
      If Frame1.Visible = True And cboSTime.Text <> "" Then
         m_B1028 = Val(Format(cboSTime.Text, "hhmm"))
      End If
      'If cboETime.Visible = True And cboETime.Text <> "" Then
      If Frame1.Visible = True And cboETime.Text <> "" Then
         m_B1029 = Val(Format(cboETime.Text, "hhmm"))
      End If
   End If
   If CboB1008.Visible = True Then
      m_B1008 = Left(CboB1008, 2)
   Else
      m_B1008 = ""
   End If
   
   'If CboEmp(0).Text <> "" And CboBoss(0).Text <> "" Then Exit Function '因有可能人員自已輸入簽核人員
   'Modify By Sindy 2018/8/3 因有可能人員尚未設定資料,自已輸入職代/簽核人員
'   '先清空欄位值
'   Call ClearFieldCbo
   'Add By Sindy 2023/11/29 81040閰所長112/12/22要出差因總經理出差,自行輸入職代會在此被清掉
   If Not (cboEmp(0).Enabled = True And cboEmp(0) <> "") Then
   '2023/11/29 END
      If bolSetCboEmp = True Then '因有可能人員尚未設定資料,自已輸入職代/簽核人員
         '先清空欄位值
         Call ClearFieldCbo
      End If
      '讀取職代及審核主管
      Call GetPersonBossData(dblDay)
   End If
End Function

Private Sub Chk1Day_LostFocus()
   Call GetCountDayHour(False)
End Sub

Private Sub cmdDel_Click()
Dim strTo As String
Dim rsTmp As New ADODB.Recordset
   
On Error GoTo ErrHand
   
   If m_B1017 <> "" And m_B1017 <> txtB1003 Then
      MsgBox "表單尚在檢核中，不可刪除！", vbExclamation
      Exit Sub
   End If
   
   If m_B1003 <> txtB1003 Then
      MsgBox "表單當事人，才可以刪除！", vbExclamation
      Exit Sub
   End If
   
   If MsgBox("確定是否刪除資料？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then Exit Sub
   
'   strTo = GetBossB1107_All(txtB1001,strSubject)
'   strContent = GetEMailContent(txtB1001, 刪除)
   
   Screen.MousePointer = vbHourglass
   
   cmdDel.Enabled = False
   cnnConnection.BeginTrans
   
   '出缺勤主檔
   strSql = "DELETE FROM ABS010 WHERE B1001='" & txtB1001 & "' "
   Pub_SeekTbLog strSql '記錄刪除Log
   cnnConnection.Execute strSql
   
   '簽核檔
   strSql = "SELECT * FROM ABS011 WHERE B1101='" & txtB1001 & "' order by B1102,B1103 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            strSql = "DELETE FROM ABS011 WHERE B1101='" & txtB1001 & "' and B1102='" & rsTmp.Fields("B1102") & "' and B1103=" & rsTmp.Fields("B1103")
            Pub_SeekTbLog strSql '記錄刪除Log
            cnnConnection.Execute strSql
            .MoveNext
         Loop
      End With
   End If
   rsTmp.Close
   
   '流程備註檔
   strSql = "SELECT * FROM ABS012 WHERE B1201='" & txtB1001 & "' order by B1202 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            strSql = "DELETE FROM ABS012 WHERE B1201='" & txtB1001 & "' and B1202=" & rsTmp.Fields("B1202")
            Pub_SeekTbLog strSql '記錄刪除Log
            cnnConnection.Execute strSql
            .MoveNext
         Loop
      End With
   End If
   rsTmp.Close
   
   'Add By Sindy 2025/11/6 刪除加班單一併拿掉逾30分鐘原因
   If Left(CboB1002, 2) = 表單類別_加班 Then
      strSql = "SELECT * FROM abs015 WHERE B1501='" & txtB1003 & "' AND B1502=" & DBDATE(txtB1004)
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If "" & rsTmp.Fields("B1504") <> "" Then
            strSql = "update abs015 set B1504=null WHERE B1501='" & txtB1003 & "' AND B1502=" & DBDATE(txtB1004)
            Pub_SeekTbLog strSql '記錄Log
            cnnConnection.Execute strSql
         End If
      End If
      rsTmp.Close
   End If
   '2025/11/6 END
   
   Set rsTmp = Nothing
   cnnConnection.CommitTrans
   
'   '發E-Mail通知已簽核的職代和審核主管
'   If strTo <> "" Then
'      PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , ,, , True
'   End If
   
   Screen.MousePointer = vbDefault
   
   cmdDel.Enabled = True
   Call cmdExit_Click '結束
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cmdDel.Enabled = True
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdExit_Click()
'   frm180101.Hide
'   frm180101.QueryData
'   frm180101.Show
   '打卡異常-確認處理方式作業
   If UCase(m_PrevForm.Name) = UCase("frm180105_1") Then
      m_PrevForm.Show
   Else
      'Add By Sindy 2025/10/30 +下班逾30分鐘原因確認
      m_PrevForm.Hide
      m_PrevForm.QueryData
      m_PrevForm.Show
   End If
   Unload Me
End Sub

Private Sub cmdModify_Click()
   m_EditMode = 2 '修改
   Call SetCtrlReadOnly(True, False)
   cmdagainSend.Enabled = True 'Add By Sindy 2011/10/11
   'Add By Sindy 2016/7/7 職代及審核主管欄位無值時,開放可以輸入
   For i = 0 To cboEmp.UBound
      If Trim(cboEmp(i).Text) <> "" Then
         cboEmp(i).Enabled = False
      Else
         cboEmp(i).Enabled = True
      End If
   Next i
   For i = 0 To CboBoss.UBound
      If Trim(CboBoss(i).Text) <> "" Then
         CboBoss(i).Enabled = False
      Else
         CboBoss(i).Enabled = True
      End If
   Next i
   '2016/7/7 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   Me.txtB1001.BackColor = &H8000000F
   Me.txtB1003.BackColor = &H8000000F
   Me.txtB1003_2.BackColor = &H8000000F
   Me.txtB1018.BackColor = &H8000000F
   Me.txtB1009.BackColor = &H8000000F
   Me.txtB1010.BackColor = &H8000000F
   
   '清空欄位值
   ClearField
      
   '預設值
   SetB1002Combo CboB1002
   SetB1008Combo CboB1008
'   For i = 0 To CboEmp.UBound
      SetABS001_1Combo txtB1003
'   Next i
'   For i = 0 To CboBoss.UBound
      SetABS001_2Combo txtB1003
'   Next i
   'Add By Sindy 2021/8/11
   SetB102829Combo cboSTime, 1
   SetB102829Combo cboETime, 2
   '2021/8/11 END
   
   Me.cmdModify.Enabled = False
   Me.cmdDel.Enabled = False
   Me.cmdSend.Enabled = True
   Me.cmdagainSend.Enabled = False
   m_EditMode = 1 '新增
   
   Me.txtB1008_2 = GetCurrSpecRestDay(txtB1003)
   Me.txtB1008_14 = GetCurrFor14RestDay(txtB1003)
   Call CboB1002_Click
   
   'Add By Sindy 2012/1/6 系統日前2個月
   'strSysDtBef2M = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(strSrvDate(1))))
   'Modify By Sindy 2023/6/8 '當月前1個月
   strSysDtBef1M = DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(Left(strSrvDate(1), 6) & "01")))
   'Add By Sindy 2014/10/7 系統日後3個月
   strSysDtAft3M = DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(strSrvDate(1))))
   m_ST13 = PUB_GetST13(txtB1003) 'Add By Sindy 2013/3/20
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing 'Add By Sindy 2013/7/12
   Set frm180102 = Nothing
End Sub

'帶出表單當事人,必須送簽核的職代及審核主管資料
Private Sub GetPersonBossData(dblDay As Double)
Dim strData As String, strCon As String
Dim strTemp As Variant
Dim bolReadAllPer As Boolean '要抓取全部職代
Dim strST03 As String 'Add By Sindy 2016/12/9
Dim strA0915 As String 'Add By Sindy 2017/1/5
Dim tmpDblDay As Double 'Add By Sindy 2020/5/28
Dim bolSerialRest As Boolean 'Add By Sindy 2023/9/15
Dim ii As Integer
Dim strA0911 As String, strA0925 As String
   
   strA0911 = GetStaffA0911(txtB1003, strA0925) 'Modify By Sindy 2023/12/20
   bolReadAllPer = False 'Add By Sindy 2016/4/20
   m_BossNum = 0
   strCon = ""
   If Left(CboB1002, 2) = 表單類別_請假 Then
      strCon = "and B0301='" & Left(CboB1008, 2) & "' "
   ElseIf Left(CboB1002, 2) = 表單類別_出差 Then
      strCon = "and B0301='" & txtB1014 & "' "
   End If
   
   'Add By Sindy 2020/10/26 A6022-20201026-重送時,計算天數有誤,簽核主管帶出3天假單才簽核的王副總
'   strSql = "select b1001 from abs010 where b1003='" & txtB1003 & "' and b1002='" & Left(CboB1002, 2) & "' and b1019 is null"
'   If Trim(txtB1001) <> "" Then
'      strSql = strSql & " and b1001<>'" & txtB1001 & "'"
'   End If
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
      'Add By Sindy 2020/5/28 檢查是否有連續假單
      bolSerialRest = PUB_ChkSerialRest_ToSir(txtB1001, False, tmpDblDay, Left(CboB1002, 2), _
                     txtB1003, DBDATE(txtB1004), DBDATE(txtB1006), Val(txtB1009), Val(txtB1010)) = True
      If bolSerialRest = True Then
         If tmpDblDay > dblDay Then '加上連續假單的時數,時數應該只會更大才是合理的
            dblDay = tmpDblDay
         End If
      End If
'   End If
   '2020/10/26 END
   
   '************************* 取得 審核主管 *************************
   '*****************************************************************
   'Add By Sindy 2016/12/9
   '105年12月23日起國外部加班管理程序:假日加班(不論是周六休息日或國定假日加班)均需由國外部副總核准
   'strST03 = PUB_GetST03(txtB1003)
'   If Left(CboB1002, 2) = 表單類別_加班 And ChkWorkDay(DBDATE(txtB1004)) = False And _
'      Left(strST03, 1) = "F" And (strST03 <> "F31" And strST03 <> "F41") Then
'      CboBoss(0) = "81040"
'      Call CboBoss_LostFocus(0)
'   Else
   '2016/12/9 END
      'Add By Sindy 2017/1/5 + A0915 假日加班簽核主管
      'Modify By Sindy 2023/12/20
      If strSrvDate(1) >= 新部門啟用日 Then
         strST03 = PUB_GetST93(txtB1003)
         strSql = "select nvl(a0928,'') from acc090NEW where a0921='" & strST03 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         strA0915 = ""
         If intI = 1 Then
            strA0915 = "" & RsTemp.Fields(0)
         End If
      Else
      '2023/12/20 END
         strST03 = PUB_GetST03(txtB1003)
         strSql = "select nvl(a0915,'') from acc090 where a0901='" & strST03 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         strA0915 = ""
         If intI = 1 Then
            strA0915 = "" & RsTemp.Fields(0)
         End If
      End If
      'Modify By Sindy 2020/2/11 劉經理說取消此限制,統一做法
'      If Left(CboB1002, 2) = 表單類別_加班 And ChkWorkDay(DBDATE(txtB1004)) = False And _
'         Left(strST03, 1) = "F" And (strST03 <> "F31" And strST03 <> "F41") Then
'         '105年12月23日起國外部加班管理程序:假日加班(不論是周六休息日或國定假日加班)均需由國外部副總核准
'      Else
      '2017/1/5 END
      '2020/2/11 END
         
         'Modify By Sindy 2020/2/11
         '　　假日加班、13.公傷假、3.大陸出差、4.國外出差、24.防疫照顧假，
         '　　請在林總核准之前須先經過”所有審核主管”，之後才給林總簽核。
         If (Left(CboB1002, 2) = 表單類別_請假 And (Left(CboB1008.Text, 2) = "24" Or Left(CboB1008.Text, 2) = "13")) Or _
            (Left(CboB1002, 2) = 表單類別_出差 And (txtB1014 = "3" Or txtB1014 = "4")) Or _
            (Left(CboB1002, 2) = 表單類別_加班 And ChkWorkDay(DBDATE(txtB1004)) = False) Then
            strData = GetABS001_2(txtB1003)
         Else
         '2020/2/11 END
            '增加判斷天數
            strData = GetABS001_2(txtB1003, dblDay)
         End If
         If strData <> "" Then
            strTemp = Split(strData, ",")
            For i = 0 To UBound(strTemp)
               m_BossNum = m_BossNum + 1
               CboBoss(i) = strTemp(i)
               CboBoss(i).Enabled = False
               Call CboBoss_LostFocus(i)
            Next i
         End If
'      End If
      'Add By Sindy 2017/1/5
      '非工作日加班須由總經理核決
      If Left(CboB1002, 2) = 表單類別_加班 And ChkWorkDay(DBDATE(txtB1004)) = False Then
         If strA0915 <> "" Then
            strTemp = Split(strA0915, ";")
            For i = 0 To UBound(strTemp)
               If InStr(strData, strTemp(i)) = 0 Then
                  m_BossNum = m_BossNum + 1
                  CboBoss(m_BossNum - 1) = strTemp(i)
                  CboBoss(m_BossNum - 1).Enabled = False
                  Call CboBoss_LostFocus(m_BossNum - 1)
               End If
            Next i
         End If
         If InStr(strData & strA0915, Pub_GetSpecMan("總經理員工編號")) = 0 Then
            m_BossNum = m_BossNum + 1
            CboBoss(m_BossNum - 1) = Pub_GetSpecMan("總經理員工編號")
            CboBoss(m_BossNum - 1).Enabled = False
            Call CboBoss_LostFocus(m_BossNum - 1)
         End If
      End If
      '2017/1/5 END
      'Add By Sindy 2020/2/4
      '請在請假假別增設「24.防疫照顧假」，不發給薪資，但也不影響考績及全勤。
      '最終核准主管為總經理。
      'Modify By Sindy 2021/7/27 + And txtB1003 <> Pub_GetSpecMan("總經理員工編號")
      If Left(CboB1002, 2) = 表單類別_請假 And Left(CboB1008.Text, 2) = "24" And txtB1003 <> Pub_GetSpecMan("總經理員工編號") Then
         strData = Pub_GetSpecMan("總經理員工編號")
         If Left(CboBoss(0), 5) <> strData And _
            Left(CboBoss(1), 5) <> strData And _
            Left(CboBoss(2), 5) <> strData And _
            Left(CboBoss(3), 5) <> strData And _
            Left(CboBoss(4), 5) <> strData Then
            m_BossNum = m_BossNum + 1
            CboBoss(m_BossNum - 1) = strData
            CboBoss(m_BossNum - 1).Enabled = False
            Call CboBoss_LostFocus(m_BossNum - 1)
         End If
      End If
      '2020/2/4 END
'   End If
   
   '公司另外規定的核決權限
   'Add By Sindy 2012/4/12 董事長、副董事長、76012.所長、81040.副所長除外＜ST20=01,02,11,12＞
   'Modify By Sindy 2022/7/18 + 15.名譽所長
   strSql = "select st01 from staff where ST20 in ('01','02','11','12','15') and st01='" & txtB1003 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI <> 1 Then
   '2012/4/12 End
      If strCon <> "" Then
         'Modify By Sindy 2023/9/15
         If bolSerialRest = True Then
            For ii = 1 To 10
               If g_strB1002(ii) <> "" Then
                  strCon = "and B0301='" & g_strB1008(ii) & "' "
                  'Modify By Sindy 2024/10/22
                  'B0303='TOT' => B0303 in('TOT','" & IIf(strA0925 <> "", strA0925, strA0911) & "')
                  '排序加 decode(B0303,'TOT',1,0) asc,
                  strSql = "select B0304 from ABS003 where B0303 in('TOT','" & IIf(strA0925 <> "", strA0925, strA0911) & "') and " & g_strDay(ii) & ">B0302 " & strCon & "order by decode(B0303,'TOT',1,0) asc,B0302 asc "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     With RsTemp
                        .MoveFirst
                        Do While Not .EOF
                           If Not IsNull(RsTemp.Fields(0)) Then
                              'Modify By Sindy 2024/10/30 增加排除當事者 +txtB1003 <> RsTemp.Fields(0)
                              If Left(CboBoss(0), 5) <> RsTemp.Fields(0) And _
                                 Left(CboBoss(1), 5) <> RsTemp.Fields(0) And _
                                 Left(CboBoss(2), 5) <> RsTemp.Fields(0) And _
                                 Left(CboBoss(3), 5) <> RsTemp.Fields(0) And _
                                 Left(CboBoss(4), 5) <> RsTemp.Fields(0) And _
                                 txtB1003 <> RsTemp.Fields(0) Then
                                 For i = 0 To CboBoss.UBound
                                    If CboBoss(i) = "" Then
                                       m_BossNum = m_BossNum + 1
                                       CboBoss(i) = RsTemp.Fields(0)
                                       CboBoss(i).Enabled = False
                                       Call CboBoss_LostFocus(i)
                                       Exit For
                                    End If
                                 Next i
                              End If
                           End If
                           .MoveNext
                        Loop
                     End With
                  End If
               Else
                  Exit For
               End If
            Next ii
         Else
         '2023/9/15 END
            'Modify By Sindy 2024/10/22
            'B0303='TOT' => B0303 in('TOT','" & IIf(strA0925 <> "", strA0925, strA0911) & "')
            '排序加 decode(B0303,'TOT',1,0) asc,
            strSql = "select B0304 from ABS003 where B0303 in('TOT','" & IIf(strA0925 <> "", strA0925, strA0911) & "') and " & dblDay & ">B0302 " & strCon & "order by decode(B0303,'TOT',1,0) asc,B0302 asc "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               With RsTemp
                  .MoveFirst
                  Do While Not .EOF
                     If Not IsNull(RsTemp.Fields(0)) Then
                        'Modify By Sindy 2024/10/30 增加排除當事者 +txtB1003 <> RsTemp.Fields(0)
                        If Left(CboBoss(0), 5) <> RsTemp.Fields(0) And _
                           Left(CboBoss(1), 5) <> RsTemp.Fields(0) And _
                           Left(CboBoss(2), 5) <> RsTemp.Fields(0) And _
                           Left(CboBoss(3), 5) <> RsTemp.Fields(0) And _
                           Left(CboBoss(4), 5) <> RsTemp.Fields(0) And _
                           txtB1003 <> RsTemp.Fields(0) Then
                           For i = 0 To CboBoss.UBound
                              If CboBoss(i) = "" Then
                                 m_BossNum = m_BossNum + 1
                                 CboBoss(i) = RsTemp.Fields(0)
                                 CboBoss(i).Enabled = False
                                 Call CboBoss_LostFocus(i)
                                 Exit For
                              End If
                           Next i
                        End If
                     End If
                     .MoveNext
                  Loop
               End With
            End If
         End If
      End If
   End If
   
   '************************* 取得 職代 **************************
   '**************************************************************
   If cboEmp(0).Visible = True Then
      'Add By Sindy 2018/8/3 + if
      If bolSetCboEmp = True Then
      '2018/8/3 END
         SetABS001_1Combo txtB1003 'Add By Sindy 2017/1/10 王副總提若職代的請假時間含蓋了請假人的請假時間, 則不可以出現
         Call GetABS001_1(txtB1003, m_ABS001_1, m_ABS001_2, m_ABS001_3)
         'Add By Sindy 2017/8/21 設定人事雙職代
         If InStr(m_ABS001_1, ",") > 0 Then
            For j = 1 To intDutyItem
               If PubABS001_1(j) <> "" Then
                  strData = PubABS001_1(j)
               Else
                  Exit For
               End If
               strTemp = Split(strData, ",")
               For i = 0 To UBound(strTemp)
                  '檢查取得的職代和表單當事人是否有相同的請假區間,若有,則找下一組職代
                  If CheckIsPersonRestSector(CStr(strTemp(i)), txtB1004, Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), txtB1006, Trim(txtB1007_1.Text & ":" & txtB1007_2.Text), txtB1001) = False And _
                     CheckIsPersonRest(CStr(strTemp(i)), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
                     If i = UBound(strTemp) Then '最後一位時才一起寫入職代資料
                        For h = 0 To UBound(strTemp)
                           For k = 0 To cboEmp.UBound
                              If cboEmp(k) = "" Then
                                 cboEmp(k) = strTemp(h)
                                 cboEmp(k).Enabled = False
                                 Call CboEmp_LostFocus(k)
                                 Exit For
                              End If
                           Next k
                        Next h
                     End If
                  Else
                     Exit For
                  End If
               Next i
               If cboEmp(0) <> "" Then Exit For  '若有取得職代,則離開迴圈
            Next j
         Else
         '2017/8/21 END
            For j = 1 To 3 '有3組職代
               strData = ""
               If j = 1 And m_ABS001_1 <> "" Then strData = m_ABS001_1
               If j = 2 And m_ABS001_2 <> "" Then strData = m_ABS001_2
               If j = 3 And m_ABS001_3 <> "" Then strData = m_ABS001_3
               If strData <> "" Then
                  '檢查取得的職代和表單當事人是否有相同的請假區間,若有,則找下一職代
                  strTemp = Split(strData, ",")
                  For i = 0 To UBound(strTemp)
                     '1050428-01
                     '2.同仁請假或出差在二天(不含)以上，若第一職代同時有請假或出差，
                     '不論重疊時間長短(完全重壘除外仍應由第二職簽核)，則由第一職代與第二職代均需簽職代。
                     '請假同仁、第一職代、第二職代與簽核主管在簽核時均會跑出「第一職代○○日請假或出差，
                     '期間由第二職代代理」訊息；同仁若無第二職代者，系統會彈訊息「第一職代○○日請假或出差，
                     '請自行加設第二職代」。
                     If dblDay > 2 Then
                        If CheckIsPersonRestSectorSame(CStr(strTemp(i)), txtB1004, Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), txtB1006, Trim(txtB1007_1.Text & ":" & txtB1007_2.Text), txtB1001) = False Then
                           If Not IsNull(strTemp(i)) Then
                              For k = 0 To cboEmp.UBound
                                 If cboEmp(k) = "" Then
                                    cboEmp(k) = strTemp(i)
                                    cboEmp(k).Enabled = False
                                    Call CboEmp_LostFocus(k)
                                    Exit For
                                 End If
                              Next k
                           End If
                           '有請假,其他職代也要簽假單
                           If CheckIsPersonRestSector(CStr(strTemp(i)), txtB1004, Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), txtB1006, Trim(txtB1007_1.Text & ":" & txtB1007_2.Text), txtB1001) = True Then
                              bolReadAllPer = True
                           End If
                        Else
                           '完全相同的請假區間就略過,讀下一個職代
                           bolReadAllPer = True
                        End If
                     Else
                        '1.同仁請假或出差在二天(含)以內，請假的職代系統維持不變，即第一職代如同時請假或出差，
                        '不論重疊時間長短，均由第二職代簽職代；若第一、二職同時請假或出差，則由請假或出差同仁自設其他職代。
                        '*不可同請假區間不可今日休假
                        If CheckIsPersonRestSector(CStr(strTemp(i)), txtB1004, Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), txtB1006, Trim(txtB1007_1.Text & ":" & txtB1007_2.Text), txtB1001) = False And _
                           CheckIsPersonRest(CStr(strTemp(i)), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
                           If Not IsNull(strTemp(i)) Then
                              For k = 0 To cboEmp.UBound
                                 If cboEmp(k) = "" Then
                                    cboEmp(k) = strTemp(i)
                                    cboEmp(k).Enabled = False
                                    Call CboEmp_LostFocus(k)
                                    Exit For
                                 End If
                              Next k
                           End If
                        End If
                     End If
                  Next i
                  If cboEmp(0) <> "" And bolReadAllPer = False Then Exit For '若有取得職代,則離開迴圈
               End If
            Next j
         End If
      End If
      '各部門特殊職代規則
      strSql = "select * from ABS003 where B0305='Y' and B0303='" & IIf(strA0925 <> "", strA0925, strA0911) & "' and " & dblDay & ">B0302 " & strCon
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         '會簽同部門其他人員-排除畫面上已填入的職代及審核主管
         strCon = ""
         For i = 0 To cboEmp.UBound
            If cboEmp(i) <> "" Then strCon = strCon & "'" & Left(cboEmp(i), 5) & "',"
         Next i
         For i = 0 To CboBoss.UBound
            If CboBoss(i) <> "" Then strCon = strCon & "'" & Left(CboBoss(i), 5) & "',"
         Next i
         '排除當事人
         If txtB1003 <> "" Then strCon = strCon & "'" & txtB1003 & "',"
         If strCon <> "" Then
            strCon = Left(strCon, Len(strCon) - 1)
            strCon = "and st01 not in(" & strCon & ") "
         End If
         'Modify By Sindy 2014/7/22 Pub_StrUserSt03="M10"總務處主管時,要抓M11總務處人員出來
         'Modified by Lydia 2017/03/28 ST14改成多個編號 st14<>'99997'=> instr(st14,'99997')=0
         'Modify By Sindy 2023/12/20
         If strSrvDate(1) >= 新部門啟用日 Then
            strSql = "select st01 from staff " & _
                     "where st93='" & Pub_StrUserSt93 & "' and st04='1' " & strCon & _
                     "and (instr(st14,'99997')=0 or st14 is null) " & _
                     "and substr(st01,4,1)<>'9' " & _
                     "order by st01 asc"
         Else
         '2023/12/20 END
            strSql = "select st01 from staff " & _
                     "where st03='" & IIf(Pub_StrUserSt03 = "M10", "M11", Pub_StrUserSt03) & "' and st04='1' " & strCon & _
                     "and (instr(st14,'99997')=0 or st14 is null) " & _
                     "and substr(st01,4,1)<>'9' " & _
                     "order by st01 asc"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               .MoveFirst
               Do While Not .EOF
                  If Not IsNull(RsTemp.Fields(0)) Then
                     If Left(cboEmp(0), 5) <> RsTemp.Fields(0) And _
                        Left(cboEmp(1), 5) <> RsTemp.Fields(0) And _
                        Left(cboEmp(2), 5) <> RsTemp.Fields(0) Then
                        For k = 0 To cboEmp.UBound
                           If cboEmp(k) = "" Then
                              cboEmp(k) = RsTemp.Fields(0)
                              'Add By Sindy 2014/7/22
                              'CboEmp(k).Enabled = False
                              cboEmp(k).Enabled = True
                              '2014/7/22 END
                              Call CboEmp_LostFocus(k)
                              Exit For
                           End If
                        Next k
                     End If
                  End If
                  .MoveNext
               Loop
            End With
         End If
      End If
   End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 2
      GRD1.row = dblPrevRow
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   dblPrevRow = GRD1.row
   For i = 0 To GRD1.Cols - 1
      GRD1.col = i
      GRD1.CellBackColor = &HFFC0C0
   Next i
End If
GRD1.Visible = True
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("B1022")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1022")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("B1022"), True)
      End If
   End If
   m_B1023 = ""
   If IsNull(rsSrcTmp.Fields("B1023")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1023")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("B1023"))
         strCDate = Format(strTemp, "###/##/##")
         m_B1023 = rsSrcTmp.Fields("B1023")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1024")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1024")) = False Then
         strTemp = rsSrcTmp.Fields("B1024")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1025")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1025")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("B1025"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1026")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1026")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("B1026"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1027")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1027")) = False Then
         strTemp = rsSrcTmp.Fields("B1027")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label26.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Sub cmdSend_Click()
Dim strUpdDate As String, strUpdTime As String
Dim bolConn As Boolean
   
On Error GoTo ErrHand

   '檢查條件
   If TxtValidate = False Then Exit Sub

   Screen.MousePointer = vbHourglass
   
   cmdSend.Enabled = False
   cnnConnection.BeginTrans: bolConn = True
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   If txtB1001 = "" Then
      '表單編號自動給號
      txtB1001 = AutoNo_ABS("ABS", 5)
      
      '檢查是否還有自動給號資料不一致的問題
      strSql = "select AU03 from autonumber where AU01='ABS'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Val(RsTemp.Fields("AU03")) <> Val(Right(txtB1001, Len(txtB1001.Text) - 3)) Then
            MsgBox "系統自動給號(" & txtB1001 & ")更新有誤，請洽電腦中心！", vbInformation, "系統錯誤"
            txtB1001 = ""
            GoTo ErrHand
            Exit Sub
         End If
      End If
   End If
   
   If SaveABS010() = False Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   If SaveABS011() = False Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   'Add By Sindy 2014/2/13 檢查是否有連續假單
   cnnConnection.CommitTrans: bolConn = False
   'Modify By Sindy 2020/5/27 + , IIf(Left(CboB1002, 2) = 表單類別_請假, Left(CboB1008, 2), txtB1014)
   If PUB_ChkSerialRest(txtB1001, CboB1002.Text, txtB1003, , IIf(Left(CboB1002, 2) = 表單類別_請假, Left(CboB1008, 2), txtB1014)) = True Then
      Screen.MousePointer = vbDefault
      QueryData
      Exit Sub
   Else
      cnnConnection.BeginTrans: bolConn = True
   End If
   
   'Add By Sindy 2022/11/18 嘉渝說有關主管代填假單，當事人入所時再修改假單送出時，增加當事人送出記錄
   If m_B1018 = 主管代填 Then
      strSql = GetInsertABS012Sql(Trim(txtB1001), strUserNum, strUpdDate, strUpdTime, "", "為主管代填的假單，修改後送出")
      cnnConnection.Execute strSql
   End If
   '2022/11/18 END
   
'   '送呈下一處理人員
'   If GetSendNextPerson(Trim(txtB1001), Trim(txtB1003), m_B1017, strUserNum) = False Then GoTo ErrHand
   '讀取下一處理人員
   If GetNextProPerson(Trim(txtB1001), Trim(txtB1003), m_B1017, strUserNum) = False Then GoTo ErrHand
   
   'Add By Sindy 2012/2/13
   '檢查表單職代是否符合人事職代裡的設定資料
   If ChkIsDutyAgent(Trim(txtB1001), Trim(txtB1003)) = False Then
      strSql = GetInsertABS012Sql(Trim(txtB1001), strUserNum, strUpdDate, strUpdTime, "01", "自行填寫職代")
      cnnConnection.Execute strSql
   End If
   '2012/2/13 End
   
   cnnConnection.CommitTrans: bolConn = False
   
   '發E-Mail通知下一處理人員
   strContent = GetEMailContent(txtB1001, strSubject)
   PUB_SendMail strUserNum, m_B1017, "", strSubject, strContent, , , , , , , , , , True
   
   Screen.MousePointer = vbDefault
   
   cmdSend.Enabled = True
   Call cmdExit_Click '結束
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cmdSend.Enabled = True
   If bolConn = True Then
      cnnConnection.RollbackTrans
   End If
   MsgBox " 送出失敗！" & vbCrLf & Err.Description
End Sub

'更新出缺勤電子簽核主檔
Private Function SaveABS010() As Boolean
Dim strB1008 As String, strB1014 As String, strB1015 As String
Dim strB1006 As String, strB1009 As String, strB1010 As String
Dim strB1012 As String, strB1013 As String
Dim strErrText As String
   
On Error GoTo ErrHand
   
   strErrText = ""
   SaveABS010 = True
   
   '假別
   If CboB1008.Visible = True Then
      strB1008 = Left(CboB1008, 2)
   End If
   '迄止日期
   If txtB1006.Visible = True Then
      strB1006 = DBDATE(txtB1006)
   End If
   '日,時
   If Frame01.Visible = True Then
      strB1009 = txtB1009
      strB1010 = txtB1010
   End If
   '時數-平日,假日
   If Left(CboB1002, 2) = 表單類別_加班 Then
'      If txtB1012 <> "" Then strB1012 = txtB1012
'      If txtB1013 <> "" Then strB1013 = txtB1013
      'Modify By Sindy 2016/12/26
      'Modify By Sindy 2019/10/3 不需檢查颱風假
      'If ChkWorkDay(ChangeTStringToWString(txtB1004), txtB1003, True) = False Then '假日
      If ChkWorkDay(ChangeTStringToWString(txtB1004), txtB1003) = False Then '假日
      '2019/10/3 END
         strB1012 = ""
         strB1013 = txtB101213
      Else
         strB1012 = txtB101213
         strB1013 = ""
      End If
      '2016/12/26 END
   End If
   '差程,地點
   If Left(CboB1002, 2) = 表單類別_出差 Then
      strB1014 = txtB1014
      strB1015 = txtB1015
   End If
   
   strSql = "SELECT * FROM ABS010 WHERE B1001=" & CNULL(txtB1001)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strSql = "SELECT * FROM ABS010 WHERE B1001=" & CNULL(txtB1001) & " and B1003=" & CNULL(txtB1003)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         '修改
         'Modify By Sindy 2016/12/26 + ,B1030= " & CNULL(txtB1030)
         strSql = "update ABS010 set " & _
                  "B1002= " & CNULL(Left(CboB1002, 2)) & _
                  ",B1004= " & CNULL(DBDATE(txtB1004)) & _
                  ",B1005= " & CNULL(txtB1005_1 & Format("00" & txtB1005_2, "00")) & _
                  ",B1006= " & CNULL(strB1006) & _
                  ",B1007= " & CNULL(txtB1007_1 & Format("00" & txtB1007_2, "00")) & _
                  ",B1008= " & CNULL(strB1008) & _
                  ",B1009= " & CNULL(strB1009) & _
                  ",B1010= " & CNULL(strB1010) & _
                  ",B1011= " & CNULL(ChgSQL(Trim(txtB1011))) & _
                  ",B1012= " & CNULL(strB1012) & _
                  ",B1013= " & CNULL(strB1013) & _
                  ",B1014= " & CNULL(strB1014) & _
                  ",B1015= " & CNULL(strB1015) & _
                  ",B1028= " & CNULL(IIf(Frame1.Visible = True And cboSTime.Text <> "", Format(cboSTime.Text, "hhmm"), "")) & _
                  ",B1029= " & CNULL(IIf(Frame1.Visible = True And cboETime.Text <> "", Format(cboETime.Text, "hhmm"), "")) & _
                  ",B1030= " & CNULL(txtB1030) & _
                  " where B1001=" & CNULL(txtB1001)
      Else
         '表單編號已存在,但該表單人員並非目前操作人員
         strErrText = "表單編號已存在,但該表單人員並非目前操作人員！"
         GoTo ErrHand
      End If
   Else
      '新增
      'Modify By Sindy 2016/12/26 + ,B1030
      strSql = "insert into ABS010(B1001,B1002,B1003,B1004,B1005,B1006,B1007,B1008,B1009,B1010,B1011,B1012,B1013,B1014,B1015,B1016,B1017,B1018,B1028,B1029,B1030) " & _
               "values(" & CNULL(txtB1001) & "," & CNULL(Left(CboB1002, 2)) & "," & CNULL(txtB1003) & "," & _
               CNULL(DBDATE(txtB1004)) & "," & CNULL(txtB1005_1 & Format("00" & txtB1005_2, "00")) & "," & CNULL(strB1006) & "," & _
               CNULL(txtB1007_1 & Format("00" & txtB1007_2, "00")) & "," & CNULL(strB1008) & "," & CNULL(strB1009) & "," & _
               CNULL(strB1010) & "," & CNULL(ChgSQL(Trim(txtB1011))) & "," & CNULL(strB1012) & "," & _
               CNULL(strB1013) & "," & CNULL(strB1014) & "," & CNULL(strB1015) & "," & _
               CNULL(txtB1003) & "," & CNULL(m_B1017) & "," & CNULL(m_B1018) & "," & _
               CNULL(IIf(Frame1.Visible = True And cboSTime.Text <> "", Format(cboSTime.Text, "hhmm"), "")) & "," & _
               CNULL(IIf(Frame1.Visible = True And cboETime.Text <> "", Format(cboETime.Text, "hhmm"), "")) & "," & _
               CNULL(txtB1030) & ")"
   End If
   cnnConnection.Execute strSql
   Exit Function
   
ErrHand:
   SaveABS010 = False
   cnnConnection.RollbackTrans
   MsgBox " 更新ABS010失敗！" & vbCrLf & strErrText & Err.Description
End Function

'更新表單簽核檔
Private Function SaveABS011() As Boolean

On Error GoTo ErrHand
   
   SaveABS011 = True
   
   strSql = "delete From ABS011 where B1101=" & CNULL(txtB1001)
   cnnConnection.Execute strSql
   'If strType = "1" Then '送出
      '1.職代
      For i = 0 To cboEmp.UBound
         If cboEmp(i) <> "" And cboEmp(i).Visible = True Then
            strSql = "insert into ABS011 (B1101,B1102,B1103,B1104,B1108) values(" & CNULL(txtB1001) & ",'1'," & (i + 1) & "," & CNULL(Left(cboEmp(i), 5)) & ",'(代" & GetPersonSeqno(txtB1003, Left(cboEmp(i), 5)) & ")')"
            cnnConnection.Execute strSql
         End If
      Next i
      '2.審核主管
      For i = 0 To CboBoss.UBound
         If CboBoss(i) <> "" Then
            'Modify By Sindy 2015/11/13 + B1109
            strSql = "insert into ABS011 (B1101,B1102,B1103,B1104,B1109) values(" & CNULL(txtB1001) & ",'2'," & (i + 1) & "," & CNULL(Left(CboBoss(i), 5)) & "," & CNULL(Left(CboBoss(i), 5)) & ")"
            cnnConnection.Execute strSql
         End If
      Next i
'   Else '重送
'      strSql = "update ABS011 set " & _
'               "B1105=null" & _
'               ",B1106=null" & _
'               ",B1107=null" & _
'               " where B1101=" & CNULL(txtB1001)
'      cnnConnection.Execute strSql
'   End If
   
   Exit Function
   
ErrHand:
   SaveABS011 = False
   cnnConnection.RollbackTrans
   MsgBox " 新增ABS011失敗！" & vbCrLf & Err.Description
End Function

Private Sub cmdagainSend_Click()
Dim strOldB1017 As String
Dim strUpdDate As String, strUpdTime As String, strB1207 As String
Dim strText As String
Dim strEmp As String, strBoss As String, strTemp As Variant 'Add By Sindy 2016/7/7
   
On Error GoTo ErrHand
   
   'Modify By Sindy 2015/12/25 發生人員重送只因職代設定錯誤,而被退回再重送;
   '                           因此在此處增加重新讀取職代及簽核主管資料
   If Left(CboB1002, 2) = 表單類別_加班 Then
      'Modify By Sindy 2016/12/26
'      If Val(IIf(txtB1012 = "", 0, txtB1012)) > 0 Then dblDay = Val(txtB1012) * 0.1
'      If Val(IIf(txtB1013 = "", 0, txtB1013)) > 0 Then dblDay = Val(txtB1013) * 0.1
      If Val(txtB101213) > 0 Then
         dblDay = Val(txtB101213) * 0.1
      End If
   Else
      dblDay = Val(txtB1009) + (Val(txtB1010) * 0.1)
   End If
   'Add By Sindy 2016/7/7 修改時,記錄人員自行輸入的職代和審核主管
   strEmp = "": strBoss = ""
   If Me.cmdModify.Enabled = True Then
      For i = 0 To cboEmp.UBound
         If cboEmp(i).Enabled = True And Trim(cboEmp(i).Text) <> "" Then
            strEmp = strEmp & ";" & Trim(cboEmp(i).Text)
         End If
      Next i
      If strEmp <> "" Then strEmp = Mid(strEmp, 2)
      For i = 0 To CboBoss.UBound
         If CboBoss(i).Enabled = True And Trim(CboBoss(i).Text) <> "" Then
            strBoss = strBoss & ";" & Trim(CboBoss(i).Text)
         End If
      Next i
      If strBoss <> "" Then strBoss = Mid(strBoss, 2)
   End If
   '2016/7/7 END
   Call ClearFieldCbo '先清空欄位值
   Call GetPersonBossData(dblDay) '讀取職代及審核主管
   '2015/12/25 END
   'Add By Sindy 2016/7/7 修改時,記錄人員自行輸入的職代和審核主管
   If strEmp <> "" Then
      strTemp = Split(strEmp, ";")
      For i = 0 To UBound(strTemp)
         For j = 0 To cboEmp.UBound
            If Trim(cboEmp(j).Text) <> "" Then
               If Left(Trim(cboEmp(j).Text), 5) = Left(Trim(strTemp(i)), 5) Then
                  Exit For
               End If
            Else
               cboEmp(j).Text = strTemp(i)
               Exit For
            End If
         Next j
      Next i
   End If
   If strBoss <> "" Then
      strTemp = Split(strBoss, ";")
      For i = 0 To UBound(strTemp)
         For j = 0 To CboBoss.UBound
            If Trim(CboBoss(j).Text) <> "" Then
               If Left(Trim(CboBoss(j).Text), 5) = Left(Trim(strTemp(i)), 5) Then
                  Exit For
               End If
            Else
               CboBoss(j).Text = strTemp(i)
               Exit For
            End If
         Next j
      Next i
   End If
   '2016/7/7 END
   
   strOldB1017 = m_B1017 '記錄原下一處理人員
   '檢查條件
   If TxtValidate = False Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   cmdagainSend.Enabled = False
   cnnConnection.BeginTrans
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   '表單退回到當事人時,重送
   If cmdModify.Enabled = True Then
      If SaveABS011() = False Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      If SaveABS010() = False Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      
'      '送呈下一處理人員
'      If GetSendNextPerson(Trim(txtB1001), Trim(txtB1003), m_B1017, strUserNum) = False Then GoTo ErrHand
      '讀取下一處理人員
      If GetNextProPerson(Trim(txtB1001), Trim(txtB1003), m_B1017, strUserNum) = False Then GoTo ErrHand
      
      '記錄10.重送訊息
      strSql = GetInsertABS012Sql(Trim(txtB1001), strUserNum, strUpdDate, strUpdTime, "07", "")
      cnnConnection.Execute strSql
      
      '發E-Mail通知下一處理人員
      strContent = GetEMailContent(txtB1001, strSubject, 重送)
      PUB_SendMail strUserNum, m_B1017, "", strSubject, strContent, , , , , , , , , , True
      
   '表單簽核中,異動簽核人員(卡單)
   Else
      '2.審核主管
      For i = CboBoss.UBound To 0 Step -1
         If CboBoss(i).Enabled = True Then '欄位為可使用狀態
            If m_cboBoss(i) <> Left(Trim(CboBoss(i)), 5) Then '比對原資料與目前資料是否相同,不同時才更新
               If m_cboBoss(i) <> "" Then '原始資料有值時,先刪除
                  strSql = "delete From ABS011 where B1101=" & CNULL(txtB1001) & " and B1102='2' and B1104=" & CNULL(m_cboBoss(i)) & " and B1107 is null "
                  cnnConnection.Execute strSql
               End If
               If Left(CboBoss(i), 5) <> "" Then '新增目前資料
                  'Modify By Sindy 2015/11/13 + B1109
                  strSql = "insert into ABS011 (B1101,B1102,B1103,B1104,B1109) values(" & CNULL(txtB1001) & ",'2'," & (i + 1) & "," & CNULL(Left(CboBoss(i), 5)) & "," & CNULL(Left(CboBoss(i), 5)) & ")"
                  cnnConnection.Execute strSql
               End If
            End If
         End If
      Next i
      '1.職代
      For i = cboEmp.UBound To 0 Step -1
         If cboEmp(i).Enabled = True Then '欄位為可使用狀態
            If m_cboEmp(i) <> Left(Trim(cboEmp(i)), 5) Then '比對原資料與目前資料是否相同,不同時才更新
               If m_cboEmp(i) <> "" Then '原始資料有值時,先刪除
                  strSql = "delete From ABS011 where B1101=" & CNULL(txtB1001) & " and B1102='1' and B1104=" & CNULL(m_cboEmp(i)) & " and B1107 is null "
                  cnnConnection.Execute strSql
               End If
               If Left(cboEmp(i), 5) <> "" And cboEmp(i).Visible = True Then '新增目前資料
                  strSql = "insert into ABS011 (B1101,B1102,B1103,B1104,B1108) values(" & CNULL(txtB1001) & ",'1'," & (i + 1) & "," & CNULL(Left(cboEmp(i), 5)) & ",'(代" & GetPersonSeqno(txtB1003, Left(cboEmp(i), 5)) & ")')"
                  cnnConnection.Execute strSql
               End If
            End If
         End If
      Next i
      
'      '送呈下一處理人員
'      If GetSendNextPerson(Trim(txtB1001), Trim(txtB1003), m_B1017, strUserNum) = False Then GoTo ErrHand
      '讀取下一處理人員
      If GetNextProPerson(Trim(txtB1001), Trim(txtB1003), m_B1017, strUserNum) = False Then GoTo ErrHand
      
      If strOldB1017 <> m_B1017 Then
         If ChkStaffST04(strOldB1017, False) = True Then
            strText = "人員離職，"
         Else
            'strText = "人員休假，"
            strText = ""
         End If
         
         '記錄10.重送訊息
         strB1207 = strText & "更改簽核人員" & GetPrjSalesNM(strOldB1017) & "->" & GetPrjSalesNM(m_B1017)
         strSql = GetInsertABS012Sql(Trim(txtB1001), strUserNum, strUpdDate, strUpdTime, "07", strB1207)
         cnnConnection.Execute strSql
         
         '有異動下一處理人員時,發E-Mail通知下一處理人員及原下一處理人員
         If m_B1017 <> "" Then
            strContent = GetEMailContent(txtB1001, strSubject)
            PUB_SendMail strUserNum, m_B1017, "", strSubject, strContent, , , , , , , , , , True
         End If
         If strOldB1017 <> "" And strText <> "人員離職" Then
            strContent = GetEMailContent(txtB1001, strSubject, 重送更改通知, "，" & strText & "更改簽核人員為" & GetPrjSalesNM(m_B1017))
            PUB_SendMail strUserNum, strOldB1017, "", strSubject, strContent, , , , , , , , , , True
         End If
      End If
   End If
   
   cnnConnection.CommitTrans
   
   Screen.MousePointer = vbDefault
   cmdagainSend.Enabled = True
   Call cmdExit_Click '結束
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cmdagainSend.Enabled = True
   cnnConnection.RollbackTrans
   MsgBox " 重送失敗！" & vbCrLf & Err.Description
End Sub

Public Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim strST20 As String
Dim intBossNum As Integer
Dim intMaxi As Integer

TxtValidate = False
PUB_FilterFormText Me 'Add by Sindy 2011/10/14 修正畫面所有含跳行符號的文字框

If CboB1002.Text = "" Then
    MsgBox "表單類別不可以空白！", vbExclamation
    CboB1002.SetFocus
    Exit Function
End If

'Add By Sindy 2014/10/7
If DBDATE(txtB1004) > strSysDtAft3M Then
   If MsgBox("表單起迄日期超過3個月，是否確定送出？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
      Exit Function
   End If
End If
'2014/10/7 End
'Add By Sindy 2018/3/12
If Left(CboB1002, 2) = 表單類別_請假 Or Left(CboB1002, 2) = 表單類別_加班 Then
   strSql = "select BR01 from BookRecord " & _
             "where BR01='" & Left(DBDATE(txtB1004), 6) & "'"
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If MsgBox("此日期薪資已發放，是否確定送出？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
         Exit Function
      End If
   End If
End If
'2018/3/12 END

'Add By Sindy 2011/10/17
'Modify By Sindy 2024/12/25 +補休
If CboB1008.Visible = True And (Left(CboB1008.Text, 2) = "08" Or Left(CboB1008.Text, 2) = "14") Then
   '特別假之外，其他均須填事由包括加班、出差
   'Add By Sindy 2017/11/3
   If Val(txtB1010.Text) <> 4 And Val(txtB1010.Text) > 0 Then
      MsgBox IIf(Left(CboB1008.Text, 2) = "08", "特別假", "補休") & "只能請4小時或整日!!!", vbExclamation + vbOKOnly
      If txtB1007_1.Enabled = True Then txtB1007_1.SetFocus
      Exit Function
   End If
   '2017/11/3 END
Else
   'Modify By Sindy 2018/8/28 +Trim
   If Trim(txtB1011.Text) = "" Then
       MsgBox "需載明事由！", vbExclamation
       txtB1011.SetFocus
       Exit Function
   End If
End If

If Left(CboB1002, 2) = 表單類別_出差 Then
   If txtB1014.Enabled = True Then
      If txtB1014.Text = "" Then
         MsgBox "差程不可以空白！", vbExclamation
         txtB1014.SetFocus
         Exit Function
      End If
   End If
'   If txtB1015.Enabled = True Then
'      If txtB1015.Text = "" Then
'         MsgBox "出差地點不可以空白！", vbExclamation
'         txtB1015.SetFocus
'         Exit Function
'      End If
'   End If
End If
If txtB1004.Text = "" Then
    MsgBox "日期起不可以空白！", vbExclamation
    txtB1004.SetFocus
    Exit Function
End If
If txtB1005_1.Text = "" Or txtB1005_1.Text = "00" Then
    MsgBox "必須輸入起始(時)！", vbExclamation
    txtB1005_1.SetFocus
    Exit Function
End If
If txtB1006.Visible = True Then
   If txtB1006.Text = "" Then
       MsgBox "日期迄不可以空白！", vbExclamation
       txtB1006.SetFocus
       Exit Function
   End If
End If
If txtB1007_1.Text = "" Or txtB1007_1.Text = "00" Then
    MsgBox "必須輸入迄止(時)！", vbExclamation
    txtB1007_1.SetFocus
    Exit Function
End If
'檢查起迄日期時間區間是否有重覆
If m_EditMode = 1 Or m_EditMode = 2 Then
   If Left(CboB1002, 2) = 表單類別_加班 Then
      If IsRecordExist(txtB1003, DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1004), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
         txtB1004.SetFocus
         Exit Function
      End If
   Else
      If IsRecordExist(txtB1003, DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1006), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
         txtB1006.SetFocus
         Exit Function
      End If
   End If
End If

'Add By Sindy 2011/10/17 增加判斷員工代號+日期是否人員已離職
If ChkStaffST04(txtB1003, True, txtB1004) = True Then
   txtB1004.SetFocus
   Exit Function
End If
If txtB1006.Visible = True Then
   If ChkStaffST04(txtB1003, True, txtB1006) = True Then
      txtB1006.SetFocus
      Exit Function
   End If
End If

If Me.txtB1004.Enabled = True Then
   Cancel = False
   txtB1004_Validate Cancel
   If Cancel = True Then
      txtB1004.SetFocus
      Exit Function
   End If
End If
If Me.txtB1005_1.Enabled = True Then
   Cancel = False
   txtB1005_1_Validate Cancel
   If Cancel = True Then
      txtB1005_1.SetFocus
      Exit Function
   End If
End If
If Me.txtB1005_2.Enabled = True Then
   Cancel = False
   txtB1005_2_Validate Cancel
   If Cancel = True Then
      txtB1005_2.SetFocus
      Exit Function
   End If
End If
If txtB1006.Visible = True Then
   If Me.txtB1006.Enabled = True Then
      Cancel = False
      txtB1006_Validate Cancel
      If Cancel = True Then
         txtB1006.SetFocus
         Exit Function
      End If
   End If
End If
If Me.txtB1007_1.Enabled = True Then
   Cancel = False
   txtB1007_1_Validate Cancel
   If Cancel = True Then
      txtB1007_1.SetFocus
      Exit Function
   End If
End If
If Me.txtB1007_2.Enabled = True Then
   Cancel = False
   txtB1007_2_Validate Cancel
   If Cancel = True Then
      txtB1007_2.SetFocus
      Exit Function
   End If
End If

If CboB1008.Visible = True Then
   If CboB1008.Text = "" Then
       MsgBox "假別不可以空白！", vbExclamation
       CboB1008.SetFocus
       Exit Function
   End If
   If Me.CboB1008.Enabled = True Then
      Cancel = False
      CboB1008_Validate Cancel
      If Cancel = True Then
         CboB1008.SetFocus
         Exit Function
      End If
   End If
End If

''Add By Sindy 2017/11/14 計算日,時
'If GetCountDayHour(False) = False Then
'   Exit Function
'End If

'Add By Sindy 2013/10/18
If PUB_bWkSpec = False And Left(CboB1002, 2) = 表單類別_請假 Then
   '起始時間不可填中午休息時間 1210~1329
   'Modify By Sindy 2017/4/17 中午休息時間 1200~1329
   If strSrvDate(1) >= 中午休息時間改1200 Then
      If Val(txtB1005_1 & txtB1005_2) >= 1200 And Val(txtB1005_1 & txtB1005_2) <= 1329 Then
         MsgBox "起始時間不可填中午休息時間！", vbExclamation + vbOKOnly
         txtB1005_1.SetFocus
         Exit Function
      End If
      '迄止時間不可填中午休息時間 1211~1330
      '                           1201~1330
      If Val(txtB1007_1 & txtB1007_2) >= 1201 And Val(txtB1007_1 & txtB1007_2) <= 1330 Then
         MsgBox "迄止時間不可填中午休息時間！", vbExclamation + vbOKOnly
         txtB1007_1.SetFocus
         Exit Function
      End If
   Else
   '2017/4/17 END
      If Val(txtB1005_1 & txtB1005_2) >= 1210 And Val(txtB1005_1 & txtB1005_2) <= 1329 Then
         MsgBox "起始時間不可填中午休息時間！", vbExclamation + vbOKOnly
         txtB1005_1.SetFocus
         Exit Function
      End If
      '迄止時間不可填中午休息時間 1211~1330
      If Val(txtB1007_1 & txtB1007_2) >= 1211 And Val(txtB1007_1 & txtB1007_2) <= 1330 Then
         MsgBox "迄止時間不可填中午休息時間！", vbExclamation + vbOKOnly
         txtB1007_1.SetFocus
         Exit Function
      End If
   End If
End If
'2013/10/18 END

'Add By Sindy 2015/1/5 檢查請假日,時是否為空的
If Left(CboB1002, 2) <> 表單類別_加班 Then
   If txtB1009.Text = "" Then txtB1009.Text = "0"
   If txtB1010.Text = "" Then txtB1010.Text = "0"
   If Val(txtB1009.Text) = 0 And Val(txtB1010.Text) = 0 Then
      If PUB_bWkSpec = False Then
         MsgBox "無時數！", vbExclamation
      Else
         MsgBox "請輸入時數！", vbExclamation
         txtB1009.SetFocus
      End If
      Exit Function
   End If
End If

If Me.Frame1.Visible = True Then
   'Modify By Sindy 2011/11/16
   'Modify By Sindy 2012/5/25 +and Left(CboB1002, 2) = 表單類別_請假
   If cboSTime <> "" And Left(CboB1002, 2) = 表單類別_請假 Then
      If Val(Format(cboSTime.Text, "hhmm")) > Val(Right("00" & txtB1005_1, 2) & Right("00" & txtB1005_2, 2)) Then
         MsgBox "起日請假時間必須大於或等於起日上班時段!", vbExclamation + vbOKOnly
         txtB1005_1.SetFocus
         Exit Function
      End If
   End If
   '2011/11/16 End
   Cancel = False
   cboSTime_Validate Cancel
   If Cancel = True Then
      cboSTime.SetFocus
      Exit Function
   End If
   
   'Modify By Sindy 2011/11/16
   'Modify By Sindy 2012/5/25 +and Left(CboB1002, 2) = 表單類別_請假
   If cboETime <> "" And Left(CboB1002, 2) = 表單類別_請假 Then
      If Val(Format(cboETime.Text, "hhmm")) < Val(Right("00" & txtB1007_1, 2) & Right("00" & txtB1007_2, 2)) Then
         MsgBox "迄日請假時間必須小於或等於迄日下班時段!", vbExclamation + vbOKOnly
         txtB1007_1.SetFocus
         Exit Function
      End If
   End If
   '2011/11/16 End
   Cancel = False
   cboETime_Validate Cancel
   If Cancel = True Then
      cboETime.SetFocus
      Exit Function
   End If

   If txtB1009.Text = "" Or txtB1010.Text = "" Or _
      (txtB1009.Text = "0" And txtB1010.Text = "0") Then
       MsgBox "無時數！", vbExclamation
'       txtB1009.SetFocus
       Exit Function
   End If
   If Me.txtB1009.Enabled = True Then
      Cancel = False
      txtB1009_Validate Cancel
      If Cancel = True Then
'         txtB1009.SetFocus
         Exit Function
      End If
   End If
   If Me.txtB1010.Enabled = True Then
      Cancel = False
      txtB1010_Validate Cancel
      If Cancel = True Then
'         txtB1010.SetFocus
         Exit Function
      End If
   End If
   If Left(CboB1002, 2) = 表單類別_請假 Then
      'Modify By Sindy 2015/2/10 人事已先行作業時,則不需要再檢查特別假天數
      If CboB1008.Enabled = True Then
      '2015/2/10 END
         If ChkSA06_08(txtB1009, txtB1010, txtB1003, txtB1004, txtB1005_1, txtB1005_2, txtB1006, txtB1007_1, txtB1007_2, CboB1008, 0, , txtB1001) = False Then
            CboB1008.Text = ""
            CboB1008.SetFocus
            Exit Function
         End If
         'Add By Sindy 2014/12/31 +健檢假
         If ChkSA06_23(txtB1009, txtB1010, txtB1003, txtB1004, txtB1005_1, txtB1005_2, txtB1006, txtB1007_1, txtB1007_2, CboB1008, 0, 0) = False Then
            CboB1008.Text = ""
            CboB1008.SetFocus
            Exit Function
         End If
         'Add By Sindy 2024/12/10 檢查可補休
         If ChkSA06_14(txtB1009, txtB1010, txtB1003, txtB1004, txtB1005_1, txtB1005_2, txtB1006, txtB1007_1, txtB1007_2, CboB1008, 0, , txtB1001) = False Then
            CboB1008.Text = ""
            CboB1008.SetFocus
            Exit Function
         End If
      End If
   End If
End If

'Add By Sindy 2012/1/6 陪產假檢查天數
'Modify By Sindy 2015/4/2 陪產假3天改為可休5休
'Modify By Sindy 2022/10/31 + 劉柏翰:請將陪產假之控管由5日調整為7日
If Left(CboB1008.Text, 2) = "19" Then '19.陪產假
'   If Val(txtB1009) + (Val(txtB1010) * 0.1) > 3 Then
'       MsgBox "陪產假不可超過3天！", vbExclamation
'       Exit Function
'   End If
   'If Val(txtB1009) + (Val(txtB1010) * 0.1) > 5 Then
   If Val(txtB1009) + (Val(txtB1010) * 0.1) > 7 Then
       MsgBox "陪產假不可超過 7 天！", vbExclamation
       Exit Function
   End If
   strSql = "select sum(nvl(SA07,0))+(sum(nvl(SA08,0))/" & PUB_intWkHour & ") from staff_Absence " & _
             "where SA01='" & txtB1003 & "' and SA06='19' " & _
               "and SA02 between " & Left(strSrvDate(1), 4) & "0101 and " & Left(strSrvDate(1), 4) & "1231"
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
'      If Val("" & adoRecordset.Fields(0)) > 3 Then
'          MsgBox "同一年中陪產假不可超過3天！", vbExclamation
'          Exit Function
'      End If
      'If Val("" & adoRecordset.Fields(0)) > 5 Then
      If Val("" & adoRecordset.Fields(0)) > 7 Then
          MsgBox "同一年中陪產假不可超過 7 天！", vbExclamation
          Exit Function
      End If
   End If
End If
'Add By Sindy 2015/4/2 產檢假可休5天
'Modify By Sindy 2021/10/25 + 劉柏翰:請將產檢假之控管由5日調整為7日
If Left(CboB1008.Text, 2) = "21" Then '21.產檢假
   If Val(txtB1009) + (Val(txtB1010) * 0.1) > 7 Then
       MsgBox "產檢假不可超過 7 天！", vbExclamation
       Exit Function
   End If
   strSql = "select sum(nvl(SA07,0))+(sum(nvl(SA08,0))/" & PUB_intWkHour & ") from staff_Absence " & _
             "where SA01='" & txtB1003 & "' and SA06='21' " & _
               "and SA02 between " & Left(strSrvDate(1), 4) & "0101 and " & Left(strSrvDate(1), 4) & "1231"
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Val("" & adoRecordset.Fields(0)) > 7 Then
          MsgBox "同一年中產檢假不可超過 7 天！", vbExclamation
          Exit Function
      End If
   End If
End If

If Left(CboB1002, 2) = 表單類別_加班 Then
'   If Me.txtB1012.Enabled = True Then
'      Cancel = False
'      txtB1012_Validate Cancel
'      If Cancel = True Then
''         txtB1012.SetFocus
'         Exit Function
'      End If
'   End If
'   If Me.txtB1013.Enabled = True Then
'      Cancel = False
'      txtB1013_Validate Cancel
'      If Cancel = True Then
''         txtB1013.SetFocus
'         Exit Function
'      End If
'   End If
   'Modify By Sindy 2016/12/26
   If Me.txtB1030.Enabled = True Then
      Cancel = False
      txtB1030_Validate Cancel
      If Cancel = True Then
'         txtB1030.SetFocus
         Exit Function
      End If
   End If
   
   'Add By Sindy 2018/4/24
   If Val(txtB101213) = 0 Then
      MsgBox "加班時數不可空白！", vbExclamation
      Exit Function
   End If
   '2018/4/24 END
   
   'Add By Sindy 2015/12/25 增加檢查同仁加班合計是否有超過46小時
   'Modify By Sindy 2016/12/26
   'Call PUB_PerFormRemindMsg(Left(CboB1002, 2), "0", txtB1003, txtB1004, txtB1012, txtB1013, True)
   Call PUB_PerFormRemindMsg(Left(CboB1002, 2), "0", txtB1003, txtB1004, txtB101213, True)
End If

If Left(CboB1002, 2) = 表單類別_出差 Then
   If Me.txtB1014.Enabled = True Then
      Cancel = False
      txtB1014_Validate Cancel
      If Cancel = True Then
         txtB1014.SetFocus
         Exit Function
      End If
   End If
   If Me.txtB1015.Enabled = True Then
      Cancel = False
      txtB1015_Validate Cancel
      If Cancel = True Then
         txtB1015.SetFocus
         Exit Function
      End If
   End If
End If

If Me.txtB1011.Enabled = True Then
   Cancel = False
   txtB1011_Validate Cancel
   If Cancel = True Then
      txtB1011.SetFocus
      Exit Function
   End If
End If

If cboEmp(0).Visible = True Then
   If cboEmp(0) = "" Then
      'Add By Sindy 2018/1/29
      If m_ABS001_1 <> "" Or m_ABS001_2 <> "" Or m_ABS001_3 <> "" Then
         'MsgBox "職代也休假；職務代理人不可以空白！", vbExclamation
         MsgBox "職代休假或出差；請自行加設其他職代！", vbExclamation
      Else
      '2018/1/29 END
         MsgBox "職務代理人不可空白！", vbExclamation
      End If
      cboEmp(0).SetFocus
      Exit Function
   End If
   For i = 0 To cboEmp.UBound
      If Me.cboEmp(i).Enabled = True And cboEmp(i).Text <> "" Then
         Cancel = False
         bolChk = False
         cboEmp_Validate i, Cancel
         If Cancel = True Then
            cboEmp(i).SetFocus
            Exit Function
         End If
      End If
      If cboEmp(i).Text <> "" Then
         intMaxi = i
      End If
   Next i
   'Add By Sindy 2021/3/29 總經理只有一職代,但剛好職代有請假重覆幾天,要彈提醒
   If Me.cboEmp(intMaxi).Enabled = False Then
      For i = cboEmp.UBound To 0 Step -1
         If cboEmp(i).Text <> "" Then
            If CheckIsPersonRestSector(Left(cboEmp(i), 5), txtB1004, Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), txtB1006, Trim(txtB1007_1.Text & ":" & txtB1007_2.Text), txtB1001) = True Then
               '職代也有休假,是最後一位職代,彈提醒
               If intMaxi = i Then
                  'Modify By Sindy 2025/3/25 依所長指示修改
                  If txtB1003 <> Pub_GetSpecMan("所長員工編號") Then
                  '2025/3/25 END
                     MsgBox "該請假區間【" & Trim(Mid(cboEmp(i), 6)) & "】和您重覆，請自行加設其他職代！", vbExclamation
   '                  Me.cboEmp(intMaxi + 1).Enabled = True
   '                  Me.cboEmp(intMaxi + 1).SetFocus
                     Exit Function
                  End If
               End If
            Else
               Exit For '有職代無休假,但不需往下檢查
            End If
         End If
      Next i
   End If
   '2021/3/29 END
End If

'所長可以無審核主管
strSql = "SELECT st20 FROM staff WHERE ST01='" & txtB1003 & "' "
intI = 1: strST20 = ""
Set RsTemp = ClsLawReadRstMsg(intI, strSql)
If intI = 1 Then
   If Not IsNull(RsTemp("ST20")) Then strST20 = RsTemp("ST20")
End If
'Modify By Sindy 2022/7/18 + 15.名譽所長
If strST20 <> "11" And strST20 <> "15" And CboBoss(0) = "" Then
   MsgBox "審核主管不可以空白！", vbExclamation
   CboBoss(0).SetFocus
   Exit Function
End If

intBossNum = 0
For i = 0 To CboBoss.UBound
   If Me.CboBoss(i).Enabled = True Then
      Cancel = False
      CboBoss_Validate i, Cancel
      If Cancel = True Then
         CboBoss(i).SetFocus
         Exit Function
      End If
   End If
   If Me.CboBoss(i).Text <> "" Then
      intBossNum = intBossNum + 1
   End If
Next i

'增加判斷權責主管人數是否足夠，若不足，則不可簽核
If Val(m_BossNum) > 0 Then
   If intBossNum < Val(m_BossNum) Then
      MsgBox "審核主管人數應為" & m_BossNum & "人，人數不足不可簽核！", vbExclamation
      Exit Function
   End If
End If

'Add by Sindy 2021/5/28 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me) = False Then
   Exit Function
End If
'2021/5/28 END

TxtValidate = True
End Function

'Modify By Sindy 2025/10/30
'Private Sub CboB1002_Click()
Public Sub CboB1002_Click()
'2025/10/30 END
   TxtOverTimeNote.Visible = False 'Add By Sindy 2017/1/5
   If Left(CboB1002.Text, 2) = 表單類別_請假 Or Left(CboB1002.Text, 2) = "" Then
      FrameNote.Visible = True 'Modify By Sindy 2024/11/5 txtB1008_2.Visible = True
      Label10.Visible = True
      CboB1008.Visible = True
      Label1(2).Visible = True
      txtB1006.Visible = True
      LblEndW.Visible = True 'Add By Sindy 2020/8/21
      Frame01.Visible = True
      Frame02.Visible = False
      Frame03.Visible = False
      Label13.Visible = True
      cboEmp(0).Visible = True
      Label25.Visible = True
      cboEmp(1).Visible = True
      Label27.Visible = True
      cboEmp(2).Visible = True
      Label28.Visible = True
      'Add By Sindy 2013/6/28 教召令
      Label33.Visible = True
      Label34.Visible = True
      '2013/6/28 END
      Label30.Visible = False
      txtB1002_01Note.Visible = True
      Chk1Day.Visible = True: If txtB1001 = "" Then Chk1Day.Value = 0
   '      Label1(3).Visible = True
   '      cboSTime.Visible = True
   '      Label1(4).Visible = True
   '      cboETime.Visible = True
   ElseIf Left(CboB1002.Text, 2) = 表單類別_加班 Then
      FrameNote.Visible = False 'Modify By Sindy 2024/11/5 txtB1008_2.Visible = False
      Label10.Visible = False
      CboB1008.Visible = False
      CboB1008.Text = ""
      Label1(2).Visible = False
      txtB1006.Visible = False
      LblEndW.Visible = False 'Add By Sindy 2020/8/21
      Frame01.Visible = False
      Frame02.Visible = True
      Frame03.Visible = False
      Frame02.Left = 3090 '900
      Frame02.Top = 2070
      Label13.Visible = False
      cboEmp(0).Visible = False
      Label25.Visible = False
      cboEmp(1).Visible = False
      Label27.Visible = False
      cboEmp(2).Visible = False
      Label28.Visible = False
      'Add By Sindy 2013/6/28 教召令
      Label33.Visible = False
      Label34.Visible = False
      '2013/6/28 END
      Label30.Visible = True
      txtB1002_01Note.Visible = False
      Chk1Day.Visible = False: Chk1Day.Value = 1
'      Label1(3).Visible = False
'      cboSTime.Visible = False
'      Label1(4).Visible = False
'      cboETime.Visible = False
      Frame1.Visible = False
      TxtOverTimeNote.Visible = True 'Add By Sindy 2017/1/5
   ElseIf Left(CboB1002.Text, 2) = 表單類別_出差 Then
      FrameNote.Visible = False 'Modify By Sindy 2024/11/5 txtB1008_2.Visible = False
      Label10.Visible = False
      CboB1008.Visible = False
      CboB1008.Text = ""
      Label1(2).Visible = True
      txtB1006.Visible = True
      LblEndW.Visible = True 'Add By Sindy 2020/8/21
      Frame01.Visible = True
      Frame02.Visible = False
      Frame03.Visible = True
      Label13.Visible = True
      cboEmp(0).Visible = True
      Label25.Visible = True
      cboEmp(1).Visible = True
      Label27.Visible = True
      cboEmp(2).Visible = True
      Label28.Visible = False
      'Add By Sindy 2013/6/28 教召令
      Label33.Visible = False
      Label34.Visible = False
      '2013/6/28 END
      Label30.Visible = True 'False
      txtB1002_01Note.Visible = False
      'Modify By Sindy 2012/4/13
      Chk1Day.Visible = True: If txtB1001 = "" Then Chk1Day.Value = 0
      'Chk1Day.Visible = False: Chk1Day.Value = 1
      '2012/4/13 End
   '      Label1(3).Visible = True
   '      cboSTime.Visible = True
   '      Label1(4).Visible = True
   '      cboETime.Visible = True
      Frame1.Visible = False
   End If
   'Modify By Sindy 2012/7/9 +尤春彬
   If txtB1003 = "99029" Or txtB1003 = "84043" Then '伊恩
      Chk1Day.Visible = False: Chk1Day.Value = 1
   End If
   Call Chk1Day_Click
End Sub

Private Sub CboB1002_GotFocus()
   InverseTextBox CboB1002
End Sub

Private Sub CboB1002_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboB1002_LostFocus()
   If Left(CboB1002, 2) = 表單類別_請假 Then
      If CboB1008.Enabled = True Then CboB1008.SetFocus
   ElseIf Left(CboB1002, 2) = 表單類別_加班 Then
      If txtB1004.Enabled = True Then txtB1004.SetFocus
   ElseIf Left(CboB1002, 2) = 表單類別_出差 Then
      If txtB1014.Enabled = True Then txtB1014.SetFocus
   End If
   If CboB1002.Text > "" Then
      For i = 0 To CboB1002.ListCount - 1
         If Left(CboB1002.List(i), 2) = CboB1002.Text Then CboB1002.Text = CboB1002.List(i): Exit For
      Next i
   End If
End Sub

Private Sub CboB1002_Validate(Cancel As Boolean)
Dim bolComp As Boolean
   
   If CboB1002 <> "" Then
      bolComp = False
      For i = 0 To CboB1002.ListCount
         If Left(CboB1002, 2) = Left(CboB1002.List(i), 2) Then
            bolComp = True
            Exit For
         End If
      Next i
      If bolComp = False Then
         MsgBox "表單類別有誤!!!", vbExclamation + vbOKOnly
         Call CboB1002_GotFocus
         Cancel = True
         Exit Sub
      End If
   Else
      MsgBox "表單類別不可以空白！", vbExclamation
      Call CboB1002_GotFocus
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub CboB1008_GotFocus()
   InverseTextBox CboB1008
End Sub

Private Sub CboB1008_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboB1008_LostFocus()
   'Call Chk1Day_Click 'Modify By Sindy 2025/9/5 mark
   If CboB1008.Text > "" Then
      For i = 0 To CboB1008.ListCount - 1
         If Left(CboB1008.List(i), 2) = CboB1008.Text Then CboB1008.Text = CboB1008.List(i): Exit For
      Next i
      'Modify By Sindy 2017/11/14 Mark
'      If GetCountDayHour(False) Then
'         Call CboB1008_GotFocus
'         Exit Sub
'      End If
      If Left(CboB1008.Text, 2) = "08" Then
         txtB1004.SetFocus
         'Modify By Sindy 2025/9/5
         If Chk1Day.Value = 1 Then
            Chk1Day.Tag = ""
         Else
            Chk1Day.Tag = "1"
         End If
         Call Chk1Day_Click
         '2025/9/5 END
      End If
   End If
End Sub

Private Sub CboB1008_Validate(Cancel As Boolean)
Dim MyRs As New ADODB.Recordset
Dim MyArr As Variant
   
   If CboB1008.Text <> "" Then
      MyArr = Split(CboB1008, " ")
      Set MyRs = New ADODB.Recordset
      If MyRs.State = 1 Then MyRs.Close
      ' 排除不須要的代碼 : 01.忘打卡 02.遲到 03.曠職 04.出差 16.加班 17.扣年終產假 18.扣年終流產假
      strSql = "select ac02||' '||ac03 from allcode where ac01='04' and ac02='" & MyArr(0) & "' and ac02 not in ('01','02','03','04','16','17','18') order by ac02"
      MyRs.CursorLocation = adUseClient
      MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If MyRs.RecordCount <> 0 Then
         CboB1008.Text = "" & MyRs.Fields(0).Value
      Else
         MsgBox "假別代號輸入錯誤!!!", vbExclamation + vbOKOnly
         Call CboB1008_GotFocus
         Cancel = True
         Exit Sub
      End If
   Else
      MsgBox "假別不可以空白！", vbExclamation
      Call CboB1008_GotFocus
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub txtB1004_GotFocus()
   InverseTextBox txtB1004
   CloseIme
End Sub

Private Sub txtB1004_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'Private Sub txtB1004_LostFocus()
'   If txtB1006.Visible = True Then
'      If txtB1004 <> "" And txtB1006 <> "" And txtB1004 <> txtB1006 Then
'         Call Chk1Day_Click
'      End If
'   End If
'End Sub

Private Sub txtB1004_Validate(Cancel As Boolean)
Dim strTime As String

'Modify By Sindy 2011/12/7 Mark
'If txtB1006.Visible = True Then
'   If txtB1004 <> "" And txtB1006 <> "" And txtB1004 <> txtB1006 Then
'      Call Chk1Day_Click
'   End If
'End If

If txtB1004 <> "" Then
   If CheckIsTaiwanDate(txtB1004, False) = False Then
      MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
      Call txtB1004_GotFocus
      Cancel = True
      Exit Sub
   End If
   'Add By Sindy 2012/1/6 不允許輸入2個月前的假單
   'Modify By Sindy 2023/6/8 '當月前1個月
   'If DBDATE(txtB1004) < strSysDtBef2M Then
   If DBDATE(txtB1004) < strSrvDate(1) Then
      If Left(DBDATE(txtB1004), 6) <> Left(strSrvDate(1), 6) And _
         Left(DBDATE(txtB1004), 6) <> Left(DBDATE(strSysDtBef1M), 6) Then
         'MsgBox "起始日期不可小於系統日2個月！", vbInformation, "輸入日期錯誤"
         MsgBox "起始日期只能輸入當月和上個月的日期！", vbInformation, "輸入日期錯誤"
   '2023/6/8 END
         Call txtB1004_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   '2012/1/6 End
   If Left(CboB1002, 2) = 表單類別_請假 Then
      Me.txtB1008_2 = GetCurrSpecRestDay(txtB1003, , Left(txtB1004, 3)) 'Add By Sindy 2014/12/3
      Me.txtB1008_14 = GetCurrFor14RestDay(txtB1003, , txtB1004) 'Add By Sindy 2024/12/10
      'Add By Sindy 2012/10/22 產假和流產假可輸入工作天
      If Left(CboB1008, 2) <> "10" And Left(CboB1008, 2) <> "11" Then
      '2012/10/22 End
         If ChkWorkDay(DBDATE(txtB1004)) = False Then
            MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
            Call txtB1004_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   'Add By Sindy 2014/4/29
   ElseIf Left(CboB1002, 2) = 表單類別_加班 Then
      'Modify By Sindy 2016/12/26
      'Modify By Sindy 2019/10/3 不需檢查颱風假
      'If ChkWorkDay(ChangeTStringToWString(txtB1004), txtB1003, True) = False Then '假日
      If ChkWorkDay(ChangeTStringToWString(txtB1004), txtB1003) = False Then '假日
      '2019/10/3 END
         'txtB1012.Text = ""
         'txtB1012.Enabled = False
         Label16.Caption = "假日-共                     時"
      Else '平日
         'txtB1013.Text = ""
         'txtB1013.Enabled = False
         Label16.Caption = "平日-共                     時"
      End If
      '2016/12/26 END
   '2014/4/29 END
      'Add By Sindy 2014/11/20 禁止同仁周日加班,固鎖住不得填寫,只有人事處可以代同仁填寫
      If Weekday(Format(DBDATE(txtB1004), "####-##-##")) = 1 Then '星期日
         MsgBox "不可填寫周日加班單，有問題請洽人事處！", vbInformation, "輸入日期錯誤"
         Call txtB1004_GotFocus
         Cancel = True
         Exit Sub
      End If
      'Add By Sindy 2015/7/7 加班日期不可輸入大於系統日
      If Val(txtB1004) > Val(strSrvDate(2)) Then
         MsgBox "加班日期不可輸入大於系統日！", vbInformation, "輸入日期錯誤"
         Call txtB1004_GotFocus
         Cancel = True
         Exit Sub
      End If
      
      'Add By Sindy 2025/11/3 下班逾30分鐘原因確認非處理公務時,不可填寫加班單
      strSql = "select * from abs015 " & _
                "where B1501='" & txtB1003 & "' and B1502=" & DBDATE(txtB1004) & _
                  "and B1504<>'2'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         MsgBox "此日期已確認下班逾30分鐘原因為 非處理公務，因此不可填寫加班單！請洽人事處。", vbInformation, "輸入日期錯誤"
         Call txtB1004_GotFocus
         Cancel = True
         Exit Sub
      End If
      '2025/11/3 END
   End If
   
   If txtB1004 <> "" And txtB1006 <> "" Then
      If Val(txtB1004) > Val(txtB1006) Then
         txtB1006 = ""
      Else
         If RunNick2(txtB1004, txtB1006) Then
            Call txtB1004_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
   If Left(CboB1008, 2) = "08" Then '特別假必須要提早1天前請(指工作天)
'      'Add By Sindy 2013/3/20
'      If DBDATE(txtB1004) < DBDATE(DateAdd("yyyy", 1, ChangeWStringToWDateString(m_ST13))) Then
'         Call txtB1004_GotFocus
'         Cancel = True
'         MsgBox "特別假必須到職滿一年才可使用！", vbInformation, "輸入日期錯誤"
'         Exit Sub
'      End If
'      '2013/3/20 End
      If txtB1001 <> "" Then '有表單編號時
         If DBDATE(txtB1004) <= CompWorkDay(2, DBDATE(m_B1023), 0) Then
            Call txtB1004_GotFocus
            Cancel = True
            MsgBox "特別假須提早 1 個工作天！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
      Else
         strTime = Right("000000" & ServerTime, 6)
         If DBDATE(txtB1004) <= CompWorkDay(2, DBDATE(strSrvDate(1)), 0) Then
            Call txtB1004_GotFocus
            Cancel = True
            MsgBox "特別假須提早 1 個工作天！", vbInformation, "輸入日期錯誤"
            Exit Sub
         ElseIf DBDATE(txtB1004) = CompWorkDay(3, DBDATE(strSrvDate(1)), 0) And Val(Left(strTime, Len(strTime) - 2)) >= 1800 Then
            Call txtB1004_GotFocus
            Cancel = True
            MsgBox "特別假已超出可以請假的時間！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
      End If
   End If
   'Modify By Sindy 2011/11/16
   If Frame1.Visible = True And cboSTime <> "" Then
      If txtB1004 = txtB1006 Then
         cboSTime.ListIndex = cboETime.ListIndex
      End If
   End If
   '2011/11/16 End
   If GetCountDayHour(True) = False Then
'      Call txtB1004_GotFocus
'      Cancel = True
'      Exit Sub
   End If
   
   'Add By Sindy 2020/8/14 顯示星期幾
   If Val(txtB1004) > 0 Then
      LblStarW.Caption = "(" & GetWeekDay(CDate(Format(DBDATE(txtB1004), "####/##/##"))) & ")"
   End If
   '2020/8/14 END
End If
End Sub

Private Sub txtB1005_1_GotFocus()
   InverseTextBox txtB1005_1
End Sub

Private Sub txtB1005_1_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1005_1_Validate(Cancel As Boolean)
If txtB1005_1 = "" Then txtB1005_1 = "00"

If txtB1005_1 <> "" Then
   If CheckLengthIsOK(txtB1005_1, txtB1005_1.MaxLength) = False Then
      Call txtB1005_1_GotFocus
      Cancel = True
      Exit Sub
   End If
   If Val(txtB1005_1.Text) = 0 And txtB1004 <> "" Then
      MsgBox "請輸入時分!", vbExclamation + vbOKOnly
      Call txtB1005_1_GotFocus
      Cancel = True
      Exit Sub
   End If
   If txtB1005_1.Text > 24 Then
      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
      Call txtB1005_1_GotFocus
      Cancel = True
      Exit Sub
   End If
   'Modify By Sindy 2014/3/10 檢查日/時欄位有值時，才需要分欄位離開時再重新計算。
   'If Val(txtB1009) > 0 Or Val(txtB1010) > 0 Then
      If GetCountDayHour(True) = False Then
'         Call txtB1005_1_GotFocus
'         Cancel = True
'         Exit Sub
      End If
   'End If
End If
CloseIme
End Sub

Private Sub txtB1005_2_GotFocus()
   InverseTextBox txtB1005_2
End Sub

Private Sub txtB1005_2_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1005_2_Validate(Cancel As Boolean)
If txtB1005_2 = "" Then txtB1005_2 = "00"

If txtB1005_2 <> "" Then
   If CheckLengthIsOK(txtB1005_2, txtB1005_2.MaxLength) = False Then
      Call txtB1005_2_GotFocus
      Cancel = True
      Exit Sub
   End If
   If txtB1005_2.Text > 59 Then
      Call txtB1005_2_GotFocus
      MsgBox "不可超過59分!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   If GetCountDayHour(True) = False Then
'      Call txtB1005_2_GotFocus
'      Cancel = True
'      Exit Sub
   End If
End If
CloseIme
End Sub

Private Sub txtB1006_GotFocus()
   InverseTextBox txtB1006
End Sub

Private Sub txtB1006_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'Private Sub txtB1006_LostFocus()
'   If txtB1006.Visible = True Then
'      If txtB1004 <> "" And txtB1006 <> "" And txtB1004 <> txtB1006 Then
'         Call Chk1Day_Click
'      End If
'   End If
'End Sub

Private Sub txtB1006_Validate(Cancel As Boolean)
Dim strTime As String

'Modify By Sindy 2011/12/7 Mark
'If txtB1006.Visible = True Then
'   If txtB1004 <> "" And txtB1006 <> "" And txtB1004 <> txtB1006 Then
'      Call Chk1Day_Click
'   End If
'End If

'Chk1Day_Click 'Add By Sindy 2017/11/3
If txtB1006 <> "" Then
   If CheckIsTaiwanDate(txtB1006, False) = False Then
      MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
      Call txtB1006_GotFocus
      Cancel = True
      Exit Sub
   End If
   'Modify By Sindy 2014/10/7 Mark:應判斷在起始日期即可
'   'Add By Sindy 2012/1/6 不允許輸入2個月前的假單
'   If DBDATE(txtB1006) < strSysDtBef2M Then
'      MsgBox "迄止日期不可小於系統日2個月！", vbInformation, "輸入日期錯誤"
'      Call txtB1006_GotFocus
'      Cancel = True
'      Exit Sub
'   End If
'   '2012/1/6 End
   If Left(CboB1002, 2) = 表單類別_請假 Then
      'Add By Sindy 2012/10/22 產假和流產假可輸入工作天
      If Left(CboB1008, 2) <> "10" And Left(CboB1008, 2) <> "11" Then
      '2012/10/22 End
         If ChkWorkDay(DBDATE(txtB1006)) = False Then
            Call txtB1006_GotFocus
            Cancel = True
            MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
      End If
   End If
   If txtB1004 <> "" And txtB1006 <> "" Then
      If RunNick2(txtB1004, txtB1006) Then
         Call txtB1006_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
'   If txtB1007_2.Enabled = False Then '特別假時,若輸入迄止日期即可計算天,時
'      Cancel = False
'      txtB1007_2_Validate Cancel
'      If Cancel = True Then
'         Exit Sub
'      End If
'   End If
   If Left(CboB1008, 2) = "08" Then '特別假必須要提早1天前請(指工作天)
      If txtB1001 <> "" Then '有表單編號時
         If DBDATE(txtB1006) <= CompWorkDay(2, DBDATE(m_B1023), 0) Then
            Call txtB1006_GotFocus
            Cancel = True
            MsgBox "特別假必須要提早1天前請(指工作天)！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
      Else
         strTime = Right("000000" & ServerTime, 6)
         If DBDATE(txtB1006) <= CompWorkDay(2, DBDATE(strSrvDate(1)), 0) Then
            Call txtB1006_GotFocus
            Cancel = True
            MsgBox "特別假必須要提早1天前請(指工作天)！", vbInformation, "輸入日期錯誤"
            Exit Sub
         ElseIf DBDATE(txtB1006) = CompWorkDay(3, DBDATE(strSrvDate(1)), 0) And Val(Left(strTime, Len(strTime) - 2)) >= 1800 Then
            Call txtB1006_GotFocus
            Cancel = True
            MsgBox "特別假已超出可以請假的時間！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
      End If
   End If
   'Modify By Sindy 2011/11/16
   If Frame1.Visible = True And cboETime <> "" Then
      If txtB1004 = txtB1006 Then
         cboETime.ListIndex = cboSTime.ListIndex
      End If
   End If
   '2011/11/16 End
   If GetCountDayHour(True) = False Then
'      Call txtB1006_GotFocus
'      Cancel = True
'      Exit Sub
   End If
   
   'Add By Sindy 2017/12/13 特別假107年開始可以請半天(4小時)
   If CboB1008.Visible = True And Left(CboB1008, 2) = "08" And DBDATE(txtB1006) >= "20180101" Then
      Chk1Day.Enabled = True
   End If
   
   'Add By Sindy 2020/8/14 顯示星期幾
   If Val(txtB1006) > 0 Then
      LblEndW.Caption = "(" & GetWeekDay(CDate(Format(DBDATE(txtB1006), "####/##/##"))) & ")"
   End If
   '2020/8/14 END
End If
End Sub

Private Sub txtB1007_1_GotFocus()
   InverseTextBox txtB1007_1
End Sub

Private Sub txtB1007_1_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1007_1_Validate(Cancel As Boolean)
If txtB1007_1 = "" Then txtB1007_1 = "00"

If txtB1007_1 <> "" Then
   If CheckLengthIsOK(txtB1007_1, txtB1007_1.MaxLength) = False Then
      Call txtB1007_1_GotFocus
      Cancel = True
      Exit Sub
   End If
   If Val(txtB1007_1.Text) = 0 And _
      ((txtB1004 <> "" And Left(CboB1002, 2) = 表單類別_加班) Or (txtB1006 <> "" And Left(CboB1002, 2) <> 表單類別_加班)) Then
      Call txtB1007_1_GotFocus
      MsgBox "請輸入時分!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   If txtB1007_1.Text > 24 Then
      Call txtB1007_1_GotFocus
      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   'Modify By Sindy 2014/3/10 檢查日/時欄位有值時，才需要分欄位離開時再重新計算。
   'ex.（103/3/10 9:00 ~ 103/3/10 9:30）輸入迄止的9點時,會因Run計算,算出來的日/時是0而傳回false,而鎖定在該欄位不能動彈
   'If Val(txtB1009) > 0 Or Val(txtB1010) > 0 Then
      If GetCountDayHour(True) = False Then
'         Call txtB1007_1_GotFocus
'         Cancel = True
'         Exit Sub
      End If
   'End If
End If
CloseIme
End Sub

Private Sub txtB1007_2_GotFocus()
   InverseTextBox txtB1007_2
End Sub

Private Sub txtB1007_2_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1007_2_Validate(Cancel As Boolean)
If txtB1007_2 = "" Then txtB1007_2 = "00"

If txtB1007_2 <> "" Then
   If CheckLengthIsOK(txtB1007_2, txtB1007_2.MaxLength) = False Then
      Call txtB1007_2_GotFocus
      Cancel = True
      Exit Sub
   End If
   If Left(CboB1002, 2) = 表單類別_加班 Then '止日的分：只接受10,20,30,40,50,00之值
      If txtB1007_2 <> "10" And txtB1007_2 <> "20" And txtB1007_2 <> "30" And txtB1007_2 <> "40" And _
         txtB1007_2 <> "50" And txtB1007_2 <> "00" Then
         MsgBox "填加班單時，止日的分，只接受10,20,30,40,50,00之值!!!", vbExclamation + vbOKOnly
         Call txtB1007_2_GotFocus
         Cancel = True
         Exit Sub
      End If
   Else
      If txtB1007_2.Text > 59 Then
         Call txtB1007_2_GotFocus
         MsgBox "不可超過59分!", vbExclamation + vbOKOnly
         Cancel = True
         Exit Sub
      End If
   End If
   'If Trim(txtB1004) <> "" And Trim(txtB1005_1) <> "" And Trim(txtB1005_2) <> "" And Trim(txtB1006) <> "" And Trim(txtB1007_1) <> "" And Trim(txtB1007_2) <> "" Then
      If CheckIsTaiwanDate(txtB1004, False) = True And CheckIsTaiwanDate(txtB1006, False) = True Then
         If CompDateTime(txtB1004 & Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), txtB1006 & Format(txtB1007_1, "00") & Format(txtB1007_2, "00")) = False Then
            Call txtB1007_2_GotFocus
            MsgBox "日期時間設定錯誤！", vbInformation, "輸入錯誤！"
            'Cancel = True
            Exit Sub
         End If
      End If
   'End If
   If GetCountDayHour(True) = False Then
'      Call txtB1007_2_GotFocus
'      Cancel = True
'      Exit Sub
   End If
End If
CloseIme
End Sub

Private Function GetCountDayHour(bolChkExist As Boolean) As Boolean
Dim dblSTime As Double, dblETime As Double, strB1008 As String, strB1014 As String
   
   GetCountDayHour = True
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If Left(CboB1002, 2) = 表單類別_加班 Then
         '欄位值尚未輸入完整
         If Val(txtB1004) = 0 Or Val(txtB1005_1) = 0 Or _
            Val(txtB1007_1) = 0 Then
            Exit Function
         End If
         '無異動欄位值
         If Val(txtB1004) = Val(m_B1004) And Val(txtB1005_1) = Val(m_B1005_1) And Val(txtB1005_2) = Val(m_B1005_2) And _
            Val(txtB1007_1) = Val(m_B1007_1) And Val(txtB1007_2) = Val(m_B1007_2) And _
            CboBoss(0) <> "" And Val(txtB101213) <> 0 Then
            Exit Function
         Else
            'Add By Sindy 2014/4/29 有異動欄位值時,清空加班時數重算
'            txtB1012.Text = ""
'            txtB1013.Text = ""
            txtB101213.Text = "" 'Modify By Sindy 2016/12/26
            '2014/4/29 END
         End If
         If bolChkExist = True Then
            If IsRecordExist(txtB1003, DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1004), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
               GetCountDayHour = False
               Exit Function
            End If
         End If
         'If txtB1012 = "" And txtB1013 = "" Then Call AutoCount
      Else
         '欄位值尚未輸入完整
         If Val(txtB1004) = 0 Or Val(txtB1005_1) = 0 Or _
            Val(txtB1006) = 0 Or Val(txtB1007_1) = 0 Or _
            (Frame1.Visible = True And cboSTime = "" And Chk1Day.Value = 1 And txtB1004 <> txtB1006) Or _
            (Frame1.Visible = True And cboETime = "" And Chk1Day.Value = 1 And txtB1004 <> txtB1006) Then
            Exit Function
         End If
         '無異動欄位值
         If Frame1.Visible = True And cboSTime.Text <> "" Then dblSTime = Val(Format(cboSTime.Text, "hhmm"))
         If Frame1.Visible = True And cboETime.Text <> "" Then dblETime = Val(Format(cboETime.Text, "hhmm"))
         If CboB1008.Visible = True And CboB1008.Text <> "" Then strB1008 = Left(CboB1008, 2)
         If Frame03.Visible = True And txtB1014 <> "" Then strB1014 = txtB1014
         If Val(txtB1004) = Val(m_B1004) And Val(txtB1005_1) = Val(m_B1005_1) And Val(txtB1005_2) = Val(m_B1005_2) And _
            Val(txtB1006) = Val(m_B1006) And Val(txtB1007_1) = Val(m_B1007_1) And Val(txtB1007_2) = Val(m_B1007_2) And _
            dblSTime = Val(m_B1028) And dblETime = Val(m_B1029) And strB1008 = m_B1008 And _
            cboEmp(0) <> "" And CboBoss(0) <> "" And strB1014 = m_B1014 And Not (Val(txtB1009) = 0 And Val(txtB1010) = 0) Then
            Exit Function
         End If
         If bolChkExist = True Then
            If IsRecordExist(txtB1003, DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1006), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
               GetCountDayHour = False
               Exit Function
            End If
         End If
         'If (txtB1009 = "" And txtB1010 = "") Or (txtB1009 = "0" And txtB1010 = "0") Then Call AutoCount
      End If
      'Modify By Sindy 2017/11/14
      'If AutoCount = False Then GetCountDayHour = False: Exit Function
      If AutoCount = False Then
         GetCountDayHour = False
'         'bolChkExist = True:還在檢查輸入的欄位值,
'         '                   因人員還在調整欄位值,不要回傳False不然會鎖住游標在此欄位上不能移動
'         If bolChkExist = True Then
'            GetCountDayHour = True
'         End If
'         'END
         Exit Function
      End If
      '2017/11/14 END
   End If
End Function

Private Sub txtB1009_GotFocus()
   InverseTextBox txtB1009
End Sub

Private Sub txtB1009_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1009_Validate(Cancel As Boolean)
If txtB1009 <> "" Then
   If CheckLengthIsOK(txtB1009, txtB1009.MaxLength) = False Then
      Call txtB1009_GotFocus
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
End Sub

Private Sub txtB1010_GotFocus()
   InverseTextBox txtB1010
End Sub

Private Sub txtB1010_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub txtB1010_Validate(Cancel As Boolean)
If txtB1010 <> "" Then
   If CheckLengthIsOK(txtB1010, txtB1010.MaxLength) = False Then
      Call txtB1010_GotFocus
      Cancel = True
      Exit Sub
   End If
   'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
   'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
   'Modify By Sindy 2012/7/9 上班時數為特殊者
'   Call Pub_GetSpecWorkHour(txtB1003, txtB1004)
'   If txtB1003 = "99029" Then
'      If txtB1010.Text >= 5 Then
'         Call txtB1010_GotFocus
'         MsgBox "請假時數-共(時)不可超過5小時!!!", vbExclamation + vbOKOnly
'         Cancel = True
'         Exit Sub
'      End If
   If Val(txtB1010.Text) >= Val(PUB_intWkHour) Then
      Call txtB1010_GotFocus
      MsgBox "請假時數-共(時)不可超過" & PUB_intWkHour & "小時!!!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
End Sub

Private Sub txtB1011_GotFocus()
   InverseTextBox txtB1011
   OpenIme
End Sub

'Add By Sindy 2021/5/31
Private Sub txtB1011_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtB1011
End Sub

Private Sub txtB1011_Validate(Cancel As Boolean)
If txtB1011 <> "" Then
   If CheckLengthIsOK(txtB1011, txtB1011.MaxLength) = False Then
      Call txtB1011_GotFocus
      Cancel = True
      Exit Sub
   End If
'   'Modify By Sindy 2011/12/6
'   If m_B1018 = 主管代填 Then Call GetCountDayHour(False)
'   '2011/12/6 End
   'Add By Sindy 2017/12/28
   If GetCountDayHour(True) = False Then
'      Cancel = True
'      Exit Sub
   End If
   '2017/12/28 END
End If
CloseIme
End Sub

Private Sub txtB1014_GotFocus()
   InverseTextBox txtB1014
End Sub

Private Sub txtB1014_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'Add By Sindy 2011/10/14
Private Sub txtB1014_LostFocus()
   'Modify By Sindy 2017/11/14 Mark
'   Call GetCountDayHour(False)
End Sub

Private Sub txtB1014_Validate(Cancel As Boolean)
   If txtB1014 <> "" Then
      If CheckLengthIsOK(txtB1014, txtB1014.MaxLength) = False Then
          Call txtB1014_GotFocus
          Cancel = True
          Exit Sub
      End If
      If Trim(txtB1014) <> "" Then
        If txtB1014 <> "1" And txtB1014 <> "2" And txtB1014 <> "3" And txtB1014 <> "4" Then
           MsgBox "差程代碼有誤!!!", vbExclamation + vbOKOnly
           Call txtB1014_GotFocus
           Cancel = True
           Exit Sub
        End If
      End If
   Else
      MsgBox "差程不可以空白！", vbExclamation
      Call txtB1014_GotFocus
      Cancel = True
      Exit Sub
   End If
CloseIme
End Sub

Private Sub txtB1015_GotFocus()
   InverseTextBox txtB1015
   OpenIme
End Sub

'Add By Sindy 2021/5/31
Private Sub txtB1015_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtB1015
End Sub

Private Sub txtB1015_Validate(Cancel As Boolean)
   If txtB1015 <> "" Then
      If CheckLengthIsOK(txtB1015, txtB1015.MaxLength) = False Then
         Call txtB1015_GotFocus
         Cancel = True
         Exit Sub
      End If
'   Else
'      MsgBox "出差地點不可以空白！", vbExclamation
'      Call txtB1015_GotFocus
'      Cancel = True
'      Exit Sub
   End If
CloseIme
End Sub

Private Sub CboEmp_GotFocus(Index As Integer)
   InverseTextBox cboEmp(Index)
   bolChk = True
End Sub

Private Sub CboEmp_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboEmp_LostFocus(Index As Integer)
   If cboEmp(Index).Text > "" And Len(Trim(cboEmp(Index).Text)) = 5 Then
      '抓取員工姓名
      cboEmp(Index).Text = SetCboStaffName(cboEmp(Index).Text)
   End If
End Sub

Private Sub cboEmp_Validate(Index As Integer, Cancel As Boolean)
Dim strMsgText As String
   
   If cboEmp(Index) <> "" Then
      If Left(cboEmp(Index), 5) = txtB1003 Then
         MsgBox "不可為本人！", vbExclamation
         If cboEmp(Index).Enabled = True Then cboEmp(Index).SetFocus
         Call CboEmp_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(cboEmp(Index), 5)) = True Then
         Call CboEmp_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查 員工不可為”不寄信”
      If ChkStaffST14(Left(cboEmp(Index), 5)) = True Then
         Call CboEmp_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      Call CboEmp_LostFocus(Index) 'Add By Sindy 2025/2/17
      '檢查職代輸入順序
      If (Trim(cboEmp(1)) <> "" And Trim(cboEmp(0)) = "") Or _
         (Trim(cboEmp(2)) <> "" And Trim(cboEmp(1)) = "") Then
         MsgBox "請依序輸入職務代理人！", vbExclamation
         If cboEmp(Index).Enabled = True Then cboEmp(Index).SetFocus
         Call CboEmp_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      If (cboEmp(1) <> "" And Left(cboEmp(1), 5) = Left(cboEmp(0), 5)) Or _
         (cboEmp(2) <> "" And Left(cboEmp(2), 5) = Left(cboEmp(1), 5)) Then
         MsgBox "資料重覆！", vbExclamation
         If cboEmp(Index).Enabled = True Then cboEmp(Index).SetFocus
         Call CboEmp_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      If bolChk = True Then
         bolChk = False
         strMsgText = ""
         '檢查取得的職代和表單當事人是否有相同的請假區間
         'Modify By Sindy 2017/1/10
         If CheckIsPersonRestSectorSame(Left(cboEmp(Index), 5), txtB1004, Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), txtB1006, Trim(txtB1007_1.Text & ":" & txtB1007_2.Text), txtB1001) = True Then
            'Modify By Sindy 2025/2/17 依所長指示修改:其職代為總經理，因職權關係即使職代休假亦維持總經理為其請假時之職代
            If txtB1003 = Pub_GetSpecMan("所長員工編號") Then
               If MsgBox("該請假區間【" & Trim(Mid(cboEmp(Index), 6)) & "】和您重覆，確定選為職代嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                  If cboEmp(Index).Enabled = True Then cboEmp(Index).SetFocus
                  Call CboEmp_GotFocus(Index)
                  Cancel = True
                  Exit Sub
               End If
            Else
            '2025/2/17 END
               MsgBox "該請假區間【" & Trim(Mid(cboEmp(Index), 6)) & "】休假或出差，不可為職代！", vbExclamation
               If cboEmp(Index).Enabled = True Then cboEmp(Index).SetFocus
               Call CboEmp_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         ElseIf CheckIsPersonRestSector(Left(cboEmp(Index), 5), txtB1004, Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), txtB1006, Trim(txtB1007_1.Text & ":" & txtB1007_2.Text), txtB1001) = True Then
'            strMsgText = "該請假區間此人員休假"
            'MsgBox "該請假區間此人員休假，不可為職代！", vbExclamation
            If MsgBox("該請假區間【" & Trim(Mid(cboEmp(Index), 6)) & "】和您重覆，確定選為職代嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
               If cboEmp(Index).Enabled = True Then cboEmp(Index).SetFocus
               Call CboEmp_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         '2017/1/10 END
         'Modify By Sindy 2024/10/7 mark:多餘的程式控管
'         If CheckIsPersonRest(Left(CboEmp(Index), 5), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = True Then
''            If strMsgText <> "" Then strMsgText = strMsgText & "，並且"
'            'strMsgText = strMsgText & "此人員今日休假，會延後簽核"
'            'Modify By Sindy 2019/7/12
'            'MsgBox Trim(Mid(cboEmp(Index), 6)) & "今日休假，不可為職代！", vbExclamation
'            If MsgBox("【" & Trim(Mid(CboEmp(Index), 6)) & "】今日休假，會延後簽核，確定繼續嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
'               'Modify By Sindy 2024/10/7
'               'cboEmp(Index).SetFocus
'               If CboEmp(Index).Enabled = True Then CboEmp(Index).SetFocus
'               '2024/10/7 END
'               Call CboEmp_GotFocus(Index)
'               Cancel = True
'               Exit Sub
'            End If
'         End If
         '2024/10/7 END
'         If strMsgText <> "" Then
'            'MsgBox "此人員休假，不可為職代！", vbExclamation
'            If MsgBox(strMsgText & "，確定為職代嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then Exit Sub
'            CboEmp(Index).SetFocus
'            Call CboEmp_GotFocus(Index)
'            Cancel = True
'            Exit Sub
'         End If
      End If
   End If
End Sub

Private Sub CboBoss_GotFocus(Index As Integer)
   InverseTextBox CboBoss(Index)
End Sub

Private Sub CboBoss_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboBoss_LostFocus(Index As Integer)
   If CboBoss(Index).Text > "" And Len(Trim(CboBoss(Index).Text)) = 5 Then
      '抓取員工姓名
      CboBoss(Index).Text = SetCboStaffName(CboBoss(Index).Text)
   End If
End Sub

Private Sub CboBoss_Validate(Index As Integer, Cancel As Boolean)
   If CboBoss(Index) <> "" Then
      If Left(CboBoss(Index), 5) = txtB1003 Then
         MsgBox "不可為本人！", vbExclamation
         CboBoss(Index).SetFocus
         Call CboBoss_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(CboBoss(Index), 5)) = True Then
         Call CboBoss_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查 員工不可為”不寄信”
      If ChkStaffST14(Left(CboBoss(Index), 5)) = True Then
         Call CboBoss_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      If (Trim(CboBoss(1)) <> "" And Trim(CboBoss(0)) = "") Or _
         (Trim(CboBoss(2)) <> "" And Trim(CboBoss(1)) = "") Or _
         (Trim(CboBoss(3)) <> "" And Trim(CboBoss(2)) = "") Or _
         (Trim(CboBoss(4)) <> "" And Trim(CboBoss(3)) = "") Then
         MsgBox "請依序輸入審核主管！", vbExclamation
         CboBoss(Index).SetFocus
         Call CboBoss_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      If (CboBoss(1) <> "" And Left(CboBoss(1), 5) = Left(CboBoss(0), 5)) Or _
         (CboBoss(2) <> "" And Left(CboBoss(2), 5) = Left(CboBoss(1), 5)) Or _
         (CboBoss(3) <> "" And Left(CboBoss(3), 5) = Left(CboBoss(2), 5)) Or _
         (CboBoss(4) <> "" And Left(CboBoss(4), 5) = Left(CboBoss(3), 5)) Then
         MsgBox "資料重覆！", vbExclamation
         CboBoss(Index).SetFocus
         Call CboBoss_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

'計算時數
Private Function CountDayHour()
   'Modify By Sindy 2012/7/9
   Dim strSTime As String, strETime As String
   Dim strB1009 As String, strB1010 As String
   strSTime = "": strETime = ""
   If Frame1.Visible = True Then strSTime = Format(cboSTime.Text, "hhmm")
   If Frame1.Visible = True Then strETime = Format(cboETime.Text, "hhmm")
   strB1009 = txtB1009
   strB1010 = txtB1010
   'Modify by Sindy 2012/10/12
   'Call PUB_CountDayHour(txtB1003, DBDATE(txtB1004), Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), DBDATE(txtB1006), Format(txtB1007_1, "00") & Format(txtB1007_2, "00"), strSTime, strETime, strB1009, strB1010, True)
   Call PUB_CountDayHour(txtB1003, DBDATE(txtB1004), Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), DBDATE(txtB1006), Format(txtB1007_1, "00") & Format(txtB1007_2, "00"), strSTime, strETime, strB1009, strB1010, CboB1008, True)
   txtB1009 = strB1009
   txtB1010 = strB1010
   '2012/7/9 End
End Function

'計算加班時數
'Modify By Sindy 2015/11/19 + Optional bolGetVal As Boolean = False
Private Function CountHour(Optional bolGetVal As Boolean = False)
'Dim calTime
Dim strB101213 As String
   
   If txtB1004.Text <> "" Then
      'TimeSerial : 此函數不可使用, 因它會依各PC的時間設定格式有其不同的結果
      'calTime = Trim(TimeSerial(Val(txtB1007_1) - Val(txtB1005_1), Val(txtB1007_2) - Val(txtB1005_2), 0))
      'dHour = Val(Mid(calTime, 1, 2)) + ((Val(Mid(calTime, 4, 2)) \ 30) * 0.5)
      
      '以半小時為單位
      'dHour = ((((Val(txtB1007_1) * 60) + Val(txtB1007_2)) - ((Val(txtB1005_1) * 60) + Val(txtB1005_2))) \ 30) * 0.5
      '以分鐘為單位, 取至小數第一位, 四捨五入
'      dHour = Round((((Val(txtB1007_1) * 60) + Val(txtB1007_2)) - ((Val(txtB1005_1) * 60) + Val(txtB1005_2))) / 60, 1)
'      If dHour < 0 Then dHour = 0
      dHour = PUB_CountHour_Overtime(txtB1004, txtB1003, txtB1007_1, txtB1007_2, txtB1005_1, txtB1005_2, strB101213)
      If bolGetVal = True Then Exit Function '只為取得dHour變數的值
      
'      '無異動欄位值
'      If Val(txtB1004) = Val(m_B1004) And Val(txtB1005_1) = Val(m_B1005_1) And Val(txtB1005_2) = Val(m_B1005_2) And _
'         Val(txtB1007_1) = Val(m_B1007_1) And Val(txtB1007_2) = Val(m_B1007_2) And _
'         Val(txtB1030.Text) > 0 Then
'      Else
'         '假日時數
'         'Modify By Sindy 2012/8/15 增加檢查颱風假
'         'If ChkWorkDay(ChangeTStringToWString(txtB1004)) = False Then
'         'Modify By Sindy 2016/12/26
         txtB1030.Text = dHour
         txtB101213.Text = strB101213
'         '2016/12/26 END
''         If ChkWorkDay(ChangeTStringToWString(txtB1004), txtB1003, True) = False Then
''         '2012/8/15 End
'''             If txtB1013.Text = "" Or txtB1013.Text <= "0" Then
''               txtB1012.Text = ""
''               txtB1012.Enabled = False
''               'Modify By Sindy 2014/4/29 人員輸入時數後又讓系統重算蓋掉
''               'txtB1013.Text = dHour
''               If Val(txtB1013.Text) = 0 Then
''                  txtB1013.Text = dHour
''               End If
''               '2014/4/29 END
''               txtB1013.Enabled = True
'''             End If
''         Else '平日時數
'''             If txtB1012.Text = "" Or txtB1012.Text <= "0" Then
''               'Modify By Sindy 2014/4/29 人員輸入時數後又讓系統重算蓋掉
''               'txtB1012.Text = dHour
''               If Val(txtB1012.Text) = 0 Then
''                  txtB1012.Text = dHour
''               End If
''               '2014/4/29 END
''               txtB1012.Enabled = True
''               txtB1013.Text = ""
''               txtB1013.Enabled = False
'''             End If
''         End If
'      End If
''      '記錄計算完畢當時的日期及時間,方便比對是否有需要重新計算
''      m_B1004 = Val(txtB1004)
''      m_B1005_1 = Val(txtB1005_1)
''      m_B1005_2 = Val(txtB1005_2)
''      m_B1006 = ""
''      m_B1007_1 = Val(txtB1007_1)
''      m_B1007_2 = Val(txtB1007_2)
''      m_B1014 = ""
''      m_B1028 = ""
''      m_B1029 = ""
''      m_B1008 = ""
   End If
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String, ByVal strKEY04 As String, ByVal strKEY05 As String) As Boolean
   IsRecordExist = False
   
   If IsNull(strKEY01) Or strKEY01 = "" Then Exit Function
   If IsNull(strKEY02) Or strKEY02 = "" Then Exit Function
   If IsNull(strKEY03) Or strKEY03 = "" Then Exit Function
   If IsNull(strKEY04) Or strKEY04 = "" Then
      strKEY04 = strKEY02
      strKEY05 = strKEY03
   End If
   
   If CheckIsAbsenceExist(strKEY01, strKEY02, strKEY03, strKEY04, strKEY05, txtB1001, Left(Trim(CboB1002), 2)) = True Then IsRecordExist = True
   If IsRecordExist = True Then
      MsgBox "該筆記錄已存在", vbOKOnly, "新增資料"
      '先清空欄位值
      txtB1009 = Empty
      txtB1010 = Empty
      txtB1030 = Empty
      txtB101213 = Empty
      Call ClearFieldCbo
   End If
End Function

Private Sub cboSTime_GotFocus()
'   InverseTextBox cboSTime
End Sub

Private Sub cboSTime_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub cboSTime_Validate(Cancel As Boolean)
If Frame1.Visible = True Then
   If cboSTime <> "" Then
   '   If Val(Format(cboSTime.Text, "hhmm")) > 2400 Then
   '      Call cboSTime_GotFocus
   '      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
   '      Cancel = True
   '      Exit Sub
   '   End If
      'Modify By Sindy 2011/11/16
      If txtB1004 = txtB1006 Then
         cboETime.ListIndex = cboSTime.ListIndex
      End If
      '2011/11/16 End
      'Modify By Sindy 2018/2/9
      If GetCountDayHour(False) = False Then
'         Call cboSTime_GotFocus
'         Cancel = True
'         Exit Sub
      End If
   Else
'      If Chk1Day.Value = 1 Then '非整日
'         If txtB1004 <> txtB1006 then
            If cboSTime = "" Then
               Call cboSTime_GotFocus
               MsgBox "請輸入起日上班時段!", vbExclamation + vbOKOnly
               Cancel = True
               Exit Sub
            End If
'         End If
'      End If
   End If
End If
CloseIme
End Sub

Private Sub cboETime_GotFocus()
'   InverseTextBox cboETime
End Sub

Private Sub cboETime_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub cboETime_Validate(Cancel As Boolean)
If Frame1.Visible = True Then
   If cboETime <> "" Then
   '   If Val(Format(cboETime.Text, "hhmm")) > 2400 Then
   '      Call cboETime_GotFocus
   '      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
   '      Cancel = True
   '      Exit Sub
   '   End If
      'Modify By Sindy 2011/11/16
      If txtB1004 = txtB1006 Then
         cboSTime.ListIndex = cboETime.ListIndex
      End If
      '2011/11/16 End
      'Modify By Sindy 2018/2/9
      If GetCountDayHour(False) = False Then
'         Call cboETime_GotFocus
'         Cancel = True
'         Exit Sub
      End If
   Else
'      If Chk1Day.Value = 1 Then '非整日
'         If txtB1004 <> txtB1006 Then
            If cboETime = "" Then
               Call cboETime_GotFocus
               MsgBox "請輸入迄日下班時段!", vbExclamation + vbOKOnly
               Cancel = True
               Exit Sub
            End If
'         End If
'      End If
   End If
End If
CloseIme
End Sub

'設定職務代理人的下拉式選單
Private Sub SetABS001_1Combo(StrST01 As String)
Dim strText As String
Dim kk As Integer
   
   For i = 0 To cboEmp.UBound
      cboEmp(i).Clear
      cboEmp(i).AddItem ""
   Next i
   strSql = "SELECT B0102,1 FROM ABS001,Staff WHERE B0101='" & StrST01 & "' AND B0102 is not null AND B0102=ST01(+) AND ST04='1' " & _
      "Union SELECT B0103,2 FROM ABS001,Staff WHERE B0101='" & StrST01 & "' AND B0103 is not null AND B0103=ST01(+) AND ST04='1' " & _
      "Union SELECT B0104,3 FROM ABS001,Staff WHERE B0101='" & StrST01 & "' AND B0104 is not null AND B0104=ST01(+) AND ST04='1' " & _
      "Union SELECT B0105,4 FROM ABS001,Staff WHERE B0101='" & StrST01 & "' AND B0105 is not null AND B0105=ST01(+) AND ST04='1' " & _
      "Union SELECT B0106,5 FROM ABS001,Staff WHERE B0101='" & StrST01 & "' AND B0106 is not null AND B0106=ST01(+) AND ST04='1' " & _
      "Union SELECT B0107,6 FROM ABS001,Staff WHERE B0101='" & StrST01 & "' AND B0107 is not null AND B0107=ST01(+) AND ST04='1' " & _
      "order by 2 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         bolSetCboEmp = True 'Add By Sindy 2018/8/3
         .MoveFirst
         Do While Not .EOF
            If Not IsNull(RsTemp.Fields(0)) Then
               strText = SetCboStaffName(RsTemp.Fields(0))
               For i = 0 To cboEmp.UBound
                  cboEmp(i).AddItem strText
               Next i
            End If
            .MoveNext
         Loop
      End With
   Else
      bolSetCboEmp = False 'Add By Sindy 2018/8/3
   End If
   For i = 0 To cboEmp.UBound
      If cboEmp(i).ListCount > 0 Then cboEmp(i).ListIndex = 0
   Next i
   
   'Modify By Sindy 2017/1/10 王副總提若職代的請假時間含蓋了請假人的請假時間, 則不可以出現
'   If (cmdSend.Visible = True And cmdSend.Enabled = True) Or _
'      (cmdagainSend.Visible = True And cmdagainSend.Enabled = True) Then
      If txtB1004 <> "" And txtB1005_1 <> "" And txtB1005_2 <> "" And _
         txtB1006 <> "" And txtB1007_1 <> "" And txtB1007_2 <> "" Then
         For i = 0 To cboEmp.UBound
            For kk = cboEmp(i).ListCount - 1 To 0 Step -1
               If Trim(cboEmp(i).List(kk)) <> "" Then
                  If CheckIsPersonRestSectorSame(CStr(Left(Trim(cboEmp(i).List(kk)), 5)), txtB1004, Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), txtB1006, Trim(txtB1007_1.Text & ":" & txtB1007_2.Text), txtB1001) = True Then
                     cboEmp(i).RemoveItem (kk)
                  End If
               End If
            Next kk
         Next i
      End If
'   End If
   '2017/1/10 END
End Sub

'設定審核主管的下拉式選單
Private Sub SetABS001_2Combo(StrST01 As String)
Dim strText As String
   
   For i = 0 To CboBoss.UBound
      CboBoss(i).Clear
      CboBoss(i).AddItem ""
   Next i
   strSql = "SELECT B0108,1 FROM ABS001,Staff WHERE B0101='" & StrST01 & "' AND B0108 is not null AND B0108=ST01(+) AND ST04='1' " & _
      "Union SELECT B0109,2 FROM ABS001,Staff WHERE B0101='" & StrST01 & "' AND B0109 is not null AND B0109=ST01(+) AND ST04='1' " & _
      "Union SELECT B0110,3 FROM ABS001,Staff WHERE B0101='" & StrST01 & "' AND B0110 is not null AND B0110=ST01(+) AND ST04='1' " & _
      "Union SELECT B0111,4 FROM ABS001,Staff WHERE B0101='" & StrST01 & "' AND B0111 is not null AND B0111=ST01(+) AND ST04='1' " & _
      "order by 2 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         bolSetCboEmpSir = True 'Add By Sindy 2018/8/3
         .MoveFirst
         Do While Not .EOF
            If Not IsNull(RsTemp.Fields(0)) Then
               strText = SetCboStaffName(RsTemp.Fields(0))
               For i = 0 To CboBoss.UBound
                  CboBoss(i).AddItem strText
               Next i
            End If
            .MoveNext
         Loop
      End With
   Else
      bolSetCboEmpSir = False 'Add By Sindy 2018/8/3
   End If
   For i = 0 To CboBoss.UBound
      If CboBoss(i).ListCount > 0 Then CboBoss(i).ListIndex = 0
   Next i
End Sub

'Add By Sindy 2016/12/28 計算加班時數
Private Sub txtB1030_Click()
   Call AutoCount
End Sub

Private Sub txtB1030_GotFocus()
   InverseTextBox txtB1030
   CloseIme
End Sub

Private Sub txtB1030_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub txtB1030_Validate(Cancel As Boolean)
Dim nResponse

If txtB1030 <> "" Then
   If CheckLengthIsOK(txtB1030, txtB1030.MaxLength) = False Then
      Call txtB1030_GotFocus
      Cancel = True
      Exit Sub
   End If
   'Add By Sindy 2015/11/19
   'Modify By Sindy 2016/12/28
   'If Val(dHour) = 0 Then Call CountHour(True) '為取得dHour的值
   Call CountHour(True) '為取得dHour的值
   '2016/12/28 END
   '2015/11/19 END
   If Val(dHour) > 0 Then
      If Val(txtB1030) > Val(dHour) Then
         MsgBox "您輸入的加班時數" & txtB1030 & "大於系統計算時數" & dHour & "，只能改少!!!", vbExclamation + vbOKOnly
         Call txtB1030_GotFocus
         Cancel = True
         Exit Sub
      ElseIf Val(txtB1030) < Val(dHour) Then
         nResponse = MsgBox("系統計算時數為" & dHour & "，您確定改為" & txtB1030 & "嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問")
         If nResponse = vbNo Then
            txtB1030 = dHour
         End If
         'Modify By Sindy 2017/1/3 換算加班時數
         txtB101213 = PUB_Overtime_TransDay(txtB1004, txtB1003, txtB1030)
         '2017/1/3 END
      End If
   End If
End If
CloseIme
End Sub
