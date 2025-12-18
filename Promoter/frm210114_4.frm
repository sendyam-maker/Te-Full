VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210114_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "委任契約書-CFT 及 T大陸案 及 TF馬德里案"
   ClientHeight    =   7090
   ClientLeft      =   650
   ClientTop       =   1840
   ClientWidth     =   10220
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7090
   ScaleWidth      =   10220
   Begin VB.CheckBox ChkSeal 
      Caption         =   "用印"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   80
      Top             =   6690
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0FFFF&
      Caption         =   "空白列印"
      Height          =   330
      Index           =   5
      Left            =   3720
      Style           =   1  '圖片外觀
      TabIndex        =   79
      Top             =   30
      Width           =   920
   End
   Begin VB.CheckBox ChkDou 
      Caption         =   "多申請人"
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   9000
      TabIndex        =   78
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm210114_4.frx":0000
      Left            =   7200
      List            =   "frm210114_4.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   31
      Top             =   6690
      Width           =   2475
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋申請人(&Q)"
      Height          =   330
      Left            =   8880
      TabIndex        =   76
      Top             =   1050
      Width           =   1365
   End
   Begin VB.TextBox txtPCnt 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   32
      Text            =   "2"
      Top             =   60
      Width           =   270
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印        份"
      Height          =   330
      Index           =   0
      Left            =   7416
      TabIndex        =   37
      Top             =   30
      Width           =   1100
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面"
      Height          =   330
      Index           =   1
      Left            =   8520
      TabIndex        =   38
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "清空資料"
      Height          =   330
      Index           =   2
      Left            =   6492
      TabIndex        =   36
      Top             =   30
      Width           =   920
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   360
      Left            =   30
      TabIndex        =   74
      Top             =   0
      Width           =   3645
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   660
         Style           =   2  '單純下拉式
         TabIndex        =   33
         Top             =   30
         Width           =   2940
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "印表機"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   75
         Top             =   90
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "儲存文檔"
      Height          =   330
      Index           =   3
      Left            =   4644
      TabIndex        =   34
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "讀取文檔"
      Height          =   330
      Index           =   4
      Left            =   5568
      TabIndex        =   35
      Top             =   30
      Width           =   920
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9450
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   30
      Left            =   1575
      TabIndex        =   9
      Top             =   3240
      Width           =   7260
      VariousPropertyBits=   671105051
      MaxLength       =   76
      Size            =   "12806;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   29
      Left            =   1095
      TabIndex        =   20
      Top             =   4845
      Width           =   1035
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1826;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   25
      Left            =   600
      TabIndex        =   16
      Top             =   4485
      Width           =   1665
      VariousPropertyBits=   671105051
      MaxLength       =   20
      BorderStyle     =   1
      Size            =   "2937;529"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   26
      Left            =   2250
      TabIndex        =   17
      Top             =   4485
      Width           =   1995
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "3519;529"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   27
      Left            =   4230
      TabIndex        =   18
      Top             =   4485
      Width           =   1665
      VariousPropertyBits=   671105051
      MaxLength       =   20
      BorderStyle     =   1
      Size            =   "2937;529"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   28
      Left            =   5880
      TabIndex        =   19
      Top             =   4485
      Width           =   1995
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "3519;529"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   24
      Left            =   1575
      TabIndex        =   8
      Top             =   2925
      Width           =   7260
      VariousPropertyBits=   671105051
      MaxLength       =   76
      Size            =   "12806;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   23
      Left            =   1580
      TabIndex        =   6
      Top             =   2300
      Width           =   7260
      VariousPropertyBits=   671105051
      MaxLength       =   76
      Size            =   "12806;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   22
      Left            =   1575
      TabIndex        =   4
      Top             =   1665
      Width           =   7260
      VariousPropertyBits=   671105051
      MaxLength       =   76
      Size            =   "12806;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   21
      Left            =   1440
      TabIndex        =   2
      Top             =   1035
      Width           =   7395
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "13044;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   11
      Left            =   2295
      TabIndex        =   21
      ToolTipText     =   "請輸入數字"
      Top             =   4845
      Width           =   2145
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "3784;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   2865
      TabIndex        =   11
      Top             =   3615
      Width           =   1035
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1826;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1980
      TabIndex        =   10
      Top             =   3615
      Width           =   675
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1191;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   20
      Left            =   4200
      TabIndex        =   30
      Top             =   6675
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1235;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   19
      Left            =   3120
      TabIndex        =   29
      Top             =   6675
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1235;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   18
      Left            =   2040
      TabIndex        =   28
      Top             =   6675
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1235;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   17
      Left            =   2040
      TabIndex        =   27
      Top             =   6370
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   16
      Left            =   5070
      TabIndex        =   26
      Top             =   6056
      Width           =   2445
      VariousPropertyBits=   671105051
      MaxLength       =   18
      Size            =   "4313;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   15
      Left            =   2040
      TabIndex        =   25
      Top             =   6056
      Width           =   2280
      VariousPropertyBits=   671105051
      MaxLength       =   22
      Size            =   "4022;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   14
      Left            =   2040
      TabIndex        =   24
      Top             =   5760
      Width           =   7995
      VariousPropertyBits=   671105051
      MaxLength       =   88
      Size            =   "14102;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   13
      Left            =   2040
      TabIndex        =   23
      Top             =   5455
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   48
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   12
      Left            =   2040
      TabIndex        =   22
      Top             =   5150
      Width           =   5475
      VariousPropertyBits=   671105051
      MaxLength       =   48
      Size            =   "9657;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   5880
      TabIndex        =   15
      Top             =   4200
      Width           =   1995
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "3519;529"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   4230
      TabIndex        =   14
      Top             =   4200
      Width           =   1665
      VariousPropertyBits=   671105051
      MaxLength       =   20
      BorderStyle     =   1
      Size            =   "2937;529"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   2250
      TabIndex        =   13
      Top             =   4200
      Width           =   1995
      VariousPropertyBits=   671105051
      MaxLength       =   10
      BorderStyle     =   1
      Size            =   "3519;529"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   600
      TabIndex        =   12
      Top             =   4200
      Width           =   1665
      VariousPropertyBits=   671105051
      MaxLength       =   20
      BorderStyle     =   1
      Size            =   "2937;529"
      SpecialEffect   =   0
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1575
      TabIndex        =   7
      Top             =   2610
      Width           =   7260
      VariousPropertyBits=   671105051
      MaxLength       =   76
      Size            =   "12806;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1575
      TabIndex        =   5
      Top             =   1980
      Width           =   7260
      VariousPropertyBits=   671105051
      MaxLength       =   76
      Size            =   "12806;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1575
      TabIndex        =   3
      Top             =   1350
      Width           =   7260
      VariousPropertyBits=   671105051
      MaxLength       =   76
      Size            =   "12806;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   7395
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "13044;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1140
      TabIndex        =   0
      Top             =   405
      Width           =   7695
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "13573;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "受任人："
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   6450
      TabIndex        =   77
      Top             =   6720
      Width           =   720
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "6.其他"
      Height          =   180
      Left            =   9195
      TabIndex        =   73
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      Caption         =   "5.救濟程序"
      Height          =   180
      Left            =   9195
      TabIndex        =   72
      Top             =   2460
      Width           =   855
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "4.領證程序"
      Height          =   180
      Left            =   9195
      TabIndex        =   71
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      Caption         =   "3.中間程序"
      Height          =   180
      Left            =   9195
      TabIndex        =   70
      Top             =   2100
      Width           =   855
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      Caption         =   "2.申請程序"
      Height          =   180
      Left            =   9195
      TabIndex        =   69
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "第二條 委辦範圍"
      Height          =   180
      Left            =   8880
      TabIndex        =   68
      Top             =   1500
      Width           =   1305
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "1.查名程序"
      Height          =   180
      Left            =   9195
      TabIndex        =   67
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "請輸入數字"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   7980
      TabIndex        =   66
      Top             =   4875
      Width           =   900
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "(英)"
      Height          =   180
      Left            =   1245
      TabIndex        =   65
      Top             =   2985
      Width           =   300
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "(中)"
      Height          =   180
      Left            =   1245
      TabIndex        =   64
      Top             =   2670
      Width           =   300
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "(英)"
      Height          =   180
      Left            =   1245
      TabIndex        =   63
      Top             =   2355
      Width           =   300
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "(中)"
      Height          =   180
      Left            =   1245
      TabIndex        =   62
      Top             =   2040
      Width           =   300
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "(英)"
      Height          =   180
      Left            =   1245
      TabIndex        =   61
      Top             =   1710
      Width           =   300
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "(中)"
      Height          =   180
      Left            =   1245
      TabIndex        =   60
      Top             =   1410
      Width           =   300
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "(英)"
      Height          =   180
      Left            =   1140
      TabIndex        =   59
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "(中)"
      Height          =   180
      Left            =   1140
      TabIndex        =   58
      Top             =   780
      Width           =   300
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   1  '單線固定
      Caption         =   "金額"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5880
      TabIndex        =   57
      Top             =   3960
      Width           =   1995
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   1  '單線固定
      Caption         =   "金額"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2250
      TabIndex        =   56
      Top             =   3960
      Width           =   1995
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   1  '單線固定
      Caption         =   "國　　別"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4230
      TabIndex        =   55
      Top             =   3960
      Width           =   1665
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      BorderStyle     =   1  '單線固定
      Caption         =   "國　　別"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   600
      TabIndex        =   54
      Top             =   3960
      Width           =   1665
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "列舉商品："
      Height          =   180
      Left            =   165
      TabIndex        =   53
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "中　華　民　國　　　　　年　　　　　月　　　　　日"
      Height          =   180
      Left            =   690
      TabIndex        =   52
      Top             =   6720
      Width           =   4500
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "乙方：經手人："
      Height          =   180
      Left            =   675
      TabIndex        =   51
      Top             =   6418
      Width           =   1260
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "傳　真："
      Height          =   180
      Left            =   4335
      TabIndex        =   50
      Top             =   6116
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "電　話："
      Height          =   180
      Left            =   1215
      TabIndex        =   49
      Top             =   6116
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "地　址："
      Height          =   180
      Left            =   1215
      TabIndex        =   48
      Top             =   5814
      Width           =   720
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代表人："
      Height          =   180
      Left            =   1215
      TabIndex        =   47
      Top             =   5512
      Width           =   720
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "甲方：委任人："
      Height          =   180
      Left            =   675
      TabIndex        =   46
      Top             =   5210
      Width           =   1260
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "元整，於本契約簽定同時由甲方一次付清。"
      Height          =   180
      Left            =   4515
      TabIndex        =   45
      Top             =   4875
      Width           =   3420
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "合計(幣別)"
      Height          =   180
      Left            =   252
      TabIndex        =   44
      Top             =   4860
      Width           =   840
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "一、乙方受委辦前條第　　　　款　　　　　　程序之費用（包括國外代理人費用），約定如下："
      Height          =   180
      Left            =   165
      TabIndex        =   43
      Top             =   3660
      Width           =   7740
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請人地址："
      Height          =   180
      Left            =   165
      TabIndex        =   42
      Top             =   2850
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "代  表  人："
      Height          =   180
      Left            =   165
      TabIndex        =   41
      Top             =   2190
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請人名稱："
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   40
      Top             =   1530
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   165
      TabIndex        =   39
      Top             =   450
      Width           =   900
   End
End
Attribute VB_Name = "frm210114_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/22 改成Form2.0 ; txt1(index)、Printer改成Word列印
'Memo by Lydia 2019/07/01 表單名稱:案件委任契約書=>委任契約書
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim SeekPrint As Integer, SeekPrintL As Integer
Dim iCount As Integer
'Add By Sindy 2010/4/30
Public m_strCustCode As String
Public m_blnOneRec As Boolean
'2010/4/30 End
Dim strNowCustNo As String 'Add by Amy 2016/08/19 客戶編號
Dim iPrintC As Integer 'Added by Lydia 2017/03/28 目前列印第幾份
Dim bolAddSeal As Boolean 'Added by Lydia 2017/03/28 是否用印
Dim d_Left As Double, d_Top As Double 'Added by Lydia 2017/04/25 印表機實際列印的左邊界、右邊界
Dim strPrinter As String 'Added by Lydia 2017/04/28
Dim strDetail As String 'Move by Lydia 2017/05/16 記錄內容(從StrMenu移出來)
Dim strCompSeal As String 'Added by Lydia 2020/03/25 記錄"公司名稱|用印編號",用,區隔
'Added by Lydia 2022/01/24  加入圖片用(Word)
Const msoFalse = 0
Const msoLineSolid = 1
Const msoLineSingle = 1
Const msoTrue = -1
Const msoPictureAutomatic = 1
'end 2022/01/24
Dim m_TempPDF As String 'Added by Lydia 2022/01/24
Dim m_TempFN As String 'Added by Lydia 2022/01/24

'Add By Sindy 2010/4/29
Private Sub cmdFind_Click()
   Dim strCmpName As String, strMsg As String 'Add by Amy 2016/08/19
   
   If Me.txt1(2).Text = "" Then
      MsgBox "請輸入申請人中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txt1(2).SetFocus
      Exit Sub
   End If
   
   frm090801_1.m_Type = 0  'add by Lydia 2014/9/22
   If ChkDou.Value = 1 Then
     frm090801_1.m_DouChk = True '可複選
   Else
     frm090801_1.m_DouChk = False
   End If
   Set frm090801_1.m_frm0908A = Me
   
   frm090801_1.m_strCustChnName = Me.txt1(2).Text
   frm090801_1.lblName.Caption = Me.txt1(2).Text
   
   m_blnOneRec = False
   m_strCustCode = ""
   If frm090801_1.StrMenu = True Then
      If frm090801_1.m_blnOneRec = False Then
         frm090801_1.Show vbModal
      End If
      m_blnOneRec = frm090801_1.m_blnOneRec
      m_strCustCode = frm090801_1.m_strCustCode
      Unload frm090801_1
   Else
      Unload frm090801_1
   End If
    Combo2.Tag = "": strNowCustNo = "" 'Add by Amy 2016/08/19
   If m_blnOneRec = True And m_strCustCode <> "" Then
     'Add by Amy 2016/08/19 記錄收據公司別(放於SetCustTxt前避免m_strCustCode被清空)
      strNowCustNo = m_strCustCode
      strCmpName = "Y"
      Combo2.Tag = GetReceiptCmp(Left(strNowCustNo, 8), Mid(strNowCustNo, 9, 1), "CFT", "020", False, strCmpName, Me.Name)
      If Combo2.Tag <> MsgText(601) And Combo2 <> MsgText(601) And Combo2.Tag <> frm210114_1.GetComp(Combo2) Then
        strMsg = "您輸入之收據公司別「" & Combo2 & "」與客戶檔設定值「" & strCmpName & "」不同" & vbCrLf & _
                     "是否依客戶檔設定覆蓋您的輸入值？"
        If MsgBox(strMsg, vbYesNo + vbCritical) = vbYes Then
            'Modified by Lydia 2024/08/06
            'Combo2 = strCmpName
            Call Pub_SetCboListIdx(Me.Combo2, strCmpName)
        End If
      ElseIf strCmpName = MsgText(601) Then
        Combo2.ListIndex = 0
      Else
        'Modified by Lydia 2024/08/06
        'Combo2 = strCmpName
        Call Pub_SetCboListIdx(Me.Combo2, strCmpName)
      End If
      'end 2016/08/19
      'Modify by Amy 2021/05/13 +if 讀取文檔要保留原文檔內容
      If Me.ActiveControl.Name = "cmdFind" Then
        Call SetCustTxt(m_strCustCode)
      End If
      txt1(2).Tag = txt1(2)
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim tb As Control
Dim op As OptionButton
Dim fN As Integer
Dim strBuffer As String
'Modified by Lydia 2023/02/17
'Dim AllObj(0 To 30) As String
Dim AllObj(0 To 31) As String
Dim AllObjV As Variant
'Add by Amy 2016/08/19 目前收據公司別
Dim strNowCmp As String

   Select Case Index
      Case 0
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(0)) = "" Then
              MsgBox "案件名稱不可空白！", vbInformation, "錯誤！"
              txt1(0).SetFocus
              txt1_GotFocus 0
              Exit Sub
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(1)) = "" And Trim(txt1(21)) = "" Then
              MsgBox "列舉商品不可空白！", vbInformation, "錯誤！"
              txt1(1).SetFocus
              txt1_GotFocus 1
              Exit Sub
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(2)) = "" And Trim(txt1(22)) = "" Then
              MsgBox "申請人名稱不可空白！", vbInformation, "錯誤！"
              txt1(2).SetFocus
              txt1_GotFocus 2
              Exit Sub
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(3)) = "" And Trim(txt1(23)) = "" Then
              MsgBox "代表人不可空白！", vbInformation, "錯誤！"
              txt1(3).SetFocus
              txt1_GotFocus 3
              Exit Sub
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(7)) = "" And Trim(txt1(8)) = "" And Trim(txt1(9)) = "" And Trim(txt1(10)) = "" Then
              MsgBox "約定事項最少輸入一項！", vbInformation, "錯誤！"
              txt1(7).SetFocus
              txt1_GotFocus 7
              Exit Sub
          End If
          If (txt1(7) = "" Or txt1(8) = "") And txt1(7) & txt1(8) <> "" Then
              MsgBox "約定事項應該完整！", vbInformation, "錯誤！"
              If txt1(7) = "" Then
                  txt1(7).SetFocus
                  txt1_GotFocus 7
              End If
              If txt1(8) = "" Then
                  txt1(8).SetFocus
                  txt1_GotFocus 8
              End If
              Exit Sub
          End If
          If (txt1(9) = "" Or txt1(10) = "") And txt1(9) & txt1(10) <> "" Then
              MsgBox "約定事項應該完整！", vbInformation, "錯誤！"
              If txt1(9) = "" Then
                  txt1(9).SetFocus
                  txt1_GotFocus 9
              End If
              If txt1(10) = "" Then
                  txt1(10).SetFocus
                  txt1_GotFocus 10
              End If
              Exit Sub
          End If
          If txt1(11) = "" Then
              MsgBox "費用不可空白！", vbInformation, "錯誤！"
              txt1(11).SetFocus
              txt1_GotFocus 11
              Exit Sub
          End If
      '    If txt1(12) = "" Then
      '        MsgBox "委任人不可空白！", vbInformation, "錯誤！"
      '        txt1(12).SetFocus
      '        txt1_GotFocus 12
      '        Exit Sub
      '    End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(17)) = "" Then
              MsgBox "經手人不可空白！", vbInformation, "錯誤！"
              txt1(17).SetFocus
              txt1_GotFocus 17
              Exit Sub
          End If
                
         '2011/10/18 ADD BY SONIA 檢查四縣市地址
         If txt1(4) <> "" Then
           If CheckTaiwanAddr(txt1(4), "000", "申請人地址") = False Then
              txt1(4).SetFocus
              txt1_GotFocus (4)
              Exit Sub
           End If
         End If
         If txt1(14) <> "" Then
           If CheckTaiwanAddr(txt1(14), "000", "甲方委任人地址") = False Then
              txt1(14).SetFocus
              txt1_GotFocus (14)
              Exit Sub
           End If
         End If
         '2011/10/18 END
         'Add by Amy 2016/08/19 +受任人不可為空
         If Combo2 = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
         End If
         
         'Add By Sindy 2013/12/15 檢查是否為不開發票之客戶
          If Combo2 = "台一智權股份有限公司" Then
            'Added by Lydia 2020/07/20 國別若有"台灣"或"中華民國"時，受任人不可選擇"台一智權股份有限公司"
            If Trim(txt1(7) & txt1(25)) <> "" And (InStr(txt1(7) & txt1(25), "台灣") > 0 Or InStr(txt1(7) & txt1(25), "臺灣") > 0 _
                   Or InStr(txt1(7) & txt1(25), "中華民國") > 0) Then
               MsgBox "國別若有""台灣""或""中華民國""時，受任人不可選擇智權公司!!!", vbInformation
               Combo2.SetFocus
               Exit Sub
            End If
            'end 2020/07/20
            'Modify By Sindy 2024/9/25 增加傳入公司別做判斷
            If PUB_ChkCU144isN("", "", txt1(2), "J", , "受任人") = True Then
               Combo2.SetFocus
               Exit Sub
            End If
          End If
          '2013/12/15 END
           '2009/11/13 MODIFY BY SONIA 杜副總提出
      '    If txt1(18) = "" Or txt1(19) = "" Or txt1(20) = "" Then
      '        MsgBox "日期需要正確！", vbInformation, "錯誤！"
      '        txt1(18).SetFocus
      '        txt1_GotFocus 18
      '        Exit Sub
      '    End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(18)) = "" Or Trim(txt1(19)) = "" Or Trim(txt1(20)) = "" Then
             If MsgBox("契約書日期不完整，是否確定？", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
               txt1(18).SetFocus
               txt1_GotFocus 18
               Exit Sub
             End If
          End If
      '2009/11/13 END
                
          'Added by Lydia 2017/03/28
          If ChkSeal.Value = 1 Then
            'Modified by Lydia 2017/04/27 PDF印表機不需詢問,並且份數改為1份
            If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
            Else
               If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If
            'end 2017/04/27
             bolAddSeal = True
          Else
             bolAddSeal = False
          End If
          'end 2017/03/28
          
          'Modified by Lydia 2017/04/13
'          For iCount = 1 To Val(txtPCnt) 'edit by nickc 2006/09/27 2
'              'add by nickc 2006/06/05
'              Set Printer = Printers(Combo1.ListIndex)
'              Screen.MousePointer = vbHourglass
'              DoEvents
'              StrMenu
'          Next iCount
          'Modified by Lydia 2022/01/25 改成Word直接印
          'Call Print2PDF(False)
          Call runWordProc(False)
          PUB_SetOsDefaultPrinter strPrinter
          'end 2022/01/25
          
          'Add by Amy 2016/08/19 畫面與客戶檔收據公司別不同更新客戶檔
          strNowCmp = frm210114_1.GetComp(Combo2)
          If Combo2.Tag <> strNowCmp Then
             Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          End If
          'end 2016/08/19
          Screen.MousePointer = vbDefault
          Call RunEndProc(True) 'Added by Lydia 2022/01/24 刪除暫存檔
          'Modified by Lydia 2022/01/24 判斷是否有列印
          'ShowPrintOk 'Added by Lydia 2017/04/11
          If m_TempPDF <> "" Then ShowPrintOk
      Case 1
          frm210114.Show
          Unload Me
      Case 2
          For Each tb In txt1
              tb.Text = Empty
          Next
      Case 3
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowSave
          If cd1.FileName <> "" Then
              AllObj(0) = "案件委任契約書-CFT"
              'Modified by Lydia 2023/02/17
              'For iCount = 1 To 29
              For iCount = 1 To 30
                  AllObj(iCount) = txt1(iCount - 1).Text
              Next iCount
              'Modified by Lydia 2023/02/17
              'AllObj(30) = Combo2.Text 'Add By Sindy 2011/3/23
              AllObj(31) = Combo2.Text
              strBuffer = Join(AllObj, Chr(30))
              strBuffer = StrEncrypt(strBuffer)
              fN = FreeFile
              Open cd1.FileName For Output As fN
              Print #fN, strBuffer
              Close #fN
          End If
          'Add by Amy 2016/08/19 畫面與客戶檔收據公司別不同更新客戶檔
          strNowCmp = frm210114_1.GetComp(Combo2)
          If Combo2 <> MsgText(601) And Combo2.Tag <> strNowCmp Then
             Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          End If
          'end 2016/08/19
      Case 4
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowOpen
          If cd1.FileName <> "" Then
              fN = FreeFile
              Open cd1.FileName For Input As fN
              Input #fN, strBuffer
              Close #fN
              strBuffer = StrDecrypt(strBuffer)
              AllObjV = Split(strBuffer, Chr(30))
              If AllObjV(0) = "案件委任契約書-CFT" Then
                  cmdOK_Click 2
                  'Modified by Lydia 2023/02/17
                  'For iCount = 1 To 29
                  For iCount = 1 To 30
                       txt1(iCount - 1).Text = AllObjV(iCount)
                  Next iCount
                  'Modify by Amy 2016/08/19 避免空值會Error
                  'Modified by Lydia 2023/02/17
                  'If AllObjV(30) = MsgText(601) Then
                  If AllObjV(31) = MsgText(601) Then
                    Combo2.ListIndex = 0
                  Else
                    'Modified by Lydia 2023/02/17
                    'Combo2.Text = AllObjV(30) 'Add By Sindy 2011/3/23
                    Combo2.Text = AllObjV(31)
                  End If
                  'end 2016/08/19
                  'Add By Sindy 2011/1/21 檢查地址欄
                  '申請人地址(中)
                  If txt1(2).Text <> "" And txt1(4).Text <> "" Then
                     If CheckCustomerAddr(1, Trim(txt1(2).Text), Trim(txt1(4).Text), "申請人中文", True) = False Then
                        txt1(4).SetFocus
                     End If
                  End If
                  '申請人地址(英)
                  If txt1(22).Text <> "" And txt1(24).Text <> "" Then
                     If CheckCustomerAddr(2, Trim(txt1(22).Text), Trim(txt1(24).Text), "申請人英文", True) = False Then
                        txt1(24).SetFocus
                     End If
                  End If
                  '委任人地址
                  If txt1(12).Text <> "" And txt1(14).Text <> "" Then
                     If CheckCustomerAddr(1, Trim(txt1(12).Text), Trim(txt1(14).Text), "委任人", True) = False Then
                        txt1(14).SetFocus
                     End If
                  End If
                  '2011/1/21 End
                  'Add by Amy 2016/08/19 讀取收據公司別
                  cmdFind_Click
              Else
                  MsgBox "錯誤格式，此份內容並非 CFT 格式！", vbExclamation
              End If
          End If
      'Added by Lydia 2017/03/28 空白委任書
      Case 5
          If Trim(Combo2) = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
          End If
          'Modified by Lydia 2017/04/17 文雄表示用印由下方勾選,可直接空白列印
          If ChkSeal.Value = 1 Then
            If (InStr(UCase(Combo1.Text), "BATCH") > 0 Or InStr(UCase(Combo1.Text), "WRITER") > 0 Or InStr(UCase(Combo1.Text), "PDF") > 0) And Pub_StrUserSt03 <> "M51" Then
               MsgBox "空白用印的印表機不可選擇PDF列印！", vbInformation, "錯誤！"
               Combo1.SetFocus
               Exit Sub
            End If
            'Modified by Lydia 2017/04/27 PDF印表機不需詢問,並且份數改為1份
            If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
            Else
               If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If
            'end 2017/04/27
            bolAddSeal = True
          End If
          'end 2017/04/17
          Call cmdOK_Click(2) '清空資料
          'Modified by Lydia 2022/01/25 改成Word直接印
          'Call Print2PDF(True)
          Call runWordProc(True)
          PUB_SetOsDefaultPrinter strPrinter
          'end 2022/01/25
          
          m_strCustCode = ""
          bolAddSeal = False
          Screen.MousePointer = vbDefault
          Call RunEndProc(True) 'Added by Lydia 2022/01/24 刪除暫存檔
          'Modified by Lydia 2022/01/24 判斷是否有列印
          'ShowPrintOk 'Added by Lydia 2017/04/11
          If m_TempPDF <> "" Then ShowPrintOk
      'end 2017/03/28
      Case Else
   End Select
   Exit Sub
DialogCancel:
End Sub

'Add by Morgan 2011/2/24 只要鍵盤有動作就不斷線
Private Sub Form_KeyPress(KeyAscii As Integer)
   If UCase(Forms(0).Name) = "MDIMAIN" Then Forms(0).tmrConnect.Tag = 0
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
   
   PUB_InitForm210114 Forms(0), Me 'Added by Lydia 2017/05/19 委任契約書表單大於主表單，控制主表單放大。
   MoveFormToCenter Me
   'Modified by Lydia 2017/04/28 改用模組
   'strSql = Printer.DeviceName
   'SeekPrintL = Printer.Orientation
   'For i = 0 To Printers.Count - 1
   '    Set Printer = Printers(i)
   '    Combo1.AddItem Printer.DeviceName, j
   '    j = j + 1
   '    If Printer.DeviceName = strSql Then
   '        SeekPrint = i
   '    End If
   'Next i
   'Set Printer = Printers(SeekPrint)
   'Combo1.Text = Combo1.List(SeekPrint)
   PUB_SetPrinter Me.Name, Me.Combo1, strPrinter, , , , , True 'Modified by Morgan 2020/10/30 +只顯示有效的印表機參數

    'Added by Lydia 2017/04/17 先用模組抓所有印表機後,排除特定印表機
    'Remove by Lydia 2017/06/07 改直接列印
    'For i = 0 To Combo1.ListCount - 1
    '   If InStr(UCase(Combo1.List(i)), "PDFCREATOR") > 0 And Trim(Combo1.List(i)) <> "" Then
    '      Combo1.RemoveItem i
    '      'If i = SeekPrint Then Combo1.Text = Combo1.List(0) 'Remove by Lydia 2017/04/28
    '   End If
    'Next
    'end 2017/04/17
    'end 2017/06/07
    
   'Add By Sindy 2013/12/15
   'Remove Lydia 2020/03/25
   'If strSrvDate(1) >= InvoiceStartDate Then
   '   Combo2.AddItem "台一智權股份有限公司"
   'End If
   ''2013/12/15 END
   ''Modify by Amy 2016/08/19
   'Combo2.Text = Combo2.List(1) 'Add By Sindy 2011/3/23
   'end 2020/03/25
   
   'Added by Lydia 2020/03/25 設定公司別下拉選項
   Call PUB_SetCboTofrm210114(Me.Name, Me.Combo2, strCompSeal)
   
   Combo2.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '還原預設印表機
   'Modified by Lydia 2017/04/28 記錄表單的印表機
   'Set Printer = Printers(SeekPrint)
   'Printer.Orientation = SeekPrintL
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   'end 2017/04/28
   
   Call RunEndProc(False) 'Added by Lydia 2022/01/24 刪除暫存檔
   
   Set frm210114_4 = Nothing
End Sub

'Modified by Lydia 2017/03/28
'Sub StrMenu()
Sub StrMenu(Optional ByVal bolSpace As Boolean = False)
Dim iY As Integer
Dim tmpI As Integer
'Modified by Lydia 2020/08/10
'Dim iStr(1 To 54) As String
Dim iStr(1 To 55) As String
Dim tBoxTop As Integer
'Added by Lydia 2017/03/28
Dim tObj As New StdPicture
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String
'end 2017/03/28
'Added by Lydia 2020/08/10
Dim tmpPosY As Integer '公司章的Y軸起點
Dim tmpMaxY As Integer '版面的最大列數

   iStr(1) = "國外商標案件委任契約書"
   iStr(2) = "委任人(甲方)茲委任受任人(乙方)辦理國外商標案件，雙方同意條件如下："
   iStr(3) = "第一條" '??■□
   iStr(4) = "　　　　　　　　　　　　　　　　　　　　　　　　　 (中)：" & StrToStr(txt1(1) & String(38, " "), 19)
   iStr(5) = "    　案 件 名稱：" & StrToStr(StrToStr(txt1(0), 12) & String(24, " "), 12) & "　列舉商品"
   iStr(6) = "                  " & StrToStr(Replace(txt1(0), StrToStr(txt1(0), 12), "") & String(24, " "), 12) & "　　　　 (英)：" & StrToStr(txt1(21) & String(38, " "), 19)
   iStr(7) = "             (中)：" & StrToStr(txt1(2) & String(40, " "), 20) & "        (中)：" & StrToStr(txt1(3) & String(22, " "), 11)
   iStr(8) = "　　　申請人名稱　　　　　　　　　　　　　　　　　　　　　　代 表 人"
   iStr(9) = "             (英)：" & StrToStr(txt1(22) & String(40, " "), 20) & "        (英)：" & StrToStr(txt1(23) & String(22, " "), 11)
   iStr(10) = "               (中)：" & StrToStr(txt1(4) & String(74, " "), 37)
   iStr(11) = "　　　申請人住址"
   iStr(12) = "               (英)：" & StrToStr(txt1(24) & String(74, " "), 37)
   'Added by Lydia 2020/08/10
   If Trim(txt1(30)) <> "" Then
         iStr(13) = "                     " & StrToStr(txt1(30) & String(74, " "), 37)
   Else
         iStr(13) = "SPACEZERO"
   End If
   'end 2020/08/10
   'Modified by Lydia 2020/08/10 英文地址放成兩行,index + 1
   iStr(14) = "第二條　委辦範圍："
   iStr(15) = "　　1.查名程序：乙方根據甲方所提供之商標圖樣，委請各該查名國商標代理人代為查尋前案資料。"
   iStr(16) = "　　2.申請程序：乙方根據甲方所提供之資料代撰必要書件及製作印版、圖樣並為甲方委請各該申請國之"
   iStr(17) = "　　　　　　　　商標代理人，代向各該國提出商標申請。"
   iStr(18) = "　　3.中間程序：提出商標申請後，依甲方請求或各該申請國商標主管機關之指示所需提出之修正、補正、"
   iStr(19) = "　　　　　　　　更正等各程序。"
   iStr(20) = "　　4.領證程序：核准後之領取證書程序。"
   iStr(21) = "　　5.救濟程序：審定不予核准時所需進行之救濟程序。"
   iStr(22) = "　　6.其　　他：如討論案情、異議、撤銷、答辯、授權、移轉、陳報、延展、變更、參加聽證會、外國"
   iStr(23) = "　　　　　　　　商標主管機關文件及引證資料之翻譯等程序。"
   iStr(24) = "第三條　委辦費用"
   iStr(25) = "　　一、乙方受委辦前條第" & String(Int((6 - LenB(StrConv(txt1(5).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(5).Text) & String(6 - LenB(StrConv(String(Int((6 - LenB(StrConv(txt1(5).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(5).Text), vbFromUnicode)), " ") & "款" & String(Int((10 - LenB(StrConv(txt1(6).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(6).Text) & String(10 - LenB(StrConv(String(Int((10 - LenB(StrConv(txt1(6).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(6).Text), vbFromUnicode)), " ") & "程序之費用（包括國外代理人費用），約定如下："
   'Modified by Lydia 2017/03/28 幣別改成可輸入
   'iStr(26) = "　　　　　　　國　　別　　　　　　金額（新台幣）　　　　　國　　別　　　　　金額（新台幣）"
   iStr(26) = "　　　　　　　國　　別　　　　　　金額　　　　　　　　　　國　　別　　　　　金額　　　　　"
   iStr(27) = "　　　　" & String(20 - LenB(StrConv(String(Int((20 - LenB(StrConv(txt1(7).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(7).Text), vbFromUnicode)), " ") & Trim(txt1(7).Text) & String(Int((20 - LenB(StrConv(txt1(7).Text, vbFromUnicode))) / 2), " ") & String(24 - LenB(StrConv(String(Int((24 - LenB(StrConv(txt1(8).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(8).Text), vbFromUnicode)), " ") & Trim(txt1(8).Text) & String(Int((24 - LenB(StrConv(txt1(8).Text, vbFromUnicode))) / 2), " ") & String(Int((20 - LenB(StrConv(txt1(9).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(9).Text) & String(20 - LenB(StrConv(String(Int((20 - LenB(StrConv(txt1(9).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(9).Text), vbFromUnicode)), " ") & String(24 - LenB(StrConv(String(Int((24 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(10).Text), vbFromUnicode)), " ") & Trim(txt1(10).Text) & String(Int((24 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ")
   iStr(28) = "　　　　" & String(20 - LenB(StrConv(String(Int((20 - LenB(StrConv(txt1(25).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(25).Text), vbFromUnicode)), " ") & Trim(txt1(25).Text) & String(Int((20 - LenB(StrConv(txt1(25).Text, vbFromUnicode))) / 2), " ") & String(24 - LenB(StrConv(String(Int((24 - LenB(StrConv(txt1(26).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(26).Text), vbFromUnicode)), " ") & Trim(txt1(26).Text) & String(Int((24 - LenB(StrConv(txt1(26).Text, vbFromUnicode))) / 2), " ") & String(Int((20 - LenB(StrConv(txt1(27).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(27).Text) & String(20 - LenB(StrConv(String(Int((20 - LenB(StrConv(txt1(27).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(27).Text), vbFromUnicode)), " ") & String(24 - LenB(StrConv(String(Int((24 - LenB(StrConv(txt1(28).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(28).Text), vbFromUnicode)), " ") & Trim(txt1(28).Text) & String(Int((24 - LenB(StrConv(txt1(28).Text, vbFromUnicode))) / 2), " ")
   'Modified by Lydia 2017/03/28 幣別改成可輸入
   'iStr(29) = "　　　　合計新台幣　" & String(LenB(StrConv(ChangeNumber(txt1(11)), vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清。"
   If Val(Trim(txt1(11))) = 0 Then
       strSpaceAmt = String(12, "　") & "元整"
   Else
       strSpaceAmt = ChangeNumber(txt1(11))
   End If
   iStr(29) = "　　　　合計" & IIf(Trim(txt1(29)) = "", "　　　", Trim(txt1(29))) & "　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清。"
   'end 2017/03/28
   iStr(30) = ""
   iStr(31) = ""
   iStr(32) = "　　二、第二條所列之委辦範圍除本條第一款特予載明外，其費用由甲方負擔，其金額依當時外國代理人費用"
   iStr(33) = "　　　　及本所服務費標準收取之。"
   iStr(34) = "　　三、本條所約定之費用如甲方未於所指定之期限內付清，則乙方無義務辦理所受任之事項，且經乙方限期"
   iStr(35) = "　　　　催告後，如甲方仍不履行時，則本契約當然終止，乙方並得通知各該外國代理人終止進行該程序及嗣"
   iStr(36) = "　　　　後之一切程序。另乙方已先行代辦之服務費用，甲方仍應照付。"
   iStr(37) = "第四條　乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，並不得發生足以影響甲方權益之疏"
   iStr(38) = "　　　　誤，否則應對甲方負損害賠償責任。但以不超過第三條所載前酬金金額的三倍為限。"
   iStr(39) = "第五條　甲方確保所交付予乙方之資料，均無虛偽情事，如因不實致生損害或法律責任時，概由甲方負責，與"
   iStr(40) = "　　　　乙方無關。"
   iStr(41) = "第六條　乙方於辦理過程中，應隨時將辦理經過如申請日、案號及其他重要函件，儘速通知或交付甲方。但甲"
   iStr(42) = "　　　　方簽約後變更連絡處所，未即時通知乙方，因而連絡不及致誤時限者，乙方不負責任。"
   iStr(43) = "第七條　凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時限者，乙方不負責任。經乙方通知甲"
   iStr(44) = "　　　　方繳費而未依限繳納者，亦同。"
   iStr(45) = "第八條　甲方如逕自撤回所委辦之程序，或未經乙方同意終止契約時，所約定之費用，仍應全數給付。"
   iStr(46) = "第九條　本約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，但有增刪修改時，需甲乙雙方於更動處蓋"
   iStr(47) = "　　　　章始生效力，並由雙方各執乙份為憑。"
   iStr(48) = "    "
   iStr(49) = "甲　方　　　　　　　　　　　　　　　　 　　　　　乙　方"
   iStr(50) = "委任人：" & StrToStr(txt1(12) & String(48, " "), 24) & "　受任人：" & Combo2.Text 'Add By Sindy 2011/3/23 台一國際專利商標事務所"
   iStr(51) = "代表人：" & StrToStr(txt1(13) & String(48, " "), 24) & "　經手人：" & StrToStr(txt1(17) & String(30, " "), 15)
   'Add By Sindy 2013/12/15
   'Modified by Lydia 2020/04/09 改用模組控制
   'If Combo2 = "台一智權股份有限公司" Then
   '   iStr(52) = "地  址：" & StrToStr(StrConv(MidB(StrConv(txt1(14), vbFromUnicode), 1, 48), vbUnicode) & String(48, " "), 24) & "　地　址：台北市長安東路二段一一０號四樓"
   'Else
   ''2013/12/15 END
   '   iStr(52) = "地  址：" & StrToStr(StrConv(MidB(StrConv(txt1(14), vbFromUnicode), 1, 48), vbUnicode) & String(48, " "), 24) & "　地　址：台北市長安東路二段一一二號九樓"
   'End If
   iStr(52) = "地  址：" & StrToStr(StrConv(MidB(StrConv(txt1(14), vbFromUnicode), 1, 48), vbUnicode) & String(48, " "), 24) & "　地　址：" & PUB_SetAddrTofrm210114(Combo2.Text)
   'end 2020/04/09
   iStr(53) = "　　　　" & StrToStr(StrConv(MidB(StrConv(txt1(14), vbFromUnicode), 49, 48), vbUnicode) & String(48, " "), 24) & "　電  話：(02)25061023(總機)"
   iStr(54) = "電  話：" & StrToStr(txt1(15) & String(22, " "), 11) & "　傳  真：" & StrToStr(txt1(16) & String(18, " "), 9) & "傳  真：(02)25011666、(02)25068147"
   iStr(55) = "  中    華    民    國 " & String((10 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & txt1(18) & String((10 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(19)), vbFromUnicode))) / 2, " ") & txt1(19) & String((10 - LenB(StrConv((txt1(19)), vbFromUnicode))) / 2, " ") & "月" & String((10 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & txt1(20) & String((10 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & "日"
   'Added by Lydia 2020/08/10 英文地址可能1~2行,調整資料
   tmpMaxY = UBound(iStr)
   tmpPosY = 43
   If Trim(iStr(13)) = "SPACEZERO" Then
      For intI = 13 To UBound(iStr)
          If intI < UBound(iStr) Then
             iStr(intI) = iStr(intI + 1)
          Else
             iStr(intI) = ""
          End If
      Next intI
      tmpPosY = tmpPosY - 1
      tmpMaxY = tmpMaxY - 1
   End If
   'end 2020/08/10
   
   'Added by Lydia 2017/03/28 有用印就記錄列印內容
   If iPrintC = 1 And bolAddSeal = True Then
      '有無費用項目
      'strExc(1) = Trim(Replace(Replace(txt1(7) & txt1(8) & txt1(9) & txt1(10) & txt1(25) & txt1(26) & txt1(27) & txt1(28), "　", ""), " ", ""))
      If tmpMaxY = 54 Then 'Added by Lydia 2020/08/10 英文地址只有一行
            For intI = 1 To 54
                If Trim(Replace(Replace(iStr(intI), "　", ""), " ", "")) <> "" Then
                   If intI <= 12 Or (intI >= 23 And intI < 28) Or (intI >= 48 And intI <= 54) Then
                       If intI = 48 Then strDetail = strDetail & vbCrLf
                      strDetail = strDetail & RTrim(iStr(intI)) & vbCrLf
                   ElseIf intI = 28 Then
                      strDetail = strDetail & "　　　　合計" & Trim(txt1(29)) & "　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清。" & vbCrLf
                   End If
                End If
            Next
      'Added by Lydia 2020/08/10 英文地址放成兩行
      Else
            For intI = 1 To tmpMaxY
                If Trim(Replace(Replace(iStr(intI), "　", ""), " ", "")) <> "" Then
                   If intI <= 13 Or (intI >= 24 And intI < 29) Or (intI >= 49 And intI <= 55) Then
                       If intI = 49 Then strDetail = strDetail & vbCrLf
                      strDetail = strDetail & RTrim(iStr(intI)) & vbCrLf
                   ElseIf intI = 29 Then
                      strDetail = strDetail & "　　　　合計" & Trim(txt1(29)) & "　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清。" & vbCrLf
                   End If
                End If
            Next
      End If
      'end 2020/08/10
      'Modified by Lydia 2017/04/17 空白用印改由勾選項目控制
      'If PUB_AddRecSeal("4", txtPCnt.Text, IIf(ChkSeal.Value = 1, "", "Y"), strDetail, Combo2.Text) Then
      'Remove by Lydia 2017/05/16 用印記錄移到pdf建立
      'If PUB_AddRecSeal("4", txtPCnt.Text, IIf(bolSpace = True, "Y", ""), strDetail, Combo2.Text) Then
      'End If
   End If
   'end 2017/03/28
        
   iY = 0
   Printer.PaperSize = 9
   
   'add by nickc 2007/05/04
   'edit by nickc 2007/07/12 試著解決第二頁的格子線會不見的問題
   'If iCount = 1 Then
       Printer.Orientation = 1
   'End If
   Printer.FontName = "標楷體"
   Printer.FontSize = 20
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(iStr(1))) / 2
   iY = iY + Printer.TextHeight(iStr(1))
   Printer.CurrentY = iY
   iY = iY + ((Printer.TextHeight(iStr(1)) / 3) * 4)
   Printer.Print iStr(1)
   Printer.FontSize = 14
   Printer.CurrentX = 1000
   iY = iY + Printer.TextHeight(iStr(2))
   Printer.CurrentY = iY
   iY = iY + ((Printer.TextHeight(iStr(2)) / 3) * 4)
   Printer.Print iStr(2)
   Printer.FontSize = 10
   'Added by Lydia 2017/03/28 同步用印
   If bolAddSeal = True Then
      '列印座置抓乙方資料的起始
      'X軸
      strExc(1) = 1000 + (Printer.TextWidth("　") * 42) - 30
      'Y軸
      'Modified by Lydia 2020/08/10 英文地址放成兩行
      'strExc(2) = iY + ((Printer.TextHeight("　") / 3) * 4) * 42
      strExc(2) = iY + ((Printer.TextHeight("　") / 3) * 4) * tmpPosY
      'Added by Lydia 2017/04/25 圖片尺寸
      strExc(3) = 1600 'width
      strExc(4) = 1600 'height
      
      'Added by Lydia 2020/03/25 已記錄公司名稱|用印編號
      intI = InStr(strCompSeal, Combo2.Text)
      If intI > 0 Then
         strExc(9) = Mid(strCompSeal, intI + Len(Combo2.Text))
         If InStr(strExc(9), ",") > 0 Then
             strExc(9) = Mid(strExc(9), 2, InStr(strExc(9), ",") - 2)
         Else
             strExc(9) = Mid(strExc(9), 2)
         End If
          If PUB_ReadDB2File(strSealFile, Val(strExc(9))) Then
             Set tObj = pvGetStdPicture(strSealFile)
             Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
          End If
      Else
      'end 2020/03/25
            If InStr(Combo2.Text, "專利法律") > 0 Then
              If PUB_ReadDB2File(strSealFile, 51) Then
                 Set tObj = pvGetStdPicture(strSealFile)
                 'Modified by Lydia 2017/04/25
                 'Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), 1570, 1570
                 Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
              End If
            End If
            If InStr(Combo2.Text, "專利商標") > 0 Then
              If PUB_ReadDB2File(strSealFile, 52) Then
                 Set tObj = pvGetStdPicture(strSealFile)
                 'Modified by Lydia 2017/04/25
                 'Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), 1570, 1570
                 Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
              End If
            End If
            If InStr(Combo2.Text, "台一智權") > 0 Then
              If PUB_ReadDB2File(strSealFile, 53) Then
                 Set tObj = pvGetStdPicture(strSealFile)
                 'Modified by Lydia 2017/04/25
                 'Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), 1570, 1570
                 Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
              End If
            End If
      End If 'Added by Lydia 2020/03/25
   End If
   'end 2017/03/28
   
   For tmpI = 3 To UBound(iStr) - 1
       If iStr(tmpI) <> "" Then
           'Modified by Lydia 2020/08/10 英文地址放成兩行
           'If tmpI = 41 Then
           If (tmpMaxY = 54 And tmpI = 41) Or (tmpMaxY = 55 And tmpI = 42) Then
               tBoxTop = iY
           End If
           '畫格子
           'Modified by Lydia 2020/08/10 英文地址放成兩行
           'Select Case tmpI
           'Case 25
           If (tmpMaxY = 54 And tmpI = 25) Or (tmpMaxY = 55 And tmpI = 26) Then
           'end 2020/08/10
               Printer.Line (1000 + (Printer.TextWidth("　") * 4), iY)-(1000 + (Printer.TextWidth("　") * 48), iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
               Printer.Line (1000 + (Printer.TextWidth("　") * 4), iY)-(1000 + (Printer.TextWidth("　") * 48), iY + (((Printer.TextHeight("　") / 3) * 4) * 2)), , B
               Printer.Line (1000 + (Printer.TextWidth("　") * 4), iY)-(1000 + (Printer.TextWidth("　") * 48), iY + (((Printer.TextHeight("　") / 3) * 4) * 1)), , B
               Printer.Line (1000 + (Printer.TextWidth("　") * 4), iY)-(1000 + (Printer.TextWidth("　") * 36), iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
               Printer.Line (1000 + (Printer.TextWidth("　") * 4), iY)-(1000 + (Printer.TextWidth("　") * 26), iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
               Printer.Line (1000 + (Printer.TextWidth("　") * 4), iY)-(1000 + (Printer.TextWidth("　") * 14), iY + (((Printer.TextHeight("　") / 3) * 4) * 3)), , B
           'Modified by Lydia 2020/08/10
           'Case Else
           'End Select
           End If
           'end 2020/08/10

           'Modified by Lydia 2020/08/10 英文地址放成兩行
           'If tmpI = 48 Then
           If (tmpMaxY = 54 And tmpI = 48) Or (tmpMaxY = 55 And tmpI = 49) Then
               Printer.FontSize = 12
           End If
           Printer.CurrentX = 1000
           Printer.CurrentY = iY
           Printer.Print iStr(tmpI)
           'Modified by Lydia 2020/08/10 英文地址放成兩行
           'If tmpI = 48 Then
           If (tmpMaxY = 54 And tmpI = 48) Or (tmpMaxY = 55 And tmpI = 49) Then
               Printer.FontSize = 10
           End If
           
           'Modified by Lydia 2020/08/10 英文地址放成兩行
           'If tmpI = 28 Then
           If (tmpMaxY = 54 And tmpI = 28) Or (tmpMaxY = 55 And tmpI = 29) Then
               Printer.FontBold = True
               Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 10) - 30
               Printer.CurrentY = iY
               'Modified by Lydia 2017/03/28
               'Printer.Print ChangeNumber(txt1(11))
               Printer.Print strSpaceAmt
               Printer.FontBold = False
           End If
           iY = iY + ((Printer.TextHeight(iStr(tmpI)) / 3) * 4)
           '畫線
           'Modified by Lydia 2020/08/10 英文地址放成兩行
           'Select Case tmpI
           'Case 4
           If tmpI = 4 Then
                Printer.Line (1000 + (Printer.TextWidth("　") * 28.5), iY - 50)-(1000 + (Printer.TextWidth("　") * 48), iY - 50)
           'Modified by Lydia 2020/08/10 英文地址放成兩行
           'Case 6
           ElseIf tmpI = 6 Then
                Printer.Line (1000 + (Printer.TextWidth("　") * 9), iY - 50)-(1000 + (Printer.TextWidth("　") * 22), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 28.5), iY - 50)-(1000 + (Printer.TextWidth("　") * 48), iY - 50)
           'Modified by Lydia 2020/08/10 英文地址放成兩行
           'Case 7, 9
           ElseIf tmpI = 7 Or tmpI = 9 Then
                Printer.Line (1000 + (Printer.TextWidth("　") * 9.5), iY - 50)-(1000 + (Printer.TextWidth("　") * 30), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 36.5), iY - 50)-(1000 + (Printer.TextWidth("　") * 48), iY - 50)
           'Modified by Lydia 2020/08/10 英文地址放成兩行
           'Case 10, 12
           ElseIf (tmpMaxY = 54 And InStr("10,12,", Format(tmpI, "00")) > 0) Or (tmpMaxY = 55 And InStr("10,12,13,", Format(tmpI, "00")) > 0) Then
                Printer.Line (1000 + (Printer.TextWidth("　") * 10.5), iY - 50)-(1000 + (Printer.TextWidth("　") * 48), iY - 50)
           'Modified by Lydia 2020/08/10 英文地址放成兩行
           'Case 49, 50, 51, 52
           ElseIf (tmpMaxY = 54 And InStr("49,50,51,52,", Format(tmpI, "00")) > 0) Or (tmpMaxY = 55 And InStr("50,51,52,53,", Format(tmpI, "00")) > 0) Then
                Printer.Line (1000 + (Printer.TextWidth("　") * 4), iY - 50)-(1000 + (Printer.TextWidth("　") * 29), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 33), iY - 50)-(1000 + (Printer.TextWidth("　") * 48), iY - 50)
           'Modified by Lydia 2020/08/10 英文地址放成兩行
           'Case 53
           ElseIf (tmpMaxY = 54 And InStr("53,", Format(tmpI, "00")) > 0) Or (tmpMaxY = 55 And InStr("54,", Format(tmpI, "00")) > 0) Then
                Printer.Line (1000 + (Printer.TextWidth("　") * 4), iY - 50)-(1000 + (Printer.TextWidth("　") * 16), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 20), iY - 50)-(1000 + (Printer.TextWidth("　") * 29), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 33), iY - 50)-(1000 + (Printer.TextWidth("　") * 48), iY - 50)
           'Modified by Lydia 2020/08/10
           'Case Else
           'End Select
           End If
           'end 2020/08/10
       End If
   Next tmpI
   iY = iY + ((Printer.TextHeight(iStr(tmpI)) / 3) * 4)
   Printer.FontSize = 14
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(iStr(UBound(iStr)))) / 2
   Printer.CurrentY = iY
   Printer.Print iStr(UBound(iStr))
   'add by nickc 2007/05/04
   'edit by nickc 2007/07/12 試著解決第二頁的格子線會不見的問題
   'If iCount = Val(txtPCnt) Then
       Printer.EndDoc
   'Else
   '    Printer.NewPage
   'End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

'Modified by Lydia 2022/01/24 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
'Add By Sindy 98/02/11
Dim intLen As Integer
   
   If KeyAscii <> 8 Then
      intLen = GetTextLength(txt1(Index))
      intLen = intLen + GetTextLength(Chr(KeyAscii))
      '2014/5/13 modify by sonia
      'If intLen > txt1(Index).MaxLength Then KeyAscii = 0
      If CheckLengthIsOK(txt1(Index).Text & Chr(KeyAscii), txt1(Index).MaxLength) = False Then
         KeyAscii = 0
      End If
      'end 2014/5/13
   End If
   '98/02/11 End
   If Index = 8 Or Index = 10 Or Index = 11 Or Index = 26 Or Index = 28 Then
       If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 46 Then
           KeyAscii = 0
       End If
   'ElseIf Index = 14 Then
   '    If KeyAscii >= 48 And KeyAscii <= 57 Then
   '        KeyAscii = ChangeZIP(KeyAscii)
   '    End If
   End If
   '2009/11/13 ADD BY SONIA
   If Index = 18 Or Index = 19 Or Index = 20 Then
      If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
      End If
   End If
   '2009/11/13 END
End Sub


Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   'Modified by Lydia 2018/04/13
   'txt1(Index).Text = Replace(Replace(txt1(Index).Text, Chr(10), ""), Chr(13), "")
   txt1(Index).Text = PUB_StringFilter(txt1(Index).Text)
   Cancel = False
   If CheckLengthIsOK(txt1(Index).Text, txt1(Index).MaxLength) = False Then
       txt1(Index).SetFocus
       txt1_GotFocus Index
       Cancel = True
   End If
   'Add By Sindy 2013/12/16
   If strSrvDate(1) >= InvoiceStartDate Then
      If Index = 7 Or Index = 9 Or Index = 25 Or Index = 27 Then 'Add By Sindy 2014/5/6 +if
         If InStr(txt1(7), "大陸") > 0 Or InStr(txt1(9), "大陸") > 0 Or InStr(txt1(25), "大陸") > 0 _
             Or InStr(txt1(27), "大陸") > 0 Then
            'Modify By Sindy 2024/9/25 增加傳入公司別做判斷
            If PUB_ChkCU144isN("", "", txt1(2), IIf(Combo2 = "台一智權股份有限公司", "J", ""), False) = False Then
               Combo2.Text = Combo2.List(2)
            End If
         End If
      End If
   End If
   '2013/12/16 END
End Sub

Private Sub txtPCnt_GotFocus()
   txtPCnt.SelStart = 0
   txtPCnt.SelLength = Len(txtPCnt)
End Sub

Private Sub txtPCnt_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 46 Then
       KeyAscii = 0
   End If
End Sub

'Add By Sindy 2010/4/29
Private Function SetCustTxt(strCUCode As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Add by Lydia 2014/9/22
Dim rsB As New ADODB.Recordset, part1 As Integer, ppart1 As Integer, partCust As String
partCust = strCUCode

   SetCustTxt = False
   strCUCode = Left(strCUCode & "000000000", 9)
   'Modified by Morgan 2021/5/5
   'StrSQLa = "Select * From Customer,nation,potcustcont Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "' and CU10=na01(+) and pcc01(+)=substr(CU08, 1, 8) And pcc02(+)=substr(CU08, 9, 1) "
   StrSQLa = "Select * From Customer Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "'"
   'end 2021/5/5
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      SetCustTxt = True
      '申請人中文
      Me.txt1(2).Text = "" & rsA("CU04").Value
      Me.txt1(12).Text = "" & rsA("CU04").Value
     
      '申請人英文
      Me.txt1(22).Text = "" & rsA("CU05").Value & "" & rsA("CU88").Value & "" & rsA("CU89").Value & "" & rsA("CU90").Value
      
     
     'Add by Lydia 2014/9/22 複選人員
      part1 = 1
      ppart1 = GetSubStringCount(partCust) '取得字串以逗點分隔的Sub字串總數
      Do While part1 < ppart1
        strCUCode = Mid(partCust, (part1 * 10) + 1, 9) '從第幾組代號開始，截取下一組代號
        'Modified by Morgan 2021/5/5
        'StrSQLa = " Select CU04,CU05,CU88,CU89,CU90 From Customer,nation,potcustcont " & _
                  " Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "' and CU10=na01(+) and pcc01(+)=substr(CU08, 1, 8) And pcc02(+)=substr(CU08, 9, 1) "
        StrSQLa = " Select CU04,CU05,CU88,CU89,CU90 From Customer Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "'"
        'end 2021/5/5
        If rsB.State <> adStateClosed Then rsB.Close
        Set rsB = Nothing
        rsB.CursorLocation = adUseClient
        rsB.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsB.RecordCount > 0 Then
         rsB.MoveFirst
         Me.txt1(2).Text = LTrim(RTrim(Me.txt1(2).Text)) + "、" & rsB("CU04").Value
        End If
        If Len(LTrim(RTrim(Me.txt1(22).Text))) > 0 Then
          Me.txt1(22).Text = LTrim(RTrim(Me.txt1(22).Text)) + ", " + "" & rsB("CU05").Value & "" & rsB("CU88").Value & "" & rsB("CU89").Value & "" & rsB("CU90").Value
        Else
          Me.txt1(22).Text = "" & rsB("CU05").Value & "" & rsB("CU88").Value & "" & rsB("CU89").Value & "" & rsB("CU90").Value
        End If
        
        part1 = part1 + 1
        
        If part1 = ppart1 Then
    
           If CheckLengthIsOK(txt1(2).Text, txt1(2).MaxLength) = False Then
                Me.txt1(2).SetFocus
           ElseIf CheckLengthIsOK(txt1(22).Text, txt1(22).MaxLength) = False Then
                Me.txt1(22).SetFocus
           End If
        End If
     Loop
     
'      'ID No.
'      Me.txt1(6).Text = "" & rsA("CU11").Value
      '申請地址
      Me.txt1(4).Text = "" & rsA("CU23").Value
      Me.txt1(14).Text = "" & rsA("CU23").Value
      '申請英文地址
      'Modify by Amy 2014/10/06 +cu102及欄位間加空白
      'Modified by Lydia 2020/08/10 英文地址自動分成兩行,並且彈訊息. (ex. 生展X41336070)
      'Me.txt1(24).Text = Trim("" & rsA("CU24").Value & " " & rsA("CU25").Value & " " & rsA("CU26").Value & " " & rsA("CU27").Value & " " & rsA("CU28").Value & " " & rsA("CU102").Value)
      StrSQLa = Trim("" & rsA("CU24").Value & " " & rsA("CU25").Value & " " & rsA("CU26").Value & " " & rsA("CU27").Value & " " & rsA("CU28").Value & " " & rsA("CU102").Value)
      If GetTextLength(StrSQLa) <= Me.txt1(24).MaxLength Then
          Me.txt1(24).Text = StrSQLa
      Else
          Me.txt1(24).Text = Trim(convForm(StrSQLa, Me.txt1(24).MaxLength))
          StrSQLa = Replace(StrSQLa, Me.txt1(24).Text, "")
          Me.txt1(30).Text = Trim(convForm(StrSQLa, Me.txt1(30).MaxLength))
          StrSQLa = Replace(StrSQLa, Me.txt1(30).Text, "")
          If Trim(StrSQLa) = "" Then
               MsgBox "申請人英文地址超過第一行最大長度，自動將後段地址填入第二行，請自行調整！", vbInformation + vbOKOnly
          Else
               MsgBox "申請人英文地址超過二行最大長度，後段地址無法填入！", vbInformation + vbOKOnly
          End If
      End If
      'end 2020/08/10
      
'      '國籍
'      Me.txt1(8).Text = "" & rsA("NA03").Value
'      '聯絡人地址
'      If "" & rsA("CU08").Value <> "" Then
'         Me.txt1(9).Text = "" & rsA("pcc22").Value
'      Else
'         Me.txt1(9).Text = "" & rsA("CU31").Value
'      End If
      '電話1
      Me.txt1(15).Text = "" & rsA("CU16").Value
      '傳真1
      Me.txt1(16).Text = "" & rsA("CU18").Value
      '代表人1中文
      Me.txt1(3).Text = "" & rsA("CU07").Value
      Me.txt1(13).Text = "" & rsA("CU07").Value
      '代表人1英文
      Me.txt1(23).Text = "" & rsA("CU103").Value    'modify by sonia 2023/10/3 原抓CU07
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Add by Amy 2016/08/19
Private Sub UpdReceiptCmp(ByVal stNowCustNo As String, ByVal stNowCmp As String)
    Dim strUpd As String
        
    'Add by Amy 2016/12/30 +同業務區或為MCTF同組人員才可回寫收據公司別
    If ChkSameCuArea(stNowCustNo, strUserNum) = False Then Exit Sub
    
    'Added by Lydia 2022/08/30 受任人若選擇台一國際智慧財產事務所時，更新客戶檔之相關欄位請更新為NULL，台一智權股份有限公司才更新為J。
    If stNowCmp <> "J" Then
        stNowCmp = ""
    End If
    'end 2022/08/30
    
    'Modified by Lydia 2019/04/12 拿掉UpdateID,Date,Time(CU84,CU85,CU86)
    'strUpd = "Update Customer Set CU84='" & strUserNum & "',CU85=to_number(to_char(sysdate,'YYYYMMDD')),CU86=to_number(to_char(sysdate,'HH24MI')),CU163='" & stNowCmp & "' " & _
                    "Where CU01='" & Left(stNowCustNo, 8) & "' And CU02='" & Mid(stNowCustNo, 9, 1) & "' "
    strUpd = "Update Customer Set CU163='" & stNowCmp & "' " & _
                    "Where CU01='" & Left(stNowCustNo, 8) & "' And CU02='" & Mid(stNowCustNo, 9, 1) & "' "
    Pub_SeekTbLog strUpd
    'Modified by Lydia 2019/04/23 觸發Trigger
    'cnnConnection.Execute strUpd
    cnnConnection.Execute "begin user_data.user_enabled:=1; " & strUpd & " ; end; "
End Sub

'Added by Lydia 2017/04/13 列印:先轉PDF,列印後刪檔
Private Sub Print2PDF(ByVal bSpace As Boolean)
Dim strFileName As String
Dim strOldName As String 'Added by Lydia 2017/06/07

'Added by Lydia 2017/04/25 VB印表機實際列印的左邊界、右邊界
Set Printer = Printers(PUB_PrinterIndex(Combo1.Text))
d_Top = Format((Printer.Height - Printer.ScaleHeight) / 2, "0") '直印
d_Left = Format((Printer.Width - Printer.ScaleWidth) / 2, "0")
'end 2017/04/25

strDetail = "" 'Added by Lydia 2017/05/16
strOldName = App.Title 'Added by Lydia 2017/06/07

Screen.MousePointer = vbHourglass
    'Modified by Lydia 2022/01/24 先產生Word檔，後轉成PDF檔逐一列印
'    For iCount = 1 To Val(txtPCnt)
'        iPrintC = iCount
'        'Modified by Lydia 2017/06/06 改用App.Title變更印表機列印文件名稱(執行exe檔有效,VB跑無效)
'        'strFileName = strUserNum & "_CFT_" & IIf(bSpace = False, IIf(Trim(txt1(2)) <> "", Mid(Trim(txt1(2)), 1, 4), Mid(Trim(txt1(22)), 1, 4)), "空白") & iCount & ".pdf"
'        'If Dir(App.path & "\" & strFileName) <> "" Then
'        '   Kill App.path & "\" & strFileName
'        'End If
'        ''轉PDF
'        'frmPDF.Show
'        'frmPDF.StartProcess App.path, strFileName
'        'Call StrMenu(bSpace)
'        'frmPDF.EndtProcess
'        'Unload frmPDF
'        strFileName = strUserNum & "_CFT_" & IIf(bSpace = False, IIf(Trim(txt1(5)) <> "", Mid(Trim(txt1(5)), 1, 4), Mid(Trim(txt1(6)), 1, 4)), "空白") & iCount
'        App.Title = strFileName
'        Call StrMenu(bSpace)
'        'end 2017/06/07
'
'        'Added by Lydia 2017/05/16 用印記錄移到pdf建立
'        If iCount = 1 And strDetail <> "" Then
'           'If Dir(App.path & "\" & strFileName) <> "" Then 'Remove by Lydia 2020/03/16 因為不存檔案所以取消檔案檢查(自2017/06/08~2020/03/16無用印記錄)
'              If PUB_AddRecSeal("4", txtPCnt.Text, IIf(bSpace = True, "Y", ""), strDetail, Combo2.Text) Then
'              End If
'           'End If 'Remove by Lydia 2020/03/16
'        End If
'        'end 2017/05/16
'
'        'Remove by Lydia 2017/06/07
'        ''列印PDF
'        'PUB_PrintPDF App.path & "\" & strFileName, Me.Combo1
'        ''刪除PDF
'        'Kill App.path & "\" & strFileName
'    Next iCount
    Call runWordProc(bSpace)
    If m_TempPDF <> "" Then
        For iCount = 1 To Val(txtPCnt)
            iPrintC = iCount
            strFileName = strUserNum & "_CFT_" & m_TempFN & iCount
            PUB_PrintPDF App.path & "\" & strUserNum & "\" & m_TempPDF, Combo1.Text
            App.Title = strFileName
        Next iCount
    End If
'--------------先產生Word檔，後轉成PDF檔逐一列印

    App.Title = strOldName 'Added by Lydia 2017/06/07
    
End Sub

'Added by Lydia 2022/01/18 下載Word範本套印
Private Sub runWordProc(ByVal pSpace As Boolean)
Dim iStr(1 To 55) As String  '用印記錄(全文)
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String
Dim strName As String
Dim strText As String
Dim intA As Integer
Dim m_FileName As String, m_TempFileName As String
Dim m_DefPath As String
Dim oShape

On Error GoTo ErrHand
   
   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000300-0-04 智權部委任契約書_CFT.docx", "M51", "000300", "0", "04", "4", "1")
   
   m_DefPath = App.path & "\" & strUserNum
   'Added by Lydia 2022/01/25
   m_TempPDF = ""
   '變更Word印表機
   PUB_SetOsDefaultPrinter Combo1
   PUB_SetWordActivePrinter
   'end 2022/01/25
   
   strDetail = ""
   '下載範本檔: M51-000300-0-04 智權部委任契約書_CFT.docx
   m_TempFN = Pub_RepFileName(IIf(pSpace = False, Mid(Trim(txt1(2)), 1, 4), "空白"))   'Move by Lydia 2022/01/25 從m_TempFileName移過來
   'Modified by Lydia 2022/01/25 改成Word直接印，所以範本一開始就先命名好
   'm_FileName = "$$" & Me.Name & ".docx"
   m_FileName = "$$" & strUserNum & "_CFT_" & m_TempFN & ".docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000300-0-04", , m_DefPath) = False Then
        Exit Sub
   End If
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   'Remove by Lydia 2022/01/25 不用改存PDF檔
'   m_TempFileName = "$$" & strUserNum & "_CFT_" & m_TempFN & ".pdf"
'   If Dir(m_DefPath & "\" & m_TempFileName) <> "" Then
'      Kill m_DefPath & "\" & m_TempFileName
'   End If
   'end 2022/01/25
   
   '改成直接用範本檔 Q: AddToRecentFiles:=False還是會新增到最近開啟記錄
   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
   
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 1 To 32
         strName = "PS" & Format(intA, "000")
         strText = ""
'-------第一條
         If intA = 1 Then
              '案 件 名稱
              strText = PUB_StrToStr(txt1(0), 80)
         ElseIf intA = 2 Then
              '列舉商品(中)
              strText = PUB_StrToStr(txt1(1), 80)
         ElseIf intA = 3 Then
              '列舉商品(英)
              strText = PUB_StrToStr(txt1(21), 80)
         ElseIf intA = 4 Then
              '申請人名稱(中)
              strText = PUB_StrToStr(txt1(2), 76)
         ElseIf intA = 5 Then
              '申請人名稱(英)
              strText = PUB_StrToStr(txt1(22), 76)
         ElseIf intA = 6 Then
              '代　表　人(中)
              strText = PUB_StrToStr(txt1(3), 76)
         ElseIf intA = 7 Then
              '代　表　人(英)
              strText = PUB_StrToStr(txt1(23), 76)
         ElseIf intA = 8 Then
               '申請人地址(中)
              strText = PUB_StrToStr(txt1(4), 76)
         ElseIf intA = 9 Then
               '申請人地址(英)
              strText = PUB_StrToStr(txt1(24), 76)
         ElseIf intA = 10 Then
               '申請人地址(中)
              strText = PUB_StrToStr(txt1(30), 76)
'-------第三條
         ElseIf intA = 11 Then
              '前條第X款: 置中
              'Added by Lydia 2023/02/17 判斷長度不用置中;
              If LenB(StrConv(txt1(5).Text, vbFromUnicode)) >= 6 Then
                  strText = " " & Trim(txt1(5)) & " "
              Else
              'end 2023/02/17
                  strText = String(Int((6 - LenB(StrConv(txt1(5).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(5).Text) & String(6 - LenB(StrConv(String(Int((6 - LenB(StrConv(txt1(5).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(5).Text), vbFromUnicode)), " ")
              End If 'Added by Lydia 2023/02/17
         ElseIf intA = 12 Then
              'XX程序: 置中
              'Added by Lydia 2023/02/17 判斷長度不用置中;
              If LenB(StrConv(txt1(6).Text, vbFromUnicode)) >= 10 Then
                  strText = " " & Trim(txt1(6)) & " "
              Else
              'end 2023/02/17
                  strText = String(Int((10 - LenB(StrConv(txt1(6).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(6).Text) & String(10 - LenB(StrConv(String(Int((10 - LenB(StrConv(txt1(6).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(6).Text), vbFromUnicode)), " ")
              End If 'Added by Lydia 2023/02/17
         ElseIf intA >= 13 And intA <= 16 Then
              '國別、金額
              strText = PUB_StrToStr(txt1(intA - 6), 20)
         ElseIf intA >= 17 And intA <= 20 Then
              '國別、金額
              strText = PUB_StrToStr(txt1(intA + 8), 20)
         ElseIf intA = 21 Then
              '合計
              If Val(Trim(txt1(11))) = 0 Then
                  strExc(1) = String(12, "　")
              Else
                  'Modified by Lydia 2023/08/10 改變數控制
                  'strExc(1) = Replace(ChangeNumber(txt1(11)), "元整", "")
                  strExc(1) = ChangeNumber(txt1(11), False)
              End If
              'Added by Lydia 2022/06/24 判斷超過字元長度不限制
              If GetTextLength(strExc(1)) > 12 Then
                 strText = IIf(Trim(txt1(29).Text) = "", "　　　", Trim(txt1(29).Text)) & "　" & strExc(1) & "　元整"
              Else
              'end 2022/06/24
                 strText = IIf(Trim(txt1(29).Text) = "", "　　　", Trim(txt1(29).Text)) & "　" & PUB_StrToStr(strExc(1), 12, True, True) & "　元整"
              End If 'Added by Lydia 2022/06/24
'------其他
         ElseIf intA = 22 Then
              '委任人
              strText = PUB_StrToStr(txt1(12) & " ", 44)
         ElseIf intA = 23 Then
              '委任人-代表人
              strText = PUB_StrToStr(txt1(13) & " ", 44)
         ElseIf intA = 24 Then
              '委任人-地址
              strText = PUB_StrToStr(txt1(14) & " ", 44)
         ElseIf intA = 25 Then
              '委任人-地址2
              'Modified by Lydia 2022/11/18
              'strExc(1) = PUB_StrToStr(txt1(14) & " ", 44)
              'strText = PUB_StrToStr(Replace(txt1(14), strExc(1), ""), 44)
              If Len(txt1(14)) > 14 Then
                  strExc(1) = PUB_StrToStr(txt1(14) & " ", 44, True)
                  strText = PUB_StrToStr(Replace(txt1(14), strExc(1), ""), 44)
              End If
         ElseIf intA = 26 Then
              '委任人-電話
              strText = PUB_StrToStr(txt1(15) & " ", 20, True)
         ElseIf intA = 27 Then
              '委任人-傳真
              strText = PUB_StrToStr(txt1(16) & " ", 20, True)
         ElseIf intA = 28 Then
              '受任人
              strText = Combo2.Text
         ElseIf intA = 29 Then
              '經手人
              strText = PUB_StrToStr(txt1(17) & " ", 30)
         ElseIf intA = 30 Then
              '受任人-地址
              strText = PUB_SetAddrTofrm210114(Combo2.Text)
         ElseIf intA = 31 Then
              strText = "        中    華    民    國 " & String((6 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & txt1(18) & String((6 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & "年" & String((6 - LenB(StrConv((txt1(19)), vbFromUnicode))) / 2, " ") & txt1(19) & String((6 - LenB(StrConv((txt1(19)), vbFromUnicode))) / 2, " ") & "月" & String((6 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & txt1(20) & String((6 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & "日"
         ElseIf intA = 32 Then
              strText = ""
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
            .Selection.Find.Replacement.Text = ""
            .Selection.Find.Forward = True
            .Selection.Find.Wrap = wdFindContinue
            .Selection.Find.Format = False
            .Selection.Find.MatchCase = False
            .Selection.Find.MatchWholeWord = False
            .Selection.Find.MatchWildcards = False
            .Selection.Find.MatchSoundsLike = False
            .Selection.Find.MatchAllWordForms = False
            .Selection.Find.MatchByte = True
            .Selection.Find.Execute
            .Selection.Delete
            '保留;因為先全部以細明體-ExtB,最後全選改字型;
            If (intA >= 1 And intA <= 10) Or (intA >= 22 And intA <= 25) Or intA = 28 Then
               '有Unicode字需要換字型
               .Selection.Font.Name = "細明體-ExtB"
            End If
            If intA = 21 Then
                '合計要粗體
                .Selection.Font.Bold = True
            End If
            If intA = 26 Or intA = 27 Then
                '委任人-電話、傳真加底線
                .Selection.Font.Underline = True
            End If
            
            If intA = 32 And bolAddSeal = True Then  '公司章: 放在甲方的儲存格
                strExc(9) = Mid(strCompSeal, InStr(strCompSeal, Combo2))
                If InStr(strExc(9), ",") > 0 Then
                    strExc(9) = Right(Mid(strExc(9), 1, InStr(strExc(9), ",") - 1), 2)
                Else
                    strExc(9) = Right(strExc(9), 2)
                End If
                If PUB_ReadDB2File(m_DefPath & "\$$" & Me.Name & "TempFile", Val(strExc(9))) Then
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=m_DefPath & "\$$" & Me.Name & "TempFile", LinkToFile:=False, SaveWithDocument:=True)
                    '--------設定圖片=文蓋圖(文字在前)
                        oShape.Fill.Visible = msoFalse
                        oShape.Fill.Solid
                        oShape.Fill.Transparency = 0#
                        oShape.Line.Weight = 0.75
                        oShape.Line.DashStyle = msoLineSolid
                        oShape.Line.Style = msoLineSingle
                        oShape.Line.Transparency = 0#
                        oShape.Line.Visible = msoFalse
                        oShape.LockAspectRatio = msoTrue
                        oShape.Rotation = 0#
                        oShape.PictureFormat.Brightness = 0.5
                        oShape.PictureFormat.Contrast = 0.5
                        oShape.PictureFormat.ColorType = msoPictureAutomatic
                        oShape.PictureFormat.CropLeft = 0#
                        oShape.PictureFormat.CropRight = 0#
                        oShape.PictureFormat.CropTop = 0#
                        oShape.PictureFormat.CropBottom = 0#
                        oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                        oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                        oShape.Left = .CentimetersToPoints(15.5)
                        oShape.Top = .CentimetersToPoints(-0.2)
                        oShape.LockAnchor = False
                        oShape.LayoutInCell = True
                        oShape.WrapFormat.AllowOverlap = True
                        oShape.WrapFormat.Side = wdWrapBoth
                        oShape.WrapFormat.DistanceTop = .CentimetersToPoints(0)
                        oShape.WrapFormat.DistanceBottom = .CentimetersToPoints(0)
                        oShape.WrapFormat.DistanceLeft = .CentimetersToPoints(0.32)
                        oShape.WrapFormat.DistanceRight = .CentimetersToPoints(0.32)
                        oShape.WrapFormat.Type = 3
                        oShape.ZOrder 5 '文蓋圖(文字在前)
                        '---------------------------
                End If
          
            End If
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
            '保留;因為先全部以細明體-ExtB,最後全選改字型;
            If (intA >= 1 And intA <= 10) Or (intA >= 22 And intA <= 25) Or intA = 28 Then
               '有Unicode字需要換字型=>還原
               .Selection.Font.Name = "標楷體"
            End If
            If intA = 21 Then
                '合計要粗體=>還原
                .Selection.Font.Bold = False
            End If
            If intA = 26 Or intA = 27 Then
                '委任人-電話、傳真加底線=>還原
                .Selection.Font.Underline = False
            End If
         End If

      Next intA
'      '因為先全部以細明體-ExtB,最後全選改字型;
      .Selection.WholeStory
      .Selection.Font.Name = "標楷體"
   End With
   
   
   '改存成PDF檔
   'Memo by Lydia 2022/01/25  因為受PDF redirect設定灰階列印影響，改成Word直接印
   intA = IIf(Val(txtPCnt) = 0, 1, Val(txtPCnt))
   For intI = 1 To intA
       g_WordAp.PrintOut Background:=False, Range:=4, Item:=0, Copies:=1, Pages:="1", Collate:=True
   Next intI
   
   '保留: 存檔
   'g_WordAp.ActiveDocument.Close wdSaveChanges
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing '避免快速開啟Word,程式出錯
   m_TempPDF = m_FileName 'Added by Lydia 2022/01/25
   
   'Mark by Lydia 2022/01/25 因為受PDF redirect設定灰階列印影響，改成Word直接印
   'If PUB_PrintWord2PDF(g_WordAp, m_DefPath, m_TempFileName, m_TempPDF) = False Then
   '    Exit Sub
   'End If
   'end 2022/01/19
   
If bolAddSeal = True Then  '用印記錄
   strDetail = ""
   iStr(1) = "國外商標案件委任契約書"
   iStr(2) = "委任人(甲方)茲委任受任人(乙方)辦理國外商標案件，雙方同意條件如下："
   iStr(3) = "第一條" '??■□
   iStr(4) = "    　案 件 名稱：" & PUB_StrToStr(txt1(0), 80)
   iStr(5) = "　　列舉商品(中)：" & PUB_StrToStr(txt1(1), 80)
   iStr(6) = "　　　　　　(英)：" & PUB_StrToStr(txt1(21), 80)
   iStr(7) = "　　申請人名稱(中)：" & PUB_StrToStr(txt1(2), 76)
   iStr(8) = "　　　　　　　(英)：" & PUB_StrToStr(txt1(22), 76)
   iStr(9) = "　　代　表　人(中)：" & PUB_StrToStr(txt1(3), 76)
   iStr(10) = "　　　　　　　(英)：" & PUB_StrToStr(txt1(23), 76)
   iStr(11) = "　　申請人地址(中)：" & PUB_StrToStr(txt1(4), 76)
   iStr(12) = "　　　　　　　(英)：" & PUB_StrToStr(txt1(24), 76)
   iStr(13) = "　　　　　　　　　　" & PUB_StrToStr(txt1(30), 76)
   iStr(14) = "第二條　委辦範圍："
   iStr(15) = "　　1.查名程序：乙方根據甲方所提供之商標圖樣，委請各該查名國商標代理人代為查尋前案資料。"
   iStr(16) = "　　2.申請程序：乙方根據甲方所提供之資料代撰必要書件及製作印版、圖樣並為甲方委請各該申請國之"
   iStr(17) = "　　　　　　　　商標代理人，代向各該國提出商標申請。"
   iStr(18) = "　　3.中間程序：提出商標申請後，依甲方請求或各該申請國商標主管機關之指示所需提出之修正、補正、"
   iStr(19) = "　　　　　　　　更正等各程序。"
   iStr(20) = "　　4.領證程序：核准後之領取證書程序。"
   iStr(21) = "　　5.救濟程序：審定不予核准時所需進行之救濟程序。"
   iStr(22) = "　　6.其　　他：如討論案情、異議、撤銷、答辯、授權、移轉、陳報、延展、變更、參加聽證會、外國"
   iStr(23) = "　　　　　　　　商標主管機關文件及引證資料之翻譯等程序。"
   iStr(24) = "第三條　委辦費用"
   'Modified by Lydia 2023/02/17 直接代入
   'iStr(25) = "　　一、乙方受委辦前條第" & String(Int((6 - LenB(StrConv(txt1(5).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(5).Text) & String(6 - LenB(StrConv(String(Int((6 - LenB(StrConv(txt1(5).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(5).Text), vbFromUnicode)), " ") & "款" & String(Int((10 - LenB(StrConv(txt1(6).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(6).Text) & String(10 - LenB(StrConv(String(Int((10 - LenB(StrConv(txt1(6).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(6).Text), vbFromUnicode)), " ") & "程序之費用（包括國外代理人費用），約定如下："
   iStr(25) = "　　一、乙方受委辦前條第 " & Trim(txt1(5).Text) & " 款 " & Trim(txt1(6).Text) & " 程序之費用（包括國外代理人費用），約定如下："
   iStr(26) = "　　　　　　　國　　別　　　　　　金額　　　　　　　　　　國　　別　　　　　金額　　　　　"
   iStr(27) = "　　　　" & String(20 - LenB(StrConv(String(Int((20 - LenB(StrConv(txt1(7).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(7).Text), vbFromUnicode)), " ") & Trim(txt1(7).Text) & String(Int((20 - LenB(StrConv(txt1(7).Text, vbFromUnicode))) / 2), " ") & String(24 - LenB(StrConv(String(Int((24 - LenB(StrConv(txt1(8).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(8).Text), vbFromUnicode)), " ") & Trim(txt1(8).Text) & String(Int((24 - LenB(StrConv(txt1(8).Text, vbFromUnicode))) / 2), " ") & String(Int((20 - LenB(StrConv(txt1(9).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(9).Text) & String(20 - LenB(StrConv(String(Int((20 - LenB(StrConv(txt1(9).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(9).Text), vbFromUnicode)), " ") & String(24 - LenB(StrConv(String(Int((24 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(10).Text), vbFromUnicode)), " ") & Trim(txt1(10).Text) & String(Int((24 - LenB(StrConv(txt1(10).Text, vbFromUnicode))) / 2), " ")
   iStr(28) = "　　　　" & String(20 - LenB(StrConv(String(Int((20 - LenB(StrConv(txt1(25).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(25).Text), vbFromUnicode)), " ") & Trim(txt1(25).Text) & String(Int((20 - LenB(StrConv(txt1(25).Text, vbFromUnicode))) / 2), " ") & String(24 - LenB(StrConv(String(Int((24 - LenB(StrConv(txt1(26).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(26).Text), vbFromUnicode)), " ") & Trim(txt1(26).Text) & String(Int((24 - LenB(StrConv(txt1(26).Text, vbFromUnicode))) / 2), " ") & String(Int((20 - LenB(StrConv(txt1(27).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(27).Text) & String(20 - LenB(StrConv(String(Int((20 - LenB(StrConv(txt1(27).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(27).Text), vbFromUnicode)), " ") & String(24 - LenB(StrConv(String(Int((24 - LenB(StrConv(txt1(28).Text, vbFromUnicode))) / 2), " ") & Trim(txt1(28).Text), vbFromUnicode)), " ") & Trim(txt1(28).Text) & String(Int((24 - LenB(StrConv(txt1(28).Text, vbFromUnicode))) / 2), " ")
   If Val(Trim(txt1(11))) = 0 Then
       strSpaceAmt = String(12, "　") & "元整"
   Else
       strSpaceAmt = ChangeNumber(txt1(11))
   End If
   iStr(29) = "　　　　合計" & IIf(Trim(txt1(29)) = "", "　　　", Trim(txt1(29))) & "　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清。"
   iStr(30) = ""
   iStr(31) = ""
   iStr(32) = "　　二、第二條所列之委辦範圍除本條第一款特予載明外，其費用由甲方負擔，其金額依當時外國代理人費用"
   iStr(33) = "　　　　及本所服務費標準收取之。"
   iStr(34) = "　　三、本條所約定之費用如甲方未於所指定之期限內付清，則乙方無義務辦理所受任之事項，且經乙方限期"
   iStr(35) = "　　　　催告後，如甲方仍不履行時，則本契約當然終止，乙方並得通知各該外國代理人終止進行該程序及嗣"
   iStr(36) = "　　　　後之一切程序。另乙方已先行代辦之服務費用，甲方仍應照付。"
   iStr(37) = "第四條　乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，並不得發生足以影響甲方權益之疏"
   iStr(38) = "　　　　誤，否則應對甲方負損害賠償責任。但以不超過第三條所載前酬金金額的三倍為限。"
   iStr(39) = "第五條　甲方確保所交付予乙方之資料，均無虛偽情事，如因不實致生損害或法律責任時，概由甲方負責，與"
   iStr(40) = "　　　　乙方無關。"
   iStr(41) = "第六條　乙方於辦理過程中，應隨時將辦理經過如申請日、案號及其他重要函件，儘速通知或交付甲方。但甲"
   iStr(42) = "　　　　方簽約後變更連絡處所，未即時通知乙方，因而連絡不及致誤時限者，乙方不負責任。"
   iStr(43) = "第七條　凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時限者，乙方不負責任。經乙方通知甲"
   iStr(44) = "　　　　方繳費而未依限繳納者，亦同。"
   iStr(45) = "第八條　甲方如逕自撤回所委辦之程序，或未經乙方同意終止契約時，所約定之費用，仍應全數給付。"
   iStr(46) = "第九條　本約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，但有增刪修改時，需甲乙雙方於更動處蓋"
   iStr(47) = "　　　　章始生效力，並由雙方各執乙份為憑。"
   iStr(48) = "    "
   iStr(49) = "甲　方　　　　　　　　　　　　　　　　 　　　　　　　乙　方"
   iStr(50) = "委任人：" & PUB_StrToStr(txt1(12) & " ", 44, True, True) & "　受任人：" & Combo2.Text
   iStr(51) = "代表人：" & PUB_StrToStr(txt1(13) & " ", 44, True, True) & "　經手人：" & PUB_StrToStr(txt1(17) & " ", 30, True, True)
   strExc(1) = PUB_StrToStr(txt1(14) & " ", 44)
   strExc(2) = Replace(PUB_StrToStr(txt1(14) & " ", 88), strExc(1), "")
   iStr(52) = "地  址：" & strExc(1) & "　地　址：" & PUB_SetAddrTofrm210114(Combo2.Text)
   iStr(53) = "　　　　" & PUB_StrToStr(strExc(2) & " ", 44) & "　電  話：(02)25061023(總機)"
   iStr(54) = "電  話：" & PUB_StrToStr(txt1(15) & " ", 44, True, True) & "　傳  真：" & PUB_StrToStr(txt1(16) & " ", 18) & "傳  真：(02)25011666、(02)25068147"
   iStr(55) = "  中    華    民    國 " & String((10 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & txt1(18) & String((10 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(19)), vbFromUnicode))) / 2, " ") & txt1(19) & String((10 - LenB(StrConv((txt1(19)), vbFromUnicode))) / 2, " ") & "月" & String((10 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & txt1(20) & String((10 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & "日"
'
'   'Added by Lydia 2017/03/28 有用印就記錄列印內容
'   If iPrintC = 1 And bolAddSeal = True Then
'      '有無費用項目
'      If tmpMaxY = 54 Then 'Added by Lydia 2020/08/10 英文地址只有一行
'            For intI = 1 To 54
'                If Trim(Replace(Replace(iStr(intI), "　", ""), " ", "")) <> "" Then
'                   If intI <= 12 Or (intI >= 23 And intI < 28) Or (intI >= 48 And intI <= 54) Then
'                       If intI = 48 Then strDetail = strDetail & vbCrLf
'                      strDetail = strDetail & RTrim(iStr(intI)) & vbCrLf
'                   ElseIf intI = 28 Then
'                      strDetail = strDetail & "　　　　合計" & Trim(txt1(29)) & "　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清。" & vbCrLf
'                   End If
'                End If
'            Next
'      'Added by Lydia 2020/08/10 英文地址放成兩行
'      Else
            For intI = 1 To UBound(iStr)
                If Trim(Replace(Replace(iStr(intI), "　", ""), " ", "")) <> "" Then
                   If intI <= 13 Or (intI >= 24 And intI < 29) Or (intI >= 49 And intI <= 55) Then
                       If intI = 49 Then strDetail = strDetail & vbCrLf
                      strDetail = strDetail & RTrim(iStr(intI)) & vbCrLf
                   ElseIf intI = 29 Then
                      strDetail = strDetail & "　　　　合計" & Trim(txt1(29)) & "　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清。" & vbCrLf
                   End If
                End If
            Next
'      End If
      'end 2020/08/10
      If PUB_AddRecSeal("4", txtPCnt.Text, IIf(pSpace = True, "Y", ""), strDetail, Combo2.Text) Then
      End If
End If
          
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Number & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub

'Added by Lydia 2022/01/20 刪除暫存檔
Private Sub RunEndProc(ByVal bolSleep As Boolean)
   
   If bolSleep = True Then Sleep 3000
   PUB_KillTempFile (strUserNum & "\$$" & strUserNum & "*_CFT*.*")
   PUB_KillTempFile (strUserNum & "\$$" & Me.Name & "*.*")
    
End Sub

