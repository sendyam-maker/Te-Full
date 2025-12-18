VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210114_7 
   BorderStyle     =   1  '單線固定
   Caption         =   "著作權案件委任契約書"
   ClientHeight    =   7704
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9684
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7704
   ScaleWidth      =   9684
   Begin VB.OptionButton opt1 
      Caption         =   "不會稿"
      Height          =   210
      Index           =   1
      Left            =   5925
      TabIndex        =   45
      Top             =   5055
      Width           =   1080
   End
   Begin VB.OptionButton opt1 
      Caption         =   "會稿"
      Height          =   210
      Index           =   0
      Left            =   4785
      TabIndex        =   44
      Top             =   5055
      Width           =   1080
   End
   Begin VB.CheckBox chkItem 
      Caption         =   "專屬授權或處分之限制登記"
      Height          =   255
      Index           =   6
      Left            =   1200
      TabIndex        =   41
      Top             =   4650
      Width           =   2955
   End
   Begin VB.CheckBox chkItem 
      Caption         =   "其他"
      Height          =   255
      Index           =   5
      Left            =   4650
      TabIndex        =   39
      Top             =   4395
      Width           =   765
   End
   Begin VB.CheckBox chkItem 
      Caption         =   "首次公開發表日或發行日登記"
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   38
      Top             =   4395
      Width           =   2955
   End
   Begin VB.CheckBox chkItem 
      Caption         =   "變更登記"
      Height          =   255
      Index           =   3
      Left            =   4650
      TabIndex        =   37
      Top             =   4155
      Width           =   1305
   End
   Begin VB.CheckBox chkItem 
      Caption         =   "著作財產權人登記"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   36
      Top             =   4155
      Width           =   1965
   End
   Begin VB.CheckBox chkItem 
      Caption         =   "讓與登記"
      Height          =   255
      Index           =   1
      Left            =   4650
      TabIndex        =   35
      Top             =   3900
      Width           =   1305
   End
   Begin VB.CheckBox chkItem 
      Caption         =   "著作人登記"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   34
      Top             =   3900
      Width           =   1305
   End
   Begin VB.CheckBox ChkKind 
      Caption         =   "其他"
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   33
      Top             =   3240
      Width           =   765
   End
   Begin VB.CheckBox ChkKind 
      Caption         =   "戲劇、舞蹈著作"
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   32
      Top             =   3270
      Width           =   1575
   End
   Begin VB.CheckBox ChkKind 
      Caption         =   "電腦程式著作"
      Height          =   255
      Index           =   8
      Left            =   1230
      TabIndex        =   31
      Top             =   3270
      Width           =   1515
   End
   Begin VB.CheckBox ChkKind 
      Caption         =   "建築著作"
      Height          =   255
      Index           =   7
      Left            =   6600
      TabIndex        =   30
      Top             =   2970
      Width           =   1065
   End
   Begin VB.CheckBox ChkKind 
      Caption         =   "錄音著作"
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   29
      Top             =   2970
      Width           =   1065
   End
   Begin VB.CheckBox ChkKind 
      Caption         =   "視聽著作"
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   28
      Top             =   3000
      Width           =   1065
   End
   Begin VB.CheckBox ChkKind 
      Caption         =   "圖形著作"
      Height          =   255
      Index           =   4
      Left            =   1230
      TabIndex        =   27
      Top             =   3000
      Width           =   1065
   End
   Begin VB.CheckBox ChkKind 
      Caption         =   "攝影著作"
      Height          =   255
      Index           =   3
      Left            =   6600
      TabIndex        =   26
      Top             =   2700
      Width           =   1065
   End
   Begin VB.CheckBox ChkKind 
      Caption         =   "美術著作"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   25
      Top             =   2700
      Width           =   1065
   End
   Begin VB.CheckBox ChkKind 
      Caption         =   "音樂著作"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   24
      Top             =   2730
      Width           =   1065
   End
   Begin VB.CheckBox ChkKind 
      Caption         =   "語文著作"
      Height          =   255
      Index           =   0
      Left            =   1230
      TabIndex        =   23
      Top             =   2730
      Width           =   1065
   End
   Begin VB.CheckBox ChkSeal 
      Caption         =   "用印"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   4890
      TabIndex        =   57
      Top             =   7380
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0FFFF&
      Caption         =   "空白列印"
      Height          =   330
      Index           =   5
      Left            =   3720
      Style           =   1  '圖片外觀
      TabIndex        =   77
      Top             =   30
      Width           =   920
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm210114_7.frx":0000
      Left            =   6510
      List            =   "frm210114_7.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   58
      Top             =   7350
      Width           =   2475
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋委託人(&Q)"
      Height          =   330
      Left            =   8250
      TabIndex        =   47
      Top             =   5640
      Width           =   1300
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "讀取文檔"
      Height          =   330
      Index           =   4
      Left            =   5544
      TabIndex        =   61
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "儲存文檔"
      Height          =   330
      Index           =   3
      Left            =   4632
      TabIndex        =   60
      Top             =   30
      Width           =   920
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   360
      Left            =   0
      TabIndex        =   66
      Top             =   30
      Width           =   3645
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   660
         Style           =   2  '單純下拉式
         TabIndex        =   59
         Top             =   30
         Width           =   2940
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "印表機"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   67
         Top             =   90
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "清空資料"
      Height          =   330
      Index           =   2
      Left            =   6456
      TabIndex        =   62
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面"
      Height          =   330
      Index           =   1
      Left            =   8460
      TabIndex        =   65
      Top             =   30
      Width           =   920
   End
   Begin VB.TextBox txtPCnt 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Left            =   7890
      MaxLength       =   1
      TabIndex        =   63
      Text            =   "2"
      Top             =   60
      Width           =   270
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9420
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印       份"
      Height          =   330
      Index           =   0
      Left            =   7368
      TabIndex        =   64
      Top             =   30
      Width           =   1100
   End
   Begin VB.Label lblProperty 
      Caption         =   "申請登記"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1260
      TabIndex        =   104
      Top             =   3600
      Width           =   810
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   23
      Left            =   5460
      TabIndex        =   103
      Top             =   3240
      Width           =   3585
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "6324;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   3
      Left            =   -60
      TabIndex        =   102
      Top             =   7020
      Width           =   180
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   -30
      TabIndex        =   101
      Top             =   5700
      Width           =   180
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   100
      Top             =   510
      Width           =   180
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   24
      Left            =   5430
      TabIndex        =   40
      Top             =   4410
      Width           =   3585
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "6324;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "地　址："
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   99
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊為必填欄位"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   8280
      TabIndex        =   98
      Top             =   6060
      Width           =   1080
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   34
      Left            =   1860
      TabIndex        =   54
      Top             =   7350
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   36
      Left            =   3720
      TabIndex        =   56
      Top             =   7350
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   35
      Left            =   2760
      TabIndex        =   55
      Top             =   7350
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   33
      Left            =   1380
      TabIndex        =   53
      Top             =   6960
      Width           =   6795
      VariousPropertyBits=   671105051
      MaxLength       =   52
      Size            =   "11994;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   32
      Left            =   5370
      TabIndex        =   52
      Top             =   6630
      Width           =   2805
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4948;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "傳　真："
      Height          =   180
      Index           =   8
      Left            =   4620
      TabIndex        =   97
      Top             =   6690
      Width           =   720
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   31
      Left            =   1380
      TabIndex        =   51
      Top             =   6630
      Width           =   2805
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4948;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   30
      Left            =   1380
      TabIndex        =   50
      Top             =   6300
      Width           =   6795
      VariousPropertyBits=   671105051
      MaxLength       =   52
      Size            =   "11994;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   29
      Left            =   5370
      TabIndex        =   49
      Top             =   5970
      Width           =   2805
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4948;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ID NO："
      Height          =   180
      Index           =   7
      Left            =   630
      TabIndex        =   96
      Top             =   6030
      Width           =   645
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   28
      Left            =   1380
      TabIndex        =   48
      Top             =   5970
      Width           =   2805
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4948;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   27
      Left            =   1380
      TabIndex        =   46
      Top             =   5640
      Width           =   6795
      VariousPropertyBits=   671105051
      MaxLength       =   52
      Size            =   "11994;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   26
      Left            =   1230
      TabIndex        =   43
      Top             =   5310
      Width           =   1245
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "2196;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   25
      Left            =   1230
      TabIndex        =   42
      Top             =   5010
      Width           =   1245
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "2196;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "前   酬   金："
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   95
      Top             =   5070
      Width           =   990
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "後   酬   金："
      Height          =   180
      Left            =   240
      TabIndex        =   94
      Top             =   5370
      Width           =   990
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "元整"
      Height          =   180
      Left            =   2550
      TabIndex        =   93
      Top             =   5070
      Width           =   360
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "元整"
      Height          =   180
      Left            =   2550
      TabIndex        =   92
      Top             =   5370
      Width           =   360
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "請輸入數字"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2970
      TabIndex        =   91
      Top             =   5070
      Width           =   900
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "請輸入數字"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2970
      TabIndex        =   90
      Top             =   5370
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "登記事項："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   6
      Left            =   240
      TabIndex        =   89
      Top             =   3900
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "性　　質："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   88
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "著作種類："
      Height          =   180
      Index           =   4
      Left            =   240
      TabIndex        =   87
      Top             =   2970
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "委辦範圍："
      Height          =   180
      Index           =   3
      Left            =   45
      TabIndex        =   86
      Top             =   2760
      Width           =   900
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   22
      Left            =   1005
      TabIndex        =   22
      Top             =   2340
      Width           =   7995
      VariousPropertyBits=   671105051
      MaxLength       =   62
      Size            =   "14102;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   20
      Left            =   7680
      TabIndex        =   20
      Top             =   2040
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   21
      Left            =   8640
      TabIndex        =   21
      Top             =   2040
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   19
      Left            =   6780
      TabIndex        =   19
      Top             =   2040
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "著作財產權人："
      Height          =   180
      Index           =   1
      Left            =   45
      TabIndex        =   84
      Top             =   2100
      Width           =   1260
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   18
      Left            =   1410
      TabIndex        =   18
      Top             =   2040
      Width           =   4125
      VariousPropertyBits=   671105051
      MaxLength       =   24
      Size            =   "7276;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "地　址："
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   83
      Top             =   1770
      Width           =   720
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   17
      Left            =   1005
      TabIndex        =   17
      Top             =   1710
      Width           =   7995
      VariousPropertyBits=   671105051
      MaxLength       =   62
      Size            =   "14102;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   15
      Left            =   7680
      TabIndex        =   15
      Top             =   1410
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   16
      Left            =   8640
      TabIndex        =   16
      Top             =   1410
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   14
      Left            =   6780
      TabIndex        =   14
      Top             =   1410
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "著作人："
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   81
      Top             =   1470
      Width           =   720
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   13
      Left            =   1005
      TabIndex        =   13
      Top             =   1410
      Width           =   4125
      VariousPropertyBits=   671105051
      MaxLength       =   32
      Size            =   "7276;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   11
      Left            =   7680
      TabIndex        =   11
      Top             =   1050
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   12
      Left            =   8640
      TabIndex        =   12
      Top             =   1050
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   6780
      TabIndex        =   10
      Top             =   1050
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   2550
      TabIndex        =   8
      Top             =   1050
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   3510
      TabIndex        =   9
      Top             =   1050
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1590
      TabIndex        =   7
      Top             =   1050
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   7680
      TabIndex        =   5
      Top             =   750
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   8640
      TabIndex        =   6
      Top             =   750
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   6780
      TabIndex        =   4
      Top             =   750
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "公開發表日：　　　　年　　　 　月　　　　日"
      Height          =   180
      Index           =   0
      Left            =   5670
      TabIndex        =   78
      Top             =   810
      Width           =   3825
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1110
      TabIndex        =   0
      Top             =   420
      Width           =   7995
      VariousPropertyBits=   671105051
      MaxLength       =   62
      Size            =   "14102;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1590
      TabIndex        =   1
      Top             =   750
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   3510
      TabIndex        =   3
      Top             =   750
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   2550
      TabIndex        =   2
      Top             =   750
      Width           =   600
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1058;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "受任人："
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5760
      TabIndex        =   76
      Top             =   7410
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "著作名稱："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   75
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "甲方：委任人："
      Height          =   180
      Left            =   150
      TabIndex        =   74
      Top             =   5700
      Width           =   1260
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代表人："
      Height          =   180
      Left            =   4620
      TabIndex        =   73
      Top             =   6030
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "地　址："
      Height          =   180
      Left            =   630
      TabIndex        =   72
      Top             =   6330
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "電　話："
      Height          =   180
      Left            =   630
      TabIndex        =   71
      Top             =   6630
      Width           =   720
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "乙方：經手人："
      Height          =   180
      Left            =   105
      TabIndex        =   70
      Top             =   7020
      Width           =   1260
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "中　華　民　國　　　　　年　　　　月　　　　　日"
      Height          =   180
      Left            =   330
      TabIndex        =   69
      Top             =   7410
      Width           =   4320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "著作完成日：　　　　年　　　 　月　　　　日"
      Height          =   180
      Left            =   480
      TabIndex        =   68
      Top             =   810
      Width           =   3825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "發　行　日：　　　　年　　　 　月　　　　日"
      Height          =   180
      Left            =   480
      TabIndex        =   79
      Top             =   1110
      Width           =   3825
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "著作權轉讓日：　　　　年　　　 　月　　　　日"
      Height          =   180
      Left            =   5490
      TabIndex        =   80
      Top             =   1110
      Width           =   4005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "出生/設立日：　　　　年　　　 　月　　　　日"
      Height          =   180
      Index           =   1
      Left            =   5640
      TabIndex        =   82
      Top             =   1470
      Width           =   3870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "出生/設立日：　　　　年　　　 　月　　　　日"
      Height          =   180
      Index           =   2
      Left            =   5640
      TabIndex        =   85
      Top             =   2130
      Width           =   3870
   End
End
Attribute VB_Name = "frm210114_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2022/04/15 著作權案件委任契約書
Option Explicit

Dim iCount As Integer
Public m_strCustCode As String
Public m_blnOneRec As Boolean
Dim strNowCustNo As String '客戶編號
Dim bolAddSeal As Boolean '是否用印
Dim strPrinter As String
Dim strDetail As String '記錄內容
Dim strCompSeal As String '記錄"公司名稱|用印編號",用,區隔
'加入圖片用(Word)
Const msoFalse = 0
Const msoLineSolid = 1
Const msoLineSingle = 1
Const msoTrue = -1
Const msoPictureAutomatic = 1
Dim m_TempPDF As String
Dim oControl As Control

Private Sub chkItem_Click(Index As Integer)
   If Index = 5 Then
       If chkItem(Index).Value = vbChecked Then
           txt1(24).SetFocus
       End If
   End If
End Sub

Private Sub ChkKind_Click(Index As Integer)
   If Index = 10 Then
       If ChkKind(Index).Value = vbChecked Then
           txt1(23).SetFocus
       End If
   End If
End Sub

Private Sub cmdFind_Click()
Dim strCmpName As String, strMsg As String

   If Me.txt1(27).Text = "" Then
      MsgBox "請輸入申請人中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txt1(27).SetFocus
      Exit Sub
   End If
   
   Set frm090801_1.m_frm0908A = Me
   frm090801_1.m_DouChk = False
   
   frm090801_1.m_strCustChnName = Me.txt1(27).Text
   frm090801_1.lblName.Caption = Me.txt1(27).Text
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
   Combo2.Tag = "": strNowCustNo = ""
   If m_blnOneRec = True And m_strCustCode <> "" Then
     '記錄收據公司別(放於SetCustTxt前避免m_strCustCode被清空)
      strNowCustNo = m_strCustCode
      strCmpName = "Y"
      Combo2.Tag = GetReceiptCmp(Left(strNowCustNo, 8), Mid(strNowCustNo, 9, 1), "TC", "000", False, strCmpName, Me.Name)
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
      '讀取文檔要保留原文檔內容
      If Me.ActiveControl.Name = "cmdFind" Then
        Call SetCustTxt(m_strCustCode)
      End If
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim tb As Control
Dim fN As Integer
Dim strBuffer As String
Dim AllObj(0 To 56) As String
Dim AllObjV As Variant
Dim strNowCmp As String '目前收據公司別
Dim iErr As Integer
   
   iErr = -1
   
   Select Case Index
      Case 0  '列印
          If txt1(0) = "" Then
              MsgBox "著作名稱不可空白！", vbInformation, "錯誤！"
              iErr = 0
              GoTo EXITSUB
          End If
          If txt1(13) = "" And txt1(18) = "" Then
              MsgBox "著作人和著作財產權人至少輸入一項！", vbInformation, "錯誤！"
              iErr = 13
              GoTo EXITSUB
          End If
          If Trim(txt1(14) & txt1(15) & txt1(16) & txt1(17)) <> "" And txt1(13) = "" Then
              MsgBox "著作人不可空白！", vbInformation, "錯誤！"
              iErr = 13
              GoTo EXITSUB
          End If
          If Trim(txt1(19) & txt1(20) & txt1(21) & txt1(22)) <> "" And txt1(18) = "" Then
              MsgBox "著作財產權人不可空白！", vbInformation, "錯誤！"
              iErr = 18
              GoTo EXITSUB
          End If
          strBuffer = ""
          For Each oControl In ChkKind
             If oControl.Value = 1 Then
                 strBuffer = strBuffer & "," & oControl.Index
             End If
          Next
          If strBuffer = "" Then
              MsgBox "著作種類至少勾一項！", vbInformation, "錯誤！"
              ChkKind(0).SetFocus
              Exit Sub
          End If
          strBuffer = ""
          For Each oControl In chkItem
             If oControl.Value = 1 Then
                 strBuffer = strBuffer & "," & oControl.Index
             End If
          Next
          If strBuffer = "" Then
              MsgBox "申請登記之登記事項至少勾一項！", vbInformation, "錯誤！"
              chkItem(0).SetFocus
              Exit Sub
          End If
          If Trim(txt1(25) & txt1(26)) = "" Then
              MsgBox "前酬金和前酬金至少輸入一項！！", vbInformation, "錯誤！"
              iErr = 26
              GoTo EXITSUB
          End If
                    
          If Trim(txt1(33)) = "" Then
              MsgBox "經手人不可空白！", vbInformation, "錯誤！"
              iErr = 33
              GoTo EXITSUB
          End If
                
         '檢查四縣市地址
         If txt1(30) <> "" Then
           If CheckTaiwanAddr(txt1(30), "000", "甲方委任人地址") = False Then
              iErr = 30
              GoTo EXITSUB
           End If
         End If
         If Combo2 = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
         End If
          If Trim(txt1(34)) = "" Or Trim(txt1(35)) = "" Or Trim(txt1(36)) = "" Then
             If MsgBox("契約書日期不完整，是否確定？", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
               txt1(34).SetFocus
               txt1_GotFocus 34
               Exit Sub
             End If
          End If
          If ChkSeal.Value = 1 Then
            If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
            Else
               If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If
             bolAddSeal = True
          Else
             bolAddSeal = False
          End If

          Call runWordProc(False)
          PUB_SetOsDefaultPrinter strPrinter
         
          '畫面與客戶檔收據公司別不同更新客戶檔
          'Mark by Lydia 2022/08/30 Lydia 2022/08/30 新功能且為台灣案用的，所以受任人下拉選單預設只能出現台一國際智慧財產事務所，請刪除更新客戶檔的程式碼。
          'strNowCmp = frm210114_1.GetComp(Combo2)
          'If Combo2.Tag <> strNowCmp Then
          '   Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          'End If
          'end 2022/08/30
          
          Screen.MousePointer = vbDefault
          Call RunEndProc(True) ' 刪除暫存檔
          If m_TempPDF <> "" Then ShowPrintOk
          
      Case 1  '回前畫面
          frm210114.Show
          Unload Me
      Case 2 '清空資料
          For Each tb In txt1
              tb.Text = Empty
          Next
          For Each oControl In ChkKind
              oControl.Value = False
          Next
          For Each oControl In chkItem
              oControl.Value = False
          Next
          opt1(0).Value = False
          opt1(1).Value = False

      Case 3  '儲存文檔
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowSave
          If cd1.FileName <> "" Then
              AllObj(0) = "著作權案件委任契約書"
              'TextBox
              For iCount = 1 To 36
                  AllObj(iCount) = txt1(iCount - 1).Text
              Next iCount
              '著作種類
              intI = 38
              For iCount = 0 To 10
                  AllObj(iCount + intI) = ChkKind(iCount).Value
              Next iCount
              '登記事項
              intI = 49
              For iCount = 0 To 6
                  AllObj(iCount + intI) = chkItem(iCount).Value
              Next iCount
              
              AllObj(56) = Combo2.Text '受任人
              
              strBuffer = Join(AllObj, Chr(30))
              strBuffer = StrEncrypt(strBuffer)
              fN = FreeFile
              Open cd1.FileName For Output As fN
              Print #fN, strBuffer
              Close #fN
          End If
          '畫面與客戶檔收據公司別不同更新客戶檔
          'Mark by Lydia 2022/08/30 Lydia 2022/08/30 新功能且為台灣案用的，所以受任人下拉選單預設只能出現台一國際智慧財產事務所，請刪除更新客戶檔的程式碼。
          'strNowCmp = frm210114_1.GetComp(Combo2)
          'If Combo2 <> MsgText(601) And Combo2.Tag <> strNowCmp Then
          '   Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          'End If
          'end 2022/08/30
          
      Case 4  '讀取文檔
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
              If AllObjV(0) = "著作權案件委任契約書" Then
                  cmdOK_Click 2
                  'TextBox
                  For iCount = 1 To 36
                    txt1(iCount - 1).Text = AllObjV(iCount)
                  Next iCount
                  '著作種類
                  intI = 38
                  For iCount = 0 To 10
                    ChkKind(iCount).Value = AllObjV(iCount + intI)
                  Next iCount
                  '登記事項
                  intI = 49
                  For iCount = 0 To 6
                    chkItem(iCount).Value = AllObjV(iCount + intI)
                  Next iCount
                  '受任人
                  If AllObjV(56) = MsgText(601) Then
                      Combo2.ListIndex = 0
                  Else
                      Combo2.Text = AllObjV(56)
                  End If
                  
                  '委任人地址=>檢查地址欄
                  If txt1(27).Text <> "" And txt1(30).Text <> "" Then
                     If CheckCustomerAddr(1, Trim(txt1(27).Text), Trim(txt1(30).Text), "委任人", True) = False Then
                        txt1(30).SetFocus
                     End If
                  End If
                  '讀取收據公司別
                  cmdFind_Click
              Else
                  MsgBox "錯誤格式，此份內容並非 著作權案件委任契約書 格式！", vbExclamation
              End If
          End If

      Case 5   '空白委任書
          If Trim(Combo2) = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
          End If
          '文雄表示用印由下方勾選,可直接空白列印
          If ChkSeal.Value = 1 Then
            If (InStr(UCase(Combo1.Text), "BATCH") > 0 Or InStr(UCase(Combo1.Text), "WRITER") > 0 Or InStr(UCase(Combo1.Text), "PDF") > 0) And Pub_StrUserSt03 <> "M51" Then
               MsgBox "空白用印的印表機不可選擇PDF列印！", vbInformation, "錯誤！"
               Combo1.SetFocus
               Exit Sub
            End If
            'PDF印表機不需詢問,並且份數改為1份
            If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
            Else
               If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If

            bolAddSeal = True
          End If

          Call cmdOK_Click(2) '清空資料
          Call runWordProc(True)
          PUB_SetOsDefaultPrinter strPrinter
          
          m_strCustCode = ""
          bolAddSeal = False
          Screen.MousePointer = vbDefault
          Call RunEndProc(True) ' 刪除暫存檔
          If m_TempPDF <> "" Then ShowPrintOk
      Case Else
   End Select
   Exit Sub
DialogCancel:

EXITSUB:
   If iErr >= 0 Then
       txt1(iErr).SetFocus
       txt1_GotFocus iErr
   End If
End Sub

'只要鍵盤有動作就不斷線
Private Sub Form_KeyPress(KeyAscii As Integer)
   If UCase(Forms(0).Name) = "MDIMAIN" Then Forms(0).tmrConnect.Tag = 0
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
   
   PUB_InitForm210114 Forms(0), Me '委任契約書表單大於主表單，控制主表單放大。
   MoveFormToCenter Me

   PUB_SetPrinter Me.Name, Me.Combo1, strPrinter, , , , , True
 
   '設定公司別下拉選項
   Call PUB_SetCboTofrm210114(Me.Name, Me.Combo2, strCompSeal)
   
   Combo2.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '記錄表單的印表機
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If

   Call RunEndProc(False) '刪除暫存檔
   Set frm210114_7 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)

Dim intLen As Integer
   
   If KeyAscii <> 8 Then
      intLen = GetTextLength(txt1(Index))
      intLen = intLen + GetTextLength(Chr(KeyAscii))
      If CheckLengthIsOK(txt1(Index).Text & Chr(KeyAscii), txt1(Index).MaxLength) = False Then
         KeyAscii = 0
      End If

   End If
   '限數字
   If InStr("01,02,03,04,05,06,07,08,09,10,11,12,14,15,16,19,20,21,26,27,35,36,", Format(Index, "00") & ",") > 0 Then
       If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
           KeyAscii = 0
       End If
   End If

End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) <> "" Then
       txt1(Index).Text = PUB_StringFilter(txt1(Index).Text)
       Cancel = False
       If CheckLengthIsOK(txt1(Index).Text, txt1(Index).MaxLength) = False Then
           txt1(Index).SetFocus
           txt1_GotFocus Index
           Cancel = True
           Exit Sub
       End If
       If InStr("02,05,08,11,15,20,35,", Format(Index, "00") & ",") > 0 Then
           If Val(txt1(Index)) > 12 Or Val(txt1(Index)) < 1 Then
               MsgBox "月份輸入錯誤！", vbExclamation, "操作錯誤！"
               txt1(Index).SetFocus
               txt1_GotFocus Index
               Cancel = True
               Exit Sub
           End If
       ElseIf InStr("03,06,09,12,16,21,36,", Format(Index, "00") & ",") > 0 Then
           If Val(txt1(Index)) > 31 Or Val(txt1(Index)) < 1 Then
               MsgBox "日輸入錯誤！", vbExclamation, "操作錯誤！"
               txt1(Index).SetFocus
               txt1_GotFocus Index
               Cancel = True
               Exit Sub
           End If
       End If
   End If
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

Private Function SetCustTxt(strCUCode As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   SetCustTxt = False
   strCUCode = Left(strCUCode & "000000000", 9)
   StrSQLa = "Select * From Customer Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      SetCustTxt = True
      '申請人中文
      Me.txt1(27).Text = "" & rsA("CU04").Value
      '申請地址
      Me.txt1(30).Text = "" & rsA("CU23").Value
      '電話1
      Me.txt1(31).Text = "" & rsA("CU16").Value
      '代表人1中文
      Me.txt1(29).Text = "" & rsA("CU07").Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Mark by Lydia 2022/08/30 Lydia 2022/08/30 新功能且為台灣案用的，所以受任人下拉選單預設只能出現台一國際智慧財產事務所，請刪除更新客戶檔的程式碼。
'Private Sub UpdReceiptCmp(ByVal stNowCustNo As String, ByVal stNowCmp As String)
'    Dim strUpd As String
'
'    '同業務區或為MCTF同組人員才可回寫收據公司別
'    If ChkSameCuArea(stNowCustNo, strUserNum) = False Then Exit Sub
'
'    strUpd = "Update Customer Set CU164='" & stNowCmp & "' " & _
'                  "Where CU01='" & Left(stNowCustNo, 8) & "' And CU02='" & Mid(stNowCustNo, 9, 1) & "' "
'    Pub_SeekTbLog strUpd
'    cnnConnection.Execute "begin user_data.user_enabled:=1; " & strUpd & " ; end; "
'End Sub
'end 2022/08/30

'下載Word範本套印
Private Sub runWordProc(ByVal pSpace As Boolean)
Dim iStr(1 To 29) As String    '用印記錄(全文)
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String
Dim strName As String
Dim strText
Dim intA As Integer
Dim m_FileName As String, m_TempFileName As String
Dim m_DefPath As String
Dim oShape
Dim oWord

On Error GoTo ErrHand

   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & Pub_GetSpecMan("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000300-0-07 智權部委任契約書_著作權.docx", "M51", "000300", "0", "07", "4", "1")

   m_DefPath = App.path & "\" & strUserNum
   Call Pub_ChkExcelPath(m_DefPath)
   
   m_TempPDF = ""
   '變更Word印表機
   PUB_SetOsDefaultPrinter Combo1
   PUB_SetWordActivePrinter
   
    strDetail = ""
    
   '下載範本檔: M51-000300-0-07 智權部委任契約書_著作權.docx
   m_FileName = Pub_RepFileName(IIf(pSpace = False, IIf(Trim(txt1(27)) <> "", Mid(Trim(txt1(27)), 1, 4), Mid(Trim(txt1(0)), 1, 4)), "空白"))
   m_FileName = "$$" & strUserNum & "_著作權_" & m_FileName & ".docx"

   If PUB_GetSampleFile(m_FileName, "M51-000300-0-07", , m_DefPath) = False Then
        Exit Sub
   End If
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If

   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
   
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 1 To 30
         strName = "PS" & Format(intA, "000")
         strText = ""
'-------第一條
         If intA = 1 Then
              '著作名稱
              strText = PUB_StrToStr(txt1(0) & " ", 62)
         ElseIf intA = 2 Then
              '著作完成日+公開發表日
              strText = "著作完成日：" & Pub_StrToCenter(txt1(1), 5) & "年" & Pub_StrToCenter(txt1(2), 5) & "月" & Pub_StrToCenter(txt1(3), 5) & "日" & String(6, " ") & _
                           "公開發表日：" & Pub_StrToCenter(txt1(4), 5) & "年" & Pub_StrToCenter(txt1(5), 5) & "月" & Pub_StrToCenter(txt1(6), 5) & "日"
         ElseIf intA = 3 Then
              '發　行　日+著作權轉讓日
              strText = "發　行　日：" & Pub_StrToCenter(txt1(7), 5) & "年" & Pub_StrToCenter(txt1(8), 5) & "月" & Pub_StrToCenter(txt1(9), 5) & "日" & String(6, " ") & _
                           "著作權轉讓日：" & Pub_StrToCenter(txt1(10), 5) & "年" & Pub_StrToCenter(txt1(11), 5) & "月" & Pub_StrToCenter(txt1(12), 5) & "日"
         ElseIf intA = 4 Then
              '著作人
              strText = PUB_StrToStr(txt1(13), 32, True)
         ElseIf intA = 5 Then
              '(著作人)出生/設立日
              strText = Pub_StrToCenter(txt1(14), 4) & "年" & Pub_StrToCenter(txt1(15), 4) & "月" & Pub_StrToCenter(txt1(16), 4) & "日"
         ElseIf intA = 6 Then
              '(著作人)地址
              strText = PUB_StrToStr(txt1(17), 62)
         ElseIf intA = 7 Then
              '著作財產權人
              strText = PUB_StrToStr(txt1(18), 28, True)
         ElseIf intA = 8 Then
              '(著作財產權人)出生/設立日
              strText = Pub_StrToCenter(txt1(19), 4) & "年" & Pub_StrToCenter(txt1(20), 4) & "月" & Pub_StrToCenter(txt1(21), 4) & "日"
         ElseIf intA = 9 Then
              '(著作財產權人)地址
              strText = PUB_StrToStr(txt1(22), 62)
'-------第二條
         ElseIf intA = 10 Then
              '著作種類1
              strText = IIf(ChkKind(0).Value = 1, "■", "□") & "語文著作" & String(6, " ") & IIf(ChkKind(1).Value = 1, "■", "□") & "音樂著作" & String(8, " ") & IIf(ChkKind(2).Value = 1, "■", "□") & "美術著作" & String(6, " ") & IIf(ChkKind(3).Value = 1, "■", "□") & "攝影著作"
         ElseIf intA = 11 Then
              '著作種類2
              strText = IIf(ChkKind(4).Value = 1, "■", "□") & "圖形著作" & String(6, " ") & IIf(ChkKind(5).Value = 1, "■", "□") & "視聽著作" & String(8, " ") & IIf(ChkKind(6).Value = 1, "■", "□") & "錄音著作" & String(6, " ") & IIf(ChkKind(7).Value = 1, "■", "□") & "建築著作"
         ElseIf intA = 12 Then
              '著作種類3
              strText = IIf(ChkKind(8).Value = 1, "■", "□") & "電腦程式著作" & String(2, " ") & IIf(ChkKind(9).Value = 1, "■", "□") & "戲劇、舞蹈著作" & String(2, " ") & IIf(ChkKind(10).Value = 1, "■", "□") & "其他：" & PUB_StrToStr(txt1(23), 20)
         ElseIf intA = 13 Then
              '登記事項1
              strText = IIf(chkItem(0).Value = 1, "■", "□") & "著作人登記" & String(22, " ") & IIf(chkItem(1).Value = 1, "■", "□") & "讓與登記"
         ElseIf intA = 14 Then
              '登記事項2
              strText = IIf(chkItem(2).Value = 1, "■", "□") & "著作財產權人登記" & String(16, " ") & IIf(chkItem(3).Value = 1, "■", "□") & "變更登記"
         ElseIf intA = 15 Then
              '登記事項3
              strText = IIf(chkItem(4).Value = 1, "■", "□") & "首次公開發表日或發行日登記" & String(6, " ") & IIf(chkItem(5).Value = 1, "■", "□") & "其他：" & PUB_StrToStr(txt1(24), 20)
         ElseIf intA = 16 Then
              '登記事項4
              strText = IIf(chkItem(6).Value = 1, "■", "□") & "專屬授權或處分之限制登記"
'-------第三條        委辦費用
         ElseIf intA = 17 Then
                strExc(1) = "": strExc(2) = ""
                If Val(Trim(txt1(25))) = 0 Then
                     If pSpace = True Then strExc(1) = "　　　　　　　　　　　　元整"
                Else
                     strExc(1) = ChangeNumber(txt1(25))
                End If
               strExc(3) = PUB_StrToStr("　　一、前酬金新台幣　|#PS018#|，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 88 - Len(strExc(1)) * 2 + 9)
               strExc(4) = "　　　　" & Replace("　　一、前酬金新台幣　|#PS018#|，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", PUB_StrToStr("　　一、前酬金新台幣　|#PS018#|，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 88 - Len(strExc(1)) * 2 + 9), "")
          
              If pSpace = True Or (Val(Trim(txt1(25))) > 0 And Val(Trim(txt1(26))) > 0) Then
                   strText = strExc(3) & vbCrLf & strExc(4) & vbCrLf & _
                                "　　二、後酬金新台幣　|#PS019#|，於本程序終結時，由甲方一次付清。"
              Else
                   If Val(Trim(txt1(25))) > 0 Then
                        strText = Replace(strExc(3) & vbCrLf & strExc(4), "一、", "　　") & "|#PS019#|"
                   ElseIf Val(Trim(txt1(26))) > 0 Then
                        strText = "|#PS018#|　　　　後酬金新台幣　|#PS019#|，於本程序終結時，由甲方一次付清。"
                   Else
                        strText = "　　|#PS018#||#PS019#|"
                   End If
              End If
         ElseIf intA = 18 Then
              '前酬金
              If Val(Trim(txt1(25))) = 0 Then
                   If pSpace = True Then strText = "　　　　　　　　　　　　元整"
              Else
                   strText = ChangeNumber(txt1(25))
              End If
         ElseIf intA = 19 Then
              '後酬金
              If Val(Trim(txt1(26))) = 0 Then
                   If pSpace = True Then strText = "　　　　　　　　　　　　元整"
              Else
                   strText = ChangeNumber(txt1(26))
              End If
         ElseIf intA = 20 Then
              '會稿 (文字方塊)
              strText = IIf(opt1(0).Value = True, "V", "")
              .ActiveDocument.Shapes("Text Box 1").Select
         ElseIf intA = 21 Then
              '不會稿(文字方塊)
              strText = IIf(opt1(1).Value = True, "V", "")
         ElseIf intA = 22 Then
              '委任人（甲方）
              strText = PUB_StrToStr(txt1(27), 70)
         ElseIf intA = 23 Then
              '委任人ID+代表人
              strText = PUB_StrToStr(txt1(28) & " ", 25, True) & "代表人：" & PUB_StrToStr(txt1(29), 18)
         ElseIf intA = 24 Then
              '委任人（地　址）
              strText = PUB_StrToStr(txt1(30) & " ", 70)
         ElseIf intA = 25 Then
              '委任人電話+傳真
              strText = PUB_StrToStr(txt1(31) & " ", 25, True) & "傳　真：" & PUB_StrToStr(txt1(32), 18)
         ElseIf intA = 26 Then
              '受任人（乙方）
              strText = Combo2.Text
         ElseIf intA = 27 Then
              '經手人
              strText = PUB_StrToStr(txt1(33), 70)
         ElseIf intA = 28 Then
              '地　址
              strText = PUB_SetAddrTofrm210114(Combo2.Text)
         ElseIf intA = 29 Then
              strText = "        中    華    民    國 " & Pub_StrToCenter(txt1(34), 8) & "年" & Pub_StrToCenter(txt1(35), 8) & "月" & Pub_StrToCenter(txt1(36), 8) & "日"
         ElseIf intA = 32 Then  '用印
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
            If intA = 18 Or intA = 19 Then
                '金額要粗體
                .Selection.Font.Bold = True
            End If
            '保留;因為先全部以細明體-ExtB,最後全選改字型;
            If (intA >= 22 And intA <= 24) Or intA = 26 Then
               '有Unicode字需要換字型
               .Selection.Font.Name = "細明體-ExtB"
            End If
            
            If intA <> 30 Then  '用印記錄
               iStr(intA) = strText
            End If
            If intA = 30 And bolAddSeal = True Then  '公司章: 放在受任人的儲存格
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
                        oShape.Left = .CentimetersToPoints(8.25)
                        oShape.Top = .CentimetersToPoints(0.3)
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

            If intA = 18 Or intA = 19 Then
                '金額要粗體 =>還原
                .Selection.Font.Bold = False
            End If
            If (intA >= 22 And intA <= 24) Or intA = 26 Then
               '有Unicode字需要換字型 =>還原
               .Selection.Font.Name = "標楷體"
            End If
            If intA = 21 Then  '會稿／不會稿
                .Selection.HomeKey Unit:=wdStory
            End If
            
         End If
      Next intA
      '因為先全部以細明體-ExtB,最後全選改字型;
      .Selection.WholeStory
      .Selection.Font.Name = "標楷體"
   End With

   '因為受PDF redirect設定灰階列印影響，改成Word直接印
   intA = IIf(Val(txtPCnt) = 0, 1, Val(txtPCnt))
   For intI = 1 To intA
       g_WordAp.PrintOut Background:=False, Range:=4, Item:=0, Copies:=1, Pages:="1", Collate:=True
   Next intI
   
   '保留: 存檔
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing '避免快速開啟Word,程式出錯
   m_TempPDF = m_FileName

   
If bolAddSeal = True Then  '用印記錄
                   strDetail = "著作權案件委任契約書" & vbCrLf
   strDetail = strDetail & "委任人(甲方)茲委任受任人(乙方)辦理國內著作權案件，經雙方同意條件如下：" & vbCrLf
   strDetail = strDetail & "第一條　 一、著作名稱：" & iStr(1) & vbCrLf
   strDetail = strDetail & "　　　　　" & iStr(2) & vbCrLf
   strDetail = strDetail & "　　　　　" & iStr(3) & vbCrLf
   strDetail = strDetail & "　　　　 二、著作人：" & iStr(4) & "　　　　出生/設立日：" & iStr(5) & vbCrLf
   strDetail = strDetail & "　　　　 　地　址：" & iStr(6) & vbCrLf
   strDetail = strDetail & "　　　　 三、著作財產權人：" & iStr(7) & "　　　　出生/設立日：" & iStr(8) & vbCrLf
   strDetail = strDetail & "　　　　 　地　址：" & iStr(9) & vbCrLf
   strDetail = strDetail & "第二條　委辦範圍：" & vbCrLf
   strDetail = strDetail & "　　　　著作種類：" & iStr(10) & vbCrLf
   strDetail = strDetail & "　　　　　　　　　" & iStr(11) & vbCrLf
   strDetail = strDetail & "　　　　　　　　　" & iStr(12) & vbCrLf
   strDetail = strDetail & "　　　　性　　質：申請登記" & vbCrLf
   strDetail = strDetail & "　　　　登記事項：" & iStr(13) & vbCrLf
   strDetail = strDetail & "　　　　　　　　　" & iStr(14) & vbCrLf
   strDetail = strDetail & "　　　　　　　　　" & iStr(15) & vbCrLf
   strDetail = strDetail & "　　　　　　　　　" & iStr(16) & vbCrLf
   '前酬金
   If Trim(iStr(18)) = "" And pSpace = False Then
       strExc(3) = ""
       strExc(4) = ""
   Else
       strExc(3) = StrToStr("　　一、前酬金新台幣　" & iStr(18) & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40)
       strExc(4) = "　　　　" & Replace("　　一、前酬金新台幣　" & iStr(18) & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", StrToStr("　　一、前酬金新台幣　" & iStr(18) & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40), "")
   End If
   '後酬金
   If Trim(iStr(19)) = "" And pSpace = False Then
       strExc(5) = ""
       strExc(6) = ""
       strExc(3) = Replace(strExc(3), "一、", "　　")
       strExc(4) = Replace(strExc(4), "一、", "　　")
   Else
       strExc(5) = StrToStr("　　二、後酬金新台幣　" & iStr(19) & "，於本程序終結時，由甲方一次付清。", 40)
       strExc(6) = "　　　　" & Replace("　　二、後酬金新台幣　" & iStr(19) & "，於本程序終結時，由甲方一次付清。", StrToStr("　　二、後酬金新台幣　" & iStr(19) & "，於本程序終結時，由甲方一次付清。", 40), "")
       If Trim(iStr(18)) = "" Then
         strExc(5) = Replace(strExc(5), "二、", "　　")
         strExc(6) = Replace(strExc(6), "二、", "　　")
       End If
   End If
   strDetail = strDetail & "第二條　委辦費月：" & vbCrLf & IIf(strExc(3) & strExc(4) <> "", strExc(3) & vbCrLf & strExc(4) & vbCrLf, "") & _
                                    IIf(strExc(5) & strExc(6) <> "", strExc(5) & vbCrLf & strExc(6) & vbCrLf, "")
   strDetail = strDetail & "第四條　乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，並不得發生足以影響甲方" & vbCrLf
   strDetail = strDetail & "　　　　權益之疏誤，否則應對甲方負損害賠償責任。但以不超過第三條所載前酬金金額之三倍為限。" & vbCrLf
   strDetail = strDetail & "第五條　甲方確保所交付乙方之資料均無虛偽情事，如因不實致生損害或法律責任時，概由甲方負責，" & vbCrLf
   strDetail = strDetail & "　　　　與乙方無關。" & vbCrLf
   strDetail = strDetail & "第六條　乙方於辦理過程中，應隨時將辦理經過如申請日期、案號及其他重要函件，儘速通知或交付甲" & vbCrLf
   strDetail = strDetail & "　　　　方。但甲方於簽約後變更連絡處所，未即時通知乙方，因而連絡不及致延誤時限時，乙方不負" & vbCrLf
   strDetail = strDetail & "　　　　責任。" & vbCrLf
   strDetail = strDetail & "第七條　凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時限，乙方不負責。經乙方通知" & vbCrLf
   strDetail = strDetail & "　　　　甲方繳費而未依限繳納者，亦同。" & vbCrLf
   strDetail = strDetail & "第八條　甲方如逕自撤回所委辦程序，或未經乙方同意終止契約時，所約定之費用，仍應全數給付。" & vbCrLf
   strDetail = strDetail & "第九條　本契約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，但有增刪修改時，需甲乙雙方於" & vbCrLf
   strDetail = strDetail & "　　　　更動處蓋章始生效力，並由雙方各執乙份為憑。" & vbCrLf
   If opt1(0).Value = True Then strDetail = strDetail & "案件需會稿" & vbCrLf
   If opt1(1).Value = True Then strDetail = strDetail & "案件不需會稿" & vbCrLf
   strDetail = strDetail & "甲方：　委任人：" & iStr(22) & vbCrLf
   strDetail = strDetail & "　　　　ID NO.：" & iStr(23) & vbCrLf
   strDetail = strDetail & "　　　　地　址：" & iStr(24) & vbCrLf
   strDetail = strDetail & "　　　　電　話：" & iStr(25) & vbCrLf
   strDetail = strDetail & "乙方：　受任人：" & iStr(26) & vbCrLf
   strDetail = strDetail & "　　　　經手人：" & iStr(27) & vbCrLf
   strDetail = strDetail & "　　　　地　址：" & iStr(28) & vbCrLf
   strDetail = strDetail & "　　　　電　話：(02)25061023(總機)　　　FAX：(02)25011666" & vbCrLf
   strDetail = strDetail & "　　　　網　址：www.taie.com.tw　　　　 E-mail：ipdept@taie.com.tw" & vbCrLf
   strDetail = strDetail & "　　" & iStr(29)
   If PUB_AddRecSeal("7", txtPCnt.Text, IIf(pSpace = True, "Y", ""), strDetail, Combo2.Text) Then
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

'刪除暫存檔
Private Sub RunEndProc(ByVal bolSleep As Boolean)
   If bolSleep = True Then Sleep 3000
   PUB_KillTempFile (strUserNum & "\$$" & strUserNum & "*_著作權*.*")
   PUB_KillTempFile (strUserNum & "\$$" & Me.Name & "*.*")
    
End Sub

