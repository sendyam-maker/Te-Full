VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210114_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "委任契約書-P"
   ClientHeight    =   6540
   ClientLeft      =   1788
   ClientTop       =   2832
   ClientWidth     =   9312
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9312
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0FFFF&
      Caption         =   "空白列印"
      Height          =   330
      Index           =   5
      Left            =   3600
      Style           =   1  '圖片外觀
      TabIndex        =   83
      Top             =   30
      Width           =   920
   End
   Begin VB.CheckBox ChkSeal 
      Caption         =   "用印"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   4860
      TabIndex        =   82
      Top             =   6225
      Width           =   735
   End
   Begin VB.CheckBox ChkDou 
      Caption         =   "多發明人 多申請人"
      ForeColor       =   &H00FF00FF&
      Height          =   510
      Left            =   7950
      TabIndex        =   81
      Top             =   3570
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind2 
      Caption         =   "搜尋發明人(&I)"
      Height          =   330
      Left            =   4920
      TabIndex        =   80
      Top             =   3810
      Width           =   1365
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋申請人(&Q)"
      Height          =   330
      Left            =   6360
      TabIndex        =   79
      Top             =   3810
      Width           =   1365
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm210114_1.frx":0000
      Left            =   6570
      List            =   "frm210114_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   42
      Top             =   6195
      Width           =   2475
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "本案所涉之技術內容並非在中國大陸境內完成的發明或者實用新型"
      Height          =   240
      Index           =   12
      Left            =   1095
      TabIndex        =   27
      Top             =   3330
      Width           =   7380
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "實體審查"
      Height          =   240
      Index           =   11
      Left            =   1095
      TabIndex        =   24
      Top             =   3060
      Width           =   1080
   End
   Begin VB.TextBox txtPCnt 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Left            =   7800
      MaxLength       =   1
      TabIndex        =   43
      Text            =   "2"
      Top             =   60
      Width           =   270
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8640
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "讀取文檔"
      Height          =   330
      Index           =   4
      Left            =   5448
      TabIndex        =   46
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "儲存文檔"
      Height          =   330
      Index           =   3
      Left            =   4524
      TabIndex        =   45
      Top             =   30
      Width           =   920
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   360
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   3645
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "frm210114_1.frx":0043
         Left            =   660
         List            =   "frm210114_1.frx":0045
         Style           =   2  '單純下拉式
         TabIndex        =   44
         Top             =   30
         Width           =   2840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "印表機"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   77
         Top             =   90
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "清空資料"
      Height          =   330
      Index           =   2
      Left            =   6372
      TabIndex        =   47
      Top             =   30
      Width           =   920
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "新型技術報告"
      Height          =   240
      Left            =   6660
      TabIndex        =   13
      Top             =   2520
      Width           =   1380
   End
   Begin VB.OptionButton opt1 
      Caption         =   "不會稿"
      Height          =   210
      Index           =   1
      Left            =   2220
      TabIndex        =   31
      Top             =   4290
      Width           =   930
   End
   Begin VB.OptionButton opt1 
      Caption         =   "會稿"
      Height          =   210
      Index           =   0
      Left            =   1125
      TabIndex        =   30
      Top             =   4290
      Width           =   900
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "繳年費"
      Height          =   240
      Index           =   9
      Left            =   4620
      TabIndex        =   23
      Top             =   2790
      Width           =   885
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "答辯"
      Height          =   240
      Index           =   7
      Left            =   2010
      TabIndex        =   21
      Top             =   2790
      Width           =   720
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "舉發"
      Height          =   240
      Index           =   6
      Left            =   1095
      TabIndex        =   20
      Top             =   2790
      Width           =   840
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "領證"
      Height          =   240
      Index           =   8
      Left            =   3300
      TabIndex        =   22
      Top             =   2790
      Width           =   1080
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "行政訴訟"
      Height          =   240
      Index           =   5
      Left            =   30
      TabIndex        =   19
      Top             =   2790
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "訴願"
      Height          =   240
      Index           =   4
      Left            =   5700
      TabIndex        =   18
      Top             =   2520
      Width           =   840
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "其他："
      Height          =   240
      Index           =   10
      Left            =   2175
      TabIndex        =   25
      Top             =   3060
      Width           =   870
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "再審申復"
      Height          =   240
      Index           =   3
      Left            =   4380
      TabIndex        =   17
      Top             =   2520
      Width           =   1080
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "初審申復"
      Height          =   240
      Index           =   2
      Left            =   3090
      TabIndex        =   16
      Top             =   2520
      Width           =   1080
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "再審"
      Height          =   240
      Index           =   1
      Left            =   2175
      TabIndex        =   15
      Top             =   2520
      Width           =   840
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "申請"
      Height          =   240
      Index           =   0
      Left            =   1095
      TabIndex        =   14
      Top             =   2520
      Width           =   840
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "設計"
      Height          =   240
      Index           =   2
      Left            =   3090
      TabIndex        =   12
      Top             =   2265
      Width           =   840
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "新型"
      Height          =   240
      Index           =   1
      Left            =   2175
      TabIndex        =   11
      Top             =   2265
      Width           =   840
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "發明"
      Height          =   240
      Index           =   0
      Left            =   1095
      TabIndex        =   10
      Top             =   2265
      Width           =   840
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面"
      Height          =   330
      Index           =   1
      Left            =   8400
      TabIndex        =   49
      Top             =   30
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印       份"
      Height          =   330
      Index           =   0
      Left            =   7296
      TabIndex        =   48
      Top             =   30
      Width           =   1100
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   10
      Left            =   3090
      TabIndex        =   26
      Top             =   3030
      Width           =   4995
      VariousPropertyBits=   671105051
      MaxLength       =   42
      Size            =   "8811;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   9
      Left            =   1845
      TabIndex        =   9
      Top             =   1950
      Width           =   7425
      VariousPropertyBits=   671105051
      MaxLength       =   54
      Size            =   "13097;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   22
      Left            =   3660
      TabIndex        =   41
      Top             =   6180
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1244;547"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   21
      Left            =   2580
      TabIndex        =   40
      Top             =   6180
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1244;547"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   20
      Left            =   1545
      TabIndex        =   39
      Top             =   6180
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1244;547"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   19
      Left            =   1545
      TabIndex        =   38
      Top             =   5850
      Width           =   7800
      VariousPropertyBits=   671105051
      MaxLength       =   54
      Size            =   "13758;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   18
      Left            =   4260
      TabIndex        =   37
      Top             =   5520
      Width           =   5085
      VariousPropertyBits=   671105051
      MaxLength       =   26
      Size            =   "8969;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   17
      Left            =   1545
      TabIndex        =   36
      Top             =   5520
      Width           =   1905
      VariousPropertyBits=   671105051
      MaxLength       =   18
      Size            =   "3360;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   16
      Left            =   1545
      TabIndex        =   35
      Top             =   5190
      Width           =   7800
      VariousPropertyBits=   671105051
      MaxLength       =   54
      Size            =   "13758;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   15
      Left            =   4260
      TabIndex        =   34
      Top             =   4860
      Width           =   5085
      VariousPropertyBits=   671105051
      MaxLength       =   26
      Size            =   "8969;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   14
      Left            =   1545
      TabIndex        =   33
      Top             =   4860
      Width           =   1905
      VariousPropertyBits=   671105051
      MaxLength       =   18
      Size            =   "3360;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   13
      Left            =   1545
      TabIndex        =   32
      Top             =   4530
      Width           =   7800
      VariousPropertyBits=   671105051
      MaxLength       =   54
      Size            =   "13758;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   12
      Left            =   1275
      TabIndex        =   29
      ToolTipText     =   "輸入數字"
      Top             =   3930
      Width           =   2145
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "3784;547"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   11
      Left            =   1275
      TabIndex        =   28
      ToolTipText     =   "輸入數字"
      Top             =   3600
      Width           =   2145
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "3784;547"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   8
      Left            =   8340
      TabIndex        =   8
      Top             =   1644
      Width           =   930
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1640;547"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   7
      Left            =   705
      TabIndex        =   7
      Top             =   1644
      Width           =   7035
      VariousPropertyBits=   671105051
      MaxLength       =   50
      Size            =   "12409;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   6
      Left            =   7290
      TabIndex        =   6
      Top             =   1338
      Width           =   1980
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "3492;547"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   5
      Left            =   1425
      TabIndex        =   5
      Top             =   1338
      Width           =   5040
      VariousPropertyBits=   671105051
      MaxLength       =   38
      Size            =   "8890;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   4
      Left            =   8340
      TabIndex        =   4
      Top             =   1032
      Width           =   930
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1640;547"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   3
      Left            =   705
      TabIndex        =   3
      Top             =   1032
      Width           =   7035
      VariousPropertyBits=   671105051
      MaxLength       =   50
      Size            =   "12409;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   2
      Left            =   7290
      TabIndex        =   2
      Top             =   726
      Width           =   1980
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "3492;547"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   1
      Left            =   1425
      TabIndex        =   1
      Top             =   726
      Width           =   5040
      VariousPropertyBits=   671105051
      MaxLength       =   38
      Size            =   "8890;547"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   310
      Index           =   0
      Left            =   1425
      TabIndex        =   0
      Top             =   420
      Width           =   7845
      VariousPropertyBits=   671105051
      MaxLength       =   58
      Size            =   "13838;547"
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
      Index           =   6
      Left            =   5610
      TabIndex        =   90
      Top             =   6255
      Width           =   180
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊為必填欄位"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   3180
      TabIndex        =   89
      Top             =   4290
      Width           =   1080
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   1350
      TabIndex        =   88
      Top             =   5915
      Width           =   180
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   3
      Left            =   1350
      TabIndex        =   87
      Top             =   4595
      Width           =   180
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   900
      TabIndex        =   86
      Top             =   4290
      Width           =   180
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   870
      TabIndex        =   85
      Top             =   2550
      Width           =   180
   End
   Begin VB.Label lblStar 
      AutoSize        =   -1  'True
      Caption         =   "＊"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   1110
      TabIndex        =   84
      Top             =   1403
      Width           =   180
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "受任人："
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5820
      TabIndex        =   78
      Top             =   6255
      Width           =   720
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "請輸入數字"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3870
      TabIndex        =   75
      Top             =   3997
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "請輸入數字"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3885
      TabIndex        =   74
      Top             =   3667
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "指定聯絡人及地址："
      Height          =   180
      Left            =   105
      TabIndex        =   73
      Top             =   2010
      Width           =   1620
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "中　華　民　國　　　　　年　　　　　月　　　　　日"
      Height          =   180
      Left            =   195
      TabIndex        =   72
      Top             =   6255
      Width           =   4500
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "乙方：經手人"
      Height          =   180
      Left            =   195
      TabIndex        =   71
      Top             =   5915
      Width           =   1080
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "傳　真："
      Height          =   180
      Left            =   3555
      TabIndex        =   70
      Top             =   5585
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "電　話："
      Height          =   180
      Left            =   735
      TabIndex        =   69
      Top             =   5585
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "地　址："
      Height          =   180
      Left            =   735
      TabIndex        =   68
      Top             =   5255
      Width           =   720
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代表人："
      Height          =   180
      Left            =   3555
      TabIndex        =   67
      Top             =   4925
      Width           =   720
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "ID  NO.："
      Height          =   180
      Left            =   720
      TabIndex        =   66
      Top             =   4925
      Width           =   735
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "甲方：委任人"
      Height          =   180
      Left            =   195
      TabIndex        =   65
      Top             =   4595
      Width           =   1080
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "元整"
      Height          =   180
      Left            =   3465
      TabIndex        =   64
      Top             =   3997
      Width           =   360
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "元整"
      Height          =   180
      Left            =   3465
      TabIndex        =   63
      Top             =   3667
      Width           =   360
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "後   酬   金："
      Height          =   180
      Left            =   135
      TabIndex        =   62
      Top             =   3997
      Width           =   990
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "前   酬   金："
      Height          =   180
      Left            =   135
      TabIndex        =   61
      Top             =   3667
      Width           =   990
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "性　質："
      Height          =   180
      Left            =   105
      TabIndex        =   60
      Top             =   2550
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請種類："
      Height          =   180
      Left            =   105
      TabIndex        =   59
      Top             =   2250
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "國籍："
      Height          =   180
      Left            =   7800
      TabIndex        =   58
      Top             =   1709
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "地址："
      Height          =   180
      Left            =   105
      TabIndex        =   57
      Top             =   1709
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "I.D.NO.："
      Height          =   180
      Left            =   6525
      TabIndex        =   56
      Top             =   1403
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "申  請　人："
      Height          =   180
      Left            =   105
      TabIndex        =   55
      Top             =   1403
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "國籍："
      Height          =   180
      Left            =   7785
      TabIndex        =   54
      Top             =   1097
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "地址："
      Height          =   180
      Left            =   105
      TabIndex        =   53
      Top             =   1097
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "I.D.NO.："
      Height          =   180
      Left            =   6510
      TabIndex        =   52
      Top             =   791
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "發明或創作人："
      Height          =   180
      Index           =   0
      Left            =   105
      TabIndex        =   51
      Top             =   791
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發明創作名稱："
      Height          =   180
      Left            =   105
      TabIndex        =   50
      Top             =   485
      Width           =   1260
   End
End
Attribute VB_Name = "frm210114_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/14 改成Form2.0 ; txt1(index)、Printer改成Word列印
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
'Added by Lydia 2022/01/14  加入圖片用(Word)
Const msoFalse = 0
Const msoLineSolid = 1
Const msoLineSingle = 1
Const msoTrue = -1
Const msoPictureAutomatic = 1
'end 2022/01/14
Dim m_TempPDF As String 'Added by Lydia 2022/01/19
Dim m_TempFN As String 'Added by Lydia 2022/01/22

Private Sub Chk1_Click(Index As Integer)
'先清空
Dim i As Integer
   
   If Chk1(Index).Value = vbChecked Then
      'cancel by sonia 2015/11/18 專利種類開放可複選
      'For i = 0 To 2
      '   If i <> Index Then
      '      Chk1(i).Value = vbUnchecked
      '   End If
      'Next i
      'end 2015/11/18
   End If
End Sub

Private Sub Chk2_Click(Index As Integer)
   If Index = 10 Then
       If Chk2(Index).Value = vbChecked Then
           txt1(10).SetFocus
       End If
   End If
End Sub

'Add By Sindy 2010/4/29
Private Sub cmdFind_Click()
   Dim strCmpName As String, strMsg As String 'Add by Amy 2016/08/19
   
   If Me.txt1(5).Text = "" Then
      MsgBox "請輸入申請人中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txt1(5).SetFocus
      Exit Sub
   End If

   frm090801_1.m_Type = 0  'add by Lydia 2014/9/22
   If ChkDou.Value = 1 Then
     frm090801_1.m_DouChk = True '可複選
   Else
     frm090801_1.m_DouChk = False
   End If

   Set frm090801_1.m_frm0908A = Me
   
   frm090801_1.m_strCustChnName = Me.txt1(5).Text
   frm090801_1.lblName.Caption = Me.txt1(5).Text

   
   m_blnOneRec = False
   m_strCustCode = ""
   txt1(5).Tag = "" 'Added by Morgan 2012/9/11
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
      Combo2.Tag = GetReceiptCmp(Left(strNowCustNo, 8), Mid(strNowCustNo, 9, 1), "P", "000", False, strCmpName, Me.Name)
      If Combo2.Tag <> MsgText(601) And Combo2 <> MsgText(601) And Combo2.Tag <> GetComp(Combo2) Then
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
      txt1(5).Tag = txt1(5) 'Added by Morgan 2012/9/11
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
'Modified by Lydia 20220/01/14
'Dim tb As TextBox
'Dim op As OptionButton
'Dim ck As CheckBox
Dim tb As Control, op As Control, ck As Control
'end 2022/01/14
Dim fN As Integer
Dim strBuffer As String
'Modify By Sindy 2010/3/17
'Dim AllObj(0 To 40) As String
'Modify By Sindy 2010/9/9
'Dim AllObj(0 To 42) As String
'Modified by Lydia 2022/01/14
'Dim AllObj(0 To 43) As String
Dim AllObj(0 To 43)
Dim AllObjV As Variant
'Add by Amy 2016/08/19 目前收據公司別
Dim strNowCmp As String
Dim strMsg As String 'Added by Lydia 2019/04/16
   
   strMsg = "" 'Added by Lydia 2019/04/16

   Select Case Index
      Case 0
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(0)) = "" Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "發明創作名稱不可空白！", vbInformation, "錯誤！"
              'txt1(0).SetFocus
              'Txt1_GotFocus 0
              'Exit Sub
              strMsg = strMsg & "、發明創作名稱"
          End If
          'Modified by Lydia 2017/03/28 +Trim
          'Modified by Lydia 2017/04/21 拿掉Trim
          If txt1(1) = "" Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "發明或創作人不可空白！", vbInformation, "錯誤！"
              'txt1(1).SetFocus
              'Txt1_GotFocus 1
              'Exit Sub
              strMsg = strMsg & "、發明或創作人"
          End If
         
          'Modified by Lydia 2017/03/28 +Trim
          If Trim(txt1(5)) = "" Then
              MsgBox "申請人不可空白！", vbInformation, "錯誤！"
              txt1(5).SetFocus
              txt1_GotFocus 5
              Exit Sub
          End If
          'Modify By Sindy 2009/10/22
          'If txt1(11) = "" And txt1(12) = "" Then
          'Modified by Lydia 2019/04/16 開放部分欄位空白
'          If Val(txt1(11)) = 0 And Val(txt1(12)) = 0 Then
'          '2009/10/22 End
'              MsgBox "費用最少輸一個不可空白！", vbInformation, "錯誤！"
'              txt1(11).SetFocus
'              Txt1_GotFocus 11
'              Exit Sub
'          End If
          If Val(txt1(11)) = 0 Then
               strMsg = strMsg & "、前酬金"
          End If
          'end 2019/04/16
          If opt1(0).Value <> True And opt1(1).Value <> True Then
              MsgBox "會不會稿要選擇一項！", vbInformation, "錯誤！"
              Exit Sub
          End If
          'Modified by Lydia 2019/04/16 必填欄位
      '    If txt1(13) = "" Then
      '        MsgBox "委任人不可空白！", vbInformation, "錯誤！"
      '        txt1(13).SetFocus
      '        txt1_GotFocus 13
      '        Exit Sub
      '    End If
          If Trim(txt1(13)) = "" Then
              MsgBox "委任人不可空白！", vbInformation, "錯誤！"
              txt1(13).SetFocus
              txt1_GotFocus 13
              Exit Sub
          End If
          'end 2019/04/16
          
          'Modified by Lydia 2017/03/28 +Trim
          If Trim(txt1(19)) = "" Then
              MsgBox "經手人不可空白！", vbInformation, "錯誤！"
              txt1(19).SetFocus
              txt1_GotFocus 19
              Exit Sub
          End If
          
          'Modified by Lydia 2019/04/16 開放部分欄位空白
          If Chk1(0).Value = vbUnchecked And Chk1(1).Value = vbUnchecked And Chk1(2).Value = vbUnchecked Then
              'Modified by Lydia 2019/04/16 開放部分欄位空白
              'MsgBox "專利種類最少選一個！", vbInformation, "錯誤！"
              'Exit Sub
              strMsg = strMsg & "、申請種類"
          End If
          If Chk2(0).Value = vbUnchecked And Chk2(1).Value = vbUnchecked And Chk2(2).Value = vbUnchecked And Chk2(3).Value = vbUnchecked And Chk2(4).Value = vbUnchecked And Chk2(5).Value = vbUnchecked And Chk2(6).Value = vbUnchecked And Chk2(7).Value = vbUnchecked And Chk2(8).Value = vbUnchecked And Chk2(9).Value = vbUnchecked And Chk2(10).Value = vbUnchecked And Chk3.Value = vbUnchecked Then
              MsgBox "案件性質最少選一個！", vbInformation, "錯誤！"
              Exit Sub
          End If
          If Chk2(10).Value = vbChecked Then
              If Trim(txt1(10)) = "" Then
                  MsgBox "請輸入其他案件性質！", vbInformation, "錯誤！"
                  txt1(10).SetFocus
                  txt1_GotFocus 10
                  Exit Sub
              End If
          End If
          
          '2011/10/18 ADD BY SONIA 檢查四縣市地址(指定聯絡人及地址因地址欄位置不明故不檢查)
          'Modified by Lydia 2017/03/28 +Trim
          If Trim(txt1(3)) <> "" And (txt1(4) = "" Or InStr(txt1(4), "台灣") > 0 Or InStr(txt1(4), "臺灣") > 0 Or InStr(txt1(4), "中華民國") > 0) Then
            If CheckTaiwanAddr(txt1(3), "000", "發明人地址") = False Then
               txt1(3).SetFocus
               txt1_GotFocus (3)
               Exit Sub
            End If
          End If
          'Modified by Lydia 2017/03/28 +Trim
          If Trim(txt1(7)) <> "" And (txt1(8) = "" Or InStr(txt1(8), "台灣") > 0 Or InStr(txt1(8), "臺灣") > 0 Or InStr(txt1(8), "中華民國") > 0) Then
            If CheckTaiwanAddr(txt1(7), "000", "申請人地址") = False Then
               txt1(7).SetFocus
               txt1_GotFocus (7)
               Exit Sub
            End If
          End If
          'Modified by Lydia 2017/03/28 +Trim
          If Trim(txt1(16)) <> "" Then
            If CheckTaiwanAddr(txt1(16), "000", "甲方委任人地址") = False Then
               txt1(16).SetFocus
               txt1_GotFocus (16)
               Exit Sub
            End If
          End If
          '2011/10/18 END

          'Add by Amy 2016/08/19 +受任人不可為空
          'Modified by Lydia 2017/03/28 +Trim
          If Trim(Combo2) = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
          End If
          
          'Added by Lydia 2019/04/16  開放部分欄位空白,統一彈訊息
          If strMsg <> "" Then
              If MsgBox("下列欄位空白，是否繼續列印？" & Replace(strMsg, "、", vbCrLf), vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                  Exit Sub
              End If
          End If
          'end 2019/04/16
          
           '2009/11/13 MODIFY BY SONIA 杜副總提出
      '    If txt1(20) = "" Or txt1(21) = "" Or txt1(22) = "" Then
      '        MsgBox "日期需要正確！", vbInformation, "錯誤！"
      '        txt1(20).SetFocus
      '        txt1_GotFocus 20
      '        Exit Sub
      '    End If
            'Modified by Lydia 2017/03/28 +Trim
            If Trim(txt1(20)) = "" Or Trim(txt1(21)) = "" Or Trim(txt1(22)) = "" Then
               If MsgBox("契約書日期不完整，是否確定？", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
                 txt1(20).SetFocus
                 txt1_GotFocus 20
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
          
          m_strCustCode = "" 'Added by Morgan 2012/9/11
          'Add by Amy 2016/08/19 畫面與客戶檔收據公司別不同更新客戶檔
          strNowCmp = GetComp(Combo2)
          If Combo2.Tag <> strNowCmp Then
             Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          End If
          'end 2016/08/19
          Screen.MousePointer = vbDefault
          Call RunEndProc(True)  'Added by Lydia 2022/01/20 刪除暫存檔
          'Modified by Lydia 2022/01/22 判斷是否有列印
          'ShowPrintOk 'Added by Lydia 2017/04/11
          If m_TempPDF <> "" Then ShowPrintOk
      Case 1
          frm210114.Show
          Unload Me
      Case 2
          For Each tb In txt1
              tb.Text = Empty
          Next
          For Each op In opt1
              op.Value = False
          Next
          For Each ck In Chk1
              ck.Value = vbUnchecked
          Next
          For Each ck In Chk2
              ck.Value = vbUnchecked
          Next
          'For Each ck In Chk3
          '    ck.Value = vbUnchecked
          'Next
          Chk3.Value = vbUnchecked
          
      Case 3
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowSave
          If cd1.FileName <> "" Then
              AllObj(0) = "案件委任契約書-P"
              For iCount = 1 To 23
                  AllObj(iCount) = txt1(iCount - 1).Text
              Next iCount
              AllObj(24) = Chk1(0).Value
              AllObj(25) = Chk1(1).Value
              AllObj(26) = Chk1(2).Value
              AllObj(27) = Chk2(0).Value
              AllObj(28) = Chk2(1).Value
              AllObj(29) = Chk2(2).Value
              AllObj(30) = Chk2(3).Value
              AllObj(31) = Chk2(4).Value
              AllObj(32) = Chk2(5).Value
              AllObj(33) = Chk2(6).Value
              AllObj(34) = Chk2(7).Value
              AllObj(35) = Chk2(8).Value
              AllObj(36) = Chk2(9).Value
              AllObj(37) = Chk2(10).Value
              AllObj(38) = Chk3.Value
              AllObj(39) = IIf(opt1(0).Value = True, "0", "1")
              AllObj(40) = IIf(opt1(1).Value = True, "0", "1")
              'Add By Sindy 2010/3/17
              AllObj(41) = Chk2(11).Value
              AllObj(42) = Chk2(12).Value
              '2010/3/17 End
              AllObj(43) = Combo2.Text 'Add By Sindy 2010/9/9
              strBuffer = Join(AllObj, Chr(30))
              strBuffer = StrEncrypt(strBuffer)
              fN = FreeFile
              Open cd1.FileName For Output As fN
              Print #fN, strBuffer
              Close #fN
          End If
          'Add by Amy 2016/08/19 畫面與客戶檔收據公司別不同更新客戶檔
          strNowCmp = GetComp(Combo2)
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
              If AllObjV(0) = "案件委任契約書-P" Then
                  cmdOK_Click 2
                  For iCount = 1 To 23
                       txt1(iCount - 1).Text = AllObjV(iCount)
                  Next iCount
                  Chk1(0).Value = AllObjV(24)
                  Chk1(1).Value = AllObjV(25)
                  Chk1(2).Value = AllObjV(26)
                  Chk2(0).Value = AllObjV(27)
                  Chk2(1).Value = AllObjV(28)
                  Chk2(2).Value = AllObjV(29)
                  Chk2(3).Value = AllObjV(30)
                  Chk2(4).Value = AllObjV(31)
                  Chk2(5).Value = AllObjV(32)
                  Chk2(6).Value = AllObjV(33)
                  Chk2(7).Value = AllObjV(34)
                  Chk2(8).Value = AllObjV(35)
                  Chk2(9).Value = AllObjV(36)
                  Chk2(10).Value = AllObjV(37)
                  Chk3.Value = AllObjV(38)
                  opt1(0).Value = IIf(Val(AllObjV(39)) = 0, True, False)
                  opt1(1).Value = IIf(Val(AllObjV(40)) = 0, True, False)
                  'Add By Sindy 2010/3/17
                  Chk2(11).Value = AllObjV(41)
                  Chk2(12).Value = AllObjV(42)
                  '2010/3/17 End
                  'Modify by Amy 2016/08/19 避免空值會Error
                  If AllObjV(43) = MsgText(601) Then
                    Combo2.ListIndex = 0
                  Else
                    Combo2.Text = AllObjV(43) 'Add By Sindy 2010/9/9
                  End If
                  'end 2016/08/19
                  
                  'Add By Sindy 2011/1/21 檢查地址欄
                  '申請人地址
                  If txt1(5).Text <> "" And txt1(7).Text <> "" Then
                     If CheckCustomerAddr(1, Trim(txt1(5).Text), Trim(txt1(7).Text), "申請人", True) = False Then
                        txt1(7).SetFocus
                     End If
                  End If
                  '委任人地址
                  If txt1(13).Text <> "" And txt1(16).Text <> "" Then
                     If CheckCustomerAddr(1, Trim(txt1(13).Text), Trim(txt1(16).Text), "委任人", True) = False Then
                        txt1(16).SetFocus
                     End If
                  End If
                  '2011/1/21 End
                  'Add by Amy 2016/08/19 讀取收據公司別
                  cmdFind_Click
              Else
                  MsgBox "錯誤格式，此份內容並非 內專 格式！", vbExclamation
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
          Call RunEndProc(True)  'Added by Lydia 2022/01/20 刪除暫存檔
          'Modified by Lydia 2022/01/22 判斷是否有列印
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
    
   'Added by Lydia 2020/03/25 設定公司別下拉選項
   Call PUB_SetCboTofrm210114(Me.Name, Me.Combo2, strCompSeal)
   
   'Modify by Amy 2016/08/19
   'Combo2.Text = Combo2.List(0)
   Combo2.ListIndex = 0
   
   'Added by Lydia 2022/01/14 保留
'   If Pub_StrUserSt03 = "M51" Then
'       cmdTest.Visible = True
'   End If
   
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
   
   Call RunEndProc(False)  'Added by Lydia 2022/01/20 刪除暫存檔
   
   Set frm210114_1 = Nothing
End Sub

'Memo by Lydia 2021/04/19 因為第五條內文增加，所以版面調整
Sub StrMenu(Optional ByVal bolSpace As Boolean = False)
Dim iY As Integer
Dim tmpI As Integer
Dim iStr(1 To 46) As String
Dim tBoxTop As Integer
Dim tObj As New StdPicture
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String

   iStr(1) = "專利案件委任契約書"
   iStr(2) = "委任人(甲方)茲委任受任人(乙方)辦理國內專利案件，經雙方同意條件如下："
   iStr(3) = "第一條"
   iStr(4) = "　　一、發明創作名稱：" & StrToStr(txt1(0) & String(60, " "), 30)
   iStr(5) = "　　二、發明或創作人：" & StrToStr(txt1(1) & String(38, " "), 19) & "　I.D.NO：" & StrToStr(txt1(2) & String(10, " "), 5)
   iStr(6) = "　　　　地址：" & StrToStr(txt1(3) & String(50, " "), 25) & "　國籍：" & StrToStr(txt1(4) & String(8, " "), 4)
   iStr(7) = "　　三、申   請   人：" & StrToStr(txt1(5) & String(38, " "), 19) & "　I.D.NO：" & StrToStr(txt1(6) & String(10, " "), 5)
   iStr(8) = "　　　　地址：" & StrToStr(txt1(7) & String(50, " "), 25) & "　國籍：" & StrToStr(txt1(8) & String(8, " "), 4)
   iStr(9) = "　　四、指定聯絡人及地址：" & StrToStr(txt1(9) & String(54, " "), 27)
   iStr(10) = "第二條　委辦範圍"
   iStr(11) = "　　　　申請種類：" & IIf(Chk1(0).Value = 1, "■", "□") & "發明  　　" & IIf(Chk1(1).Value = 1, "■", "□") & "新型  " & IIf(Chk1(2).Value = 1, "■", "□") & "設計　　　"
   iStr(12) = "　　　　性    質：" & IIf(Chk2(0).Value = 1, "■", "□") & "申請  　　" & IIf(Chk2(1).Value = 1, "■", "□") & "再審  " & IIf(Chk2(2).Value = 1, "■", "□") & "初審申復  " & IIf(Chk2(3).Value = 1, "■", "□") & "再審申復  " & IIf(Chk2(4).Value = 1, "■", "□") & "訴願" & IIf(Chk3.Value = 1, "■", "□") & "新型技術報告　"
   iStr(13) = "　　　　　　　　　" & IIf(Chk2(5).Value = 1, "■", "□") & "行政訴訟  " & IIf(Chk2(6).Value = 1, "■", "□") & "舉發　" & IIf(Chk2(7).Value = 1, "■", "□") & "答辯　　  " & IIf(Chk2(8).Value = 1, "■", "□") & "領證　　　" & IIf(Chk2(9).Value = 1, "■", "□") & "繳年費"
   iStr(14) = "　　　　　　　　　" & IIf(Chk2(11).Value = 1, "■", "□") & "實體審查  " & IIf(Chk2(10).Value = 1, "■", "□") & "其他：" & txt1(10).Text
   iStr(15) = "　　　　　　　　　" & IIf(Chk2(12).Value = 1, "■", "□") & "本案所涉之技術內容並非在中國大陸境內完成的發明或者實用新型"
   iStr(16) = "　　　　乙方根據甲方所提供資料，依前項約定之範圍，代撰必要書件，向本程序主管機關"
   iStr(17) = "　　　　提出，並代為收受有關文件。"
   iStr(18) = "第三條　委辦費用"  '第三條(起)
   '空白委任書要保留費用
   If Val(Trim(txt1(11))) = 0 And bolSpace = False Then
       iStr(19) = ""
       iStr(20) = ""
   Else
       If Val(Trim(txt1(11))) = 0 Then
          strSpaceAmt = "　　　　　　　　　　　　元整"
       Else
          strSpaceAmt = ChangeNumber(txt1(11))
       End If
       iStr(19) = StrToStr("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 44)
       iStr(20) = "　　　　" & Replace("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", StrToStr("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 44), "")
       '金額直接併入說明
       strExc(3) = StrToStr("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 44)
       strExc(4) = "　　　　" & Replace("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", StrToStr("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 44), "")
   End If
   '空白委任書要保留費用
   If Val(Trim(txt1(12))) = 0 And bolSpace = False Then
       iStr(21) = ""
       iStr(22) = ""
   Else
       If Val(Trim(txt1(12))) = 0 Then
          strSpaceAmt = "　　　　　　　　　　　　元整"
       Else
          strSpaceAmt = ChangeNumber(txt1(12))
       End If
       iStr(21) = StrToStr("　　二、後酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本程序終結時，由甲方一次付清。", 44)
       strExc(5) = StrToStr("　　二、後酬金新台幣　" & strSpaceAmt & "，於本程序終結時，由甲方一次付清。", 44)
   End If
   iStr(22) = "第四條　乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，並不得發生足以影響甲方"
   iStr(23) = "　　　　權益之疏誤，否則應對甲方負損害賠償責任。但以不超過第三條所載前酬金金額之三倍為限。"
   'Memo by Lydia 2021/04/19 改第五條內容
   iStr(24) = "第五條　甲方應確保所交付予乙方之資料及本契約書所載內容(包括發明人或創作人、申請人等資訊)"
   iStr(25) = "　　　　均無虛偽情事，且甲方確實得到與委辦案件相關共同發明人及第三人之同意，有權委託乙方"
   iStr(26) = "　　　　辦理案件，如因不實致生損害或法律責任時，概由甲方負責，與乙方無關。"
   iStr(27) = "第六條　乙方於辦理過程中，應隨時將辦理經過如申請日期、案號及其他重要函件，儘速通知或交付"
   iStr(28) = "　　　　甲方。但甲方於簽約後變更連絡處所，未即時通知乙方，因而連絡不及致延誤時限者，乙方"
   iStr(29) = "　　　　不負責任。" '
   iStr(30) = "第七條　凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時限者，乙方不負責。經乙方"
   iStr(31) = "　　　　通知甲方繳費而未依限繳納者，亦同。"
   iStr(32) = "第八條　甲方如逕自撤回所委辦程序，或未經乙方同意終止契約時，所約定之費用，仍應全數給付。"
   iStr(33) = "第九條　本契約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，但有增刪修改時，需甲乙雙方"
   iStr(34) = "　　　　於更動處蓋章始生效力，並由雙方各執乙份為憑。"
   iStr(35) = "          "   '第九條(止)
   iStr(36) = "　　　　　　甲方：委任人：" & StrToStr(txt1(13) & String(54, " "), 27)
   iStr(37) = "　　　　　　　　　ID NO.：" & StrToStr(txt1(14) & String(20, " "), 10) & "代表人：" & StrToStr(txt1(15) & String(26, " "), 13)
   iStr(38) = "　　　　　　      地  址：" & StrToStr(txt1(16) & String(54, " "), 27)
   iStr(39) = "　　　　　　　　　電  話：" & StrToStr(txt1(17) & String(20, " "), 10) & "傳  真：" & StrToStr(txt1(18) & String(26, " "), 13)
   iStr(40) = "　　　　　　乙方：受任人：" & Combo2.Text '"台一國際專利法律事務所"
   iStr(41) = "　　　　　　　　　經手人：" & StrToStr(txt1(19) & String(54, " "), 27)
   iStr(42) = "　　　　　　　　　地　址：" & PUB_SetAddrTofrm210114(Combo2.Text) '改用模組控制
   iStr(43) = "　　　　　　　　　電  話：(02)25061023(總機)   FAX:(02)25011666"
   iStr(44) = "　　　　　　　　　網  址：www.taie.com.tw"
   iStr(45) = "　　　　　　　　　E-mail：ipdept@taie.com.tw"
   iStr(46) = "  中    華    民    國 " & String((10 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & txt1(20) & String((10 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(21)), vbFromUnicode))) / 2, " ") & txt1(21) & String((10 - LenB(StrConv((txt1(21)), vbFromUnicode))) / 2, " ") & "月" & String((10 - LenB(StrConv((txt1(22)), vbFromUnicode))) / 2, " ") & txt1(22) & String((10 - LenB(StrConv((txt1(22)), vbFromUnicode))) / 2, " ") & "日"
   
   '有用印就記錄列印內容
   If iPrintC = 1 And bolAddSeal = True Then
           strDetail = ""
           For intI = 1 To UBound(iStr)
              If Trim(iStr(intI)) <> "" Then
                 If intI >= 19 And intI <= 21 Then
                    Select Case intI
                        Case 19: strDetail = strDetail & vbCrLf & RTrim(strExc(3))
                        Case 20: strDetail = strDetail & vbCrLf & RTrim(strExc(4))
                        Case 21: strDetail = strDetail & vbCrLf & RTrim(strExc(5))
                    End Select
                 Else
                    If intI <= 18 Or (intI >= 36 And intI <= 41) Or intI = 46 Then
                        If intI = 40 Then
                          strDetail = strDetail & vbCrLf & vbCrLf & IIf(opt1(0).Value = True, "會稿", "不會稿") & vbCrLf
                        End If
                       strDetail = strDetail & vbCrLf & RTrim(iStr(intI))
                    End If
                 End If
              End If
           Next
   End If
   

   iY = 0
   Printer.PaperSize = 9
   Printer.Orientation = 1
  
   Printer.FontName = "標楷體"
   Printer.FontSize = 20
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(iStr(1))) / 2
   iY = iY + Printer.TextHeight(iStr(1))
   Printer.CurrentY = iY
   iY = iY + ((Printer.TextHeight(iStr(1)) / 3) * 4)
   Printer.Print iStr(1)
   Printer.FontSize = 12
   '同步用印
   If bolAddSeal = True Then
      '列印座置抓乙方資料的起始
      'X軸
      strExc(1) = 1000 + (Printer.TextWidth("　") * 30)
      'Y軸
      intI = 0
      If bolSpace = True Then
         If Val(Trim(txt1(12))) = 0 Then
            intI = 4
         Else
            intI = 3
         End If
      Else
         '顯示費用，資料列數不同
         If Val(Trim(txt1(11))) > 0 Then intI = intI + 2
         If Val(Trim(txt1(12))) > 0 Then intI = intI + 1
      End If
      strExc(2) = iY + ((Printer.TextHeight("　") / 3) * 4) * (34 + intI) + 50
      '圖片尺寸
      strExc(3) = 1600 'width
      strExc(4) = 1600 'height
      
      '已記錄公司名稱|用印編號
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
        If InStr(Combo2.Text, "專利法律") > 0 Then
          If PUB_ReadDB2File(strSealFile, 51) Then
             Set tObj = pvGetStdPicture(strSealFile)
             Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
          End If
        End If
        If InStr(Combo2.Text, "專利商標") > 0 Then
          If PUB_ReadDB2File(strSealFile, 52) Then
             Set tObj = pvGetStdPicture(strSealFile)
             Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
          End If
        End If
      End If
   End If
   
   For tmpI = 2 To UBound(iStr) - 1
        '第三條到第九條改為11號字
        If tmpI = 18 Then '第三條(起)
             Printer.FontSize = 11
        ElseIf tmpI = 35 Then  '第九條(止)
             Printer.FontSize = 12
        End If
       If Trim(iStr(tmpI)) <> "" Or tmpI = 35 Then
           If tmpI = 37 Then
               tBoxTop = iY
           End If
           Printer.CurrentX = 1000
           Printer.CurrentY = iY
           Printer.Print iStr(tmpI)
           If tmpI = 19 Then
               Printer.FontBold = True
               Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 11) - 50
               Printer.CurrentY = iY
               '判斷空白列印
                If bolSpace = True And Val(Trim(txt1(11))) = 0 Then
                   Printer.Print String(12, "　") & "元整"
               Else
                   Printer.Print ChangeNumber(txt1(11))
               End If
               Printer.FontBold = False
           End If
           If tmpI = 21 Then
               Printer.FontBold = True
               Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 11) - 50
               Printer.CurrentY = iY
               '判斷空白列印
               If bolSpace = True And Val(Trim(txt1(12))) = 0 Then
                   Printer.Print String(12, "　") & "元整"
               Else
                   Printer.Print ChangeNumber(txt1(12))
               End If
               Printer.FontBold = False
           End If
           iY = iY + ((Printer.TextHeight(iStr(tmpI)) / 3) * 4)
           '畫線
           Select Case tmpI
           Case 4
                Printer.Line (1000 + (Printer.TextWidth("　") * 11), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 5, 7
                Printer.Line (1000 + (Printer.TextWidth("　") * 11), iY - 50)-(1000 + (Printer.TextWidth("　") * 31), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 35), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 6, 8
                Printer.Line (1000 + (Printer.TextWidth("　") * 7), iY - 50)-(1000 + (Printer.TextWidth("　") * 33), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 36), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 9, 36, 38, 40, 41
                Printer.Line (1000 + (Printer.TextWidth("　") * 13), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 14
                Printer.Line (1000 + (Printer.TextWidth("　") * 19), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 19, 21
           Case 37, 39
                Printer.Line (1000 + (Printer.TextWidth("　") * 13), iY - 50)-(1000 + (Printer.TextWidth("　") * 23), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 27), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case Else
           End Select
       End If
   Next tmpI
   '畫格子
   Printer.Line (1000, tBoxTop)-(1000 + (Printer.TextWidth("　") * 5), tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 8)), , B
   Printer.Line (1000, tBoxTop)-(1000 + (Printer.TextWidth("　") * 2.5), tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 8)), , B
   Printer.Line (1000, tBoxTop)-(1000 + (Printer.TextWidth("　") * 5), tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 4)), , B
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 0.5)
   Printer.Print "會"
   If opt1(0).Value = True Then
       Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 3.25)
       Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 1.5)
       Printer.Print "Ｖ"
   End If
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 2.5)
   Printer.Print "稿"
   
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 4.5)
   Printer.Print "不"
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 5.5)
   Printer.Print "會"
   If opt1(1).Value = True Then
       Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 3.25)
       Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 5.5)
       Printer.Print "Ｖ"
   End If
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 6.5)
   Printer.Print "稿"
   Printer.FontSize = 16
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(iStr(UBound(iStr)))) / 2
   Printer.CurrentY = iY + 200
   Printer.Print iStr(UBound(iStr))

   Printer.EndDoc

   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox Err.Description
      Resume Next
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

'Modified by Lydia 2022/01/14 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
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
   If Index = 11 Or Index = 12 Then
       If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 46 Then
           KeyAscii = 0
       End If
   End If
   '2009/11/13 ADD BY SONIA
   If Index = 20 Or Index = 21 Or Index = 22 Then
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
   StrSQLa = "Select * From Customer,nation,potcustcont Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "' and CU10=na01(+) and pcc01(+)=CU01 And pcc02(+)=CU127 "
   'end 2021/5/5
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      SetCustTxt = True
      '申請人中文
      Me.txt1(5).Text = "" & rsA("CU04").Value
      Me.txt1(13).Text = "" & rsA("CU04").Value
       
      'Add by Lydia 2014/9/22 複選人員
      part1 = 1
      ppart1 = GetSubStringCount(partCust) '取得字串以逗點分隔的Sub字串總數
         Do While part1 <= ppart1
            strCUCode = Mid(partCust, (part1 * 10) + 1, 9) '從第幾組代號開始，截取下一組代號
            'Modified by Morgan 2021/5/5
            'StrSQLa = " Select CU04 From Customer,nation,potcustcont " & _
                      " Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "' and CU10=na01(+) and pcc01(+)=substr(CU08, 1, 8) And pcc02(+)=substr(CU08, 9, 1) "
            StrSQLa = " Select CU04 From Customer Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "'"
            'end 2021/5/5
            If rsB.State <> adStateClosed Then rsB.Close
            Set rsB = Nothing
            rsB.CursorLocation = adUseClient
            rsB.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsB.RecordCount > 0 Then
             rsB.MoveFirst
             Me.txt1(5).Text = LTrim(RTrim(Me.txt1(5).Text)) + "、" & rsB("CU04").Value
            End If
            
            part1 = part1 + 1
            
            If part1 = ppart1 Then

               If CheckLengthIsOK(txt1(5).Text, txt1(5).MaxLength) = False Then
                    Me.txt1(5).SetFocus
               End If
            End If
        Loop
     

      'ID No.
      Me.txt1(6).Text = "" & rsA("CU11").Value
      Me.txt1(14).Text = "" & rsA("CU11").Value
      '申請地址
      Me.txt1(7).Text = "" & rsA("CU23").Value
      Me.txt1(16).Text = "" & rsA("CU23").Value
      '國籍
      'Modified by Lydia 2019/08/05 台灣地區改顯示為中華民國
      'Me.txt1(8).Text = "" & rsA("NA03").Value
      Me.txt1(8).Text = IIf(Val("" & rsA("NA01")) < 10, "中華民國", "" & rsA("NA03").Value)
      
      '聯絡人地址
      'Modified by Morgan 2021/5/5
      'If "" & rsA("CU08").Value <> "" Then
      If "" & rsA("pcc22").Value <> "" Then
      'end 2021/5/5
         Me.txt1(9).Text = "" & rsA("pcc22").Value
      Else
         Me.txt1(9).Text = "" & rsA("CU31").Value
      End If
      '電話1
      Me.txt1(17).Text = "" & rsA("CU16").Value
      '傳真1
      Me.txt1(18).Text = "" & rsA("CU18").Value
      '代表人1中文
      Me.txt1(15).Text = "" & rsA("CU07").Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Add by Lydia 2014/9/22 發明人
Private Sub cmdFind2_Click()
   If Me.txt1(1).Text = "" Then
      MsgBox "請輸入發明人中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txt1(1).SetFocus
      Exit Sub
   End If
   
   frm090801_1.m_Type = 3
   If ChkDou.Value = 1 Then
     frm090801_1.m_DouChk = True '可複選
   Else
     frm090801_1.m_DouChk = False
   End If

   Set frm090801_1.m_frm0908A = Me
   
   frm090801_1.m_strCustChnName = Me.txt1(1).Text
   frm090801_1.lblName.Caption = Me.txt1(1).Text
 
   m_blnOneRec = False
   m_strCustCode = ""
'   txt1(1).Tag = ""
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
   
   If CheckLengthIsOK(txt1(1).Text, txt1(1).MaxLength) = False Then
      Me.txt1(1).SetFocus
   End If
   
End Sub

'Add by Amy 2016/08/19
Public Function GetComp(ByVal stComp As String) As String
    If stComp = MsgText(601) Then Exit Function
    
    GetComp = ""
    Select Case stComp
        Case "台一國際專利商標事務所"
            GetComp = "1"
        'Moified by Lydia 2022/04/15 + 台一國際智慧財產事務所
        Case "台一國際專利法律事務所", "台一國際智慧財產事務所"
            GetComp = "2"
        Case "台一智權股份有限公司"
            GetComp = "J"
    End Select
End Function

Private Sub UpdReceiptCmp(ByVal stNowCustNo As String, ByVal stNowCmp As String)
    Dim strUpd As String
    
    Exit Sub 'Added by Lydia 2022/08/30 受任人下拉預設只剩下台一國際智慧財產事務所，所以不必再更新客戶檔的了。
    
    'Add by Amy 2016/12/30 +同業務區或為MCTF同組人員才可回寫收據公司別
    If ChkSameCuArea(stNowCustNo, strUserNum) = False Then Exit Sub
    
    'Modified by Lydia 2019/04/12 拿掉UpdateID,Date,Time(CU84,CU85,CU86)
    'strUpd = "Update Customer Set CU84='" & strUserNum & "',CU85=to_number(to_char(sysdate,'YYYYMMDD')),CU86=to_number(to_char(sysdate,'HH24MI')),CU160='" & stNowCmp & "' " & _
                            "Where CU01='" & Left(stNowCustNo, 8) & "' And CU02='" & Mid(stNowCustNo, 9, 1) & "' "
    strUpd = "Update Customer Set CU160='" & stNowCmp & "' " & _
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
    
    'Modified by Lydia 2022/01/19 先產生Word檔，後轉成PDF檔逐一列印
'    For iCount = 1 To Val(txtPCnt)
'        iPrintC = iCount
'        'Modified by Lydia 2017/06/06 改用App.Title變更印表機列印文件名稱(執行exe檔有效,VB跑無效)
'        'strFileName = strUserNum & "_P_" & IIf(bSpace = False, Mid(Trim(txt1(5)), 1, 4), "空白") & iCount & ".pdf"
'        'If Dir(App.path & "\" & strFileName) <> "" Then
'        '   Kill App.path & "\" & strFileName
'        'End If
'        ''轉PDF
'        'frmPDF.Show
'        'frmPDF.StartProcess App.path, strFileName
'        'Call StrMenu(bSpace)
'        'frmPDF.EndtProcess
'        'Unload frmPDF
'        strFileName = strUserNum & "_P_" & IIf(bSpace = False, Mid(Trim(txt1(5)), 1, 4), "空白") & iCount
'        App.Title = strFileName
'        Call StrMenu(bSpace)
'        'end 2017/06/07
'
'        'Added by Lydia 2017/05/16 用印記錄移到pdf建立
'        If iCount = 1 And strDetail <> "" Then
'           'If Dir(App.path & "\" & strFileName) <> "" Then 'Remove by Lydia 2020/03/16 因為不存檔案所以取消檔案檢查(自2017/06/08~2020/03/16無用印記錄)
'              If PUB_AddRecSeal("1", txtPCnt.Text, IIf(bSpace = True, "Y", ""), strDetail, Combo2.Text) Then
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
            strFileName = strUserNum & "_P_" & IIf(bSpace = False, m_TempFN, "空白") & iCount
            PUB_PrintPDF App.path & "\" & strUserNum & "\" & m_TempPDF, Combo1.Text
            App.Title = strFileName
        Next iCount
    End If
'--------------先產生Word檔，後轉成PDF檔逐一列印
    App.Title = strOldName 'Added by Lydia 2017/06/07
    
End Sub

'Modified by Lydia 2017/03/28
'Sub StrMenu()
'Memo by Lydia 2021/04/19 保留舊版面；因為第五條內文增加，所以版面調整
Sub StrMenu_P(Optional ByVal bolSpace As Boolean = False)
Dim iY As Integer
Dim tmpI As Integer
Dim iStr(1 To 48) As String
Dim tBoxTop As Integer
'Added by Lydia 2017/03/28
Dim tObj As New StdPicture
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String
'end 2017/03/28

   iStr(1) = "專利案件委任契約書"
   iStr(2) = "委任人(甲方)茲委任受任人(乙方)辦理國內專利案件，經雙方同意條件如下："
   iStr(3) = "第一條"
   iStr(4) = "　　一、發明創作名稱：" & StrToStr(txt1(0) & String(60, " "), 30)
   iStr(5) = "　　二、發明或創作人：" & StrToStr(txt1(1) & String(38, " "), 19) & "　I.D.NO：" & StrToStr(txt1(2) & String(10, " "), 5)
   iStr(6) = "　　　　地址：" & StrToStr(txt1(3) & String(50, " "), 25) & "　國籍：" & StrToStr(txt1(4) & String(8, " "), 4)
   iStr(7) = "　　三、申   請   人：" & StrToStr(txt1(5) & String(38, " "), 19) & "　I.D.NO：" & StrToStr(txt1(6) & String(10, " "), 5)
   iStr(8) = "　　　　地址：" & StrToStr(txt1(7) & String(50, " "), 25) & "　國籍：" & StrToStr(txt1(8) & String(8, " "), 4)
   iStr(9) = "　　四、指定聯絡人及地址：" & StrToStr(txt1(9) & String(54, " "), 27)
   iStr(10) = "第二條　委辦範圍"
   '2013/10/24 modify by sonia 新型技術報告 自申請種類移至性質
   iStr(11) = "　　　　申請種類：" & IIf(Chk1(0).Value = 1, "■", "□") & "發明  　　" & IIf(Chk1(1).Value = 1, "■", "□") & "新型  " & IIf(Chk1(2).Value = 1, "■", "□") & "設計　　　"
   iStr(12) = "　　　　性    質：" & IIf(Chk2(0).Value = 1, "■", "□") & "申請  　　" & IIf(Chk2(1).Value = 1, "■", "□") & "再審  " & IIf(Chk2(2).Value = 1, "■", "□") & "初審申復  " & IIf(Chk2(3).Value = 1, "■", "□") & "再審申復  " & IIf(Chk2(4).Value = 1, "■", "□") & "訴願" & IIf(Chk3.Value = 1, "■", "□") & "新型技術報告　"
   iStr(13) = "　　　　　　　　　" & IIf(Chk2(5).Value = 1, "■", "□") & "行政訴訟  " & IIf(Chk2(6).Value = 1, "■", "□") & "舉發　" & IIf(Chk2(7).Value = 1, "■", "□") & "答辯　　  " & IIf(Chk2(8).Value = 1, "■", "□") & "領證　　　" & IIf(Chk2(9).Value = 1, "■", "□") & "繳年費"
   'edit by nickc 2008/05/09 加入實體審查
   'iStr(14) = "　　　　　　　　　" & IIf(Chk2(10).Value = 1, "■", "□") & "其他：" & txt1(10).Text
   iStr(14) = "　　　　　　　　　" & IIf(Chk2(11).Value = 1, "■", "□") & "實體審查  " & IIf(Chk2(10).Value = 1, "■", "□") & "其他：" & txt1(10).Text
   'Add By Sindy 2009/10/19
   iStr(15) = "　　　　　　　　　" & IIf(Chk2(12).Value = 1, "■", "□") & "本案所涉之技術內容並非在中國大陸境內完成的發明或者實用新型"
   '2009/10/19 End
   iStr(16) = "　　　　乙方根據甲方所提供資料，依前項約定之範圍，代撰必要書件，向本程序主管機關"
   iStr(17) = "　　　　提出，並代為收受有關文件。"
   iStr(18) = "第三條　委辦費用"
   'Modify By Sindy 2009/10/22
   'If Trim(txt1(11)) = "" Then
   'Modified by Lydia 2017/03/28 空白委任書要保留費用
   'If Val(Trim(txt1(11))) = 0 Then
   If Val(Trim(txt1(11))) = 0 And bolSpace = False Then
   '2009/10/22 End
       iStr(19) = ""
       iStr(20) = ""
   Else
       'Added by Lydia 2017/03/28
       If Val(Trim(txt1(11))) = 0 Then
          strSpaceAmt = "　　　　　　　　　　　　元整"
       Else
          strSpaceAmt = ChangeNumber(txt1(11))
       End If
       'Modified by Lydia 2017/03/28 ChangeNumber(txt1(11)) => ChangeNumber(strSpaceAmt)
       iStr(19) = StrToStr("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40)
       iStr(20) = "　　　　" & Replace("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", StrToStr("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40), "")
       'Added by Lydia 2017/03/28 金額直接併入說明
       strExc(3) = StrToStr("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40)
       strExc(4) = "　　　　" & Replace("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", StrToStr("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40), "")
   End If
   'Modify By Sindy 2009/10/22
   'If Trim(txt1(12)) = "" Then
   'Modified by Lydia 2017/03/28 空白委任書要保留費用
   'If Val(Trim(txt1(12))) = 0 Then
   If Val(Trim(txt1(12))) = 0 And bolSpace = False Then
   '2009/10/22 End
       iStr(21) = ""
       iStr(22) = ""
       'Added by Morgan 2022/1/18
       iStr(19) = Replace(iStr(19), "一、", "　　")
       iStr(20) = Replace(iStr(20), "一、", "　　")
       strExc(3) = Replace(strExc(3), "一、", "　　")
       strExc(4) = Replace(strExc(4), "一、", "　　")
       'end 2022/1/18
   Else
       'Added by Lydia 2017/03/28
       If Val(Trim(txt1(12))) = 0 Then
          strSpaceAmt = "　　　　　　　　　　　　元整"
       Else
          strSpaceAmt = ChangeNumber(txt1(12))
       End If
       'Modified by Lydia 2017/03/28 ChangeNumber(txt1(12)) => ChangeNumber(strSpaceAmt)
       iStr(21) = StrToStr("　　二、後酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本程序終結時，由甲方一次付清。", 40)
       iStr(22) = "　　　　" & Replace("　　二、後酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本程序終結時，由甲方一次付清。", StrToStr("　　二、後酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本程序終結時，由甲方一次付清。", 40), "")
       'Added by Lydia 2017/03/28 金額直接併入說明
       strExc(5) = StrToStr("　　二、後酬金新台幣　" & strSpaceAmt & "，於本程序終結時，由甲方一次付清。", 40)
       strExc(6) = "　　　　" & Replace("　　二、後酬金新台幣　" & strSpaceAmt & "，於本程序終結時，由甲方一次付清。", StrToStr("　　二、後酬金新台幣　" & strSpaceAmt & "，於本程序終結時，由甲方一次付清。", 40), "")
       
       
       'Added by Morgan 2022/1/18
       If iStr(19) = "" Then
         iStr(21) = Replace(iStr(21), "二、", "　　")
         iStr(22) = Replace(iStr(22), "二、", "　　")
         strExc(5) = Replace(strExc(5), "二、", "　　")
         strExc(6) = Replace(strExc(6), "二、", "　　")
       End If
       'end 2022/1/18
   End If
   iStr(23) = "第四條　乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，並不得發生足以"
   iStr(24) = "　　　　影響甲方權益之疏誤，否則應對甲方負損害賠償責任。但以不超過第三條所載前酬"
   iStr(25) = "　　　　金金額之三倍為限。"
   iStr(26) = "第五條　甲方確保所交付予乙方之資料均無虛偽情事，如因不實致生損害或法律責任時，概"
   iStr(27) = "　　　　由甲方負責，與乙方無關。"
   iStr(28) = "第六條　乙方於辦理過程中，應隨時將辦理經過如申請日期、案號及其他重要函件，儘速通"
   iStr(29) = "　　　　知或交付甲方。但甲方於簽約後變更連絡處所，未即時通知乙方，因而連絡不及致"
   iStr(30) = "　　　　延誤時限者，乙方不負責任。"
   iStr(31) = "第七條　凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時限者，乙方不負責。"
   iStr(32) = "　　　　經乙方通知甲方繳費而未依限繳納者，亦同。"
   iStr(33) = "第八條　甲方如逕自撤回所委辦程序，或未經乙方同意終止契約時，所約定之費用，仍應全"
   iStr(34) = "　　　　數給付。"
   iStr(35) = "第九條  本契約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，但有增刪修改時，需"
   iStr(36) = "　　　　甲乙雙方於更動處蓋章始生效力，並由雙方各執乙份為憑。"
   iStr(37) = "          "
   iStr(38) = "　　　　　　甲方：委任人：" & StrToStr(txt1(13) & String(54, " "), 27)
   iStr(39) = "　　　　　　　　　ID NO.：" & StrToStr(txt1(14) & String(20, " "), 10) & "代表人：" & StrToStr(txt1(15) & String(26, " "), 13)
   iStr(40) = "　　　　　　      地  址：" & StrToStr(txt1(16) & String(54, " "), 27)
   iStr(41) = "　　　　　　　　　電  話：" & StrToStr(txt1(17) & String(20, " "), 10) & "傳  真：" & StrToStr(txt1(18) & String(26, " "), 13)
   iStr(42) = "　　　　　　乙方：受任人：" & Combo2.Text '"台一國際專利法律事務所"
   iStr(43) = "　　　　　　　　　經手人：" & StrToStr(txt1(19) & String(54, " "), 27)
   'Modified by Lydia 2020/04/09 改用模組控制
   'iStr(44) = "　　　　　　　　　地　址：台北市長安東路二段一一二號九樓"
   iStr(44) = "　　　　　　　　　地　址：" & PUB_SetAddrTofrm210114(Combo2.Text)
   iStr(45) = "　　　　　　　　　電  話：(02)25061023(總機)   FAX:(02)25011666"
   iStr(46) = "　　　　　　　　　網  址：www.taie.com.tw"
   iStr(47) = "　　　　　　　　　E-mail：ipdept@taie.com.tw"   'modify by sonia 2020/4/8 原為lawoffice
   iStr(48) = "  中    華    民    國 " & String((10 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & txt1(20) & String((10 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(21)), vbFromUnicode))) / 2, " ") & txt1(21) & String((10 - LenB(StrConv((txt1(21)), vbFromUnicode))) / 2, " ") & "月" & String((10 - LenB(StrConv((txt1(22)), vbFromUnicode))) / 2, " ") & txt1(22) & String((10 - LenB(StrConv((txt1(22)), vbFromUnicode))) / 2, " ") & "日"
   
   'Added by Lydia 2017/03/28 有用印就記錄列印內容
   If iPrintC = 1 And bolAddSeal = True Then
           strDetail = ""
           For intI = 1 To UBound(iStr)
              If Trim(iStr(intI)) <> "" Then
                 If intI >= 19 And intI <= 22 Then
                    Select Case intI
                        Case 19: strDetail = strDetail & vbCrLf & RTrim(strExc(3))
                        Case 20: strDetail = strDetail & vbCrLf & RTrim(strExc(4))
                        Case 21: strDetail = strDetail & vbCrLf & RTrim(strExc(5))
                        Case 22: strDetail = strDetail & vbCrLf & RTrim(strExc(6))
                    End Select
                 Else
                    If intI <= 18 Or (intI >= 38 And intI <= 43) Or intI = 48 Then
                        If intI = 38 Then
                          strDetail = strDetail & vbCrLf & vbCrLf & IIf(opt1(0).Value = True, "會稿", "不會稿") & vbCrLf
                        End If
                       strDetail = strDetail & vbCrLf & RTrim(iStr(intI))
                    End If
                 End If
              End If
           Next
        'Modified by Lydia 2017/04/17 空白用印改由勾選項目控制
        'If PUB_AddRecSeal("1", txtPCnt.Text, IIf(ChkSeal.Value = 1, "", "Y"), strDetail, Combo2.Text) Then
        'Remove by Lydia 2017/05/16 用印記錄移到pdf建立
        'If PUB_AddRecSeal("1", txtPCnt.Text, IIf(bolSpace = True, "Y", ""), strDetail, Combo2.Text) Then
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
   Printer.FontSize = 12
   'Added by Lydia 2017/03/28 同步用印
   If bolAddSeal = True Then
      '列印座置抓乙方資料的起始
      'X軸
      strExc(1) = 1000 + (Printer.TextWidth("　") * 30)
      'Y軸
      intI = 0
      If bolSpace = True Then
         If Val(Trim(txt1(12))) = 0 Then
            intI = 4
         Else
            intI = 3
         End If
      Else
         '顯示費用，資料列數不同
         If Val(Trim(txt1(11))) > 0 Then intI = intI + 2
         If Val(Trim(txt1(12))) > 0 Then intI = intI + 1
      End If
      strExc(2) = iY + ((Printer.TextHeight("　") / 3) * 4) * (36 + intI) + 50
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
      End If 'Added by Lydia 2020/03/25
   End If
   'end 2017/03/28
   
   For tmpI = 2 To UBound(iStr) - 1
       'Modify By Sindy 2009/10/22
       'If iStr(tmpI) <> "" Then
       If Trim(iStr(tmpI)) <> "" Or tmpI = 37 Then
       '2009/10/22 End
           If tmpI = 39 Then
               tBoxTop = iY
           End If
           Printer.CurrentX = 1000
           Printer.CurrentY = iY
           Printer.Print iStr(tmpI)
           If tmpI = 19 Then
               Printer.FontBold = True
               Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 11) - 50
               Printer.CurrentY = iY
               'Modified by Lydia 2017/03/28 +判斷空白列印
               'Printer.Print ChangeNumber(txt1(11))
               If bolSpace = True And Val(Trim(txt1(11))) = 0 Then
                   Printer.Print String(12, "　") & "元整"
               Else
                   Printer.Print ChangeNumber(txt1(11))
               End If
               'end 2017/03/28
               Printer.FontBold = False
           End If
           If tmpI = 21 Then
               Printer.FontBold = True
               Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 11) - 50
               Printer.CurrentY = iY
               'Added by Lydia 2017/03/28 +判斷空白列印
               'Printer.Print ChangeNumber(txt1(12))
               If bolSpace = True And Val(Trim(txt1(12))) = 0 Then
                   Printer.Print String(12, "　") & "元整"
               Else
                   Printer.Print ChangeNumber(txt1(12))
               End If
               'end 2017/03/28
               Printer.FontBold = False
           End If
           iY = iY + ((Printer.TextHeight(iStr(tmpI)) / 3) * 4)
           '畫線
           Select Case tmpI
           Case 4
                Printer.Line (1000 + (Printer.TextWidth("　") * 11), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 5, 7
                Printer.Line (1000 + (Printer.TextWidth("　") * 11), iY - 50)-(1000 + (Printer.TextWidth("　") * 31), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 35), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 6, 8
                Printer.Line (1000 + (Printer.TextWidth("　") * 7), iY - 50)-(1000 + (Printer.TextWidth("　") * 33), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 36), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 9, 38, 40, 42, 43
                Printer.Line (1000 + (Printer.TextWidth("　") * 13), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 14
                'edit by nickc 2008/05/09 加入實體審查，往後移
                'Printer.Line (1000 + (Printer.TextWidth("　") * 13), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 19), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 19, 21
           Case 39, 41
                Printer.Line (1000 + (Printer.TextWidth("　") * 13), iY - 50)-(1000 + (Printer.TextWidth("　") * 23), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 27), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case Else
           End Select
       End If
   Next tmpI
   '畫格子
   Printer.Line (1000, tBoxTop)-(1000 + (Printer.TextWidth("　") * 5), tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 8)), , B
   Printer.Line (1000, tBoxTop)-(1000 + (Printer.TextWidth("　") * 2.5), tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 8)), , B
   Printer.Line (1000, tBoxTop)-(1000 + (Printer.TextWidth("　") * 5), tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 4)), , B
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 0.5)
   Printer.Print "會"
   If opt1(0).Value = True Then
       Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 3.25)
       Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 1.5)
       Printer.Print "Ｖ"
   End If
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 2.5)
   Printer.Print "稿"
   
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 4.5)
   Printer.Print "不"
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 5.5)
   Printer.Print "會"
   If opt1(1).Value = True Then
       Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 3.25)
       Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 5.5)
       Printer.Print "Ｖ"
   End If
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 6.5)
   Printer.Print "稿"
   Printer.FontSize = 16
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(iStr(UBound(iStr)))) / 2
   'Modified by Lydia 2017/06/07 縮短空白間距
   'Printer.CurrentY = iY + 500
   Printer.CurrentY = iY + 200
   Printer.Print iStr(UBound(iStr))
   'add by nickc 2007/05/04
   'edit by nickc 2007/07/12 試著解決第二頁的格子線會不見的問題
   'If iCount = Val(txtPCnt) Then
       Printer.EndDoc
   'Else
   '    Printer.NewPage
   'End If
   
'Added by Lydia 2017/04/11
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox Err.Description
      Resume Next
   End If
'end 2017/04/11
End Sub

'Added by Lydia 2022/01/14 下載Word範本套印
Private Sub runWordProc(ByVal pSpace As Boolean)
Dim iStr(1 To 46) As String    '用印記錄(全文)
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
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000300-0-01 智權部委任契約書_P.docx", "M51", "000300", "0", "01", "4", "1")

   m_DefPath = App.path & "\" & strUserNum
   'Added by Lydia 2022/01/25
   m_TempPDF = ""
   '變更Word印表機
   PUB_SetOsDefaultPrinter Combo1
   PUB_SetWordActivePrinter
   'end 2022/01/25
  
   '下載範本檔: M51-000300-0-01 智權部委任契約書_P.docx
   m_TempFN = Pub_RepFileName(IIf(pSpace = False, Mid(Trim(txt1(5)), 1, 4), "空白")) 'Move by Lydia 2022/01/25 從m_TempFileName移過來
   'Modified by Lydia 2022/01/25 改成Word直接印，所以範本一開始就先命名好
   'm_FileName = "$$" & Me.Name & ".docx"
   m_FileName = "$$" & strUserNum & "_P_" & m_TempFN & ".docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000300-0-01", , m_DefPath) = False Then
        Exit Sub
   End If
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   'Remove by Lydia 2022/01/25 不用改存PDF檔
   'm_TempFileName = "$$" & strUserNum & "_P_" & m_TempFN & ".pdf"
   'If Dir(m_DefPath & "\" & m_TempFileName) <> "" Then
   '   Kill m_DefPath & "\" & m_TempFileName
   'End If
   'end 2022/01/25
   
   'Modified by Lydia 2022/01/20 改成直接用範本檔
'   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, , , False
'   g_WordAp.ActiveDocument.SaveAs m_DefPath & "\" & m_TempFileName
'   g_WordAp.ActiveDocument.Close
'   g_WordAp.Documents.Open m_DefPath & "\" & m_TempFileName
   'Q: AddToRecentFiles:=False還是會新增到最近開啟記錄
  'g_WordAp.Documents.Open FileName:=m_DefPath & "\" & m_FileName, ConfirmConversions:=False, ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", Format:=wdOpenFormatAuto
  g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
  
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 0 To 31
         strName = "PS" & Format(intA, "000")
         strText = ""
'-------第一條
         If intA = 1 Then
              '發明創作名稱
              strText = PUB_StrToStr(txt1(0), 60)
         ElseIf intA = 2 Then
              '發明或創作人
              strText = PUB_StrToStr(txt1(1), 38)
         ElseIf intA = 3 Then
              '發明或創作人-I.D.NO
              strText = PUB_StrToStr(txt1(2), 10)
         ElseIf intA = 4 Then
               '發明或創作人-地址
              strText = PUB_StrToStr(txt1(3), 50)
         ElseIf intA = 5 Then
              '發明或創作人-國籍
              strText = PUB_StrToStr(txt1(4), 8)
         ElseIf intA = 6 Then
              '申   請   人
              strText = PUB_StrToStr(txt1(5), 38)
         ElseIf intA = 7 Then
              '申請人-I.D.NO
              strText = PUB_StrToStr(txt1(6), 10)
         ElseIf intA = 8 Then
              '申請人-地址
              strText = PUB_StrToStr(txt1(7), 50)
         ElseIf intA = 9 Then
              '申請人-國籍
              strText = PUB_StrToStr(txt1(8), 8)
         ElseIf intA = 10 Then
              '指定聯絡人及地址
              strText = PUB_StrToStr(txt1(9), 54)
'-------第二條
         ElseIf intA = 11 Then
              '申請種類:
              strText = IIf(Chk1(0).Value = 1, "■", "□") & "發明  　　" & IIf(Chk1(1).Value = 1, "■", "□") & "新型  " & IIf(Chk1(2).Value = 1, "■", "□") & "設計　　　"
         ElseIf intA = 12 Then
              '性    質：1
              'Modified by Lydia 2022/08/30 不顯示訴願
              'Memo by Lydia 2022/09/23 (還原)經協商後，專利、商標案件之訴願程序將由智慧所承辦
              strText = IIf(Chk2(0).Value = 1, "■", "□") & "申請  　　" & IIf(Chk2(1).Value = 1, "■", "□") & "再審  " & IIf(Chk2(2).Value = 1, "■", "□") & "初審申復  " & IIf(Chk2(3).Value = 1, "■", "□") & "再審申復  " & IIf(Chk2(4).Value = 1, "■", "□") & "訴願" & IIf(Chk3.Value = 1, "■", "□") & "新型技術報告"
              'Mark by Lydia 2022/09/23 經協商後，專利、商標案件之訴願程序將由智慧所承辦
              'strText = IIf(Chk2(0).Value = 1, "■", "□") & "申請  　　" & IIf(Chk2(1).Value = 1, "■", "□") & "再審  " & IIf(Chk2(2).Value = 1, "■", "□") & "初審申復  " & IIf(Chk2(3).Value = 1, "■", "□") & "再審申復  " & IIf(Chk3.Value = 1, "■", "□") & "新型技術報告"
         ElseIf intA = 13 Then
              '性    質：2
              'Modified by Lydia 2022/08/30 不顯示行政訴訟 'Memo by Lydia 2022/10/04 只要還原訴願(2022/09/23 )
              'strText = IIf(chk2(5).Value = 1, "■", "□") & "行政訴訟  " & IIf(chk2(6).Value = 1, "■", "□") & "舉發　" & IIf(chk2(7).Value = 1, "■", "□") & "答辯　　  " & IIf(chk2(8).Value = 1, "■", "□") & "領證　　　" & IIf(chk2(9).Value = 1, "■", "□") & "繳年費"
              strText = IIf(Chk2(6).Value = 1, "■", "□") & "舉發  　　" & IIf(Chk2(7).Value = 1, "■", "□") & "答辯  " & IIf(Chk2(8).Value = 1, "■", "□") & "領證　　  " & IIf(Chk2(9).Value = 1, "■", "□") & "繳年費"
         ElseIf intA = 14 Then
              '性    質：3
              strText = IIf(Chk2(11).Value = 1, "■", "□") & "實體審查  " & IIf(Chk2(10).Value = 1, "■", "□") & "其他："
         ElseIf intA = 15 Then
              '性    質：其他 設定 底線
              strText = PUB_StrToStr(txt1(10) & " ", 42, True)
         ElseIf intA = 16 Then
              '性    質：4
              strText = IIf(Chk2(12).Value = 1, "■", "□") & "本案所涉之技術內容並非在中國大陸境內完成的發明或者實用新型"
'-------第三條
         'Added by Lydia 2022/01/19 酬金要區分項目描述
         ElseIf intA = 0 Then
                strExc(1) = "": strExc(2) = ""
                If Val(Trim(txt1(11))) = 0 Then
                     If pSpace = True Then strExc(1) = "　　　　　　　　　　　　元整"
                Else
                     strExc(1) = ChangeNumber(txt1(11))
                End If
               strExc(3) = PUB_StrToStr("　　一、前酬金新台幣　|#PS017#|，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 88 - Len(strExc(1)) * 2 + 9)
               strExc(4) = "　　　　" & Replace("　　一、前酬金新台幣　|#PS017#|，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", PUB_StrToStr("　　一、前酬金新台幣　|#PS017#|，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 88 - Len(strExc(1)) * 2 + 9), "")
       
              If pSpace = True Or (Val(Trim(txt1(11))) > 0 And Val(Trim(txt1(12))) > 0) Then
                   strText = strExc(3) & vbCrLf & strExc(4) & vbCrLf & _
                                "　　二、後酬金新台幣　|#PS018#|，於本程序終結時，由甲方一次付清。"
              Else
                   If Val(Trim(txt1(11))) > 0 Then
                        strText = Replace(strExc(3) & vbCrLf & strExc(4), "一、", "　　") & "|#PS018#|"
                   ElseIf Val(Trim(txt1(12))) > 0 Then
                        strText = "|#PS017#|　　　　後酬金新台幣　|#PS018#|，於本程序終結時，由甲方一次付清。"
                   Else
                        strText = "　　|#PS017#||#PS018#|"
                   End If
              End If
         'end 2022/01/19
         ElseIf intA = 17 Then
              '前酬金
              If Val(Trim(txt1(11))) = 0 Then
                   If pSpace = True Then strText = "　　　　　　　　　　　　元整"
              Else
                   strText = ChangeNumber(txt1(11))
              End If
         ElseIf intA = 18 Then
              '後酬金
              If Val(Trim(txt1(12))) = 0 Then
                   If pSpace = True Then strText = "　　　　　　　　　　　　元整"
              Else
                   strText = ChangeNumber(txt1(12))
              End If
'-------
         ElseIf intA = 19 Then
              '會稿
              strText = IIf(opt1(0).Value = True, "V", "")
         ElseIf intA = 20 Then
              '不會稿
              strText = IIf(opt1(1).Value = True, "V", "")
         ElseIf intA = 21 Then
              '委任人
              strText = PUB_StrToStr(txt1(13).Text, 50)
         ElseIf intA = 22 Then
              '委任人-ID NO.
              strText = PUB_StrToStr(txt1(14).Text, 20)
         ElseIf intA = 23 Then
              '委任人-代表人
              strText = PUB_StrToStr(txt1(15).Text, 24)
         ElseIf intA = 24 Then
              '委任人-地址
              strText = PUB_StrToStr(txt1(16).Text, 50)
         ElseIf intA = 25 Then
              '委任人-電話
              strText = PUB_StrToStr(txt1(17).Text, 20)
         ElseIf intA = 26 Then
              '委任人-傳真
              strText = PUB_StrToStr(txt1(18).Text, 24)
         ElseIf intA = 27 Then
              '受任人
              strText = Combo2.Text
         ElseIf intA = 28 Then
              '經手人
              strText = PUB_StrToStr(txt1(19).Text, 50)
         ElseIf intA = 29 Then
              '受任人-地址
              strText = PUB_SetAddrTofrm210114(Combo2.Text)
         ElseIf intA = 30 Then
              strText = "    中    華    民    國 " & String((8 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & txt1(20) & String((8 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & "年" & String((8 - LenB(StrConv((txt1(21)), vbFromUnicode))) / 2, " ") & txt1(21) & String((8 - LenB(StrConv((txt1(21)), vbFromUnicode))) / 2, " ") & "月" & String((8 - LenB(StrConv((txt1(22)), vbFromUnicode))) / 2, " ") & txt1(22) & String((8 - LenB(StrConv((txt1(22)), vbFromUnicode))) / 2, " ") & "日"
         ElseIf intA = 31 Then
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
            If (intA >= 1 And intA <= 10) Or intA = 16 Or (intA >= 21 And intA <= 24) Or intA = 28 Then
               '有Unicode字需要換字型
               .Selection.Font.Name = "細明體-ExtB"
            End If
            If intA = 15 Then
               '性    質：其他 設定 底線
               .Selection.Font.Underline = True
            End If
            If intA = 17 Or intA = 18 Then
                '金額要粗體
                .Selection.Font.Bold = True
            End If
            If intA = 19 Or intA = 20 Then
                '會稿/不會稿勾選
                .Selection.Font.Size = 14
            End If
            If intA = 31 And bolAddSeal = True Then  '公司章: 放在經手人的儲存格
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
            '保留;因為先全部以細明體-ExtB,最後全選改字型;
            If (intA >= 1 And intA <= 10) Or intA = 16 Or (intA >= 21 And intA <= 24) Or intA = 28 Then
               '有Unicode字需要換字型=>還原
               .Selection.Font.Name = "標楷體"
            End If
            If intA = 15 Then
               '性    質：其他 設定 底線
               .Selection.Font.Underline = False
            End If
            If intA = 17 Or intA = 18 Then
                '金額要粗體=>還原
                .Selection.Font.Bold = False
            End If
            If intA = 19 Or intA = 20 Then
                '會稿/不會稿勾選=>還原
                .Selection.Font.Size = 12
            End If
         End If
         
      Next intA
'      '因為先全部以細明體-ExtB,最後全選改字型;
      .Selection.WholeStory
      .Selection.Font.Name = "標楷體"
   End With
   
   'Modified by Lydia 2022/01/19 改存成PDF檔
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
   iStr(1) = "專利案件委任契約書"
   iStr(2) = "委任人(甲方)茲委任受任人(乙方)辦理國內專利案件，經雙方同意條件如下："
   iStr(3) = "第一條"
   iStr(4) = "　　一、發明創作名稱：" & StrToStr(txt1(0) & String(60, " "), 30)
   iStr(5) = "　　二、發明或創作人：" & StrToStr(txt1(1) & String(38, " "), 19) & "　I.D.NO：" & StrToStr(txt1(2) & String(10, " "), 5)
   iStr(6) = "　　　　地址：" & StrToStr(txt1(3) & String(50, " "), 25) & "　國籍：" & StrToStr(txt1(4) & String(8, " "), 4)
   iStr(7) = "　　三、申   請   人：" & StrToStr(txt1(5) & String(38, " "), 19) & "　I.D.NO：" & StrToStr(txt1(6) & String(10, " "), 5)
   iStr(8) = "　　　　地址：" & StrToStr(txt1(7) & String(50, " "), 25) & "　國籍：" & StrToStr(txt1(8) & String(8, " "), 4)
   iStr(9) = "　　四、指定聯絡人及地址：" & StrToStr(txt1(9) & String(54, " "), 27)
   iStr(10) = "第二條　委辦範圍"
   iStr(11) = "　　　　申請種類：" & IIf(Chk1(0).Value = 1, "■", "□") & "發明  　　" & IIf(Chk1(1).Value = 1, "■", "□") & "新型  " & IIf(Chk1(2).Value = 1, "■", "□") & "設計　　　"
   'Modified by Lydia 2022/08/30 不顯示訴願,行政訴訟
   'Memo by Lydia 2022/09/23 (還原)經協商後，專利、商標案件之訴願程序將由智慧所承辦
   'Modified by Lydia 2022/10/04 (debug) 只要還原訴願(2022/09/23 )
   'iStr(12) = "　　　　性    質：" & IIf(chk2(0).Value = 1, "■", "□") & "申請  　　" & IIf(chk2(1).Value = 1, "■", "□") & "再審  " & IIf(chk2(2).Value = 1, "■", "□") & "初審申復  " & IIf(chk2(3).Value = 1, "■", "□") & "再審申復  " & IIf(chk2(4).Value = 1, "■", "□") & "訴願" & IIf(Chk3.Value = 1, "■", "□") & "新型技術報告　"
   'iStr(13) = "　　　　　　　　　" & IIf(chk2(5).Value = 1, "■", "□") & "行政訴訟  " & IIf(chk2(6).Value = 1, "■", "□") & "舉發　" & IIf(chk2(7).Value = 1, "■", "□") & "答辯　　  " & IIf(chk2(8).Value = 1, "■", "□") & "領證　　　" & IIf(chk2(9).Value = 1, "■", "□") & "繳年費"
   iStr(12) = "　　　　性    質：" & IIf(Chk2(0).Value = 1, "■", "□") & "申請　　  " & IIf(Chk2(1).Value = 1, "■", "□") & "再審  " & IIf(Chk2(2).Value = 1, "■", "□") & "初審申復  " & IIf(Chk2(3).Value = 1, "■", "□") & "再審申復  " & IIf(Chk3.Value = 1, "■", "□") & "新型技術報告　"
   iStr(13) = "　　　　　　　　　" & IIf(Chk2(6).Value = 1, "■", "□") & "舉發　　  " & IIf(Chk2(7).Value = 1, "■", "□") & "答辯　　  " & IIf(Chk2(8).Value = 1, "■", "□") & "領證　　　" & IIf(Chk2(9).Value = 1, "■", "□") & "繳年費"
   'end 2022/10/04
   iStr(14) = "　　　　　　　　　" & IIf(Chk2(11).Value = 1, "■", "□") & "實體審查  " & IIf(Chk2(10).Value = 1, "■", "□") & "其他：" & txt1(10).Text
   iStr(15) = "　　　　　　　　　" & IIf(Chk2(12).Value = 1, "■", "□") & "本案所涉之技術內容並非在中國大陸境內完成的發明或者實用新型"
   iStr(16) = "　　　　乙方根據甲方所提供資料，依前項約定之範圍，代撰必要書件，向本程序主管機關"
   iStr(17) = "　　　　提出，並代為收受有關文件。"
   iStr(18) = "第三條　委辦費用"  '第三條(起)
   '空白委任書要保留費用
   If Val(Trim(txt1(11))) = 0 And pSpace = False Then
       iStr(19) = ""
       iStr(20) = ""
   Else
       If Val(Trim(txt1(11))) = 0 Then
          strSpaceAmt = "　　　　　　　　　　　　元整"
       Else
          strSpaceAmt = ChangeNumber(txt1(11))
       End If
       iStr(19) = StrToStr("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 44)
       iStr(20) = "　　　　" & Replace("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", StrToStr("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 44), "")
       '金額直接併入說明
       strExc(3) = StrToStr("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 44)
       strExc(4) = "　　　　" & Replace("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", StrToStr("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 44), "")
   End If
   '空白委任書要保留費用
   If Val(Trim(txt1(12))) = 0 And pSpace = False Then
       iStr(21) = ""
       iStr(22) = ""
       'Added by Morgan 2022/1/18
       iStr(19) = Replace(iStr(19), "一、", "　　")
       iStr(20) = Replace(iStr(20), "一、", "　　")
       strExc(3) = Replace(strExc(3), "一、", "　　")
       strExc(4) = Replace(strExc(4), "一、", "　　")
       'end 2022/1/18
   Else
       If Val(Trim(txt1(12))) = 0 Then
          strSpaceAmt = "　　　　　　　　　　　　元整"
       Else
          strSpaceAmt = ChangeNumber(txt1(12))
       End If
       iStr(21) = StrToStr("　　二、後酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本程序終結時，由甲方一次付清。", 44)
       strExc(5) = StrToStr("　　二、後酬金新台幣　" & strSpaceAmt & "，於本程序終結時，由甲方一次付清。", 44)
       
       'Added by Morgan 2022/1/18
       If iStr(19) = "" Then
         iStr(21) = Replace(iStr(21), "二、", "　　")
         iStr(22) = Replace(iStr(22), "二、", "　　")
         strExc(5) = Replace(strExc(5), "二、", "　　")
         strExc(6) = Replace(strExc(6), "二、", "　　")
       End If
       'end 2022/1/18
   End If
   iStr(22) = "第四條　乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，並不得發生足以影響甲方"
   iStr(23) = "　　　　權益之疏誤，否則應對甲方負損害賠償責任。但以不超過第三條所載前酬金金額之三倍為限。"
   'Memo by Lydia 2021/04/19 改第五條內容
   iStr(24) = "第五條　甲方應確保所交付予乙方之資料及本契約書所載內容(包括發明人或創作人、申請人等資訊)"
   iStr(25) = "　　　　均無虛偽情事，且甲方確實得到與委辦案件相關共同發明人及第三人之同意，有權委託乙方"
   iStr(26) = "　　　　辦理案件，如因不實致生損害或法律責任時，概由甲方負責，與乙方無關。"
   iStr(27) = "第六條　乙方於辦理過程中，應隨時將辦理經過如申請日期、案號及其他重要函件，儘速通知或交付"
   iStr(28) = "　　　　甲方。但甲方於簽約後變更連絡處所，未即時通知乙方，因而連絡不及致延誤時限者，乙方"
   iStr(29) = "　　　　不負責任。" '
   iStr(30) = "第七條　凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時限者，乙方不負責。經乙方"
   iStr(31) = "　　　　通知甲方繳費而未依限繳納者，亦同。"
   iStr(32) = "第八條　甲方如逕自撤回所委辦程序，或未經乙方同意終止契約時，所約定之費用，仍應全數給付。"
   iStr(33) = "第九條　本契約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，但有增刪修改時，需甲乙雙方"
   iStr(34) = "　　　　於更動處蓋章始生效力，並由雙方各執乙份為憑。"
   iStr(35) = "          "   '第九條(止)
   iStr(36) = "　　　　　　甲方：委任人：" & StrToStr(txt1(13) & String(54, " "), 27)
   iStr(37) = "　　　　　　　　　ID NO.：" & StrToStr(txt1(14) & String(20, " "), 10) & "代表人：" & StrToStr(txt1(15) & String(26, " "), 13)
   iStr(38) = "　　　　　　      地  址：" & StrToStr(txt1(16) & String(54, " "), 27)
   iStr(39) = "　　　　　　　　　電  話：" & StrToStr(txt1(17) & String(20, " "), 10) & "傳  真：" & StrToStr(txt1(18) & String(26, " "), 13)
   iStr(40) = "　　　　　　乙方：受任人：" & Combo2.Text '"台一國際專利法律事務所"
   iStr(41) = "　　　　　　　　　經手人：" & StrToStr(txt1(19) & String(54, " "), 27)
   iStr(42) = "　　　　　　　　　地　址：" & PUB_SetAddrTofrm210114(Combo2.Text) '改用模組控制
   iStr(43) = "　　　　　　　　　電  話：(02)25061023(總機)   FAX:(02)25011666"
   iStr(44) = "　　　　　　　　　網  址：www.taie.com.tw"
   iStr(45) = "　　　　　　　　　E-mail：ipdept@taie.com.tw"
   iStr(46) = "  中    華    民    國 " & String((10 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & txt1(20) & String((10 - LenB(StrConv((txt1(20)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(21)), vbFromUnicode))) / 2, " ") & txt1(21) & String((10 - LenB(StrConv((txt1(21)), vbFromUnicode))) / 2, " ") & "月" & String((10 - LenB(StrConv((txt1(22)), vbFromUnicode))) / 2, " ") & txt1(22) & String((10 - LenB(StrConv((txt1(22)), vbFromUnicode))) / 2, " ") & "日"

   '有用印就記錄列印內容
    strDetail = ""
    For intI = 1 To UBound(iStr)
       If Trim(iStr(intI)) <> "" Then
          If intI >= 19 And intI <= 21 Then
             Select Case intI
                 Case 19: strDetail = strDetail & vbCrLf & RTrim(strExc(3))
                 Case 20: strDetail = strDetail & vbCrLf & RTrim(strExc(4))
                 Case 21: strDetail = strDetail & vbCrLf & RTrim(strExc(5))
             End Select
          Else
             If intI <= 18 Or (intI >= 36 And intI <= 41) Or intI = 46 Then
                 If intI = 40 Then
                   strDetail = strDetail & vbCrLf & vbCrLf & IIf(opt1(0).Value = True, "會稿", "不會稿") & vbCrLf
                 End If
                strDetail = strDetail & vbCrLf & RTrim(iStr(intI))
             End If
          End If
       End If
    Next
       If PUB_AddRecSeal("1", txtPCnt.Text, IIf(pSpace = True, "Y", ""), strDetail, Combo2.Text) Then
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
   PUB_KillTempFile (strUserNum & "\$$" & strUserNum & "*_P*.*")
   PUB_KillTempFile (strUserNum & "\$$" & Me.Name & "*.*")
End Sub

