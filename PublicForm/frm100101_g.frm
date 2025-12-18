VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_g 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖進度資料查詢"
   ClientHeight    =   5130
   ClientLeft      =   320
   ClientTop       =   1380
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9000
   Begin VB.CommandButton cmd 
      Caption         =   "回前畫面"
      Default         =   -1  'True
      Height          =   400
      Left            =   7620
      TabIndex        =   65
      Top             =   70
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "草圖是否計件："
      Height          =   255
      Index           =   41
      Left            =   3480
      TabIndex        =   80
      Top             =   210
      Width           =   1260
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   17
      Left            =   4725
      TabIndex        =   79
      Top             =   210
      Width           =   270
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "476;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(N：不計)"
      Height          =   255
      Index           =   40
      Left            =   5115
      TabIndex        =   78
      Top             =   210
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "墨圖加乘註記修改理由："
      Height          =   255
      Index           =   39
      Left            =   6630
      TabIndex        =   77
      Top             =   2790
      Width           =   1980
   End
   Begin MSForms.Label lbl1 
      Height          =   795
      Index           =   36
      Left            =   6615
      TabIndex        =   76
      Top             =   3015
      Width           =   2340
      Caption         =   "lblFM2"
      Size            =   "4128;1402"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "墨圖加乘註記："
      Height          =   255
      Index           =   38
      Left            =   6630
      TabIndex        =   75
      Top             =   2520
      Width           =   1260
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   35
      Left            =   7950
      TabIndex        =   74
      Top             =   2520
      Width           =   800
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1411;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "墨圖計件值："
      Height          =   255
      Index           =   37
      Left            =   6630
      TabIndex        =   73
      Top             =   2250
      Width           =   1080
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   34
      Left            =   7815
      TabIndex        =   72
      Top             =   2250
      Width           =   800
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1411;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "草圖加乘註記修改理由："
      Height          =   255
      Index           =   36
      Left            =   6630
      TabIndex        =   71
      Top             =   1155
      Width           =   1980
   End
   Begin MSForms.Label lbl1 
      Height          =   765
      Index           =   33
      Left            =   6615
      TabIndex        =   70
      Top             =   1395
      Width           =   2340
      Caption         =   "lblFM2"
      Size            =   "4128;1349"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "草圖加乘註記："
      Height          =   255
      Index           =   35
      Left            =   6630
      TabIndex        =   69
      Top             =   870
      Width           =   1260
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   32
      Left            =   7950
      TabIndex        =   68
      Top             =   885
      Width           =   800
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1411;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "草圖計件值："
      Height          =   255
      Index           =   0
      Left            =   6630
      TabIndex        =   67
      Top             =   600
      Width           =   1080
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   31
      Left            =   7875
      TabIndex        =   66
      Top             =   600
      Width           =   800
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1411;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(N：不計)"
      Height          =   255
      Index           =   32
      Left            =   5115
      TabIndex        =   64
      Top             =   1245
      Width           =   780
   End
   Begin MSForms.Label lbl1 
      Height          =   1155
      Index           =   30
      Left            =   4095
      TabIndex        =   63
      Top             =   3900
      Width           =   3030
      Caption         =   "lblFM2"
      Size            =   "5345;2037"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   25
      Left            =   5055
      TabIndex        =   62
      Top             =   2295
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   20
      Left            =   4725
      TabIndex        =   61
      Top             =   990
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   23
      Left            =   4725
      TabIndex        =   60
      Top             =   2025
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   22
      Left            =   4725
      TabIndex        =   59
      Top             =   1770
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   24
      Left            =   4725
      TabIndex        =   58
      Top             =   1245
      Width           =   270
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "476;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   19
      Left            =   4725
      TabIndex        =   57
      Top             =   735
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   18
      Left            =   4725
      TabIndex        =   56
      Top             =   465
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   37
      Left            =   5205
      TabIndex        =   55
      Top             =   3570
      Width           =   270
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "476;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   255
      Index           =   31
      Left            =   3480
      TabIndex        =   54
      Top             =   3900
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "複雜時數："
      Height          =   255
      Index           =   22
      Left            =   3480
      TabIndex        =   53
      Top             =   2295
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "草圖張數："
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   52
      Top             =   990
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "草圖完稿日："
      Height          =   255
      Index           =   26
      Left            =   3480
      TabIndex        =   51
      Top             =   735
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "草圖齊備日："
      Height          =   255
      Index           =   23
      Left            =   3480
      TabIndex        =   50
      Top             =   465
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "墨圖是否計件："
      Height          =   255
      Index           =   17
      Left            =   3480
      TabIndex        =   49
      Top             =   1245
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶是否提供圖檔：          (Y：提供)"
      Height          =   255
      Index           =   7
      Left            =   3480
      TabIndex        =   48
      Top             =   3570
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "墨圖齊備日："
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   47
      Top             =   1515
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "墨圖完稿日："
      Height          =   255
      Index           =   28
      Left            =   3480
      TabIndex        =   46
      Top             =   1770
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "墨圖張數："
      Height          =   255
      Index           =   34
      Left            =   3480
      TabIndex        =   45
      Top             =   2025
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   21
      Left            =   4725
      TabIndex        =   44
      Top             =   1515
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "草圖："
      Height          =   255
      Index           =   5
      Left            =   4425
      TabIndex        =   43
      Top             =   2295
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "墨圖："
      Height          =   255
      Index           =   24
      Left            =   4425
      TabIndex        =   42
      Top             =   2550
      Width           =   540
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   26
      Left            =   5055
      TabIndex        =   41
      Top             =   2550
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "2."
      Height          =   255
      Index           =   25
      Left            =   4530
      TabIndex        =   40
      Top             =   3075
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "3."
      Height          =   255
      Index           =   27
      Left            =   4530
      TabIndex        =   39
      Top             =   3330
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "繪圖人員："
      Height          =   255
      Index           =   4
      Left            =   90
      TabIndex        =   38
      Top             =   210
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國外案承辦人："
      Height          =   255
      Index           =   9
      Left            =   90
      TabIndex        =   37
      Top             =   4095
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國外案本所案號："
      Height          =   255
      Index           =   10
      Left            =   90
      TabIndex        =   36
      Top             =   4350
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點數："
      Height          =   255
      Index           =   11
      Left            =   90
      TabIndex        =   35
      Top             =   2010
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   255
      Index           =   12
      Left            =   90
      TabIndex        =   34
      Top             =   3045
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   255
      Index           =   13
      Left            =   90
      TabIndex        =   33
      Top             =   2535
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   255
      Index           =   14
      Left            =   90
      TabIndex        =   32
      Top             =   2265
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   255
      Index           =   15
      Left            =   90
      TabIndex        =   31
      Top             =   1755
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利/商標種類："
      Height          =   255
      Index           =   16
      Left            =   90
      TabIndex        =   30
      Top             =   1485
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   18
      Left            =   90
      TabIndex        =   29
      Top             =   1230
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Index           =   19
      Left            =   90
      TabIndex        =   28
      Top             =   975
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   255
      Index           =   20
      Left            =   90
      TabIndex        =   27
      Top             =   735
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "總收文號："
      Height          =   255
      Index           =   21
      Left            =   90
      TabIndex        =   26
      Top             =   465
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "取消收文日："
      Height          =   255
      Index           =   29
      Left            =   90
      TabIndex        =   25
      Top             =   3825
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "草圖作業天數："
      Height          =   255
      Index           =   33
      Left            =   90
      TabIndex        =   24
      Top             =   3315
      Width           =   1290
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   10
      Left            =   1050
      TabIndex        =   23
      Top             =   2790
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   1050
      TabIndex        =   22
      Top             =   465
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1050
      TabIndex        =   21
      Top             =   210
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1035
      TabIndex        =   20
      Top             =   735
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   1050
      TabIndex        =   19
      Top             =   975
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   4
      Left            =   1050
      TabIndex        =   18
      Top             =   1230
      Width           =   2355
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4154;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   5
      Left            =   1500
      TabIndex        =   17
      Top             =   1485
      Width           =   1905
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3360;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   7
      Left            =   1050
      TabIndex        =   16
      Top             =   2010
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   8
      Left            =   1050
      TabIndex        =   15
      Top             =   2265
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   9
      Left            =   1050
      TabIndex        =   14
      Top             =   2535
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   11
      Left            =   1050
      TabIndex        =   13
      Top             =   3045
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   6
      Left            =   1050
      TabIndex        =   12
      Top             =   1755
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   16
      Left            =   1605
      TabIndex        =   11
      Top             =   4350
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   12
      Left            =   1410
      TabIndex        =   10
      Top             =   3315
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   14
      Left            =   1395
      TabIndex        =   9
      Top             =   3825
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   15
      Left            =   1395
      TabIndex        =   8
      Top             =   4095
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "墨圖作業天數："
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   3570
      Width           =   1290
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   13
      Left            =   1410
      TabIndex        =   6
      Top             =   3570
      Width           =   2000
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3528;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   28
      Left            =   4725
      TabIndex        =   5
      Top             =   3075
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   29
      Left            =   4725
      TabIndex        =   4
      Top             =   3330
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   255
      Index           =   8
      Left            =   90
      TabIndex        =   3
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "修改時數："
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   2
      Top             =   2805
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   27
      Left            =   4725
      TabIndex        =   1
      Top             =   2805
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1."
      Height          =   255
      Index           =   30
      Left            =   4530
      TabIndex        =   0
      Top             =   2805
      Width           =   135
   End
End
Attribute VB_Name = "frm100101_g"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/24 改成Form2.0 ; lbl1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer, s As Integer, strSql As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer


'92.04.16 nick
Public Sub PubShowNextData()
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
End Sub

Private Sub cmd_Click()
'92.04.16 nick 紀錄作用按鍵
cmdState = 100
PubShowNextData
Exit Sub
'92.04.16 nick 以下無效
Me.Hide
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
'92.04.16 nick
cmdState = -1
End Sub

Sub Process(strText As String)
pub_QL05 = ";總收文號：" & strText & "(繪圖進度)" 'Add By Sindy 2025/8/7
'Modify by Morgan 2004/1/19
'strSQL = "SELECT S1.ST02,cp09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa01,'CFP',ptm03,DECODE(PA09,'000',PTM03,PTM04)),decode(pa09,'000',cpm03,cpm04),cp18," & SQLDate("cp06") & "," & SQLDate("cp07") & ",S2.ST02,s3.st02,0,0," & SQLDate("CP57") & ",'',''," & SQLDate("cP48") & "," & SQLDate("eP14") & "," & SQLDate("eP15") & ",ep16," & SQLDate("ep17") & "," & SQLDate("ep18") & ",ep19,ep20,ep21,ep22,ep23,ep24,ep25,ep26 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND ep13=S1.ST01(+) AND eP05=S2.ST01(+) AND cP13=S3.ST01(+) AND cp09='" & StrText & "' "
'edit by nickc 2005/03/31 加欄位
'StrSql = "SELECT EP13||'  '||S1.ST02 as EP13N,cp09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa01,'CFP',ptm03,DECODE(PA09,'000',PTM03,PTM04)),decode(pa09,'000',cpm03,cpm04),cp18," & SQLDate("cp06") & "," & SQLDate("cp07") & ",EP05||'  '||S2.ST02 as EP05N,CP13||'  '||s3.st02 AS CP13N,0,0," & SQLDate("CP57") & ",'',''," & SQLDate("cP48") & "," & SQLDate("eP14") & "," & SQLDate("eP15") & ",ep16," & SQLDate("ep17") & "," & SQLDate("ep18") & ",ep19,ep20,ep21,ep22,ep23,ep24,ep25,ep26 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND ep13=S1.ST01(+) AND eP05=S2.ST01(+) AND cP13=S3.ST01(+) AND cp09='" & StrText & "' "
strSql = "SELECT EP13||'  '||S1.ST02 as EP13N,cp09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),decode(pa01,'CFP',ptm03,DECODE(PA09,'000',PTM03,PTM04)),decode(pa09,'000',cpm03,cpm04),cp18," & SQLDate("cp06") & "," & SQLDate("cp07") & ",EP05||'  '||S2.ST02 as EP05N,CP13||'  '||s3.st02 AS CP13N,0,0," & SQLDate("CP57") & ",'','',EP20," & SQLDate("eP14") & "," & SQLDate("eP15") & ",ep16," & SQLDate("ep17") & "," & SQLDate("ep18") & ",ep19,ep29,ep21,ep22,ep23,ep24,ep25,ep26,cp100,cp101,cp102,cp103,cp104,cp105,CP106 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND ep13=S1.ST01(+) AND eP05=S2.ST01(+) AND cP13=S3.ST01(+) AND cp09='" & strText & "' "
'Modify end 2004/1/19
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        pub_QL05 = ";本所案號：" & .Fields(3) & pub_QL05 'Add By Sindy 2025/8/7
        If pub_QL04 <> "" Then InsertQueryLog (.RecordCount) 'Add By Sindy 2025/8/7
        .MoveFirst
        'edit by nickc 2005/03/31
        'For i = 0 To 30
        For i = 0 To 37
            lbl1(i) = CheckStr(.Fields(i))
            If i = 6 Then
                Me.lbl1(i).Caption = Me.lbl1(i).Caption & PUB_GetRelateCasePropertyName(Me.lbl1(1).Caption, "1")
            End If
            'add by nickc 2008/02/04
            If i = 23 Then lbl1(i) = Val(lbl1(i))
            
            'Add by Morgan 2004/1/19
            Select Case .Fields(i).Name
                Case "EP13N", "EP05N", "CP13N"
                    If Len(.Fields(i)) = 8 Then
                        lbl1(i).ForeColor = vbRed
                    Else
                        lbl1(i).ForeColor = vbBlack
                    End If
            End Select
            'Add end 2004/1/19
        Next i
        '92.04.03 nick add left join
        'strSQL = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01 AND CP02=CM02 AND CP03=CM03 AND CP04=CM04 AND CP14=ST01(+) AND CP31='Y' and cp09='" & StrText & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
        'edit by nickc 2005/03/31 修正，因為串錯了
        'StrSql = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01(+) AND CP02=CM02(+) AND CP03=CM03(+) AND CP04=CM04(+) AND CP14=ST01(+) AND CP31='Y' and cp09='" & StrText & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
        strSql = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS C1,caseprogress C2,STAFF WHERE c1.cp01=cm05(+) AND c1.cp02=CM06(+) AND c1.cp03=CM07(+) AND c1.cp04=CM08(+) AND C2.CP14=ST01(+) AND c2.CP31='Y' and c1.cp09='" & strText & "' and cm01=c2.cp01(+) and cm02=c2.cp02(+) and cm03=c2.cp03(+) and cm04=c2.cp04(+) order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            lbl1(16) = CheckStr(adoRecordset1.Fields(0))
            lbl1(15) = CheckStr(adoRecordset1.Fields(1))
        Else
            lbl1(16) = ""
            lbl1(15) = ""
        End If
        CheckOC2
        '計算草圖作業天數
        If Len(lbl1(18)) <> 0 And Len(lbl1(19)) <> 0 And Val(lbl1(18)) <> 0 And Val(lbl1(19)) <> 0 Then
            lbl1(12) = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(lbl1(19))), ChangeTStringToWString(ChangeTDateStringToTString(lbl1(18))))
        End If
        '計算墨圖作業天數
        If Len(lbl1(21)) <> 0 And Len(lbl1(22)) <> 0 And Val(lbl1(21)) <> 0 And Val(lbl1(22)) <> 0 Then
            lbl1(13) = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(lbl1(22))), ChangeTStringToWString(ChangeTDateStringToTString(lbl1(21))))
        End If

    Else
        If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/7
    End If
End With
CheckOC

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm100101_g = Nothing
End Sub

