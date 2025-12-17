VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21r0 
   AutoRedraw      =   -1  'True
   Caption         =   "客戶/代理人財務EMail資料維護"
   ClientHeight    =   6000
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9108
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6009.513
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   9198.484
   Begin VB.TextBox txtNameNoUni 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   68
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.FileListBox File1 
      Height          =   180
      Left            =   8760
      TabIndex        =   67
      Top             =   4320
      Visible         =   0   'False
      Width           =   297
   End
   Begin VB.CheckBox Check3 
      Caption         =   "　"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   7800
      TabIndex        =   63
      Top             =   4332
      Width           =   198
   End
   Begin VB.CheckBox Check6 
      Caption         =   "不索取CF對帳單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6576
      TabIndex        =   61
      Top             =   3312
      Width           =   1848
   End
   Begin VB.CheckBox Check5 
      Caption         =   "國外特殊收據"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   13
      Top             =   3252
      Width           =   1683
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   255
      Left            =   2550
      TabIndex        =   54
      Top             =   4656
      Width           =   5595
      Begin VB.CheckBox chkA4228 
         Caption         =   "單筆收據稅額超過2000元"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Index           =   1
         Left            =   2400
         TabIndex        =   21
         Top             =   30
         Width           =   2625
      End
      Begin VB.CheckBox chkA4228 
         Caption         =   "每筆代繳"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   20
         Top             =   30
         Width           =   1185
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "代填方式："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   252
         Index           =   4
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   1050
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "境外公司"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3330
      TabIndex        =   15
      Top             =   4032
      Width           =   1188
   End
   Begin VB.CheckBox Check2 
      Caption         =   "零稅率"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4560
      TabIndex        =   16
      Top             =   4032
      Width           =   990
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   324
      Left            =   59
      TabIndex        =   47
      Top             =   3600
      Width           =   8796
      Begin VB.TextBox txtCU196 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6096
         MaxLength       =   1
         TabIndex        =   50
         Top             =   0
         Width           =   408
      End
      Begin VB.ComboBox Combo4 
         Height          =   276
         ItemData        =   "Frmacc21r0.frx":0000
         Left            =   6552
         List            =   "Frmacc21r0.frx":000D
         Style           =   2  '單純下拉式
         TabIndex        =   51
         Top             =   24
         Width           =   2016
      End
      Begin VB.CheckBox Check4 
         Caption         =   "寄電子檔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2310
         TabIndex        =   49
         Top             =   30
         Visible         =   0   'False
         Width           =   1188
      End
      Begin VB.CheckBox Check4 
         Caption         =   "寄紙本"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1290
         TabIndex        =   48
         Top             =   30
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "請款匯入銀行資料"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   4152
         TabIndex        =   60
         Top             =   48
         Width           =   1872
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "發票寄送方式"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   57
         Top             =   30
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin VB.TextBox txtInform 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   7776
      MaxLength       =   1
      TabIndex        =   19
      Top             =   3936
      Width           =   345
   End
   Begin VB.TextBox txtInform 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   1950
      MaxLength       =   1
      TabIndex        =   14
      Top             =   3936
      Width           =   345
   End
   Begin VB.TextBox textCU16 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   40
      Top             =   30
      Width           =   1815
   End
   Begin VB.TextBox TextCU168 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   39
      Text            =   "Frmacc21r0.frx":002D
      Top             =   5592
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox textCU169 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   17
      Top             =   4296
      Width           =   345
   End
   Begin VB.CheckBox Check3 
      Caption         =   "　"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   6708
      TabIndex        =   18
      Top             =   4332
      Width           =   218
   End
   Begin VB.CommandButton cmdA49 
      Caption         =   "會計師資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      Style           =   1  '圖片外觀
      TabIndex        =   25
      Top             =   30
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CheckBox Check1 
      Caption         =   "不同意抵帳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   5850
      TabIndex        =   12
      Top             =   3012
      Width           =   1392
   End
   Begin VB.CheckBox Check1 
      Caption         =   "宣告破產"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4260
      TabIndex        =   11
      Top             =   3012
      Width           =   1287
   End
   Begin VB.CheckBox Check1 
      Caption         =   "帳款處理中"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2865
      TabIndex        =   10
      Top             =   3012
      Width           =   1366
   End
   Begin VB.CheckBox Check1 
      Caption         =   "同意抵帳中"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1350
      TabIndex        =   9
      Top             =   3012
      Width           =   1505
   End
   Begin VB.TextBox txtInform 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   7440
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2616
      Width           =   345
   End
   Begin VB.TextBox txtInform 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   5040
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2616
      Width           =   345
   End
   Begin VB.TextBox txtInform 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2616
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   300
      Left            =   3630
      Picture         =   "Frmacc21r0.frx":003E
      Style           =   1  '圖片外觀
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   350
   End
   Begin VB.TextBox txtKey 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      MaxLength       =   9
      TabIndex        =   0
      Top             =   30
      Width           =   1572
   End
   Begin VB.Label lblCU168 
      AutoSize        =   -1  'True
      Caption         =   "法律所"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   7981
      TabIndex        =   66
      Top             =   4332
      Width           =   648
   End
   Begin VB.Label lblCU168 
      AutoSize        =   -1  'True
      Caption         =   "智慧所"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   6911
      TabIndex        =   65
      Top             =   4332
      Width           =   648
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "每月代填繳款同意書"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   4812
      TabIndex        =   64
      Top             =   4356
      Width           =   1836
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "財務副本信箱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   315
      TabIndex        =   62
      Top             =   1920
      Width           =   1440
   End
   Begin MSForms.TextBox txtBox 
      Height          =   348
      Index           =   2
      Left            =   1800
      TabIndex        =   5
      Top             =   1872
      Width           =   6828
      VariousPropertyBits=   671105051
      MaxLength       =   200
      Size            =   "12044;614"
      FontName        =   "新細明體"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBox 
      Height          =   348
      Left            =   1800
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2220
      Width           =   6828
      VariousPropertyBits=   679495707
      BackColor       =   14737632
      DisplayStyle    =   3
      Size            =   "12039;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtName 
      Height          =   336
      Index           =   0
      Left            =   0
      TabIndex        =   45
      Top             =   60
      Visible         =   0   'False
      Width           =   240
      VariousPropertyBits=   671105051
      BackColor       =   14737632
      Size            =   "423;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtBox 
      Height          =   348
      Index           =   5
      Left            =   1800
      TabIndex        =   4
      Top             =   1512
      Width           =   6828
      VariousPropertyBits=   671105051
      MaxLength       =   50
      Size            =   "12044;614"
      FontName        =   "新細明體"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU170 
      Height          =   348
      Left            =   1080
      TabIndex        =   22
      Top             =   4932
      Width           =   7812
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "13785;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU171 
      Height          =   348
      Left            =   6096
      TabIndex        =   24
      Top             =   5556
      Width           =   2808
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "4948;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   636
      Left            =   1092
      TabIndex        =   23
      Top             =   5292
      Width           =   4968
      VariousPropertyBits=   -1466941413
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "8758;1129"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtBox 
      Height          =   345
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   1170
      Width           =   6825
      VariousPropertyBits=   671105051
      MaxLength       =   200
      Size            =   "12039;609"
      FontName        =   "新細明體"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtBox 
      Height          =   345
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   6825
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "12039;609"
      FontName        =   "新細明體"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboCusName 
      Height          =   345
      Left            =   810
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   390
      Width           =   7890
      VariousPropertyBits=   679495707
      BackColor       =   14737632
      DisplayStyle    =   3
      Size            =   "13917;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl_fa101 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H008080FF&
      BorderStyle     =   1  '單線固定
      Caption         =   "？"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   7860
      TabIndex        =   56
      Top             =   2652
      Width           =   252
   End
   Begin VB.Label Lbl_Inf1 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H008080FF&
      BorderStyle     =   1  '單線固定
      Caption         =   "？"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   5556
      TabIndex        =   53
      Top             =   2976
      Width           =   252
   End
   Begin VB.Label Lbl_Inf 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H008080FF&
      BorderStyle     =   1  '單線固定
      Caption         =   "？"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   5556
      TabIndex        =   52
      Top             =   3996
      Width           =   252
   End
   Begin VB.Label Lbl_InfMail 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H008080FF&
      BorderStyle     =   1  '單線固定
      Caption         =   "？"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1410
      TabIndex        =   46
      Top             =   1170
      Width           =   255
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "財務信箱(CF)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   315
      TabIndex        =   44
      Top             =   1560
      Width           =   1428
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收據列印統一編號       (Y:印)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   14
      Left            =   6096
      TabIndex        =   43
      Top             =   4032
      Width           =   2604
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "不寄發扣繳核對資料      (N:不寄)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   60
      TabIndex        =   42
      Top             =   4032
      Width           =   3036
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶電話1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5970
      TabIndex        =   41
      Top             =   90
      Width           =   1110
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "繳款書寄件處      (1.客戶 2.會計師 3.特殊)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   60
      TabIndex        =   38
      Top             =   4356
      Width           =   3852
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "特殊地址"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   60
      TabIndex        =   37
      Top             =   4992
      Width           =   960
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "特殊收件人:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6096
      TabIndex        =   36
      Top             =   5352
      Width           =   1272
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "會計備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   60
      TabIndex        =   35
      Top             =   5316
      Width           =   960
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "帳款處理情形"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   60
      TabIndex        =   34
      Top             =   3012
      Width           =   1260
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "不寄催款單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   6360
      TabIndex        =   33
      Top             =   2676
      Width           =   1056
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否寄發電匯通知      (N:不寄)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   3360
      TabIndex        =   32
      Top             =   2676
      Width           =   2820
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否寄發收據      (N:不寄 B:Ｅ+印)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   60
      TabIndex        =   31
      Top             =   2676
      Width           =   3192
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "其他信箱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   312
      TabIndex        =   30
      Top             =   2220
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "財務信箱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   315
      TabIndex        =   29
      Top             =   1170
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "代表信箱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   315
      TabIndex        =   28
      Top             =   840
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1788
      Left            =   180
      Top             =   816
      Width           =   8580
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   27
      Top             =   420
      Width           =   915
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "客戶/代理人編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   26
      Top             =   60
      Width           =   1755
   End
End
Attribute VB_Name = "Frmacc21r0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modified by Lydia 2021/12/02 改成Form 2.0; txtBox(index)、textCU170、textCU171、Text1、txtName(0)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Dim ado21r0 As New ADODB.Recordset
'Add by Amy 2013/10/18
Dim oCheck As CheckBox
Dim oldDizhang As String '記錄帳款處理情形
Dim m_CU11 As String 'Add By Sindy 2015/12/11
Dim oText As Object
Dim m_PrevForm As Form '前一畫面'Add By Sindy 2016/11/29
Dim bolMod As Boolean 'Added by Lydia 2016/12/26 是否可維護本筆記錄
Dim strDizhangRec As String 'Added by Lydia 2017/01/16 帳款處理情形修改記錄
Dim bolReadAfterSave As Boolean 'Added by Morgan 2024/10/18
Dim m_DefColor As Long, m_SetColor As Long  'Add by Amy 2025/02/20

Private Sub FormClear()
   For Each oText In txtName
      oText = ""
   Next
   For Each oText In txtBox
      oText = ""
   Next
   For Each oText In txtInform
      oText = ""
   Next
   
   For Each oCheck In Check1 'Add by Amy 2013/10/18 + 帳款處理情形
      oCheck.Value = 0
   Next
   'Add by Amy 2014/09/23 +境外公司
   Check2(0).Value = 0
   Check2(0).Tag = ""
   Text1.Text = "" 'Add By Sindy 2014/10/15 +會計備註
   'Add By Sindy 2016/11/8
   'Modify by Amy 2025/02/20 原Check3=每月代填繳款書,改可勾智慧所/法律所
   Check3(0).Value = 0: TextCU168.Visible = False
   Check3(1).Value = 0
   Check3(0).BackColor = m_DefColor: lblCU168(0).BackColor = m_DefColor
   Check3(1).BackColor = m_DefColor: lblCU168(1).BackColor = m_DefColor
   'end 2025/02/20
   chkA4228(0).Enabled = False: chkA4228(1).Enabled = False 'Add By Sindy 2019/12/18
   textCU169.Text = ""
   textCU170.Text = ""
   textCU171.Text = ""
   '2016/11/8 END
   textCU16.Text = "" 'Add By Sindy 2017/1/6
   'Add by Amy 2019/07/23 零稅率
   Check2(1).Value = 0
   Check2(1).Tag = ""
   'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
'   'Add 2019/07/25 電子發票寄送方式
'   For Each oCheck In Check4
'      oCheck.Value = 0
'      oCheck.Tag = ""
'   Next
   
   Frmacc0000.StatusBar1.Panels(2).Text = ""
   cmdA49.Visible = False 'Add By Sindy 2016/11/1
   
   'Added by Lydia 2023/11/13
   txtCU196.Text = "": txtCU196.Tag = ""
   Combo4.ListIndex = 0
End Sub

'Add by Amy 2013/10/18 帳款處理情形擇一選擇
Private Sub Check1_Click(Index As Integer)
    Dim oCheck As CheckBox
    If Check1(Index).Value = vbChecked Then
      For Each oCheck In Check1
          If oCheck.Index <> Index Then
              oCheck.Value = 0
          End If
      Next
    End If
End Sub
'end 2013/10/18

'Add By Sindy 2016/3/21 由未勾選改為有勾選存檔時,若當年已有扣繳資料, 顯示訊息, 但仍可操作
Private Sub Check2_Click(Index As Integer)

If bolReadAfterSave Then Exit Sub 'Added by Morgan 2024/10/18 存檔後讀取資料時不要檢查否則會誤解且與實際資料不符

'Modify by Amy 2019/07/22 +零稅率
If Index = 1 Then
   If Check2(1).Enabled = False Then Exit Sub 'Add By Sindy 2019/11/18
   If strSaveConfirm = MsgText(601) Then Exit Sub
   '有統編不可勾選零稅率
   If Val(Check2(1).Tag) <> Val(Check2(1).Value) Then
       If Check2(1).Value = vbChecked And Len(m_CU11) > 0 Then
           Check2(1).Value = 0
           MsgBox "有統一編號不可勾選零稅率", , MsgText(5)
           Exit Sub
       End If
   End If
'境外公司
Else
   If Check2(0).Enabled = False Then Exit Sub 'Add By Sindy 2019/11/18
   If Val(Check2(0).Tag) = 0 And Check2(0).Value = 1 Then
      strExc(0) = "select a0k01 from acc0k0 where a0k04='" & txtName(0) & "' and a0k05<>'1' and nvl(a0k16,0)=" & Left(strSrvDate(2), 3) & " and nvl(a0k09,0)=0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("此客戶 " & txtName(0) & " 今年有(可扣繳)資料, 確定是境外公司嗎？", vbYesNo + vbDefaultButton1 + vbExclamation) = vbNo Then
            Check2(0).Value = 0
         End If
      End If
   End If
End If
End Sub

'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
''Add by Amy 2019/07/25
'Private Sub Check4_Click(Index As Integer)
'    If strSaveConfirm = MsgText(601) Then Exit Sub
'
'    If Check4(0).Value = vbChecked And Check4(1).Value = vbChecked Then
'        MsgBox "電子發票寄送方式只能擇一選擇", , MsgText(5)
'    End If
'End Sub

'Add by Amy 2020/10/06
Private Sub Check5_Click()
    If Check5.Value = vbChecked And InStr(Text1, "特殊收據") = 0 Then
        Text1 = "特殊收據：" & ";" & Text1
    End If
End Sub

'Add By Sindy 2020/11/24
Private Sub chkA4228_Click(Index As Integer)
   If Index = 0 Then
      If chkA4228(Index).Value = 1 Then
         chkA4228(1).Value = 0
      End If
   ElseIf Index = 1 Then
      If chkA4228(Index).Value = 1 Then
         chkA4228(0).Value = 0
      End If
   End If
End Sub

'Add By Sindy 2016/11/1
Private Sub cmdA49_Click()
   'Add by Amy 2019/07/24
   Dim i As Integer
   Dim strCusName As String
      
   Frmacc21v0.Hide
   Frmacc21v0.textA4901.Visible = False
   Frmacc21v0.textA4901_C.Visible = True
   Frmacc21v0.textA4901_C.Text = txtKey
   'Modify by Amy 2019/07/24 客戶名稱 改為下拉選單
   'Frmacc21v0.lblA4901_C.Caption = IIf(txtName(0).Text <> "", txtName(0).Text, IIf(txtName(1).Text <> "", txtName(1).Text, txtName(2).Text))
   For i = 0 To CboCusName.ListCount - 1
        If Mid(CboCusName.List(i), 5) <> MsgText(601) Then
            strCusName = Mid(CboCusName.List(i), 5)
            Exit For
        End If
   Next i
   Frmacc21v0.lblA4901_C.Caption = strCusName
   'end 2019/07/24
   Frmacc21v0.Tag = Me.Name 'Add By Sindy 2024/9/4
   Frmacc21v0.OpenTable
   Frmacc21v0.Show vbModal
   'Add By Sindy 2016/11/8 有資料按鈕變顏色
   cmdA49.BackColor = &H8000000F
   strExc(0) = "select A4901 from ACC490 where A4901='" & txtKey & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      cmdA49.BackColor = &HC0FFC0
   End If
   '2016/11/8 END
End Sub

'Modify By Sindy 2014/9/26 因Frmacc44t0會呼叫此Form所以改為Public
'Private Sub Command1_Click()
Public Sub Command1_Click()
'2014/9/26 END
   cmdA49.Visible = False 'Add By Sindy 2016/11/1
   If txtKey = "" Then
      MsgBox "請輸入欲查詢編號！"
   ElseIf GetRec(txtKey) = True Then
      'Add by Amy 2014/09/23 +境外公司
      If Left(txtKey, 1) = "X" Then
          Check2(0).Visible = True
          'Add by Amy 2019/07/23 零稅率
          Check2(1).Visible = True
          Lbl_Inf.Visible = True
          'Add 2019/07/25 發票寄送選項
          'Memo by Lydia 2023/11/13 +請款匯入銀行資料CU196
          Frame1.Visible = True
      Else
          Check2(0).Visible = False
          'Add by Amy 2019/07/23 零稅率
          Check2(1).Visible = False
          Lbl_Inf.Visible = False
          'Add 2019/07/25 發票寄送選項
          Frame1.Visible = False
      End If
      FormShow
   Else
      MsgBox "查無資料！"
   End If
End Sub

Private Sub Form_Activate()
    strFormName = Name
End Sub

'Added by Lydia 2021/12/02
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Added by Lydia 2021/12/02 Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/12/02 Form2.0 記錄鍵盤傳入順序
   
   If KeyCode = vbKeyF12 Then
      If Command1.Enabled = True Then
         Command1_Click
      End If
   Else
      KeyEnter KeyCode
   End If
End Sub

Private Sub Form_Load()
   '表單初始化
   'Modified by Lydia 2018/07/20
   'PUB_InitForm Me, 9045, 5700, strBackPicPath1 'Modify by Amy 2013/10/18 原5130 改5430
   'Modify by Amy 2023/08/18 W9045 H6090
   'Modified by Lydia 2024/09/18 9210, 6260
   PUB_InitForm Me, 9210, 6440, strBackPicPath1
   tool6_enabled
   SetCheck1 (False) 'Add by Amy 2014/04/03
   Check2(0).Visible = False 'Add by Amy 2014/09/23 +境外公司
   'add by sonia 2014/11/14 分所只能輸財務信箱及會計備註,分所只能改該所資料故取消前後筆功能
   If pub_strUserOffice <> "1" Then
      Me.Caption = "客戶財務EMail資料維護"
      Label19.Caption = "客戶編號"
      txtInform(0).Enabled = False
      txtInform(1).Enabled = False
      txtInform(2).Enabled = False
      txtInform(3).Enabled = False 'Add By Sindy 2017/3/16
      txtInform(4).Enabled = False 'Add By Sindy 2017/3/24
      Check1(0).Enabled = False
      Check1(1).Enabled = False
      Check1(2).Enabled = False
      Check1(3).Enabled = False
      'modify BY SONIA 2016/12/20 瑞婷說境外公司開給分所維護,但是否寄發收據,是否寄發電匯通知,是否寄發催款單,每月提醒代填繳款書四欄不開放
      'Check2(0).Enabled = False
      'txtInform(0).Enabled = False
      'txtInform(1).Enabled = False
      'txtInform(2).Enabled = False
      'Modify by Amy 2025/02/20 原Check3=每月代填繳款書,改可勾智慧所/法律所
      Check3(0).Enabled = False
      Check3(1).Enabled = False
      'end 2025/02/20
      chkA4228(0).Enabled = False: chkA4228(1).Enabled = False 'Add By Sindy 2019/12/18
      'end 2016/12/20
      'Remove by Lydia 2016/12/26 改為分所人員可查詢所有客戶資料, 但只可修改該所資料
      'tool3_enabled                   '取消前後筆功能
   End If
   'end 2014/11/14
   
   Call SetCombo4 'Added by Lydia 2023/11/13 請款匯入銀行資料預設清單
   'Add by Amy 2025/02/20 有同意書檔變色
   m_DefColor = &H8000000F
   m_SetColor = RGB(215, 117, 117)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
       Cancel = 1
       Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   strTrackMode = "" 'Added by Lydia 2021/12/02 Form2.0 記錄鍵盤傳入順序(清除)
   KeyEnter vbKeyEscape
   MenuEnabled
   
   'Add By Sindy 2014/9/26
   If UCase(strUserLevel) = UCase("Frmacc44t0") Then
      Frmacc44t0.Show
      Frmacc44t0.cmdQuery_Click
      tool3_enabled
   ElseIf TypeName(m_PrevForm) <> "Nothing" Then
      m_PrevForm.Show
      tool3_enabled
   End If
   '2014/9/26 END
   
   strUserLevel = MsgText(601) 'Add By Sindy 2015/12/10
   Set m_PrevForm = Nothing 'Add By Sindy 2016/11/29
   Set Frmacc21r0 = Nothing
End Sub

'Add by Amy 2020/09/18 不寄催款單說明
Private Sub Lbl_fa101_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Lbl_fa101.ToolTipText = "1.每月寄對帳單" & _
                                            "2.客戶要求不寄對帳單" & _
                                            "3.其他"
End Sub

'Add by Amy 2019/07/23  零稅率說明
Private Sub Lbl_Inf_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     Lbl_Inf.ToolTipText = "符合條件：" & _
                        "(1)國外匯入款　(2)國外公司　(3)有簽約　就適用零稅率。" & _
                        "是否符合條件主要在於款項是否由國外匯入，故僅能在收到款項後設定。"
End Sub

'Add by Amy 2019/08/01
Private Sub Lbl_Inf1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Lbl_Inf1.ToolTipText = "勾選宣告破產會加註於客戶/代理人備註欄。"
End Sub

'Add by Amy 2019/07/25 財務信箱說明
Private Sub Lbl_InfMail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Modify By Sindy 2025/9/5 +、「不使用信箱」
    Lbl_InfMail.ToolTipText = "Key 「NO」不寄「收據」、「付款明細」、「催款單」、「不使用信箱」。"
End Sub

Private Sub lblCU168_Click(Index As Integer)
   Dim stComp As String, stMsg As String
   
   stComp = "1"
   If Index = 1 Then stComp = "L"
   If lblCU168(Index).BackColor = m_SetColor Then
      'Modify by Amy 2025/11/03 +txtNameNoUni,避免檔名有UniCode字無法開啟
      If ChkWithholdingTaxConsent(1, Me.Name, stComp, txtName(0), File1, stMsg, txtNameNoUni) = False Then
         MsgBox "檔案開啟有誤！" & vbCrLf & "請洽電腦中心！" & vbCrLf & _
                          "(錯誤:" & stMsg & ")"
      End If
   End If
End Sub

'Add By Sindy 2014/10/15
Private Sub Text1_GotFocus()
   InverseTextBox Text1
   OpenIme
End Sub
Private Sub Text1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(Text1, Text1.MaxLength) = False Then
      Cancel = True
      Text1_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub
'2014/10/15 END

Private Sub txtBox_GotFocus(Index As Integer)
   TextInverse txtBox(Index)
End Sub

Private Sub txtInform_GotFocus(Index As Integer)
   TextInverse txtInform(Index)
End Sub

Private Sub txtInform_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Moragn 2015/6/12
   If Index = 0 Then
      If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("B") Then
         KeyAscii = 0
         Beep
      End If
   'end 2015/6/12
   'Add By Sindy 2017/3/24
   ElseIf Index = 4 Then
      If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
         KeyAscii = 0
         Beep
      End If
   'Add by Amy 2020/09/18 不寄催款單 改輸1-3
   ElseIf Index = 2 Then
      If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
         KeyAscii = 0
         Beep
      End If
   '2017/3/24 END
   Else
      If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub txtKey_Change()
   Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = False
End Sub

Private Sub txtKey_GotFocus()
   TextInverse txtKey
   If Command1.Enabled = True Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   End If
   CloseIme
End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtKey_Validate(Cancel As Boolean)
   If Len(txtKey) = 6 Then
      txtKey = txtKey & "000"
   End If
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
    
On Error GoTo ErrorHandler
    If ado21r0.RecordCount = 0 Then
        Exit Sub
    End If
    CountShow ado21r0.AbsolutePosition, ado21r0.RecordCount
Exit Sub
ErrorHandler:
    MsgBox Err.Description
    
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   With ado21r0
      txtName(0) = "" & .Fields("name1")
      txtNameNoUni = "" & .Fields("name1")
      'Modify by Amy 2019/07/24 客戶名稱 改成下拉選單
'      txtName(1) = "" & .Fields("name2")
'      txtName(2) = "" & .Fields("name3")
      CboCusName.Clear
      CboCusName.AddItem "(中):" & "" & .Fields("name1")
      CboCusName.AddItem "(英):" & "" & .Fields("name2")
      CboCusName.AddItem "(日):" & "" & .Fields("name3")
      If "" & .Fields("name1") <> MsgText(601) Then
        CboCusName = "(中):" & "" & .Fields("name1")
      ElseIf "" & .Fields("name2") <> MsgText(601) Then
        CboCusName = "(英):" & "" & .Fields("name2")
      Else
        CboCusName = "(日):" & "" & .Fields("name3")
      End If
      'end 2019/07/24
      txtBox(0) = "" & .Fields("box1")
      txtBox(1) = "" & .Fields("box2")
      'Modify by Amy 2019/07/24 其他信箱 改為下拉選單
'      txtBox(2) = "" & .Fields("box3")
'      txtBox(3) = "" & .Fields("box4")
'      txtBox(4) = "" & .Fields("box5")
      CboBox.Clear
      CboBox.AddItem "1:" & .Fields("box3")
      CboBox.AddItem "2:" & .Fields("box4")
      CboBox.AddItem "3:" & .Fields("box5")
      If "" & .Fields("box3") <> MsgText(601) Then
        CboBox = "1:" & .Fields("box3")
      ElseIf "" & .Fields("box4") <> MsgText(601) Then
        CboBox = "2:" & .Fields("box4")
      ElseIf "" & .Fields("box5") <> MsgText(601) Then
        CboBox = "3:" & .Fields("box5")
      Else
        CboBox = "1:" & .Fields("box3")
      End If
      'end 2019/07/24
      txtBox(5) = "" & .Fields("box6") 'Added by Lydia 2018/07/20 財務信箱(CF)
      txtBox(2) = "" & .Fields("box7") 'Added by Lydia 2024/09/18 財務副本信箱
      txtKey = "" & .Fields("No")
      'Add by Morgan 2007/3/3
      txtInform(0) = "" & .Fields("inf1")
      txtInform(1) = "" & .Fields("inf2")
      'end 2007/3/3
      txtInform(2) = "" & .Fields("inf3") 'Add by Sindy 2010/7/12
      txtInform(3) = "" & .Fields("inf7") 'Add By Sindy 2017/3/16
      txtInform(4) = "" & .Fields("inf8") 'Add By Sindy 2017/3/24
      
      'Add by Amy 2013/10/18 +帳款處理情形
      For Each oCheck In Check1
        oCheck.Value = 0
      Next
       oldDizhang = "" & .Fields("inf4")
       
      strDizhangRec = "" 'Added by Lydia 2017/01/16
      Select Case oldDizhang
        Case "A"
            Check1(0).Value = vbChecked
            strDizhangRec = Check1(0).Caption 'Added by Lydia 2017/01/16
        Case "B"
            Check1(2).Value = vbChecked
            strDizhangRec = Check1(2).Caption 'Added by Lydia 2017/01/16
        Case "C"
            Check1(1).Value = vbChecked
            strDizhangRec = Check1(1).Caption 'Added by Lydia 2017/01/16
        'Add by Amy 2014/04/03
        Case "D"
            Check1(3).Value = vbChecked
            strDizhangRec = Check1(3).Caption 'Added by Lydia 2017/01/16
      End Select
      'end 2013/10/18
      'Add by Amy 2014/09/23 +境外公司
      If Left(txtKey, 1) = "X" And Not IsNull(.Fields("inf5")) Then
           Check2(0).Value = vbChecked
      Else
           Check2(0).Value = 0
      End If
      Check2(0).Tag = Check2(0).Value
      'end 2014/09/23
      m_CU11 = "" & .Fields("cu11") 'Add By Sindy 2015/12/11
      Text1.Text = "" & .Fields("inf6") 'Add by Sindy 2014/10/15
      
      'Add By Sindy 2016/11/8
      'Modify by Amy 2025/02/20 原Check3=每月代填繳款書,改可勾智慧所/法律所,且已有同意書檔變色
      Check3(0).BackColor = m_DefColor: lblCU168(0).BackColor = m_DefColor
      Check3(1).BackColor = m_DefColor: lblCU168(1).BackColor = m_DefColor
      If IsNull(.Fields("cu168").Value) Then
         Check3(0).Value = 0: TextCU168.Visible = False
         Check3(1).Value = 0
      Else
         If InStr("," & .Fields("cu168"), ",1") > 0 Then
            Check3(0).Value = 1
            If ChkWithholdingTaxConsent(0, Me.Name, "1", txtName(0)) = True Then
               Check3(0).BackColor = m_SetColor: lblCU168(0).BackColor = m_SetColor
            End If
         End If
         If InStr("," & .Fields("cu168"), ",L") > 0 Then
            Check3(1).Value = 1
            If ChkWithholdingTaxConsent(0, Me.Name, "L", txtName(0)) = True Then
               Check3(1).BackColor = m_SetColor: lblCU168(1).BackColor = m_SetColor
            End If
         End If
         TextCU168.Visible = True
      End If
      'end 2025/02/20
      'Add By Sindy 2019/12/18
      If IsNull(.Fields("cu181").Value) Then
         chkA4228(0).Value = 0: chkA4228(1).Value = 0
      Else
         If .Fields("cu181").Value = "1" Then
            chkA4228(0).Value = 1
         ElseIf .Fields("cu181").Value = "2" Then
            chkA4228(1).Value = 1
         End If
      End If
      '2019/12/18 END
      
      textCU169 = "" & .Fields("cu169")
      textCU170 = "" & .Fields("cu170")
      textCU171 = "" & .Fields("cu171")
      '2016/11/8 END
      textCU16.Text = "" & .Fields("cu16") 'Add By Sindy 2017/1/6
      'Add by Amy 2019/07/23 零稅率
      Check2(1).Value = 0
      Check2(1).Tag = ""
      If "" & .Fields("cu178") <> MsgText(601) Then
            Check2(1).Value = vbChecked
      End If
      Check2(1).Tag = Check2(1).Value
      'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
'      'Add 2019/07/25 電子發票寄送方式 null-未設定/1-紙本/2-電子檔
'      If "" & .Fields("cu179") <> MsgText(601) Then
'            Check4(Val(.Fields("cu179")) - 1).Value = vbChecked
'      End If
'      Check4(0).Tag = "" & .Fields("cu179")
'      'end 2019/07/25
      'end 2024/06/13
      'Add by Amy 2020/10/06 +特殊收據
      Check5.Value = 0
      If "" & .Fields("cu184") <> MsgText(601) Then
            Check5.Value = vbChecked
      End If
      'end 2020/10/06
      
      'Added by Morgan 2024/5/28 不索取CF對帳單
      Check6.Value = vbUnchecked
      If Left(txtKey, 1) = "Y" Then
         Check6.Visible = True
         If "" & .Fields("FA133") = "N" Then
            Check6.Value = vbChecked
         End If
      Else
         Check6.Visible = False
      'end 2024/5/28
      End If
      'end 2024/5/28
      
      'Added by Lydia 2023/11/13 請款匯入銀行資料
      txtCU196 = "" & .Fields("cu196")
      Call SetCombo4(IIf(txtCU196.Text = "", "0", txtCU196.Text))
      txtCU196.Tag = txtCU196.Text
      'end 2023/11/13
      
      txtKey.Tag = txtKey
      
      'Add By Sindy 2016/11/8 有資料按鈕變顏色
      cmdA49.BackColor = &H8000000F
      strExc(0) = "select A4901 from ACC490 where A4901='" & txtKey & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         cmdA49.BackColor = &HC0FFC0
      End If
      '2016/11/8 END
      
      'Modified by Lydia 2016/12/26 改為分所人員可查詢所有客戶資料, 但只可修改該所資料
      'Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
      If bolMod = True Then
         Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
      Else
         Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = False
      End If
      'end 2016/12/26
      
      RecordShow
   End With
End Sub

Private Function GetRec(ByVal p_Key As String, Optional p_Way As Integer = 0) As Boolean
   Dim stCompSign As String, stSort As String
   FormClear
   'Modified by Lydia 2024/09/18
   'If Len(p_Key) = 6 Then
   '   p_Key = p_Key & "000"
   'End If
   p_Key = ChangeCustomerL(p_Key)
   'end 2024/09/18
   If p_Way = 0 Then
      stCompSign = "="
   ElseIf p_Way > 0 Then
      stCompSign = ">"
      stSort = "asc"
   Else
      stCompSign = "<"
      stSort = "desc"
   End If
   'Modify by Morgan 2007/3/3 加cu119,cu120,fa83,fa84
   'Modify by Sindy 2010/7/12 加fa101,cu140
   'Modify by Amy 2013/10/18 +fa103,cu142
   'Modify by Amy 2014/09/23 +cu158
   'Modify By Sindy 2014/10/15 +CU159
   'Modify By Sindy 2017/1/6 +cu16
   'Modify by Amy 2019/07/23 +cu178 零稅率
   'Modify by Amy 2019/07/25 +cu179 電子發票寄送方式
   'Modify by Amy 2020/10/06 +特殊收據 fa125/cu184
   
   If Left(p_Key, 1) = "X" Then
      'modify by sonia 2014/11/14 加業務所別
      'Modify By sindy 2015/12/11 + cu11
      'Modify By Sindy 2016/11/8 + cu168,cu169,cu170,cu171
      'Modify By Sindy 2017/3/16 + CU172 inf7
      'Modify By Sindy 2017/3/24 + CU173 inf8
      'Modified by Lydia 2018/07/20 +box6
      'strExc(0) = "select cu01||cu02 No,cu04 name1,cu05||cu88||cu89||cu90 name2,cu06 name3,cu20 box1,cu115 box2,cu116 box3,cu117 box4,cu118 box5,CU119 inf1,CU120 inf2,CU140 inf3,CU172 inf7,CU173 inf8,Nvl(CU142,'') inf4,cu158 inf5,cu159 inf6,st06,cu11,cu168,cu169,cu170,cu171,cu16" & _
         " from customer,staff where cu01||cu02" & stCompSign & "'" & p_Key & "' and cu13=st01(+) "
      'Modified by Lydia 2023/11/13 +cu196
      'Modified by Lydia 2024/09/18 +cu200=box7
      'Modify By Sindy 2025/9/19 cu16 => nvl(cu16,nvl(cu17,cu22)) 客戶電話1優先,再客戶電話2,最後為手機
      strExc(0) = "select cu01||cu02 No,cu04 name1,cu05||cu88||cu89||cu90 name2,cu06 name3," & _
                        "cu20 box1,cu115 box2,cu116 box3,cu117 box4,cu118 box5, null box6," & _
                        "CU119 inf1,CU120 inf2,CU140 inf3,CU172 inf7,CU173 inf8,Nvl(CU142,'') inf4,cu158 inf5,cu159 inf6," & _
                        "st06,cu11,cu168,cu169,cu170,cu171,nvl(cu16,nvl(cu17,cu22)) cu16,cu178,cu179,cu181,cu184,cu196" & _
                        ",cu200 as box7 from customer,staff where cu01||cu02" & stCompSign & "'" & p_Key & "' and cu13=st01(+) "
   Else
      'Modify By Sindy 2014/10/15 +FA118
      'Modify By Sindy 2016/11/8 + '' cu168,'' cu169,'' cu170,'' cu171
      'Modified by Lydia 2017/01/16 原cu16 -> fa12 cu16
      'Modify By Sindy 2017/3/16 + '' inf7
      'Modify By Sindy 2017/3/24 + '' inf8
      'Modified by Lydia 2018/07/20 +box6
      'strExc(0) = "select fa01||fa02 No,fa04 name1,fa05||fa63||fa64||fa65 name2,fa06 name3,fa16 box1,fa79 box2,fa80 box3,fa81 box4,fa82 box5,FA83 inf1,FA84 inf2,FA101 inf3,'' inf7,'' inf8,Nvl(FA103,'') inf4,'' inf5,FA118 inf6,'' cu11,'' cu168,'' cu169,'' cu170,'' cu171,fa12 cu16" & _
         " from fagent where fa01||fa02" & stCompSign & "'" & p_Key & "'"
      'Modified by Lydia 2023/11/13 +cu196
      'Modified by Morgan 2024/5/28 +FA133
      'Modified by Lydia 2024/09/18 +FA134=box7
      'Modify By Sindy 2025/9/19 fa12 cu16 => nvl(fa12,fa13) 客戶電話1優先,再客戶電話2,最後為手機
      strExc(0) = "select fa01||fa02 No,fa04 name1,fa05||fa63||fa64||fa65 name2,fa06 name3," & _
                       "fa16 box1,fa79 box2,fa80 box3,fa81 box4,fa82 box5,fa105 box6," & _
                       "FA83 inf1,FA84 inf2,FA101 inf3,'' inf7,'' inf8,Nvl(FA103,'') inf4,'' inf5,FA118 inf6," & _
                       "'' cu11,'' cu168,'' cu169,'' cu170,'' cu171,nvl(fa12,fa13) cu16,'' cu178,'' cu179,'' cu181,fa125 cu184,'' as cu196,fa133" & _
                       ",fa134 as box7 from fagent where fa01||fa02" & stCompSign & "'" & p_Key & "'"
   End If
   'end 2007/3/3
   strExc(0) = strExc(0) & " order by 1 " & stSort
   If ado21r0.State = adStateOpen Then
      ado21r0.Close
   End If
   '設定回傳筆數
   ado21r0.CursorLocation = adUseClient
   ado21r0.MaxRecords = 50
   ado21r0.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   If ado21r0.RecordCount > 0 Then
      'Add By Sindy 2016/11/1
      If Left(Trim(ado21r0.Fields("No")), 1) = "X" Then
         cmdA49.Visible = True
      End If
      '2016/11/1 END
      
      bolMod = True 'Added by Lydia 2016/12/26
      
      'add by sonia 2014/11/14 分所操作判斷所別
      If pub_strUserOffice <> "1" Then
         If pub_strUserOffice <> ado21r0.Fields("st06") Then
            'MsgBox "所別不同, 不可操作！"
            'Modified by Lydia 2016/12/26 改為分所人員可查詢所有客戶資料, 但只可修改該所資料
            'Exit Function
            bolMod = False
            MsgBox "非" & IIf(pub_strUserOffice = "2", "中所", IIf(pub_strUserOffice = "3", "南所", IIf(pub_strUserOffice = "4", "高所", "分所"))) & "資料, 請通知北所修改！", vbExclamation
            'end 2016/12/26
         End If
      End If
      'end 2014/11/14
      GetRec = True
      If p_Way < 0 Then
         ado21r0.Sort = "No asc"
         ado21r0.MoveLast
      End If
   End If
End Function

Public Sub MoveNext()
   If txtKey.Tag = "" Then
      MsgBox "尚未有查詢紀錄，無法搜尋上下筆！"
      Exit Sub
   End If
   With ado21r0
   If .State = adStateOpen Then
      If Not .EOF Then
         .MoveNext
      End If
      If .EOF Then
         GetRec txtKey.Tag, 1
      End If
      If .EOF Then
         MsgBox MsgText(8), , MsgText(5)
         MoveLast
      Else
         FormShow
      End If
   End If
   End With
End Sub

Public Sub MovePrevious()
   If txtKey.Tag = "" Then
      MsgBox "尚未有查詢紀錄，無法搜尋上下筆！"
      Exit Sub
   End If
   With ado21r0
   If .State = adStateOpen Then
      If Not .BOF Then
         .MovePrevious
      End If
      If .BOF Then
         GetRec txtKey.Tag, -1
      End If
      If .BOF Then
         MsgBox MsgText(7), , MsgText(5)
         MoveFirst
      Else
         FormShow
      End If
   End If
   End With
End Sub

Public Sub MoveFirst()
   Dim strKey
   If txtKey.Tag = "" Then
      strKey = "Y"
   Else
      strKey = Left(txtKey.Tag, 1)
   End If
   With ado21r0
      If GetRec(strKey, 1) = True Then
         FormShow
      Else
         MsgBox "查無資料！"
      End If
   End With
End Sub

Public Sub MoveLast()
   Dim strKey
   If txtKey.Tag = "" Then
      strKey = "Y"
   Else
      strKey = Left(txtKey.Tag, 1)
   End If
   strKey = strKey & "Z"
   With ado21r0
      If GetRec(strKey, -1) = True Then
         FormShow
      Else
         MsgBox "查無資料！"
      End If
   End With
End Sub

Public Function FormSave() As Boolean
Dim lngEff As Long
'Add by Amy 2013/10/18
Dim strDizhang As String '帳款處理代號 A.同意抵帳 B.帳款處理中 C.宣告破產 D.不同意抵帳
Dim oCheck As CheckBox
Dim StrSQLa As String
Dim bolSendMail As Boolean, strNo As String 'Add By Sindy 2016/3/21
Dim Cancel As Boolean 'Add By Sindy 2016/11/8
Dim strCU168 As String 'Add by Amy 2025/02/20

'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
'Dim strCU179 As String 'Add by Amy 2019/07/25 電子發票寄送方式 null-未設定/1-紙本/2電子檔
   
    'Added by Lydia 2021/12/02 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
    If PUB_ChkTrackMode = False Then
        Exit Function
    End If
    'end 2021/12/02
    
   'Added by Lydia 2018/07/20 txtBox改成迴圈
   For Each oText In txtBox
     'If oText.Index = 1 Or oText.Index = 0 Then 'Mark by Lydia 2024/09/18
         If StrLength(oText) > oText.MaxLength Then
            MsgBox "長度不可超過" & oText.MaxLength & "！"
            Exit Function
         'Modified by Morgan 2013/8/6
         'ElseIf otext <> "" And InStr(otext, "@") = 0 Then
         '   'Modify by Morgan 2009/7/7 不檢查,可放說明
         '   If UCase(otext) = "NO" Then
         '      otext = "NO"
         '   Else
         '      MsgBox "格式錯誤！"
         '      Exit Function
         '   End If
         ElseIf oText <> "" Then
            If UCase(oText) = "NO" Then
               oText = "NO"
            ElseIf PUB_CheckMail(oText.Text) = False Then
               txtBox_GotFocus oText.Index
               oText.SetFocus
               Exit Function
            End If
         'end 2013/8/6
         End If
     'End If 'Mark by Lydia 2024/09/18
   Next

   'Add By Sindy 2016/11/8　繳款書寄件處不可空白
   If Left(txtKey, 1) = "X" Then
      'Add by Amy 2019/07/23 零稅率
      If Val(Check2(1).Tag) <> Val(Check2(1).Value) Then
        If Check2(1).Value = vbChecked And Len(m_CU11) > 0 Then
            Check2(1).Value = 0
            MsgBox "有統一編號不可勾選零稅率", , MsgText(5)
            Exit Function
        End If
      End If
      'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
'      'Add by Amy 2019/07/25 電子發票寄送方式
'      If Check4(0).Value = vbChecked And Check4(1).Value = vbChecked Then
'         Check4(0).SetFocus
'         MsgBox "發票寄送方式只能擇一選擇", , MsgText(5)
'         Exit Function
'      End If
      If textCU169 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         textCU169.SetFocus
         Exit Function
      Else
         If textCU169 = "3" Then '選擇3特殊時，特殊地址和收件人要同時有值
            If textCU170 = MsgText(601) Or textCU171 = MsgText(601) Then
               MsgBox MsgText(10), , MsgText(5)
               If textCU171 = MsgText(601) Then
                  textCU171.SetFocus
               ElseIf textCU170 = MsgText(601) Then
                  textCU170.SetFocus
               End If
               Exit Function
            End If
         Else
            If textCU170 <> MsgText(601) Or textCU171 <> MsgText(601) Then
               MsgBox "繳款書寄件處非選特殊，不需輸入特殊地址和收件人", , MsgText(5)
               If textCU171 <> MsgText(601) Then
                  textCU171.SetFocus
               ElseIf textCU170 <> MsgText(601) Then
                  textCU170.SetFocus
               End If
               Exit Function
            End If
         End If
      End If
      Call textCU169_Validate(Cancel)
      If Cancel = True Then
         textCU169.SetFocus
         Exit Function
      End If
      Call textCU171_Validate(Cancel)
      If Cancel = True Then
         textCU171.SetFocus
         Exit Function
      End If
      Call textCU170_Validate(Cancel)
      If Cancel = True Then
         textCU170.SetFocus
         Exit Function
      End If
      'Added by Lydia 2023/11/13 請款匯入銀行資料
      Call SetCombo4(IIf(txtCU196.Text = "", "0", txtCU196.Text))
      If txtCU196.Tag <> "" And txtCU196.Text = "" Then
         If MsgBox("取消請款匯入銀行資料設定．是否繼續存檔？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            txtCU196.SetFocus
            txtCU196_GotFocus
            Exit Function
         End If
      End If
      'end 2023/11/13
   End If
   '2016/11/8 END
   
   'Add By Sindy 2019/12/18 無每月提醒代填繳款書,代填方式則不須點選
   'Modify by Amy 2025/02/20 原Check3=每月代填繳款書,改可勾智慧所/法律所
   If Check3(0).Value = 0 And Check3(1).Value = 0 Then
      chkA4228(0).Value = 0
      chkA4228(1).Value = 0
   End If
   'Add by Amy 2020/10/06 +特殊收據
   If Check5.Value = vbChecked Then
        If InStr(Text1, "特殊收據") = 0 Or InStr(Text1, "特殊收據：;") > 0 Then
            MsgBox "勾選「特殊收據」需寫特殊收據內容"
            Text1.SetFocus
            Exit Function
        End If
   Else
        If InStr(Text1, "特殊收據") > 0 Then
            MsgBox "未勾選「特殊收據」請將內容拿掉"
            Text1.SetFocus
            Exit Function
        End If
   End If
   
   'Added by Lydia 2024/09/18
   'Modified by Lydia 2025/03/05 不一定有財務正本信箱
   'If Trim(txtBox(2)) <> "" And Trim(txtBox(1) & txtBox(5)) = "" Then
   '   MsgBox "財務信箱或財務信箱(CF)不可空白！"
   '   txtBox(1).SetFocus
   '   txtBox_GotFocus 1
   '   Exit Function
   'End If
   ''end 2024/09/18
   'end 2025/03/05
   
   'Added by Lydia 2021/12/02 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
   'Add by Amy 2023/06/30 避免存檔時,未Run到地址欄位_Validate未轉全型,故再轉一次
   If Trim(textCU170) <> MsgText(601) Then textCU170 = PUB_ChangeZIPToSir(textCU170) '特殊地址
    
   'Add by Amy 2013/10/18 帳款處理代號
   strDizhang = "": strExc(1) = "": StrSQLa = ""
   For Each oCheck In Check1
       If oCheck.Value = vbChecked Then
           strDizhang = oCheck.Index
           strDizhangRec = strDizhangRec & "->" & oCheck.Caption 'Added by Lydia 2017/01/16
           Exit For
       End If
   Next
   Select Case strDizhang
        Case "0"
            strDizhang = "A"
        Case "1"
            strDizhang = "C"
        Case "2"
            strDizhang = "B"
        'Add by Amy 2014/04/03 +不同意抵帳
        Case "3"
            strDizhang = "D"
        Case Else
            strDizhang = ""
            strDizhangRec = strDizhangRec & "->" 'Added by Lydia 2017/01/16
   End Select
   If oldDizhang <> strDizhang Then
        '更新編號前8碼相同之fa103,cu142
        If strDizhang = "B" Then
            '選「宣告破產」需更新 fa29/cu79 (備註)、fa69/cu80 (狀態)、fa77/cu111 (呆帳記錄)
            If Left(txtKey, 1) = "Y" Then
                 StrSQLa = ",FA29='" & strSrvDate(1) & "財務處加註宣告破產;'||FA29, FA69='宣告破產',FA77='Y' "
            Else
                StrSQLa = ",CU79='" & strSrvDate(1) & "財務處加註宣告破產;'||CU79, CU80='宣告破產',CU111='Y' "
            End If
        End If
        If Left(txtKey, 1) = "Y" Then
            strExc(1) = "Update FAgent Set FA103=" & CNULL(ChgSQL(strDizhang)) & StrSQLa & " Where FA01='" & Left(txtKey, 8) & "'"
            StrSQLa = ""
        Else
            strExc(1) = "Update Customer Set CU142=" & CNULL(ChgSQL(strDizhang)) & StrSQLa & " Where CU01='" & Left(txtKey, 8) & "'"
        End If
   End If
   'end 2013/10/18
   
   'Modify by Morgan 2007/3/3 加fa83,fa84,cu119,cu120
   'Modify by Sindy 2010/7/12 加fa101,cu140
   If Left(txtKey, 1) = "Y" Then
      'Modify By Sindy 2014/10/15 +FA118=" & CNULL(ChgSQL(Text1)) & ",
      'Modify By Sindy 2018/1/5 + and FA02='" & Mid(txtKey, 9) & "'
      'Modified by Lydia 2018/07/20
      'strExc(0) = "Update FAgent Set FA118=" & CNULL(ChgSQL(Text1)) & ",FA79=" & CNULL(ChgSQL(txtBox(1))) & ",FA83=" & CNULL(txtInform(0)) & ",FA84=" & CNULL(txtInform(1)) & ",FA101=" & CNULL(txtInform(2)) & " where FA01='" & Left(txtKey, 8) & "'"
      'Modify by Amy 2020/10/06 +特殊收據
      'Modified by Lydia 2024/09/18 +財務副本信箱FA134
      strExc(0) = ",FA125=" & CNULL(ChgSQL(IIf(Check5.Value = 1, "Y", "")))
      strExc(0) = strExc(0) & ",FA133='" & IIf(Check6.Value = vbChecked, "N", "") & "'" 'Added by Morgan 2024/5/28
      strExc(0) = "Update FAgent Set FA118=" & CNULL(ChgSQL(Text1)) & _
                        ",FA79=" & CNULL(ChgSQL(txtBox(1))) & ", FA105=" & CNULL(ChgSQL(txtBox(5))) & _
                        ",FA83=" & CNULL(ChgSQL(txtInform(0))) & ",FA84=" & CNULL(ChgSQL(txtInform(1))) & _
                        ",FA101=" & CNULL(ChgSQL(txtInform(2))) & strExc(0) & ",FA134=" & CNULL(ChgSQL(txtBox(2))) & " where FA01='" & Left(txtKey, 8) & "'"
      'end 2020/10/06
   Else
      'Modify By Sindy 2014/10/15 +CU159=" & CNULL(ChgSQL(Text1)) & ",
      'Modify By Sindy 2017/3/16 +CU172=" & CNULL(txtInform(3))
      'Modify By Sindy 2017/3/24 +CU173=" & CNULL(txtInform(4))
      'Modified by Lydia 2024/09/18 +財務副本信箱CU200
      strExc(0) = "Update Customer Set CU159=" & CNULL(ChgSQL(Text1)) & _
                  ",CU115=" & CNULL(ChgSQL(txtBox(1))) & ",CU119=" & CNULL(txtInform(0)) & _
                  ",CU120=" & CNULL(txtInform(1)) & ",CU140=" & CNULL(txtInform(2)) & _
                  ",CU172=" & CNULL(txtInform(3)) & ",CU173=" & CNULL(txtInform(4)) & _
                  ",CU200=" & CNULL(ChgSQL(txtBox(2)))
      If Left(txtKey, 1) = "X" Then
         'Modify By Sindy 2019/12/18 +CU181=" & CNULL(IIf(optA4228(0).Value = True, "1", IIf(optA4228(1).Value = True, "2", "")))
         'Modify by Amy 2025/02/20 原Check3=每月代填繳款書(原:CNULL(IIf(Check3.Value = 1, "Y", ""))),改可勾智慧所/法律所(存公司代碼)
         If Check3(0).Value = 1 Then strCU168 = strCU168 & ",1"
         If Check3(1).Value = 1 Then strCU168 = strCU168 & ",L"
         If strCU168 <> MsgText(601) Then strCU168 = Mid(strCU168, 2)
         strExc(0) = strExc(0) & ",CU168=" & CNULL(strCU168) & _
                                 ",CU169=" & CNULL(textCU169) & _
                                 ",CU170=" & CNULL(textCU170) & _
                                 ",CU171=" & CNULL(textCU171) & _
                                 ",CU181=" & CNULL(IIf(chkA4228(0).Value = 1, "1", IIf(chkA4228(1).Value = 1, "2", "")))
          'end 2025/02/20
          'Add by Amy 2019/07/23 零稅率
          If Check2(1).Tag <> Check2(1).Value Then
                strExc(0) = strExc(0) & ",CU178=" & CNULL(ChgSQL(IIf(Check2(1).Value = 1, "Y", "")))
          End If
          'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
'          'Add by Amy 2019/07/25 電子發票寄送方式 null-未設定/1-紙本/2電子檔
'          If Check4(0).Value = vbChecked Then
'                strCU179 = "1"
'          ElseIf Check4(1).Value = vbChecked Then
'                strCU179 = "2"
'          End If
'          If Val(Check4(0).Tag) <> Val(strCU179) Then
'                strExc(0) = strExc(0) & ",CU179=" & CNULL(strCU179)
'          End If
'          strExc(0) = strExc(0) & ",CU184=" & CNULL(ChgSQL(IIf(Check5.Value = 1, "Y", "")))
'          'end 2019/07/25
          'end 2024/06/13
          'Added by Lydia 2023/11/13 請款匯入銀行資料
          strExc(0) = strExc(0) & ",CU196=" & CNULL(txtCU196)
      End If
      'Modify By Sindy 2018/1/5 + and CU02='" & Mid(txtKey, 9) & "'
      strExc(0) = strExc(0) & " where CU01='" & Left(txtKey, 8) & "'"
   End If
   'end 2007/3/3
   
   adoTaie.BeginTrans
   Pub_SeekTbLog strExc(0) 'Added by Morgan 2017/9/6
   adoTaie.Execute strExc(0), lngEff
   'Add by Amy 2013/10/18 +if
   If strExc(1) <> "" Then
        If strDizhang = "B" Then 'Modify by Amy 2013/11/1 +if 判斷只有選 「宣告破產」才寫log
            Pub_SeekTbLog strExc(1) '修改 帳款處理情形 記log
        End If
        adoTaie.Execute strExc(1)
        
        'Added by Lydia 2017/01/16 記錄帳款處理情形的變更
        strExc(0) = "insert into DizhangRecord(DR01,DR02,DR03,DR04,DR05) values ('" & ChangeCustomerL(txtKey.Text) & "','" & strDizhangRec & "','" & strUserNum & "'," & strSrvDate(1) & "," & Format(ServerTime, "000000") & ") "
        adoTaie.Execute strExc(0)
        'end 2017/01/16
   End If
   'Add by Amy 2014/09/23 +境外公司
   If Left(txtKey, 1) = "X" And Val(Check2(0).Tag) <> Check2(0).Value Then
        If Check2(0).Value = 0 Then
            strExc(1) = ""
        Else
            strExc(1) = "Y"
        End If
        strExc(0) = "Update Customer Set CU158=" & CNULL(strExc(1)) & " Where CU01='" & Left(txtKey, 8) & "'"
        Pub_SeekTbLog strExc(0) 'Added by Morgan 2017/9/6
        adoTaie.Execute strExc(0)
        'Add By Sindy 2015/12/11 更新境外公司時，讀取相同ID一併更新
        If m_CU11 <> "" Then
           'Modify by Amy 2025/09/25 與Sindy 討論後,覺得不應該於此加條件,應該於log改(再和秀玲確認如何改),故先改回
           'Modify by Amy 2025/09/24 發現更新資料 X0096400 ,不會記錄Log,但會執行更新語法
           'strExc(0) = "Update Customer Set CU158=" & CNULL(strExc(1)) & " Where CU11='" & m_CU11 & "'"
'           '因Pub_SeekTbLog中查詢 select  CU158 from Customer where CU11='35048016' ,比對到 X0096400 cu158(上面語法)已更新,而不會記錄log
            'Memo by Amy ID=02750963 有X13175/X29726 編號,故+CU158條件
'           If strExc(1) = "" Then
'               strExc(0) = "Update Customer Set CU158=" & CNULL(strExc(1)) & " Where CU11='" & m_CU11 & "' And CU158 is not Null "
'           Else
'               strExc(0) = "Update Customer Set CU158=" & CNULL(strExc(1)) & " Where CU11='" & m_CU11 & "' And CU158 is Null "
'           End If
            'Memo by by Amy LOG有更新ID8個6且不是境外公司 ex:X9099200 (瑞婷:有時為開收據,無資料可建,而先輸8個6),若修改會將原本境外公司資料一同更新
            If InStr(";00000000;66666666;", ";" & m_CU11 & ";") = 0 Then
               strExc(0) = "Update Customer Set CU158=" & CNULL(strExc(1)) & " Where CU11='" & m_CU11 & "'"
            End If
            'end 2025/09/25
           
           'Modify by Amy 2025/09/24 傳入8碼客戶編號,記錄dl05
           Pub_SeekTbLog strExc(0), , , , , Left(txtKey, 8)   'Added by Morgan 2017/9/6
           adoTaie.Execute strExc(0)
        End If
        '2015/12/11 END
        
        'Add By Sindy 2016/2/17 境外公司欄由未勾選改為有勾選存檔時,
         If Check2(0).Value = 1 Then
            bolSendMail = CU158isYUpdAccData(txtName(0), strNo, txtKey)
         End If
         '2016/2/17 END
   End If
   'end 2014/09/23
   
   'Add By Sindy 2018/1/8
   'If lngEff = 1 Then
   If lngEff > 0 Then
   '2018/1/8 END
      adoTaie.CommitTrans
      
      'Add By Sindy 2016/3/21
      If bolSendMail = True Then
         PUB_SendMail strUserNum, strUserNum, "", txtKey & txtName(0) & "改為境外公司,過去年度之收據請再確認是否改為個人收據", "Dear Sirs," & vbCrLf & vbCrLf & strNo & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
      End If
      '2016/3/21 END
      
      FormSave = True
      If oldDizhang <> strDizhang And oldDizhang = "B" Then '有修改 帳款處理情形且取消「宣告破產」
         PUB_SendMail strUserNum, "83002", "", Left(txtKey, 8) & " 財務處取消宣告破產設定", "請至客戶或代理人檔取消狀態、呆帳記錄及備註欄之加註！"
      End If
      bolReadAfterSave = True 'Added by Morgan 2024/10/18
      If GetRec(txtKey) = True Then
         FormShow
      End If
      bolReadAfterSave = False 'Added by Morgan 2024/10/18
   Else
      adoTaie.RollbackTrans
      MsgBox "更新失敗！"
   End If
End Function

'Add by Amy 2014/04/03 由acc_var搬回修改
Public Sub SetCheck1(ByVal bolEnabled As Boolean)
   For Each oCheck In Check1
      oCheck.Enabled = bolEnabled
   Next
   
   For Each oText In txtInform
      oText.Enabled = bolEnabled
   Next
   'Add by Amy 2020/10/06 +特殊收據
   Check5.Enabled = bolEnabled
   If Check6.Visible Then Check6.Enabled = bolEnabled   'Added by Morgan 2024/5/28
   
   'Add By Sindy 2016/11/8
   If Left(txtKey, 1) = "Y" Then bolEnabled = False '**** Y(代理人)沒有代填繳款書機制
   'Modify by Amy 2025/02/20 原Check3=每月代填繳款書,改可勾智慧所/法律所
   Check3(0).Enabled = bolEnabled
   Check3(1).Enabled = bolEnabled
   'end 2025/02/20
   chkA4228(0).Enabled = bolEnabled: chkA4228(1).Enabled = bolEnabled 'Add By Sindy 2019/12/18
   textCU169.Enabled = bolEnabled
   textCU170.Enabled = bolEnabled
   textCU171.Enabled = bolEnabled
   '2016/11/8 END
   
   'Add by Amy 2014/09/23 +境外公司
   Check2(0).Enabled = bolEnabled
   Check2(1).Enabled = bolEnabled 'Add by Amy 2019/07/23 零稅率
   'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再印地址條
   'Add by Amy 2019/07/25 電子發票寄送方式
'   Check4(0).Enabled = bolEnabled
'   Check4(1).Enabled = bolEnabled
   'Added by Lydia 2023/11/13 請款匯入銀行資料
   txtCU196.Enabled = bolEnabled
   Combo4.Enabled = bolEnabled
   'end 2023/11/13
   
   'Text1.Enabled = bolEnabled 'Add By Sindy 2014/10/15 +會計備註  '2016/12/21 cancel by sonia 此欄都可以輸
   'add by sonia 2014/11/14 分所只能輸財務信箱及會計備註
   If pub_strUserOffice <> "1" Then
      Check1(0).Enabled = False
      Check1(1).Enabled = False
      Check1(2).Enabled = False
      Check1(3).Enabled = False
      'modify BY SONIA 2016/12/20 瑞婷說境外公司開給分所維護,但是否寄發收據,是否寄發電匯通知,是否寄發催款單,每月提醒代填繳款書四欄不開放
      'Check2(0).Enabled = False
      txtInform(0).Enabled = False
      txtInform(1).Enabled = False
      txtInform(2).Enabled = False
      txtInform(3).Enabled = False 'Add By Sindy 2017/3/16
      txtInform(4).Enabled = False 'Add By Sindy 2017/3/24
      'Modify by Amy 2025/02/20 原Check3.Enabled = False,改可勾智慧所/法律所
      Check3(0).Enabled = False
      Check3(1).Enabled = False
      'end 2025/02/20
      chkA4228(0).Enabled = 0: chkA4228(1).Enabled = 0 'Add By Sindy 2019/12/18
      'end 2016/12/20
   End If
   'end 2014/11/14
   
End Sub

'Add By Sindy 2016/11/8
Private Sub textCU169_GotFocus()
   TextInverse textCU169
   CloseIme
End Sub
Private Sub textCU169_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
      KeyAscii = 0
   End If
End Sub
Private Sub textCU169_Validate(Cancel As Boolean)
   If textCU169.Text = "" Then Exit Sub
   If textCU169.Enabled = False Then Exit Sub
   If textCU169 = "2" Then
      If cmdA49.BackColor <> &HC0FFC0 Then
         MsgBox "無會計師資料,繳款書寄件處不可選擇2.會計師!!", vbCritical
         Cancel = True
      Else
         strExc(0) = "select A4902,A4912,A4913 from ACC490 where A4901='" & txtKey & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If "" & RsTemp.Fields("A4902") = "" And "" & RsTemp.Fields("A4912") = "" Then
               MsgBox "會計師無輸入姓名及事務所名稱資料!!", vbCritical
               Cancel = True
            ElseIf "" & RsTemp.Fields("A4913") = "" Then
               MsgBox "會計師無輸入地址資料!!", vbCritical
               Cancel = True
            End If
         End If
      End If
   End If
End Sub
Private Sub textCU170_GotFocus()
   OpenIme
   TextInverse textCU170
End Sub

'Modified by Lydia 2021/12/02 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub textCU170_KeyPress(KeyAscii As MSForms.ReturnInteger)
   'Modified by Lydia 2021/12/14 +物件名稱
   KeyAscii = ChangeZIP(KeyAscii, textCU170)
End Sub
Private Sub textCU170_Validate(Cancel As Boolean)
   If textCU170.Text = "" Then Exit Sub
   If textCU170.Enabled = False Then Exit Sub
   
   'Add by Amy 2023/07/04 避免貼上的未轉全型 or KeyPress事件因輸入法失效沒執行到,故再轉一次
   textCU170 = PUB_ChangeZIPToSir(textCU170)
   
   If Not CheckLengthIsOK(textCU170, textCU170.MaxLength) Then
      Cancel = True
   End If
End Sub
Private Sub textCU171_GotFocus()
   OpenIme
   TextInverse textCU171
End Sub
Private Sub textCU171_Validate(Cancel As Boolean)
   If textCU171.Enabled = False Then Exit Sub
   If textCU171.Text = "" Then Exit Sub
   '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textCU171, textCU171.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
'2016/11/8 END

'Add By Sindy 2016/11/29
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'Added by Lydia 2023/11/13 請款匯入銀行資料
Private Sub SetCombo4(Optional ByVal pTxt As String)
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
   If pTxt = "" Then  '預設清單
      Combo4.Clear
      stSQL = "select * from rptaccount where ra01='CU196' order by ra02"
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         rsQuery.MoveFirst
         Combo4.AddItem "   "
         Do While Not rsQuery.EOF
            Combo4.AddItem "" & rsQuery.Fields("ra03")
            rsQuery.MoveNext
         Loop
      End If
      Set rsQuery = Nothing
      Combo4.ListIndex = 0
   Else
      If Val(pTxt) > 0 And Val(pTxt) < Combo4.ListCount Then
         Combo4.ListIndex = Val(pTxt)
      Else
         Combo4.ListIndex = 0
      End If
   End If
End Sub
'Added by Lydia 2023/11/13
Private Sub txtCU196_GotFocus()
   TextInverse txtCU196
End Sub
'Added by Lydia 2023/11/13
Private Sub txtCU196_Validate(Cancel As Boolean)
   If txtCU196.Tag <> txtCU196.Text Then
      Call SetCombo4(IIf(txtCU196.Text = "", "0", txtCU196.Text))
   End If
End Sub

'Added by Lydia 2023/11/13
Private Sub Combo4_Click()
   If txtCU196.Locked = False And txtCU196.Enabled = True Then
      'Added by Lydia 2024/12/10
      If Combo4.ListIndex = 0 Then
         txtCU196.Text = ""
      Else
      'end 2024/12/10
         txtCU196.Text = Combo4.ListIndex
      End If
   End If
End Sub
