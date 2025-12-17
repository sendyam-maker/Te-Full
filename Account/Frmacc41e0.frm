VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41e0 
   AutoRedraw      =   -1  'True
   Caption         =   "簽收作業"
   ClientHeight    =   5880
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   8760
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   6168
      ScaleHeight     =   252
      ScaleWidth      =   2364
      TabIndex        =   71
      Top             =   720
      Width           =   2412
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "北"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   24
         TabIndex        =   75
         Top             =   48
         Width           =   468
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   624
         TabIndex        =   74
         Top             =   48
         Width           =   468
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "南"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   1224
         TabIndex        =   73
         Top             =   48
         Width           =   468
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "高"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   1800
         TabIndex        =   72
         Top             =   48
         Width           =   468
      End
   End
   Begin VB.TextBox txtA2324 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      MaxLength       =   16
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   380
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "查詢"
      Height          =   315
      Left            =   5055
      TabIndex        =   69
      Top             =   710
      Width           =   765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "搜尋(&Q)"
      Default         =   -1  'True
      Height          =   315
      Left            =   5040
      TabIndex        =   65
      Top             =   380
      Width           =   765
   End
   Begin VB.TextBox txtA2328 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7065
      MaxLength       =   12
      TabIndex        =   13
      Top             =   2400
      Width           =   1440
   End
   Begin VB.TextBox txtA2327 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2115
      MaxLength       =   10
      TabIndex        =   12
      Top             =   2400
      Width           =   1260
   End
   Begin VB.TextBox txtA2326 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3555
      MaxLength       =   8
      TabIndex        =   11
      Top             =   2070
      Width           =   1125
   End
   Begin VB.CheckBox Check1 
      Caption         =   "已處理"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6075
      TabIndex        =   58
      Top             =   1350
      Width           =   1050
   End
   Begin VB.CommandButton cmdFind2 
      Height          =   300
      Left            =   8235
      Picture         =   "Frmacc41e0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2790
      Width           =   350
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   5175
      Style           =   2  '單純下拉式
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2790
      Width           =   3030
   End
   Begin VB.TextBox txtA2304 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      MaxLength       =   9
      TabIndex        =   3
      Top             =   1030
      Width           =   1305
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   5175
      ScaleHeight     =   264
      ScaleWidth      =   1812
      TabIndex        =   54
      Top             =   3120
      Width           =   1860
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "提款機"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   855
         TabIndex        =   18
         Top             =   0
         Width           =   960
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "電匯"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   45
         TabIndex        =   16
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.TextBox txtA2308 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      MaxLength       =   16
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   50
      Width           =   1395
   End
   Begin VB.CommandButton cmdEmail 
      Caption         =   "Email"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7245
      TabIndex        =   53
      Top             =   1100
      Width           =   1365
   End
   Begin VB.TextBox txtDif 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7695
      MaxLength       =   16
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3450
      Width           =   900
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5175
      MaxLength       =   16
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3450
      Width           =   1395
   End
   Begin VB.TextBox txtA2301 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   0
      Top             =   50
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Height          =   300
      Left            =   2385
      Picture         =   "Frmacc41e0.frx":0102
      Style           =   1  '圖片外觀
      TabIndex        =   46
      Top             =   57
      Width           =   350
   End
   Begin VB.TextBox txtA2321 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1890
   End
   Begin VB.TextBox txtA2306 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   7695
      MaxLength       =   16
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   900
   End
   Begin VB.TextBox txtA2306 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   6210
      MaxLength       =   16
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Width           =   900
   End
   Begin VB.TextBox txtA2306 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   4680
      MaxLength       =   16
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   900
   End
   Begin VB.TextBox txtA2306 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1680
      MaxLength       =   16
      TabIndex        =   5
      Top             =   1680
      Width           =   900
   End
   Begin VB.CommandButton cmdClear 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4770
      Picture         =   "Frmacc41e0.frx":0204
      Style           =   1  '圖片外觀
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "清除畫面"
      Top             =   5220
      Width           =   550
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "E"
      Top             =   5355
      Width           =   270
   End
   Begin VB.TextBox txtA2309 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1185
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3780
      Width           =   7425
   End
   Begin VB.TextBox txtNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      MaxLength       =   8
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5355
      Width           =   1710
   End
   Begin VB.CommandButton cmdCut 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5355
      Picture         =   "Frmacc41e0.frx":0ACE
      Style           =   1  '圖片外觀
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "取消"
      Top             =   5220
      Width           =   550
   End
   Begin VB.TextBox txtA2307 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      MaxLength       =   16
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3450
      Width           =   945
   End
   Begin VB.TextBox txtA2306 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1185
      MaxLength       =   16
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3450
      Width           =   1440
   End
   Begin VB.TextBox txtA2303 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   4
      Text            =   "99999"
      Top             =   1350
      Width           =   810
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   1000
      Left            =   135
      TabIndex        =   30
      Top             =   4110
      Width           =   8505
      _ExtentX        =   15007
      _ExtentY        =   1778
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSMask.MaskEdBox txtA2302 
      Height          =   300
      Left            =   4230
      TabIndex        =   1
      Top             =   50
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtA2325 
      Height          =   330
      Left            =   1680
      TabIndex        =   10
      Top             =   2070
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   572
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtA2306 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   3210
      MaxLength       =   16
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Width           =   900
   End
   Begin MSForms.TextBox txtSales 
      Height          =   315
      Left            =   1980
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1125
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA2330 
      Height          =   315
      Left            =   1170
      TabIndex        =   67
      Top             =   700
      Width           =   3870
      VariousPropertyBits=   679493659
      MaxLength       =   100
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtTitle 
      Height          =   315
      Left            =   1170
      TabIndex        =   2
      Top             =   380
      Width           =   3870
      VariousPropertyBits=   679493659
      MaxLength       =   40
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtBankName 
      Height          =   315
      Left            =   3375
      TabIndex        =   63
      Top             =   2400
      Width           =   2565
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA2310 
      Height          =   615
      Left            =   1185
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2790
      Width           =   2970
      VariousPropertyBits=   -1466941413
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "5239;1085"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCustomer 
      Height          =   315
      Left            =   2475
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1030
      Width           =   4680
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label29 
      BackStyle       =   0  '透明
      Caption         =   "模糊比對"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   6000
      TabIndex        =   70
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label28 
      BackStyle       =   0  '透明
      Caption         =   "電匯資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   68
      Top             =   745
      Width           =   900
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   66
      Top             =   425
      Width           =   900
   End
   Begin VB.Label Label26 
      BackStyle       =   0  '透明
      Caption         =   "票據資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   64
      Top             =   2100
      Width           =   960
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   750
      Left            =   90
      Top             =   2010
      Width           =   8550
   End
   Begin VB.Label Label22 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收票帳號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6075
      TabIndex        =   62
      Top             =   2460
      Width           =   900
   End
   Begin VB.Label Label25 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收票銀行"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1185
      TabIndex        =   61
      Top             =   2460
      Width           =   900
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "票號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3015
      TabIndex        =   60
      Top             =   2100
      Width           =   450
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "票期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1185
      TabIndex        =   59
      Top             =   2100
      Width           =   450
   End
   Begin VB.Label Label21 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "銀存科目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4275
      TabIndex        =   57
      Top             =   2835
      Width           =   900
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "Eail日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6165
      TabIndex        =   55
      Top             =   425
      Width           =   975
   End
   Begin VB.Label Label19 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "差額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7215
      TabIndex        =   52
      Top             =   3495
      Width           =   450
   End
   Begin VB.Label Label18 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "未收合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4275
      TabIndex        =   51
      Top             =   3495
      Width           =   900
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "簽收單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   49
      Top             =   102
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "輸入日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3285
      TabIndex        =   48
      Top             =   102
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "繳收據日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6165
      TabIndex        =   47
      Top             =   102
      Width           =   975
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "簽收確認"
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
      Left            =   3285
      TabIndex        =   44
      Top             =   1370
      Width           =   840
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "總額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   43
      Top             =   3495
      Width           =   450
   End
   Begin VB.Label Label16 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "其他"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7215
      TabIndex        =   42
      Top             =   1725
      Width           =   450
   End
   Begin VB.Label Label15 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "暫存"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5670
      TabIndex        =   41
      Top             =   1725
      Width           =   450
   End
   Begin VB.Label Label11 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "銀存"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4200
      TabIndex        =   40
      Top             =   1725
      Width           =   450
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "現金"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2715
      TabIndex        =   39
      Top             =   1725
      Width           =   450
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "票額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1185
      TabIndex        =   38
      Top             =   1725
      Width           =   450
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "收據編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   225
      TabIndex        =   37
      Top             =   5400
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   600
      Left            =   135
      Top             =   5170
      Width           =   5850
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收據編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   36
      Top             =   3810
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "扣繳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2760
      TabIndex        =   35
      Top             =   3495
      Width           =   450
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   34
      Top             =   2850
      Width           =   645
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "收款金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   33
      Top             =   1725
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "客戶"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   32
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   31
      Top             =   1370
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4716
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc41e0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 txtTitle/txtA2330/txtCustomer/txtSales/txtBankName/txtA2310/grdDataList
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Modified by Morgan 2016/8/11 客戶(編號)欄位合併為一欄(第一碼X不再單獨顯示)
Option Explicit
Public m_Status As String '狀態
Dim m_LstAddIndex As Integer 'Added by Morgan 2015/7/20 最後新增收款種類：票額1、現金2...
'Removed by Morgan 2013/8/30 改抓特殊設定
'Const cNonSalesList As String = "86048;79041;73014;78048;78027;99020;65001;69001;67002;69004;84027;70003;A0038" '非智權人員要通知清單 Added by Morgan 2013/4/29
Dim m_LstIndex As Integer 'Added by Morgan 2014/1/13
Dim m_bolOffDutySales As Boolean 'Added by Morgan 2015/8/5 智權人員是否已離職
Dim bolMatch As Boolean     'modify by sonia 2020/5/19 由Command1_Click移上來
Const Color1 As Long = &HFFFFFF
Const Color2 As Long = &HE0E0E0
Dim m_lstOption1 As Integer 'Added by Morgan 2025/6/13

Public Sub MailCheck()
   Dim strVer As String
   
   'Mark by Amy 2020/06/08 取消-辜
   'Add by Amy 2020/04/23 +if 銀存科目110502 不發mail-辜
'   If Trim(Combo1.Text) <> MsgText(601) Then
'        If Left(Combo1.Text, InStr(Combo1, " ") - 1) = "110502" Then Exit Sub
'   End If
   
   If m_Status = "2" And txtA2321 <> "" Then
      If Trim(txtA2310.Tag) <> Trim(txtA2310) Then
         strVer = ""
         strVer = strVer & "繳款日：" & txtA2302 & vbCrLf
         strVer = strVer & "客戶：" & txtA2304 & " " & txtCustomer & vbCrLf
         strVer = strVer & "原備註：" & txtA2310.Tag & vbCrLf
         strVer = strVer & "新備註：" & txtA2310 & vbCrLf
         PUB_SendMail strUserNum, txtA2303, "", "簽收單<" & txtA2301 & ">備註修改！", strVer
      End If
   End If
End Sub

Private Sub cmdClear_Click()
   txtNo.Text = ""
End Sub

Private Sub cmdCut_Click()
   Dim ii As Integer
   With grdDataList
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) = (txtCode & txtNo) Then
            If txtA2306(1).Enabled = True Then
               txtA2306(1) = Val(txtA2306(1)) - Format(.TextMatrix(ii, 4), "0")
               txtA2306(0) = Val(txtA2306(1)) + Val(txtA2306(2)) + Val(txtA2306(3)) + Val(txtA2306(4)) + Val(txtA2306(5)) + Val(txtA2307)
               txtTot = Val(txtTot) - Format(.TextMatrix(ii, 4), "0") 'Add by Morgan 2006/10/17
            End If
            If .Rows = 2 Then
               SetDataListWidth
               txtA2308.Text = ""
               txtNo.Text = ""
            Else
               .RemoveItem (ii)
               txtNo.Text = Right(.TextMatrix(.row, 0), 8)
            End If
            If txtA2309 <> "" Then
               txtA2309 = Replace(txtA2309, txtCode & txtNo & ",", "")
               txtA2309 = Replace(txtA2309, "," & txtCode & txtNo, "")
               txtA2309 = Replace(txtA2309, txtCode & txtNo, "")
            End If
            
            Exit For
         End If
      Next
   End With
End Sub

Public Function ReadData(ByVal p_A2301 As String, Optional ByVal iAct As Integer = 0) As Boolean
   Dim stMsg As String, iA2305 As Integer, ii As Integer, stCon As String
   
   'Added by Morgan 2013/4/25
   cmdEmail.Enabled = False
   SetOption False, True
   'end 2013/4/25
   
   'Modified by Morgan 2015/6/17 改北所人員可看全部
   'If Pub_StrUserSt03 <> "M51" Then
   If pub_strUserOffice <> "1" Then
   'end 2015/6/17
      stCon = stCon & " and A2305='" & pub_strUserOffice & "'"
   End If
   
   Select Case iAct
      '指定
      Case 0
         strSql = "Select * From ACC230,ACC010 Where A2301='" & p_A2301 & "'"
         stMsg = "簽收單號[" & p_A2301 & "]不存在！"
      '首筆
      Case 1
         strSql = "Select * From ACC230 WHERE A2301=(SELECT MIN(A2301) FROM ACC230 WHERE 1=1" & stCon & ")"
         stMsg = "無簽收資料！"
      '上筆
      Case 2
         strSql = "Select * From ACC230 Where A2301=(SELECT MAX(A2301) FROM ACC230 WHERE A2301<'" & p_A2301 & "'" & stCon & ")"
         stMsg = "已經是第一筆簽收資料！"
      '下筆
      Case 3
         strSql = "Select * From ACC230 Where A2301=(SELECT MIN(A2301) FROM ACC230 WHERE A2301>'" & p_A2301 & "'" & stCon & ")"
         stMsg = "已經是最後一筆簽收資料！"
      '末筆
      Case Else
         strSql = "Select * From ACC230 WHERE A2301=(SELECT MAX(A2301) FROM ACC230 WHERE 1=1" & stCon & ")"
         stMsg = "無簽收資料！"
   End Select
   
On Error GoTo ErrHnd
    
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         Me.FormClear
         txtA2301 = "" & .Fields("A2301")
         'Modified by Morgan 2014/3/12
         'txtA2302 = Format("" & .Fields("A2302"), "0##/##/##")
         txtA2302.Mask = MsgText(601)
         If IsNull(.Fields("A2302").Value) Then
            txtA2302.Text = MsgText(601)
         Else
            txtA2302.Text = CFDate(Trim(str(.Fields("A2302").Value)))
         End If
         txtA2302.Mask = DFormat
         'end 2014/3/12
         txtA2303 = "" & .Fields("A2303")
         'Modified by Morgan 2013/4/29
         'txtA2304 = "" & .Fields("A2304")
         txtA2304 = "" & .Fields("A2304")
         txtA2306(1) = "" & .Fields("A2306")
         txtA2306(2) = "" & .Fields("A2317")
         txtA2306(3) = "" & .Fields("A2318")
         txtA2306(4) = "" & .Fields("A2319")
         txtA2306(5) = "" & .Fields("A2320")
         txtA2307 = "" & .Fields("A2307")
         txtA2306(0) = Val(txtA2306(1)) + Val(txtA2306(2)) + Val(txtA2306(3)) + Val(txtA2306(4)) + Val(txtA2306(5)) + Val(txtA2307)
         txtA2308 = Format("" & .Fields("A2308"), "0##/##/##")
         txtA2309 = "" & .Fields("A2309")
         txtA2310 = "" & .Fields("A2310")
         txtA2310.Tag = txtA2310 '紀錄原備註
         If IsNull(.Fields("A2321")) Then
            txtA2321 = ""
            Check1.Value = 0 'Added by Morgan 2015/6/17
         Else
            txtA2321 = Format(.Fields("A2321"), "YYYY") - 1911 & Format(.Fields("A2321"), "/MM/DD hh:mm:ss")
            Check1.Value = 1 'Added by Morgan 2015/6/17
         End If
         
         If txtA2309 <> "" Then
            Call ReadGridData(txtA2309)
         End If
         'Added by Morgan 2013/4/25
         'If pub_strUserOffice = "1" Then 'Removed by Morgan 2014/1/23
            If txtA2309 = "" Then cmdEmail.Enabled = True
         'End If 'Removed by Morgan 2014/1/23
         
         'Added by Morgan 2013/12/23 改放科目代號
         'If .Fields("A2322") = "1" Then
         '   Option1(0).Value = True
         'ElseIf .Fields("A2322") = "2" Then
         '   Option1(1).Value = True
         'End If
         If IsNull(.Fields("a2322")) Then
            Combo1.ListIndex = -1
         Else
            SetCombo1 .Fields("a2322")
         End If
         'end 2013/12/23
        
         If .Fields("A2323") = "1" Then
            Option2(0).Value = True
         ElseIf .Fields("A2323") = "2" Then
            Option2(1).Value = True
         End If
         'end 2013/4/25
         'Added by Morgan 2013/5/8
         If IsNull(.Fields("A2324")) Then
            txtA2324 = ""
         Else
            txtA2324 = Format(.Fields("A2324"), "EE/MM/DD")
         End If
         
         'Added by Morgan 2015/7/20
         txtA2325.Mask = MsgText(601)
         If IsNull(.Fields("A2325").Value) Then
            txtA2325.Text = MsgText(601)
         Else
            txtA2325.Text = CFDate(Trim(str(.Fields("A2325").Value)))
         End If
         txtA2325.Mask = DFormat
         txtA2326 = "" & .Fields("A2326")
         txtA2327 = "" & .Fields("A2327")
         If txtA2327 <> "" Then
            txtA2327_Validate False
         End If
         txtA2328 = "" & .Fields("A2328")
         'end 2015/7/20
         txtA2330 = "" & .Fields("A2330") 'Add by Amy 2017/12/07 電匯資料
         
         'Added by Morgan 2022/6/1
         txtA2326.Tag = txtA2326
         txtA2327.Tag = txtA2327
         txtA2328.Tag = txtA2328
         'end 2022/6/1
         
         'Added by Morgan 2025/6/11
         intI = Val("" & .Fields("A2305"))
         If intI >= 1 And intI <= 4 Then
            Option1(intI).Value = True
         End If
         'end 2025/6/11
         ReadData = True
      Else
         MsgBox stMsg, vbExclamation
      End If
ErrHnd:
      CheckOC
   End With
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   'Resume
End Function

Private Sub SetCombo1(pCode As String)
   Dim idx As Integer
   For idx = 0 To Combo1.ListCount - 1
      'Modified by Morgan 2025/6/10
      'If InStr(Combo1.List(idx), pCode) = 1 Then
      If InStr(Combo1.List(idx), pCode & " ") = 1 Then
      'end 2025/6/10
         Combo1.ListIndex = idx
         Exit For
      End If
   Next
   If idx = Combo1.ListCount Then
      'Modified by Morgan 2025/5/29
      'Combo1.AddItem pCode
      'Combo1 = pCode
      Combo1.AddItem pCode & " " & A0102Query(pCode)
      Combo1.ListIndex = idx
      'end 205/5/29
   End If
End Sub

Private Function ReadGridData(ByVal p_A2309 As String, Optional ByVal p_bCheck As Boolean = False) As Boolean
   Dim A0K
   Dim ii As Integer, stA0k01 As String
   A0K = Split(p_A2309, ",")
   'Modify by Morgan 2011/8/18
   'strSql = "Select a0k01,a0k02,a0k03,a0k04,NVL(a0k06,0)+NVL(a0k07,0)-NVL(X2,0) Amt1,a0k20,a0k09,X1 Amt2" & _
      " From ACC0k0,(select CP60,sum(CP79) X1,SUM(CP77) X2" & _
      " from CASEPROGRESS where CP60='" & A0K(0) & "' group by CP60) X" & _
      " Where A0k01='" & A0K(0) & "' and CP60(+)=a0k01"
   
   'For ii = LBound(A0K) + 1 To UBound(A0K)
      'strSql = strSql & " UNION ALL " & _
         "Select a0k01,a0k02,a0k03,a0k04,NVL(a0k06,0)+NVL(a0k07,0)-NVL(X2,0) Amt1,a0k20,a0k09,X1 Amt2" & _
      " From ACC0k0,(select CP60,sum(CP79) X1,SUM(CP77) X2" & _
      " from CASEPROGRESS where CP60='" & A0K(ii) & "' group by CP60) X" & _
      " Where A0k01='" & A0K(ii) & "' and CP60(+)=a0k01"
   'Next
   '收據將改為可一收文多收據,故改抓acc1u0
   strSql = ""
   For ii = LBound(A0K) To UBound(A0K)
      If strSql <> "" Then strSql = strSql & " UNION ALL "
      strSql = strSql & " Select a0k01,a0k02,a0k03,a0k04,NVL(a0k06,0)+NVL(a0k07,0)-NVL(X2,0) Amt1,a0k20,a0k09" & _
         ",nvl(a0k06,0)+nvl(a0k07,0)-nvl(X1,0)-nvl(X2,0)+nvl(X3,0) Amt2" & _
         " From ACC0k0,( select a1u02,nvl(sum(a1u04),0)+nvl(sum(a1u05),0) X1" & _
         ",nvl(sum(a1u07),0)+nvl(sum(a1u09),0) X2,nvl(sum(a1u08),0)+nvl(sum(a1u10),0) X3" & _
         " from acc1u0 where a1u02='" & A0K(ii) & "' group by a1u02" & _
         ") X Where A0k01='" & A0K(ii) & "' and a1u02(+)=a0k01"
   Next
   'end 2011/8/18
   strSql = strSql & " ORDER BY 1"
   
On Error GoTo ErrHnd
   
   CheckOC2
   With adoRecordset1
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         'Add by Morgan 2005/5/19
         If p_bCheck = True And Val("" & .Fields("a0k09")) > 0 Then
            MsgBox "收據已作廢", vbExclamation
         Else
            Do While Not .EOF
               '新增
               If p_bCheck = True Then
                  If txtA2303 = "" Then
                     txtA2303 = "" & .Fields("A0k20")
                  'Add by Morgan 2005/5/19
                  ElseIf txtA2303 <> "" & .Fields("A0k20") Then
                     MsgBox "收據智權人員與此次簽收智權人員不同！", vbExclamation
                     GoTo ErrHnd
                  End If
                  If txtA2304 = "" Then
                  'Modified by Morgan 2013/4/29
                  '   txtA2304 = "" & .Fields("A0k03")
                  'ElseIf Left(txtA2304, 6) <> Left("" & .Fields("A0k03"), 6) Then
                     txtA2304 = "" & .Fields("A0k03")
                  ElseIf Left(txtA2304, 6) <> Left("" & .Fields("A0k03"), 6) Then
                  'end 2013/4/29
                     'Modify by Morgan 2005/8/11 不同客戶收據要可簽收--瑞婷
                     'MsgBox "收據客戶與此次簽收客戶不同！", vbExclamation
                     If MsgBox("收據客戶與此次簽收客戶不同，確定要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                        GoTo ErrHnd
                     End If
                  End If
                  For ii = 1 To grdDataList.Rows - 1
                     If grdDataList.TextMatrix(ii, 0) = "" & .Fields(0) Then
                        Exit For
                     End If
                  Next
                  If ii = grdDataList.Rows Then
                     If txtA2306(1).Enabled = True Then
                        txtA2306(1) = Val(txtA2306(1)) + Val("" & .Fields("Amt2"))
                        txtA2306(0) = Val(txtA2306(1)) + Val(txtA2306(2)) + Val(txtA2306(3)) + Val(txtA2306(4)) + Val(txtA2306(5)) + Val(txtA2307)
                     End If
                     If txtA2308.Text = "" Then
                        txtA2308.Text = CFDate(strSrvDate(2))
                     End If
                  End If
               '查詢
               Else
                  ii = grdDataList.Rows
               End If
               If ii = grdDataList.Rows Then
                  'Modify by Morgan 2005/5/19 改後輸入的放上面
                  If grdDataList.TextMatrix(1, 0) <> "" Then
                     'grdDataList.AddItem "", grdDataList.Rows
                     grdDataList.AddItem "", 1
                  End If
                  grdDataList.TextMatrix(1, 0) = "" & .Fields("A0k01")
                  grdDataList.TextMatrix(1, 1) = Format("" & .Fields("A0k02"), "###/##/##")
                  grdDataList.TextMatrix(1, 2) = "" & .Fields("A0k03")
                  grdDataList.TextMatrix(1, 3) = "" & .Fields("A0k04")
                  grdDataList.TextMatrix(1, 4) = Format(Val("" & .Fields("Amt2")), FDollar) '未收
                  grdDataList.TextMatrix(1, 5) = Format(Val("" & .Fields("Amt1")), FDollar) '應收
                  
               End If
               'grdDataList.TopRow = grdDataList.Rows - 1
               'Add by Morgan 2006/10/17
               txtTot = Val(txtTot) + Val("" & .Fields("Amt2"))
               txtDif = Val(txtTot) - Val(txtA2306(0))
               'end 2006/10/17
               .MoveNext
            Loop
            grdDataList.row = grdDataList.Rows - 1
            ReadGridData = True
         End If
      ElseIf p_bCheck = True Then
         MsgBox "收據號碼[" & txtCode & txtNo & "]不存在！", vbExclamation
      End If
ErrHnd:
      CheckOC2
      .MaxRecords = 0
   End With
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Added by Morgan 2013/4/29
Private Sub cmdEmail_Click()
   Mail2Sales False
End Sub

Private Sub cmdFind_Click()
   Me.FormClear False
   If Len(txtA2301) = 10 Then
      Call ReadData(txtA2301)
   End If
End Sub

'Added by Morgan 2014/3/12
Private Sub cmdFind2_Click()
   If Combo1 = "" Then
      MsgBox "請先點選銀存科目!!", vbExclamation
   Else
      strExc(1) = Left(Combo1.Text, InStr(Combo1, " ") - 1)
      'Modified by Morgan 2023/2/8 1911-1913改以科目判斷所別
      'strExc(0) = "select count(*),nvl(max(a2301),'S') from acc230 where a2322='" & strExc(1) & "' and A2305='" & pub_strUserOffice & "'"
      strExc(0) = "select count(*),nvl(max(a2301),'S') from acc230 where a2322='" & strExc(1) & "'"
      If pub_strUserOffice <> "1" And Not ((strExc(1) = "1911" And pub_strUserOffice = "2") _
         Or (strExc(1) = "1912" And pub_strUserOffice = "3") _
         Or (strExc(1) = "1913" And pub_strUserOffice = "4")) Then
         strExc(0) = strExc(0) & " and A2305='" & pub_strUserOffice & "'"
      End If
      'end 2023/2/8
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) > 0 Then
            txtA2301 = RsTemp(1)
            cmdFind.Value = True
         Else
            MsgBox "無該科目資料!!", vbExclamation
         End If
      End If
   End If
End Sub

Public Sub GetSelect()
   Me.Enabled = True
   
   With Frmacc1220
      txtA2304 = .Adodc1.Recordset("A0K03")
      'Modify by Amy 2022/06/27 原以txtTitle.Tag 記錄
      txtTitle = "" & .Adodc1.Recordset("A0K04") 'Add by Amy 2017/12/25 記錄所選抬頭
      If txtCustomer <> .Adodc1.Recordset("A0K04") Then
         txtA2310 = .Adodc1.Recordset("A0K04")
      End If
      txtA2303 = .Adodc1.Recordset("A0K20")
      'add by sonia 2020/5/19 只有一組且為法律所案源資料改帶介紹智權人員los04
      If bolMatch = True Then
         'modify by sonia 2020/6/9 介紹智權人員los04只帶第一人
         'strExc(0) = "select los04 from lawofficesource where los06='" & .Adodc1.Recordset("A0j01") & "' "
         strExc(0) = "select decode(instr(los04,','),0,los04,substr(los04,1,instr(los04,',')-1)) los04 from lawofficesource where los06='" & .Adodc1.Recordset("A0j01") & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then txtA2303 = RsTemp("los04")
      End If
      'end 2020/5/19
   End With
   
   If txtA2304 = "" Then
      txtA2304.SetFocus
   ElseIf txtA2303 = "" Then
      txtA2303.SetFocus
   Else
      For intI = 1 To 5
         If txtA2306(intI).TabStop = True Then
            txtA2306(intI).SetFocus
            Exit For
         End If
      Next
   End If
End Sub

Private Sub Command1_Click()
   
   Dim ii As Integer
   
   If txtTitle = "" Then
      MsgBox "請輸入收據抬頭的關鍵字!!!", vbExclamation + vbOKOnly
      txtTitle.SetFocus
      txtTitle_GotFocus
      Exit Sub
   End If
   
   txtA2304 = ""
   txtA2303 = ""
   txtA2310 = ""
   
   With Frmacc1220
   Set .frmCall = Me
   .SetForm
   'Modify by Amy 2022/06/27 若輸抬頭搜尋,抬頭不是要的資料,直接輸客戶編號,收據抬頭.Tag不會被清mail會帶錯(搜尋字改.tag記錄)
   txtTitle.Tag = txtTitle
   txtTitle = ""
   .Text3 = txtTitle.Tag
   'end 2022/06/27
   .MaskEdBox1 = CFDate(TransDate(CompDate(0, -2, strSrvDate(1)), 1))
   .MaskEdBox2 = CFDate(strSrvDate(2))
   .Text6 = "1"
   .Check1.Value = vbChecked 'Added by Morgan 2020/11/12 未列印收據也要--辜
   .KeyDefine vbKeyF12
   If .Adodc1.Recordset.State <> adStateOpen Then
      Unload Frmacc1220
      txtTitle.SetFocus
      txtTitle_GotFocus
   Else
      Set RsTemp = .Adodc1.Recordset.Clone
      With RsTemp
      bolMatch = True
      strExc(1) = Left("" & .Fields("a0k03"), 6) & .Fields("a0k04") & .Fields("a0k20")
      Do While Not .EOF
         If Left("" & .Fields("a0k03"), 6) & .Fields("a0k04") & .Fields("a0k20") <> strExc(1) Then
            bolMatch = False
            Exit Do
         End If
         .MoveNext
      Loop
      .MoveFirst
      End With
      
      '只有一組的 "客戶編號6碼+智權人員+收據抬頭"
      If bolMatch = True Then
         GetSelect
         Unload Frmacc1220
      Else
         Me.Enabled = False
         .Show
      End If
   End If
   End With
   
End Sub

'Add by Amy 2017/12/07 電匯資料查詢
Private Sub Command2_Click()
    Dim txtSearch As String
    
    If txtA2330 = MsgText(601) Then
        MsgBox "請輸入電匯資料的關鍵字!!!", vbExclamation + vbOKOnly
        Exit Sub
    End If
    If txtTitle <> MsgText(601) Then txtTitle = "" 'Add by Amy 2017/12/25  for 發mail 判斷帶的內容
    txtA2304 = "": txtCustomer = ""
    txtA2303 = ""
    
    txtSearch = txtA2330
    Call Frmacc11n1.SetParent(Me)
    Frmacc11n1.Text1 = txtSearch
    Frmacc11n1.KeyDefine vbKeyF12
    Me.Enabled = False
    'Modify by Amy 2017/12/18
    If Frmacc11n1.adoadodc1.RecordCount = 0 Then
        Unload Frmacc11n1
        Me.Enabled = True
    ElseIf Frmacc11n1.adoadodc1.RecordCount = 1 Then
        Frmacc11n1.DataGrid1_DblClick
        Me.Enabled = True
    Else
        Frmacc11n1.Show
    End If
   
End Sub

Private Sub Form_Activate()
   strFormName = Name
   strFormLink = ""
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub
Private Sub Form_Load()
   Dim intX As Integer, intY As Integer, sglWidth As Single, sglHeight As Single, ii As Integer
   Dim stAccAll As String, stAccField As String
      
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   txtA2302.Mask = DFormat 'Added by Morgan 2014/3/12
   
   'Modify by Amy 2023/05/25 改共用
'   'Added by Morgan 2013/12/23
'   Combo1.Clear
'   If pub_strUserOffice = "2" Then
'      strExc(0) = "select a0101||' '||a0102 from acc010 where  a0101='1911'"
'   ElseIf pub_strUserOffice = "3" Then
'      strExc(0) = "select a0101||' '||a0102 from acc010 where  a0101='1912'"
'   ElseIf pub_strUserOffice = "4" Then
'      strExc(0) = "select a0101||' '||a0102 from acc010 where  a0101='1913'"
'   Else
'      'Modify by Amy 2014/05/08 +會計科目110209
'      'Modify by Amy 2020/04/08 +會計科目110230/110231
'      'Modify by Amy 2020/04/09 110230改為110602/110231改為110502
'      'Modify by Amy 2023/05/23 刪除110202/110223/110209 並改為自動產生序號
''      strExc(0) = "select a0101||' '||a0102,decode(a0101,'110202',1,'110207',2,'110303',3,'110223',4,'110208',5,'110208',6,'110204',7,'110205',8,'110301',9,'110302',10,'110209',11,'110602',12,'110502',13,14) Srt" & _
''         " from acc010 where  a0101 in ('110202','110204','110205','110207','110208','110209','110223','110602','110502','110301','110302','110303','1911','1912','1913') order by 2"
'      stAccField = GetAccSeq(stAccAll)
'      strExc(0) = "select a0101||' '||a0102" & stAccField & _
'         " from acc010 where  a0101 in (" & stAccAll & ") order by 2"
'   End If
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      Do While Not RsTemp.EOF
'         Combo1.AddItem RsTemp(0)
'         RsTemp.MoveNext
'      Loop
'   End If
'   'end 2013/12/23
   Pub_AccBankTit Combo1, Me.Name
   'end 2023/05/25
   
   SetDataListWidth
   FormEnable
   ReadData "", 4
   
'   If pub_strUserOffice <> "1" Then
'      txtA2310.Width = 7380
'      cmdEmail.Visible = False
'      txtA2324.Visible = False
'      Label20.Visible = False
'   End If
   
   m_LstIndex = -1 'Added by Morgan 2014/1/13
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   'Modify by Amy 2014/04/16 +if
   If Me.Tag = MsgText(601) Then
        strConTitle = MsgText(601)
        strFormName = MsgText(601)
        KeyEnter vbKeyEscape
        MenuEnabled
        Set Frmacc41e0 = Nothing
        Exit Sub
   End If
   
   'Add by Amy 2014/04/16
   tool3_enabled
   Me.Tag = MsgText(601)
   Frmacc42b0.Show
   Set Frmacc41e0 = Nothing
End Sub
'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim ii As Integer
   With grdDataList
      .Visible = False
      If p_bolHeaderOnly = False Then
         .Clear
         .Rows = 2: .Cols = 6: .FixedRows = 1: .FixedCols = 0
      End If
      .row = 0
      
      .col = 0: .ColWidth(.col) = 1100: .Text = "收據編號"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .col = 1: .ColWidth(.col) = 1100: .Text = "收據日期"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .col = 2: .ColWidth(.col) = 1100: .Text = "客戶編號"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .col = 3: .ColWidth(.col) = 2300: .Text = "收據抬頭"
      .ColAlignment(.col) = flexAlignLeftCenter
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .col = 4: .ColWidth(.col) = 1300: .Text = "未收金額"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .col = 5: .ColWidth(.col) = 1300: .Text = "金額"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      '.Refresh
      .Visible = True
   End With
End Sub
'p_bAll:是否全部清除
Public Sub FormClear(Optional ByVal p_bAll As Boolean = True)
   Dim ii As Integer
   If p_bAll Then
      txtA2301.Text = ""
   End If
   'Modified by Morgan 2014/3/12
   'txtA2302.Text = ""
   txtA2302.Mask = ""
   txtA2302.Text = ""
   txtA2302.Mask = DFormat
   'end 2014/3/12
   txtA2303.Text = "": txtSales.Text = ""
   txtA2304.Text = "": txtCustomer.Text = ""
   For ii = 0 To 5
      txtA2306(ii).Text = ""
   Next
   txtA2307.Text = ""
   txtA2308.Text = ""
   txtA2309.Text = ""
   txtA2310.Text = ""
   txtA2324.Text = "" 'Added by Morgan 2013/5/9
   txtA2321.Text = "" 'Added by Morgan 2014/1/23
   
   'Added by Morgan 2015/7/20
   txtA2325.Mask = ""
   txtA2325.Text = ""
   txtA2325.Mask = DFormat
   txtA2326.Text = ""
   txtA2327.Text = ""
   txtA2328.Text = ""
   'end 2015/7/20
   
   'Added by Morgan 2022/6/1
   txtA2326.Tag = txtA2326
   txtA2327.Tag = txtA2327
   txtA2328.Tag = txtA2328
   'end 2022/6/1
         
   txtA2330.Text = "" 'Add by Amy 2017/12/07
   txtTitle = "" 'Added by Morgan 2015/7/23
   txtTitle.Tag = "" 'Add by Amy 2017/12/25
   txtNo.Text = ""
   txtTot = "": txtDif = "" 'Add by Morgan 2006/10/17
   SetDataListWidth
   SetOption False, True 'Added by Morgan 2013/4/25
End Sub
Public Sub FormEnable(Optional ByVal p_Status As String = "0")
   Dim bolValue As Boolean, ii As Integer
   
   m_Status = p_Status
   Select Case p_Status
      Case "0" '查詢
         bolValue = True
      Case "1" '新增
         bolValue = False
         'Modified by Morgan 2014/3/12
         'txtA2302 = CFDate(strSrvDate(2))
         txtA2302.Mask = MsgText(601)
         txtA2302.Text = CFDate(strSrvDate(2))
         txtA2302.Mask = DFormat
         'end 2014/3/12
      Case "2" '修改
         '智權人員已確認簽收只能修改備註
         If txtA2321 <> "" Then
            bolValue = True
         Else
            bolValue = False
         End If
   End Select
   
   'Added by Morgan 2015/6/17
   If p_Status = "2" Then
      If txtA2321 = "" Then
         Check1.Enabled = True
      ElseIf Left(txtA2321, 8) = "11/11/11" Then
         Check1.Enabled = True
      Else
         Check1.Enabled = False
      End If
   Else
      Check1.Enabled = False
   End If
   'end 2015/6/17
            
   txtA2302.Enabled = Not bolValue 'Added by Morgan 2014/3/12
   
   '簽收單號
   If p_Status = "0" Then
      txtA2301.Enabled = True
      If Me.Visible Then txtA2301.SetFocus
      cmdFind.Enabled = True
      cmdFind2.Enabled = True 'Added by Morgan 2014/3/12
      'Modify by Amy 2021/12/13 原:Enabled = False,導致 已是灰色底的TextBox 顯示的字更淺,字看不清-瑞婷
      txtA2310.Locked = True
      txtTitle.Locked = True 'Added by Morgan 2015/7/23
      'end 2021/12/13
      Command1.Enabled = False 'Added by Morgan 2015/7/23
   Else
      txtA2301.Enabled = False
      cmdFind.Enabled = False
      cmdFind2.Enabled = False 'Added by Morgan 2014/3/12
      'Modify by Amy 2021/12/13 原:Enabled = False,導致 已是灰色底的TextBox 顯示的字更淺,字看不清-瑞婷
      txtA2310.Locked = False
      'Added by Morgan 2015/7/23
      txtTitle.Locked = False
      'end 2021/12/13
      If Me.Visible Then txtTitle.SetFocus
      Command1.Enabled = True
      'end 2015/7/23
   End If
   '智權人員
   txtA2303.Enabled = Not bolValue
   '客戶編號
   txtA2304.Enabled = Not bolValue
   '收款金額
   For ii = 1 To 5
      txtA2306(ii).Enabled = Not bolValue
   Next
   '扣繳金額
   txtA2307.Enabled = Not bolValue
   '繳收據日
   txtA2308.Enabled = False
   
   'Add by Amy 2017/12/07 電匯資料
   'Modify by Amy 2021/12/13 原:Enabled = False,導致 已是灰色底的TextBox 顯示的字更淺,字看不清-瑞婷
   txtA2330.Locked = bolValue
   Command2.Enabled = Not bolValue
   'end 2017/12/07
   
   '暫收的修改
   If p_Status = "2" And txtA2308 = "" Then
      '收據編號
      txtNo.Enabled = True
      '刪除收據編號
      cmdCut.Enabled = True
      '清除收據編號
      cmdClear.Enabled = True
   Else
      '收據編號
      txtNo.Enabled = Not bolValue
      '刪除收據編號
      cmdCut.Enabled = Not bolValue
      '清除收據編號
      cmdClear.Enabled = Not bolValue
   End If
   
   'Removed by Morgan 2015/7/23
   'If txtNo.Enabled = True Then
   '   txtNo.SetFocus
   'ElseIf txtA2310.Enabled = True Then
   '   txtA2310.SetFocus
   'End If
   'end 2015/7/23
   
   'Added by Morgan 2013/4/24
   'If pub_strUserOffice = "1" Then 'Removed by Morgan 2014/1/13 分所也開放輸入電匯
      If txtA2309 = "" Then
         cmdEmail.Enabled = bolValue
      Else
         cmdEmail.Enabled = False
      End If
   
   
      If Val(txtA2306(3)) > 0 Then
         SetOption Not bolValue
      Else
         SetOption False
      End If
   'End If 'Removed by Morgan 2014/1/13 分所也開放輸入電匯
   'end 2013/4/24
   
   'Added by Morgan 2015/7/20
   SetCheckData
   'end 2015/7/20
   
   'Added by Morgan 2025/6/13
   If pub_strUserOffice <> "1" Or p_Status = "0" Then
      Picture1.Enabled = False
      Picture1.BackColor = Color2
      Option1(1).BackColor = Color2
      Option1(2).BackColor = Color2
      Option1(3).BackColor = Color2
      Option1(4).BackColor = Color2
      
   Else
      Picture1.Enabled = True
      Picture1.BackColor = Color1
      Option1(1).BackColor = Color1
      Option1(2).BackColor = Color1
      Option1(3).BackColor = Color1
      Option1(4).BackColor = Color1
   End If
   'end 2025/6/13
   
   'Added by Morgan 2015/7/20
   '會連續新增收款種類,預設前次種類--辜
   If m_Status = "1" Then
      Option1(pub_strUserOffice).Value = True 'Added by Morgan 2025/6/24 所別預設操作人員所別
      If m_LstAddIndex = 0 Then m_LstAddIndex = 1
      For intI = 1 To 5
         If m_LstAddIndex = intI Then
            txtA2306(intI).TabStop = True
            'txtA2306(intI).SetFocus 'Removed by Morgan 2015/7/23 改先輸收據抬頭
         Else
            txtA2306(intI).TabStop = False
         End If
      Next
   End If
   'end 2015/7/20
End Sub

Private Sub GrdDataList_Click()
   Dim ii As Integer, jj As Integer
   With grdDataList
      If .row > 0 And .row < .Rows Then
         txtNo = Mid(.TextMatrix(.row, 0), 2)
      End If
   End With
End Sub

Private Sub Option1_Click(Index As Integer)
   Dim strUserOffice As String
   If m_Status <> 0 Then
      If Val(txtA2306(3)) > 0 Then
         SetOption True
      Else
         SetOption False, True
      End If
      
      If m_lstOption1 <> Index Then
         strUserOffice = pub_strUserOffice
         pub_strUserOffice = Index
         Pub_AccBankTit Combo1, Me.Name
         pub_strUserOffice = strUserOffice
      End If
      
   End If
   m_lstOption1 = Index
   
   If m_Status <> 0 Then
      If Val(txtA2306(3)) > 0 Then
         SetOption True
      End If
   End If
End Sub


'Added by Morgan 2014/3/12
Private Sub txtA2302_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(601) Then Exit Sub 'Add by Amy 2017/12/18
   If txtA2302.Text = MsgText(601) Or txtA2302.Text = MsgText(29) Then
      MsgBox Label3 & MsgText(52), , MsgText(5)
      Cancel = True
      txtA2302.SetFocus
      Exit Sub
   End If
   If DateCheck(txtA2302.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      Cancel = True
      txtA2302.SetFocus
      Exit Sub
   End If
End Sub

Private Sub txtA2301_Change()
   If txtA2301.Enabled = True And txtA2303 <> "" Then
      Me.FormClear False
   End If
End Sub

Private Sub txtA2301_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA2303_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Removed by Morgan 2015/10/26
'Private Sub txtA2304_Validate(Cancel As Boolean)
'   If txtCustomer.Text = "" Then
'      'Modified by Morgan 2013/4/29
'      'If Len(Left(txtA2304 & "000", 9)) = 9 Then
'      '   txtA2304 = Left(txtA2304 & "000", 9)
'      '   txtCustomer.Text = GetCustTitle(txtA2304)
'      'End If
'      If Len(Left(txtA2304 & "000", 8)) = 8 Then
'         txtA2304 = Left(txtA2304 & "000", 8)
'         txtCustomer.Text = GetCustTitle(txtCuCode & Left(txtA2304 & "000", 8))
'      End If
'      'end 2013/4/29
'   End If
'
'End Sub
'Added by Morgan 2015/7/17
Private Sub txtA2306_Change(Index As Integer)
   If Val(txtA2306(Index)) > 0 Then
      For intI = 1 To 5
         If Index <> intI Then
            txtA2306(intI).TabStop = False
         End If
      Next
   End If
   
   If Index = 1 Then
      SetCheckData
   End If
End Sub
'Added by Morgan 2015/7/20
Private Sub SetCheckData()
   If Val(txtA2306(1)) > 0 Then
      txtA2325.Mask = DFormat
      If txtA2306(1).Enabled = True Then
         txtA2325.Enabled = True
         txtA2326.Enabled = True
         txtA2327.Enabled = True
         txtA2328.Enabled = True
      End If
   Else
      txtA2325.Mask = ""
      txtA2325 = ""
      txtA2325.Mask = DFormat
      txtA2325.Enabled = False
      txtA2326 = "": txtA2326.Enabled = False
      txtA2327 = "": txtA2327.Enabled = False
      txtA2328 = "": txtA2328.Enabled = False
   End If
End Sub
Private Sub txtA2306_GotFocus(Index As Integer)
   TextInverse txtA2306(Index)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtA2306(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtA2306_Validate(Index As Integer, Cancel As Boolean)
   If Index <> 0 Then
      txtA2306(0) = Val(txtA2306(1)) + Val(txtA2306(2)) + Val(txtA2306(3)) + Val(txtA2306(4)) + Val(txtA2306(5)) + Val(txtA2307)
      txtDif = Val(txtTot) - Val(txtA2306(0)) 'Add by Morgan 2006/10/17
   End If
   'Added by Morgan 2013/4/25
   'If pub_strUserOffice = "1" Then 'Removed by Morgan 2014/1/13 分所也開放輸入電匯
      If m_Status <> 0 Then
         If Val(txtA2306(3)) > 0 Then
            SetOption True
         Else
            SetOption False, True
         End If
      End If
   'End If'Removed by Morgan 2014/1/13 分所也開放輸入電匯
   'end 2013/4/25
End Sub

Private Sub txtA2307_Validate(Cancel As Boolean)
   txtA2306(0) = Val(txtA2306(1)) + Val(txtA2306(2)) + Val(txtA2306(3)) + Val(txtA2306(4)) + Val(txtA2306(5)) + Val(txtA2307)
   txtDif = Val(txtTot) - Val(txtA2306(0)) 'Add by Morgan 2006/10/17
End Sub

Private Sub txtA2325_Validate(Cancel As Boolean)
    If strSaveConfirm = MsgText(601) Then Exit Sub 'Add by Amy 2017/12/18
   If txtA2325.Text = MsgText(601) Or txtA2325.Text = MsgText(29) Then
      MsgBox Label23 & MsgText(52), , MsgText(5)
      Cancel = True
      txtA2325.SetFocus
      Exit Sub
   End If
   If DateCheck(txtA2325.Text) = MsgText(603) Then
      MsgBox Label23 & MsgText(63), , MsgText(5)
      Cancel = True
      txtA2325.SetFocus
      Exit Sub
   End If
   
   'Added by Morgan 2025/6/13
   '輸入支票若超過票期規定請於備註欄內加上(票期超過規定請轉呈主管簽核)
   If txtA2325 <> "" Then
      'Modified by Morgan 2025/6/18 規定有修正，改用函數
      'strExc(1) = CompDate(1, 2, DBDATE(txtA2302))
      strExc(1) = PUB_GetCheckMaxDate(txtA2302)
      'end 2025/6/18
      If DBDATE(txtA2325) > strExc(1) Then
         If InStr(txtA2310, "票期超過規定，請轉呈主管簽核。") = 0 Then
            txtA2310 = "票期超過規定，請轉呈主管簽核。" & txtA2310
         End If
      Else
         If InStr(txtA2310, "票期超過規定，請轉呈主管簽核。") > 0 Then
            txtA2310 = Replace(txtA2310, "票期超過規定，請轉呈主管簽核。", "")
         End If
      End If
   End If
   'end 2025/6/13
End Sub

Private Sub txtA2326_GotFocus()
   TextInverse txtA2326
End Sub

Private Sub txtA2326_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA2327_Change()
   txtBankName = ""
End Sub

Private Sub txtA2327_GotFocus()
   TextInverse txtA2327
End Sub

Private Sub txtA2327_Validate(Cancel As Boolean)
   txtBankName = A0g02Query(txtA2327)
End Sub

Private Sub txtA2328_GotFocus()
   TextInverse txtA2328
End Sub

'Add by Amy 2017/12/18 設定Command Default
Private Sub txtA2330_GotFocus()
    Command2.Default = True
    TextInverse txtTitle
    OpenIme
End Sub

Private Sub txtTitle_GotFocus()
   Command1.Default = True 'Add by Amy 2017/12/18 設定Command Default
   TextInverse txtTitle
   OpenIme
End Sub

Private Sub txtNo_GotFocus()
   TextInverse txtNo
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtNo.IMEMode = 2
   CloseIme
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA2303_Change()
   If Len(txtA2303) > 4 Then
      txtSales.Text = StaffQuery(txtA2303)
   Else
      txtSales.Text = ""
   End If
End Sub

Private Sub txtA2303_GotFocus()
   TextInverse txtA2303
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtA2303.IMEMode = 2
   CloseIme
End Sub

Private Sub txtA2304_Change()
   'Modify by Amy 2022/06/27 若輸抬頭(墨)搜尋,抬頭不是要的資料,直接輸客戶編號(X84479000),收據抬頭.Tag不會被清mail會帶錯
   txtTitle = "": txtTitle.Tag = ""
   'Modified by Morgan 2013/4/29
   'If Len(txtA2304) > 8 Then
   '   txtCustomer.Text = GetCustTitle(txtA2304)
   If Len(txtA2304) > 8 Then
      'Modified by Morgan 2015/7/23 新增抬頭查詢功能,改放客戶名稱
      'txtCustomer.Text = GetCustTitle(txtCuCode & txtA2304)
      txtCustomer.Text = GetCustomerName(txtA2304)
   Else
      txtCustomer.Text = ""
   End If
  
End Sub

Private Sub txtA2304_GotFocus()
   TextInverse txtA2304
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtA2304.IMEMode = 2
   CloseIme
End Sub
'Removed by Morgan 2015/10/26
''取客戶最近收據抬頭
'Private Function GetCustTitle(ByVal p_A0k03 As String) As String
'
'   p_A0k03 = Left(p_A0k03 & "000", 9)
'   strSql = "Select NVL(A0K04,CU04) From CUSTOMER,ACC0K0 Where CU01='" & Left(p_A0k03, 8) & "' AND CU02='" & Right(p_A0k03, 1) & "' AND A0K03(+)=CU01||CU02 Order By A0K02 Desc "
'
'On Error GoTo ErrHnd
'
'   CheckOC3
'   With AdoRecordSet3
'      .CursorLocation = adUseClient
'      .MaxRecords = 1
'      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
'      If .RecordCount > 0 Then
'         GetCustTitle = "" & .Fields(0)
'      End If
'ErrHnd:
'      CheckOC3
'      .MaxRecords = 0
'   End With
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
'
'End Function

Private Sub txtA2304_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Function CheckOffDutySales() As Boolean
   strExc(0) = "select st04,st02 from staff where st01='" & txtA2303 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp(0) = "2" Then
         CheckOffDutySales = True
      End If
   End If
End Function

Private Function CheckData() As Boolean
   Dim ii As Integer
   Dim bCancel As Boolean
   CheckData = True
   
   '重新計算繳款金額
   txtA2306_Validate 1, False
   
   '收款金額
   If Val(txtA2306(0)) = 0 Then
      CheckData = False
      MsgBox Label12 & "必須大於0！", vbExclamation
      txtA2306(1).Enabled = True
      txtA2306(1).SetFocus
      txtA2306_GotFocus 1
      Exit Function
   End If
   '扣繳金額
   txtA2307 = Val(txtA2307)
   
   '智權人員
   If txtA2303 = "" Then
      CheckData = False
      MsgBox Label9 & "不可空白！", vbExclamation
   ElseIf txtSales = "" Then
      CheckData = False
      MsgBox Label9 & "輸入錯誤！", vbExclamation
   
   'Added by Morgan 2015/2/5
   '必須為在職員工
   'Modified by Morgan 2015/8/5 改提醒可繼續
   'ElseIf PUB_GetStaffNameDept(txtA2303, "", "", True) = False Then
   ElseIf CheckOffDutySales() = True Then
      If MsgBox(txtSales & "已離職！是否確定要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         CheckData = False
      End If
   'end 2015/8/5
   End If
   'end 2015/2/5
   
   If CheckData = False Then
      txtA2303.SetFocus
      txtA2303_GotFocus
      Exit Function
   End If
   '客戶
   If txtA2304 <> "" Then
      'Call txtA2304_Validate(False) 'Removed by Morgan 2015/10/26
      If txtCustomer = "" Then
         CheckData = False
         MsgBox Label1 & "輸入錯誤！", vbExclamation
      End If
   'Add by Moragn 2005/4/18
   'Else
   ElseIf grdDataList.TextMatrix(1, 0) <> "" Then
      CheckData = False
      MsgBox "已繳收據，" & Label1 & "不可空白！", vbExclamation
   End If

   
   If CheckData = False Then
      txtA2304.SetFocus
      txtA2304_GotFocus
   End If
   
   'Added by Morgan 2013/4/29
   'If pub_strUserOffice = "1" Then 'Removed by Morgan 2014/1/13 分所也開放輸入電匯
      If Val(txtA2306(3)) > 0 Then
         'Modified by Morgan 2013/12/23
         'If Option1(0).Value = False And Option1(1).Value = False Then
         '   MsgBox "銀存請點瑞興或華銀！", vbExclamation
         If Combo1.Text = "" Then
            MsgBox "請點選科目代碼！", vbExclamation
            CheckData = False
            
         ElseIf Option2(0).Value = False And Option2(1).Value = False Then
            CheckData = False
            MsgBox "銀存請點選電匯或提款機！", vbExclamation
         
         ElseIf txtA2310 = "" Then
            strExc(1) = PUB_GetST03(txtA2303)
            'Modified by Morgan 2014/1/23 +分所出納
            If strExc(1) = "M31" Or strExc(1) = "M71" Then
               CheckData = False
               MsgBox "銀存待認領請於備註欄輸入" & IIf(Option2(0).Value, "匯款人", "匯款帳號尾碼") & "！", vbExclamation
               txtA2310.SetFocus
            End If
         End If
      End If
   'End If 'Removed by Morgan 2014/1/13 分所也開放輸入電匯
   'end 2013/4/29
   
   'Added by Morgan 2015/7/20
   '票據檢查
   If Val(txtA2306(1)) > 0 Then
      txtA2325_Validate bCancel
      If bCancel = True Then
         CheckData = False
         If txtA2325.Enabled Then txtA2325.SetFocus
         Exit Function
      End If
      If txtA2326 = "" Then
         MsgBox "請輸入票號！", vbExclamation
         CheckData = False
         If txtA2326.Enabled Then txtA2326.SetFocus
         Exit Function
      End If
      If txtA2327 = "" Then
         MsgBox "請輸入收票銀行！", vbExclamation
         CheckData = False
         If txtA2327.Enabled Then txtA2327.SetFocus
         Exit Function
      Else
         txtA2327_Validate bCancel
         If txtBankName = "" Then
            MsgBox "收票銀行錯誤！", vbExclamation
            CheckData = False
            If txtA2327.Enabled Then txtA2327.SetFocus
            Exit Function
         End If
      End If
      If txtA2328 = "" Then
         MsgBox "請輸入收票帳號！", vbExclamation
         CheckData = False
         If txtA2328.Enabled Then txtA2328.SetFocus
         Exit Function
      End If
      
      If Check1.Value = 0 Then 'Added by Morgan 2015/9/3 非整批收款後要能手動設已處理
         If txtA2326 & txtA2327 & txtA2328 <> txtA2326.Tag & txtA2327.Tag & txtA2328.Tag Then 'Added by Morgan 2022/6/1 修改時若票據資料沒改不要檢查，否則都不會過
            strExc(0) = "select a0e02 from acc0e0 where a0e02 = '" & txtA2326 & "' and a0e01 = '" & txtA2327 & "' and a0e07 = '" & txtA2328 & "' and a0e04 = 'R' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "票號已存在!", vbExclamation
               CheckData = False
               If txtA2326.Enabled Then txtA2326.SetFocus
               Exit Function
            End If
         End If
      End If
   End If
   'end 2015/7/20
   
   If CheckData = False Then
      Exit Function
   End If
   
End Function

Public Function SaveData() As Boolean
   
   Dim stA23(1 To 30) As String 'Modify by Amy 2017/12/07 +電匯資料
   Dim ii As Integer, iYear As Integer, iMonth As Integer
   Dim stCols As String, stValues As String, bolMail As Boolean
   
   If CheckData = False Then Exit Function
   
On Error GoTo ErrHnd

   adoTaie.BeginTrans 'Added by Morgan 2018/2/9
   
   Erase stA23
   stA23(1) = txtA2301
   stA23(2) = Val(ChangeTDateStringToTString(txtA2302.Text))
   stA23(2) = IIf(stA23(2) = 0, stA23(2) = "NULL", stA23(2))
   stA23(3) = "'" & txtA2303 & "'"
   'Modified by Morgan 2013/4/29
   'stA23(4) = "'" & txtA2304 & "'"
   If txtA2304 <> "" Then
      stA23(4) = "'" & txtA2304 & "'"
   Else
      stA23(4) = "''"
   End If
   'end 2013/4/29
   
   
   'Modified by Morgan 2015/6/23
   'Removed by Morgan 2025/6/16
   'If stA23(1) = "" Then
   '   stA23(5) = pub_strUserOffice
   'Else
   '   stA23(5) = "A2305"
   'End If
   'end 2025/6/16
   'end 2015/6/23
   
   'Added by Morgan 2023/7/13
   'Modified by Morgan 2025/6/16
   'strExc(2) = Left(Combo1.Text, 4)
   'If strExc(2) = "1911" Then
   '   stA23(5) = "2"
   'ElseIf strExc(2) = "1912" Then
   '   stA23(5) = "3"
   'ElseIf strExc(2) = "1913" Then
   '   stA23(5) = "4"
   'End If
   If Option1(1).Value = True Then
      stA23(5) = "1"
   ElseIf Option1(2).Value = True Then
      stA23(5) = "2"
   ElseIf Option1(3).Value = True Then
      stA23(5) = "3"
   ElseIf Option1(4).Value = True Then
      stA23(5) = "4"
   End If
   'end 2025/6/16
   'end 2023/7/13
   
   stA23(6) = Val(txtA2306(1))
   stA23(7) = Val(txtA2307)
   
   'Removed by Morgan 2016/6/17 已改由智權人員繳款時回寫,此處取消(可能會因時間差導致繳款資料被清除)
   'stA23(8) = Val(ChangeTDateStringToTString(txtA2308))
   'stA23(8) = IIf(stA23(8) = 0, "NULL", stA23(8))
   'If grdDataList.TextMatrix(1, 0) <> "" Then
   '   grdDataList.col = 0
   '   grdDataList.Sort = flexSortGenericAscending
   '   stA23(9) = grdDataList.TextMatrix(1, 0)
   '   For ii = 2 To grdDataList.Rows - 1
   '      If grdDataList.TextMatrix(ii, 0) <> "" Then
   '         stA23(9) = stA23(9) & "," & grdDataList.TextMatrix(ii, 0)
   '      End If
   '   Next ii
   'End If
   'txtA2309 = stA23(9)
   'stA23(9) = CNULL(stA23(9))
   'end 2016/6/17
   
   stA23(10) = "'" & ChgSQL(txtA2310) & "'"
   stA23(17) = Val(txtA2306(2))
   stA23(18) = Val(txtA2306(3))
   stA23(19) = Val(txtA2306(4))
   stA23(20) = Val(txtA2306(5))
   
   'Added by Morgan 2015/7/20
   stA23(25) = Val(ChangeTDateStringToTString(txtA2325))
   stA23(26) = "'" & txtA2326 & "'"
   stA23(27) = "'" & txtA2327 & "'"
   stA23(28) = "'" & txtA2328 & "'"
   'end 2015/7/20
   stA23(30) = IIf(txtA2330 = "", "NULL", "'" & ChgSQL(txtA2330) & "'") 'Add by Amy 2017/12/07
   
   'Added by Morgan 2013/4/25
   'Modified by Morgan 2013/12/23
   'If Option1(0).Value = True Then
   '   stA23(22) = "'1'"
   'ElseIf Option1(1).Value = True Then
   '   stA23(22) = "'2'"
   'Else
   '   stA23(22) = "''"
   'End If
   If Combo1 <> "" Then
      stA23(22) = "'" & Left(Combo1.Text, InStr(Combo1, " ") - 1) & "'"
      m_LstIndex = Combo1.ListIndex
   Else
      stA23(22) = "''"
   End If
   'end 2013/12/23
   
   If Option2(0).Value = True Then
      stA23(23) = "'1'"
   ElseIf Option2(1).Value = True Then
      stA23(23) = "'2'"
   Else
      stA23(23) = "''"
   End If
   
   'end 2013/4/25
   '新增
   If stA23(1) = "" Then
      'Modified by Morgan 2015/7/22 票據/現金也要EMail通知
      If Val(txtA2306(1)) + Val(txtA2306(2)) + Val(txtA2306(3)) > 0 Then bolMail = True 'Added by Morgan 2013/5/9
      'end 2015/7/22
      iYear = (stA23(2) \ 10000)
      iMonth = ((stA23(2) \ 100) Mod 100)
      stA23(1) = AccAutoNo("S", 4, iYear, iMonth)
      AccSaveAutoNo "S", Right(stA23(1), 4), iYear, iMonth 'Added by Morgan 2019/1/7
      stA23(11) = "'" & strUserNum & "'"
      stA23(12) = "to_number(to_char(sysdate,'YYYYMMDD'))"
      stA23(13) = "to_number(to_char(sysdate,'HH24MI'))"
      
      stCols = "A2301"
      stValues = "'" & stA23(1) & "'"
      'Modified by Morgan 2013/4/25 +22-23
      'Modified by Morgan 2015/7/20 +25-28
      'Modify by Amy 2017/12/07 +30剔除29
      For ii = 2 To 30
         'Modified by Morgan 2016/6/17 +剔除8,9
         If Not (ii > 13 And ii < 17) And ii <> 21 And ii <> 24 And ii <> 8 And ii <> 9 And ii <> 29 Then
            stCols = stCols & ",A23" & Format(ii, "00")
            stValues = stValues & "," & stA23(ii)
         End If
      Next
      strSql = "insert into ACC230 (" & stCols & ") VALUES (" & stValues & ")"
      'AccSaveAutoNo "S", Right(stA23(1), 4), iYear, iMonth 'Removed by Morgan 2019/1/7 移到上面
      
      'Added by Morgan 2015/7/20
      For ii = 1 To 5
         If Val(txtA2306(ii)) > 0 Then
            m_LstAddIndex = ii
            Exit For
         End If
      Next
      'end 2015/7/20
   '修改
   Else
      stA23(14) = "'" & strUserNum & "'"
      stA23(15) = "to_number(to_char(sysdate,'YYYYMMDD'))"
      stA23(16) = "to_number(to_char(sysdate,'HH24MI'))"
      'Modified by Morgan 2014/3/12
      'strSql = "Update ACC230 Set A2303=" & stA23(3)
      strSql = "Update ACC230 Set A2302=" & stA23(2) & ",A2303=" & stA23(3)
      'end 2014/3/12
      
      'Modified by Morgan 2013/4/25 +22-23
      'Modified by Morgan 2015/7/20 +25-28
      'Modify by Amy 2017/12/07 +30剔除29
      For ii = 4 To 30
         'Modified by Morgan 2016/6/7  修改記錄不可略過,+剔除8,9
         'If Not (ii > 10 And ii < 17) And ii <> 21 And ii <> 24 Then
         If Not (ii > 10 And ii < 14) And ii <> 21 And ii <> 24 And ii <> 8 And ii <> 9 And ii <> 29 Then
            strSql = strSql & ",A23" & Format(ii, "00") & "=" & stA23(ii)
         End If
      Next
      
      'Added by Morgan 2015/6/17
      If Check1.Value = 1 And txtA2321 = "" Then
         strSql = strSql & ",A2321=to_date(19221111,'yyyymmdd')"
      ElseIf Check1.Value = 0 And txtA2321 <> "" Then
         strSql = strSql & ",A2321=null"
      End If
      'end 2015/6/17
      
      strSql = strSql & " Where A2301='" & stA23(1) & "'"
   End If
   
   adoTaie.Execute strSql
   txtA2301 = stA23(1)
   adoTaie.CommitTrans 'Added by Morgan 2018/2/9
   SaveData = True
   
   'Modified by Morgan 2019/5/24 南所發生有 EMail 但查無資料情形,改從 Transaction 內移出來
   'Modified by Morgan 2014/1/23 不限制只有北所
   'If pub_strUserOffice = "1" And bolMail = True Then Mail2Sales True 'Added by Morgan 2013/5/9
   If bolMail = True Then Mail2Sales True
   'end 2019/5/24
   
ErrHnd:
   If Err.Number <> 0 Then
      adoTaie.RollbackTrans
      MsgBox Err.Description, vbCritical
      'adoTaie.BeginTrans 'Removed by Morgan 2018/2/9
   End If
   
End Function

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyInsert
         If txtNo.Enabled = True And txtNo.Text <> "" Then
            If ReadGridData(txtCode & txtNo, True) = True Then
               txtNo.Text = Left(txtNo, 3) & Format(Val(Right(txtNo.Text, 5)) + 1, "00000")
               txtNo_GotFocus
               grdDataList.row = 1
            End If
         End If
      Case Else
         KeyEnter KeyCode
   End Select
End Sub

Private Sub txtA2307_GotFocus()
   TextInverse txtA2307
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtA2307.IMEMode = 2
   CloseIme
End Sub

Private Sub SetOption(bEnable As Boolean, Optional bClear As Boolean)
   Dim lColor As Long
   If bEnable Then
      lColor = Color1
   Else
      lColor = Color2
   End If
   
   'Modified by Morgan 2013/12/23
   'Picture1.Enabled = bEnable
   'Picture1.BackColor = lColor
   'Option1(0).BackColor = lColor
   'Option1(1).BackColor = lColor
   'Modified by Morgan 2014/3/12
   'Combo1.Enabled = bEnable
   Combo1.Enabled = bEnable Or cmdFind2.Enabled
   'end2014/3/12
   'end 2013/12/23
   
   Picture2.Enabled = bEnable
   Picture2.BackColor = lColor
   Option2(0).BackColor = lColor
   Option2(1).BackColor = lColor
  
   If bClear = True Then
      'Modified by Morgan 2013/12/23
      'Option1(0).Value = False
      'Option1(1).Value = False
      Combo1.ListIndex = -1
      'end 2013/12/23
      Option2(0).Value = False
      Option2(1).Value = False
   End If
   
   'Added by Morgan 2013/5/3 預設瑞興,電匯--瑞婷
   If bEnable = True Then
      'Modified by Morgan 2013/12/23
      'If Option1(0).Value = False And Option1(1).Value = False Then
      '   Option1(0).Value = True
      'End If
      If Combo1.ListIndex = -1 Then
         'Modified by Morgan 2025/6/13
         'If pub_strUserOffice = "1" Then
         If m_lstOption1 = "1" Then
         'end 2025/6/13
            'Added by Morgan 2014/1/13 預設前次存檔選項
            If m_LstIndex > -1 Then
               Combo1.ListIndex = m_LstIndex
            Else
            'end 2014/1/13
               SetCombo1 "110602"  'modify by sonia 2020/12/24 原為110202，婉莘說改110602
            End If 'Added by Morgan 2014/1/13
         'Added by Morgan 2014/1/13 分所也開放輸入電匯
         ElseIf Combo1.ListCount > 0 Then
            Combo1.ListIndex = 0
         'end 2014/1/13
         End If
      End If
      'end 2013/12/23
      
      If Option2(0).Value = False And Option2(1).Value = False Then
         Option2(0).Value = True
      End If
   End If
   'End 2013/5/3
End Sub

Private Sub Mail2Sales(pAutoMail As Boolean)
   
   Dim ii As Integer, bolNonSales As Boolean
   Dim strSubject As String, strContent As String
   Dim strReceiverName As String, strReceiverID As String
   Dim strCopyName As String, strCopyID As String
   Dim stNonSalesList As String
   Dim stST06 As String
   'add by sonia 2017/8/24 增加簽收固定不通知清單
   Dim stNoticeList  As String
   Dim stNotNoticeList  As String
   Dim i As Integer, j As Integer, varTmp1 As Variant, varTmp2 As Variant, InList As Boolean
   'end 2017/8/24
   Dim strTrueRecID
   Dim stST06_1 As String 'Added by Morgan 2023/2/7
   
On Error GoTo ErrHnd
   'Mark by Amy 2020/06/08 取消-辜
   'Add by Amy 2020/04/23 +if 銀存科目110502 不發mail-辜
'   If Trim(Combo1.Text) <> MsgText(601) Then
'        If Left(Combo1.Text, InStr(Combo1, " ") - 1) = "110502" Then Exit Sub
'   End If
   
   strContent = ""
   strExc(1) = PUB_GetST03(txtA2303)
   
   stST06 = PUB_GetST06(txtA2303) 'Added by Morgan 2014/12/11
   
   'Modified by Morgan 2015/10/12 --辜
   'strSubject = "客戶"
   If txtA2304 <> "" Then
      strSubject = txtCustomer
   ElseIf txtA2310 <> "" Then
      strSubject = txtA2310
   'cancel by sonia 2017/8/24
   'Else
   '   strSubject = "客戶"
   'end 2017/8/24
   End If
   'end 2015/10/12
   
   If Val(txtA2306(1)) > 0 Then
      strSubject = strSubject & " 票據 "
   End If
   
   If Val(txtA2306(2)) > 0 Then
      strSubject = strSubject & " 現金 "
   End If
   
   If Val(txtA2306(3)) > 0 Then
      strSubject = strSubject & " 匯款 "
   End If
   
   'Modified by Morgan 2014/1/23 +分所出納
   If strExc(1) = "M31" Or strExc(1) = "M71" Then
      bolNonSales = True
      
      'Added by Morgan 2023/2/7 認領改以銀存科目判斷通知的所別
      stST06_1 = PUB_GetST06(txtA2303)
      If Combo1.Text <> "" Then
         'Modified by Morgan 2025/6/16
         'strExc(2) = Left(Combo1.Text, InStr(Combo1, " ") - 1)
         strExc(2) = Left(Combo1.Text, 4)
         'end 2025/6/16
      Else
         strExc(2) = ""
      End If
      '中所
      If strExc(2) = "1911" Then
         stST06 = "2"
      '南所
      ElseIf strExc(2) = "1912" Then
         stST06 = "3"
      '高所
      ElseIf strExc(2) = "1913" Then
         stST06 = "4"
      '智權公司與法律所認領時發全所
      ElseIf strExc(2) = "110303" Or strExc(2) = "110502" Then
         stST06 = "0"
      Else
         stST06 = stST06_1
      End If
      'end 2023/2/7
   
      'modify by sonia 2017/9/1 取消前面其他字眼
      'strSubject = strSubject & "請認領　不明匯入款"
      strSubject = strSubject & "請認領　不明匯入款"
      
      'Modified by Morgan 2014/12/11
      'Mail財務/出納人員所別的所有智權人員
      'If pub_strUserOffice = "1" Then
      If stST06 = "1" Then
      'end 2014/12/11
      
         'Modified by Morgan 2013/5/20
         'strReceiverName = "智權部北所全部人員"
         'strReceiverID = "sale_taipei@taie.com.tw"
         strReceiverName = "智權部北所各區人員"
         strReceiverID = "sale_s1@taie.com.tw"
         
      'Added by Morgan 2023/2/7
      ElseIf stST06 = "0" Then
         strReceiverName = "智權部人員"
         strReceiverID = "sales@taie.com.tw"
         
      'Added by Morgan 2014/1/23
      Else
         'Modified by Morgan 2014/12/11
         'strExc(0) = "select st01,st02 from staff where st06='" & pub_strUserOffice & "' and st15 like 'S%' and st04='1' and st01>'6' and st01<'F' order by 1"
         strExc(0) = "select st01,st02 from staff where st06='" & stST06 & "' and st15 like 'S%' and st04='1' and st01>'6' and st01<'F' order by 1"
         'end 2014/12/11
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            Do While Not .EOF
               strReceiverName = strReceiverName & ";" & .Fields("st02")
               strReceiverID = strReceiverID & ";" & .Fields("st01")
               .MoveNext
            Loop
            End With
            strReceiverName = Mid(strReceiverName, 2)
            strReceiverID = Mid(strReceiverID, 2)
         End If
      End If
      'end 2014/1/23
      
      'Modified by Morgan 2013/8/30
      'stNonSalesList = cNonSalesList
      stNonSalesList = Pub_GetSpecMan("簽收非智權通知清單")
      stNotNoticeList = Pub_GetSpecMan("簽收固定不通知清單")  'add by sonia 2017/8/24
      
      'add by sonia 2017/6/27 再加有國內應收帳款之非智權部人員
      'modify by sonia 2018/12/22專利處人員都不通知
      'strExc(0) = "SELECT A0K20 FROM STAFF,(SELECT DISTINCT A0K20 FROM ACC0K0 WHERE A0K37 IS NULL AND NVL(A0K09,0)=0 AND A0K02<>0)" & _
                  "WHERE A0K20=ST01(+) AND ST04='1' AND (ST14 IS NULL OR ST14<>'99997') AND SUBSTR(ST15,1,1)<>'S' ORDER BY 1"
      strExc(0) = "SELECT A0K20 FROM STAFF,(SELECT DISTINCT A0K20 FROM ACC0K0 WHERE A0K37 IS NULL AND NVL(A0K09,0)=0 AND A0K02<>0)" & _
                  "WHERE A0K20=ST01(+) AND ST04='1' AND (ST14 IS NULL OR ST14<>'99997') AND SUBSTR(ST15,1,1)<>'S' AND SUBSTR(ST15,1,2)<>'P1' ORDER BY 1"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            stNonSalesList = stNonSalesList & ";" & .Fields("A0K20")
            .MoveNext
         Loop
         End With
      End If
      'end 2017/6/27
      
      'add by sonia 2017/8/24 剔除簽收固定不通知清單
      stNoticeList = ""
      If stNonSalesList <> "" And stNotNoticeList <> "" Then
         varTmp1 = Split(stNonSalesList, ";")
         varTmp2 = Split(stNotNoticeList, ";")
         For i = 0 To UBound(varTmp1)
            InList = False
            For j = 0 To UBound(varTmp2)
               If varTmp1(i) = varTmp2(j) Then
                  InList = True
               End If
            Next
            If InList = False Then
               stNoticeList = stNoticeList & ";" & varTmp1(i)
            End If
         Next
      ElseIf stNotNoticeList = "" Then
         stNoticeList = stNonSalesList
      End If
      'end 2017/8/24
      
      'Added by Morgan 2023/2/26
      '非法律所帳號時排除法律所人員
      If strExc(2) <> "110502" Then
         strExc(0) = "select st01 from staff where instr('" & stNoticeList & "',st01)>0 and st93 not like 'L%'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            stNonSalesList = RsTemp.GetString(, , , ";")
            If Right(stNonSalesList, 1) = ";" Then stNonSalesList = Left(stNonSalesList, Len(stNonSalesList) - 1)
         End If
      End If
      'end 2023/2/26
      
      'modify by sonia 改用新變數stNonSalesList->stNoticeList
      'If stNonSalesList <> "" Then
      If stNoticeList <> "" Then
         'Modified by Morgan 2014/12/11
         'strExc(0) = "select st01,st02 from staff where st01 in ('" & Replace(stNonSalesList, ";", "','") & "') and st06='" & pub_strUserOffice & "' and st04='1' order by 1"
         'modify by sonia 2017/6/28 取消離職條件,才會發給主管
         'strExc(0) = "select st01,st02 from staff where st01 in ('" & Replace(stNonSalesList, ";", "','") & "') and st06='" & stST06 & "' and st04='1' order by 1"
         'modify by sonia 改用新變數stNonSalesList->stNoticeList
         'strExc(0) = "select st01,st02 from staff where st01 in ('" & Replace(stNonSalesList, ";", "','") & "') and st06='" & stST06 & "' order by 1"
         'Modified by Morgan 2023/2/7 +全所的判斷
         'strExc(0) = "select st01,st02 from staff where st01 in ('" & Replace(stNoticeList, ";", "','") & "') and st06='" & stST06 & "' order by 1"
         strExc(0) = "select st01,st02 from staff where st01 in ('" & Replace(stNoticeList, ";", "','") & "')" & IIf(stST06 <> "0", " and st06='" & stST06 & "'", "") & " order by 1"
         'end 2014/12/11
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            Do While Not .EOF
               strReceiverName = strReceiverName & ";" & .Fields("st02")
               strReceiverID = strReceiverID & ";" & .Fields("st01")
               .MoveNext
            Loop
            End With
         End If
      End If
      
   '其他
   Else
      bolNonSales = False
      strSubject = strSubject & "通知"
      strReceiverName = txtSales
      strReceiverID = txtA2303
      'Added by Morgan 2015/8/5 +若智權人員已離職系統會自動改寄主管
      If CheckOffDutySales() = True Then
         strContent = "智權人員：" & txtSales
      End If
      'end 2015/8/5
      
      'Added by Morgan 2022/9/16
      '特殊收件者彈提醒
      strExc(0) = ChkSpecMailReciver(strReceiverID, "", True)
      'Modified by Morgan 2023/5/10 回傳值最後不會再加分號
      'If strReceiverID & ";" <> strExc(0) Then
      If strReceiverID <> strExc(0) Then
      'end 2023/5/10
         strExc(1) = strReceiverID & "(" & strReceiverName & ")有特殊設定，實際會通知下列人員：" & vbCrLf & vbCrLf & PUB_ReadUserData(strExc(0))
         MsgBoxU strExc(1), vbInformation
      End If
      'end 2022/9/16
      
   'Modified by Morgan 2014/12/11
   '   stST06 = PUB_GetST06(txtA2303)
   End If
      'Added by Morgan 2023/2/7
      If stST06 = "0" Then
         If stST06_1 <> "1" Then
            strExc(1) = Pub_GetSpecMan("財務處出納人員")
            strExc(2) = GetPrjSalesNM(strExc(1))
            
            strCopyID = strCopyID & strExc(1) & ";"
            strCopyName = strCopyName & strExc(2) & ";"
         End If
         If stST06_1 <> "2" Then
            strExc(1) = Pub_GetSpecMan("出納人員-中所")
            strExc(2) = GetPrjSalesNM(strExc(1))
            
            strCopyID = strCopyID & strExc(1) & ";"
            strCopyName = strCopyName & strExc(2) & ";"
         End If
         If stST06_1 <> "3" Then
            strExc(1) = Pub_GetSpecMan("出納人員-南所")
            strExc(2) = GetPrjSalesNM(strExc(1))
            
            strCopyID = strCopyID & strExc(1) & ";"
            strCopyName = strCopyName & strExc(2) & ";"
         End If
         If stST06_1 <> "4" Then
            'Modified by Morgan 2023/6/30 玉瑛留停3個月,暫改 A8029 呂麗君
            'Modified by Morgan 2023/10/4 玉瑛復職
            strExc(1) = Pub_GetSpecMan("出納人員-高所")
            strExc(2) = GetPrjSalesNM(strExc(1))
            strCopyID = strCopyID & strExc(1) & ";"
            strCopyName = strCopyName & strExc(2) & ";"
         End If
         If Right(strCopyID, 1) = ";" Then strCopyID = Left(strCopyID, Len(strCopyID) - 1)
         If Right(strCopyName, 1) = ";" Then strCopyName = Left(strCopyName, Len(strCopyName) - 1)
      Else
      'end 2023/2/7
      
         'Modified by Morgan 2014/1/23 輸入人員所別與智權人員不同時通知
         If pub_strUserOffice <> stST06 Then
            '中南高要寄副本給助理
            '2015/5/26 MODIFY BY SONIA 各所出納人員改抓系統特殊設定
            If stST06 = "2" Then
               'strCopyName = "江文鸞"
               'strCopyID = "85003"
               strCopyID = Pub_GetSpecMan("出納人員-中所")
               strCopyName = GetPrjSalesNM(strCopyID)
            ElseIf stST06 = "3" Then
               'strCopyName = "唐惠琴"
               'strCopyID = "71002"
               strCopyID = Pub_GetSpecMan("出納人員-南所")
               strCopyName = GetPrjSalesNM(strCopyID)
            ElseIf stST06 = "4" Then
               'strCopyName = "余玉瑛"
               'strCopyID = "68008"
               'Modified by Morgan 2023/6/30 玉瑛留停3個月,暫改 A8029 呂麗君
               'Modified by Morgan 2023/10/4 玉瑛復職還原
               strCopyID = Pub_GetSpecMan("出納人員-高所")
               strCopyName = GetPrjSalesNM(strCopyID)
               'end 2023/6/30
            Else
               'strCopyName = "辜苑琪"
               'strCopyID = "71005"
               strCopyID = Pub_GetSpecMan("財務處出納人員")
               strCopyName = GetPrjSalesNM(strCopyID)
            End If
         End If
         
      End If 'Added by Morgan 2023/2/7
   'End If 'Removed by Morgan 2014/12/11
   
   '銀存
   If Val(txtA2306(3)) > 0 Then
      If strContent <> "" Then strContent = strContent & vbCrLf & vbCrLf
      'Modified by Morgan 2013/12/23
      'If Option1(0).Value = True Then
      '   strContent = "瑞興"
      'ElseIf Option1(1).Value = True Then
      '   strContent = "華銀"
      'End If
      'Modified by Morgan 2015/3/2 改中文都帶
      'strContent = Mid(Combo1, 8, 2)
      strContent = strContent & Mid(Combo1, 8)
      'end 2013/12/23
      If Option2(0).Value = True Then
         strContent = strContent & "-電匯"
      ElseIf Option2(1).Value = True Then
         strContent = strContent & "-提款機"
      End If
      If strContent <> "" Then
         strContent = "(" & strContent & ")"
      End If
      
      strContent = strContent & vbCrLf & "匯款日：" & txtA2302
      '電匯
      If Option2(0).Value = True Then
         '2015/3/11 modify by sonia 加傳備註
         'strContent = strContent & _
         '   vbCrLf & "匯款人："
         'If txtA2304 <> "" Then
         '   strContent = strContent & txtCustomer
         'Else
         '   strContent = strContent & txtA2310
         'End If
         'Modify by Amy 2017/12/25 有電匯資料且智權非財務或分所出納者顯示客戶資料,有收抬頭者顯示收據抬頭
'         If txtA2304 <> "" And txtA2310 <> "" Then
'            strContent = strContent & vbCrLf & "客　戶：" & txtCustomer
'            strContent = strContent & vbCrLf & "匯款人：" & txtA2310
'         ElseIf txtA2304 <> "" Then
'            strContent = strContent & vbCrLf & "客　戶：" & txtCustomer
'         Else
'            strContent = strContent & vbCrLf & "匯款人：" & txtA2310
'         End If
        If bolNonSales = True Then
           strContent = strContent & vbCrLf & "電匯資料：" & txtA2330
        ElseIf txtA2330 <> MsgText(601) And bolNonSales = False Then
           strContent = strContent & vbCrLf & "客　戶：" & txtCustomer
           strContent = strContent & vbCrLf & "電匯資料：" & txtA2330
        'Modify by Amy 2022/06/27 原:txtTitle.Tag
        ElseIf txtTitle <> MsgText(601) Then
           strContent = strContent & vbCrLf & "匯款人：" & txtTitle
        'end 2022/06/27
        End If
        'end 2017/12/25
         '2015/3/11 end
      End If
      
      strContent = strContent & vbCrLf & "金　額：" & Format(txtA2306(3), DDollar)
      
      If bolNonSales = True Then
         strContent = strContent & vbCrLf & vbCrLf & "請同仁認領！"
         If Option2(1).Value = True Then
            strContent = strContent & _
               vbCrLf & vbCrLf & "本次匯款帳號尾數為：" & txtA2310 & _
               vbCrLf & "請詢問客戶帳號尾數5碼以確認款項是否是您的,以免繳錯。"
         End If
      End If
   End If
   
   'Added by Morgan 2015/7/22
   '票據
   If Val(txtA2306(1)) > 0 Then
      If strContent <> "" Then strContent = strContent & vbCrLf & vbCrLf
      
      strContent = strContent & "(票據)" & vbCrLf & "繳款日：" & txtA2302
      'Modify by Amy 2017/12/25 有電匯資料且智權非財務或分所出納者顯示客戶資料,有收抬頭者顯示收據抬頭
'      If txtA2304 <> "" And txtA2310 <> "" Then
'         strContent = strContent & vbCrLf & "客　戶：" & txtCustomer
'         strContent = strContent & vbCrLf & "付款人：" & txtA2310
'      ElseIf txtA2304 <> "" Then
'         strContent = strContent & vbCrLf & "客　戶：" & txtCustomer
'      Else
'         strContent = strContent & vbCrLf & "付款人：" & txtA2310
'      End If
      If bolNonSales = True Then
         strContent = strContent & vbCrLf & "電匯資料：" & txtA2330
      ElseIf txtA2330 <> MsgText(601) And bolNonSales = False Then
         strContent = strContent & vbCrLf & "客　戶：" & txtCustomer
         strContent = strContent & vbCrLf & "電匯資料：" & txtA2330
      'Modify by Amy 2022/06/27 原:txtTitle.Tag
      ElseIf txtTitle <> MsgText(601) Then
         strContent = strContent & vbCrLf & "付款人：" & txtTitle
      End If
      'end 2017/12/25
      strContent = strContent & vbCrLf & "金額：" & Format(txtA2306(1), DDollar)
      
      strContent = strContent & vbCrLf & "票期：" & txtA2325
      strContent = strContent & vbCrLf & "票號：" & txtA2326
      strContent = strContent & vbCrLf & "收票銀行：" & txtBankName
      strContent = strContent & vbCrLf & "收票帳號：" & txtA2328
   End If
   'END 2015/7/22
      
   '現金
   If Val(txtA2306(2)) > 0 Then
      If strContent <> "" Then strContent = strContent & vbCrLf & vbCrLf
      
      '2015/3/11 modify by sonia 加傳備註
      'strContent = "(現金)" & _
      '   vbCrLf & "繳款日：" & txtA2302 & _
      '   vbCrLf & "客戶："
      'If txtA2304 <> "" Then
      '   strContent = strContent & txtCustomer
      'Else
      '   strContent = strContent & txtA2310
      'End If
      strContent = strContent & "(現金)" & vbCrLf & "繳款日：" & txtA2302
      'Modify by Amy 2017/12/25 有電匯資料且智權非財務或分所出納者顯示客戶資料,有收抬頭者顯示收據抬頭
'      If txtA2304 <> "" And txtA2310 <> "" Then
'         strContent = strContent & vbCrLf & "客　戶：" & txtCustomer
'         strContent = strContent & vbCrLf & "付款人：" & txtA2310
'      ElseIf txtA2304 <> "" Then
'         strContent = strContent & vbCrLf & "客　戶：" & txtCustomer
'      Else
'         strContent = strContent & vbCrLf & "付款人：" & txtA2310
'      End If
      If bolNonSales = True Then
         strContent = strContent & vbCrLf & "電匯資料：" & txtA2330
      ElseIf txtA2330 <> MsgText(601) And bolNonSales = False Then
         strContent = strContent & vbCrLf & "客　戶：" & txtCustomer
         strContent = strContent & vbCrLf & "電匯資料：" & txtA2330
      'Modify by Amy 2022/06/27 原:txtTitle.Tag
      ElseIf txtTitle <> MsgText(601) Then
         strContent = strContent & vbCrLf & "付款人：" & txtTitle
      End If
      'end 2017/12/25
      '2015/3/11 end
      strContent = strContent & vbCrLf & "金額：" & Format(txtA2306(2), DDollar)
   End If
   
   'Add by Amy 2017/12/25 備註有值需顯示
   If txtA2310 <> MsgText(601) Then
      strContent = strContent & vbCrLf & vbCrLf & "備　註：" & txtA2310
   End If
   
   'Modified by Morgan 2022/9/16
   'ChkSpecMailReciver
   'end 2022/9/16
   
   If pAutoMail = True Then
      PUB_SendMail strUserNum, strReceiverID, "", strSubject, strContent, , , , , , strCopyID, , , , True
      If bolMailSendOk = True Then
         strSql = "update acc230 set a2324=sysdate where a2301='" & txtA2301 & "'"
         cnnConnection.Execute strSql, intI
      End If
   Else
      Frmacc41e2.txtSubject = strSubject
      Frmacc41e2.txtReceiver = strReceiverName
      Frmacc41e2.txtReceiver.Tag = strReceiverID
      Frmacc41e2.txtCopy = strCopyName
      Frmacc41e2.txtCopy.Tag = strCopyID
      Frmacc41e2.txtContent = strContent
      Frmacc41e2.Show vbModal
      strFormName = Me.Name
   End If
   
   strExc(0) = "select a2324 from acc230 where a2301='" & txtA2301 & "' and a2324 is not null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      txtA2324 = Format(RsTemp(0), "EE/MM/DD")
   End If
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

'Add by Amy 2014/04/16 由frmacc42b0過來
Sub StrMenu()
    Me.Height = 3840
    Me.txtA2301.Locked = True
    Me.cmdFind.Visible = False
    Me.cmdFind2.Visible = False
    Me.Combo1.Locked = True
   
    If Len(txtA2301) = 10 Then
        Call ReadData(txtA2301)
    End If
    Me.cmdEmail.Enabled = False
End Sub
'end 2014/04/16

Public Function DeleteCheck() As Boolean
   strExc(0) = "select A4416 from acc440 where a4421='" & txtA2301 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp(0)) Then
         MsgBox "本簽收資料已收款不可刪除！", vbCritical
      Else
         MsgBox "本簽收資料智權人員已繳款不可刪除！", vbCritical
      End If
   Else
      DeleteCheck = True
   End If
   
End Function

'Add by Amy 2023/05/23
Public Function GetAccSeq(ByRef stAllAccNo As String) As String
    Dim ii As Integer, stFieldSeq(1 To 3) As String, stFieldLast As String, arrTmp
    
    GetAccSeq = "": stAllAccNo = ""
    '1-5
    stFieldSeq(1) = stFieldSeq(1) & "110207,110303,110208,110204,110205"
    '6-10
    stFieldSeq(2) = stFieldSeq(2) & "110301,110302,110602,110502"
    '11-15
    stFieldSeq(3) = stFieldSeq(3) & ""
    '放最後
    stFieldLast = stFieldLast & "'1911','1912','1913'"
    arrTmp = Split(stFieldSeq(1) & IIf(stFieldSeq(2) = "", "", "," & stFieldSeq(2)) & IIf(stFieldSeq(3) = "", "", "," & stFieldSeq(3)), ",")
   
    For ii = LBound(arrTmp) To UBound(arrTmp)
        GetAccSeq = GetAccSeq & ",'" & arrTmp(ii) & "'," & ii + 1
        stAllAccNo = stAllAccNo & "','" & arrTmp(ii)
    Next ii
    If GetAccSeq <> MsgText(601) Then
        GetAccSeq = ",Decode(a0101,'" & Mid(GetAccSeq, 3) & ",99) Srt"
    End If
    If stAllAccNo <> MsgText(601) Then
        stAllAccNo = Mid(stAllAccNo, 3) & "'" & IIf(stFieldLast = "", "", "," & stFieldLast)
    End If
  
End Function


