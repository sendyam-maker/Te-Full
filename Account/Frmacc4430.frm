VERSION 5.00
Begin VB.Form Frmacc4430 
   AutoRedraw      =   -1  'True
   Caption         =   "科目明細表(對沖)"
   ClientHeight    =   5412
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6108
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5412
   ScaleWidth      =   6108
   Begin VB.ComboBox CboComp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1185
      TabIndex        =   0
      Text            =   "CboComp"
      Top             =   0
      Width           =   4720
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "Frmacc4430.frx":0000
      Left            =   3360
      List            =   "Frmacc4430.frx":0002
      TabIndex        =   18
      Top             =   2940
      Width           =   2640
   End
   Begin VB.TextBox Text7 
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
      Left            =   1860
      TabIndex        =   17
      Top             =   2940
      Width           =   390
   End
   Begin VB.TextBox txtAX211 
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
      Left            =   4320
      TabIndex        =   13
      Top             =   2250
      Width           =   1572
   End
   Begin VB.TextBox txtAX211 
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
      Left            =   2490
      TabIndex        =   12
      Top             =   2250
      Width           =   1572
   End
   Begin VB.TextBox txtAX208 
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
      Left            =   4320
      MaxLength       =   9
      TabIndex        =   5
      Top             =   990
      Width           =   1572
   End
   Begin VB.TextBox txtAX209 
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
      Left            =   4320
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1305
      Width           =   1572
   End
   Begin VB.TextBox txtAX214 
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
      Left            =   4320
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1620
      Width           =   1572
   End
   Begin VB.TextBox txtAX213 
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
      Left            =   4320
      TabIndex        =   11
      Top             =   1935
      Width           =   1572
   End
   Begin VB.TextBox txtAX213 
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
      Left            =   2490
      TabIndex        =   10
      Top             =   1935
      Width           =   1572
   End
   Begin VB.ComboBox cboSort 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   4
      Left            =   1848
      TabIndex        =   22
      Top             =   4476
      Width           =   2844
   End
   Begin VB.ComboBox cboSort 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   3
      Left            =   1836
      TabIndex        =   21
      Top             =   4092
      Width           =   2844
   End
   Begin VB.TextBox txtAX214 
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
      Left            =   2490
      MaxLength       =   12
      TabIndex        =   8
      Top             =   1620
      Width           =   1572
   End
   Begin VB.TextBox txtAX209 
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
      Left            =   2490
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1305
      Width           =   1572
   End
   Begin VB.TextBox Text4 
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
      Left            =   1785
      TabIndex        =   36
      Top             =   345
      Width           =   4110
   End
   Begin VB.TextBox Text3 
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
      TabIndex        =   1
      Top             =   345
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   156
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   4932
      Width           =   5790
   End
   Begin VB.ComboBox cboSort 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   1
      ItemData        =   "Frmacc4430.frx":0004
      Left            =   1836
      List            =   "Frmacc4430.frx":0006
      TabIndex        =   19
      Top             =   3336
      Width           =   2832
   End
   Begin VB.ComboBox cboSort 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   2
      Left            =   1836
      TabIndex        =   20
      Top             =   3720
      Width           =   2844
   End
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   4320
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Frmacc4430.frx":0008
      Left            =   4455
      List            =   "Frmacc4430.frx":000A
      TabIndex        =   16
      Top             =   2587
      Width           =   1440
   End
   Begin VB.TextBox Text6 
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
      Left            =   2355
      MaxLength       =   5
      TabIndex        =   15
      Top             =   2595
      Width           =   852
   End
   Begin VB.TextBox Text6 
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
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   14
      Top             =   2595
      Width           =   852
   End
   Begin VB.TextBox txtAX208 
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
      Left            =   2490
      MaxLength       =   9
      TabIndex        =   4
      Top             =   990
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3105
      TabIndex        =   3
      Top             =   660
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1185
      TabIndex        =   2
      Top             =   660
      Width           =   1572
   End
   Begin VB.Label Label25 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   49
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label24 
      BackStyle       =   0  '透明
      Caption         =   "對沖條件(1~5)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   48
      Top             =   2970
      Width           =   1515
   End
   Begin VB.Label Label23 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4155
      TabIndex        =   47
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "5. 沖帳傳票號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   46
      Top             =   2280
      Width           =   2265
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "4. 對沖代號(其他)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   45
      Top             =   1965
      Width           =   2265
   End
   Begin VB.Label Label21 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4155
      TabIndex        =   44
      Top             =   1965
      Width           =   255
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1548
      TabIndex        =   43
      Top             =   4488
      Width           =   252
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1536
      TabIndex        =   42
      Top             =   4116
      Width           =   252
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4155
      TabIndex        =   41
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4155
      TabIndex        =   40
      Top             =   1335
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4140
      TabIndex        =   39
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "3. 對沖代號(本所案號)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   38
      Top             =   1650
      Width           =   2265
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "2. 對沖代號(業)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   37
      Top             =   1335
      Width           =   2265
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -36
      Top             =   5004
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   372
      TabIndex        =   35
      Top             =   3360
      Width           =   972
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1536
      TabIndex        =   34
      Top             =   3372
      Width           =   252
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1536
      TabIndex        =   33
      Top             =   3732
      Width           =   252
   End
   Begin VB.Image Image3 
      Height          =   252
      Left            =   3720
      Picture         =   "Frmacc4430.frx":000C
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1590
      Left            =   135
      Top             =   3270
      Width           =   5835
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "列表選擇"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3510
      TabIndex        =   32
      Top             =   2625
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   31
      Top             =   2625
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "年度月份"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   30
      Top             =   2625
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "1. 對沖代號(客)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   29
      Top             =   1020
      Width           =   2265
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2865
      TabIndex        =   28
      Top             =   660
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   27
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   26
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   25
      Top             =   45
      Width           =   735
   End
End
Attribute VB_Name = "Frmacc4430"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
'Modified by Lydia cntX=500 = > 350
Const cntX As Long = 350   '橫印起始位置
Const cntY As Long = 500   '直印起始位置
'Modified by Lydia 2017/07/18 改成A4直印
'Const cntL As Long = 300   '列距
Const cntL As Long = 280   '列距
'Modified by Lydia 2017/07/18 改成A4直印
'Const cPageRec As Integer = 35 '每頁筆數
Const cPageRec As Integer = 45 '每頁筆數
Dim iPrint As Integer '印表位置
Dim PLeft(0 To 15) As Integer '各欄位位置
Dim m_stPageCol As String '是否跳頁判斷欄位值
Dim m_stGpCol As String, m_stGpDtlCol As String
Dim m_stGpAX214 As String   'add by sonia 2022/5/5
'Added by Lydia 2017/07/18
Dim strPrinter As String
Const ciGap As Integer = 80 '欄位間距
Const ciFontSize As Integer = 9 '明細字型大小

Sub GetPleft()
   Erase PLeft
   PLeft(0) = cntX   '對沖代號
   'Modified by Lydia 2017/07/18 改成A4直印+欄位間距
'   pLeft(1) = pLeft(0) + 3000 '傳票日期
'   pLeft(2) = pLeft(1) + 1200 '傳票編號
'   pLeft(3) = pLeft(2) + 1600 '摘要
'   pLeft(4) = pLeft(3) + 3500 '借方金額
'   pLeft(5) = pLeft(4) + 2000 '貸方金額
'   pLeft(6) = pLeft(5) + 2000 '累計餘額
'   pLeft(7) = pLeft(6) + 2000 '右邊
   PLeft(1) = PLeft(0) + Printer.TextWidth(String(11, "　")) + ciGap '傳票日期
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(5, "　")) + ciGap '傳票編號
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(5, "　")) + ciGap  '摘要
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(13, "　")) + ciGap '借方金額
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(8, "　")) + ciGap '貸方金額
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(8, "　")) + ciGap '累計餘額
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(8, "　")) + ciGap '右邊
   'end 2017/07/18
End Sub

'Add by Amy 2020/04/14
Private Sub CboComp_GotFocus()
    TextInverse cboComp
End Sub

Private Sub CboComp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboComp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(cboComp) = MsgText(601) Then Exit Sub
    
    strCmp = cboComp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label2 & MsgText(63), , MsgText(5)
        Cancel = True
        cboComp.SetFocus
        Exit Sub
    ElseIf Len(Trim(cboComp)) = 1 Then
        cboComp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub

Private Sub cboSort_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         If Index < 4 Then
            cboSort(Index + 1).SetFocus
         Else
            'Modify by Amy 2020/04/14 公司別改下拉 原:Text1
            cboComp.SetFocus
         End If
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Function Process() As Boolean
   Dim stCon As String, stCon1 As String
   Dim stSqlTblX As String, stSort As String
   Dim ii As Integer
   Dim strCmp As String 'Add by Amy 2020/04/14
   
   stCon = ""
   '部門別
   Text3.Text = Trim(Text3.Text)
   If Text3.Text <> "" Then
      stCon = stCon & " and ax204='" & Text3.Text & "'"
   End If
   '會計科目
   If Text5(0).Text <> "" Then
      stCon = stCon & " and ax205>='" & Text5(0).Text & "'"
   End If
   If Text5(1).Text <> "" Then
      stCon = stCon & " and ax205<='" & Text5(1).Text & "'"
   End If
   '1. 對沖代號(客)
   If txtAX208(0).Text <> "" Then
      stCon = stCon & " and ax208>='" & txtAX208(0).Text & "'"
   End If
   If txtAX208(1).Text <> "" Then
      stCon = stCon & " and ax208<='" & txtAX208(1).Text & "'"
   End If
    '2. 對沖代號(業)
   If txtAX209(0).Text <> "" Then
      stCon = stCon & " and ax209>='" & txtAX209(0).Text & "'"
   End If
   If txtAX209(1).Text <> "" Then
      stCon = stCon & " and ax209<='" & txtAX209(1).Text & "'"
   End If
   '3. 對沖代號(本所案號)
   If txtAX214(0).Text <> "" Then
      stCon = stCon & " and ax214>='" & txtAX214(0).Text & "'"
   End If
   If txtAX214(1).Text <> "" Then
      stCon = stCon & " and ax214<='" & txtAX214(1).Text & "'"
   End If
   '4. 對沖代號(其他)
   If txtAX213(0).Text <> "" Then
      stCon = stCon & " and ax213>='" & txtAX213(0).Text & "'"
   End If
   If txtAX213(1).Text <> "" Then
      stCon = stCon & " and ax213<='" & txtAX213(1).Text & "'"
   End If
   '5. 沖帳傳票號碼
   If txtAX211(0).Text <> "" Then
      stCon = stCon & " and ax211>='" & txtAX211(0).Text & "'"
   End If
   If txtAX211(1).Text <> "" Then
      stCon = stCon & " and ax211<='" & txtAX211(1).Text & "'"
   End If
   '年度月份
   If Text6(0).Text <> "" Then
      stCon = stCon & " and a0205>=" & Format(Text6(0).Text & "01", "0")
   End If
   If Text6(1).Text <> "" Then
      stCon = stCon & " and a0205<=" & Format(Text6(1).Text & "31", "0")
   End If
   
   '對沖條件
   m_stGpCol = "": m_stGpDtlCol = ""
   Select Case Text7.Text
      Case "1"
         m_stGpCol = "ax208"
         m_stGpDtlCol = "cust"
      Case "2"
         m_stGpCol = "ax209"
         m_stGpDtlCol = "salesname"
      Case "3"
         m_stGpCol = "ax214"
         m_stGpDtlCol = "case"
         m_stGpAX214 = "substr(ax214,1,length(ax214)-2)" 'add by sonia 2022/5/5 子案與母案合併計算
      Case "4"
         m_stGpCol = "ax213"
         m_stGpDtlCol = "ax213"
      Case "5"
         m_stGpCol = "ax211"
         m_stGpDtlCol = "ax211"
   End Select
      
   '排序
   stSort = " order by AX201, AX205"
   'If m_stGpCol <> "" Then stSort = stSort & "," & m_stGpCol   '2010/5/18 cancel by sonia CFP指定國家不排序
   'Modify by Amy 2020/04/14 公司別改下拉 原:Text1
   strCmp = cboComp
   If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
   End If
   
   '列表選擇
   '未沖平或沖平
   If Combo1.ListIndex = 0 Or Combo1.ListIndex = 1 Then
      
      '20140123START Modify By eric
      'Modify by Amy 2020/04/14 公司別改下拉 原:Text1
      If Trim(cboComp) <> MsgText(601) Then
         'add by sonia 2022/5/5 子案與母案合併計算
         If m_stGpCol = "ax214" Then
            stSqlTblX = ",( select AX201 K1, AX205 K2, " & m_stGpAX214 & " K3,nvl(SUM(AX206),0) S1, nvl(SUM(AX207),0) S2" & _
               " From acc021, acc020" & _
               " where ''||ax201='" & strCmp & "'" & stCon & _
               " and a0201(+) = ax201 and a0202(+) = ax202" & _
               " GROUP BY AX201, AX205, " & m_stGpAX214 & ") X "
         Else
         'end 2022/5/5
            stSqlTblX = ",( select AX201 K1, AX205 K2, " & m_stGpCol & " K3,nvl(SUM(AX206),0) S1, nvl(SUM(AX207),0) S2" & _
               " From acc021, acc020" & _
               " where ''||ax201='" & strCmp & "'" & stCon & _
               " and a0201(+) = ax201 and a0202(+) = ax202" & _
               " GROUP BY AX201, AX205, " & m_stGpCol & ") X "
         End If
      End If
      If Trim(cboComp) = MsgText(601) Then
         'add by sonia 2022/5/5 子案與母案合併計算
         If m_stGpCol = "ax214" Then
            stSqlTblX = ",( select AX201 K1, AX205 K2, " & m_stGpAX214 & " K3,nvl(SUM(AX206),0) S1, nvl(SUM(AX207),0) S2" & _
               " From acc021, acc020" & _
               " where a0201(+) = ax201 and a0202(+) = ax202 " & stCon & _
               " GROUP BY AX201, AX205, " & m_stGpAX214 & ") X "
         Else
         'end 2022/5/5
            stSqlTblX = ",( select AX201 K1, AX205 K2, " & m_stGpCol & " K3,nvl(SUM(AX206),0) S1, nvl(SUM(AX207),0) S2" & _
               " From acc021, acc020" & _
               " where a0201(+) = ax201 and a0202(+) = ax202 " & stCon & _
               " GROUP BY AX201, AX205, " & m_stGpCol & ") X "
         End If
      End If
      'end 2020/04/14
      'stSqlTblX = ",( select AX201 K1, AX205 K2, " & m_stGpCol & " K3,nvl(SUM(AX206),0) S1, nvl(SUM(AX207),0) S2" & _
      '   " From acc021, acc020" & _
      '   " where ''||ax201='" & Text1.Text & "'" & stCon & _
      '   " and a0201(+) = ax201 and a0202(+) = ax202" & _
      '   " GROUP BY AX201, AX205, " & m_stGpCol & ") X "
      '20140123END
         
      'add by sonia 2022/5/5 子案與母案合併計算
      If m_stGpCol = "ax214" Then
         stCon1 = " and K1(+) = ax201 AND K2(+)= ax205 and K3(+) = " & m_stGpAX214 & " and K1 is not null"
      Else
      'end 2022/5/5
         stCon1 = " and K1(+) = ax201 AND K2(+)= ax205 and K3(+) = " & m_stGpCol & " and K1 is not null"
      End If
      If Combo1.ListIndex = 1 Then
         stCon1 = stCon1 & " AND S1=S2"
      Else
         stCon1 = stCon1 & " AND S1<>S2"
      End If
   End If
   
  For ii = 1 To 4
      If cboSort(ii).Text <> "" Then
         Select Case cboSort(ii).ListIndex
            Case 0
               stSort = stSort & ",ax214"
            Case 1
               stSort = stSort & ",ax208"
            Case 2
               stSort = stSort & ",ax209"
            Case 3
               stSort = stSort & ",ax213"
            Case 4
               stSort = stSort & ",ax211"
         End Select
      End If
   Next
   
   'Add by Morgan 2005/9/13 最後加傳票日期編號排序
   stSort = stSort & ",A0205,A0202"
   
   '20140123START Modify By eric
   'Modify by Amy 2020/04/14 公司別改下拉 原:Text1
   If Trim(cboComp) <> MsgText(601) Then
      strSql = "select ax201,ax205,ax208, ax209" & _
      ", substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as case" & _
      ", a0205, ax202, ax206, ax207, substrb(ax212,1,28) as memo, substrb(ax208||nvl(nvl(cu04,a0i02),st2.st02),1,22) as cust" & _
      ", ax209||st1.st02 as salesname, ax213, ax214" & _
      " From acc021, acc020" & stSqlTblX & ", staff st1, customer, acc0i0, staff st2" & _
      " where ''||ax201='" & strCmp & "'" & stCon & stCon1 & _
      " and a0201(+) = ax201 and a0202(+) = ax202" & _
      " and st1.st01(+) = ax209" & _
      " and cu01(+) = substr(ax208, 1, length(ax208) - 1) and cu02(+) = substr(ax208, length(ax208), 1)" & _
      " and a0i01(+) = ax208" & _
      " and st2.st01(+) = ax208" & stSort
   End If
   If Trim(cboComp) = MsgText(601) Then
      strSql = "select ax201,ax205,ax208, ax209" & _
      ", substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as case" & _
      ", a0205, ax202, ax206, ax207, substrb(ax212,1,28) as memo, substrb(ax208||nvl(nvl(cu04,a0i02),st2.st02),1,22) as cust" & _
      ", ax209||st1.st02 as salesname, ax213, ax214" & _
      " From acc021, acc020" & stSqlTblX & ", staff st1, customer, acc0i0, staff st2" & _
      " where  a0201(+) = ax201 and a0202(+) = ax202" & stCon & stCon1 & _
      " and st1.st01(+) = ax209" & _
      " and cu01(+) = substr(ax208, 1, length(ax208) - 1) and cu02(+) = substr(ax208, length(ax208), 1)" & _
      " and a0i01(+) = ax208" & _
      " and st2.st01(+) = ax208" & stSort
   End If
   'end 2020/04/14
   'strSql = "select ax201,ax205,ax208, ax209" & _
   '   ", substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as case" & _
   '   ", a0205, ax202, ax206, ax207, substrb(ax212,1,28) as memo, substrb(ax208||nvl(nvl(cu04,a0i02),st2.st02),1,22) as cust" & _
   '   ", ax209||st1.st02 as salesname, ax213, ax214" & _
   '   " From acc021, acc020" & stSqlTblX & ", staff st1, customer, acc0i0, staff st2" & _
   '   " where ''||ax201='" & Text1.Text & "'" & stCon & stCon1 & _
   '   " and a0201(+) = ax201 and a0202(+) = ax202" & _
   '   " and st1.st01(+) = ax209" & _
   '   " and cu01(+) = substr(ax208, 1, length(ax208) - 1) and cu02(+) = substr(ax208, length(ax208), 1)" & _
   '   " and a0i01(+) = ax208" & _
   '   " and st2.st01(+) = ax208" & stSort
   '20140123END
      
      
      
On Error GoTo ErrHnd

   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      '若有資料
      If .RecordCount > 0 Then
         Process = DoPrint(AdoRecordSet3)
      Else
         MsgBox "無資料可列印！"
      End If
   End With
   CheckOC3
   
ErrHnd:
      If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
      
End Function

Private Function DoPrint(ByRef p_Recordset As ADODB.Recordset) As Boolean

   Dim strTitle As String
   Dim iPage As Integer, iPages As Integer
   Dim iRow As Long
   Dim stPageKey As String
   Dim dblBTot As Double, dblLTot As Double, dblTot As Double
   Dim dblSubTot As Double
   
On Error GoTo ErrHnd
   
   PUB_RestorePrinter Combo2 'Added by Lydia 2017/07/18
   
   'Add by Morgan 2005/8/29 為了控制未沖平報表累計餘額0以前的資料不印改回先寫暫存檔
   Dim adoaccrpt404 As New ADODB.Recordset
   adoTaie.Execute "delete from accrpt404 where R40401='" & strUserNum & "'"
   adoaccrpt404.CursorLocation = adUseClient
   adoaccrpt404.Open "select * from accrpt404", adoTaie, adOpenDynamic, adLockBatchOptimistic
   With p_Recordset
      .MoveFirst
      If m_stGpCol <> "" Then
         'modify by sonia 2022/5/5 EPC或TF子案與母案合併計算CFP-028265-0-11,所以案號最後2或4碼不取
         'stPageKey = "" & .Fields("ax201") & .Fields("ax205") & .Fields(m_stGpCol)
         If m_stGpCol = "ax214" Then
            If txtAX214(0) <> "" And Mid(txtAX214(0), 1, 2) = "TF" Then
               stPageKey = "" & .Fields("ax201") & .Fields("ax205") & Mid(.Fields(m_stGpCol), 1, Len(.Fields(m_stGpCol)) - 4)
            Else
               stPageKey = "" & .Fields("ax201") & .Fields("ax205") & Mid(.Fields(m_stGpCol), 1, Len(.Fields(m_stGpCol)) - 2)
            End If
         Else
            stPageKey = "" & .Fields("ax201") & .Fields("ax205") & .Fields(m_stGpCol)
         End If
         'end 2022/5/5
      End If
      Do While Not .EOF
         iRow = iRow + 1
         If m_stGpCol <> "" Then
            'modify by sonia 2022/5/5 EPC或TF子案與母案合併計算CFP-028265-0-11,所以案號最後2或4碼不取
            'If stPageKey <> "" & .Fields("ax201") & .Fields("ax205") & .Fields(m_stGpCol) Then
            '   adoaccrpt404.UpdateBatch
            '   stPageKey = "" & .Fields("ax201") & .Fields("ax205") & .Fields(m_stGpCol)
            '   dblSubTot = 0
            'End If
            If m_stGpCol = "ax214" Then
               If Mid(stPageKey, 8, 2) = "TF" Then
                  If stPageKey <> "" & .Fields("ax201") & .Fields("ax205") & Mid(.Fields(m_stGpCol), 1, Len(.Fields(m_stGpCol)) - 4) Then
                     adoaccrpt404.UpdateBatch
                     If Mid(stPageKey, 8, 2) = "TF" Then
                        stPageKey = "" & .Fields("ax201") & .Fields("ax205") & Mid(.Fields(m_stGpCol), 1, Len(.Fields(m_stGpCol)) - 4)
                     Else
                        stPageKey = "" & .Fields("ax201") & .Fields("ax205") & Mid(.Fields(m_stGpCol), 1, Len(.Fields(m_stGpCol)) - 2)
                     End If
                     dblSubTot = 0
                  End If
               Else
                  If stPageKey <> "" & .Fields("ax201") & .Fields("ax205") & Mid(.Fields(m_stGpCol), 1, Len(.Fields(m_stGpCol)) - 2) Then
                     adoaccrpt404.UpdateBatch
                     If Left(m_stGpCol, 2) = "TF" Then
                        stPageKey = "" & .Fields("ax201") & .Fields("ax205") & Mid(.Fields(m_stGpCol), 1, Len(.Fields(m_stGpCol)) - 4)
                     Else
                        stPageKey = "" & .Fields("ax201") & .Fields("ax205") & Mid(.Fields(m_stGpCol), 1, Len(.Fields(m_stGpCol)) - 2)
                     End If
                     dblSubTot = 0
                  End If
               End If
            Else
               If stPageKey <> "" & .Fields("ax201") & .Fields("ax205") & .Fields(m_stGpCol) Then
                  adoaccrpt404.UpdateBatch
                  stPageKey = "" & .Fields("ax201") & .Fields("ax205") & .Fields(m_stGpCol)
                  dblSubTot = 0
               End If
            End If
            'end 2022/5/5
         End If
         adoaccrpt404.AddNew
         adoaccrpt404.Fields("R40401") = strUserNum
         If m_stGpCol <> "" Then
            '對沖代號
            adoaccrpt404.Fields("R40402") = "" & .Fields(m_stGpDtlCol)
         End If
         adoaccrpt404.Fields("R40403") = stPageKey
         If m_stGpCol = "ax214" Then
            'modify by sonia 2022/5/5 EPC或TF子案與母案合併計算CFP-028265-0-11,所以案號最後2或4碼不取
            'adoaccrpt404.Fields("R40403") = "" & Left(.Fields(m_stGpCol), 10)  'add by sonia 2021/4/22
            If Mid(stPageKey, 8, 2) = "TF" Then
               adoaccrpt404.Fields("R40403") = "" & Mid(.Fields(m_stGpCol), 1, Len(.Fields(m_stGpCol)) - 4)
            Else
               adoaccrpt404.Fields("R40403") = "" & Mid(.Fields(m_stGpCol), 1, Len(.Fields(m_stGpCol)) - 2)
            End If
         End If
         '會計科目
         adoaccrpt404.Fields("R40404") = "" & .Fields("ax205")
         '傳票日期
         adoaccrpt404.Fields("R40405") = Val("" & .Fields("a0205"))
         '傳票編號
         adoaccrpt404.Fields("R40406") = "" & .Fields("ax202")
         '摘　　要
         adoaccrpt404.Fields("R40407") = "" & .Fields("memo")
         '借方金額
         adoaccrpt404.Fields("R40408") = Val("" & .Fields("ax206"))
         '貸方金額
         adoaccrpt404.Fields("R40409") = Val("" & .Fields("ax207"))
         '累計餘額 1,6,8開頭的=借-貸;其他=貸-借
         If InStr("1,6,8", Left("" & .Fields("ax205"), 1)) > 0 Then
            dblSubTot = dblSubTot + Val("" & .Fields("ax206")) - Val("" & .Fields("ax207"))
         Else
            dblSubTot = dblSubTot + Val("" & .Fields("ax207")) - Val("" & .Fields("ax206"))
         End If
         adoaccrpt404.Fields("R40410") = dblSubTot
         adoaccrpt404.Fields("R40413") = iRow
         
         '若選沖平時累計餘額0以前的資料不存檔
         If Combo1.ListIndex = 0 Then
            If dblSubTot = 0 Then
               adoaccrpt404.CancelBatch
            End If
         End If
         .MoveNext
      Loop
      adoaccrpt404.UpdateBatch
   End With
   adoaccrpt404.Close
  
   'Modified by Lydia 2017/07/03 改成A4橫印
   'GetPleft
   'Printer.Orientation = 1
   Printer.PaperSize = 9
   'Modified by Lydia 2017/07/18 改成直印
   'Printer.Orientation = 2
   Printer.Orientation = 1
   Printer.Font = "細明體"
   'Modified by Lydia 2017/07/18 字型縮小
   'Printer.FontSize = 12
   Printer.FontSize = ciFontSize
   GetPleft
   'end 2017/07/03
   'Printer.PaperSize = 39  'US SF
   'Printer.Font = "細明體" 'Remove by Lydia 2017/07/03
   iPage = 1: iRow = 0
   
   strTitle = "*** 科目明細表(對沖) ***"
   
   With adoaccrpt404
      strSql = "select * from accrpt404 where R40401='" & strUserNum & "' order by R40413"
      .CursorLocation = adUseClient
      .Open strSql, adoTaie, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         
         iPages = .RecordCount \ cPageRec + IIf((.RecordCount Mod cPageRec) = 0, 0, 1)
         PrintHead strTitle, iPage, iPages
         .MoveFirst
         'If m_stGpCol <> "" Then stPageKey = "" & .Fields("ax201") & .Fields("ax205") & .Fields(m_stGpCol)
         '2010/5/18 add by sonia
         'modify by sonia 2022/5/5 原語法不適用於TF
         'If m_stGpCol = "ax214" And Left(txtAX214(0), 10) = Left(txtAX214(1), 10) Then
         '   stPageKey = "" & Left(.Fields("R40403"), 10)
         'cancel by sonia 2022/5/5 上面寫入R40403時已區分，此處不必再做
         'If m_stGpCol = "ax214" And Mid(txtAX214(0), 1, Len(txtAX214(0)) - 2) = Mid(txtAX214(1), 1, Len(txtAX214(1)) - 2) Then
         '   stPageKey = "" & Mid(.Fields("R40403"), 1, Len(.Fields("R40403")) - 2)
         '2010/5/18 end
         'ElseIf m_stGpCol <> "" Then
         If m_stGpCol <> "" Then
         'end 2022/5/5
            stPageKey = "" & .Fields("R40403")
         End If
         
         Do While Not .EOF
            iRow = iRow + 1
            
            '2010/5/18 add by sonia
            'cancel by sonia 2022/5/5 上面寫入R40403時已區分，此處不必再做
            'If m_stGpCol = "ax214" And Left(txtAX214(0), 10) = Left(txtAX214(1), 10) Then
            '   If stPageKey <> "" & Left(.Fields("R40403"), 10) Then
            '      iPrint = iPrint + cntL
            '      Printer.Line (PLeft(4), iPrint)-(PLeft(7), iPrint)
            '      stPageKey = "" & Left(.Fields("R40403"), 10)
            '      iRow = iRow + 1
            '      dblTot = 0
            '   End If
            ''2010/5/18 end
            'ElseIf m_stGpCol <> "" Then
            'end 2022/5/5
            If m_stGpCol <> "" Then
               If stPageKey <> "" & .Fields("R40403") Then
                  iPrint = iPrint + cntL
                  Printer.Line (PLeft(4), iPrint)-(PLeft(7), iPrint)
                  stPageKey = "" & .Fields("R40403")
                  iRow = iRow + 1
                  dblTot = 0
               End If
            End If
            
            If iRow > cPageRec Then
               iPrint = iPrint + cntL
               Printer.Line (PLeft(4), iPrint)-(PLeft(7), iPrint)
               Printer.NewPage
               iPage = iPage + 1
               PrintHead strTitle, iPage, iPages
               iRow = 0
            End If
            iPrint = iPrint + cntL
            If m_stGpCol <> "" Then
               '對沖代號
               Printer.CurrentX = PLeft(0)
               Printer.CurrentY = iPrint
               'Modified by Lydia 2017/07/18 限定字串長度
               'Printer.Print "" & .Fields("R40402")
               Printer.Print PUB_StrToStr("" & .Fields("R40402"), 22)
            End If
            '傳票日期
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print Format("" & .Fields("R40405"), "###/##/##")
            '傳票編號
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = iPrint
            Printer.Print "" & .Fields("R40406")
            '摘　　要
            Printer.CurrentX = PLeft(3)
            Printer.CurrentY = iPrint
            'Modified by Lydia 2017/07/18 限定字串長度
            'Printer.Print "" & .Fields("R40407")
            Printer.Print PUB_StrToStr("" & .Fields("R40407"), 26)
            '借方金額
            dblBTot = dblBTot + Val("" & .Fields("R40408"))
            Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format("" & .Fields("R40408"), FDollar))
            Printer.CurrentY = iPrint
            Printer.Print Format("" & .Fields("R40408"), FDollar)
            '貸方金額
            dblLTot = dblLTot + Val("" & .Fields("R40409"))
            Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format("" & .Fields("R40409"), FDollar))
            Printer.CurrentY = iPrint
            Printer.Print Format("" & .Fields("R40409"), FDollar)
            '累計餘額 1,6,8開頭的=借-貸;其他=貸-借
            If InStr("1,6,8", Left("" & .Fields("R40404"), 1)) > 0 Then
               dblTot = dblTot + Val("" & .Fields("R40408")) - Val("" & .Fields("R40409"))
            Else
               dblTot = dblTot + Val("" & .Fields("R40409")) - Val("" & .Fields("R40408"))
            End If
            Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(dblTot, FDollar))
            Printer.CurrentY = iPrint
            Printer.Print Format(dblTot, FDollar)
            .MoveNext
         Loop
         '表尾
         iPrint = iPrint + cntL
         Printer.Line (PLeft(4), iPrint)-(PLeft(7), iPrint)
         iPrint = iPrint + cntL
         Printer.CurrentX = PLeft(4) - Printer.TextWidth("合計：")
         Printer.CurrentY = iPrint
         Printer.Print "合計："
         Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblBTot, FDollar))
         Printer.CurrentY = iPrint
         Printer.Print Format(dblBTot, FDollar)
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblLTot, FDollar))
         Printer.CurrentY = iPrint
         Printer.Print Format(dblLTot, FDollar)
         iPrint = iPrint + cntL
         Printer.Line (PLeft(4), iPrint)-(PLeft(7), iPrint)
         Printer.Line (PLeft(4), iPrint + 50)-(PLeft(7), iPrint + 50)
         
         Printer.EndDoc
         
      End If
   End With
   
   DoPrint = True
   
ErrHnd:
   If Err.Number <> 0 Then
      Printer.EndDoc
      MsgBox Err.Description, vbCritical
   End If
   PUB_RestorePrinter strPrinter 'Added by Lydia 2017/07/18
   Set adoaccrpt404 = Nothing
End Function

Private Sub PrintHead(stTitle As String, iPage As Integer, iPageTot As Integer)
   Dim stDesc As String, iXPlus As Integer, iYPlus As Integer
   Dim startX1 As Integer, startX2 As Integer  'Added by Lydia 2017/07/18
   Dim strCmpNo As String 'Add by Amy 2020/04/14
   
   iPrint = cntY
   Printer.Font.Name = "細明體"
   '報表名稱
   Printer.Font.Size = 14
   Printer.Font.Bold = True
   
   'Modified by Lydia 2017/07/18
   'Printer.CurrentX = 7000 - (Printer.TextWidth(stTitle) / 2)
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(stTitle) / 2)
   Printer.CurrentY = iPrint
   Printer.Print stTitle
   
   'Modified by Lydia 2017/07/18
   'Printer.Font.Size = 12
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   
   'Modify by Morgan 2005/8/29 加控制條件有輸才要印
   '條件
   iPrint = Printer.CurrentY
   'Modified by Lydia 2017/07/18
   'iXPlus = 3000: iYPlus = cntL
   startX1 = 2500: startX2 = 7000 'startX1 + Printer.TextWidth(String(30, "　"))
   iXPlus = startX1: iYPlus = cntL
   'end  2017/07/18
   
   '20140123START Modify By eric
   'Modify by Amy 2020/04/14 公司別改下拉
'   If Text1.Text <> MsgText(601) Then
'      stDesc = Text2.Text
'   Else
'      stDesc = "台一　專利商標/智權"
'   End If
   'stDesc = Text1.Text & Text2.Text
   '20140123END
   strCmpNo = cboComp
   If InStr(strCmpNo, "　") > 0 Then
        strCmpNo = Mid(strCmpNo, 1, Val(InStr(strCmpNo, "　")) - 1)
   End If
   stDesc = GetAccReportCmpN(strCmpNo, True)
   
   If stDesc <> "" Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "公司別：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If
   
   stDesc = Text3.Text & Text4.Text
   If stDesc <> "" Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "部門別：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If
   
   stDesc = Text5(0).Text & " ~ " & Text5(1).Text
   If stDesc <> " ~ " Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "會計科目：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If
   
   stDesc = Text6(0).Text & " ~ " & Text6(1).Text
   If stDesc <> " ~ " Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "年度月份：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If
   
   stDesc = txtAX208(0).Text & " ~ " & txtAX208(1).Text
   If stDesc <> " ~ " Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "對沖代號(客)：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If
   
   stDesc = txtAX209(0).Text & " ~ " & txtAX209(1).Text
   If stDesc <> " ~ " Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "對沖代號(業)：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If
   
   stDesc = txtAX214(0).Text & " ~ " & txtAX214(1).Text
   If stDesc <> " ~ " Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "對沖代號(本)：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If
   
   stDesc = txtAX213(0).Text & " ~ " & txtAX213(1).Text
   If stDesc <> " ~ " Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "對沖代號(他)：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If
   
   stDesc = txtAX211(0).Text & " ~ " & txtAX211(1).Text
   If stDesc <> " ~ " Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "沖帳傳票號碼：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If

   stDesc = Combo1.Text
   If stDesc <> "" Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "列表選擇：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If
   
   stDesc = Text7.Text
   If stDesc <> "" Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "對沖條件：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If

   stDesc = cboSort(1).Text & "," & cboSort(2).Text & "," & cboSort(3).Text & "," & cboSort(4).Text
   If stDesc <> "" Then
      iPrint = iPrint + iYPlus
      Printer.CurrentX = cntX + iXPlus
      Printer.CurrentY = iPrint
      Printer.Print "排序方式：" & stDesc
      'Modified by Lydia 2017/07/18 改成A4直印
      'If iXPlus = 3000 Then
      '   iXPlus = 8000: iYPlus = 0
      'Else
      '   iXPlus = 3000: iYPlus = cntL
      'End If
      If iXPlus = startX1 Then
         iXPlus = startX2: iYPlus = 0
      Else
         iXPlus = startX1: iYPlus = cntL
      End If
      'end 2017/07/18
   End If
   '2005/8/29 end
   
   iPrint = iPrint + cntL
   Printer.CurrentX = cntX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & GetStaffName(strUserNum)
   'Modified by Lydia 2017/07/18
   'Printer.CurrentX = cntX + 13000
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(String(13, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
  
   iPrint = iPrint + cntL
   'Modified by Lydia 2017/07/18
   'Printer.CurrentX = cntX + 13000
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(String(13, "　"))
   Printer.CurrentY = iPrint
   'Modify by Morgan 2005/9/13 頁數會有錯
   'Printer.Print "頁　　次：" & Format(iPage) & " / " & Format(iPageTot)
   Printer.Print "頁　　次：" & Format(iPage)
   
   iPrint = iPrint + cntL
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "對沖代號"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "傳票日期"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "傳票編號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "摘　　要"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "借方金額"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "貸方金額"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "累計餘額"
   iPrint = iPrint + cntL
   Printer.Line (PLeft(0), iPrint)-(PLeft(7), iPrint)
   iPrint = iPrint - 200
End Sub
Private Sub Command1_Click()

   Screen.MousePointer = vbHourglass
   If FormCheck = False Then
      'MsgBox MsgText(181), , MsgText(5)
   Else
      If Process = True Then
         MsgBox "列印完成！"
         FormClear
      End If
   End If
   Screen.MousePointer = vbDefault
   
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102) 'Remove by Lydia 2017/07/31
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   'Remove by Lydia 2017/07/31
   'If KeyCode <> vbKeyEscape Then
   '   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   'End If
End Sub

Private Sub Form_Load()

   Dim intX As Integer
   Dim intY As Integer
   Dim sglWidth As Single
   Dim sglHeight As Single
   Dim ii As Integer
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 6350 'Modify by Amy 2023/07/19 原:6200
   Me.Height = 6000 'Modify by Amy 2023/07/19 原:5800
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2020/04/14
   cboComp.Clear
   cboComp.AddItem "", 0
   Call Pub_SetCboCmp(cboComp, False, False, False, , 1)
   'end 2020/04/14
   
   '列表選擇
   Combo1.AddItem ComboItem(141) '未沖平
   Combo1.AddItem ComboItem(142) '沖平
   Combo1.AddItem ComboItem(143) '全部
   Combo1.ListIndex = 2
   
   '排序
   For ii = 1 To 4
      cboSort(ii).AddItem ComboItem(231)  '1--本所案號
      cboSort(ii).AddItem ComboItem(232)  '2--客戶/廠商/員工
      cboSort(ii).AddItem ComboItem(233)  '3--智權人員
      cboSort(ii).AddItem ComboItem(234)  '4--其它
      cboSort(ii).AddItem "5--沖帳傳票號碼"
      cboSort(ii).AddItem "6--傳票日期"     '2010/5/18 add by sonia
   Next
   cboSort(1) = ComboItem(233)
   
   '20140123Remark By eric
   '公司別
   'Text1.Text = "1"
   
   '對沖條件(1~5)
   Text7.Text = "1"
   
   'Remove by Lydia 2017/07/20 因為已預設印表機,所以不用顯示更換紙張
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   
   PUB_SetPrinter Me.Name, Combo2, strPrinter 'Added by Lydia 2017/07/18 預設印表機選單
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   'Added by Lydia 2017/07/18 若印表機變動, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   'end 2017/07/18
   
   Set Frmacc4430 = Nothing
End Sub

'Mark by Amy 2020/04/14 公司別改下拉
'Private Sub Text1_GotFocus()
'   TextInverse Text1
'   'edit by nickc 2007/06/11  切換輸入法改用API
'   'Text1.IMEMode = 2
'   CloseIme
'End Sub
'
''20140123START Add By eric
'Private Sub Text1_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub

'Private Sub Text1_Change()
'   '20140123START Modify By eric
'   'Text2 = A0802Query(Text1)
'   Select Case Text1
'      Case "1"
'         Text2 = A0802Query(Text1)
'      Case "2"
'         Text2 = A0802Query("J")
'   End Select
'   '20140123END
'End Sub
'end 2020/04/14

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_GotFocus(Index As Integer)
   If Index = 1 Then
      If Text5(Index).Text = "" Then Text5(Index).Text = Text5(Index - 1).Text
   End If
   TextInverse Text5(Index)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text5(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub Text6_GotFocus(Index As Integer)
   If Index = 1 Then
      If Text6(Index).Text = "" Then Text6(Index).Text = Text6(Index - 1).Text
   End If
   TextInverse Text6(Index)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text6(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text7.IMEMode = 2
   CloseIme
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   If KeyAscii <> vbKeyBack And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> Asc("4") And KeyAscii <> Asc("5") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text3_Change()
   Text4 = A0902Query(Text3)
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
'Private Sub ProduceData()
'Dim lngStartDate As Long
'Dim lngEndDate As Long
'Dim strSubNo As String
'Dim strOrder1 As String
'Dim strOrder2 As String
'Dim strSQL As String
'Dim strSort As String
'Dim strGroup As String
'Dim strName As String
'Dim lngCounter As Long
'Dim douAmount As Double
'
'On Error GoTo Checking
'   Me.MousePointer = vbHourglass
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   adoaccrpt404.CursorLocation = adUseClient
'   adoaccrpt404.Open "select * from accrpt404", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   lngStartDate = Val(Text4 & MsgText(12))
'   lngEndDate = Val(Text5 & MsgText(13))
'   adoacc021.CursorLocation = adUseClient
'   If Text3 <> MsgText(601) Then
'      strSubNo = " and ax208 = '" & Text3 & "'"
'   Else
'      If Text10 <> MsgText(601) Then
'         strSubNo = " and ax209 = '" & Text10 & "'"
'      Else
'         If Text11 <> MsgText(601) Then
'            strSubNo = " and ax214 = '" & Text11 & "'"
'         Else
'            strSubNo = MsgText(601)
'         End If
'      End If
'   End If
'   If Text8 <> MsgText(601) Then
'      strSQL = strSQL & " and ax204 = '" & Text8 & "'"
'   End If
'   If Text2 <> MsgText(601) Then
'      strSQL = strSQL & " and ax205 >= '" & Text2 & "'"
'   End If
'   If Text1 <> MsgText(601) Then
'      strSQL = strSQL & " and ax205 <= '" & Text1 & "'"
'   End If
'   If Text4 <> MsgText(601) Then
'      strSQL = strSQL & " and a0205 >= " & lngStartDate & ""
'   End If
'   If Text5 <> MsgText(601) Then
'      strSQL = strSQL & " and a0205 <= " & lngEndDate & ""
'   End If
'   If Text3 <> MsgText(601) Then
'      strSQL = strSQL & " and ax208 >= '" & Text3 & "'"
'   End If
'   If Text12 <> MsgText(601) Then
'      strSQL = strSQL & " and ax208 <= '" & Text12 & "'"
'   End If
'   If Text10 <> MsgText(601) Then
'      strSQL = strSQL & " and ax209 >= '" & Text10 & "'"
'   End If
'   If Text13 <> MsgText(601) Then
'      strSQL = strSQL & " and ax209 <= '" & Text13 & "'"
'   End If
'   If Text11 <> MsgText(601) Then
'      strSQL = strSQL & " and ax214 >= '" & Text11 & "'"
'   End If
'   If Text14 <> MsgText(601) Then
'      strSQL = strSQL & " and ax214 <= '" & Text14 & "'"
'   End If
'   If Text17 <> MsgText(601) Then
'      strSQL = strSQL & " and ax213 >= '" & Text17 & "'"
'   End If
'   If Text16 <> MsgText(601) Then
'      strSQL = strSQL & " and ax213 <= '" & Text16 & "'"
'   End If
''   If Text6 <> MsgText(601) Then
''      strSQL = strSQL & " and ax201 = '" & Text6 & "'"
''   End If
'
'   'Add by Morgan 2004/10/19
'   If Text18.Text <> "" Then
'      strSQL = strSQL & " and ax211 >= '" & Text18.Text & "'"
'   End If
'   If Text19.Text <> "" Then
'      strSQL = strSQL & " and ax211 <= '" & Text19.Text & "'"
'   End If
'
'   Select Case Text20.Text
'      Case "1"
'         strSort = strSort & " ax208 asc,"
'      Case "2"
'         strSort = strSort & " ax209 asc,"
'      Case "3"
'         strSort = strSort & " ax214 asc,"
'      Case "4"
'         strSort = strSort & " ax213 asc,"
'      Case "5"
'         strSort = strSort & " ax211 asc,"
'   End Select
'   '2004/10/19 end
'
'   If Combo13 <> MsgText(601) Then
'      Select Case Mid(Combo13, 1, 1)
'         Case "1"
'            strSort = strSort & " ax214 asc,"
'         Case "2"
'            strSort = strSort & " ax208 asc,"
'         Case "3"
'            strSort = strSort & " ax209 asc,"
'         Case "4"
'            strSort = strSort & " ax213 asc,"
'         'Add by Morgan 2004/10/19
'         Case "5"
'            strSort = strSort & " ax211 asc,"
'      End Select
'   End If
'   If Combo5 <> MsgText(601) Then
'      Select Case Mid(Combo5, 1, 1)
'         Case "1"
'            strSort = strSort & " ax214 asc,"
'         Case "2"
'            strSort = strSort & " ax208 asc,"
'         Case "3"
'            strSort = strSort & " ax209 asc,"
'         Case "4"
'            strSort = strSort & " ax213 asc,"
'         'Add by Morgan 2004/10/19
'         Case "5"
'            strSort = strSort & " ax211 asc,"
'      End Select
'   End If
'   If Combo1 <> MsgText(601) Then
'      Select Case Mid(Combo1, 1, 1)
'         Case "1"
'            strSort = strSort & " ax214 asc,"
'         Case "2"
'            strSort = strSort & " ax208 asc,"
'         Case "3"
'            strSort = strSort & " ax209 asc,"
'         Case "4"
'            strSort = strSort & " ax213 asc,"
'         'Add by Morgan 2004/10/19
'         Case "5"
'            strSort = strSort & " ax211 asc,"
'      End Select
'   End If
'   If Combo2 <> MsgText(601) Then
'      Select Case Mid(Combo2, 1, 1)
'         Case "1"
'            strSort = strSort & " ax214 asc,"
'         Case "2"
'            strSort = strSort & " ax208 asc,"
'         Case "3"
'            strSort = strSort & " ax209 asc,"
'         Case "4"
'            strSort = strSort & " ax213 asc,"
'         'Add by Morgan 2004/10/19
'         Case "5"
'            strSort = strSort & " ax211 asc,"
'      End Select
'   End If
'   If strSort <> MsgText(601) Then
'      strSort = " order by " & Mid(strSort, 1, Len(strSort) - 1)
'   End If
'   lngCounter = 1
'   Select Case Mid(Combo3, 1, 1)
'      Case "1"
'         adoacc021.Open "select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, cu04 as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, (select ax201 as A, ax211 as B, ax205 as C, sum(ax206) as D, sum(ax207) as E from acc021 where ax211 is not null group by ax201, ax211, ax205) new, customer, staff sales where ax201 = a0201 and ax202 = a0202 and ax201 = A (+) and ax202 = B (+) and ax205 = C (+) and substr(ax208, 1, length(ax208) - 1) = cu01 and substr(ax208, length(ax208), 1) = cu02 and ax209 = sales.st01 (+) and (ax206 <> E or ax207 <> D or E is null)" & strSQL & _
'                        " union select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, a0i02 as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, (select ax201 as A, ax211 as B, ax205 as C, sum(ax206) as D, sum(ax207) as E from acc021 where ax211 is not null group by ax201, ax211, ax205) new, acc0i0, staff sales where ax201 = a0201 and ax202 = a0202 and ax201 = A (+) and ax202 = B (+) and ax205 = C (+) and ax208 = a0i01 and ax209 = sales.st01 (+) and (ax206 <> E or ax207 <> D or E is null)" & strSQL & _
'                        " union select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, staff.st02 as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, (select ax201 as A, ax211 as B, ax205 as C, sum(ax206) as D, sum(ax207) as E from acc021 where ax211 is not null group by ax201, ax211, ax205) new, staff, staff sales where ax201 = a0201 and ax202 = a0202 and ax201 = A (+) and ax202 = B (+) and ax205 = C (+) and ax208 = staff.st01 and ax209 = sales.st01 (+) and (ax206 <> E or ax207 <> D or E is null)" & strSQL & _
'                        " union select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, '' as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, (select ax201 as A, ax211 as B, ax205 as C, sum(ax206) as D, sum(ax207) as E from acc021 where ax211 is not null group by ax201, ax211, ax205) new, staff sales where ax201 = a0201 and ax202 = a0202 and ax201 = A (+) and ax202 = B (+) and ax205 = C (+) and ax209 = sales.st01 (+) and (ax206 <> E or ax207 <> D or E is null) and ax208 is null" & strSQL & _
'                        strSort, adoTaie, adOpenStatic, adLockReadOnly
'      Case "2"
'         adoacc021.Open "select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, cu04 as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, (select ax201 as A, ax211 as B, ax205 as C, sum(ax206) as D, sum(ax207) as E from acc021 where ax211 is not null group by ax201, ax211, ax205) new, customer, staff sales where ax201 = a0201 and ax202 = a0202 and ax201 = A (+) and ax202 = B (+) and ax205 = C (+) and substr(ax208, 1, length(ax208) - 1) = cu01 and substr(ax208, length(ax208), 1) = cu02 and ax209 = sales.st01 (+) and (ax206 = E and ax207 = D and E is not null)" & strSQL & _
'                        " union select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, a0i02 as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, (select ax201 as A, ax211 as B, ax205 as C, sum(ax206) as D, sum(ax207) as E from acc021 where ax211 is not null group by ax201, ax211, ax205) new, acc0i0, staff sales where ax201 = a0201 and ax202 = a0202 and ax201 = A (+) and ax202 = B (+) and ax205 = C (+) and ax208 = a0i01 and ax209 = sales.st01 (+) and (ax206 = E and ax207 = D and E is not null)" & strSQL & _
'                        " union select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, staff.st02 as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, (select ax201 as A, ax211 as B, ax205 as C, sum(ax206) as D, sum(ax207) as E from acc021 where ax211 is not null group by ax201, ax211, ax205) new, staff, staff sales where ax201 = a0201 and ax202 = a0202 and ax201 = A (+) and ax202 = B (+) and ax205 = C (+) and ax208 = staff.st01 and ax209 = sales.st01 (+) and (ax206 = E and ax207 = D and E is not null)" & strSQL & _
'                        " union select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, '' as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, (select ax201 as A, ax211 as B, ax205 as C, sum(ax206) as D, sum(ax207) as E from acc021 where ax211 is not null group by ax201, ax211, ax205) new, staff sales where ax201 = a0201 and ax202 = a0202 and ax201 = A (+) and ax202 = B (+) and ax205 = C (+) and ax209 = sales.st01 (+) and (ax206 = E and ax207 = D and E is not null)" & strSQL & _
'                        strSort, adoTaie, adOpenStatic, adLockReadOnly
'      Case Else
'         adoacc021.Open "select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, cu04 as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, customer, staff sales where ax201 = a0201 and ax202 = a0202 and substr(ax208, 1, length(ax208) - 1) = cu01 and substr(ax208, length(ax208), 1) = cu02 and ax209 = sales.st01 (+)" & strSQL & _
'                        " union select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, a0i02 as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, acc0i0, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 = a0i01 and ax209 = sales.st01 (+)" & strSQL & _
'                        " union select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, staff.st02 as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, staff, staff sales where ax201 = a0201 and ax202 = a0202 and ax208 = staff.st01 and ax209 = sales.st01 (+)" & strSQL & _
'                        " union select ax208, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Case, a0205, ax202, ax206, ax207, ax212, '' as cust, sales.st02 as salesname, ax213, ax214 from acc021, acc020, staff sales where ax201 = a0201 and ax202 = a0202 and ax209 = sales.st01 (+) and ax208 is null" & strSQL & _
'                        strSort, adoTaie, adOpenStatic, adLockReadOnly
'   End Select
'   If adoacc021.RecordCount = 0 Then
'      adoacc021.Close
'      adoaccrpt404.Close
'      Me.MousePointer = vbDefault
'      MsgBox MsgText(28), , MsgText(5)
'      Exit Sub
'   End If
'   Do While adoacc021.EOF = False
'      strGroup = ""
'      If Combo13 <> MsgText(601) Then
'         Select Case Mid(Combo13, 1, 1)
'            Case "1"
'               If IsNull(adoacc021.Fields("ax214").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax214").Value
'               End If
'            Case "2"
'               If IsNull(adoacc021.Fields("ax208").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax208").Value
'               End If
'            Case "3"
'               If IsNull(adoacc021.Fields("ax209").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax209").Value
'               End If
'            Case "4"
'               If IsNull(adoacc021.Fields("ax213").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax213").Value
'               End If
'         End Select
'      End If
'      If Combo5 <> MsgText(601) Then
'         Select Case Mid(Combo5, 1, 1)
'            Case "1"
'               If IsNull(adoacc021.Fields("ax214").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax214").Value
'               End If
'            Case "2"
'               If IsNull(adoacc021.Fields("ax208").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax208").Value
'               End If
'            Case "3"
'               If IsNull(adoacc021.Fields("ax209").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax209").Value
'               End If
'            Case "4"
'               If IsNull(adoacc021.Fields("ax213").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax213").Value
'               End If
'         End Select
'      End If
'      If Combo1 <> MsgText(601) Then
'         Select Case Mid(Combo1, 1, 1)
'            Case "1"
'               If IsNull(adoacc021.Fields("ax214").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax214").Value
'               End If
'            Case "2"
'               If IsNull(adoacc021.Fields("ax208").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax208").Value
'               End If
'            Case "3"
'               If IsNull(adoacc021.Fields("ax209").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax209").Value
'               End If
'            Case "4"
'               If IsNull(adoacc021.Fields("ax213").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax213").Value
'               End If
'         End Select
'      End If
'      If Combo2 <> MsgText(601) Then
'         Select Case Mid(Combo2, 1, 1)
'            Case "1"
'               If IsNull(adoacc021.Fields("ax214").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax214").Value
'               End If
'            Case "2"
'               If IsNull(adoacc021.Fields("ax208").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax208").Value
'               End If
'            Case "3"
'               If IsNull(adoacc021.Fields("ax209").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax209").Value
'               End If
'            Case "4"
'               If IsNull(adoacc021.Fields("ax213").Value) = False Then
'                  strGroup = strGroup & adoacc021.Fields("ax213").Value
'               End If
'         End Select
'      End If
'      If strName <> strGroup Then
'         douAmount = Val(adoacc021.Fields("ax206").Value) - Val(adoacc021.Fields("ax207").Value)
'         strName = strGroup
'      Else
'         douAmount = douAmount + Val(adoacc021.Fields("ax206").Value) - Val(adoacc021.Fields("ax207").Value)
'      End If
'      adoaccrpt404.AddNew
'      adoaccrpt404.Fields("r40401").Value = strUserNum
'      Select Case Mid(Combo13, 1, 1)
'         Case "1"
'            If IsNull(adoacc021.Fields("Case").Value) Then
'               adoaccrpt404.Fields("r40402").Value = Null
'            Else
'               adoaccrpt404.Fields("r40402").Value = adoacc021.Fields("Case").Value
'            End If
'         Case "2"
'            If IsNull(adoacc021.Fields("ax208").Value) Then
'               adoaccrpt404.Fields("r40402").Value = Null
'            Else
'               adoaccrpt404.Fields("r40402").Value = adoacc021.Fields("ax208").Value & adoacc021.Fields("cust").Value
'            End If
'         Case "4"
'            If IsNull(adoacc021.Fields("ax213").Value) Then
'               adoaccrpt404.Fields("r40402").Value = Null
'            Else
'               adoaccrpt404.Fields("r40402").Value = adoacc021.Fields("ax213").Value
'            End If
'         Case Else
'            If IsNull(adoacc021.Fields("ax209").Value) Then
'               adoaccrpt404.Fields("r40402").Value = Null
'            Else
'               adoaccrpt404.Fields("r40402").Value = adoacc021.Fields("ax209").Value & adoacc021.Fields("salesname").Value
'            End If
'      End Select
'      If IsNull(adoacc021.Fields("a0205").Value) Then
'         adoaccrpt404.Fields("r40405").Value = Null
'      Else
'         adoaccrpt404.Fields("r40405").Value = adoacc021.Fields("a0205").Value
'      End If
'      adoaccrpt404.Fields("r40406").Value = adoacc021.Fields("ax202").Value
'      If IsNull(adoacc021.Fields("ax212").Value) Then
'         adoaccrpt404.Fields("r40407").Value = Null
'      Else
'         adoaccrpt404.Fields("r40407").Value = adoacc021.Fields("ax212").Value
'      End If
'      If IsNull(adoacc021.Fields("ax206").Value) Then
'         adoaccrpt404.Fields("r40408").Value = 0
'      Else
'         adoaccrpt404.Fields("r40408").Value = Val(adoacc021.Fields("ax206").Value)
'      End If
'      If IsNull(adoacc021.Fields("ax207").Value) Then
'         adoaccrpt404.Fields("r40409").Value = 0
'      Else
'         adoaccrpt404.Fields("r40409").Value = Val(adoacc021.Fields("ax207").Value)
'      End If
'      adoaccrpt404.Fields("r40410").Value = douAmount
'      If IsNull(adoacc021.Fields("ax213").Value) Then
'         adoaccrpt404.Fields("r40411").Value = 0
'      Else
'         adoaccrpt404.Fields("r40411").Value = Val(adoacc021.Fields("ax213").Value)
'      End If
'      adoaccrpt404.Fields("r40412").Value = strGroup
'      adoaccrpt404.Fields("r40413").Value = lngCounter
'      adoaccrpt404.UpdateBatch
'      lngCounter = lngCounter + 1
'      adoacc021.MoveNext
'   Loop
'   adoacc021.Close
'   adoaccrpt404.Close
'   Me.MousePointer = vbDefault
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   Dim bolCancel As Boolean 'Add by Amy 2020/04/14
   
   FormCheck = False
   If Text7.Text = "" Then
      MsgBox "對沖條件不可空白！"
      Text7.SetFocus
      Text7_GotFocus
      Exit Function
   End If
   'Add by Amy 2020/04/14 +公司別檢查
   If Trim(cboComp) <> MsgText(601) Then
        Call CboComp_Validate(bolCancel)
        If bolCancel = True Then
            Exit Function
        End If
   End If
   'end 2020/04/14
   FormCheck = True
End Function

Private Sub txtAX208_GotFocus(Index As Integer)
   If Index = 1 Then
      If txtAX208(Index).Text = "" Then txtAX208(Index).Text = txtAX208(Index - 1).Text
   End If
   TextInverse txtAX208(Index)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtAX208(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtAX208_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtAX208_Validate(Index As Integer, Cancel As Boolean)
   If Len(txtAX208(Index)) = 6 Then
      txtAX208(Index).Text = AfterZero(txtAX208(Index).Text)
   End If
End Sub

Private Sub txtAX209_GotFocus(Index As Integer)
   If Index = 1 Then
      If txtAX209(Index).Text = "" Then txtAX209(Index).Text = txtAX209(Index - 1).Text
   End If
   TextInverse txtAX209(Index)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtAX209(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtAX209_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtAX211_GotFocus(Index As Integer)
   If Index = 1 Then
      If txtAX211(Index).Text = "" Then txtAX211(Index).Text = txtAX211(Index - 1).Text
   End If
   TextInverse txtAX211(Index)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtAX211(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtAX211_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtAX213_GotFocus(Index As Integer)
   If Index = 1 Then
      If txtAX213(Index).Text = "" Then txtAX213(Index).Text = txtAX213(Index - 1).Text
   End If
   TextInverse txtAX213(Index)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtAX213(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtAX213_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtAX214_GotFocus(Index As Integer)
   If Index = 1 Then
      If txtAX214(Index).Text = "" Then txtAX214(Index).Text = txtAX214(Index - 1).Text
   End If
   TextInverse txtAX214(Index)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtAX214(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtAX214_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'2010/4/9 ADD BY SONIA 預設本所案號尾碼
Private Sub txtAX214_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      txtAX214(0) = CaseNoZero(txtAX214(0))
      If txtAX214(0) <> "" Then
         If Mid(txtAX214(0), 1, 2) = "TF" Then
            txtAX214(1) = Left(txtAX214(0), 7) & "9999"
         Else
            txtAX214(1) = Left(txtAX214(0), 10) & "99"
         End If
      End If
   End If
End Sub
'2010/4/9 END
