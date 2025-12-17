VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1125 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "特殊收據"
   ClientHeight    =   5832
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9108
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5832
   ScaleWidth      =   9108
   Begin VB.CommandButton Command6 
      Caption         =   "檢視接洽單"
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
      Left            =   2910
      TabIndex        =   48
      Top             =   30
      Width           =   1410
   End
   Begin VB.TextBox txtPrintNo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   19
      Top             =   5310
      Width           =   345
   End
   Begin VB.OptionButton OptKind 
      Caption         =   "母子號"
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   6
      Top             =   870
      Width           =   855
   End
   Begin VB.OptionButton OptKind 
      Caption         =   "拆新號"
      Height          =   255
      Index           =   0
      Left            =   6960
      TabIndex        =   5
      Top             =   870
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   180
      TabIndex        =   44
      Top             =   4980
      Width           =   7776
      Begin VB.CheckBox Check3 
         Caption         =   "3.代理人請款之匯款日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5244
         TabIndex        =   22
         Top             =   30
         Width           =   2544
      End
      Begin VB.CheckBox Check3 
         Caption         =   "2.代理人請款日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3324
         TabIndex        =   21
         Top             =   30
         Width           =   1872
      End
      Begin VB.CheckBox Check3 
         Caption         =   "1.送件日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2030
         TabIndex        =   20
         Top             =   30
         Width           =   1248
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "收據自動列印時間點"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   0
         TabIndex        =   45
         Top             =   60
         Width           =   1990
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "重整"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1935
      TabIndex        =   43
      Top             =   1590
      Width           =   825
   End
   Begin VB.CheckBox Check2 
      Caption         =   "有手動調整次序"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   42
      Top             =   1620
      Width           =   1680
   End
   Begin VB.TextBox txtDate 
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
      Left            =   5670
      MaxLength       =   7
      TabIndex        =   4
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  '平面
      Height          =   375
      Left            =   2745
      TabIndex        =   38
      Text            =   "Text3"
      Top             =   2550
      Width           =   1635
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8130
      Picture         =   "Frmacc1125.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   23
      ToolTipText     =   "取消"
      Top             =   4530
      Width           =   555
   End
   Begin VB.CheckBox Check1 
      Caption         =   "收據暫不列印"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8000
      TabIndex        =   18
      Top             =   4980
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox Text1 
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
      Left            =   5670
      MaxLength       =   1
      TabIndex        =   2
      Top             =   480
      Width           =   600
   End
   Begin VB.TextBox txtNo 
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
      Left            =   5670
      TabIndex        =   7
      Top             =   1200
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "建立"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6300
      TabIndex        =   8
      Top             =   1200
      Width           =   600
   End
   Begin VB.TextBox txtCustNo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1050
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
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
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   2610
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8100
      TabIndex        =   10
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton Command3 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7245
      TabIndex        =   9
      Top             =   1320
      Width           =   840
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1245
      Left            =   90
      TabIndex        =   24
      Top             =   300
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   2180
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      FormatString    =   "本所案號|案件性質|服務費|規費|合併"
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
      _Band(0).Cols   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   1965
      Left            =   90
      TabIndex        =   26
      Top             =   1890
      Width           =   8835
      _ExtentX        =   15600
      _ExtentY        =   3450
      _Version        =   393216
      Cols            =   23
      FixedCols       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   $"Frmacc1125.frx":066A
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
      _Band(0).Cols   =   23
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Left            =   6840
      TabIndex        =   15
      Top             =   4200
      Width           =   1785
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3149;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtMain 
      Height          =   315
      Index           =   3
      Left            =   5535
      TabIndex        =   14
      Top             =   4200
      Width           =   1260
      VariousPropertyBits=   671105051
      MaxLength       =   15
      Size            =   "2222;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtMain 
      Height          =   345
      Index           =   5
      Left            =   2175
      TabIndex        =   17
      Top             =   4560
      Width           =   5865
      VariousPropertyBits=   -1466941413
      MaxLength       =   1000
      ScrollBars      =   2
      Size            =   "10345;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtMain 
      Height          =   315
      Index           =   6
      Left            =   4395
      TabIndex        =   13
      Top             =   4200
      Width           =   1125
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1984;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCustName 
      Height          =   315
      Left            =   6750
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2130
      VariousPropertyBits=   671105051
      BackColor       =   14737632
      MaxLength       =   30
      Size            =   "3757;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtMain 
      Height          =   315
      Index           =   1
      Left            =   1095
      TabIndex        =   16
      Top             =   4560
      Width           =   465
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "820;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboTitle 
      Height          =   330
      Left            =   495
      TabIndex        =   12
      Top             =   4200
      Width           =   3885
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "6853;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtMain 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4200
      Width           =   330
      VariousPropertyBits=   671105051
      Size            =   "582;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印統一編號      (Y:印)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   180
      TabIndex        =   47
      Top             =   5340
      Width           =   2190
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "介紹案源同仁"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   6900
      TabIndex        =   46
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "手開收據編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5535
      TabIndex        =   41
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收據日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   4770
      TabIndex        =   40
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   1680
      TabIndex        =   39
      Top             =   4620
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "預定收款日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   4395
      TabIndex        =   37
      Top             =   3960
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4770
      TabIndex        =   36
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "張數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4770
      TabIndex        =   35
      Top             =   1200
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   4770
      TabIndex        =   34
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lblFee 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6375
      TabIndex        =   33
      Top             =   1635
      Width           =   690
   End
   Begin VB.Label lblService 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4215
      TabIndex        =   32
      Top             =   1635
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "待開規費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   5475
      TabIndex        =   31
      Top             =   1635
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "待開服務費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   3090
      TabIndex        =   30
      Top             =   1635
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1800
      Left            =   60
      Top             =   3900
      Width           =   8925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "個人/公司"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   150
      TabIndex        =   29
      Top             =   4620
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   540
      TabIndex        =   28
      Top             =   3960
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "序號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   27
      Top             =   3960
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "待開收據資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   25
      Top             =   90
      Width           =   1260
   End
End
Attribute VB_Name = "Frmacc1125"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Created by Morgan 2011/9/7
Option Explicit

Public m_OldNo As String '原收據編號
Public frmCallForm As Form '呼叫的視窗
Public adoquery As New ADODB.Recordset    'add by sonia 2017/6/22
Dim m_dftColor As Long '預設顏色
Dim m_dftColor2 As Long '預設顏色2
Dim m_dftColor3 As Long '點選顏色
Dim m_lstRow As Integer '最後點選行
Dim m_lstColor As Long '最後點選行顏色
Dim m_CP12 As String '業務區
Dim m_CP13 As String '智權人員
Dim iRow As Integer, iCol As Integer
Dim m_bolManualReceipt As Boolean '是否有設定手開收據
Dim m_bolSplitMail As Boolean '是否拆收據發Mail
Dim m_strMailDesc As String, m_strMailSubject As String
Dim m_strChkCompany As String, m_strCaseNo As String '檢查是否為專利商標公司 Added by Morgan 2012/9/12
Dim m_Nation As String 'Add By Sindy 2012/11/15
Dim m_CP31 As String 'Add By Sindy 2013/12/26
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String 'Add By Sindy 2013/12/26
Dim m_CP09 As String 'Add By Sindy 2014/2/11
Dim adocheck As New ADODB.Recordset 'Add By Sindy 2014/1/13
Public m_CRL119 As String, m_CRL49 As String, m_CP140 As String, m_CallForm 'Add By Sindy 2014/2/11
Public m_CRL02 As String 'Add By Sindy 2020/3/31
Private Const monLen As Integer = 9 'Added by Lydia 2016/09/05 母號收據號碼長度
Dim tmpfrm As Form 'Add By Sindy 2023/1/4
Dim m_CRL153 As String, strShowCRL153 As String 'Added by Lydia 2024/08/05 國內接洽單：DEBIT NOTE請款選項
Dim strPropertyCode As String, strPromoterNo As String 'Added by Lydia 2025/11/13 改全域變數

Private Sub cboTitle_GotFocus()
   OpenIme
   cboTitle.SelStart = 0
   cboTitle.SelLength = Len(cboTitle.Text)
End Sub

Private Sub cboTitle_LostFocus()
   CloseIme
End Sub

'add by sonia 2015/11/26
Private Sub Check3_Click(Index As Integer)
   If Check3(0).Value = 1 Or Check3(1).Value = 1 Or Check3(2).Value = 1 Then
      Check1.Value = 1
   Else
      Check1.Value = 0
   End If
End Sub
'end 2015/11/26

'Add By Sindy 2013/12/25
Private Sub Check3_GotFocus(Index As Integer)
   Select Case Index
      Case 0
         Check3(1).Value = 0
         Check3(2).Value = 0
      Case 1
         Check3(0).Value = 0
         Check3(2).Value = 0
      Case 2
         Check3(0).Value = 0
         Check3(1).Value = 0
   End Select
End Sub

Private Sub Command1_Click()
   Dim ii As Integer, jj As Integer, kk As Integer, idx As Integer
   Dim strItem As String, lColor As Long
   Dim m_CU173 As String
   Dim m_CU11 As String, m_CU168 As String 'Add by Amy 2025/02/20
   
   If Val(txtNo) = 0 Then
      MsgBox "請輸入張數!!"
      txtNo.SetFocus
      Exit Sub
   End If
   
   If m_OldNo = "" Then
      If Val(txtNo) > 1 And txtMain(3) <> "" Then
         MsgBox "張數超過一張時手開收據號需為空白!!"
         txtMain(3).SetFocus
         Exit Sub
      End If
   End If
   
   If MSHFlexGrid2.Rows > 1 Then
      If MSHFlexGrid2.TextMatrix(1, 0) <> "" Then
         If MsgBox("是否確定要重新建立收據資料！" & vbCrLf & "(前次建立資料將清除)", vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
         End If
      End If
   End If
   
   If InsertCheck = False Then Exit Sub
   
   MSHFlexGrid2.Clear
   GridHead2
   MSHFlexGrid2.Visible = False
   With MSHFlexGrid1
   idx = 0
   For jj = 1 To Val(txtNo)
      If lColor = m_dftColor2 Then
         lColor = m_dftColor
      Else
         lColor = m_dftColor2
      End If
      
      'Modify By Sindy 2017/3/24
      m_CU173 = ""
      'Modify By Sindy 2019/5/22 + txtCustNo
      'Modify by Amy 2025/02/20 +m_CU11
      Call GetTitleCustData(cboTitle.Text, txtCustNo, "", , , , , , , , , , , , , , , , , , , , , m_CU168, , , , m_CU11, , , m_CU173)
      txtPrintNo.Text = m_CU173
      '2017/3/24 END
      
      For ii = 1 To .Rows - 1
         strItem = jj & vbTab & vbTab & cboTitle
         strItem = strItem & vbTab & txtMain(1) & vbTab & .TextMatrix(ii, 4) & vbTab & Format(txtMain(6), "###/##/##")
         strItem = strItem & vbTab & IIf(Check1.Value = "1", "N", "")
         strItem = strItem & vbTab & .TextMatrix(ii, 0) & vbTab & .TextMatrix(ii, 1)
         strItem = strItem & vbTab & Format(Val(.TextMatrix(ii, 2)) / Val(txtNo), "0")
         strItem = strItem & vbTab & Format(Val(.TextMatrix(ii, 3)) / Val(txtNo), "0")
         strItem = strItem & vbTab & vbTab & vbTab & vbTab & vbTab 'Modified by Morgan 2017/11/17+a0j20,a0j21
         
         '拆收據
         If m_OldNo <> "" Then
            If jj = 1 Then
               strItem = strItem & vbTab & m_OldNo 'Memo by Lydia 2016/09/05 第一筆保留原收據編號
            Else
               strItem = strItem & vbTab  'Memo by Lydia 2016/09/05 第二筆以後按確定存檔時才取得編號
            End If
         Else
            strItem = strItem & vbTab & txtMain(3)
         End If
         
         'Modify By Sindy 2012/11/15
         'strItem = strItem & vbTab & .TextMatrix(ii, 1) & vbTab & .TextMatrix(ii, 5) & vbTab & .TextMatrix(ii, 6) & vbTab & vbTab
         'Modify By Sindy 2013/12/26
         'strItem = strItem & vbTab & .TextMatrix(ii, 1) & vbTab & .TextMatrix(ii, 5) & vbTab & .TextMatrix(ii, 6) & vbTab & IIf(Me.Check3(0).Value = 1, "1", IIf(Me.Check3(1).Value = 1, "2", IIf(Me.Check3(2).Value = 1, "3", ""))) & vbTab
         'strItem = strItem & vbTab & .TextMatrix(ii, 1) & vbTab & .TextMatrix(ii, 5) & vbTab & .TextMatrix(ii, 6) & vbTab & IIf(Me.Check3(0).Value = 1, "1", IIf(Me.Check3(1).Value = 1, "2", IIf(Me.Check3(2).Value = 1, "3", ""))) & vbTab & Left(Trim(Combo2.Text), 5) & vbTab
         'Modify By Sindy 2017/3/17
         'Modify by Amy 2025/02/20 +m_CU11/m_CU168
         strItem = strItem & vbTab & .TextMatrix(ii, 1) & vbTab & .TextMatrix(ii, 5) & vbTab & .TextMatrix(ii, 6) & vbTab & IIf(Me.Check3(0).Value = 1, "1", IIf(Me.Check3(1).Value = 1, "2", IIf(Me.Check3(2).Value = 1, "3", ""))) & vbTab & Left(Trim(Combo2.Text), 5) & vbTab & txtPrintNo.Text & _
                           vbTab & m_CU11 & vbTab & m_CU168 & vbTab
         
         idx = idx + 1
         MSHFlexGrid2.AddItem strItem, idx
         
         MSHFlexGrid2.TextMatrix(idx, 15) = txtMain(5) 'Added by Morgan 2023/5/11
         
         For kk = 0 To MSHFlexGrid2.Cols - 1
            MSHFlexGrid2.row = idx
            MSHFlexGrid2.col = kk
            MSHFlexGrid2.CellBackColor = lColor
         Next
      Next
   Next
   MSHFlexGrid2.Rows = idx + 1
   SortData
   End With
   MSHFlexGrid2.Visible = True
   SetRestValue
   m_lstRow = 0
   txtMain(0) = ""
   ClearInput
   
   'Added by Lydia 2016/09/07 提示
   If OptKind(1).Value = True Then MsgBox "請將收據餘額留在母號!"
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Command3_Click()
Dim i As Integer 'Add By Sindy 2014/1/29
   
   'Add by Sindy 2021/12/14 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Sub
   End If
   
   If TxtValidate = True Then
      If m_OldNo <> "" Then
         If FormSave1 = True Then
            LockForm
         End If
      Else
         If FormSave = True Then
            If m_bolSplitMail = True Then
               PUB_SendMail strUserNum, "83002", "", m_strMailSubject, m_strMailDesc
            End If
            LockForm
         End If
      End If
      'Add By Sindy 2014/1/29 若收據開J公司,但案件的特殊出名公司未輸入時,同時發E-MAIL
      With MSHFlexGrid2
         For i = 1 To .Rows - 1
            If .TextMatrix(i, 16) <> "" Then
               Call PUB_ChkJCompanyRecv_Mail(.TextMatrix(i, 16), Text1)
               Exit For
            End If
         Next i
      End With
      '2014/1/29 END
      'add by sonia 2024/11/6
      If OptKind(1).Value = True Then
         MsgBox "母子號的設定收據列印次數為1；如欲列印，請自行更改列印次數！"
      End If
      Call ChkWTConsentMail 'Add by Amy 2025/02/20
      'end 2024/11/6
      
      Call PUB_SendMailCache 'Added by Lydia 2025/11/13
   End If
   
   'Add by Sindy 2023/1/4 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q", tmpfrm) = True Then
      Unload tmpfrm 'frm090801_Q
   End If
   '2023/1/4 END
End Sub

Private Sub LockForm()
   Dim oControl As Control
   Command1.Enabled = False
   Command2.Caption = "前畫面"
   Command3.Enabled = False
   Command4.Enabled = False
   cboTitle.Enabled = False
   For Each oControl In Me.Controls
      If TypeName(oControl) = "TextBox" Then
         oControl.Locked = True
      End If
      If TypeName(oControl) = "CheckBox" Then
         oControl.Enabled = False
      End If
   Next
End Sub

Private Sub Command4_Click()
   If m_lstRow > 0 Then
      If MSHFlexGrid2.Rows > 2 Then
         MSHFlexGrid2.RemoveItem m_lstRow
         m_lstRow = 0
         SetRestValue
      Else
         MsgBox "只剩一筆資料，不可再刪除！"
         Exit Sub
      End If
   End If
End Sub

Private Sub Command5_Click()
   If Check2.Value = vbChecked Then
      ReSort
   Else
      SortData
   End If
End Sub

'Add By Sindy 2022/11/24 檢視檢洽單
Private Sub Command6_Click()
   Call PUB_Queryfrm090801(m_CP140, "", Me)
End Sub

Private Sub Form_Activate()
   If m_bolManualReceipt Then
      MsgBox "有收文號設定為手開收據，存檔前請自行確認!!", vbInformation, "手開收據提醒"
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
On Error GoTo Checking
   Select Case KeyCode
      Case vbKeyInsert
         If m_lstRow > 0 And Command1.Enabled = True Then
            UpdateRow
         End If
      Case vbKeyF12
         
      Case Else
         KeyEnter KeyCode
   End Select
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & "/" & MsgText(107)
   Exit Sub
   
Checking:
   MsgBox Err.Description, , MsgBox(5)
End Sub

Private Sub UpdateCol()
   Dim ii As Integer
   If txtInput <> txtInput.Tag Then
      With MSHFlexGrid2
      If iCol = 9 Or iCol = 10 Then
         .TextMatrix(iRow, iCol) = Val(txtInput.Text)
         SetRestValue
      ElseIf iCol = 8 Or iCol = 1 Then
         .TextMatrix(iRow, iCol) = txtInput.Text
      
      ElseIf iCol = 11 Or iCol = 12 Or iCol = 13 Or iCol = 14 Then
         For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 0) = .TextMatrix(iRow, 0) And .TextMatrix(ii, 7) = .TextMatrix(iRow, 7) Then
               .TextMatrix(ii, iCol) = txtInput.Text
               'Added by Morgan 2017/11/17
               If iCol = 13 Then
                  txtInput.Text = "N"
                  .TextMatrix(ii, 14) = txtInput.Text
               End If
               'end 2017/11/17
            End If
         Next
      Else
         For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 0) = .TextMatrix(iRow, 0) Then
               .TextMatrix(ii, iCol) = txtInput.Text
            End If
         Next
      End If
      End With
   End If
End Sub

Private Sub UpdateRow()
   Dim ii As Integer
   Dim m_CU173 As String
   Dim m_CU11 As String, m_CU168 As String 'Add by Amy 2025/02/20
   
   If InsertCheck(True) = False Then Exit Sub
   
   With MSHFlexGrid2
   For ii = 1 To .Rows - 1
      If .TextMatrix(ii, 0) = txtMain(0) Then
         .TextMatrix(ii, 2) = cboTitle
         .TextMatrix(ii, 3) = txtMain(1)
         .TextMatrix(ii, 5) = Format(txtMain(6), "###/##/##")
         .TextMatrix(ii, 6) = IIf(Check1.Value = 1, "N", "")
         .TextMatrix(ii, 15) = txtMain(5)
         .TextMatrix(ii, 16) = txtMain(3)
         'Add By Sindy 2012/11/15
         .TextMatrix(ii, 20) = IIf(Me.Check3(0).Value = 1, "1", IIf(Me.Check3(1).Value = 1, "2", IIf(Me.Check3(2).Value = 1, "3", "")))
         '2012/11/15 End
         .TextMatrix(ii, 21) = Left(Trim(Combo2.Text), 5) 'Add By Sindy 2013/12/26
         'Modify By Sindy 2017/3/24
         m_CU173 = ""
         'Modify By Sindy 2019/5/22 + txtCustNo
         'Modify by Amy 2025/02/20 +m_CU11
         Call GetTitleCustData(cboTitle.Text, txtCustNo, "", , , , , , , , , , , , , , , , , , , , , m_CU168, , , , m_CU11, , , m_CU173)
         txtPrintNo.Text = m_CU173
         '2017/3/24 END
         .TextMatrix(ii, 22) = txtPrintNo.Text 'Add By Sindy 2017/3/17
         .TextMatrix(ii, 23) = m_CU11
         .TextMatrix(ii, 24) = m_CU168
         'end 2025/02/20
      End If
   Next
   End With
   
   '新增新的收據抬頭
   If cboTitle.ListIndex = -1 Then
      For ii = 0 To cboTitle.ListCount - 1
         If cboTitle.List(ii) = cboTitle Then
            Exit For
         End If
      Next
      If ii = cboTitle.ListCount Then
         cboTitle.AddItem cboTitle, 0
      End If
   End If
   
   SetRestValue
   ClearInput
End Sub

Private Sub UnSelectRow()
   With MSHFlexGrid2
   '還原上一筆點選
   If m_lstRow > 0 Then
      .row = m_lstRow
      For intI = 0 To .Cols - 1
         .col = intI
         .CellBackColor = m_lstColor
      Next
      m_lstRow = 0
   End If
   End With
End Sub

Private Sub ClearInput()
   UnSelectRow
   txtMain(0) = ""
   cboTitle = ""
   txtMain(1) = ""
   txtMain(3) = ""
   txtMain(5) = ""
   txtMain(6) = ""
   Check1.Value = 0
   'Add By Sindy 2012/11/15
   Me.Check3(0).Value = 0
   Me.Check3(1).Value = 0
   Me.Check3(2).Value = 0
   'txtSales = "": lblSales.Caption = "" 'Add By Sindy 2013/12/26
   txtPrintNo.Text = "" 'Add By Sindy 2017/3/17
End Sub

Private Sub Form_Load()
Dim adoRst As ADODB.Recordset
Dim strCU125 As String 'Add By Sindy 2016/12/9
'Dim strCRL49 As String 'Add by Amy 2020/03/27
   
   'Modify by Amy 2023/08/23 W9150 H5955
   PUB_InitForm Me, 9195, 6270 '5700
   '底色
   m_dftColor = &HFFFFFF
   '底色2
   m_dftColor2 = RGB(&HFF, &HFA, &HCD)
   '底色3
   m_dftColor3 = &HFFC0C0
   ResetForm
   If m_OldNo <> "" Then
      OpenTable1
      'Added by Lydia 2016/09/05 拆收據可選擇拆新號或母子號
      OptKind(0).Visible = True: OptKind(1).Visible = True
      strExc(0) = GetAutoNo(m_CP09, True)
      '未拆過母子號可選擇拆新號
      If Mid(strExc(0), monLen + 1, 1) = "1" Then OptKind(0).Value = 1
      
   Else
      OpenTable
      'Added by Lydia 2016/09/05
      OptKind(0).Visible = False: OptKind(1).Visible = False
   End If
   txtInput.Visible = False
   txtDate = strSrvDate(2)
   
   'Modify by Amy 2020/03/26 Mark InvoiceStartDate,加新公司別判斷
   'Add By Sindy 2013/12/26
'   If strSrvDate(1) >= InvoiceStartDate Then
      If m_strChkCompany <> "" Then
         If strSrvDate(1) >= 智慧所更名日 Then
            Text1.Enabled = False
            If ChkAccReceiptComp(2, m_CP01) = True Then
                Text1 = "L"
            ElseIf m_strChkCompany = "J" Then
                Text1 = "J"
            Else
                If m_strChkCompany = "T" Then Text1 = "2"
            End If
            '拆收據作業進入,公司別維持不變,並鎖住;其他作業進入,公司別 不 為 J 或 L才可改
            If UCase(m_CallForm) <> UCase("frmacc11m0") And m_strChkCompany <> "J" And m_strChkCompany <> "L" Then
                Text1.Enabled = True
            End If
         Else
            If m_strChkCompany = "T" Then Text1 = "1": Text1.Enabled = False
            If m_strChkCompany = "J" Then Text1 = "J": Text1.Enabled = False
         End If
      End If
      If m_CP31 = "Y" Then
         '新案
         If m_strChkCompany <> "" Then
            MsgBox "請注意," & m_strCaseNo & "有設定特殊出名公司,請檢查與接洽記錄單是否相同!!", vbInformation, "收據公司別提醒"
         'CANCEL BY SONIA 2017/12/5 因接洽單已改預設方式,故此處不再提醒
         'Else
         '   If (m_CP01 = "P" Or m_CP01 = "T") And Left(m_CP12, 1) <> "F" And m_Nation = "020" Then
         '      Text1 = "J": Text1.Enabled = True
         '      MsgBox "請注意,大陸新案,請注意接洽記錄單的收據公司別!!", vbInformation, "收據公司別提醒"
         '   End If
         End If
      End If
'   Else
'   '2013/12/26 END
'      'Added by Morgan 2012/9/12
'      If m_strChkCompany <> "" And Text1 <> "1" And m_CP31 = "Y" Then MsgBox "請注意,專利案" & m_strCaseNo & "有設定以專利商標出名!!", vbInformation, "收據公司別提醒"
'   End If
   'end 2020/03/26
   
   'Add By Sindy 2014/2/11
   m_CallForm = ""
   If m_CP140 <> "" Then '直接收文
      strCU125 = PUB_GetApplCU125(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) 'Add By Sindy 2016/12/9 組合業務備註
      If strCU125 <> "" Then strCU125 = "業務備註：" & vbCrLf & strCU125
      'Add By Sindy 2015/8/10
      'Modified by Lydia 2024/08/05 +CRL153
      strSql = "select CRL47,CRL153 from consultrecordlist where CRL01='" & m_CP140 & "'"
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If "" & adoRst.Fields("CRL47") = "Y" Then
      'Modify By Sindy 2016/12/9
            'MsgBox "請注意, 接洽記錄單有加註「貼印花」!!", vbInformation, "提醒"
            strCU125 = "請注意, 接洽記錄單有加註「貼印花」!!" & IIf(strCU125 <> "", vbCrLf & vbCrLf & strCU125, "")
         End If
         m_CRL153 = "" & adoRst.Fields("CRL153") 'Added by Lydia 2024/08/05 國內接洽單：DEBIT NOTE請款選項
      End If
      If strCU125 <> "" Then
         MsgBox strCU125, vbInformation, "提醒"
      End If
      '2016/12/9 END
      Set adoRst = Nothing
      '2015/8/10 END
      '有特殊收據資料或收據公司與預設公司不同時,顯示接洽單的特殊收據內容給使用者看再開立收據
      'Modify By Sindy 2014/10/23 收據公司與預設公司不同時,增加檢查必須為新案(m_CP31 = "Y")
      'Modify by Amy 2020/03/27 改為公司編號判斷,增加 strCRL49=4(智慧所)
'      strCRL49 = ChgCRL49ToBKeeping(m_CRL49)
      'Modify By Sindy 2020/3/31
      If m_CRL119 = "Y" Or _
         (m_CRL02 < 事務所合併日 And m_CP31 = "Y" And IIf(m_CRL49 = "3", "智權公司", IIf(m_CRL49 = "2", "專利商標", "專利法律")) <> Text12) Or _
         (m_CRL02 >= 事務所合併日 And m_CP31 = "Y" And IIf(m_CRL49 = "J", "智權", IIf(m_CRL49 = "L", "法律所", "智慧所")) <> Text12) Then
      '2020/3/31 END
         frm090801_7.SetParent Me
         frm090801_7.m_stCRL01 = m_CP140
         m_CallForm = "frm090801_7"
         frm090801_7.Show 'vbModal Modify By Sindy 2024/3/26 開特殊收據畫面不要用強制表單方式開啟
      End If
   'Added by Morgan 2015/12/3
   Else
      strExc(1) = PUB_ReadCP64Tag(m_CP09, "開收據提醒")
      If strExc(1) <> "" Then
         MsgBox strExc(1), vbInformation, "提醒"
      End If
   'end 2015/12/3
   End If
   '2014/2/11 END
   
   'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
   Label1(7).Visible = False
   txtMain(6).Visible = False
   
   'Add By Sindy 2022/11/30
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      Command6.Visible = True
   Else
      Command6.Visible = False
   End If
   '2022/11/30 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Sindy 2022/12/17 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q", tmpfrm) = True Then
      Unload tmpfrm 'frm090801_Q
   End If
   '2022/12/17 END
   
   If UCase(m_CallForm) = UCase("frm090801_7") Then
      m_CallForm = ""
      Unload frm090801_7
   End If
   frmCallForm.Enabled = True
   frmCallForm.Show
   Set Frmacc1125 = Nothing
End Sub

Private Sub OpenTable()
   'Dim strPropertyCode As String, strPromoterNo As String 'Mark by Lydia 2025/11/13 改全域變數
   Dim strSpecCompany As String 'Add By Sindy 2013/12/26
   
   'Modify By Sindy 2012/11/15 +,cp151
   'Modify By Sindy 2013/12/26 +,cp31
   strExc(0) = "select cp01||cp02||cp03||cp04 本所案號,getcp10desc(cp01,cp10,a0j04) 案件性質,a0j09||'' 服務費,a0j10||'' 規費" & _
      ",a0j07,cp09,cp05,a0j04,a0j11,cp01,cp02,cp03,cp04,cp10,cp12,cp13,cp14,cp140,a0j08,a0j02,cp151,CP31" & _
      " from acc0j0,caseprogress,casepropertymap" & _
      " where cp09(+) = a0j01 and a0j06 = '" & MsgText(602) & "' and a0j13=a0j01" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 order by 1,cp05,cp09"
   intI = 1
   Set MSHFlexGrid1.Recordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid1.Recordset
         'Add By Sindy 2013/12/26
         m_CP31 = "" & .Fields("CP31")
         m_CP01 = .Fields("CP01")
         m_CP02 = .Fields("CP02")
         m_CP03 = .Fields("CP03")
         m_CP04 = .Fields("CP04")
         '2013/12/26 END
         m_CP12 = .Fields("cp12")
         m_CP13 = .Fields("cp13")
         m_CP09 = .Fields("CP09") 'Add By Sindy 2014/2/11
         txtCustNo = .Fields("a0j11")
         'Modify By Sindy 2012/11/15
         'strNation = .Fields("a0j04")
         m_Nation = .Fields("a0j04")
         '2012/11/15 End
         
         strPropertyCode = .Fields("cp10")
         strPromoterNo = "" & .Fields("cp14")
         '設定公司別
         SetCompany m_CP01, m_Nation, strPropertyCode, strPromoterNo, m_CP02, m_CP03, m_CP04
         '設定收據抬頭
         SetReceiptTitle
         'Add By Sindy 2014/2/11 C類該案號最新收據抬頭
         strTitle = GetReceiptTitle_C(m_CP09, m_CP01 & m_CP02 & m_CP03 & m_CP04)
         If strTitle <> "" Then
            cboTitle.Text = strTitle
         End If
         '2014/2/11 END
         
         'Modify by Sindy 2015/10/15 若為境外公司 只能為1.個人且不可改
         If PUB_GetTaxNo(cboTitle.Text, 1) = "Y" Then
            txtMain(1) = "1"
            txtMain(1).Locked = True
         Else
            txtMain(1) = "2"
            txtMain(1).Locked = False
         End If
         '2015/10/15 END
         
         '智權人員收文設定收據抬頭
         .MoveFirst
         m_strChkCompany = "": m_strCaseNo = ""
         Do While Not .EOF
            'Added by Morgan 2012/9/12
   '         If (.Fields("cp01") = "P" Or .Fields("cp01") = "CFP") Then
   '            If InStr(m_strCaseNo, .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")) = 0 Then
   '               If ChkPatentNameCompany(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04")) = True Then
   '                  m_bolChkCompany = True
   '                  If m_strCaseNo <> "" Then m_strCaseNo = m_strCaseNo & ","
   '                  m_strCaseNo = m_strCaseNo & .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")
   '               End If
   '            End If
   '         End If
            'Modify By Sindy 2013/12/26
            If (.Fields("cp01") = "P" Or .Fields("cp01") = "CFP") Or _
               strSrvDate(1) >= InvoiceStartDate Then
               If InStr(m_strCaseNo, .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")) = 0 Then
                  strSpecCompany = ChkPatentNameCompany(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"))
                  If strSpecCompany <> "" And (strSpecCompany = m_strChkCompany Or m_strChkCompany = "") Then
                     m_strChkCompany = strSpecCompany
                     If m_strCaseNo <> "" Then m_strCaseNo = m_strCaseNo & ","
                     m_strCaseNo = m_strCaseNo & .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")
                  End If
               End If
            End If
            '2013/12/26 END
            
            m_bolSplitMail = m_bolSplitMail Or IsSplitReceipt("" & .Fields("cp10"), "" & .Fields("a0j02"), "" & .Fields("cp09"))
         
            If Not IsNull(.Fields("cp140")) Then
               SetAutoTitle .Fields("cp140")
            End If
            
            'Add By Sindy 2023/9/6 ACS代收代付,不能鎖定收據暫不列印
            If .Fields("cp01").Value = "ACS" And .Fields("cp10").Value = "706" And Me.Frame1.Enabled = False Then
               Check1.Value = 0
               Check1.Visible = False
               Check1.Enabled = True
               Me.Frame1.Enabled = True
            End If
            '2023/9/6 END
            
            If .Fields("a0j08") = "Y" Then
               m_bolManualReceipt = True
            End If
            .MoveNext
         Loop
         .MoveFirst
      End With
      
      'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
      'txtMain(6) = GetRecDay
      'txtMain(6).Tag = txtMain(6) '紀錄預設預定收款日
      'end 2018/08/22
      SetRestValue
      '2014/3/10 add by sonia
      'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'      If txtMain(6) <> "" Then
'         txtMain(6).Locked = True
'         txtMain(6).Enabled = False
'      Else
'         txtMain(6).Locked = False
'         txtMain(6).Enabled = True
'      End If
'      '2014/3/10 end
      'end 2018/08/22
   End If
   GridHead1
End Sub

'預設公司別,參照Frmacc1121.OpenTable
Private Sub SetCompany(pSystem As String, pNation As String, pPropertyCode As String, pPromoterNo As String, pCaseNo2 As String, pCaseNo3 As String, pCaseNo4 As String)
   Dim bolChkA0K11 As Boolean
   Dim str000 As String
   
   Select Case pSystem
      Case "T", "CFT", "TC", "CFC", "S", "TD", "TF", "TM", "TR", "TS", "TT"
         Text1 = "1"
         'TT之文件簽證711設定為2公司
         If pSystem = "TT" And pPropertyCode = "711" Then
            Text1 = "2"
         End If
               
      Case "P"
         If pNation <> "000" Then
            Text1 = "1"
            '2012/4/24 ADD BY SONIA 此日期起非台灣P案改用專利法律2公司
            If Val(pCaseNo2) >= 101672 Then
               Text1 = "2"
            End If
            '2012/4/24 END
         Else
            Text1 = "2"
         End If
         
      Case "L"
         '蔣律師為5公司外其餘都用 2公司--辜
         If pPromoterNo = "79037" Then
            Text1 = "5"
         Else
            Text1 = "2"
         End If
               
      Case "LA"
         Text1 = "2"
         
      Case "TB"
         Text1 = "1"
            
      Case "CFP"
         If Val(pCaseNo2) >= 11051 Then
            Select Case pNation
               Case "221", "011", "239"
                  Text1 = "1"
                  If pNation = "239" And Val(pCaseNo2) > 16183 Then
                      Text1 = "2"
                  End If
                           
                  If pNation = "221" And Val(pCaseNo2) > 16183 Then
                      Text1 = "2"
                  End If
                           
                  If pNation = "011" And Val(pCaseNo2) > 23914 Then
                      Text1 = "2"
                  End If

               Case Else
                  Text1 = "2"
            End Select
         Else
            Text1 = "2"
         End If
               
      Case Else
         Text1 = "2"
   End Select
   'Add by Amy2020/03/27 SK02為法務案、顧問案,公司設為L且不可改
   If strSrvDate(1) >= 智慧所更名日 Then
        If ChkAccReceiptComp(2, pSystem) = True Then
            Text1 = "L"
            Text1.Enabled = False
        End If
   End If
   If Text1 = "1" And strSrvDate(1) >= 事務所合併日 Then
        Text1 = "2"
   End If
   'end 2020/03/27
   
   strSql = "select A0K11 from ACC0J0,ACC0K0 where A0J02='" & pSystem & pCaseNo2 & pCaseNo3 & pCaseNo4 & "' and A0J13=A0K01 Order By A0K02 Desc "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   bolChkA0K11 = False
   If intI = 1 Then
      If Trim("" & RsTemp("A0K11")) <> "" Then
         bolChkA0K11 = True
         If Trim("" & RsTemp("A0K11")) <> Trim(Text1) Then
            MsgBox "此案號最新收據的公司別" & Trim("" & RsTemp("A0K11")) & "與系統預設" & Text1 & "不同, 請注意！"
         End If
      End If
   End If
   If bolChkA0K11 = False Then
      If pSystem = "P" Or pSystem = "T" Then
         If pNation = "000" Then
            str000 = "and a0j04='" & pNation & "' "
         Else
            str000 = "and a0j04<>'000' "
         End If
      ElseIf pSystem = "CFP" Then
         If pNation = "011" Then
            str000 = "and a0j04='" & pNation & "' "
         Else
            str000 = "and a0j04<>'011' "
         End If
      Else
         str000 = ""
      End If
      strSql = "select A0K11 from ACC0J0,ACC0K0 where A0J11='" & txtCustNo & "' AND SUBSTR(A0J02, 1, Length(a0j02) - 9)='" & pSystem & "' " & str000 & "and A0J13=A0K01 Order By A0K02 Desc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Trim("" & RsTemp("A0K11")) <> "" Then
            If Trim("" & RsTemp("A0K11")) <> Trim(Text1) Then
               MsgBox "此客戶最新收據的公司別" & Trim("" & RsTemp("A0K11")) & "與系統預設" & Text1 & "不同, 請注意！"
            End If
         End If
      End If
   End If
   
End Sub

Private Sub SetRestValue()
   Dim lngService As Long, lngFee As Long
   With MSHFlexGrid1
   For intI = 1 To .Rows - 1
      lngService = lngService + Val(.TextMatrix(intI, 2))
      lngFee = lngFee + Val(.TextMatrix(intI, 3))
   Next
   End With
   With MSHFlexGrid2
   For intI = 1 To .Rows - 1
      lngService = lngService - Val(.TextMatrix(intI, 9))
      lngFee = lngFee - Val(.TextMatrix(intI, 10))
   Next
   End With
   lblService = Format(lngService, "#,##0")
   lblFee = Format(lngFee, "#,##0")
End Sub

Private Sub GridHead1()
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 1200: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1050: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 800: .Text = "服務費"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(2) = flexAlignRightCenter
      .col = 3: .ColWidth(3) = 800: .Text = "規費"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignRightCenter
      .col = 4: .ColWidth(4) = 400: .Text = "合併"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(4) = flexAlignCenterCenter
      For intI = 5 To .Cols - 1
         .ColWidth(intI) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub GridHead2()
   With MSHFlexGrid2
      .Visible = False
      .Cols = 25 'Modify by Amy 2025/02/20 原:23
      .row = 0
      .col = 0: .ColWidth(0) = 300: .Text = "序"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(0) = flexAlignCenterCenter
      
      .col = 1: .ColWidth(1) = 300: .Text = "次"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignCenterCenter
      
      .col = 2: .ColWidth(2) = 1080: .Text = "收據抬頭"
      .CellAlignment = flexAlignCenterCenter
      
      .col = 3: .ColWidth(3) = 480: .Text = "個/公"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignCenterCenter
      
      .col = 4: .ColWidth(4) = 450: .Text = "合併"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(4) = flexAlignCenterCenter
      
      'Modified by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
      '.col = 5: .ColWidth(5) = 1000: .Text = "預定收款日"
      .col = 5: .ColWidth(5) = 0: .Text = "預定收款日"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(5) = flexAlignLeftCenter
      
      .col = 6: .ColWidth(6) = 850: .Text = "暫不列印"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(6) = flexAlignCenterCenter
      
      .col = 7: .ColWidth(7) = 1200: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      
      .col = 8: .ColWidth(8) = 1000: .Text = "帳款類別"
      .CellAlignment = flexAlignCenterCenter
      
      .col = 9: .ColWidth(9) = 850: .Text = "服務費"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(9) = flexAlignRightCenter
      
      .col = 10: .ColWidth(10) = 850: .Text = "規費"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(10) = flexAlignRightCenter
      
      .col = 11: .ColWidth(11) = 800: .Text = "不印案號"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(11) = flexAlignCenterCenter
      
      .col = 12: .ColWidth(12) = 800: .Text = "不印國家"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(12) = flexAlignCenterCenter
      
      .col = 13: .ColWidth(13) = 1200: .Text = "不印案件名稱"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(13) = flexAlignCenterCenter
      
      .col = 14: .ColWidth(14) = 1200: .Text = "不印商品類別"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(14) = flexAlignCenterCenter
      
      .col = 15: .ColWidth(15) = 1600: .Text = "備註"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(15) = flexAlignLeftCenter
      
      'Modified by Lydia 2016/09/05 寬度1000->1100
      .col = 16: .ColWidth(16) = 1100: .Text = "收據編號"
      .CellAlignment = flexAlignCenterCenter
      
      .col = 17: .ColWidth(17) = 1200: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      
      .col = 18: .ColWidth(18) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      
      .col = 19: .ColWidth(19) = 900: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(19) = flexAlignCenterCenter
      
      'Add By Sindy 2012/11/15
      .col = 20: .ColWidth(20) = 900: .Text = "收據自動列印時間點"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(20) = flexAlignCenterCenter
      '2012/11/15 End
      
      'Add By Sindy 2013/12/26
      .col = 21: .ColWidth(21) = 900: .Text = "介紹案源同仁"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(21) = flexAlignCenterCenter
      '2013/12/26 End
      
      'Add By Sindy 2017/3/17
      .col = 22: .ColWidth(22) = 900: .Text = "列印統一編號"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(22) = flexAlignCenterCenter
      '2017/3/17 End
      
      'Add by Amy 2025/02/20 cu11=統一編號(ColWidth>0為測式用)
      .col = 23: .ColWidth(23) = 100: .Text = "cu11"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(23) = flexAlignCenterCenter
      'cu168=每月代填公司別cu168/a4220(ColWidth>0為測式用)
      .col = 24: .ColWidth(24) = 100: .Text = "cu168"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(24) = flexAlignCenterCenter
      'end 2025/02/20
      
      'For intI = 18 To .Cols - 1
      'For intI = 19 To .Cols - 1
      For intI = 23 To .Cols - 1
         .ColWidth(intI) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub ResetForm()
   Dim oControl As Control
   For Each oControl In Me.Controls
      If TypeName(oControl) = "TextBox" Then
         oControl.Text = ""
'      ElseIf TypeName(oControl) = "ComboBox" Then
'         oControl.Clear
      ElseIf TypeName(oControl) = "Label" Then
         If oControl.Name <> "Label1" Then
            oControl.Caption = ""
         End If
      End If
   Next
   txtMain(1) = "2" '預設公司
   
   '拆收據手開收據預設不可輸
   'Removed by Morgan 2012/11/23 開放可輸入手開收據
   'If m_OldNo <> "" Then txtMain(3).Locked = True
End Sub

'取得此客戶所開的收據抬頭, 並預設最近一次開的收據抬頭, 若無則預設申請人
Private Function SetReceiptTitle(Optional strA0K04 As String) As String
Dim strMaxA0k02 As String 'Add By Sindy 2013/12/26
   
   cboTitle.Clear
   If txtCustNo <> "" Then
      cboTitle.AddItem CustomerQuery(txtCustNo, 1)
      txtCustName = cboTitle.List(0)
      strExc(0) = "Select A0K04,max(a0k02) lstDate From ACC0K0 Where A0K03='" & txtCustNo & "' and (a0k09 is null or a0k09=0)" & _
         " and a0k04 is not null group by a0k04"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         While Not RsTemp.EOF
            If RsTemp.Fields(0) <> cboTitle.List(0) Then cboTitle.AddItem RsTemp.Fields(0)
            RsTemp.MoveNext
         Wend
         RsTemp.Sort = " lstDate desc"
         If strA0K04 <> "" Then
            cboTitle = strA0K04
         Else
            cboTitle = RsTemp.Fields(0)
         End If
      Else
         cboTitle.ListIndex = 0
      End If
      
      'Add By Sindy 2013/12/26 若客戶在上次收據日期後有更名,要提醒操作者
      strExc(0) = "SELECT nvl(max(a0k02),0) FROM acc0k0 WHERE a0k03='" & txtCustNo & "' and nvl(a0k09,0)=0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) > 0 Then
            strMaxA0k02 = RsTemp.Fields(0)
            strExc(0) = "SELECT cu04 FROM customer WHERE cu01='" & Left(txtCustNo, 8) & "' and cu02='0' and cu82>" & DBDATE(strMaxA0k02)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "請注意,客戶已更名為" & "" & RsTemp.Fields("CU04") & ",請注意欲開立之收據抬頭!!", vbInformation, "收據公司別提醒"
            End If
         End If
      End If
      '2013/12/26 END
   End If
End Function

Private Sub SetColor(pRow As Integer, pColor As Long)
   With MSHFlexGrid2
   .row = pRow
   For intI = 0 To .Cols - 1
      .col = intI
      .CellBackColor = pColor
   Next
   End With
End Sub

Private Sub SetData()
   With MSHFlexGrid2
      txtMain(0) = .TextMatrix(.row, 0)
      cboTitle = .TextMatrix(.row, 2)
      txtMain(1) = .TextMatrix(.row, 3)
      txtMain(3) = .TextMatrix(.row, 16)
      txtMain(5) = .TextMatrix(.row, 15)
      txtMain(6) = Replace(.TextMatrix(.row, 5), "/", "")
      '暫不列印
      If .TextMatrix(.row, 6) = "N" Then
         Check1.Value = 1
      Else
         Check1.Value = 0
      End If
      'Add By Sindy 2012/11/15
      Me.Check3(0).Value = 0
      Me.Check3(1).Value = 0
      Me.Check3(2).Value = 0
      If .TextMatrix(.row, 20) <> "" Then
         If .TextMatrix(.row, 20) = "1" Then Me.Check3(0).Value = 1
         If .TextMatrix(.row, 20) = "2" Then Me.Check3(1).Value = 1
         If .TextMatrix(.row, 20) = "3" Then Me.Check3(2).Value = 1
      End If
      '2012/11/15 End
'      'Add By Sindy 2013/12/26
'      txtSales = .TextMatrix(.row, 21)
'      lblSales.Caption = GetPrjSalesNM(.TextMatrix(.row, 21))
'      '2013/12/26 END
      'Modify By Sindy 2014/12/29
      Combo2 = .TextMatrix(.row, 21)
      Call Combo2_LostFocus
      '2014/12/29 END
      'Add By Sindy 2017/3/17
      '列印統一編號
      If .TextMatrix(.row, 22) = "" Then
         txtPrintNo.Text = ""
      Else
         txtPrintNo.Text = "" & .TextMatrix(.row, 22)
      End If
      '2017/3/17 END
   End With
End Sub

Private Sub MSHFlexGrid2_Click()
   
   Dim iCurCol As Integer, iCurRow As Integer
      
   With MSHFlexGrid2
   If .MouseRow > 0 And .MouseRow < .Rows And .MouseCol < 19 Then
      iCurRow = .MouseRow
      iCurCol = .MouseCol
      .Visible = False
      '還原上一筆點選
      If m_lstRow > 0 Then
         SetColor m_lstRow, m_lstColor
      End If
      .row = iCurRow
      m_lstColor = .CellBackColor
      SetColor iCurRow, m_dftColor3
      
      m_lstRow = .row
      
      .col = iCurCol
      iRow = .row: iCol = .col
      If .col = 1 Or .col = 3 Or .col = 8 Or .col = 9 Or .col = 10 Then SetBox
           
      .Visible = True
      SetData
      
   End If
   End With
End Sub

Private Sub MSHFlexGrid2_DblClick()
   If iCol = 11 Or iCol = 12 Or iCol = 13 Or iCol = 14 Then
      txtInput.Tag = MSHFlexGrid2.TextMatrix(iRow, iCol)
      txtInput = txtInput.Tag
      If txtInput = "N" Then
         txtInput = ""
      Else
         txtInput = "N"
      End If
      UpdateCol
   End If
End Sub

Private Sub MSHFlexGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   txtInput.Visible = False
End Sub

Private Sub MSHFlexGrid2_Scroll()
   If txtInput.Visible = True Then
      SetBox False
   End If
End Sub

Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   'ADD BY SONIA 2014/5/28 J公司不可暫不列印-婷
   'cancel by sonia 2017/12/6婷要求開放
   'If Text1 = "J" Then
   '   Check1.Value = 0
   '   Check1.Enabled = False
   '   'add by sonia 2015/11/26
   '   Check3(0).Value = 0: Check3(1).Value = 0: Check3(2).Value = 0
   '   Check3(0).Enabled = False: Check3(1).Enabled = False: Check3(2).Enabled = False
   '   'end 2015/11/26
   'Else
   'end 2017/12/6
'      Check1.Enabled = True
      'add by sonia 2015/11/26
      Check3(0).Enabled = True: Check3(1).Enabled = True: Check3(2).Enabled = True
      'end 2015/11/26
   'End If  'cancel by sonia 2017/12/6婷要求開放
   'END 2014/5/28
   
   Select Case Text1
      Case "1"
         Text12 = MsgText(901)
      Case "2"
         'Modify by Amy 2020/03/26
         If strSrvDate(1) >= 事務所合併日 Then
            Text12 = A0802Query(Text1, True)
         Else
            Text12 = MsgText(902)
         End If
      Case "3"
         Text12 = MsgText(903)
      Case "5"
         Text12 = MsgText(904)
      Case "7"
         Text12 = MsgText(905)
      Case "8"
         Text12 = MsgText(906)
      'Add By Sindy 2013/12/19
      Case "9"
         Text12 = MsgText(908)
      'Add By Sindy 2013/12/18
      Case "J"
         Text12 = MsgText(907)
      'Add by Amy 2020/03/26
      Case "L"
         Text12 = A0802Query(Text1, True)
   End Select
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

'Add By Sindy 2013/12/26
Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      MsgBox MsgText(188) & Label1(0), , MsgText(5)
      Cancel = True
      Text1.SetFocus
      Exit Sub
   Else
      strExc(0) = "select * from acc080 where a0801 = '" & Text1 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI <> 1 Then
         MsgBox MsgText(188) & Label1(0), , MsgText(5)
         Cancel = True
      End If
      'ADD BY SONIA 2014/5/28 J公司不可暫不列印-婷
        'cancel by sonia 2017/12/6婷要求開放
      'If Text1 = "J" Then
      '   Check1.Value = 0
      '   Check1.Enabled = False
      '   'add by sonia 2015/11/26
      '   Check3(0).Value = 0: Check3(1).Value = 0: Check3(2).Value = 0
      '   Check3(0).Enabled = False: Check3(1).Enabled = False: Check3(2).Enabled = False
      '   'end 2015/11/26
      'Else
      'end 2017/12/6
'         Check1.Enabled = True
         'add by sonia 2015/11/26
         Check3(0).Enabled = True: Check3(1).Enabled = True: Check3(2).Enabled = True
         'end 2015/11/26
      'End If  'cancel by sonia 2017/12/6婷要求開放
      'END 2014/5/28
   End If
End Sub

Private Sub txtDate_GotFocus()
   CloseIme
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
   If txtDate <> "" Then
      If ChkDate(txtDate) = False Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtInput_GotFocus()
   TextInverse txtInput
   CloseIme
End Sub

Private Sub txtInput_LostFocus()
   If txtInput.Locked = False Then UpdateCol
   txtInput.Visible = False
End Sub

Private Sub txtMain_GotFocus(Index As Integer)
   TextInverse txtMain(Index)
   Select Case Index
      Case 4, 5
         OpenIme
      Case Else
         CloseIme
   End Select
End Sub

Private Sub txtMain_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 8 Then Exit Sub
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 1
         If Chr(KeyAscii) <> "1" And Chr(KeyAscii) <> "2" Then
            KeyAscii = 0
         End If
   End Select
End Sub

Private Sub txtMain_LostFocus(Index As Integer)
   CloseIme
End Sub

Private Sub txtMain_Validate(Index As Integer, Cancel As Boolean)
   If Index = 6 Then
      If txtMain(6) <> txtMain(6).Tag Then
         '原來有日期但被清除
         If txtMain(6) = "" Then
            Cancel = True
            MsgBox "請輸入預定收款日！"
         Else
            '檢查格式
            If ChkDate(txtMain(6)) = False Then
               Cancel = True
            ElseIf Val(txtMain(6)) < Val(strSrvDate(2)) Then
               Cancel = True
               MsgBox "預定收款日不可小於系統日！"
            End If
         End If
         If Cancel = True Then
            txtMain_GotFocus 6
         End If
      End If
   End If
End Sub

'抓預定收款日期
Private Function GetRecDay() As String
   '同收文號抓最後異動的日期，多個收文號時抓最小的日期
   strExc(0) = "select rd05 from receivablesday" & _
      " where (rd01,rd02*1000+rd03) in (" & _
      " select rd01,max(rd02*1000+rd03) from acc0j0,receivablesday" & _
      " where a0j06 = '" & MsgText(602) & "' and rd01(+)=a0j01 group by rd01 ) order by 1 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetRecDay = TransDate(RsTemp(0), 1)
   End If
End Function

'設定自動收文的收據抬頭
Private Sub SetAutoTitle(pCRL01 As String)
   Dim stSQL As String, iR As Integer, adoRst As ADODB.Recordset
   
   'Modify By Sindy 2012/11/19 +CRL92
   'Modified by Lydia 2024/08/05 +CRL153
   stSQL = "select crl41,crl42,crl50,CRL92,CRL153 from consultrecordlist" & _
      " where crl01='" & pCRL01 & "'"
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      
      If adoRst("crl50") = "Y" Then Check1.Value = 1
      'Add By Sindy 2012/11/19 接洽單自動收文時要帶出智權人員當時輸入的收據自動列印時間點
      Me.Check3(0).Value = 0
      Me.Check3(1).Value = 0
      Me.Check3(2).Value = 0
      If "" & adoRst.Fields("CRL92") <> "" Then
         If "" & adoRst.Fields("CRL92") = "1" Then Me.Check3(0).Value = 1
         If "" & adoRst.Fields("CRL92") = "2" Then Me.Check3(1).Value = 1
         If "" & adoRst.Fields("CRL92") = "3" Then Me.Check3(2).Value = 1
      'Add By Sindy 2023/9/5
      ElseIf Check1.Value = 1 Then
         Check1.Visible = True
         Check1.Enabled = False
         Me.Frame1.Enabled = False
         '2023/9/5 END
      End If
      '2012/11/19 End
      m_CRL153 = "" & adoRst.Fields("CRL153") 'Added by Lydia 2024/08/05 國內接洽單：DEBIT NOTE請款選項
      If adoRst("crl41") = "1" Then
         cboTitle.Text = txtCustName
      ElseIf adoRst("crl41") = "2" Then
         'Modified by Lydia 2024/08/05
         'MsgBox "本接洽單設定為以 DEBIT NOTE 請款！", vbExclamation, "注意"
         strExc(0) = "本接洽單設定為以 DEBIT NOTE 請款！" & vbCrLf & "DEBIT NOTE請款選項："
         If m_CRL153 = "1" Then
            strExc(0) = strExc(0) & "立即開立DEBIT NOTE"
         ElseIf m_CRL153 <> "" Then
            strExc(0) = strExc(0) & "待通知後開立，" & IIf(m_CRL153 = "2", "要", "不需要") & "加印國內收據"
         End If
         If strShowCRL153 = "" Or (m_CRL153 <> strShowCRL153 And (m_CRL153 = "1" Or m_CRL153 = "3")) Then
            MsgBox strExc(0), vbExclamation, "注意"
            strShowCRL153 = m_CRL153
         End If
         'end 2024/08/05
      ElseIf adoRst("crl41") = "3" Then
         cboTitle.Text = "" & adoRst("crl42")
      Else
         MsgBox "自動收文抬頭資料設定錯誤！"
      End If
   End If
   Set adoRst = Nothing
End Sub

Private Function TxtValidate() As Boolean
   Dim ii As Integer, jj As Integer
   Dim lngSFee As Long, lngOFee As Long
   Dim lstNo As String
   Dim bCancel As Boolean
   Dim stItem As String, stItemList As String, iItemCount As Integer
   Dim stCaseNo As String, stCaseNoList As String, stCaseNoCount As Integer
   Dim strSpecCompany As String 'Add By Sindy 2014/1/13
   Dim strMsg As String 'Add by Amy 2020/03/26
   
   If Text1 = "" Then
      MsgBox "請輸入公司別！"
      Text1.SetFocus
      Exit Function
   'Add By Sindy 2013/12/26
   Else
      'Moidfy by Amy 2020/03/26 會有舊資料故判斷是否Enabled
      If Text1.Enabled = True Then
            If strSrvDate(1) >= 事務所合併日 Then
                If ChkAccReceiptComp(0, Text1, strMsg) = False Then
                    MsgBox strMsg, , MsgText(5)
                    Text1.SetFocus
                    Exit Function
                End If
            Else
                If strSrvDate(1) >= InvoiceStartDate Then
                   If Text1 <> "1" And Text1 <> "2" And Text1 <> "9" And Text1 <> "J" Then
                      MsgBox "收據公司別只可輸入１或２或９或J", , MsgText(5)
                      Text1.SetFocus
                      Exit Function
                   End If
                Else
                   If Text1 <> "1" And Text1 <> "2" And Text1 <> "9" Then
                      MsgBox "收據公司別只可輸入１或２或９", , MsgText(5)
                      Text1.SetFocus
                      Exit Function
                   End If
                End If
            End If
      End If
      'end 2020/03/26
   '2013/12/26 end
   End If
   
   If txtDate = "" Then
      MsgBox "請輸入收據日期!"
      txtDate.SetFocus
      Exit Function
   End If
   
   'Add By Sindy 2014/1/13
   If m_OldNo <> "" Then
      '拆收據
      strExc(0) = "select cp01||cp02||cp03||cp04 本所案號,getcp10desc(cp01,cp10,a0j04) 案件性質,a0j09||'' 服務費,a0j10||'' 規費" & _
                  ",a0j07,a0j01,cp05,cp01,cp02,cp03,cp04,cp12,cp13,K.*,a0j04,cp151,cp31 from acc0k0 K,acc0j0,caseprogress,casepropertymap" & _
                  " where a0k01='" & m_OldNo & "' and a0j13(+)=a0k01 and cp09(+) = a0j01" & _
                  " and cpm01(+)=cp01 and cpm02(+)=cp10 order by 1,cp05,cp09"
   Else
      strExc(0) = "select cp01||cp02||cp03||cp04 本所案號,getcp10desc(cp01,cp10,a0j04) 案件性質,a0j09||'' 服務費,a0j10||'' 規費" & _
                  ",a0j07,cp09,cp05,a0j04,a0j11,cp01,cp02,cp03,cp04,cp10,cp12,cp13,cp14,cp140,a0j08,a0j02,cp151,CP31" & _
                  " from acc0j0,caseprogress,casepropertymap" & _
                  " where cp09(+) = a0j01 and a0j06 = '" & MsgText(602) & "' and a0j13=a0j01" & _
                  " and cpm01(+)=cp01 and cpm02(+)=cp10 order by 1,cp05,cp09"
   End If
   adocheck.CursorLocation = adUseClient
   adocheck.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      With adocheck
         .MoveFirst
         m_strChkCompany = "": m_strCaseNo = ""
         Do While Not .EOF
            If InStr(m_strCaseNo, .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")) = 0 Then
               strSpecCompany = ChkPatentNameCompany(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"))
               If strSpecCompany <> "" And (strSpecCompany = m_strChkCompany Or m_strChkCompany = "") Then
                  m_strChkCompany = strSpecCompany
                  If m_strCaseNo <> "" Then m_strCaseNo = m_strCaseNo & ","
                  m_strCaseNo = m_strCaseNo & .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")
               End If
            End If
            .MoveNext
         Loop
      End With
   End If
   adocheck.Close
   '2014/1/13 END
   
   'Added by Morgan 2012/9/12
   'Add By Sindy 2013/12/17
   If strSrvDate(1) >= InvoiceStartDate Then
      If m_strChkCompany = "T" And Text1 <> "1" And m_CP31 = "Y" Then
         MsgBox "專利案" & m_strCaseNo & "有設定以專利商標出名不可開立其他公司別，請與專業部確認!!", vbCritical, "收據公司別提醒"
         Text1.SetFocus
         Exit Function
      ElseIf m_strChkCompany = "J" And Text1 <> "J" And m_CP31 = "Y" Then
         MsgBox m_strCaseNo & "有設定以智權公司出名不可開立其他公司別，請與專業部確認!!", vbCritical, "收據公司別提醒"
         Text1.SetFocus
         Exit Function
      End If
   Else
   '2013/12/17 END
      If m_strChkCompany <> "" And Text1 <> "1" And m_CP31 = "Y" Then
         MsgBox "專利案" & m_strCaseNo & "有設定以專利商標出名不可開立其他公司別，請與專業部確認!!", vbCritical, "收據公司別提醒"
         Text1.SetFocus
         Exit Function
      End If
   End If
   
   If Val(lblService) <> 0 Then
      MsgBox "待開服務費必須為 0！"
      Exit Function
   End If
   If Val(lblFee) <> 0 Then
      'Added by Morgan 2025/8/19
      If Text1 = "J" And Left(MSHFlexGrid2.TextMatrix(1, 7), 3) = "ACS" Then
         MsgBox "待開規費必須為 0！" & vbCrLf & vbCrLf & "若拆收據後規費(稅)合計與原來不同時請調整收文金額。", vbExclamation, "ACS智權公司拆收據提醒"
      Else
      'end 2025/8/19
      
         MsgBox "待開規費必須為 0！"
      End If
      Exit Function
   End If
   
   bCancel = False
   With MSHFlexGrid2
   lstNo = .TextMatrix(1, 0)
   iItemCount = 0
   stItemList = vbTab
   stCaseNoCount = 0
   stCaseNoList = vbTab
   For jj = 1 To .Rows - 1
      If .TextMatrix(jj, 8) = "" Then
         MsgBox "帳款類別不可空白!!"
         bCancel = True
         Exit For
      End If
      
'Remove by Morgan 2012/3/20
'      If Val(.TextMatrix(jj, 9)) = 0 And Val(.TextMatrix(jj, 10)) = 0 Then
'         MsgBox "服務費規費不可皆為 0 !!"
'         bCancel = True
'         Exit For
'      End If
      
      If lstNo = .TextMatrix(jj, 0) Then
      
         stItem = .TextMatrix(jj, 8)
         If InStr(stItemList, vbTab & stItem & vbTab) = 0 Then
            iItemCount = iItemCount + 1
            'Modify By Sindy 2013/12/26
            If Text1 = "J" Then
               If iItemCount > 5 Then
                  MsgBox "一張收據不可超過5個不同的帳款類別!!"
                  bCancel = True
                  Exit For
               End If
            Else
            '2013/12/26 END
               'Modified by Morgan 2022/11/10 改可印列數改用變數控制(與列印同步)
               If iItemCount > m_intRecMaxItem Then
                  MsgBox "一張收據不可超過" & m_intRecMaxItem & "個不同的帳款類別!!"
                  bCancel = True
                  Exit For
               End If
            End If
            stItemList = stItemList & stItem & vbTab
         End If
         
         If Text1 <> "J" Then 'Modify By Sindy 2013/12/26 +if
            stCaseNo = .TextMatrix(jj, 7)
            If InStr(stCaseNoList, vbTab & stCaseNo & vbTab) = 0 Then
               If .TextMatrix(jj, 11) = "" Or .TextMatrix(jj, 12) = "" Then
                  stCaseNoCount = stCaseNoCount + 1
                  If stCaseNoCount > 2 Then
                     MsgBox "一張收據不可超過2組要印的案號國家或案件名稱!!"
                     bCancel = True
                     Exit For
                  End If
                  stCaseNoList = stCaseNoList & stCaseNo & vbTab
               End If
            End If
         End If
      Else
         lstNo = .TextMatrix(jj, 0)
         iItemCount = 0
         stItemList = vbTab
         stCaseNoCount = 0
         stCaseNoList = vbTab
      End If
      
      'Added by Morgan 2025/7/7
      '智權公司ACS案件在開(拆)收據時,不能沒有輸規費(稅)--瑞婷
      If Text1 = "J" And Left(.TextMatrix(jj, 7), 3) = "ACS" Then
         If Val(.TextMatrix(jj, 10)) = 0 Then
            MsgBox "智權公司ACS案件在開(拆)收據時,不能沒有輸規費(稅)!!", vbCritical
            bCancel = True
            Exit For
         Else
            intI = Round(Val(.TextMatrix(jj, 9)) * 0.05)
            'ACS一個收文號一張收據
            If Val(.TextMatrix(jj, 10)) <> intI Then
               'Modified by Morgan 2025/8/19
               'If MsgBox("規費(稅)錯誤!!應該為" & intI & " (" & Val(.TextMatrix(jj, 9)) & "x0.05)。" & vbCrLf & vbCrLf & "是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
               '   bCancel = True
               '   Exit For
               'End If
               MsgBox "規費(稅)錯誤!!" & vbCrLf & vbCrLf & "正確金額:四捨五入(" & Val(.TextMatrix(jj, 9)) & "x0.05)=" & intI, vbCritical
               bCancel = True
               Exit For
            End If
         End If
      End If
      'end 2025/7/7
   Next
   
   If bCancel = True Then
      .Visible = False
      '還原上一筆點選
      If m_lstRow > 0 Then
         SetColor m_lstRow, m_lstColor
      End If
      .row = jj
      m_lstColor = .CellBackColor
      SetColor jj, m_dftColor3
      m_lstRow = .row
      .Visible = True
      SetData
      .TopRow = jj
      Exit Function
   End If
      
   End With
   
   For ii = 1 To MSHFlexGrid1.Rows - 1
      lngSFee = 0
      lngOFee = 0
      For jj = 1 To MSHFlexGrid2.Rows - 1
         If MSHFlexGrid2.TextMatrix(jj, 18) = MSHFlexGrid1.TextMatrix(ii, 5) Then
            lngSFee = lngSFee + Val(MSHFlexGrid2.TextMatrix(jj, 9))
            lngOFee = lngOFee + Val(MSHFlexGrid2.TextMatrix(jj, 10))
         End If
      Next
      If lngSFee <> Val(MSHFlexGrid1.TextMatrix(ii, 2)) Then
         MsgBox "[" & MSHFlexGrid1.TextMatrix(ii, 0) & " " & MSHFlexGrid1.TextMatrix(ii, 1) & "] 服務費加總與原金額不符!!'"
         Exit Function
      End If
      If lngOFee <> Val(MSHFlexGrid1.TextMatrix(ii, 3)) Then
         MsgBox "[" & MSHFlexGrid1.TextMatrix(ii, 0) & " " & MSHFlexGrid1.TextMatrix(ii, 1) & "] 規費加總與原金額不符!!'"
         Exit Function
      End If
   Next
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
   Dim ii As Integer
   Dim strLstNo As String
   Dim A0K(40) As String
   Dim bolInsert As Boolean
   Dim strCP09 As String
   Dim bolInsDate As Boolean '新增預定收款日紀錄
   Dim strCP09List As String
   
On Error GoTo ErrHnd

   adoTaie.BeginTrans
   
   If Check2.Value = vbUnchecked Then
      SortData
   End If
   
   A0K(2) = txtDate
   A0K(3) = txtCustNo
   A0K(11) = Text1
   A0K(20) = m_CP13
   'modify by sonia 2017/6/22 若屬於業績列入P1001之專利處人員則智權人員改為P1001
   'A0K(22) = m_CP12
   If adoquery.State = 1 Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * FROM SetSpecMan where ocode='P1001' and instr(oman,'" & A0K(20) & "')>0 ", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      A0K(20) = "P1001"
   End If
   adoquery.Close
   A0K(22) = PUB_GetStaffST15(A0K(20), 1)
   'end 2017/6/22
   
   With MSHFlexGrid2
   strLstNo = .TextMatrix(1, 0)
   
   If .TextMatrix(1, 16) = "" Then
      .TextMatrix(1, 16) = AutoNo(MsgText(802), 5, 1) '收據編號
   End If
   
   For ii = 1 To .Rows - 1
      A0K(1) = .TextMatrix(ii, 16)
      'Modified by Morgan 2011/12/26 取消 a0j03,a0j05,a0j12,a0j20,a0j21
      'Modified by Morgan 2017/11/16 +a0j21
      strSql = "insert into acc0j0 (a0j01,a0j02,a0j04,a0j06,a0j07,a0j08,a0j09,a0j10" & _
         ",a0j11,a0j13,a0j14,a0j15,a0j16,a0j22,a0j21,a0j23,a0j24,a0j20,a0j25) select a0j01,a0j02,a0j04,a0j06,a0j07,a0j08," & Val(.TextMatrix(ii, 9)) & "," & Val(.TextMatrix(ii, 10)) & _
         ",a0j11,'" & A0K(1) & "',a0j14,a0j15,a0j16,'" & ChgSQL(.TextMatrix(ii, 8)) & "','" & .TextMatrix(ii, 11) & "','" & .TextMatrix(ii, 12) & "','" & .TextMatrix(ii, 13) & "','" & .TextMatrix(ii, 14) & "'," & Val(.TextMatrix(ii, 1)) & _
         " from acc0j0 where a0j01='" & .TextMatrix(ii, 18) & "' and a0j13=a0j01"
         
      adoTaie.Execute strSql, intI
      
      strSql = "insert into acc1m0 values('" & .TextMatrix(ii, 16) & "','" & .TextMatrix(ii, 18) & "')"
      adoTaie.Execute strSql, intI
      
      A0K(6) = Val(A0K(6)) + Val(.TextMatrix(ii, 9)) '服務費
      A0K(7) = Val(A0K(7)) + Val(.TextMatrix(ii, 10)) '規費
      '最後一筆也要新增最後一張收據
      If ii = .Rows - 1 Then
         bolInsert = True
      '下一筆序號不同新增收據
      ElseIf .TextMatrix(ii + 1, 0) <> .TextMatrix(ii, 0) Then
         If .TextMatrix(ii + 1, 16) = "" Then
            .TextMatrix(ii + 1, 16) = AutoNo(MsgText(802), 5, 1)  '收據編號
         End If
         bolInsert = True
      ElseIf .TextMatrix(ii + 1, 0) = .TextMatrix(ii, 0) Then
         .TextMatrix(ii + 1, 16) = .TextMatrix(ii, 16)
      End If
      
      If bolInsert = True Then
         A0K(4) = .TextMatrix(ii, 2) '收據抬頭
         A0K(5) = .TextMatrix(ii, 3) '個人/公司
         A0K(8) = .TextMatrix(ii, 15) '備註
         A0K(32) = .TextMatrix(ii, 6) '收據暫不列印
         A0K(34) = .TextMatrix(ii, 21) '介紹案源同仁 Add By Sindy 2013/12/26
         'Added by Lydia 2024/08/05 國內接洽單：DEBIT NOTE請款選項
         If strShowCRL153 = "1" Or strShowCRL153 = "3" Then  '1=立即開立DEBIT NOTE, 3=待通知後開立，不需要加印國內收據
            A0K(32) = "Z" 'Z.確定不印(為了和暫不列印做取代)
         End If
         'end 2024/08/05
         
         strExc(0) = "select * from acc0k0 where a0k01='" & A0K(1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         'Modify By Sindy 2013/12/26 增加更新a0k34
         If intI = 1 Then
            strSql = "update acc0k0 set a0k02='" & A0K(2) & "', a0k03='" & A0K(3) & "', a0k04='" & ChgSQL(A0K(4)) & "'" & _
               ", a0k05='" & A0K(5) & "', a0k06=" & A0K(6) & ", a0k07=" & A0K(7) & ", a0k08='" & ChgSQL(A0K(8)) & "'" & _
               ", a0k09=0, a0k11='" & A0K(11) & "', a0k12='" & A0K(12) & "', a0k17=0, a0k18=0, a0k19=1" & _
               ", a0k20='" & A0K(20) & "',a0k22='" & A0K(22) & "', a0k27=" & strSrvDate(2) & _
               ", a0k28=to_char(sysdate,'HH24MiSS'), a0k29='" & strUserNum & "'" & _
               ", a0k32='" & A0K(32) & "', a0k33='Y', a0k34=" & CNULL(A0K(34)) & _
               " where a0k01='" & A0K(1) & "'"
         Else
            'Modified by Lydia 2016/09/06 +a0k23
            strSql = "insert into acc0k0 (a0k01, a0k02, a0k03, a0k04, a0k05, a0k06, a0k07" & _
               ", a0k08, a0k09, a0k11, a0k12, a0k17, a0k18, a0k19, a0k20,a0k22,a0k23, a0k24, a0k25, a0k26, a0k32, a0k33, a0k34) values " & _
               "('" & A0K(1) & "', " & A0K(2) & ", '" & A0K(3) & "', '" & ChgSQL(A0K(4)) & "', '" & A0K(5) & "', " & A0K(6) & ", " & A0K(7) & _
               ", '" & ChgSQL(A0K(8)) & "', 0, '" & A0K(11) & "', '" & A0K(12) & "', 0, 0, 0, '" & A0K(20) & "', '" & A0K(22) & "'," & CNULL(m_Nation) & ", " & strSrvDate(2) & ", to_char(sysdate,'HH24MiSS'), '" & strUserNum & "','" & A0K(32) & "','Y'," & CNULL(A0K(34)) & ")"
         End If
         adoTaie.Execute strSql, intI
         
         A0K(6) = 0
         A0K(7) = 0
         A0K(32) = ""
         bolInsert = False
      End If
      strLstNo = .TextMatrix(ii, 0)
      
      'Modify By Sindy 2012/11/15 +cp151=" & CNULL(.TextMatrix(ii, 20)) & ",
      strSql = "update caseprogress set cp151=" & CNULL(.TextMatrix(ii, 20)) & ",cp60='" & A0K(1) & "', cp73 = 0, cp74 = 0, cp75 = 0, cp76 = 0, cp77 = 0, cp78 = 0, cp79 = cp16 where cp09='" & .TextMatrix(ii, 18) & "' and cp60 is null"
      adoTaie.Execute strSql, intI
      
      
      'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'      If .TextMatrix(ii, 5) <> "" Then
'         If InStr(strCP09List, .TextMatrix(ii, 18)) = 0 Then '一個收文號一次
'            strCP09List = strCP09List & "," & .TextMatrix(ii, 18)
'            strExc(1) = DBDATE(.TextMatrix(ii, 5))
'            bolInsDate = False
'            strExc(0) = "select rd02*1000+rd03||rd05 from receivablesday where rd01='" & .TextMatrix(ii, 18) & "' order by 1 desc"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               '預定收款日期不同
'               If Mid(RsTemp(0), 12) <> strExc(1) Then
'                  bolInsDate = True
'               End If
'            Else
'               bolInsDate = True
'            End If
'            If bolInsDate = True Then
'               strSql = "insert into receivablesday (rd01,rd02,rd03,rd04,rd05)" & _
'                  " select '" & .TextMatrix(ii, 18) & "'," & strSrvDate(1) & ",nvl(max(rd03),0)+1,'" & strUserNum & "'," & strExc(1) & _
'                  " from receivablesday where rd01='" & .TextMatrix(ii, 18) & "' and rd02=" & strSrvDate(1)
'               adoTaie.Execute strSql, intI
'            End If
'         End If
'      End If
      'end 2018/08/22
   Next
   End With
   
   strSql = "delete acc0j0 where a0j06 = '" & MsgText(602) & "' and a0j16='" & strUserNum & "' and a0j13=a0j01"
   adoTaie.Execute strSql, intI
      
   Call ChkACStoEmail(m_CP01, m_CP09, strPropertyCode) 'Adderd by Lydia 2025/11/13
   
   adoTaie.CommitTrans
   FormSave = True
   
   'Add By Sindy 2023/10/19 ACS不管制
   'Modify By Sindy 2023/10/25 J公司才檢查
   If m_CP01 <> "ACS" And Text1 = "J" Then
   '2023/10/19 END
      'Modify By Sindy 2024/9/25 增加傳入公司別做判斷
      Call PUB_ChkCU144isN(Mid(txtCustNo, 1, 8), Mid(txtCustNo, 9, 1), "", Text1, , , "A") 'Add By Sindy 2023/9/4
   End If
   
   Exit Function
   
ErrHnd:
   adoTaie.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

Private Sub SetBox(Optional pbolSetValue As Boolean = True)
   
   Dim lngLeft As Long, lngTop As Long
   Dim ii As Integer
   
   With MSHFlexGrid2
   If .LeftCol > .col Or .TopRow > .row Then
      txtInput.Visible = False
   Else
      txtInput.FontName = .CellFontName
      txtInput.FontSize = .CellFontSize
      If .CellAlignment < 3 Then
         txtInput.Alignment = 0 '靠左
      ElseIf .CellAlignment < 6 Then
         txtInput.Alignment = 2 '置中
      ElseIf .CellAlignment < 9 Then
         txtInput.Alignment = 1 '靠右
      Else
         txtInput.Alignment = 0 '靠左
      End If
      If pbolSetValue = True Then
         txtInput.Text = .TextMatrix(.row, .col)
      End If
      txtInput.Tag = txtInput.Text
      txtInput.Width = .ColWidth(.col) + 10
      txtInput.Height = .RowHeight(.row) - 5
      'iRow = .row: iCol = .col
      lngLeft = .Left + 20
      lngTop = .Top + .RowHeight(0) + 20
      For ii = .LeftCol To .col - 1
         lngLeft = lngLeft + .ColWidth(ii)
      Next
      For ii = .TopRow To .row - 1
         lngTop = lngTop + .RowHeight(ii)
      Next
      txtInput.Left = lngLeft: txtInput.Top = lngTop
      If txtInput.Left + txtInput.Width < .Left + .Width Then
         txtInput.Visible = True
         txtInput.SetFocus
         TextInverse txtInput
         iRow = .row: iCol = .col
      Else
         txtInput.Visible = False
      End If
   End If
   End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   If KeyAscii = 8 Then Exit Sub
   
   If KeyAscii = vbKeyReturn Then
      UpdateCol
      GoNext
   ElseIf KeyAscii = vbKeyEscape Then
      txtInput = txtInput.Tag
      TextInverse txtInput
   '個人/公司
   ElseIf iCol = 3 Then
      If Chr(KeyAscii) <> "1" And Chr(KeyAscii) <> "2" Then
         KeyAscii = 0
         Beep
         Exit Sub
      End If
   '服務費規費欄位
   ElseIf iCol = 9 Or iCol = 10 Then
      If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) Then
         KeyAscii = 0
         Beep
         Exit Sub
      End If
   '是否不印案號,是否不印國家,是否不印案件名稱,是否不印商品類別
   ElseIf iCol = 11 Or iCol = 12 Or iCol = 13 Or iCol = 14 Then
      KeyAscii = UpperCase(KeyAscii)
      If Not (KeyAscii = Asc("N")) Then
         KeyAscii = 0
         Beep
         Exit Sub
      End If
      
   '排序欄位
   ElseIf iCol = 1 Then
      If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) Then
         KeyAscii = 0
         Beep
         Exit Sub
      Else
         Check2.Value = vbChecked
      End If
   End If
End Sub

Private Sub GoNext()
   With MSHFlexGrid2
      If .col = 8 Then
         .col = 9
      ElseIf .col = 9 Then
         .col = 10
      Else
         .col = 8
         If .row < .Rows - 1 Then
            .row = .row + 1
         Else
            .row = 1
         End If
      End If
      SetBox
   End With
End Sub

Private Function InsertCheck(Optional pUpdate As Boolean = False) As Boolean
Dim bCancel As Boolean 'Add By Sindy 2012/11/15
   
   If cboTitle = "" Then
      MsgBox "請輸入收據抬頭！"
      cboTitle.SetFocus
      Exit Function
   'Add By Sindy 2015/11/16 收據抬頭欄若不存在於客戶檔及抬頭檔,增加顯示訊息提醒
   Else
'      'Modify By Sindy 2016/8/22 + or cu05||' '||cu88||' '||cu89||' '||cu90='" & cboTitle & "' or cu06='" & cboTitle & "'
'      strSql = "select cu11" & _
'               " From customer" & _
'               " where (upper(cu04)=upper('" & ChgSQL(cboTitle) & "') or upper(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))=upper('" & ChgSQL(cboTitle) & "') or upper(cu06)=upper('" & ChgSQL(cboTitle) & "'))" & _
'               " and cu15<>'0'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 0 Then
'         'Modify By Sindy 2017/4/18 and A4202<>'04150022'==>and (A4202<>'04150022' or A4202 is null) 改語法不然抓不到資料
'         strSql = "select a4202" & _
'                  " From acc420" & _
'                  " where upper(a4201)=upper('" & ChgSQL(cboTitle) & "') and (A4202<>'04150022' or A4202 is null)"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'         If intI = 0 Then
'            'Modify By Sindy 2016/8/22 + or cu05||' '||cu88||' '||cu89||' '||cu90='" & cboTitle & "' or cu06='" & cboTitle & "'
'            strSql = "select cu11" & _
'                     " From customer" & _
'                     " where (upper(cu04)=upper('" & ChgSQL(cboTitle) & "') or upper(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))=upper('" & ChgSQL(cboTitle) & "') or upper(cu06)=upper('" & ChgSQL(cboTitle) & "'))"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 0 Then
'               MsgBox "此為新的收據抬頭，請聯絡智權同仁提供基本資料以利建檔!!", vbInformation
'            End If
'         End If
'      End If
      'Add By Sindy 2017/6/19 改呼叫函數 : 檢查收據抬頭是否存在
      'Modified by Sindy 2018/9/18 拿掉chgsql
      Call PUB_ChkTitleNmExist(cboTitle)
      '2017/6/19 END
   '2015/11/16 END
   End If
   If txtMain(1) = "" Then
      MsgBox "請輸入個人/公司！"
      txtMain(1).SetFocus
      Exit Function
   End If
   
   'Add By Sindy 2012/11/15
   If m_Nation = "000" And (Me.Check3(1).Value = 1 Or Me.Check3(2).Value = 1) Then
      MsgBox "非台灣案時, 收據自動列印時間點才可選擇2 或 3 !!!", vbExclamation + vbOKOnly
      Exit Function
   End If
'cancel by sonia 2015/11/26 選擇收據自動列印時間點時,一定為暫不列印,故不必再勾選
'   If Me.Check1.Value = 1 And _
'      (Me.Check3(0).Value = 0 And Me.Check3(1).Value = 0 And Me.Check3(2).Value = 0) Then
'      MsgBox "勾選收據暫不列印時, 收據自動列印時間點不可空白!!!", vbExclamation + vbOKOnly
'      Exit Function
'   End If
'   '2012/11/15 End
'
'   'add by Sindy 2013/12/25
'   If Me.Check1.Value = 0 And _
'      (Me.Check3(0).Value = 1 Or Me.Check3(1).Value = 1 Or Me.Check3(2).Value = 1) Then
'      MsgBox "點選收據自動列印時間點, 收據暫不列印一定要勾選!!!", vbExclamation + vbOKOnly
'      Exit Function
'   End If
'   '2013/12/25 end
'end 2015/11/26
   
   'Add By Sindy 2012/12/6
   '檢查是否可上收據自動列印時間點
   If PUB_ChkAccIsUpdCP151(MSHFlexGrid1.TextMatrix(1, 5), IIf(Me.Check3(0).Value = 1, "1", IIf(Me.Check3(1).Value = 1, "2", IIf(Me.Check3(2).Value = 1, "3", "")))) = False Then
      Me.Check3(0).Value = 0
      Me.Check3(1).Value = 0
      Me.Check3(2).Value = 0
      Exit Function
   End If
   '2012/12/6 End
   
   '手開收據編號檢查
   If txtMain(3) <> "" And txtMain(3) <> m_OldNo Then
      'Added by Lydia 2016/09/05
      If OptKind(1).Value = True Then
          MsgBox "母子號不需輸入手開收據編號!!"
          txtMain_GotFocus 3
          txtMain(3).SetFocus
          Exit Function
      End If
      'end 2016/09/05
      strExc(0) = "select a0k01 from acc0k0 where a0k01 = '" & txtMain(3) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         MsgBox MsgText(28), , MsgText(5)
         txtMain(3).SetFocus
         Exit Function
      End If
      
      strExc(0) = "select a1m01 from acc1m0 where a1m01 = '" & txtMain(3) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox MsgText(42), , MsgText(5)
         txtMain(3).SetFocus
         Exit Function
      End If
      If pUpdate Then
         With MSHFlexGrid2
         For intI = 1 To .Rows - 1
            If .TextMatrix(intI, 0) <> txtMain(0) Then
               If .TextMatrix(intI, 16) = txtMain(3) Then
                  MsgBox "手開收據編號重複!!"
                  txtMain_GotFocus 3
                  txtMain(3).SetFocus
                  Exit Function
               End If
            End If
         Next
         End With
      End If
   End If
   InsertCheck = True
End Function

Private Sub txtNo_GotFocus()
   CloseIme
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 8 Then Exit Sub
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) Then
      KeyAscii = 0
      Beep
      Exit Sub
   End If
End Sub

'Added by Morgan 2011/12/1 拆收據用
Private Sub OpenTable1()
Dim strSpecCompany As String 'Add By Sindy 2013/12/26
   
   'Modify By Sindy 2012/11/15 +,a0j04,cp151
   'Modify By Sindy 2013/12/26 +,cp31,cp09
   'Modified by Lydia 2025/11/13 + ,cp10,cp14
   strExc(0) = "select cp01||cp02||cp03||cp04 本所案號,getcp10desc(cp01,cp10,a0j04) 案件性質,a0j09||'' 服務費,a0j10||'' 規費" & _
      ",a0j07,a0j01,cp05,cp01,cp02,cp03,cp04,cp12,cp13,K.*,a0j04,cp151,cp31,cp09,cp10,cp14 from acc0k0 K,acc0j0,caseprogress,casepropertymap" & _
      " where a0k01='" & m_OldNo & "' and a0j13(+)=a0k01 and cp09(+) = a0j01" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 order by 1,cp05,cp09"
   intI = 1
   Set MSHFlexGrid1.Recordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid1.Recordset
         'Add By Sindy 2013/12/26
         m_CP31 = "" & .Fields("CP31")
         m_CP01 = .Fields("CP01")
         m_CP02 = .Fields("CP02")
         m_CP03 = .Fields("CP03")
         m_CP04 = .Fields("CP04")
         '2013/12/26 END
         'Modified by Morgan 2023/5/11 拆收據時用原收據資料
         'm_CP12 = .Fields("cp12")
         'm_CP13 = .Fields("cp13")
         m_CP12 = .Fields("a0k22")
         m_CP13 = .Fields("a0k20")
         'end 2023/5/11
         
         m_CP09 = .Fields("CP09") 'Add By Sindy 2014/2/11
         Text1 = "" & .Fields("a0k11")
         txtDate = "" & .Fields("a0k02")
         txtCustNo = .Fields("a0k03")
         
         'Added by Lydia 2025/11/13
         strPropertyCode = .Fields("cp10")
         strPromoterNo = "" & .Fields("cp14")
         'end 2025/11/13
         
         '設定收據抬頭
         SetReceiptTitle "" & .Fields("a0k04")
         'Add By Sindy 2014/2/11 C類該案號最新收據抬頭
         strTitle = GetReceiptTitle_C(m_CP09, m_CP01 & m_CP02 & m_CP03 & m_CP04)
         If strTitle <> "" Then
            cboTitle.Text = strTitle
         End If
         '2014/2/11 END
         
         '暫不列印
         If .Fields("a0k32") = "N" Then
            Check1.Value = 1
         Else
            Check1.Value = 0
         End If
         
         'Add By Sindy 2012/11/15
         Me.Check3(0).Value = 0
         Me.Check3(1).Value = 0
         Me.Check3(2).Value = 0
         If Check1.Value = 1 Then 'add by sonia 2024/4/29 有暫不列印時,才需要帶出值
            If "" & .Fields("cp151") <> "" Then
               If "" & .Fields("cp151") = "1" Then Me.Check3(0).Value = 1
               If "" & .Fields("cp151") = "2" Then Me.Check3(1).Value = 1
               If "" & .Fields("cp151") = "3" Then Me.Check3(2).Value = 1
            End If
         End If                    'add by sonia 2024/4/29
         m_Nation = .Fields("a0j04")
         '2012/11/15 End
'         txtSales = "" & .Fields("a0k34") 'Add By Sindy 2013/12/26
'         lblSales = GetPrjSalesNM("" & .Fields("a0k34")) 'Add By Sindy 2013/12/26
         'Modify By Sindy 2014/12/29
         Combo2 = "" & .Fields("a0k34")
         Call Combo2_LostFocus
         '2014/12/29 END
         
'         'Add By Sindy 2017/3/17
'         If .Fields("a0k40") = "" Then
'            txtPrintNo.Text = ""
'         Else
'            txtPrintNo.Text = "" & .Fields("a0k40")
'         End If
'         '2017/3/17 END
         
         txtMain(1) = "" & .Fields("a0k05")
         txtMain(5) = "" & .Fields("a0k08")
         
         'Added by Morgan 2012/9/12
         .MoveFirst
         m_strChkCompany = "": m_strCaseNo = ""
         Do While Not .EOF
            'Modify By Sindy 2013/12/26
            If (.Fields("cp01") = "P" Or .Fields("cp01") = "CFP") Or _
               strSrvDate(1) >= InvoiceStartDate Then
               If InStr(m_strCaseNo, .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")) = 0 Then
                  strSpecCompany = ChkPatentNameCompany(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"))
                  If strSpecCompany <> "" And (strSpecCompany = m_strChkCompany Or m_strChkCompany = "") Then
                     m_strChkCompany = strSpecCompany
                     If m_strCaseNo <> "" Then m_strCaseNo = m_strCaseNo & ","
                     m_strCaseNo = m_strCaseNo & .Fields("cp01") & .Fields("cp02") & .Fields("cp03") & .Fields("cp04")
                  End If
               End If
            End If
            '2013/12/26 END
            .MoveNext
         Loop
         'end 2012/9/12
      End With
      
      'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
      'txtMain(6) = GetRecDay
      'txtMain(6).Tag = txtMain(6) '紀錄預設預定收款日
      'end 2018/08/22
      SetRestValue
   End If
   GridHead1
End Sub

'Added by Morgan 2011/12/1 拆收據
Private Function FormSave1() As Boolean
   Dim ii As Integer
   Dim strLstNo As String
   Dim A0K(35) As String
   Dim bolInsert As Boolean
   Dim strCP09 As String
   Dim bolInsDate As Boolean '新增預定收款日紀錄
   Dim strCP09List As String
   
On Error GoTo ErrHnd

   adoTaie.BeginTrans
   
   If Check2.Value = vbUnchecked Then
      SortData
   End If
   'Memo by Lydia 2016/09/05 先將原本資料的收據號碼統一設定,並且將清除原本的收據記錄
   adoTaie.Execute "update acc0j0 set a0j13='X' where a0j13='" & m_OldNo & "'", intI
   adoTaie.Execute "delete acc1m0 where a1m01='" & m_OldNo & "'", intI
   
   A0K(2) = txtDate
   A0K(3) = txtCustNo
   A0K(11) = Text1
   A0K(20) = m_CP13
   'modify by sonia 2017/6/22 若屬於業績列入P1001之專利處人員則智權人員改為P1001
   'A0K(22) = m_CP12
   'Modified by Morgan 2023/5/11 拆收據時用原收據資料
   'If adoquery.State = 1 Then adoquery.Close
   'adoquery.CursorLocation = adUseClient
   'adoquery.Open "select * FROM SetSpecMan where ocode='P1001' and instr(oman,'" & A0K(20) & "')>0 ", adoTaie, adOpenStatic, adLockReadOnly
   'If adoquery.RecordCount <> 0 Then
   '   A0K(20) = "P1001"
   'End If
   'adoquery.Close
   'A0K(22) = PUB_GetStaffST15(A0K(20), 1)
   A0K(22) = m_CP12
   'end 2023/5/11
   'end 2017/6/22
   
   With MSHFlexGrid2
   strLstNo = .TextMatrix(1, 0)
   '.TextMatrix(1, 16) = m_OldNo '收據編號
   For ii = 1 To .Rows - 1
      A0K(1) = .TextMatrix(ii, 16)
      'Modified by Morgan 2011/12/26 取消 a0j03,a0j05,a0j12,a0j20,a0j21
      strSql = "insert into acc0j0 (a0j01,a0j02,a0j04,a0j06,a0j07,a0j08,a0j09,a0j10" & _
         ",a0j11,a0j13,a0j14,a0j15,a0j16,a0j22,a0j21,a0j23,a0j24,a0j20,a0j25) select a0j01,a0j02,a0j04,a0j06,a0j07,a0j08," & Val(.TextMatrix(ii, 9)) & "," & Val(.TextMatrix(ii, 10)) & _
         ",a0j11,'" & A0K(1) & "',a0j14,a0j15,a0j16,'" & ChgSQL(.TextMatrix(ii, 8)) & "','" & .TextMatrix(ii, 11) & "','" & .TextMatrix(ii, 12) & "','" & .TextMatrix(ii, 13) & "','" & .TextMatrix(ii, 14) & "'," & Val(.TextMatrix(ii, 1)) & _
         " from acc0j0 where a0j01='" & .TextMatrix(ii, 18) & "' and a0j13='X'"
         
      adoTaie.Execute strSql, intI
      
      strSql = "insert into acc1m0 values('" & .TextMatrix(ii, 16) & "','" & .TextMatrix(ii, 18) & "')"
      adoTaie.Execute strSql, intI
      
      A0K(6) = Val(A0K(6)) + Val(.TextMatrix(ii, 9)) '服務費
      A0K(7) = Val(A0K(7)) + Val(.TextMatrix(ii, 10)) '規費
      
      '最後一筆也要新增最後一張收據
      If ii = .Rows - 1 Then
         'Modified by Lydia 2016/09/05
         'If .TextMatrix(ii, 16) = "" Then
            '.TextMatrix(ii, 16) = AutoNo(MsgText(802), 5, 1)   '收據編號
         If .TextMatrix(ii, 16) = "" Or (OptKind(1).Value = True And Mid(.TextMatrix(ii, 16), 1, monLen) <> Mid(m_OldNo, 1, monLen)) Then
            .TextMatrix(ii, 16) = GetAutoNo("" & .TextMatrix(ii, 18))
         End If
         bolInsert = True
      '下一筆序號不同新增收據
      ElseIf .TextMatrix(ii + 1, 0) <> .TextMatrix(ii, 0) Then
         'Modified by Lydia 2016/09/05
         'If .TextMatrix(ii + 1, 16) = "" Then
            '.TextMatrix(ii + 1, 16) = AutoNo(MsgText(802), 5, 1)  '收據編號
         If .TextMatrix(ii + 1, 16) = "" Or (OptKind(1).Value = True And Mid(.TextMatrix(ii + 1, 16), 1, monLen) <> Mid(m_OldNo, 1, monLen)) Then
            .TextMatrix(ii + 1, 16) = GetAutoNo("" & .TextMatrix(ii, 18))
         End If
         bolInsert = True
      ElseIf .TextMatrix(ii + 1, 0) = .TextMatrix(ii, 0) Then
         .TextMatrix(ii + 1, 16) = .TextMatrix(ii, 16)
      End If
      
      If bolInsert = True Then
         A0K(4) = .TextMatrix(ii, 2) '收據抬頭
         A0K(5) = .TextMatrix(ii, 3) '個人/公司
         A0K(8) = .TextMatrix(ii, 15) '備註
         A0K(32) = .TextMatrix(ii, 6) '收據暫不列印
         A0K(34) = .TextMatrix(ii, 21) '介紹案源同仁 Add By Sindy 2013/12/26
         'Added by Lydia 2024/08/05 國內接洽單：DEBIT NOTE請款選項
         If strShowCRL153 = "1" Or strShowCRL153 = "3" Then  '1=立即開立DEBIT NOTE, 3=待通知後開立，不需要加印國內收據
            A0K(32) = "Z" 'Z.確定不印(為了和暫不列印做取代)
         End If
         'end 2024/08/05
         
         strExc(0) = "select * from acc0k0 where a0k01='" & A0K(1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         'Modify By Sindy 2013/12/26 增加更新a0k34
         If intI = 1 Then
            'Modified by Lydia 2016/09/05 母號不動列印次數
            'strSql = "update acc0k0 set a0k02='" & A0K(2) & "', a0k03='" & A0K(3) & "', a0k04='" & ChgSQL(A0K(4)) & "'" & _
               ", a0k05='" & A0K(5) & "', a0k06=" & A0K(6) & ", a0k07=" & A0K(7) & ", a0k08='" & ChgSQL(A0K(8)) & "'" & _
               ", a0k09=0, a0k11='" & A0K(11) & "', a0k12='" & A0K(12) & "', a0k17=0, a0k18=0, a0k19=0" & _
               ", a0k20='" & A0K(20) & "',a0k22='" & A0K(22) & "', a0k27=" & strSrvDate(2) & _
               ", a0k28=to_char(sysdate,'HH24MiSS'), a0k29='" & strUserNum & "'" & _
               ", a0k32='" & A0K(32) & "', a0k33='Y', a0k34=" & CNULL(A0K(34)) & _
               " where a0k01='" & A0K(1) & "'"
            strSql = "update acc0k0 set a0k02='" & A0K(2) & "', a0k03='" & A0K(3) & "', a0k04='" & ChgSQL(A0K(4)) & "'" & _
               ", a0k05='" & A0K(5) & "', a0k06=" & A0K(6) & ", a0k07=" & A0K(7) & ", a0k08='" & ChgSQL(A0K(8)) & "'" & _
               ", a0k09=0, a0k11='" & A0K(11) & "', a0k12='" & A0K(12) & "', a0k17=0, a0k18=0 " & IIf(OptKind(1).Value = True, "", ", a0k19=0") & _
               ", a0k20='" & A0K(20) & "',a0k22='" & A0K(22) & "', a0k27=" & strSrvDate(2) & _
               ", a0k28=to_char(sysdate,'HH24MiSS'), a0k29='" & strUserNum & "'" & _
               ", a0k32='" & A0K(32) & "', a0k33='Y', a0k34=" & CNULL(A0K(34)) & _
               " where a0k01='" & A0K(1) & "'"
         Else
           'Modified by Lydia 2016/09/05 拆出的子號之列印次數自動上1(因為不用列印) ; +a0k23
            strSql = "insert into acc0k0 (a0k01, a0k02, a0k03, a0k04, a0k05, a0k06, a0k07" & _
               ", a0k08, a0k09, a0k11, a0k12, a0k17, a0k18, a0k19, a0k20,a0k22,a0k23, a0k24, a0k25, a0k26, a0k32, a0k33, a0k34) values " & _
               "('" & A0K(1) & "', " & A0K(2) & ", '" & A0K(3) & "', '" & ChgSQL(A0K(4)) & "', '" & A0K(5) & "', " & A0K(6) & ", " & A0K(7) & _
               ", '" & ChgSQL(A0K(8)) & "', 0, '" & A0K(11) & "', '" & A0K(12) & "', 0, 0, " & IIf(OptKind(1).Value = True, "1", "0") & ", '" & A0K(20) & "', '" & A0K(22) & "', " & CNULL(m_Nation) & ", " & strSrvDate(2) & ", to_char(sysdate,'HH24MiSS'), '" & strUserNum & "','" & A0K(32) & "','Y'," & CNULL(A0K(34)) & ")"
         End If
         adoTaie.Execute strSql, intI
         
         A0K(6) = 0
         A0K(7) = 0
         bolInsert = False
      End If
      strLstNo = .TextMatrix(ii, 0)
      
      'Add By Sindy 2012/11/15
      strSql = "update caseprogress set cp151=" & CNULL(.TextMatrix(ii, 20)) & " where cp09='" & .TextMatrix(ii, 18) & "'"
      adoTaie.Execute strSql, intI
      '2012/11/15 End
      
      'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'      If .TextMatrix(ii, 5) <> "" Then
'         If InStr(strCP09List, .TextMatrix(ii, 18)) = 0 Then '一個收文號一次
'            strCP09List = strCP09List & "," & .TextMatrix(ii, 18)
'            strExc(1) = DBDATE(.TextMatrix(ii, 5))
'            bolInsDate = False
'            strExc(0) = "select rd02*1000+rd03||rd05 from receivablesday where rd01='" & .TextMatrix(ii, 18) & "' order by 1 desc"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               '預定收款日期不同
'               If Mid(RsTemp(0), 12) <> strExc(1) Then
'                  bolInsDate = True
'               End If
'            Else
'               bolInsDate = True
'            End If
'            If bolInsDate = True Then
'               strSql = "insert into receivablesday (rd01,rd02,rd03,rd04,rd05)" & _
'                  " select '" & .TextMatrix(ii, 18) & "'," & strSrvDate(1) & ",nvl(max(rd03),0)+1,'" & strUserNum & "'," & strExc(1) & _
'                  " from receivablesday where rd01='" & .TextMatrix(ii, 18) & "' and rd02=" & strSrvDate(1)
'               adoTaie.Execute strSql, intI
'            End If
'         End If
'      End If
      'end 2018/08/22
   Next
   End With
   
   adoTaie.Execute "update caseprogress set cp60=(select min(a0j13) from acc0j0 where a0j01=cp09) where cp09 in (select a0j01 from acc0j0 where a0j13='X')", intI
   adoTaie.Execute "delete acc0j0 where a0j13='X'", intI
      
   Call ChkACStoEmail(m_CP01, m_CP09, strPropertyCode) 'Adderd by Lydia 2025/11/13
   
   adoTaie.CommitTrans
   FormSave1 = True
   Exit Function
   
ErrHnd:
   adoTaie.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function
'依相同項次收文日最早收文號最小的排序
Private Sub SortData(Optional ByVal pID As Integer)
   Dim iSort1 As Integer, iSort2 As Integer
   Dim iRow1 As Integer, iRow2 As Integer
   Dim ii As Integer, jj As Integer, kk As Integer, mm As Integer
   Dim stItem As String, iNo As Integer
   
   With MSHFlexGrid2
   
   '設定要排序的起迄序號
   If pID > 0 Then
      iSort1 = pID
      iSort2 = pID
   Else
      iSort1 = Val(.TextMatrix(1, 0))
      iSort2 = Val(.TextMatrix(.Rows - 1, 0))
   End If
   
   For mm = iSort1 To iSort2
      '設定序號的起迄列號
      iRow1 = 0
      For ii = 1 To .Rows - 1
         If Val(.TextMatrix(ii, 0)) = mm Then
            If iRow1 = 0 Then
               iRow1 = ii
            End If
         ElseIf iRow1 > 0 Then
            Exit For
         End If
      Next
      iRow2 = ii - 1
      
      '清除原排序
      For ii = iRow1 To iRow2
         .TextMatrix(ii, 1) = ""
      Next
      
      '從1開始
      iNo = 1
      
      Do While iNo <= iRow2 - iRow1 + 1
         
         For ii = iRow1 To iRow2
            If .TextMatrix(ii, 1) = "" Then
               kk = ii
               Exit For
            End If
         Next
         
         For ii = iRow1 To iRow2
            If .TextMatrix(ii, 1) = "" Then
               '先比收文日
               If .TextMatrix(ii, 19) < .TextMatrix(kk, 19) Then
                  kk = ii
               '再比收文號
               ElseIf .TextMatrix(ii, 19) = .TextMatrix(kk, 19) Then
                  If .TextMatrix(ii, 18) < .TextMatrix(kk, 18) Then
                     kk = ii
                  End If
               End If
            End If
         Next
         
         .TextMatrix(kk, 1) = iNo
         iNo = iNo + 1
         
         stItem = .TextMatrix(kk, 8)
         '相同帳款類別排一起
         For ii = iRow1 To iRow2
            If .TextMatrix(ii, 1) = "" Then
               If .TextMatrix(ii, 8) = stItem Then
                  .TextMatrix(ii, 1) = iNo
                  iNo = iNo + 1
               End If
            End If
         Next
      Loop
   Next
   End With
   Check2.Value = vbUnchecked
   ReSort
End Sub

Private Sub ReSort()
   With MSHFlexGrid2
   If m_lstRow > 0 Then
      SetColor m_lstRow, m_lstColor
   End If
   '指定列排序沒作用,只能有左到右排
   .col = 0
   .ColSel = 1
   .Sort = 1
   End With
   m_lstRow = 0
   txtMain(0) = ""
   ClearInput
End Sub

'檢查是否為拆收據補收文[其他]
Private Function IsSplitReceipt(p_CP10 As String, p_CP1234 As String, p_CP09 As String) As Boolean
   Dim stCP01 As String, stCP34 As String, stSQL As String, stCP10 As String, intR As Integer
   stCP01 = Left(p_CP1234, Len(p_CP1234) - 9)
   Select Case stCP01
      Case "P", "PS", "FCP", "FG", "CFP", "CPS"
         stCP10 = "910"
      Case "L", "LA", "FCL", "CFL"
         stCP10 = "7"
      Case Else
         stCP10 = "706"
   End Select
   If p_CP10 = stCP10 Then
      stSQL = "select cp05 from caseprogress a where cp09='" & p_CP09 & "' and cp14 is null" & _
         " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05 and b.cp14 is not null and b.cp16>0)"
      intR = 1
      Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         IsSplitReceipt = True
         stCP34 = Right(p_CP1234, 3)
         If stCP34 = "000" Then
            stCP34 = ""
         Else
            stCP34 = "-" & Left(stCP34, 1) & "-" & Right(stCP34, 2)
         End If
         m_strMailSubject = m_strMailSubject & IIf(m_strMailSubject = "", "拆收據補收文", "") & "【" & stCP01 & "-" & Mid(p_CP1234, Len(stCP01) + 1, 6) & stCP34 & "】;"
         m_strMailDesc = m_strMailDesc & "收文日：" & ChangeWStringToTDateString(RsTemp.Fields(0)) & _
                  vbTab & "收文號：" & p_CP09 & vbCrLf
      End If
   End If
End Function
'Added by Morgan 2012/9/12
'檢查專利案是否已專利商標出名
'Private Function ChkPatentNameCompany(pPA01 As String, pPA02 As String, pPA03 As String, pPA04 As String) As Boolean
'   Dim stSQL As String, adoRst As ADODB.Recordset, intR As Integer
'   stSQL = "select pa161 from patent where pa01='" & pPA01 & "' and pa02='" & pPA02 & "' and pa03='" & pPA03 & "' and pa04='" & pPA04 & "' and pa161='Y'"
'   intR = 1
'   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
'   If intR = 1 Then
'      ChkPatentNameCompany = True
'   End If
'End Function
Private Function ChkPatentNameCompany(pPA01 As String, pPA02 As String, pPA03 As String, pPA04 As String) As String
   Dim stSQL As String, adoRst As ADODB.Recordset, intR As Integer
   ChkPatentNameCompany = ""
   'Add By Sindy 2013/12/26
   If strSrvDate(1) >= InvoiceStartDate Then
      stSQL = "select pa161 from patent where pa01='" & pPA01 & "' and pa02='" & pPA02 & "' and pa03='" & pPA03 & "' and pa04='" & pPA04 & "' and pa161 is not null" & _
              " union select tm130 from trademark where tm01='" & pPA01 & "' and tm02='" & pPA02 & "' and tm03='" & pPA03 & "' and tm04='" & pPA04 & "' and tm130 is not null" & _
              " union select sp85 from servicepractice where sp01='" & pPA01 & "' and sp02='" & pPA02 & "' and sp03='" & pPA03 & "' and sp04='" & pPA04 & "' and sp85 is not null" & _
              " union select lc48 from lawcase where lc01='" & pPA01 & "' and lc02='" & pPA02 & "' and lc03='" & pPA03 & "' and lc04='" & pPA04 & "' and lc48 is not null"
   Else
   '2013/12/26 END
      stSQL = "select pa161 from patent where pa01='" & pPA01 & "' and pa02='" & pPA02 & "' and pa03='" & pPA03 & "' and pa04='" & pPA04 & "' and pa161='Y'"
   End If
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      ChkPatentNameCompany = Trim("" & adoRst.Fields(0).Value)
   End If
End Function

''Add By Sindy 2013/12/15
'Private Sub txtSales_GotFocus()
'   txtSales.SelStart = 0
'   txtSales.SelLength = Len(txtSales.Text)
'   '儲存未修改前之值至Tag中,供再確認時使用
'   txtSales.Tag = txtSales
'   '切換輸入法
'   CloseIme
'End Sub
'
''Add By Sindy 2013/12/15
'Private Sub txtSales_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
''Add By Sindy 2013/12/15
'Private Sub txtSales_Validate(Cancel As Boolean)
'Dim strTemp As String, strTemp1 As String
'
'   lblSales.Caption = ""
'   If txtSales.Text <> "" Then
'      If Not ClsPDGetStaff(txtSales.Text, strTemp, strTemp1) Then
'         Cancel = True
'         Exit Sub
'      End If
'      lblSales.Caption = strTemp
'   End If
'End Sub
'Add By Sindy 2014/12/29
Private Sub Combo2_GotFocus()
   InverseTextBox Combo2
End Sub
Private Sub Combo2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo2_LostFocus()
   If Combo2.Text > "" And Len(Trim(Combo2.Text)) = 5 Then
      '抓取員工姓名
      Combo2.Text = SetCboStaffName(Combo2.Text)
   End If
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
   If Combo2 <> "" Then
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(Combo2, 5)) = True Then
         Call Combo2_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
'2014/12/29 END

'Added by Lydia 2016/09/05 取得新收據號碼
Private Function GetAutoNo(ByVal tCP09 As String, Optional ByVal bolChk As Boolean = False) As String
Dim tmpNo As String
Dim tmpSQL As String
Dim rsT As New ADODB.Recordset
Dim inS As Integer

If bolChk Then GoTo Jump2CHK '判斷之前是否有拆母子號

   '拆新號
   If OptKind(0).Value = True Or tCP09 = "" Then
      GetAutoNo = AutoNo(MsgText(802), 5, 1)
      
   '母子號
   ElseIf OptKind(1).Value = True Then
Jump2CHK:
      tmpSQL = "select nvl(max(a0j13),'0') mno from acc0j0 where a0j01=" & CNULL(tCP09) & " and substr(a0j13,1," & monLen & ")=" & CNULL(Mid(m_OldNo, 1, monLen))
      inS = 1
      Set rsT = ClsLawReadRstMsg(inS, tmpSQL)
      If inS = 1 Then
        tmpNo = Mid(rsT(0), monLen + 1, 1)
        If tmpNo = "" Or tmpNo = "Z" Then
           GetAutoNo = rsT(0) & "1"
        ElseIf tmpNo = "9" Then
           GetAutoNo = Mid(rsT(0), 1, monLen) & "A"
        Else
           GetAutoNo = Mid(rsT(0), 1, monLen) & Chr(Asc(tmpNo) + 1)
        End If
        If tmpNo <> "" Then
           OptKind(0).Enabled = False
           OptKind(1).Value = 1
        End If
        
      End If
      Set rsT = Nothing
   End If
End Function

''Add by Amy 2020/03/27 接洽單公司別轉為作帳公司代號
'Private Function ChgCRL49ToBKeeping(stCRL49 As String) As String
'      If stCRL49 = "3" Then stCRL49 = "J"
'      If strSrvDate(1) >= 智慧所更名日 Then
'        If stCRL49 = "4" Then
'            stCRL49 = "L"
'        ElseIf stCRL49 <> "J" Then
'            stCRL49 = "2"
'        End If
'      Else
'        'CRL49:1-專利法律/2-專利商標/3-智權
'        If stCRL49 = "2" Then
'            stCRL49 = "1"
'        ElseIf stCRL49 <> "J" Then
'            stCRL49 = "2"
'        End If
'      End If
'      ChgCRL49ToBKeeping = stCRL49
'End Function

'Add by Amy 2025/02/20 收據抬頭開立之公司別,屬於每月代填公司別(cu168/a4220有此公司別),且無其公司別之同意書者發提醒信
Private Sub ChkWTConsentMail()
   Dim i As Integer, stWTCMailContent As String, stTO As String  '信件內容/收件者
   Dim stCmp As String, stCmpN As String, stTitleN As String, stID As String, stTmp As String
  
   stCmp = Text1
   If stCmp = "J" Then Exit Sub
   
   stCmpN = "法律所"
   If stCmp = "2" Then stCmp = "1": stCmpN = "智慧所"
   With MSHFlexGrid2
      For i = 1 To .Rows - 1
         stTitleN = .TextMatrix(i, 2)
         stID = .TextMatrix(i, 23)
         '每月代填公司別 有值且為畫面公司別
         If .TextMatrix(i, 24) <> MsgText(601) And InStr(.TextMatrix(i, 24), stCmp) > 0 Then
            If ChkWithholdingTaxConsent(0, Me.Name, stCmp, stTitleN) = False Then
               stTmp = stTmp & ",收據抬頭：" & stTitleN & vbCrLf & _
                                                "客戶編號：" & txtCustNo & vbCrLf & _
                                                "智權同仁：" & GetPrjSalesNM(m_CP13) & "(" & m_CP13 & ")" & vbCrLf & _
                                                "收據號碼：" & .TextMatrix(i, 16) & vbCrLf
            End If
         End If
      Next i
   End With
   If stTmp <> MsgText(601) Then
      stTO = "taieacc@taie.com.tw"
      stWTCMailContent = "以上收據抬頭為本所代填繳款書的客戶" & vbCrLf & _
                                             "但尚未提供 (" & stCmpN & ") 代填同意書" & vbCrLf & _
                                             "請聯絡提供"
      stTmp = Mid(stTmp, 2)
      stWTCMailContent = Replace(stTmp, ",", vbCrLf) & vbCrLf & stWTCMailContent
      PUB_SendMail strUserNum, stTO, "", "收據抬頭為本所代填繳款書的客戶,但尚未提供 (" & stCmpN & ") 代填同意書", stWTCMailContent
   End If
End Sub

'Added by Lydia 2025/11/13
Private Sub ChkACStoEmail(ByVal pCP01 As String, ByVal pCP09 As String, ByVal pCP10 As String)
Dim intQ As Integer, strQ1 As String, strQ2 As String
Dim rsQuery As New ADODB.Recordset

   If pCP01 = "ACS" And InStr(ACSforTIPSstep, pCP10) > 0 And pCP09 <> "" Then
      '抓拆收據新增acc0j0的a0j14=當天，a0j02=ACS，A0j01收文性質為區塊2，開立總金額=收文總金額，用mailcache發通知，不可重複通知；
      strQ1 = "SELECT cp01,cp02,cp03,cp04,cp16,cp60,sum(amt1) as totamt," & _
              "listagg(a0j13||'('||amt1||')','、') within group (order by a0j13) as totlist " & _
              "FROM (SELECT a0j13,sum(a0j09+a0j10) amt1,cp01,cp02,cp03,cp04,cp16,cp60 " & _
                    "FROM acc0j0,caseprogress WHERE a0j01='" & pCP09 & "' and a0j14=" & strSrvDate(2) & " AND a0j01=cp09(+) " & _
                    "GROUP BY a0j13,cp01,cp02,cp03,cp04,cp16,cp60) " & _
              "GROUP BY cp01,cp02,cp03,cp04,cp16,cp60 "
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 1 Then
         If "" & rsQuery.Fields("cp60") <> "" And "" & rsQuery.Fields("cp16") = "" & rsQuery.Fields("totamt") Then
            strQ2 = rsQuery.Fields("cp01") & "-" & rsQuery.Fields("cp02") & IIf("" & rsQuery.Fields("cp03") & rsQuery.Fields("cp04") <> "000", "-" & rsQuery.Fields("cp03") & "-" & rsQuery.Fields("cp04"), "") & "號已開立收據，請續行分案及請款階段設定"
            strSql = "select * from mailcache where mc03=to_char(sysdate,'yyyymmdd') and mc07='" & strQ2 & "' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then
               strQ1 = Pub_GetSpecMan("ACS分案人員")
               If strQ1 <> "" Then
                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                     " values( '" & strUserNum & "','" & strQ1 & "',to_char(sysdate,'yyyymmdd')" & _
                     ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strQ2) & "','如旨',null)"
                  adoTaie.Execute strSql, intQ
               End If
            End If
         End If
      End If
   End If
   Set rsQuery = Nothing
End Sub
