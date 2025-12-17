VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11p0 
   AutoRedraw      =   -1  'True
   Caption         =   "收據抬頭基本資料維護"
   ClientHeight    =   5880
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9072
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9072
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
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.FileListBox File1 
      Height          =   180
      Left            =   8280
      TabIndex        =   56
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
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
      Left            =   6240
      TabIndex        =   53
      Top             =   3000
      Width           =   218
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
      Left            =   7332
      TabIndex        =   52
      Top             =   3000
      Width           =   198
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   255
      Left            =   3330
      TabIndex        =   47
      Top             =   3330
      Width           =   5385
      Begin VB.OptionButton optA4228 
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
         Height          =   252
         Index           =   1
         Left            =   2340
         TabIndex        =   49
         Top             =   0
         Width           =   2985
      End
      Begin VB.OptionButton optA4228 
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
         Height          =   252
         Index           =   0
         Left            =   1110
         TabIndex        =   48
         Top             =   0
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
         TabIndex        =   50
         Top             =   0
         Width           =   1050
      End
   End
   Begin VB.CheckBox Check4 
      Caption         =   "寄紙本"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   2190
      TabIndex        =   46
      Top             =   4020
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CheckBox Check4 
      Caption         =   "寄電子檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3330
      TabIndex        =   45
      Top             =   4020
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.CheckBox Check1 
      Caption         =   "零稅率"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   7470
      TabIndex        =   42
      Top             =   4020
      Width           =   1100
   End
   Begin VB.TextBox textA4225 
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
      Left            =   1980
      MaxLength       =   1
      TabIndex        =   18
      Top             =   3320
      Width           =   345
   End
   Begin VB.TextBox textA4224 
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
      Left            =   5880
      MaxLength       =   1
      TabIndex        =   3
      Top             =   450
      Width           =   345
   End
   Begin VB.TextBox TextA4220 
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
      Height          =   605
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   39
      Text            =   "Frmacc11p0.frx":0000
      Top             =   4950
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox textA4221 
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
      Left            =   5616
      MaxLength       =   1
      TabIndex        =   16
      Top             =   2640
      Width           =   315
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
      Height          =   345
      Left            =   5250
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "複製基本資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7080
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.OptionButton optCustomer 
      Caption         =   "特殊機構"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   6780
      TabIndex        =   11
      Top             =   1590
      Width           =   1305
   End
   Begin VB.OptionButton optCustomer 
      Caption         =   "學校"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   5850
      TabIndex        =   10
      Top             =   1590
      Width           =   825
   End
   Begin VB.OptionButton optCustomer 
      Caption         =   "個人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   3990
      TabIndex        =   8
      Top             =   1590
      Width           =   825
   End
   Begin VB.OptionButton optCustomer 
      Caption         =   "公司"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4920
      TabIndex        =   9
      Top             =   1590
      Width           =   825
   End
   Begin VB.TextBox textA4218 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5190
      MaxLength       =   200
      TabIndex        =   13
      Top             =   1890
      Width           =   3615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "境外公司"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   6090
      TabIndex        =   4
      Top             =   4020
      Width           =   1300
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   8370
      Picture         =   "Frmacc11p0.frx":0020
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   120
      Width           =   350
   End
   Begin VB.TextBox textA4204 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1515
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1530
      Width           =   2055
   End
   Begin VB.TextBox textA4205 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1515
      MaxLength       =   20
      TabIndex        =   12
      Top             =   1890
      Width           =   2055
   End
   Begin VB.TextBox textA4206 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1515
      MaxLength       =   6
      TabIndex        =   14
      Top             =   2250
      Width           =   1215
   End
   Begin VB.TextBox textA4202 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1515
      MaxLength       =   10
      TabIndex        =   2
      Top             =   450
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   1380
      Top             =   0
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   572
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   4920
      TabIndex        =   57
      Top             =   1920
      Width           =   260
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
      Height          =   288
      Index           =   0
      Left            =   6444
      TabIndex        =   55
      Top             =   3000
      Width           =   648
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
      Height          =   288
      Index           =   1
      Left            =   7512
      TabIndex        =   54
      Top             =   3000
      Width           =   648
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "每月代填繳款同意書"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   4140
      TabIndex        =   51
      Top             =   3024
      Width           =   2052
   End
   Begin MSForms.TextBox textA4223 
      Height          =   330
      Left            =   1515
      TabIndex        =   17
      Top             =   2970
      Width           =   2325
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "4101;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4222 
      Height          =   330
      Left            =   1515
      TabIndex        =   19
      Top             =   3630
      Width           =   7305
      VariousPropertyBits=   -1466941413
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "12885;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   975
      Left            =   1515
      TabIndex        =   21
      Top             =   4620
      Width           =   7305
      VariousPropertyBits=   -1466941413
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "12885;1720"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4215 
      Height          =   345
      Left            =   1515
      TabIndex        =   5
      Top             =   810
      Width           =   6975
      VariousPropertyBits=   -1466941413
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "12303;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4208 
      Height          =   330
      Left            =   1515
      TabIndex        =   20
      Top             =   4290
      Width           =   7305
      VariousPropertyBits=   -1466941413
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "12885;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4207 
      Height          =   345
      Left            =   1515
      TabIndex        =   15
      Top             =   2610
      Width           =   2325
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4101;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4201 
      Height          =   345
      Left            =   1515
      TabIndex        =   0
      Top             =   90
      Width           =   6825
      VariousPropertyBits=   671105051
      MaxLength       =   100
      Size            =   "12039;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textA4203 
      Height          =   345
      Left            =   1515
      TabIndex        =   6
      Top             =   1170
      Width           =   6975
      VariousPropertyBits=   -1466941413
      MaxLength       =   80
      ScrollBars      =   2
      Size            =   "12303;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "電子發票寄送方式："
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
      Left            =   90
      TabIndex        =   44
      Top             =   4020
      Visible         =   0   'False
      Width           =   2160
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
      Height          =   270
      Left            =   8580
      TabIndex        =   43
      Top             =   4005
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收據列印統一編號       (Y:印)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   96
      TabIndex        =   41
      Top             =   3355
      Width           =   2952
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "不寄發扣繳核對資料      (N:不寄)"
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
      Left            =   3660
      TabIndex        =   40
      Top             =   510
      Width           =   3510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "繳款書寄件處     (1.客戶 2.會計師 3.特殊)"
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
      Left            =   4140
      TabIndex        =   38
      Top             =   2700
      Width           =   4416
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "特殊收件人："
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
      Left            =   90
      TabIndex        =   37
      Top             =   3030
      Width           =   1440
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "特殊地址："
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
      Left            =   270
      TabIndex        =   36
      Top             =   3690
      Width           =   1200
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "E-mail(財務)："
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
      Index           =   3
      Left            =   3630
      TabIndex        =   35
      Top             =   1950
      Width           =   1575
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "會計備註："
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
      Index           =   2
      Left            =   270
      TabIndex        =   34
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "營業地址："
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
      Left            =   270
      TabIndex        =   33
      Top             =   870
      Width           =   1200
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "備　　註："
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
      Index           =   1
      Left            =   270
      TabIndex        =   32
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "聯 絡 人："
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
      Left            =   270
      TabIndex        =   31
      Top             =   2700
      Width           =   1110
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭："
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
      TabIndex        =   30
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "郵寄地址："
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
      Index           =   18
      Left            =   270
      TabIndex        =   29
      Top             =   1230
      Width           =   1200
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "電　　話："
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
      Index           =   9
      Left            =   270
      TabIndex        =   28
      Top             =   1590
      Width           =   1200
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "傳　　真："
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
      Index           =   10
      Left            =   270
      TabIndex        =   27
      Top             =   1950
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "智權人員："
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
      Index           =   6
      Left            =   270
      TabIndex        =   26
      Top             =   2310
      Width           =   1200
   End
   Begin MSForms.Label Label30 
      Height          =   300
      Index           =   2
      Left            =   2790
      TabIndex        =   25
      Top             =   2310
      Width           =   795
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "統一編號："
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
      Index           =   10
      Left            =   270
      TabIndex        =   24
      Top             =   510
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   120
      Top             =   4530
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc11p0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Create by Sindy 2013/12/19
Option Explicit

Dim adoadodc1 As New ADODB.Recordset
Public bolCallMe As Boolean 'Add By Sindy 2015/9/14
Public ProState As String '權限: 1.全所 2.該所 add by sonia 2025/5/22
Dim strA4902 As String, strA4903 As String, strA4904 As String, strA4905 As String 'Add By Sindy 2016/11/1
Dim strA4912 As String, strA4913 As String, strA4914 As String 'Add By Sindy 2016/11/7
Dim m_PrevForm As Form '前一畫面'Add By Sindy 2016/11/29
Dim m_DefColor As Long, m_SetColor As Long 'Add by Amy 2025/02/20

'Add By Sindy 2016/3/21 由未勾選改為有勾選存檔時,若當年已有扣繳資料, 顯示訊息, 但仍可操作
Private Sub Check1_Click(Index As Integer)
'Modify by Amy 2019/07/22 +if 零稅率
If Index = 1 Then
   '統編不為空不可勾選零稅率
   If Check1(1).Value = 1 And Len(Trim(textA4202)) > 0 Then
        MsgBox "有統一編號不可勾選零稅率", , MsgText(5)
        Check1(1).Value = 0
   End If
'境外公司
Else
   If Check1(0).Tag = "" And Check1(0).Value = 1 Then
      'Modify By Sindy 2019/10/15 有無扣繳資料,要比對到有沒有ACC1v0檔案資料
      strExc(0) = "select a0k01 from acc0k0,acc1v0 where a0k04='" & ChgSQL(textA4201) & "' and a0k05<>'1' and nvl(a0k16,0)=" & Left(strSrvDate(2), 3) & " and nvl(a0k09,0)=0" & _
                  " and a0k01=a1v02(+) and a1v02 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("此客戶 " & textA4201 & " 今年已有可扣繳資料, 確定是境外公司嗎？", vbYesNo + vbDefaultButton1 + vbExclamation) = vbNo Then
            Check1(0).Value = 0
         End If
      End If
   End If
End If
End Sub

'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再設定
''Add by Amy 2019/07/25
'Private Sub Check4_Click(Index As Integer)
'    If strSaveConfirm = MsgText(601) Then Exit Sub
'
'    If Check4(0).Value = vbChecked And Check4(1).Value = vbChecked Then
'        MsgBox "電子發票寄送方式只能擇一選擇", , MsgText(5)
'    End If
'End Sub

'Add By Sindy 2016/11/1
Private Sub cmdA49_Click()
   Frmacc21v0.Hide
   Frmacc21v0.textA4901_C.Visible = False
   Frmacc21v0.textA4901.Visible = True
   Frmacc21v0.textA4901.Text = textA4201
   Frmacc21v0.Tag = Me.Name 'Add By Sindy 2025/6/23
   Frmacc21v0.OpenTable
   Frmacc21v0.Show vbModal
   'Add By Sindy 2016/11/8 有資料按鈕變顏色
   cmdA49.BackColor = &H8000000F
   strExc(0) = "select A4901 from ACC490 where A4901='" & ChgSQL(textA4201) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      cmdA49.BackColor = &HC0FFC0
   End If
   '2016/11/8 END
End Sub

'Add By Sindy 2016/5/31 複製基本資料
Private Sub cmdCopy_Click()
   strControlButton = MsgText(602) '不清除畫面上的資料,收據抬頭除外
   textA4208 = "基本資料同原收據抬頭名稱:" & ChgSQL(textA4201.Text) & ";" & textA4208.Text
   'Add By Sindy 2016/11/1
   strSql = "select *" & _
            " From acc490" & _
            " where a4901='" & ChgSQL(textA4201) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   cmdA49.Visible = False
   cmdA49.Tag = ""
   If intI = 1 Then
      strA4902 = "" & RsTemp.Fields("a4902")
      strA4903 = "" & RsTemp.Fields("a4903")
      strA4904 = "" & RsTemp.Fields("a4904")
      strA4905 = "" & RsTemp.Fields("a4905")
      strA4912 = "" & RsTemp.Fields("a4912") 'Add By Sindy 2016/11/7
      strA4913 = "" & RsTemp.Fields("a4913") 'Add By Sindy 2016/11/7
      strA4914 = "" & RsTemp.Fields("a4914") 'Add By Sindy 2016/11/7
      cmdA49.Tag = "複製"
   End If
   '2016/11/1 END
   KeyEnter vbKeyF2
End Sub

'Modify By Sindy 2015/7/8
'Private Sub Command3_Click()
Public Sub Command3_Click()
'2015/7/8 END
Dim Rs As New ADODB.Recordset
'Dim strFindA4201 As String
   
   If textA4201 = MsgText(601) Then
      Exit Sub
   Else
      strCompanyNo = textA4201
   End If
   If strSaveConfirm = MsgText(3) Then '新增狀態時
      '先檢查在多筆視窗中是否已有電匯資料,若有不可再新增
      Adodc1.Recordset.Find "a4201 = '" & ChgSQL(textA4201) & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
         Exit Sub
      End If
   Else
      Adodc1.Recordset.Find "a4201 = '" & ChgSQL(textA4201) & "'", 0, adSearchForward, 1
      If Adodc1.Recordset.EOF = False Then
         FormShow
         RecordShow
         'Add By Sindy 2015/9/14
         If bolCallMe = True Then
            Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True '修改
         End If
         '2015/9/14 END
      Else
         MsgBox MsgText(33), , MsgText(5)
         'Add By Sindy 2015/9/14
         If bolCallMe = True Then
            strSaveConfirm = "A"
            FormEnabled
            Frmacc11p0_Clear
            'textA4201 = strFindA4201
            textA4201 = strCompanyNo
            textA4221 = "1" 'Add By Sindy 2016/12/2 繳款書寄件處預設為1客戶
            tool2_enabled
         End If
         '2015/9/14 END
         If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.MoveFirst
         End If
      End If
   End If
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strCompanyNo = MsgText(601) Then
      Exit Sub
   End If
   If Adodc1.Recordset.RecordCount <> 0 Then
      Adodc1.Recordset.MoveFirst
   End If
   Adodc1.Recordset.Find "a4201 = '" & ChgSQL(strCompanyNo) & "'", 0, adSearchForward, 1
   If Adodc1.Recordset.EOF = False Then
      FormShow
      RecordShow
   End If
'   'Add By Sindy 2024/9/2
'   PUB_WriteDebugLog ("strCompanyNo=" & strCompanyNo & "; strFormName=" & strFormName & ";")
'   '2024/9/2 END
   strCompanyNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim intCounter As Integer
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/06 原H:6120/W:9045
   Me.Height = 6345
   Me.Width = 9195
   'Modify by Amy 2023/10/06 原(lngWidth - Me.Width) / 2,切畫面不需再調,故左移-瑞婷
   Me.Move 0, (lngHeight - Me.Height) / 2 + 900
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   strCompanyNo = MsgText(601)
   Label30(2).Caption = ""
   g_LetterDebug = True 'Modify By Sindy 2024/9/2 要記錄Log
   
   OpenTable
   Call FormDisabled
   'Add by Amy 2025/02/20 有同意書檔變色
   m_DefColor = &H8000000F
   m_SetColor = RGB(215, 117, 117)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strCU11 As String
   
   g_LetterDebug = False 'Modify By Sindy 2024/9/2 取消記錄Log
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   If bolCallMe = False Then 'Add By Sindy 2015/9/14 +if
      If Adodc1.Recordset.RecordCount <> 0 Then
         strCompanyNo = textA4201 'Adodc1.Recordset.Fields("a4201").Value
      Else
         strCompanyNo = MsgText(601)
      End If
   End If
   
   StatusClear
   
   'Add By Sindy 2014/3/31 若是從收據抬頭修改進入請款單開立發票作業至此作業,要重新讀取統一編號
   If UCase(strTitle) = UCase("Frmacc1140") And UCase(strUserLevel) = UCase("Frmacc1127") Then
      strCU11 = ""
      strSql = "select cu11" & _
               " From customer" & _
               " where cu04='" & Frmacc1127.labA0K04 & "'" & _
               " and cu15<>'0'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strCU11 = "" & RsTemp.Fields("cu11")
      End If
      If strCU11 = "" Then
         'Modify By Sindy 2017/4/18 and A4202<>'04150022'==>and (A4202<>'04150022' or A4202 is null) 改語法不然抓不到資料
         strSql = "select a4202" & _
                  " From acc420" & _
                  " where a4201='" & Frmacc1127.labA0K04 & "' and (A4202<>'04150022' or A4202 is null)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strCU11 = "" & RsTemp.Fields("a4202")
         End If
      End If
      Frmacc1127.labA4303.Caption = strCU11
      Frmacc1127.CmdSave_Click
   End If
   '2014/3/31 END
   
   'Add By Sindy 2015/7/8
   If UCase(strUserLevel) = UCase("Frmacc44t0") Then
      Frmacc44t0.Show
      Frmacc44t0.cmdQuery_Click
      tool3_enabled
   '2015/7/8 END
   'Add By Sindy 2016/11/15
   'Modify By Sindy 2016/11/29
   'ElseIf UCase(strUserLevel) = UCase("Frmacc11b0") Then
   ElseIf TypeName(m_PrevForm) <> "Nothing" Then
      If UCase(m_PrevForm.Name) = UCase("Frmacc11b0") Or _
         UCase(m_PrevForm.Name) = UCase("Frmacc44w1") Then
   '2016/11/29 END
         m_PrevForm.Show
   '      m_PrevForm.cmdQuery_Click
         tool3_enabled
      End If
   Else
      KeyEnter vbKeyEscape
   End If
   '2016/11/15 END
   
   strUserLevel = MsgText(601) 'Add By Sindy 2015/12/10
   strConTitle = MsgText(601)
   strFormName = MsgText(601)
   'KeyEnter vbKeyEscape
   MenuEnabled
   
   Set m_PrevForm = Nothing 'Add By Sindy 2016/11/29
   Set Frmacc11p0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim strCon2 As String

On Error GoTo Checking

   adoadodc1.CursorLocation = adUseClient
   strSql = "select * from acc420 order by a4210,a4211 asc"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   strControlButton = MsgText(601)
   textA4201 = Adodc1.Recordset.Fields("A4201").Value
   'textA4201.Enabled = False 'Modify By Sindy 2014/3/31 Mark
   If IsNull(Adodc1.Recordset.Fields("A4202").Value) Then
      textA4202 = MsgText(601)
   Else
      textA4202 = Adodc1.Recordset.Fields("A4202").Value
      If UCase(strUserLevel) <> UCase("Frmacc44t0") Then
         cmdCopy.Visible = True 'Add By Sindy 2016/5/31 查詢出資料時,才顯示此按鈕
      End If
      cmdA49.Visible = True 'Add By Sindy 2016/11/1
      cmdA49.Tag = "" 'Add By Sindy 2016/11/1
   End If
   If IsNull(Adodc1.Recordset.Fields("A4203").Value) Then
      textA4203 = MsgText(601)
   Else
      textA4203 = Adodc1.Recordset.Fields("A4203").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("A4204").Value) Then
      textA4204 = MsgText(601)
   Else
      textA4204 = Adodc1.Recordset.Fields("A4204").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("A4205").Value) Then
      textA4205 = MsgText(601)
   Else
      textA4205 = Adodc1.Recordset.Fields("A4205").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("A4206").Value) Then
      textA4206 = MsgText(601)
   Else
      textA4206 = Adodc1.Recordset.Fields("A4206").Value
   End If
   Call textA4206_Validate(False)
   If IsNull(Adodc1.Recordset.Fields("A4207").Value) Then
      textA4207 = MsgText(601)
   Else
      textA4207 = Adodc1.Recordset.Fields("A4207").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("A4208").Value) Then
      textA4208 = MsgText(601)
   Else
      textA4208 = Adodc1.Recordset.Fields("A4208").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("A4215").Value) Then
      textA4215 = MsgText(601)
   Else
      textA4215 = Adodc1.Recordset.Fields("A4215").Value
   End If
   'Add by Amy 2014/09/23 +境外公司
   If IsNull(Adodc1.Recordset.Fields("A4216").Value) Then
      Check1(0).Value = 0
      Check1(0).Tag = "" 'Add By Sindy 2016/2/15
   Else
      Check1(0).Value = vbChecked
      Check1(0).Tag = "Y" 'Add By Sindy 2016/2/15
   End If
   Text1.Text = "" & Adodc1.Recordset.Fields("A4217").Value 'Add By Sindy 2014/10/15
   'Add By Sindy 2015/6/2
   If IsNull(Adodc1.Recordset.Fields("A4218").Value) Then
      textA4218 = MsgText(601)
   Else
      textA4218 = Adodc1.Recordset.Fields("A4218").Value
   End If
   If Not IsNull(Adodc1.Recordset.Fields("A4219").Value) Then
      optCustomer(Adodc1.Recordset.Fields("A4219").Value).Value = True
   End If
   '2015/6/2 END
   'Add By Sindy 2016/11/7
   'Modify by Amy 2025/02/20 原Check3=每月代填繳款書,改可勾智慧所/法律所,且已有同意書檔變色
   Check3(0).BackColor = m_DefColor: lblCU168(0).BackColor = m_DefColor
   Check3(1).BackColor = m_DefColor: lblCU168(1).BackColor = m_DefColor
   If IsNull(Adodc1.Recordset.Fields("A4220").Value) Then
      Check3(0).Value = 0: TextA4220.Visible = False
      Check3(1).Value = 0
   Else
      If InStr("," & Adodc1.Recordset.Fields("A4220"), ",1") > 0 Then
         Check3(0).Value = 1
         If ChkWithholdingTaxConsent(0, Me.Name, "1", textA4201, textA4202) = True Then
            Check3(0).BackColor = m_SetColor: lblCU168(0).BackColor = m_SetColor
         End If
      End If
      
      If InStr("," & Adodc1.Recordset.Fields("A4220"), ",L") > 0 Then
         Check3(1).Value = 1
         If ChkWithholdingTaxConsent(0, Me.Name, "L", textA4201, textA4202) = True Then
            Check3(1).BackColor = m_SetColor: lblCU168(1).BackColor = m_SetColor
         End If
      End If
      TextA4220.Visible = True
   End If
   'end 2025/02/20
   'Add By Sindy 2019/12/18
   If IsNull(Adodc1.Recordset.Fields("A4228").Value) Then
      optA4228(0).Value = False: optA4228(1).Value = False
   Else
      If Adodc1.Recordset.Fields("A4228").Value = "1" Then
         optA4228(0).Value = True
      ElseIf Adodc1.Recordset.Fields("A4228").Value = "2" Then
         optA4228(1).Value = True
      End If
   End If
   '2019/12/18 END
   
   If IsNull(Adodc1.Recordset.Fields("A4221").Value) Then
      textA4221 = MsgText(601)
   Else
      textA4221 = Adodc1.Recordset.Fields("A4221").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("A4222").Value) Then
      textA4222 = MsgText(601)
   Else
      textA4222 = Adodc1.Recordset.Fields("A4222").Value
   End If
   If IsNull(Adodc1.Recordset.Fields("A4223").Value) Then
      textA4223 = MsgText(601)
   Else
      textA4223 = Adodc1.Recordset.Fields("A4223").Value
   End If
   '2016/11/7 END
   'Add By Sindy 2017/3/16
   If IsNull(Adodc1.Recordset.Fields("A4224").Value) Then
      textA4224 = MsgText(601)
   Else
      textA4224 = Adodc1.Recordset.Fields("A4224").Value
   End If
   '2017/3/16 END
   'Add By Sindy 2017/3/24
   If IsNull(Adodc1.Recordset.Fields("A4225").Value) Then
      textA4225 = MsgText(601)
   Else
      textA4225 = Adodc1.Recordset.Fields("A4225").Value
   End If
   '2017/3/24 END
   'Add by Amy 2019/07/22 零稅率
   If IsNull(Adodc1.Recordset.Fields("A4226").Value) Then
        Check1(1).Value = 0
        Check1(1).Tag = ""
   Else
        Check1(1).Value = vbChecked
        Check1(1).Tag = "Y"
   End If
   'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再設定
   'Add by Amy 2019/07/25 電子發票寄送方式
'   If IsNull(Adodc1.Recordset.Fields("A4227").Value) Then
'        '未設定
'   Else
'        '紙本
'        If Adodc1.Recordset.Fields("A4227").Value = "1" Then Check4(0).Value = vbChecked
'        '電子檔
'        If Adodc1.Recordset.Fields("A4227").Value = "2" Then Check4(1).Value = vbChecked
'   End If
'   Check4(0).Tag = "" & Adodc1.Recordset.Fields("A4227").Value
'   'end 2019/07/25
   
   'Add By Sindy 2016/11/8 有資料按鈕變顏色
   cmdA49.BackColor = &H8000000F
   strExc(0) = "select A4901 from ACC490 where A4901='" & ChgSQL(textA4201) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      cmdA49.BackColor = &HC0FFC0
   End If
   '2016/11/8 END
End Sub

'*************************************************
'  關閉分錄欄位輸入狀態
'
'*************************************************
Public Sub FormDisabled()
   textA4201.Enabled = True 'False 'Modify By Sindy 2014/3/31
   textA4202.Enabled = False
   textA4203.Enabled = False
   textA4204.Enabled = False
   textA4205.Enabled = False
   textA4206.Enabled = False
   textA4207.Enabled = False
   textA4208.Enabled = False
   textA4215.Enabled = False
   'Add by Amy 2014/09/23 +境外公司
   Check1(0).Enabled = False
   Text1.Enabled = False 'Add By Sindy 2014/10/15
   'Add By Sindy 2015/6/2
   'Modify By Sindy 2020/2/15
   'textA4218.Enabled = False
   textA4218.Locked = True
   '2020/2/15 END
   optCustomer(0).Enabled = False
   optCustomer(1).Enabled = False
   optCustomer(2).Enabled = False
   optCustomer(3).Enabled = False
   '2015/6/2 END
   'Add By Sindy 2016/11/7
   'Modify by Amy 2025/02/20 原Check3=每月代填繳款書,改可勾智慧所/法律所
   Check3(0).Enabled = False
   Check3(1).Enabled = False
   'end 2025/02/20
   optA4228(0).Enabled = False: optA4228(1).Enabled = False 'Add By Sindy 2019/12/18
   textA4221.Enabled = False
   textA4222.Enabled = False
   textA4223.Enabled = False
   '2016/11/7 END
   textA4224.Enabled = False 'Add By Sindy 2017/3/16
   textA4225.Enabled = False 'Add By Sindy 2017/3/24
   Check1(1).Enabled = False 'Add by Amy 2019/07/22 零稅率
   'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再設定
'   'Add by Amy 2019/07/25 電子發票寄送方式
'   Check4(0).Enabled = False
'   Check4(1).Enabled = False
End Sub

'*************************************************
'  開啟分錄欄位輸入狀態
'
'*************************************************
Public Sub FormEnabled()
   If strSaveConfirm = MsgText(3) Then '新增狀態時
      textA4201.Enabled = True
   Else
      If textA4201 <> "" Then
         strCompanyNo = textA4201 'Add By Sindy 2015/10/23
         textA4201.Enabled = False
         Me.cmdCopy.Visible = False 'Add By Sindy 2016/5/31
         cmdA49.Visible = True 'Add By Sindy 2016/11/1
         cmdA49.Tag = "" 'Add By Sindy 2016/11/1
      Else
         textA4201.Enabled = True
      End If
   End If
   
   'Modify by Sindy 2015/10/15 Mark
   If bolCallMe = True Then
      textA4201 = strCompanyNo
      textA4201.Enabled = False
   End If
   '2015/10/15 END
   
   textA4202.Enabled = True
   textA4203.Enabled = True
   textA4204.Enabled = True
   textA4205.Enabled = True
   textA4206.Enabled = True
   textA4207.Enabled = True
   textA4208.Enabled = True
   textA4215.Enabled = True
   'Add by Amy 2014/09/23 +境外公司
   Check1(0).Enabled = True
   Text1.Enabled = True 'Add By Sindy 2014/10/15
   'Add By Sindy 2015/6/2
   'Modify By Sindy 2020/2/15
   'textA4218.Enabled = True
   textA4218.Locked = False
   '2020/2/15 END
   optCustomer(0).Enabled = True
   optCustomer(1).Enabled = True
   optCustomer(2).Enabled = True
   optCustomer(3).Enabled = True
   '2015/6/2 END
   'Add By Sindy 2016/11/7
   'Modify by Amy 2025/02/20 原Check3=每月代填繳款書,改可勾智慧所/法律所
   Check3(0).Enabled = True
   Check3(1).Enabled = True
   'end 2025/02/20
   optA4228(0).Enabled = True: optA4228(1).Enabled = True 'Add By Sindy 2019/12/18
   textA4221.Enabled = True
   textA4222.Enabled = True
   textA4223.Enabled = True
   '2016/11/7 EMD
   textA4224.Enabled = True 'Add By Sindy 2017/3/16
   textA4225.Enabled = True 'Add By Sindy 2017/3/24
   Check1(1).Enabled = True 'Add by Amy 2019/07/22 零稅率
   'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再設定
'   'Add by Amy 2019/07/25 電子發票寄送方式
'   Check4(0).Enabled = True
'   Check4(1).Enabled = True
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
Dim strCon2 As String

On Error GoTo Checking

   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   strSql = "select * from acc420 order by a4210,a4211 asc"
   adoadodc1.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount <> 0 Then
      If textA4201 <> MsgText(601) Then
         Adodc1.Recordset.Find "a4201 = '" & ChgSQL(textA4201) & "'", 0, adSearchForward, 1
         If Adodc1.Recordset.EOF = False Then
            FormShow
            RecordShow
         End If
      End If
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   'Modify By Sindy 2024/9/2 + & ";" & strCompanyNo 增加顯示資訊
   Frmacc0000.StatusBar1.Panels(2).Text = Adodc1.Recordset.Bookmark & MsgText(35) & Adodc1.Recordset.RecordCount & ";" & strCompanyNo
End Sub

Public Sub Frmacc11p0_Clear()
   With Frmacc11p0
      .cmdCopy.Visible = False 'Add By Sindy 2016/5/31
      .cmdA49.Visible = False 'Add By Sindy 2016/11/1
      .cmdA49.Tag = "" 'Add By Sindy 2016/11/1
      .textA4201 = ""
      .textA4202 = ""
      .textA4203 = ""
      .textA4204 = ""
      .textA4205 = ""
      .textA4206 = ""
      .textA4207 = ""
      .textA4208 = ""
      .textA4215 = ""
      .Label30(2).Caption = ""
      'Add by Amy 2014/09/23 +境外公司
      Check1(0).Value = 0
      Check1(0).Tag = "" 'Add By Sindy 2016/2/15
      If textA4201.Enabled = True Then
         .textA4201.SetFocus
      End If
      .Text1 = "" 'Add By Sindy 2014/10/15
      'Add By Sindy 2015/6/2
      .textA4218 = ""
      optCustomer(0).Value = False
      optCustomer(1).Value = False
      optCustomer(2).Value = False
      optCustomer(3).Value = False
      '2015/6/2 END
      'Add By Sindy 2016/11/7
      'Modify by Amy 2025/02/20 原Check3=每月代填繳款書,改可勾智慧所/法律所,且已有同意書檔變色
      Check3(0).Value = 0: TextA4220.Visible = False
      Check3(1).Value = 0
      Check1(0).BackColor = m_DefColor: lblCU168(0).BackColor = m_DefColor
      Check1(1).BackColor = m_DefColor: lblCU168(1).BackColor = m_DefColor
      'end 2025/02/20
      .optA4228(0).Enabled = False: .optA4228(1).Enabled = False 'Add By Sindy 2019/12/18
      .textA4221 = ""
      .textA4222 = ""
      .textA4223 = ""
      '2016/11/7 END
      .textA4224 = "" 'Add By Sindy 2017/3/16
      .textA4225 = "" 'Add By Sindy 2017/3/24
      
      'Modify by Sindy 2015/10/15 Mark
      If bolCallMe = True Then
         textA4201 = strCompanyNo
         textA4201.Enabled = False
      End If
      '2015/10/15 END
      'Add by Amy 2019/07/22 零稅率
      Check1(1).Value = 0
      Check1(1).Tag = ""
      'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再設定
'      'Add by Amy 2019/07/25 電子發票寄送方式
'      Check4(0).Value = 0
'      Check4(1).Value = 0
'      Check4(0).Tag = ""
   End With
End Sub

'Public Sub Frmacc11p0_Delete()
Public Function Frmacc11p0_Delete() As Boolean
On Error GoTo Checking
   With Frmacc11p0
      If .textA4203.Enabled = False And Trim(.textA4203) <> "" And Trim(.textA4201) <> "" Then
      Else
         MsgBox "尚未查出欲刪除的資料 !", , MsgText(5)
         strControlButton = MsgText(602)
         Exit Function
      End If
      
      'Add By Sindy 2016/5/31 按刪除時, 先檢查 '會計備註' 及 'E-mail(財務)' 是否有值
      If Trim(Text1) <> "" Or Trim(textA4218) <> "" Then
         If MsgBox("此筆資料有會計備註或E-mail(財務)，請複製資料至客戶基本資料檔" & vbCrLf & _
                   "是否確定要刪除？", vbYesNo + vbCritical) = vbNo Then
            strControlButton = MsgText(602)
            Exit Function
         End If
      End If
      '2016/5/31 END
      
      adoTaie.Execute "delete from acc420 where a4201='" & ChgSQL(.textA4201) & "'"
      adoTaie.Execute "delete from acc490 where a4901='" & ChgSQL(.textA4201) & "'" 'Add By Sindy 2016/11/1 刪除會計師資料
      .AdodcRefresh
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
         .RecordShow
      Else
         StatusClear
      End If
   End With
Checking:
   If Err.Number = 0 Then
      Exit Function
   End If
   MsgBox Err.Description, , MsgText(5)
End Function

Public Sub Frmacc11p0_First()
   With Frmacc11p0
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveFirst
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11p0_Last()
   With Frmacc11p0
      If .Adodc1.Recordset.RecordCount <> 0 Then
         .Adodc1.Recordset.MoveLast
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11p0_Next()
   With Frmacc11p0
      If .Adodc1.Recordset.EOF = False Then
         .Adodc1.Recordset.MoveNext
         If .Adodc1.Recordset.EOF Then
            .Adodc1.Recordset.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11p0_Previous()
   With Frmacc11p0
      If .Adodc1.Recordset.BOF = False Then
         .Adodc1.Recordset.MovePrevious
         If .Adodc1.Recordset.BOF Then
            .Adodc1.Recordset.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         .FormShow
         .RecordShow
      End If
   End With
End Sub

Public Sub Frmacc11p0_Save()
Dim strText As String
Dim Cancel As Boolean
Dim strA4216 As String 'Add by Amy 2014/09/23
Dim strA4219 As String 'Add by Sindy 2015/6/2
Dim bolSendMail As Boolean, strNo As String 'Add By Sindy 2016/3/21
Dim strA4226 As String 'Add by Amy 2019/07/22 零稅率
Dim strA4227 As String 'Add by Amy 2019/07/25 電子發票寄送方式
Dim strA4220 As String 'Add by Amy 2025/02/20
   
   On Error GoTo Checking
   
   'Add by Amy 2023/06/30 避免存檔時,未Run到地址欄位_Validate未轉全型,故再轉一次
   If Trim(textA4215) <> MsgText(601) Then textA4215 = PUB_ChangeZIPToSir(textA4215) '營業地址
   If Trim(textA4203) <> MsgText(601) Then textA4203 = PUB_ChangeZIPToSir(textA4203) '郵寄地址
   If Trim(textA4222) <> MsgText(601) Then textA4222 = PUB_ChangeZIPToSir(textA4222) '特殊地址
   
   With Frmacc11p0
      If .textA4201 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .textA4201.SetFocus
         Exit Sub
      End If
      If .textA4202 = MsgText(601) Then
         'Add By Sindy 2015/7/13 個人或境外公司, 不必一定要輸入統一編號,
         '                       但其他情形則一定要輸入
         If Check1(0).Value = 0 And optCustomer(0).Value = False Then
         '2015/7/13 END
            MsgBox MsgText(10), , MsgText(5)
            strControlButton = MsgText(602)
            .textA4202.SetFocus
            Exit Sub
         End If
      'Add by Amy 2019/07/22 有統編不可勾零稅率
      ElseIf Check1(1).Value = 1 Then
            MsgBox "有統一編號不可勾零稅率", , MsgText(5)
            strControlButton = MsgText(602)
            .Check1(1).SetFocus
            Exit Sub
      End If
      If .textA4203 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .textA4203.SetFocus
         Exit Sub
      End If
      If .textA4206 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .textA4206.SetFocus
         Exit Sub
      End If
      If .textA4215 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .textA4215.SetFocus
         Exit Sub
      End If
      'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再設定
'      'Add by Amy 2019/07/25 電子發票寄送方式
'      If Check4(0).Value = vbChecked And Check4(1).Value = vbChecked Then
'            MsgBox "電子發票寄送方式只能擇一選擇", , MsgText(5)
'            .Check4(0).SetFocus
'            strControlButton = MsgText(602)
'            Exit Sub
'      End If
      'Add By Sindy 2016/11/07　繳款書寄件處不可空白
      If .textA4221 = MsgText(601) Then
         MsgBox MsgText(10), , MsgText(5)
         strControlButton = MsgText(602)
         .textA4221.SetFocus
         Exit Sub
      Else
         If .textA4221 = "3" Then '選擇3特殊時，特殊地址和收件人要同時有值
            If .textA4222 = MsgText(601) Or .textA4223 = MsgText(601) Then
               MsgBox MsgText(10), , MsgText(5)
               strControlButton = MsgText(602)
               If .textA4223 = MsgText(601) Then
                  .textA4223.SetFocus
               ElseIf .textA4222 = MsgText(601) Then
                  .textA4222.SetFocus
               End If
               Exit Sub
            End If
         Else
            If .textA4222 <> MsgText(601) Or .textA4223 <> MsgText(601) Then
               MsgBox "繳款書寄件處非選特殊，不需輸入特殊地址和收件人", , MsgText(5)
               strControlButton = MsgText(602)
               If .textA4223 <> MsgText(601) Then
                  .textA4223.SetFocus
               ElseIf .textA4222 <> MsgText(601) Then
                  .textA4222.SetFocus
               End If
               Exit Sub
            End If
         End If
      End If
      '2016/11/07 END
      Call textA4201_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4201.SetFocus
         Exit Sub
      End If
      Call textA4202_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4202.SetFocus
         Exit Sub
      End If
      Call textA4203_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4203.SetFocus
         Exit Sub
      End If
      Call textA4206_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4206.SetFocus
         Exit Sub
      End If
      Call textA4207_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4207.SetFocus
         Exit Sub
      End If
      Call textA4208_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4208.SetFocus
         Exit Sub
      End If
      Call textA4215_Validate(Cancel)
      If Cancel = True Then
         strControlButton = MsgText(602)
         .textA4215.SetFocus
         Exit Sub
      End If
      'Add By Sindy 2015/6/2
      Call textA4218_Validate(Cancel)
      If Cancel = True Then
         .textA4218.SetFocus
         strControlButton = MsgText(602)
         Exit Sub
      'Add By Sindy 2025/9/10 no轉大寫
      ElseIf UCase(.textA4218) = "NO" Then
         .textA4218 = "NO"
      '2025/9/10 END
      End If
      If Trim(textA4204) = "" Then
         MsgBox "電話不可以空白!!", vbExclamation
         textA4204.SetFocus
         strControlButton = MsgText(602)
         Exit Sub
      End If
      If optCustomer(0).Value = False And optCustomer(1).Value = False And _
         optCustomer(2).Value = False And optCustomer(3).Value = False Then
         MsgBox "個人或公司不可以空白!!", vbExclamation
         optCustomer(0).SetFocus
         strControlButton = MsgText(602)
         Exit Sub
      End If
      If GetTextLength(textA4201.Text) <= 6 Then
         If optCustomer(0).Value = False Then
            If MsgBox("確定不是個人嗎？", vbYesNo + vbCritical) = vbNo Then
               optCustomer(0).Value = True
               strControlButton = MsgText(602)
               Exit Sub
            End If
         End If
      Else
         If GetTextLength(textA4201.Text) >= 12 Then
            If optCustomer(0).Value = True Then
               If MsgBox("確定是個人嗎？", vbYesNo + vbCritical) = vbNo Then
                  optCustomer(0).Value = False
                  strControlButton = MsgText(602)
                  Exit Sub
               End If
            End If
         End If
      End If
      '2015/6/2 END
      'Add By Sindy 2016/11/7
      Call textA4221_Validate(Cancel)
      If Cancel = True Then
         .textA4221.SetFocus
         strControlButton = MsgText(602)
         Exit Sub
      End If
      Call textA4223_Validate(Cancel)
      If Cancel = True Then
         .textA4223.SetFocus
         strControlButton = MsgText(602)
         Exit Sub
      End If
      Call textA4222_Validate(Cancel)
      If Cancel = True Then
         .textA4222.SetFocus
         strControlButton = MsgText(602)
         Exit Sub
      End If
      '2016/11/7 END
      
      'Add by Sindy 2021/12/14 檢查畫面上的物件是否含有Unicode文字
      If PUB_ChkUniText(Me, True, True) = False Then
         Exit Sub
      End If

      'Add By Sindy 2016/6/8
      If strSaveConfirm = MsgText(3) Then '新增狀態時檢查
         '檢查客戶檔的中文和英文名稱是否有重覆,若有,彈詢問訊息,可以鍵入
         strExc(0) = "SELECT cu01,cu02,cu04,rtrim(upper(cu05||' '||cu88||' '||cu89||' '||cu90)),cu06" & _
                     " From customer" & _
                     " Where cu04='" & ChgSQL(textA4201.Text) & "'" & _
                     " or (rtrim(upper(cu05||' '||cu88||' '||cu89||' '||cu90))='" & ChgSQL(textA4201.Text) & "')" & _
                     " or cu06='" & ChgSQL(textA4201.Text) & "'" & _
                     " order by cu01,cu02"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If MsgBox("此客戶名稱已存在客戶檔，確定要儲存資料嗎？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
               textA4201.SetFocus
               strControlButton = MsgText(602)
               Exit Sub
            End If
         End If
      End If
      '2016/6/8 END
      
      'Add by Amy 2014/09/23 +境外公司及修正畫面所有含跳行符號的文字框
'      'Modify By Sindy 2014/10/15
'      'PUB_FilterFormText Me
'      Dim oObj As Object
'      For Each oObj In Me.Controls
'         If TypeName(oObj) = "TextBox" Then
'            If oObj.Text <> "" And oObj.Locked = False And oObj.Enabled = True And oObj.Name <> "Text1" Then
'               oObj.Text = PUB_StringFilter(oObj.Text)
'            End If
'         End If
'      Next
'      '2014/10/15 END
      'Modify By Sindy 2014/12/29 +不過濾的文字框.name
      PUB_FilterFormText Me, "Text1"
      '2014/12/29 END
      If .Check1(0).Value = vbChecked Then strA4216 = "Y"
      'Add by Amy 2019/07/22 零稅率
      If .Check1(1).Value = vbChecked Then strA4226 = "Y"
      'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再設定
'      'Add by Amy 2019/07/25 電子發票寄送方式 null-未設定
'      If Check4(0).Value = vbChecked Then strA4227 = "1" '紙本
'      If Check4(1).Value = vbChecked Then strA4227 = "2" '電子檔
      
      'Add By Sindy 2019/12/18 無每月提醒代填繳款書,代填方式則不須點選
      'Modify by Amy 2025/02/20 原Check3=每月代填繳款書,改可勾智慧所/法律所
      If Check3(0).Value = 0 And Check3(1).Value = 0 Then
         optA4228(0).Value = False
         optA4228(1).Value = False
      End If
      
      '更新DB資料
      adoTaie.BeginTrans
      'If .textA4201.Enabled = True Then '新增
      If optCustomer(0).Value = True Then
         strA4219 = "0"
      ElseIf optCustomer(1).Value = True Then
         strA4219 = "1"
      ElseIf optCustomer(2).Value = True Then
         strA4219 = "2"
      ElseIf optCustomer(3).Value = True Then
         strA4219 = "3"
      End If
      'Modify by Amy 2025/02/20 原Check3=每月代填繳款書(原:CNULL(IIf(Check3.Value = 1, "Y", "")),改可勾智慧所/法律所(存公司代碼)
      If Check3(0).Value = 1 Then strA4220 = strA4220 & ",1"
      If Check3(1).Value = 1 Then strA4220 = strA4220 & ",L"
      If strA4220 <> MsgText(601) Then strA4220 = Mid(strA4220, 2)
      'end 2025/02/20
      
      If strSaveConfirm = MsgText(3) Then '新增狀態時
         'Modify By Sindy 2014/10/15 +a4217
         'Modify By Sindy 2015/6/2 +a4218,a4219
         'Modify By Sindy 2016/11/4 +,a4220,a4221,a4222,a4223
         'Modify By Sindy 2017/3/16 +,a4224
         'Modify By Sindy 2017/3/24 +,a4225
         'Modify by Amy 2019/07/22 +,a4226 零稅率
         'Modify by Amy 2019/07/25 +,a4227 電子發票寄送方式
         'Modify By Sindy 2019/12/18 +,a4228 繳款書代填方式
         'Modify by Amy 2025/02/20 原Check3=每月代填繳款書(原:CNULL(IIf(Check3.Value = 1, "Y", ""))
         strSql = "insert into acc420(a4201,a4202,a4203,a4204,a4205,a4206,a4207,a4208,a4215,a4216,a4217,a4218,a4219,a4220,a4221,a4222,a4223,a4224,a4225,a4226,a4227,a4228) " & _
                  "values(" & CNULL(ChgSQL(.textA4201)) & "," & CNULL(.textA4202) & "," & CNULL(.textA4203) & _
                  "," & CNULL(.textA4204) & "," & CNULL(.textA4205) & "," & CNULL(.textA4206) & _
                  "," & CNULL(.textA4207) & "," & CNULL(.textA4208) & "," & CNULL(.textA4215) & _
                  "," & CNULL(strA4216) & "," & CNULL(ChgSQL(Text1)) & "," & CNULL(ChgSQL(textA4218)) & _
                  "," & CNULL(ChgSQL(strA4219)) & "," & CNULL(strA4220) & _
                  "," & CNULL(textA4221) & "," & CNULL(ChgSQL(textA4222)) & "," & CNULL(ChgSQL(textA4223)) & "," & CNULL(textA4224) & "," & CNULL(textA4225) & _
                  "," & CNULL(strA4226) & "," & CNULL(strA4227) & _
                  "," & CNULL(IIf(optA4228(0).Value = True, "1", IIf(optA4228(1).Value = True, "2", ""))) & ")"
         'Add By Sindy 2016/11/1 複製會計師資料
         If cmdA49.Tag = "複製" Then
            adoTaie.Execute strSql
            strSql = "insert into acc490(a4901,a4902,a4903,a4904,a4905,a4912,a4913,a4914) " & _
                     "values(" & CNULL(ChgSQL(.textA4201)) & "," & CNULL(strA4902) & "," & CNULL(strA4903) & _
                     "," & CNULL(strA4904) & "," & CNULL(strA4905) & "," & CNULL(strA4912) & _
                     "," & CNULL(strA4913) & "," & CNULL(strA4914) & ")"
         End If
         '2016/11/1 End
      Else '修改
         'Modify By Sindy 2014/10/15 +a4217
         'Modify By Sindy 2015/6/2 +a4218,a4219
         'Modify By Sindy 2017/3/16 +,a4224
         'Modify By Sindy 2017/3/24 +,a4225
         'Modify by Amy 2019/07/22 +,a4226 零稅率
         strExc(0) = ""
         If Check1(1).Tag <> strA4226 Then strExc(0) = ",a4226=" & CNULL(strA4226)
         'Mark by Amy 2024/06/13 目前由盟立下載電子檔,交由業務處理,故不需再設定
'         'Add by Amy 2019/07/25 電子發票寄送方式
'         If Val(Check4(0).Tag) <> Val(strA4227) Then strExc(0) = ",a4227=" & CNULL(strA4227)
         'Modify By Sindy 2019/12/18 +,a4228 繳款書代填方式
         'Modify by Amy 2025/02/20 原Check3=每月代填繳款書(原:CNULL(IIf(Check3.Value = 1, "Y", ""))
         strSql = "update acc420 " & _
                  "set a4202=" & CNULL(.textA4202) & _
                     ",a4203=" & CNULL(.textA4203) & _
                     ",a4204=" & CNULL(.textA4204) & _
                     ",a4205=" & CNULL(.textA4205) & _
                     ",a4206=" & CNULL(.textA4206) & _
                     ",a4207=" & CNULL(.textA4207) & _
                     ",a4208=" & CNULL(.textA4208) & _
                     ",a4215=" & CNULL(.textA4215) & _
                     ",a4216=" & CNULL(strA4216) & _
                     ",a4217=" & CNULL(ChgSQL(Text1)) & _
                     ",a4218=" & CNULL(ChgSQL(textA4218)) & _
                     ",a4219=" & CNULL(ChgSQL(strA4219)) & _
                     ",a4220=" & CNULL(strA4220) & _
                     ",a4221=" & CNULL(textA4221) & _
                     ",a4222=" & CNULL(ChgSQL(textA4222)) & _
                     ",a4223=" & CNULL(ChgSQL(textA4223)) & _
                     ",a4224=" & CNULL(textA4224) & _
                     ",a4225=" & CNULL(textA4225) & _
                     strExc(0) & _
                     ",a4228=" & CNULL(IIf(optA4228(0).Value = True, "1", IIf(optA4228(1).Value = True, "2", ""))) & _
                  " where a4201='" & ChgSQL(.textA4201) & "' "
         'end 2019/07/22
      End If
      adoTaie.Execute strSql
      
      'Add By Sindy 2016/2/15 境外公司欄由未勾選改為有勾選存檔時,
      If Check1(0).Tag = "" And Check1(0).Value = 1 Then
         bolSendMail = CU158isYUpdAccData(ChgSQL(textA4201), strNo)
      End If
      '2016/2/15 END
      
      adoTaie.CommitTrans
      'end 2014/09/23
      
      If cmdA49.Tag = "複製" Then cmdA49.Tag = "" 'Add By Sindy 2016/11/1
      
      'Add By Sindy 2016/3/21
      If bolSendMail = True Then
         PUB_SendMail strUserNum, strUserNum, "", textA4201 & "改為境外公司,過去年度之收據請再確認是否改為個人收據", "Dear Sirs," & vbCrLf & vbCrLf & strNo & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
      End If
      '2016/3/21 END
 
      .AdodcRefresh
      .FormDisabled
   End With
   
Checking:
   If Err.Number = 0 Then
      Exit Sub
   Else '-2147168237
      adoTaie.RollbackTrans
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Amy 2019/07/22 零稅率說明
Private Sub Lbl_Inf_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Lbl_Inf.ToolTipText = "符合條件：" & _
                        "(1)國外匯入款　(2)國外公司　(3)有簽約　就適用零稅率。" & _
                        "是否符合條件主要在於款項是否由國外匯入，故僅能在收到款項後設定。"
End Sub

'Add By Sindy 2025/9/5 財務信箱說明
Private Sub Lbl_InfMail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Lbl_InfMail.ToolTipText = "Key 「NO」不寄「收據」、「付款明細」、「催款單」、「不使用信箱」。"
End Sub

Private Sub lblCU168_Click(Index As Integer)
   Dim stComp As String, stMsg As String
   
   stComp = "1"
   If Index = 1 Then stComp = "L"
   If lblCU168(Index).BackColor = m_SetColor Then
      'Modify by Amy 2025/11/03 +txtNameNoUni,避免檔名有UniCode字無法開啟
      If ChkWithholdingTaxConsent(1, Me.Name, stComp, textA4201, File1, stMsg, txtNameNoUni) = False Then
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

Private Sub textA4201_GotFocus()
   OpenIme
   TextInverse textA4201
End Sub

Private Sub textA4201_Validate(Cancel As Boolean)
   'Added by Morgan 2014/12/23 剔除跳行符號
   textA4201.Text = PUB_StringFilter(textA4201.Text)
   'end 2014/12/23
   
   If textA4201.Enabled = False Then Exit Sub
   
   If textA4201.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textA4201, textA4201.MaxLength) Then
      Cancel = True
   End If
   
   If strSaveConfirm = MsgText(3) Then '新增狀態時,檢查是否有重覆
      If IsRecordExist(textA4201) = True Then
         MsgBox "該筆記錄已存在!!", , MsgText(5)
         textA4201.SetFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   
   'Modify By Sindy 2022/5/9 Ex: 美商宇心生醫股份有限公司”全型”台灣分公司
   '                             美商宇心生醫股份有限公司”半型”台灣分公司

   'a4201='" & ChgSQL(strKEY01) & "'" ==> replace(replace(a4201,' ',''),'　','')
   strSql = "SELECT * FROM acc420 " & _
            "WHERE replace(replace(a4201,' ',''),'　','')='" & ChgSQL(Replace(Replace(strKEY01, " ", ""), "　", "")) & "'"
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub textA4202_GotFocus()
   CloseIme
   TextInverse textA4202
End Sub

Private Sub textA4202_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textA4202_Validate(Cancel As Boolean)
Dim strTmp As String
   
   If textA4202.Enabled = False Then Exit Sub

   If textA4202.Text = "" Then Exit Sub
   If optCustomer(1).Value = True Then '公司才檢查
      If GetTextLength(textA4202.Text) <> 8 Then
         Call textA4202_GotFocus
         strTmp = "統編必須是8碼 ! 請確定 ?"
         If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
            Cancel = True
            Exit Sub
         End If
      End If
      If CheckID(1, textA4202.Text) = False Then
         strTmp = "統一編號錯誤，是否確定 ?"
         If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
            Cancel = True
         End If
      End If
   End If
End Sub

Private Sub textA4203_GotFocus()
   OpenIme
   TextInverse textA4203
End Sub

Private Sub textA4203_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textA4203) 'Moidfy by Amy 2023/04/20 ChangeZIP(KeyAscii)
End Sub

Private Sub textA4203_Validate(Cancel As Boolean)
   If textA4203.Enabled = False Then Exit Sub
   If textA4203.Text = "" Then Exit Sub
   
   'Add by Amy 2023/06/30 避免貼上的未轉全型 or KeyPress事件因輸入法失效沒執行到,故再轉一次
   textA4203 = PUB_ChangeZIPToSir(textA4203)
   
   If Not CheckLengthIsOK(textA4203, textA4203.MaxLength) Then
      Cancel = True
   End If
   
End Sub

Private Sub textA4204_GotFocus()
   CloseIme
   TextInverse textA4204
End Sub

Private Sub textA4205_GotFocus()
   CloseIme
   TextInverse textA4205
End Sub

Private Sub textA4206_GotFocus()
   CloseIme
   TextInverse textA4206
End Sub

Private Sub textA4206_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textA4206_Validate(Cancel As Boolean)
Dim strTmp As String, strTmp1 As String
   
   'If textA4206.Enabled = False Then Exit Sub 'Mark by Amy 2014/09/23
   
   Label30(2).Caption = ""
   If textA4206.Text = "" Then Exit Sub
   
   If PUB_GetStaffNameDept(textA4206.Text, strTmp, strTmp1, False) = True Then
      Label30(2).Caption = strTmp
   Else
      Cancel = True
   End If
   'add by sonia 2025/5/22 開放分所可新增但智權人員只能是該所人員
   If ProState = "2" Then
      If PUB_GetST06(strUserNum) <> "1" And PUB_GetST06(strUserNum) <> PUB_GetST06("" & textA4206) Then
         MsgBox "不可跨所新增收據抬頭資料！", , MsgText(5)
         Cancel = True
         If textA4206.Enabled = True Then textA4206.SetFocus
      End If
   End If
   'end 2025/5/22
End Sub

Private Sub textA4207_GotFocus()
   OpenIme
   TextInverse textA4207
End Sub

Private Sub textA4207_Validate(Cancel As Boolean)
   If textA4207.Enabled = False Then Exit Sub
   
   If textA4207.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textA4207, textA4207.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textA4208_GotFocus()
   OpenIme
   TextInverse textA4208
End Sub

Private Sub textA4208_Validate(Cancel As Boolean)
   If textA4208.Enabled = False Then Exit Sub
   
   If textA4208.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textA4208, textA4208.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textA4215_GotFocus()
   OpenIme
   TextInverse textA4215
End Sub

Private Sub textA4215_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textA4215) 'Modify by Amy 2023/04/20 原:ChangeZIP(KeyAscii)
End Sub

Private Sub textA4215_Validate(Cancel As Boolean)
   If textA4215.Enabled = False Then Exit Sub
   If textA4215.Text = "" Then Exit Sub
   
   'Add by Amy 2023/06/30 避免貼上的未轉全型 or KeyPress事件因輸入法失效沒執行到,故再轉一次
   textA4215 = PUB_ChangeZIPToSir(textA4215)
   
   If Not CheckLengthIsOK(textA4215, textA4215.MaxLength) Then
      Cancel = True
   End If
   
End Sub

'Add By Sindy 2015/6/2
Private Sub textA4218_GotFocus()
   CloseIme
   TextInverse textA4218
End Sub
Private Sub textA4218_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Email輸入字元檢查
End Sub
Private Sub textA4218_Validate(Cancel As Boolean)
   If textA4218.Enabled = False Then Exit Sub
   
   If textA4218.Text = "" Then Exit Sub
   Cancel = Not PUB_CheckMail(textA4218.Text)
End Sub
'2015/6/2 END

'Add By Sindy 2016/11/7
Private Sub textA4221_GotFocus()
   TextInverse textA4221
   CloseIme
End Sub
Private Sub textA4221_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
      KeyAscii = 0
   End If
End Sub
Private Sub textA4221_Validate(Cancel As Boolean)
   If textA4221.Text = "" Then Exit Sub
   If textA4221.Enabled = False Then Exit Sub
   If textA4221 = "2" Then
      If cmdA49.BackColor <> &HC0FFC0 Then
         MsgBox "無會計師資料,繳款書寄件處不可選擇2.會計師!!", vbCritical
         Cancel = True
      Else
         strExc(0) = "select A4902,A4912,A4913 from ACC490 where A4901='" & ChgSQL(textA4201) & "'"
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

'Add By Sindy 2016/11/4
Private Sub textA4222_GotFocus()
   OpenIme
   TextInverse textA4222
End Sub
Private Sub textA4222_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textA4222) 'Modify by Amy 2023/06/30 原:ChangeZIP(KeyAscii)
End Sub
Private Sub textA4222_Validate(Cancel As Boolean)
   If textA4222.Text = "" Then Exit Sub
   If textA4222.Enabled = False Then Exit Sub
   
   'Add by Amy 2023/06/30 避免貼上的未轉全型 or KeyPress事件因輸入法失效沒執行到,故再轉一次
   textA4222 = PUB_ChangeZIPToSir(textA4222)
   
   If Not CheckLengthIsOK(textA4222, textA4222.MaxLength) Then
      Cancel = True
   End If
   
End Sub
Private Sub textA4223_GotFocus()
   OpenIme
   TextInverse textA4223
End Sub
Private Sub textA4223_Validate(Cancel As Boolean)
   If textA4223.Enabled = False Then Exit Sub
   If textA4223.Text = "" Then Exit Sub
   '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textA4223, textA4223.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
'2016/11/4 END

'Add By Sindy 2016/11/29
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'Add By Sindy 2017/3/16
Private Sub textA4224_GotFocus()
   TextInverse textA4224
End Sub

'Add By Sindy 2017/3/16
Private Sub textA4224_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add By Sindy 2017/3/24
Private Sub textA4225_GotFocus()
   TextInverse textA4225
End Sub

'Add By Sindy 2017/3/24
Private Sub textA4225_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
