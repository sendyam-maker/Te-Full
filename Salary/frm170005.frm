VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm170005 
   BorderStyle     =   1  '單線固定
   Caption         =   "薪資異動資料"
   ClientHeight    =   5316
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8364
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5316
   ScaleWidth      =   8364
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   39
      Left            =   7335
      MaxLength       =   7
      TabIndex        =   15
      Text            =   "9999999"
      Top             =   2820
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Enabled         =   0   'False
      Height          =   270
      Index           =   38
      Left            =   5328
      MaxLength       =   7
      TabIndex        =   83
      Text            =   "9999999"
      Top             =   5070
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Height          =   270
      Index           =   37
      Left            =   3015
      MaxLength       =   80
      TabIndex        =   4
      Text            =   "123456"
      Top             =   1470
      Width           =   3375
   End
   Begin VB.TextBox txtSL 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   36
      Left            =   3015
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1680
      Width           =   285
   End
   Begin VB.TextBox txtSL 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   35
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "N"
      Top             =   1470
      Width           =   285
   End
   Begin VB.TextBox txtSL 
      Height          =   270
      Index           =   34
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   16
      Text            =   "1"
      Top             =   3150
      Width           =   285
   End
   Begin VB.TextBox txtSL 
      Height          =   270
      Index           =   33
      Left            =   1485
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "1"
      Top             =   2040
      Width           =   285
   End
   Begin VB.TextBox txtSL 
      Height          =   270
      Index           =   2
      Left            =   945
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "971226"
      Top             =   930
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   8
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   28
      Text            =   "9999999"
      Top             =   4785
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   7
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   25
      Text            =   "9999999"
      Top             =   4530
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Enabled         =   0   'False
      Height          =   270
      Index           =   6
      Left            =   7350
      MaxLength       =   7
      TabIndex        =   30
      Text            =   "9999999"
      Top             =   4785
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   4
      Left            =   7350
      MaxLength       =   5
      TabIndex        =   31
      Text            =   "99.99"
      Top             =   5070
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Enabled         =   0   'False
      Height          =   270
      Index           =   5
      Left            =   7350
      MaxLength       =   7
      TabIndex        =   27
      Text            =   "9999999"
      Top             =   4530
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   10
      Left            =   5328
      MaxLength       =   7
      TabIndex        =   29
      Text            =   "9999999"
      Top             =   4785
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   9
      Left            =   5328
      MaxLength       =   7
      TabIndex        =   26
      Text            =   "9999999"
      Top             =   4530
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   26
      Left            =   5328
      MaxLength       =   7
      TabIndex        =   24
      Text            =   "9999999"
      Top             =   3930
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   18
      Left            =   5328
      MaxLength       =   7
      TabIndex        =   14
      Text            =   "9999999"
      Top             =   2820
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   25
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   23
      Text            =   "9999999"
      Top             =   3930
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   24
      Left            =   7350
      MaxLength       =   7
      TabIndex        =   22
      Text            =   "9999999"
      Top             =   3660
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   23
      Left            =   5328
      MaxLength       =   7
      TabIndex        =   21
      Text            =   "9999999"
      Top             =   3660
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   21
      Left            =   2565
      MaxLength       =   7
      TabIndex        =   20
      Text            =   "9999999"
      Top             =   3660
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   22
      Left            =   7335
      MaxLength       =   7
      TabIndex        =   19
      Text            =   "9999999"
      Top             =   3390
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   20
      Left            =   5328
      MaxLength       =   7
      TabIndex        =   18
      Text            =   "9999999"
      Top             =   3390
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   19
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   17
      Text            =   "9999999"
      Top             =   3390
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   17
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   13
      Text            =   "9999999"
      Top             =   2820
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   16
      Left            =   7320
      MaxLength       =   7
      TabIndex        =   12
      Text            =   "9999999"
      Top             =   2550
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   15
      Left            =   5328
      MaxLength       =   7
      TabIndex        =   11
      Text            =   "9999999"
      Top             =   2550
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   13
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   10
      Text            =   "9999999"
      Top             =   2550
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   14
      Left            =   7320
      MaxLength       =   7
      TabIndex        =   9
      Text            =   "9999999"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   12
      Left            =   5328
      MaxLength       =   7
      TabIndex        =   8
      Text            =   "9999999"
      Top             =   2265
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   11
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   7
      Text            =   "9999999"
      Top             =   2265
      Width           =   735
   End
   Begin VB.TextBox txtSL 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   3
      Left            =   945
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "R"
      Top             =   1200
      Width           =   285
   End
   Begin VB.TextBox txtSL 
      Height          =   270
      Index           =   1
      Left            =   930
      MaxLength       =   6
      TabIndex        =   0
      Text            =   "123456"
      Top             =   630
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7605
      Top             =   -30
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170005.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170005.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170005.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170005.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170005.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170005.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170005.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170005.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170005.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170005.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170005.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   8364
      _ExtentX        =   14753
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblSD 
      Alignment       =   2  '置中對齊
      Caption         =   "Y"
      Height          =   180
      Index           =   48
      Left            =   2712
      TabIndex        =   85
      Top             =   5112
      Width           =   252
   End
   Begin MSForms.TextBox textCUID 
      Height          =   270
      Left            =   2610
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   630
      Width           =   5700
      VariousPropertyBits=   671105055
      Size            =   "10054;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "證照津貼："
      Height          =   180
      Index           =   9
      Left            =   6375
      TabIndex        =   84
      Top             =   2865
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "健保投保金額："
      Height          =   180
      Index           =   8
      Left            =   4056
      TabIndex        =   82
      Top             =   5100
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   180
      Index           =   7
      Left            =   2450
      TabIndex        =   81
      Top             =   1500
      Width           =   540
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   8100
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(若為是則只回寫勞健退資料且會即時更新回基本檔)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   3870
      TabIndex        =   79
      Top             =   1740
      Width           =   4440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否為調薪：         ( N:否 )"
      Height          =   180
      Index           =   4
      Left            =   30
      TabIndex        =   78
      Top             =   1500
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "新進或復職的異動會直接回存薪資基本資料，其他將依規則於每日批次更新！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Index           =   3
      Left            =   4005
      TabIndex        =   77
      Top             =   900
      Width           =   4260
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDsp 
      Caption         =   "台一國際專利法律事務所"
      Height          =   180
      Index           =   3
      Left            =   1800
      TabIndex        =   76
      Top             =   3180
      Width           =   4050
   End
   Begin VB.Label lblDsp 
      Caption         =   "台一國際專利商標事務所"
      Height          =   180
      Index           =   2
      Left            =   1800
      TabIndex        =   75
      Top             =   2070
      Width           =   4050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國　　籍："
      Height          =   180
      Index           =   15
      Left            =   6570
      TabIndex        =   73
      Top             =   1485
      Width           =   900
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "L  本國"
      Height          =   180
      Index           =   4
      Left            =   7500
      TabIndex        =   72
      Top             =   1485
      Width           =   555
   End
   Begin VB.Label lblSD 
      Alignment       =   2  '置中對齊
      Caption         =   "Y"
      Height          =   180
      Index           =   16
      Left            =   6525
      TabIndex        =   71
      Top             =   4290
      Width           =   255
   End
   Begin VB.Label lblSD 
      Alignment       =   2  '置中對齊
      Caption         =   "Y"
      Height          =   180
      Index           =   11
      Left            =   3735
      TabIndex        =   70
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "異動日期："
      Height          =   180
      Index           =   2
      Left            =   30
      TabIndex        =   69
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "勞健保是否以合夥人身分投保：     ( Y:是)"
      Height          =   180
      Index           =   26
      Left            =   1245
      TabIndex        =   68
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "喪事互助："
      Height          =   180
      Index           =   36
      Left            =   6420
      TabIndex        =   67
      Top             =   4830
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "扣繳項目"
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
      Index           =   27
      Left            =   135
      TabIndex        =   66
      Top             =   4320
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所得稅率："
      Height          =   180
      Index           =   25
      Left            =   6420
      TabIndex        =   65
      Top             =   5115
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "婚事互助："
      Height          =   180
      Index           =   24
      Left            =   6420
      TabIndex        =   64
      Top             =   4575
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "健  保  費："
      Height          =   180
      Index           =   23
      Left            =   4416
      TabIndex        =   63
      Top             =   4836
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "勞  保  費："
      Height          =   180
      Index           =   22
      Left            =   4416
      TabIndex        =   62
      Top             =   4572
      Width           =   900
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   90
      X2              =   8190
      Y1              =   4230
      Y2              =   4230
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "特殊勞保投保薪資："
      Height          =   180
      Index           =   38
      Left            =   930
      TabIndex        =   61
      Top             =   4575
      Width           =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "特殊健保投保薪資："
      Height          =   180
      Index           =   39
      Left            =   945
      TabIndex        =   60
      Top             =   4830
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "適用勞退新制：       ( Y:適用 )"
      Height          =   180
      Index           =   42
      Left            =   5250
      TabIndex        =   59
      Top             =   4290
      Width           =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "特殊退休金投保薪資："
      Height          =   180
      Index           =   41
      Left            =   3516
      TabIndex        =   58
      Top             =   3972
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "特殊退休金投保薪資："
      Height          =   180
      Index           =   40
      Left            =   3504
      TabIndex        =   57
      Top             =   2868
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "兼職人員資料以時薪輸入"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   37
      Left            =   6075
      TabIndex        =   55
      Top             =   2040
      Width           =   1980
   End
   Begin MSForms.Label lblName 
      Height          =   285
      Left            =   1710
      TabIndex        =   54
      Top             =   660
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1270;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      Caption         =   "第二家"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1065
      Left            =   120
      TabIndex        =   53
      Top             =   3150
      Width           =   345
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "第一家"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1065
      Left            =   120
      TabIndex        =   52
      Top             =   2010
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "特  支  費："
      Height          =   180
      Index           =   35
      Left            =   1650
      TabIndex        =   51
      Top             =   3975
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "房租津貼："
      Height          =   180
      Index           =   34
      Left            =   6420
      TabIndex        =   50
      Top             =   3705
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "差旅津貼："
      Height          =   180
      Index           =   33
      Left            =   4416
      TabIndex        =   49
      Top             =   3708
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "午餐津貼："
      Height          =   180
      Index           =   32
      Left            =   6405
      TabIndex        =   48
      Top             =   3435
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "技術津貼："
      Height          =   180
      Index           =   31
      Left            =   1650
      TabIndex        =   47
      Top             =   3705
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "職務津貼："
      Height          =   180
      Index           =   30
      Left            =   4416
      TabIndex        =   46
      Top             =   3432
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "基本薪資："
      Height          =   180
      Index           =   29
      Left            =   1650
      TabIndex        =   45
      Top             =   3435
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所得項目"
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
      Index           =   28
      Left            =   540
      TabIndex        =   44
      Top             =   3390
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公  司  別："
      Height          =   180
      Index           =   1
      Left            =   540
      TabIndex        =   43
      Top             =   3165
      Width           =   900
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   30
      X2              =   8160
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "特  支  費："
      Height          =   180
      Index           =   21
      Left            =   1650
      TabIndex        =   42
      Top             =   2865
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "房租津貼："
      Height          =   180
      Index           =   20
      Left            =   6375
      TabIndex        =   41
      Top             =   2595
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "差旅津貼："
      Height          =   180
      Index           =   19
      Left            =   4404
      TabIndex        =   40
      Top             =   2592
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "午餐津貼："
      Height          =   180
      Index           =   18
      Left            =   6375
      TabIndex        =   39
      Top             =   2310
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "技術津貼："
      Height          =   180
      Index           =   17
      Left            =   1650
      TabIndex        =   38
      Top             =   2595
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "職務津貼："
      Height          =   180
      Index           =   16
      Left            =   4404
      TabIndex        =   37
      Top             =   2316
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "基本薪資："
      Height          =   180
      Index           =   14
      Left            =   1650
      TabIndex        =   36
      Top             =   2310
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所得項目"
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
      Index           =   13
      Left            =   510
      TabIndex        =   35
      Top             =   2310
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公  司  別："
      Height          =   180
      Index           =   11
      Left            =   510
      TabIndex        =   34
      Top             =   2055
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "編　　制：        ( R:正  T:試  P:內兼  F:外兼 )"
      Height          =   180
      Index           =   10
      Left            =   30
      TabIndex        =   33
      Top             =   1215
      Width           =   3405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   32
      Top             =   675
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否為勞健退逕行調整或申訴獲准：         ( Y:是 )"
      Height          =   180
      Index           =   6
      Left            =   30
      TabIndex        =   80
      Top             =   1740
      Width           =   3840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "勞保是否無就保：       ( Y:無就保 )"
      Height          =   180
      Index           =   12
      Left            =   1245
      TabIndex        =   86
      Top             =   5112
      Width           =   2700
   End
End
Attribute VB_Name = "frm170005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/20 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2008/12/17
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_SL As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim iR(1 To 15) As String '勞保勞退健保費率資料
Dim m_bConfirmCheck As Boolean
Dim SL() As String
Dim m_bExportDoc As Boolean 'Added by Morgan 2017/11/20 匯出並EMail薪資調整表


Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   
   InitialField
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      Form_KeyDown vbKeyF2, 0
   End If
   SetInputEntry
   UpdateToolbarState
   
'Removed by Morgan 2013/1/21 sd11 改為 勞健保是否以合夥人身分投保
'   'Add by Morgan 2012/2/24 適用一般勞健保費率已無作用，改不顯示
'   Label1(26).Visible = False
'   lblSD(11).Visible = False
'end 2013/1/21

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170005 = Nothing
End Sub

Private Sub SetIR()
   strExc(0) = "select * from InsuranceRate"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      For intI = 1 To 15
         iR(intI) = "" & RsTemp.Fields("IR" & Format(intI, "00"))
      Next
   End If
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from SalaryLog where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_SL = .Fields.Count
      ReDim m_FieldList(TF_SL) As FIELDITEM
      ReDim SL(TF_SL) As String
      For Each oText In txtSL
         idx = oText.Index
         m_FieldList(idx).fiName = "SL" & Format(idx, "00")
         'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
         'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
            m_FieldList(idx).fiType = 0
         'Else
         '   m_FieldList(idx).fiType = 1
         'End If
         'end 2017/06/29
      Next
      End With
   End If
   SetIR
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
   
   Dim stKey01 As String
   Dim adoRst As New ADODB.Recordset
   
   stKey01 = DBDATE(txtSL(2)) & txtSL(1)
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM SalaryLog" & _
            " WHERE SL01 = '" & txtSL(1) & "'" & IIf(txtSL(2) = "", "", " and SL02 = " & DBDATE(txtSL(2))) & " order by sl02 desc"
      Case -2
         strExc(0) = "SELECT * FROM SalaryLog order by 2 ASC,1 ASC"
      Case -1
         strExc(0) = "SELECT * FROM SalaryLog" & _
            " WHERE SL02||SL01 <'" & stKey01 & "' order by 2 DESC,1 DESC"
      Case 1
         strExc(0) = "SELECT * FROM SalaryLog" & _
            " WHERE SL02||SL01 >'" & stKey01 & "' order by 2 ASC,1 ASC"
      Case 2
         strExc(0) = "SELECT * FROM SalaryLog order by 2 DESC,1 DESC"
   End Select
   intI = 1
   If adoRst.State = 1 Then adoRst.Close
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      UpdateCtrlData adoRst
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         MsgBox "查無資料！", vbInformation
         ClearField
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtSL(1).SetFocus
      txtSL_GotFocus 1
   End If
End Function

Private Sub txtSL_Change(Index As Integer)
   If Index = 1 Then
      If txtSL(Index) = "" Then
         lblName = "" 'Modify By Sindy 2021/12/20
      End If
      
   'Added by Morgan 2015/1/29
   '新增勞健保逕行調整時重算勞健保費
   ElseIf Index = 36 Then
      If m_EditMode = 1 Then
         SetInsureFee
      End If
   'end 2015/1/29
   End If
End Sub

Private Sub txtSL_GotFocus(Index As Integer)
   TextInverse txtSL(Index)
   CloseIme
   If Index = 37 Then OpenIme '2010/5/4 ADD BY SONIA
End Sub

Private Sub ClearField()
   lblName.Caption = Empty 'Modify By Sindy 2021/12/20
   For Each oText In txtSL
      oText.Text = Empty
      oText.Tag = Empty 'Add by Morgan 2011/10/25
   Next
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For Each oLabel In lblSD
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_SL
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   '是否為調薪預設 N
   txtSL(35) = "N"
   m_bConfirmCheck = False
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   With p_Rst
   If .RecordCount > 0 Then
      For Each oText In txtSL
         idx = oText.Index
         m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         '日期轉民國
         If idx = 2 Then
            If m_FieldList(idx).fiOldData <> "" Then
               oText.Text = Val(m_FieldList(idx).fiOldData) - 19110000
            End If
         Else
            oText.Text = m_FieldList(idx).fiOldData
         End If
         txtSL(oText.Index) = oText
      Next
      CUID(1) = "" & .Fields("sl27")
      CUID(2) = "" & .Fields("sl28")
      CUID(3) = "" & .Fields("sl29")
      CUID(4) = "" & .Fields("sl30")
      CUID(5) = "" & .Fields("sl31")
      CUID(6) = "" & .Fields("sl32")
      SetRefData
      txtSL(7).Tag = txtSL(7)
      txtSL(8).Tag = txtSL(8)
      txtSL(11).Tag = txtSL(11)
      txtSL(12).Tag = txtSL(12)
      txtSL(39).Tag = txtSL(39) 'Add By Sindy 2020/7/15
      txtSL(14).Tag = txtSL(14)
      
      '紀錄原始值
      For Each oText In txtSL
         SL(oText.Index) = oText
      Next
   End If
   End With
   UpdateCUID CUID, textCUID
   txtSL(1).Tag = txtSL(1)
   txtSL(2).Tag = txtSL(2)
   txtSL(33).Tag = txtSL(33) 'Added by Morgan 2020/4/13
   txtSL(34).Tag = txtSL(34) 'Added by Morgan 2020/4/13
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtSL
      oText.Locked = bLocked
   Next
   'Add by Morgan 2010/7/14 互助資料一律不可輸入,改只能由人事異動回寫
   'Removed by Morgan 2025/7/29 114/7/28起廢止婚喪互助辦法改直接設Enabled=False
   'txtSL(5).Locked = True
   'txtSL(6).Locked = True
   'end 2025/7/29
   'end 2010/7/14
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
         
      Case vbKeyEscape:
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction KeyCode
            End If
         End If
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If

   End Select
End Sub

' 執行指令
Public Sub OnAction(ByVal KeyCode As Integer)
   Dim bCancel As Boolean
   
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         m_EditMode = 1
         ClearField
         SetInputEntry
         UpdateToolbarState
         
      Case vbKeyF3 ' 修改
         'Add by Morgan 2010/6/2
         '薪資異動已更新回基本檔，不可再修改！
         'Modified by Morgan 2015/12/15 年終14,端午15,中秋16將加入
         'strExc(0) = "select 1 from dual where exists(select * from bookrecord where br01>=" & Left(DBDATE(txtSL(2)), 6) & ")"
         'Modified by Morgan 2017/9/5 Ex. A2033 1060701
         'strExc(0) = "select * from bookrecord where br01>=" & Left(DBDATE(txtSL(2)), 6) & " and substr(br01,-2)<13 and rownum<2"
         If txtSL(36) = "Y" Then
            strExc(1) = Left(DBDATE(txtSL(2)), 6)
         Else
            strExc(1) = CompDate(1, -1, txtSL(2)) \ 100
         End If
         strExc(0) = "select * from bookrecord where br01>=" & strExc(1) & " and substr(br01,-2)<13 and rownum<2"
         'end 2015/12/15
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If MsgBox("本次薪資異動已更新回基本檔，是否確定要修改！", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
               Exit Sub
            End If
         End If
         'end 2010/6/2
         m_EditMode = 2
         SetInputEntry
         UpdateToolbarState

      Case vbKeyF5 ' 刪除
         'Added by Morgan 2019/9/4
         '逕行調整的異動不可直接刪除,需由電腦中心還原薪資基本資料後刪除
         If txtSL(36) = "Y" Then
            MsgBox "本異動為勞健退逕行調整或申訴獲准，不可直接刪除！" & vbCrLf & "請通知電腦中心處理！(人工還原薪資基本資料並刪除)", vbCritical
            Exit Sub
            
         'Added by Morgan 2023/5/30 基本檔已更新就不可刪除 Ex: 111/5/1 A3013
         Else
            strExc(1) = CompDate(1, -1, txtSL(2)) \ 100
            strExc(0) = "select * from bookrecord where br01>=" & strExc(1) & " and substr(br01,-2)<13 and rownum<2"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "本次薪資異動已更新回基本檔，不可直接刪除！" & vbCrLf & "請通知電腦中心處理！(人工還原薪資基本資料並刪除)", vbCritical
               Exit Sub
            End If
         'end 2023/5/30
         End If
         'end 2019/9/4
         
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = False Then
                Exit Sub
            End If
            UpdateToolbarState
         End If
         
      Case vbKeyF4 ' 查詢
         m_EditMode = 4
         ClearField
         SetInputEntry
         UpdateToolbarState
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
      Case vbKeyF9 ' 確定
         If OnWork = False Then
            Exit Sub
         End If
         SetInputEntry
         UpdateToolbarState
         
      Case vbKeyF10 ' 取消
         bCancel = False
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  bCancel = True
               End If
            Case Else
               bCancel = True
         End Select
         If bCancel = True Then
            txtSL(1) = txtSL(1).Tag
            txtSL(2) = txtSL(2).Tag
            m_EditMode = 0
            SetInputEntry
            ShowRecord
            UpdateToolbarState
         End If
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate And txtSL(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtSL(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtSL(1) <> "" Then
            TBar1.Buttons(6).Enabled = True
            TBar1.Buttons(7).Enabled = True
            TBar1.Buttons(8).Enabled = True
            TBar1.Buttons(9).Enabled = True
         Else
            TBar1.Buttons(6).Enabled = False
            TBar1.Buttons(7).Enabled = False
            TBar1.Buttons(8).Enabled = False
            TBar1.Buttons(9).Enabled = False
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
      
      Case 1, 2, 3, 4 '維護
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1
         SetCtrlReadOnly False
         'Add by Morgan 2009/11/2
         '新增異動時公司別不可改
         'Modified by Morgan 2020/4/13 改開放修改，但加提醒
         'txtSL(33).Locked = True
         'txtSL(34).Locked = True
         'end 2020/4/13
         
         If Me.Visible = True Then
            txtSL(1).SetFocus
         End If
         
      Case 2
         SetCtrlReadOnly False
         txtSL(1).Locked = True
         txtSL(2).Locked = True
         If Me.Visible = True Then
            txtSL(3).SetFocus
         End If
         
         '公司別已輸入就不可改
         'Modified by Morgan 2020/4/13 改開放修改，但加提醒
'         If txtSL(33).Text <> "" Then
'            txtSL(33).Locked = True
'            txtSL(34).Locked = True
'
'            'Added by Morgan 2018/6/21 新進或復職可改公司別
'            strSql = "select sc03 from Staff_Change where sc01='" & txtSL(1) & "' and sc02=" & DBDATE(txtSL(2)) & " and sc03 IN ('01','02')"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               txtSL(33).Locked = False
'               txtSL(34).Locked = False
'            End If
'            'end 2018/6/21
'         End If
         'end 2020/4/13
         
         'Add by Morgan 2009/9/3
         If txtSL(36) <> "" Then
            txtSL(36).Locked = True
         End If
         
      Case 4
         SetCtrlReadOnly True
         txtSL(1).Locked = False
         txtSL(2).Locked = False
         If Me.Visible = True Then
            txtSL(1).SetFocus
         End If
      Case Else
         SetCtrlReadOnly True
         If Me.Visible = True Then
            txtSL(1).SetFocus
         End If
   End Select
   PUB_ChangeCaption Me, m_EditMode
End Sub

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            m_bExportDoc = False 'Added by Morgan 2017/11/20
            If AddRecord = True Then
               If m_bExportDoc Then PUB_ExportSalaryUpdate txtSL(1), txtSL(2), True 'Added by Morgan 2017/11/20
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            m_bExportDoc = False 'Added by Morgan 2017/11/20
            If ModRecord = True Then
               If m_bExportDoc Then PUB_ExportSalaryUpdate txtSL(1), txtSL(2), True 'Added by Morgan 2017/11/20
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 3: '刪除
         If DelRecord = True Then
            OnWork = True
            m_EditMode = 0
            ShowRecord 2
         End If
      
      Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtSL(1).SetFocus
               txtSL_GotFocus 1
            End If
         End If
         
   End Select
End Function


Private Function TxtValidate() As Boolean
   
   Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   If txtSL(1) = "" Then
      ShowMsg "請輸入員工代號 !"
      txtSL(1).SetFocus
      GoTo EscPoint
   End If
      
   For Each oText In txtSL
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtSL_Validate idx, bCancel
         If bCancel = True Then
            txtSL(idx).SetFocus
            txtSL_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   '維護
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If txtSL(2) = "" Then
         ShowMsg "請輸異動日期 !"
         txtSL(2).SetFocus
         GoTo EscPoint
      End If
      
      If txtSL(11) = "" Then
         MsgBox "第一家的基本薪資不可空白!"
         txtSL(11).SetFocus
         GoTo EscPoint
      End If
      
      If m_EditMode = 1 Then
         If txtSL(3) = "T" And txtSL(3) = SL(3) Then
            If MsgBox("該員工的原編制為【T】且新編制仍為【T】,是否確定要繼續?", vbYesNo + vbDefaultButton2) = vbNo Then
               txtSL(3).SetFocus
               txtSL_GotFocus 3
               GoTo EscPoint
            End If
         End If
         If CheckExists = True Then
            MsgBox "異動資料已存在 !"
            txtSL(1).SetFocus
            txtSL_GotFocus 1
            GoTo EscPoint
         End If
         
      End If 'Added by Morgan 2017/9/5 修改也要檢查
      
         strExc(1) = ""
         strExc(2) = ""
         For Each oText In txtSL
            'Modified by Morgan 2015/5/28 只需考慮薪資欄位
            'If oText.Index > 3 And oText.Index < 27 Then
            If oText.Index >= 11 And oText.Index <= 25 And oText.Index <> 18 Then
            'end 2015/5/28
               strExc(1) = Val(strExc(1)) + Val(oText)
               strExc(2) = Val(strExc(2)) + Val(SL(oText.Index))
            End If
            
         Next
         
         If strExc(1) <> strExc(2) Then
            If txtSL(35) = "N" Then
               'Added by Morgan 2015/5/5
               If txtSL(36) = "Y" Then
                  MsgBox "若為勞健保逕行調整不可調整薪資！", vbCritical
                  txtSL(36).SetFocus
                  GoTo EscPoint
               End If
               'end 2015/5/5
         
               If MsgBox("薪資總和有變動，確定不是調薪嗎?", vbYesNo + vbDefaultButton2) = vbNo Then
                  txtSL(35).SetFocus
                  txtSL_GotFocus 35
                  GoTo EscPoint
               End If
            End If
         Else
            If txtSL(35) = "" Then
               If MsgBox("薪資總和未變動，確定是調薪嗎?", vbYesNo + vbDefaultButton2) = vbNo Then
                  txtSL(35).SetFocus
                  txtSL_GotFocus 35
                  GoTo EscPoint
               End If
            End If
         End If
      
      'End If 'Removed by Morgan 2017/9/5 修改也要檢查
     
      'Add by Morgan 2009/9/4
      If txtSL(36) = "Y" Then
         If txtSL(35) <> "N" Then
            MsgBox "若為勞健保逕行調整或申訴獲准時不可為調薪！"
            txtSL(35).SetFocus
            GoTo EscPoint
         End If
      End If
   End If
   TxtValidate = True
   
EscPoint:
   m_bConfirmCheck = False
    
End Function

Private Function AddRecord() As Boolean
   Dim stCols As String, stValues As String, stSQL As String
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtSL
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO SalaryLog (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   
   '薪資基本檔資料現為 97.11.01(含)為止的最新的狀態，補之前的資料不需回寫，之後的才要。
   If Val(txtSL(2)) > 971101 Then
      
      'Modify by Morgan 2009/6/17 更新規則調整
      ''最後的異動才要回寫
      'stSQL = "Update SalaryData set (" & _
         " SD02,SD08,SD09,SD10,SD12,SD13,SD14,SD15" & _
         ",SD20,SD21,SD22,SD23,SD24,SD25,SD26,SD27" & _
         ",SD29,SD30,SD31,SD32,SD33,SD34,SD35,SD36)=(select" & _
         " SL03,SL04,SL05,SL06,SL07,SL08,SL09,SL10" & _
         ",SL11,SL12,SL13,SL14,SL15,SL16,SL17,SL18" & _
         ",SL19,SL20,SL21,SL22,SL23,SL24,SL25,SL26" & _
         " from salarylog s1 where SL01=SD01 AND SL02=" & m_FieldList(2).fiNewData & _
         ") where sd01='" & m_FieldList(1).fiNewData & "'" & _
         " and EXISTS (select 1 from salarylog where SL01=sd01 HAVING max(SL02)=" & m_FieldList(2).fiNewData & ")"
      'cnnConnection.Execute stSQL, intI
      
      UpdateSalaryData m_FieldList(1).fiNewData, m_FieldList(2).fiNewData, m_FieldList(36).fiNewData
      'end 2009/6/17
   End If
   
   cnnConnection.CommitTrans
   
   AddRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
    
End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE SalaryLog SET "
   stSet = ""
   For Each oText In txtSL
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where sl01='" & m_FieldList(1).fiOldData & "' and sl02=" & m_FieldList(2).fiOldData & "; end; "
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
      
      '薪資基本檔資料現為 97.11.01(含)為止的最新的狀態，補之前的資料不需回寫，之後的才要。
      If Val(txtSL(2)) > 971101 Then
      
         'Modify by Morgan 2009/6/17 更新規則調整
         ''最後的異動才要回寫
         'stSQL = "Update SalaryData set (" & _
            " SD02,SD08,SD09,SD10,SD12,SD13,SD14,SD15" & _
            ",SD20,SD21,SD22,SD23,SD24,SD25,SD26,SD27" & _
            ",SD29,SD30,SD31,SD32,SD33,SD34,SD35,SD36)=(select" & _
            " SL03,SL04,SL05,SL06,SL07,SL08,SL09,SL10" & _
            ",SL11,SL12,SL13,SL14,SL15,SL16,SL17,SL18" & _
            ",SL19,SL20,SL21,SL22,SL23,SL24,SL25,SL26" & _
            " from salarylog s1 where SL01=SD01 AND SL02=" & m_FieldList(2).fiNewData & _
            ") where sd01='" & m_FieldList(1).fiNewData & "'" & _
            " and EXISTS (select 1 from salarylog where SL01=sd01 HAVING max(SL02)=" & m_FieldList(2).fiNewData & ")"
         'cnnConnection.Execute stSQL, intI
         
         UpdateSalaryData m_FieldList(1).fiNewData, m_FieldList(2).fiNewData, m_FieldList(36).fiNewData
         'end 2009/6/17
      End If
   End If
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function
'Add by Morgan 2009/6/16 更新薪資基本檔或批次更新檔
Private Sub UpdateSalaryData(pSL01 As String, pSL02 As String, pSL36 As String)
   Dim bRealTimeUpdate As Boolean '是否即時更新
   Dim bRealTimeUpdate1 As Boolean '薪資是否即時更新
   Dim bRealTimeUpdate2 As Boolean '勞退是否即時更新
   Dim bRealTimeUpdate3 As Boolean '勞健保是否即時更新
   
   '最新異動才做
   strSql = "select 1 from salarylog where sL01='" & pSL01 & "' and SL02>" & pSL02
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 0 Then
   
      strSql = "delete SALARYUPDATE where su01='" & pSL01 & "' and su02=" & pSL02
      cnnConnection.Execute strSql, intI
         
      'Modify by Morgan 2009/9/3
      '若為勞健保逕行調整或申訴獲准則勞健保資料會即時更新回基本檔
      If pSL36 = "Y" Then
      
         'Added by Morgan 2015/3/25 尚未執行的舊異動要刪除否則資料可能會被覆蓋
         'Modified by Morgan 2016/3/1 勞退欄位有更新故也要刪
         strSql = "delete SALARYUPDATE where su01='" & pSL01 & "' and su02<" & pSL02 & " and su03 in ('2','3') and su05 is null"
         cnnConnection.Execute strSql, intI
         'end 2015/3/25
      
         'Modified by Morgan 2013/1/21 +SD47,,SL38
         'Modified by Morgan 2014/10/15 +勞退欄位 SD27,SD36,SD43,SD44
         'Modified by Sindy 2020/7/15 +NVL(SL39,0)
         strSql = "Update SalaryData set ( SD12,SD13,SD14,SD15,SD45,SD47,SD27,SD36,SD43,SD44)=(SELECT SL07,SL08,SL09,SL10,NVL(SL11,0)+NVL(SL12,0)+NVL(SL14,0)+NVL(SL39,0),SL38,SL18,SL26,NVL(SL11,0)+NVL(SL12,0)+NVL(SL14,0)+NVL(SL39,0),NVL(SL19,0)+NVL(SL20,0)+NVL(SL22,0) from salarylog where SL01=SD01 AND SL02=" & pSL02 & ") where sd01='" & pSL01 & "'"
         cnnConnection.Execute strSql, intI
         
         'Added by Morgan 2025/2/27
         '調整月份的月薪資投保金額也要更新
         strSql = "Update Salarymonth set sm42=(SELECT sd47 from salarydata where sd01=sm01) where sm01='" & pSL01 & "' and sm02=" & Left(pSL02, 6)
         cnnConnection.Execute strSql, intI
         'end 2025/2/27
      Else
      
         'Added by Morgan 2015/3/25 尚未執行的舊異動要刪除否則資料可能會被覆蓋
         'Removed by Morgan 2015/10/21 移到後面改加考慮是否有即時回寫
         'strSql = "delete SALARYUPDATE where su01='" & pSL01 & "' and su02<" & pSL02 & " and su05 is null"
         'cnnConnection.Execute strSql, intI
         'end 2015/10/21
         'end 2015/3/25
      
         '若是新進或復職時直接更新基本檔
         strSql = "select sc03 from Staff_Change where sc01='" & pSL01 & "'" & _
            " and sc02=" & pSL02 & " and sc03 IN ('01','02')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            '復職
            'Modify by Morgan 2010/6/25 若為復職勞健保改批次更新
            If RsTemp(0) = "02" Then
               bRealTimeUpdate1 = True
               bRealTimeUpdate2 = True
               bRealTimeUpdate3 = False
            '新進
            Else
               bRealTimeUpdate = True
            End If
         End If
         
         If bRealTimeUpdate Then
         
            'Added by Morgan 2015/10/21 刪除尚未執行的舊異動
            strSql = "delete SALARYUPDATE where su01='" & pSL01 & "' and su05 is null"
            cnnConnection.Execute strSql, intI
            'end 2015/10/21
         
            'Modify by Morgan 2010/7/14 婚喪扣款不更新改由維護員工基本資料及人事異動時回寫
            'Modified by Morgan 2013/1/21 +SD47,SL38
            'Modified by Morgan 2020/6/22 +SD52,SL39
            'Modified by Sindy 2020/7/15 +NVL(SL39,0)
            strSql = "Update SalaryData set (" & _
               " SD02,SD08,SD12,SD13,SD14,SD15" & _
               ",SD20,SD21,SD22,SD23,SD24,SD25,SD26,SD27" & _
               ",SD29,SD30,SD31,SD32,SD33,SD34,SD35,SD36" & _
               ",SD19,SD28,SD43,SD44,SD45,SD47,SD52)=(select" & _
               " SL03,SL04,SL07,SL08,SL09,SL10" & _
               ",SL11,SL12,SL13,SL14,SL15,SL16,SL17,SL18" & _
               ",SL19,SL20,SL21,SL22,SL23,SL24,SL25,SL26" & _
               ",SL33,SL34,NVL(SL11,0)+NVL(SL12,0)+NVL(SL14,0)+NVL(SL39,0),NVL(SL19,0)+NVL(SL20,0)+NVL(SL22,0),NVL(SL11,0)+NVL(SL12,0)+NVL(SL14,0)+NVL(SL39,0)" & _
               ",SL38,SL39 from salarylog s1 where SL01=SD01 AND SL02=" & pSL02 & _
               ") where sd01='" & pSL01 & "'"
            cnnConnection.Execute strSql, intI
         Else
            
            '1.薪資:
            If Not bRealTimeUpdate1 Then
               '若上月已入帳則即時更新,否則寫紀錄以批次更新(生效日=異動日)
               strExc(1) = CompDate(1, -1, pSL02) \ 100
               'Modified by Morgan 2015/12/15 年終14,端午15,中秋16將加入
               'strSql = "select * from bookrecord where br01>=" & strExc(1)
               strSql = "select * from bookrecord where br01>=" & strExc(1) & " and substr(br01,-2)<13 and rownum<2"
               'end 2015/12/15
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               '已入帳
               If intI = 1 Then
                  bRealTimeUpdate1 = True
               End If
            End If
            
            '薪資即時更新
            If bRealTimeUpdate1 Then
            
               'Added by Morgan 2015/10/21 刪除尚未執行的舊異動
               strSql = "delete SALARYUPDATE where su01='" & pSL01 & "' and su03='1' and su05 is null"
               cnnConnection.Execute strSql, intI
               'end 2015/10/21
            
               'Modify by Morgan 2010/7/14 婚喪扣款不更新改由維護員工基本資料及人事異動時回寫
               'Modify By Sindy 2020/6/22 +SD52,SL39
               strSql = "Update SalaryData set (" & _
                  " SD02,SD08,SD19,SD28" & _
                  ",SD20,SD21,SD22,SD23,SD24,SD25,SD26" & _
                  ",SD29,SD30,SD31,SD32,SD33,SD34,SD35,SD52)=(select" & _
                  " SL03,SL04,SL33,SL34" & _
                  ",SL11,SL12,SL13,SL14,SL15,SL16,SL17" & _
                  ",SL19,SL20,SL21,SL22,SL23,SL24,SL25,SL39" & _
                  " from salarylog s1 where SL01=SD01 AND SL02=" & pSL02 & _
                  ") where sd01='" & pSL01 & "'"
               cnnConnection.Execute strSql, intI
            '入帳後更新
            Else
               strSql = "insert into SALARYUPDATE (su01,su02,su03,su04)" & _
                  " values ('" & pSL01 & "'," & pSL02 & ",'1' ," & pSL02 & ")"
               cnnConnection.Execute strSql, intI
            End If
            
            '2.勞退:
            '勞退即時更新
            If bRealTimeUpdate2 Then
            
               'Added by Morgan 2015/10/21 刪除尚未執行的舊異動
               strSql = "delete SALARYUPDATE where su01='" & pSL01 & "' and su03='2' and su05 is null"
               cnnConnection.Execute strSql, intI
               
               'Modified by Sindy 2020/7/15 +NVL(SL39,0)
               strSql = "Update SalaryData set (" & _
                  " SD27,SD36,SD43,SD44)=(select" & _
                  " SL18,SL26,NVL(SL11,0)+NVL(SL12,0)+NVL(SL14,0)+NVL(SL39,0),NVL(SL19,0)+NVL(SL20,0)+NVL(SL22,0)" & _
                  " from salarylog s1 where SL01=SD01 AND SL02=" & pSL02 & _
                  ") where sd01='" & pSL01 & "'"
               cnnConnection.Execute strSql, intI
            Else
               '寫紀錄以批次更新(生效日=異動日的隔月1號)
               strExc(1) = CompDate(1, 1, pSL02)
               strExc(1) = strExc(1) \ 100 & "01"
               strSql = "insert into SALARYUPDATE (su01,su02,su03,su04)" & _
                  " values ('" & pSL01 & "'," & pSL02 & ",'2' ," & strExc(1) & ")"
               cnnConnection.Execute strSql, intI
               
               'Added by Morgan 2017/11/29 當月最後3個工作天的異動若勞退級距有變動要製作調整表
               If Left(pSL02, 6) = Left(strSrvDate(1), 6) Then
                  'Modified by Morgan 2023/5/29
                  'strExc(0) = CompWorkDay(3, CompDate(2, -1, strExc(1)), 1)
                  intI = Val(Pub_GetSpecMan("SUWD"))
                  If intI > 0 Then
                     strExc(0) = CompWorkDay(intI, CompDate(2, -1, strExc(1)), 1)
                  'end 2023/5/26
                     If strSrvDate(1) >= strExc(0) Then
                        m_bExportDoc = True
                     End If
                  End If 'Added by Morgan 2023/5/29
               End If
               'end 2017/11/29
            End If
            
            '3.勞健保:
            If bRealTimeUpdate3 Then
            
               'Added by Morgan 2015/10/21 刪除尚未執行的舊異動
               strSql = "delete SALARYUPDATE where su01='" & pSL01 & "' and su03='3' and su05 is null"
               cnnConnection.Execute strSql, intI
               
               'Modified by Morgan 2013/1/21 +SD47,SL38
               'Modified by Sindy 2020/7/15 +NVL(SL39,0)
               strSql = "Update SalaryData set (" & _
                  " SD12,SD13,SD14,SD15,SD45,SD47)=(select" & _
                  " SL07,SL08,SL09,SL10,NVL(SL11,0)+NVL(SL12,0)+NVL(SL14,0)+NVL(SL39,0)" & _
                  ",SL38 from salarylog s1 where SL01=SD01 AND SL02=" & pSL02 & _
                  ") where sd01='" & pSL01 & "'"
               cnnConnection.Execute strSql, intI
            Else
               '寫紀錄以批次更新(生效日=下一個3月或9月)
               'Modify by Morgan 2009/8/21 2-7月的異動9月更新,8-1月的異動3月更新
               strExc(1) = Right(pSL02, 4)
               If Val(strExc(1)) < 200 Then
                  strExc(1) = pSL02 \ 10000 & "0301"
               ElseIf Val(strExc(1)) < 800 Then
                  strExc(1) = pSL02 \ 10000 & "0901"
               Else
                  strExc(1) = pSL02 \ 10000 + 1 & "0301"
               End If
               strSql = "insert into SALARYUPDATE (su01,su02,su03,su04)" & _
                  " values ('" & pSL01 & "'," & pSL02 & _
                  ",'3' ," & strExc(1) & ")"
               cnnConnection.Execute strSql, intI
            End If
         End If
      End If
   End If
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: OnAction vbKeyF5
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

Private Sub UpdateFieldNewData()
   For Each oText In txtSL
      idx = oText.Index
      If idx = 2 Then
         '年月轉西元
         m_FieldList(idx).fiNewData = DBDATE(oText.Text)
      Else
         m_FieldList(idx).fiNewData = oText.Text
      End If
   Next
End Sub

Private Sub txtSL_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 1, 33, 34, 37
         '不控制
      Case 3
         'Modify by Morgan 2009/11/24 取消 S
         'If KeyAscii <> 8 And KeyAscii <> Asc("R") And KeyAscii <> Asc("T") And KeyAscii <> Asc("S") And KeyAscii <> Asc("P") And KeyAscii <> Asc("F") Then
         If KeyAscii <> 8 And KeyAscii <> Asc("R") And KeyAscii <> Asc("T") And KeyAscii <> Asc("P") And KeyAscii <> Asc("F") Then
            KeyAscii = 0
            Beep
         End If
      'Added by Morgan 2025/6/30 所得稅率可以輸小數
      Case 4
         KeyAscii = Pub_NumAscii(KeyAscii, True)
      'end 2025/6/30
      Case 35
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
      
      Case 36
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
            Beep
         End If
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtSL_Validate(Index As Integer, Cancel As Boolean)
   If m_EditMode = 1 Or m_EditMode = 2 Then
   
      Select Case Index
         Case 1
            If txtSL(Index) <> "" Then
               If ChkStaffID(txtSL(Index)) = True Then
                  Cancel = True
               End If
               If m_EditMode = 1 Then
                  If Left(txtSL(Index), 1) = "F" Then
                     MsgBox "不可輸入翻譯人員 !"
                     Cancel = True
                  End If
               End If
               If SetRefData() = False Then
                  MsgBox "員工代碼輸入錯誤！"
                  Cancel = True
               End If
            End If
            
         Case 2
            If txtSL(Index) <> "" Then
               If ChkDate(txtSL(Index)) = False Then
                  Cancel = True
                  
               ElseIf m_EditMode = 1 Then
                  If DateDiff("m", Format(DBDATE(txtSL(Index)), "####/##/##"), Format(strSrvDate(1), "####/##/##")) > 1 Then
                     MsgBox "異動日期不可早於上月份!"
                     Cancel = True
                  'Modify by Morgan 2009/10/27 已改寫為批次更新不必再限制！
                  'ElseIf Val(txtSL(Index)) \ 100 > Val(strSrvDate(2)) \ 100 Then
                  '   MsgBox "異動日期不可大於當月份!"
                  '   Cancel = True
                  End If
               End If
            End If

               
         Case 4
            If Val(txtSL(Index)) > 99 Then
               Cancel = True
               MsgBox "所得稅率不可大於 99！"
            End If
               
         Case 7 '勞保投保薪資
            If txtSL(Index) <> txtSL(Index).Tag Then
               SetInsureFee 1
            End If
            txtSL(Index).Tag = txtSL(Index)
               
         Case 8 '健保投保薪資
            If txtSL(Index) <> txtSL(Index).Tag Then
               SetInsureFee 2
            End If
            txtSL(Index).Tag = txtSL(Index)
               
         Case 11, 12, 14
            'Modified by Morgan 2013/1/21 sd11 改為 勞健保是否以合夥人身分投保
            'If lblSD(11) = "" And (txtSL(7) = "" Or txtSL(8) = "") Then
            If Left(txtSL(1), 1) <> "F" And (txtSL(7) = "" Or txtSL(8) = "") Then
            'end 2013/1/21
               'Modify By Sindy 2020/7/15 + Val(txtSL(39))
               strExc(1) = Val(txtSL(11)) + Val(txtSL(12)) + Val(txtSL(14)) + Val(txtSL(39))
               strExc(2) = Val(txtSL(11).Tag) + Val(txtSL(12).Tag) + Val(txtSL(14).Tag) + Val(txtSL(39).Tag)
               If strExc(1) <> strExc(2) Then
                  If txtSL(7) = "" Then
                     SetInsureFee 1
                  End If
                  If txtSL(8) = "" Then
                     SetInsureFee 2
                  End If
               End If
            End If
            txtSL(Index).Tag = txtSL(Index)
               
         Case 33
            If txtSL(Index) <> "" Then
               lblDsp(2) = CompNameQuery(txtSL(Index))
               If lblDsp(2) = "" Then
                  ShowMsg "公司別錯誤 !"
                  Cancel = True
               End If
            'Added by Morgan 2020/4/13
            ElseIf txtSL(33).Tag <> "" And txtSL(33).Tag <> txtSL(33) Then
                  If MsgBox("是否確定要更改公司別？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                     Cancel = True
                  End If
            'end 2020/4/13
            End If
         Case 34
            If txtSL(Index) <> "" Then
               lblDsp(3) = CompNameQuery(txtSL(Index))
               If lblDsp(3) = "" Then
                  ShowMsg "公司別錯誤 !"
                  Cancel = True
               End If
            'Added by Morgan 2020/4/13
            ElseIf txtSL(34).Tag <> "" And txtSL(34).Tag <> txtSL(34) Then
                  If MsgBox("是否確定要更改公司別？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                     Cancel = True
                  End If
            'end 2020/4/13
            End If
            
         '2010/5/4 ADD BY SONIA
         Case 37
            If txtSL(Index) <> "" Then
               If Not CheckLengthIsOK(txtSL(Index), txtSL(Index).MaxLength) Then
                  Cancel = True
               End If
            End If
         '2010/5/4 END
            
      End Select
      
      If Cancel = True Then TextInverse txtSL(Index)
      
      '若是按確定的檢查時略過
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
            Case 1
               If txtSL(Index) <> "" Then
                  '載入最新薪資資料
                  If m_EditMode = 1 Then
                     If LoadSalaryData = False Then
                        MsgBox "薪資基本資料讀入失敗 !"
                     End If
                  End If
               End If
               
            Case 2
               If Right(txtSL(Index), 4) = "0501" Or Right(txtSL(Index), 4) = "1101" Then
                  txtSL(35) = ""
               End If
               
            Case 3
               If txtSL(3) = "R" And SL(3) = "T" Then
                  txtSL(35) = ""
               End If

         End Select
      End If
   End If
End Sub

'設定勞健保費
'iOption:0=全部,1=勞保費,2=健保費
Private Sub SetInsureFee(Optional iOption As Integer)
   Dim lngInsureSalary As Long '投保薪資
   Dim lngInsureBase As Long '投保等級
   Dim dblInsureRate As Double '投保費率
   Dim dblFreeRate As Double '補助比率
   Dim dblInsureRate2 As Double '就業保險費率
   Dim intShareRate As Integer '負擔比例
   
   'Added by Morgan 2013/11/28
   '員工編號第4碼為9者除外 Ex.68099
   'Modified by Morgan 2024/4/30 +合約翻譯人員(第5碼>='A')除外
   If Mid(txtSL(1), 4, 1) = "9" And Right(txtSL(1), 1) < "A" Then
      Exit Sub
   End If
   
   'Modified by Morgan 2013/1/21 sd11 改為 勞健保是否以合夥人身分投保
   'If lblSD(11) = "" Then '適用一般勞健保費率
      
      If iOption = 0 Or iOption = 1 Then
         'Added by Morgan 2015/6/22
         If Val(txtSL(7)) = 0 And Trim(txtSL(7)) <> "" Then
            txtSL(9) = 0
         Else
         'end 2015/6/22
         
            '勞保投保薪資
            'Modified by Morgan 2016/3/31 特殊勞保投保薪資會輸0(63001,已退休)
            'lngInsureSalary = Val(txtSL(7))
            'If lngInsureSalary = 0 Then
            If txtSL(7) <> "" Then
               lngInsureSalary = Val(txtSL(7))
            Else
            'end 2016/3/31
               'Modify By Sindy 2020/7/15 + Val(txtSL(39))
               lngInsureSalary = Val(txtSL(11)) + Val(txtSL(12)) + Val(txtSL(14)) + Val(txtSL(39))
            End If
            '勞保等級
            lngInsureBase = PUB_GetInsureBase(lngInsureSalary, "L")
            
            'Modify by Morgan 2009/6/29
            '98/5/1 起外國人也有失業給付,費率改與本國人同,只有雇主(所長)沒有(未來加65歲以上)
            'If Left(lblDsp(4), 1) = "F" Then
            'Modify by Morgan 2010/10/26 勞保費率及就業保險費率需個別計算(四捨五入)
            'Modified by Morgan 2013/1/21 改判斷 sd11 勞健保是否以合夥人身分投保
            'If CheckExceptLiRate = True Then
            'Modified by Morgan 2023/6/29 +判斷 sd48勞保是否無就保
            If lblSD(11) = "Y" Or lblSD(48) = "Y" Then
            'end 2013/1/21
               'dblInsureRate = Val(iR(2))
               dblInsureRate = Val(iR(1))
               dblInsureRate2 = 0 'Add
            
            'Added by Morgan 2015/1/28
            '超過65歲也沒有就保
            ElseIf PUB_ChkOver65(txtSL(1)) Then
               dblInsureRate = Val(iR(1))
               dblInsureRate2 = 0
            'end 2015/1/28
            
            Else
               dblInsureRate = Val(iR(1))
               dblInsureRate2 = Val(iR(2)) 'Add
            End If
            '勞保費=勞保等級*勞保費率*勞保個人負擔比例
            'txtSL(9) = Round(lngInsureBase * dblInsureRate / 100 * Val(iR(3)) / 100, 0)
            txtSL(9) = Round(lngInsureBase * dblInsureRate / 100 * Val(iR(3)) / 100, 0) + Round(lngInsureBase * dblInsureRate2 / 100 * Val(iR(3)) / 100, 0)
            'end 2010/10/27
         
         End If
      End If
      
      If iOption = 0 Or iOption = 2 Then
         'Added by Morgan 2015/6/22
         If Val(txtSL(8)) = 0 And Trim(txtSL(8)) <> "" Then
            txtSL(10) = 0
            txtSL(38) = 0
         Else
         'end 2015/6/22
         
            '健保投保薪資
            'Modified by Morgan 2016/3/31 與勞保檢查一致
            'lngInsureSalary = Val(txtSL(8))
            'If lngInsureSalary = 0 Then
            If txtSL(8) <> "" Then
               lngInsureSalary = Val(txtSL(8))
            Else
            'end 2016/3/31
               'Modify By Sindy 2020/7/15 + Val(txtSL(39))
               lngInsureSalary = Val(txtSL(11)) + Val(txtSL(12)) + Val(txtSL(14)) + Val(txtSL(39))
            End If
            dblInsureRate = Val(iR(6))
            
            'Added by Morgan 2013/1/21
            '以合夥人身分投保 100% 個人負擔
            If lblSD(11) = "Y" Then
               intShareRate = 100
            Else
               intShareRate = Val(iR(7))
            End If
            'end 2013/1/21
            
            '健保費=健保等級*健保費率*健保個人負擔比例
            'Modify by Morgan 2010/4/15 健保費調整改用共用函數
            'lngInsureBase = GetInsureBase(lngInsureSalary, "H") '健保等級
            'txtSL(10) = Round(lngInsureBase * dblInsureRate / 100 * Val(IR(7)) / 100, 0)
            lngInsureBase = PUB_GetInsureBase(lngInsureSalary, "H", dblFreeRate) '健保等級
            'Modified by Morgan 2013/1/21
            'txtSL(10) = PUB_GetHIFee(lngInsureBase, dblInsureRate, Val(IR(7)), dblFreeRate)
            txtSL(10) = PUB_GetHIFee(lngInsureBase, dblInsureRate, intShareRate, dblFreeRate)
            txtSL(38) = lngInsureBase
            'end 2013/1/21
            
         End If
      End If
      
   'End If
End Sub

'Removed by Morgan 2013/1/21
''Add by Morgan 2009/6/23
''檢查勞保費是否特別,現在只有雇主(所長)
'Private Function CheckExceptLiRate() As Boolean
'   If txtSL(1) <> "" Then
'      strExc(0) = "select st02 from staff where st01='" & txtSL(1) & "' and st20='11'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         CheckExceptLiRate = True
'      End If
'   End If
'End Function

Private Function SetRefData() As Boolean
   
   'Modified by Morgan 2023/6/29 +sd48
   strExc(0) = "select st02,st04,a0902,ac03,st24,sd11,sd16,sd19,sd48,a1.a0802 comp1,sd28,a2.a0802 comp2" & _
      " from staff,acc090,allcode,salarydata,acc080 a1,acc080 a2" & _
      " where st01='" & txtSL(1) & "' and a0901(+)=st03 and ac01(+)='01' and ac02(+)=st20" & _
      " and sd01(+)=st01 and a1.a0801(+)=sd19 and a2.a0801(+)=sd28"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      '名稱
      lblName = "" & .Fields("st02")
      '部門
      lblDsp(2) = "" & .Fields("a0902")
      '職稱
      lblDsp(3) = "" & .Fields("ac03")
      '國籍
      lblDsp(4) = "" & .Fields("st24")
      If lblDsp(4) = "L" Then
         lblDsp(4) = lblDsp(4) & " 本國"
      ElseIf lblDsp(4) = "F" Then
         lblDsp(4) = lblDsp(4) & " 外國"
      End If
      
      lblSD(11) = "" & .Fields("sd11")
      lblSD(16) = "" & .Fields("sd16") '適用勞退新制
      lblSD(48) = "" & .Fields("sd48") 'Added by Morgan 2023/6/29 勞保是否無就保
      lblDsp(2) = "" & .Fields("comp1")
      lblDsp(3) = "" & .Fields("comp2")
      If m_EditMode = "1" And "" & .Fields("st04") = "2" Then
         MsgBox "該員工已離職!"
         Exit Function
      End If
      End With
      SetRefData = True
   End If
   
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '刪除
   stSQL = "delete from SalaryLog where sl01='" & m_FieldList(1).fiOldData & "' and sl02=" & m_FieldList(2).fiOldData
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   'Add by Morgan 2009/6/17 刪除未更新的待更新記錄
   stSQL = "delete from SalaryUpdate where su01='" & m_FieldList(1).fiOldData & "' and su02=" & m_FieldList(2).fiOldData & " and su05 is null"
   cnnConnection.Execute stSQL, intI
   'end 2009/6/17
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtSL(1).Tag = ""
   txtSL(2).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

'Remove by Morgan 2010/4/14 改抓公用函數
'Private Function GetInsureBase(pInsureSalary As Long, pKind As String) As Long
'   strExc(0) = "select si02 from SalaryInsurance" & _
'      " where si01='" & pKind & "' and si03<=" & pInsureSalary & " and si04>=" & pInsureSalary
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      GetInsureBase = Val("" & RsTemp.Fields(0))
'   End If
'End Function

Private Function CheckExists() As Boolean
   CheckExists = True
   strExc(0) = "select 1 from SalaryLog where sl01='" & txtSL(1) & "' and sl02=" & DBDATE(txtSL(2))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      CheckExists = False
   End If
End Function

Private Function LoadSalaryData() As Boolean
   strExc(0) = "select * from salarydata where sd01='" & txtSL(1) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      txtSL(3) = "" & .Fields("sd02")
      txtSL(4) = "" & .Fields("sd08")
      
      'Modify by Morgan 2010/7/2 改到人事異動更新
      ''Modify by Morgan 2010/7/1 婚喪互助要重新抓
      ''txtSL(5) = "" & .Fields("sd09")
      ''txtSL(6) = "" & .Fields("sd10")
      'If PUB_GetHelpFee(txtSL(1), strExc(1)) = True Then
      '   txtSL(5) = strExc(1) '婚事互助
      '   txtSL(6) = txtSL(5) '喪事互助
      'End If
      ''end 2010/7/1
      'Removed by Morgan 2025/7/29 114/7/28起廢止婚喪互助辦法
      'txtSL(5) = "" & .Fields("sd09")
      'txtSL(6) = "" & .Fields("sd10")
      'end 2025/7/29
      'end 2010/7/2
      
      txtSL(7) = "" & .Fields("sd12")
      txtSL(8) = "" & .Fields("sd13")
      txtSL(9) = "" & .Fields("sd14")
      txtSL(10) = "" & .Fields("sd15")
      txtSL(11) = "" & .Fields("sd20")
      txtSL(12) = "" & .Fields("sd21")
      txtSL(13) = "" & .Fields("sd22")
      txtSL(14) = "" & .Fields("sd23")
      txtSL(15) = "" & .Fields("sd24")
      txtSL(16) = "" & .Fields("sd25")
      txtSL(17) = "" & .Fields("sd26")
      txtSL(18) = "" & .Fields("sd27")
      txtSL(19) = "" & .Fields("sd29")
      txtSL(20) = "" & .Fields("sd30")
      txtSL(21) = "" & .Fields("sd31")
      txtSL(22) = "" & .Fields("sd32")
      txtSL(23) = "" & .Fields("sd33")
      txtSL(24) = "" & .Fields("sd34")
      txtSL(25) = "" & .Fields("sd35")
      txtSL(26) = "" & .Fields("sd36")
      txtSL(33) = "" & .Fields("sd19")
      txtSL(33).Tag = txtSL(33) 'Added by Morgan 2020/4/13
      txtSL(34) = "" & .Fields("sd28")
      txtSL(34).Tag = txtSL(34) 'Added by Morgan 2020/4/13
      txtSL(38) = "" & .Fields("sd47") 'Added by Morgan 2013/1/21
      txtSL(39) = "" & .Fields("sd52") 'Added by Sindy 2020/6/22
      End With
      
      '紀錄原始值
      For Each oText In txtSL
         SL(oText.Index) = oText
      Next
      
      LoadSalaryData = True
   End If
End Function

