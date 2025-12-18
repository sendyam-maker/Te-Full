VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm170021 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月薪資資料"
   ClientHeight    =   5544
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8340
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5544
   ScaleWidth      =   8340
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   45
      Left            =   4500
      MaxLength       =   8
      TabIndex        =   14
      Text            =   "99999999"
      Top             =   2520
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   26
      Left            =   6765
      MaxLength       =   8
      TabIndex        =   78
      Text            =   "99999999"
      Top             =   5100
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   44
      Left            =   2250
      MaxLength       =   3
      TabIndex        =   77
      Text            =   "6"
      Top             =   4050
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   43
      Left            =   6750
      MaxLength       =   8
      TabIndex        =   74
      Text            =   "99999999"
      Top             =   3765
      Width           =   915
   End
   Begin VB.CommandButton cmdHi 
      Caption         =   "健保費明細"
      Height          =   315
      Left            =   6480
      TabIndex        =   73
      Top             =   1200
      Width           =   1635
   End
   Begin VB.TextBox txtSM 
      Height          =   270
      Index           =   37
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "1"
      Top             =   1320
      Width           =   285
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   27
      Left            =   1455
      MaxLength       =   2
      TabIndex        =   30
      Text            =   "99"
      Top             =   5130
      Width           =   735
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   30
      Left            =   2250
      MaxLength       =   8
      TabIndex        =   28
      Text            =   "99999999"
      Top             =   4785
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   29
      Left            =   6750
      MaxLength       =   8
      TabIndex        =   27
      Text            =   "99999999"
      Top             =   4515
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   28
      Left            =   2250
      MaxLength       =   8
      TabIndex        =   26
      Text            =   "99999999"
      Top             =   4515
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   25
      Left            =   6750
      MaxLength       =   8
      TabIndex        =   29
      Text            =   "99999999"
      Top             =   4785
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   24
      Left            =   4500
      MaxLength       =   8
      TabIndex        =   25
      Text            =   "99999999"
      Top             =   3765
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   23
      Left            =   2250
      MaxLength       =   8
      TabIndex        =   24
      Text            =   "99999999"
      Top             =   3765
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   22
      Left            =   6750
      MaxLength       =   8
      TabIndex        =   23
      Text            =   "99999999"
      Top             =   3495
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   21
      Left            =   4500
      MaxLength       =   8
      TabIndex        =   22
      Text            =   "99999999"
      Top             =   3495
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   20
      Left            =   2250
      MaxLength       =   8
      TabIndex        =   21
      Text            =   "99999999"
      Top             =   3495
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   19
      Left            =   6750
      MaxLength       =   8
      TabIndex        =   20
      Text            =   "99999999"
      Top             =   3225
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   18
      Left            =   4500
      MaxLength       =   8
      TabIndex        =   19
      Text            =   "99999999"
      Top             =   3225
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   17
      Left            =   2250
      MaxLength       =   8
      TabIndex        =   18
      Text            =   "99999999"
      Top             =   3225
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   16
      Left            =   6750
      MaxLength       =   8
      TabIndex        =   17
      Text            =   "99999999"
      Top             =   2970
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   15
      Left            =   4500
      MaxLength       =   8
      TabIndex        =   16
      Text            =   "99999999"
      Top             =   2970
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   14
      Left            =   2250
      MaxLength       =   8
      TabIndex        =   15
      Text            =   "99999999"
      Top             =   2970
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Height          =   270
      Index           =   3
      Left            =   3645
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1005
      Width           =   750
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   2
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "9712"
      Top             =   1005
      Width           =   750
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   13
      Left            =   2250
      MaxLength       =   7
      TabIndex        =   13
      Text            =   "9999999"
      Top             =   2520
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   12
      Left            =   6765
      MaxLength       =   8
      TabIndex        =   12
      Text            =   "99999999"
      Top             =   2250
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   11
      Left            =   4500
      MaxLength       =   7
      TabIndex        =   11
      Text            =   "9999999"
      Top             =   2250
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Height          =   270
      Index           =   1
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   0
      Text            =   "123456"
      Top             =   705
      Width           =   735
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   4
      Left            =   2250
      MaxLength       =   7
      TabIndex        =   4
      Text            =   "9999999"
      Top             =   1710
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   5
      Left            =   4500
      MaxLength       =   8
      TabIndex        =   5
      Text            =   "99999999"
      Top             =   1710
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   6
      Left            =   2250
      MaxLength       =   8
      TabIndex        =   7
      Text            =   "99999999"
      Top             =   1980
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   7
      Left            =   6750
      MaxLength       =   7
      TabIndex        =   6
      Text            =   "9999999"
      Top             =   1710
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   8
      Left            =   4500
      MaxLength       =   8
      TabIndex        =   8
      Text            =   "99999999"
      Top             =   1980
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   9
      Left            =   6750
      MaxLength       =   8
      TabIndex        =   9
      Text            =   "99999999"
      Top             =   1980
      Width           =   915
   End
   Begin VB.TextBox txtSM 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   10
      Left            =   2250
      MaxLength       =   7
      TabIndex        =   10
      Text            =   "9999999"
      Top             =   2250
      Width           =   915
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6915
      Top             =   0
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
            Picture         =   "frm170021.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170021.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170021.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170021.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170021.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170021.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170021.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170021.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170021.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170021.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170021.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   8340
      _ExtentX        =   14711
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
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   2700
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   690
      Width           =   5550
      VariousPropertyBits=   671105055
      Size            =   "9790;529"
      Value           =   "textCUID"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "證照津貼："
      Height          =   180
      Index           =   34
      Left            =   3555
      TabIndex        =   80
      Top             =   2565
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "未休假代金計算月薪："
      Height          =   180
      Index           =   38
      Left            =   4920
      TabIndex        =   79
      Top             =   5145
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "勞退自提費率："
      Height          =   180
      Index           =   37
      Left            =   945
      TabIndex        =   76
      Top             =   4080
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "補充保費："
      Height          =   180
      Index           =   35
      Left            =   5805
      TabIndex        =   75
      Top             =   3810
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公  司  別："
      Height          =   180
      Index           =   29
      Left            =   135
      TabIndex        =   72
      Top             =   1335
      Width           =   900
   End
   Begin VB.Label lblDsp 
      Caption         =   "台一國際專利商標事務所"
      Height          =   180
      Index           =   5
      Left            =   1425
      TabIndex        =   71
      Top             =   1335
      Width           =   4050
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "工作天數："
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
      Index           =   26
      Left            =   195
      TabIndex        =   69
      Top             =   5160
      Width           =   1200
   End
   Begin VB.Label lblDsp 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
      Height          =   180
      Index           =   3
      Left            =   6780
      TabIndex        =   68
      Top             =   2580
      Width           =   900
   End
   Begin VB.Label lblDsp 
      Alignment       =   1  '靠右對齊
      Caption         =   "9,999,999"
      Height          =   180
      Index           =   4
      Left            =   6795
      TabIndex        =   67
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "年終獎金基準月薪："
      Height          =   180
      Index           =   33
      Left            =   5085
      TabIndex        =   66
      Top             =   4830
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "其他"
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
      Index           =   32
      Left            =   225
      TabIndex        =   65
      Top             =   4545
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   8100
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   45
      X2              =   8145
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   8100
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "勞退公司提撥："
      Height          =   180
      Index           =   31
      Left            =   945
      TabIndex        =   64
      Top             =   4830
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "超時加班費："
      Height          =   180
      Index           =   30
      Left            =   1125
      TabIndex        =   63
      Top             =   4560
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "超時加班稅額+其他所得稅金："
      Height          =   180
      Index           =   9
      Left            =   4275
      TabIndex        =   62
      Top             =   4560
      Width           =   2430
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "貸款還款："
      Height          =   180
      Index           =   24
      Left            =   5805
      TabIndex        =   61
      Top             =   3270
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "勞退自提："
      Height          =   180
      Index           =   12
      Left            =   5805
      TabIndex        =   60
      Top             =   3015
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "其他扣款："
      Height          =   180
      Index           =   28
      Left            =   1305
      TabIndex        =   58
      Top             =   3810
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "未  打  卡："
      Height          =   180
      Index           =   15
      Left            =   5805
      TabIndex        =   57
      Top             =   3540
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "借支還款："
      Height          =   180
      Index           =   11
      Left            =   1305
      TabIndex        =   56
      Top             =   3540
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "應扣金額："
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   10
      Left            =   5805
      TabIndex        =   55
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "互  助  會："
      Height          =   180
      Index           =   8
      Left            =   3555
      TabIndex        =   54
      Top             =   3270
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "缺勤扣款："
      Height          =   180
      Index           =   7
      Left            =   3555
      TabIndex        =   53
      Top             =   3540
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "其他所得："
      Height          =   180
      Index           =   6
      Left            =   1305
      TabIndex        =   52
      Top             =   2565
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "加  班  費："
      Height          =   180
      Index           =   5
      Left            =   5805
      TabIndex        =   51
      Top             =   2295
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "加班時數："
      Height          =   180
      Index           =   4
      Left            =   3555
      TabIndex        =   50
      Top             =   2295
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "應發金額："
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   3
      Left            =   5805
      TabIndex        =   49
      Top             =   2565
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "薪資月份："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   48
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   47
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部　　門："
      Height          =   180
      Index           =   2
      Left            =   2700
      TabIndex        =   46
      Top             =   1050
      Width           =   900
   End
   Begin MSForms.Label lblName 
      Height          =   285
      Left            =   1845
      TabIndex        =   45
      Top             =   750
      Width           =   720
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1270;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "A1 台一部"
      Height          =   180
      Index           =   2
      Left            =   4470
      TabIndex        =   44
      Top             =   1050
      Width           =   795
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
      Left            =   150
      TabIndex        =   43
      Top             =   1740
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "基本薪資："
      Height          =   180
      Index           =   14
      Left            =   1305
      TabIndex        =   42
      Top             =   1755
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "職務津貼："
      Height          =   180
      Index           =   16
      Left            =   3555
      TabIndex        =   41
      Top             =   1755
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "技術津貼："
      Height          =   180
      Index           =   17
      Left            =   1305
      TabIndex        =   40
      Top             =   2025
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "午餐津貼："
      Height          =   180
      Index           =   18
      Left            =   5805
      TabIndex        =   39
      Top             =   1755
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "差旅津貼："
      Height          =   180
      Index           =   19
      Left            =   3555
      TabIndex        =   38
      Top             =   2025
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "房租津貼："
      Height          =   180
      Index           =   20
      Left            =   5805
      TabIndex        =   37
      Top             =   2025
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "特  支  費："
      Height          =   180
      Index           =   21
      Left            =   1305
      TabIndex        =   36
      Top             =   2295
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "勞  保  費："
      Height          =   180
      Index           =   22
      Left            =   1305
      TabIndex        =   35
      Top             =   3015
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "健  保  費："
      Height          =   180
      Index           =   23
      Left            =   3555
      TabIndex        =   34
      Top             =   3015
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "所  得  稅："
      Height          =   180
      Index           =   25
      Left            =   3555
      TabIndex        =   33
      Top             =   3810
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "扣除項目"
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
      Left            =   150
      TabIndex        =   32
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "婚喪互助："
      Height          =   180
      Index           =   36
      Left            =   1305
      TabIndex        =   31
      Top             =   3270
      Width           =   900
   End
End
Attribute VB_Name = "frm170021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/20 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Memo by Morgan 2024/1/31 新部門已修改
'Create by Morgan 2008/12/23
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_SM As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim m_bConfirmCheck As Boolean
Dim adoHiMonth As ADODB.Recordset '健保明細


Private Sub cmdHi_Click()
   If txtSM(1) <> "" And txtSM(2) <> "" Then
      If adoHiMonth Is Nothing Then
         GetHiMonth
      End If
      With frm170021_1
         .m_EditMode = m_EditMode
         .strSM01 = txtSM(1)
         .strSM02 = txtSM(2)
         .SetGrid adoHiMonth
         .Show vbModal
         If m_EditMode = 2 Then
            SumHiMonth
         End If
      End With
   End If
End Sub
'加總健保費明細
Private Sub SumHiMonth()
   Dim lSum As Long
   With adoHiMonth
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         lSum = lSum + Val("" & .Fields(2))
         .MoveNext
      Loop
   End If
   End With
   txtSM(15) = lSum
End Sub

Private Sub GetHiMonth()
   strSql = "select nvl(sr04,st02),decode(hm02,0,'自己',decode(sr03,'1','父親','2','母親','3','配偶','4','子女','其他'))" & _
      ",hm04,hm05,NVL(RTRIM(hm06||' '||HR04),'無') as Memo,hm01,hm02,hm03,hm04 as ohm04,hm06 as ohm06" & _
      " From himonth, staff, staff_relation, HiReduce" & _
      " where st01(+)=hm01 and sr01(+)=hm01 and sr02(+)=hm02 and hr01(+)=hm06" & _
      " and hm01='" & txtSM(1) & "' and hm03=" & (Val(txtSM(2)) + 191100) & _
      " order by hm02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   'Modify by Amy 2014/06/11 +FormName 改暫存TB
   'Set adoHiMonth = PUB_CreateRecordset(RsTemp, , , 300)
   Set adoHiMonth = PUB_CreateRecordset(RsTemp, , , 300, Me.Name)
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170021 = Nothing
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

Private Sub txtSM_Change(Index As Integer)
   Select Case Index
      Case 1
         If txtSM(Index) = "" Then
            lblName = "" 'Modify By Sindy 2021/12/20
         End If
   End Select
End Sub

Private Sub txtSM_GotFocus(Index As Integer)
   TextInverse txtSM(Index)
End Sub

Private Sub txtSM_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 1, 3, 37 '不控制
      Case 11
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") Then
         End If
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtSM_Validate(Index As Integer, Cancel As Boolean)
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Select Case Index
         Case 1
            If txtSM(Index) <> "" Then
               If ChkStaffID(txtSM(Index)) = True Then
                  Cancel = True
               End If
               If ClsPDGetStaffN(txtSM(Index), strExc(1), , True) = False Then
                  Cancel = True
               Else
                  lblName = strExc(1) 'Modify By Sindy 2021/12/20
               End If
            End If
            
         Case 2
            If txtSM(Index) <> "" Then
               If ChkDate(txtSM(Index) & "01") = False Then
                  Cancel = True
               End If
            End If
            
         Case 3
            If txtSM(Index) <> "" Then
               'Modified by Morgan 2024/1/31
               'If ClsPDGetStaffDeptName(txtSM(Index), strExc(1)) = False Then
               '   Cancel = True
               'Else
               '   lblDsp(2) = strExc(1)
               'End If
               If txtSM(2) >= (Left(新部門啟用日, 6) - 191100) Then
                  lblDsp(2) = GetPrjSalesBlack(txtSM(3), True)
               Else
                  lblDsp(2) = GetPrjSalesBlack(txtSM(3))
               End If
               If lblDsp(2) = "" Then Cancel = True
               'end 2024/1/31
            End If
            
         Case 4 To 10, 12 To 13
            Caculate1
            
         Case 14 To 24
            Caculate2
            
         'Removed by Morgan 2020/2/4 原"工作月數"取消改為"未休假代金計算月薪"
         'Case 26
         '   If Val(txtSM(Index)) > 1 Then
         '      MsgBox "工作月數不可大於 1 !"
         '      Cancel = True
         '   End If
            
         Case 37
            If txtSM(Index) <> "" Then
               lblDsp(5) = CompNameQuery(txtSM(Index))
               If lblDsp(5) = "" Then
                  ShowMsg "公司別錯誤 !"
                  Cancel = True
               End If
            End If
      End Select
      
      If Cancel = True Then TextInverse txtSM(Index)
      
      '若是案確定的檢查時略過
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
            Case 1
               '新增時預設部門
               If m_EditMode = 1 Then
                  'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
                  'txtSM(3) = PUB_GetST03(Replace(txtSM(Index), "A", "0"))
                  'Modified by Morgan 2024/1/31
                  'txtSM(3) = PUB_GetST03(Left(txtSM(Index), 1) & Replace(Mid(txtSM(Index), 2), "A", "0"))
                  'If txtSM(3) <> "" Then
                  '   If ClsPDGetStaffDeptName(txtSM(3), strExc(1)) Then
                  '      lblDsp(2) = strExc(1)
                  '   End If
                  'End If
                  If txtSM(2) = "" Or txtSM(2) >= (Left(新部門啟用日, 6) - 191100) Then
                     txtSM(3) = PUB_GetST93(txtSM(Index))
                     If txtSM(3) <> "" Then
                        lblDsp(2) = GetPrjSalesBlack(txtSM(3), True)
                     End If
                  Else
                     txtSM(3) = PUB_GetST03(txtSM(Index))
                     If txtSM(3) <> "" Then
                        lblDsp(2) = GetPrjSalesBlack(txtSM(3))
                     End If
                  End If
                  'end 2024/1/31
               End If
         End Select
      End If
   End If
End Sub

Private Sub Caculate1()
   lblDsp(3) = 0
   For intI = 4 To 13
      If intI <> 11 Then
         lblDsp(3) = lblDsp(3) + Val(txtSM(intI))
      End If
   Next
   lblDsp(3) = lblDsp(3) + Val(txtSM(45)) 'Added by Sindy 2020/8/4
   lblDsp(3) = Format(lblDsp(3), "#,###")
End Sub

Private Sub Caculate2()
   lblDsp(4) = 0
   For intI = 14 To 24
      lblDsp(4) = lblDsp(4) + Val(txtSM(intI))
   Next
   lblDsp(4) = lblDsp(4) + Val(txtSM(43)) 'Added by Morgan 2013/1/31
   lblDsp(4) = Format(lblDsp(4), "#,###")
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   With p_Rst
   If .RecordCount > 0 Then
      For Each oText In txtSM
         idx = oText.Index
         m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         '日期轉民國
         If idx = 2 Then
            If m_FieldList(idx).fiOldData <> "" Then
               oText.Text = Val(m_FieldList(idx).fiOldData) - 191100
            End If
         Else
            oText.Text = m_FieldList(idx).fiOldData
         End If
      Next
      CUID(1) = "" & .Fields("sm31")
      CUID(2) = "" & .Fields("sm32")
      CUID(3) = "" & .Fields("sm33")
      CUID(4) = "" & .Fields("sm34")
      CUID(5) = "" & .Fields("sm35")
      CUID(6) = "" & .Fields("sm36")
      
      If ClsPDGetStaffN(txtSM(1), strExc(1), , True) Then
         lblName = strExc(1) 'Modify By Sindy 2021/12/20
      End If
      'Modified by Morgan 2024/1/31
      'If ClsPDGetStaffDeptName(txtSM(3), strExc(1)) Then
      '   lblDsp(2) = strExc(1)
      'End If
      If txtSM(2) >= (Left(新部門啟用日, 6) - 191100) Then
         lblDsp(2) = GetPrjSalesBlack(txtSM(3), True)
      Else
         lblDsp(2) = GetPrjSalesBlack(txtSM(3))
      End If
      'end 2024/1/31
      lblDsp(5) = CompNameQuery(txtSM(37))
      
      Caculate1
      Caculate2
   End If
   End With
   UpdateCUID CUID, textCUID
   txtSM(1).Tag = txtSM(1)
   txtSM(2).Tag = txtSM(2)
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtSM
      oText.Locked = bLocked
   Next
   txtSM(15).Enabled = True
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
              " " & strCTime & String(6, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
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
         m_EditMode = 2
         SetInputEntry
         txtSM(15).Enabled = False 'Add by Morgan 2009/7/28 健保費要從明細改才會一致
         UpdateToolbarState

      Case vbKeyF5 ' 刪除
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         m_EditMode = 4
         SetInputEntry
         ClearField
         UpdateToolbarState
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
         UpdateToolbarState
         
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
         UpdateToolbarState
         
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
         UpdateToolbarState
         
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
         UpdateToolbarState
         
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
            txtSM(1) = txtSM(1).Tag
            txtSM(2) = txtSM(2).Tag
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
         'Added by Morgan 2015/3/27 翻譯費不可在此維護,需與其他所得同步,目前已算薪資就不可改,要由電腦中心處理
         If Left(txtSM(1), 1) = "F" Then
            TBar1.Buttons(1).Enabled = False
            TBar1.Buttons(2).Enabled = False
            TBar1.Buttons(3).Enabled = False
            TBar1.Buttons(4).Enabled = False
         Else
         'end 2015/3/27
            If m_bInsert Then
               TBar1.Buttons(1).Enabled = True
            Else
               TBar1.Buttons(1).Enabled = False
            End If
            If m_bUpdate And txtSM(1) <> "" Then
               TBar1.Buttons(2).Enabled = True
            Else
               TBar1.Buttons(2).Enabled = False
            End If
            If m_bDelete And txtSM(1) <> "" Then
               TBar1.Buttons(3).Enabled = True
            Else
               TBar1.Buttons(3).Enabled = False
            End If
            If m_bQuery Then
               TBar1.Buttons(4).Enabled = True
            Else
               TBar1.Buttons(4).Enabled = False
            End If
            
         End If 'Added by Morgan 2015/3/27
         
         If m_bQuery And txtSM(1) <> "" Then
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
         If Me.Visible = True Then
            txtSM(1).SetFocus
         End If
         
      Case 2
         SetCtrlReadOnly False
         txtSM(1).Locked = True
         txtSM(2).Locked = True
         If Me.Visible = True Then
            txtSM(3).SetFocus
         End If
      Case 4
         SetCtrlReadOnly True
         txtSM(1).Locked = False
         txtSM(2).Locked = False
         If Me.Visible = True Then
            txtSM(1).SetFocus
         End If
      Case Else
         SetCtrlReadOnly True
         If Me.Visible = True Then
            txtSM(1).SetFocus
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
            If AddRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
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
               txtSM(1).SetFocus
               txtSM_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
   
   Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   If txtSM(1) = "" Then
      ShowMsg "請輸員工代號 !"
      txtSM(1).SetFocus
      txtSM_GotFocus 1
      GoTo EscPoint
   End If
   If txtSM(2) = "" Then
      ShowMsg "請輸薪資月份 !"
      txtSM(2).SetFocus
      txtSM_GotFocus 2
      GoTo EscPoint
   End If
      
   '維護
   If m_EditMode = 1 Or m_EditMode = 2 Then
   
      'Added by Morgan 2012/9/12
      If txtSM(37) = "" Then
         ShowMsg "請輸公司別 !"
         txtSM(37).SetFocus
         txtSM_GotFocus 37
         GoTo EscPoint
      End If
      'end 2012/9/12
   
      If txtSM(3) = "" And txtSM(3).Locked = False Then
         ShowMsg "請輸入部門 !"
         txtSM(3).SetFocus
         txtSM_GotFocus 3
         GoTo EscPoint
      End If
      Caculate1
      If Val(Format(lblDsp(3))) = 0 Then
         ShowMsg "所得項目至少需輸入一項 !"
         txtSM(4).SetFocus
         txtSM_GotFocus 4
         GoTo EscPoint
      End If
      'Added by Morgan 2018/10/31
      If Val(txtSM(12)) > 0 And Val(txtSM(11)) = 0 Then
         MsgBox "有加班費時必須輸入加班時數！", vbExclamation
         txtSM(11).SetFocus
         GoTo EscPoint
      End If
      'end 2018/10/31
   End If
   
   For Each oText In txtSM
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtSM_Validate idx, bCancel
         If bCancel = True Then
            txtSM(idx).SetFocus
            txtSM_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   If m_EditMode = 1 Then
      If CheckExists(txtSM(1), txtSM(2)) Then
         ShowMsg "資料已存在 !"
         txtSM(1).SetFocus
         txtSM_GotFocus 1
         GoTo EscPoint
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
   For Each oText In txtSM
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
   stSQL = "INSERT INTO SalaryMonth (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
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
   Dim stHM04 As String, stHM06 As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE SalaryMonth SET "
   stSet = ""
   For Each oText In txtSM
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
      stSQL = stSQL & stSet & " where SM01='" & m_FieldList(1).fiNewData & "' AND SM02=" & m_FieldList(2).fiNewData & "; end; "
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   
   'Add by Morgan 2009/7/27
   If Not adoHiMonth Is Nothing Then
      With adoHiMonth
      .MoveFirst
      Do While Not .EOF
         stHM04 = "" & .Fields("hm04")
         If .Fields("memo") = "無" Then
            stHM06 = ""
         Else
            stHM06 = Left(.Fields("memo"), 2)
         End If
         If Val(stHM04) <> Val("" & .Fields("ohm04")) Or stHM06 <> "" & .Fields("ohm06") Then
            stSQL = "update Himonth set hm04=" & Val(stHM04) & ",hm06='" & stHM06 & "' where hm01='" & .Fields("hm01") & "'" & _
               " and hm02=" & .Fields("hm02") & " and hm03=" & .Fields("hm03")
            Pub_SeekTbLog stSQL
            cnnConnection.Execute stSQL, intI
         End If
         .MoveNext
      Loop
      End With
   End If
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Sub UpdateFieldNewData()
   For Each oText In txtSM
      idx = oText.Index
      If idx = 2 Then
         '年月轉西元
         m_FieldList(idx).fiNewData = Val(oText.Text) + 191100
      Else
         m_FieldList(idx).fiNewData = oText.Text
      End If
   Next
End Sub

Private Sub ClearField()
   lblName.Caption = Empty 'Modify By Sindy 2021/12/20
   For Each oText In txtSM
      oText.Text = Empty
   Next
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_SM
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   m_bConfirmCheck = False
   Set adoHiMonth = Nothing
End Sub

Private Function SetRefData(stUserNo As String) As Boolean
   
   
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '刪除
   stSQL = "delete from SalaryMonth where SM01='" & m_FieldList(1).fiNewData & "' AND SM02=" & m_FieldList(2).fiNewData
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtSM(1).Tag = ""
   txtSM(2).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from SalaryMonth where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_SM = .Fields.Count
      ReDim m_FieldList(TF_SM) As FIELDITEM
      For Each oText In txtSM
         idx = oText.Index
         m_FieldList(idx).fiName = "SM" & Format(idx, "00")
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
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
   
   Dim stKey01 As String
   Dim adoRst As New ADODB.Recordset
   
   stKey01 = m_FieldList(2).fiNewData & m_FieldList(3).fiNewData & m_FieldList(1).fiNewData
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM SalaryMonth" & _
            " WHERE SM01='" & txtSM(1) & "' AND SM02=" & (Val(txtSM(2)) + 191100)
      Case -2
         '2012/12/7 modify by sonia 辜進入很慢,加idxsm020301但未改善,故再加rownum<100,只改語法有改善但有idx更快
         'strExc(0) = "SELECT * FROM SalaryMonth order by 2 ASC,3 ASC,1 ASC"
         strExc(0) = "SELECT * FROM SalaryMonth where rownum<100 order by 2 ASC,3 ASC,1 ASC"
      Case -1
         '2012/12/7 modify by sonia 辜進入很慢,加idxsm020301但未改善,故再加rownum<100,只改語法有改善但有idx更快
         'strExc(0) = "SELECT * FROM SalaryMonth" & _
            " WHERE SM02||SM03||SM01 <'" & stKey01 & "' order by 2 DESC,3 DESC,1 DESC"
         strExc(0) = "SELECT * FROM SalaryMonth" & _
            " WHERE SM02||SM03||SM01 <'" & stKey01 & "' and rownum<100 order by 2 DESC,3 DESC,1 DESC"
      Case 1
         '2012/12/7 modify by sonia 辜進入很慢,加idxsm020301但未改善,故再加rownum<100,只改語法有改善但有idx更快
         'strExc(0) = "SELECT * FROM SalaryMonth" & _
            " WHERE SM02||SM03||SM01 >'" & stKey01 & "' order by 2 ASC,3 ASC,1 ASC"
         strExc(0) = "SELECT * FROM SalaryMonth" & _
            " WHERE SM02||SM03||SM01 >'" & stKey01 & "' and rownum<100 order by 2 ASC,3 ASC,1 ASC"
      Case 2
         '2012/12/7 modify by sonia 辜進入很慢,加idxsm020301但未改善,故再加rownum<100,只改語法有改善但有idx更快
         'strExc(0) = "SELECT * FROM SalaryMonth order by 2 DESC,3 DESC,1 DESC"
         strExc(0) = "SELECT * FROM SalaryMonth where rownum<100 order by 2 DESC,3 DESC,1 DESC"
   End Select
   intI = 1
   If adoRst.State <> 0 Then adoRst.Close
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
      txtSM(1).SetFocus
      txtSM_GotFocus 1
   End If
End Function

Private Function CheckExists(pSM01 As String, pSM02 As String) As Boolean
   CheckExists = True
   strExc(0) = "select 1 from salarymonth where sm01='" & pSM01 & "' and sm02=" & Val(pSM02) + 191100
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      CheckExists = False
   End If
End Function
