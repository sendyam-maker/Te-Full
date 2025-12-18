VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081020 
   BorderStyle     =   1  '單線固定
   Caption         =   "開拓客戶資料維護"
   ClientHeight    =   5940
   ClientLeft      =   1212
   ClientTop       =   1080
   ClientWidth     =   8076
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8076
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7470
      Top             =   30
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
            Picture         =   "frm081020.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm081020.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm081020.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm081020.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm081020.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm081020.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm081020.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm081020.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm081020.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm081020.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm081020.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   8076
      _ExtentX        =   14245
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   12
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "刪除"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "第一筆"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "前一筆"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "後一筆"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "最後筆"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "確定"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "取消"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "結束"
               EndProperty
            EndProperty
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   90
      TabIndex        =   19
      Top             =   720
      Width           =   7905
      _ExtentX        =   13949
      _ExtentY        =   9017
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   7
      TabHeight       =   420
      TabCaption(0)   =   "基本"
      TabPicture(0)   =   "frm081020.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "textECA02"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "textECA01"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "textECA03"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "內容"
      TabPicture(1)   =   "frm081020.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtEmailSameCnt"
      Tab(1).Control(1)=   "textECD14"
      Tab(1).Control(2)=   "textECD16"
      Tab(1).Control(3)=   "textECD15"
      Tab(1).Control(4)=   "textECD12"
      Tab(1).Control(5)=   "textECD04"
      Tab(1).Control(6)=   "textECD03"
      Tab(1).Control(7)=   "textECD05"
      Tab(1).Control(8)=   "textECD06"
      Tab(1).Control(9)=   "textECD07"
      Tab(1).Control(10)=   "textECD08"
      Tab(1).Control(11)=   "textECD09"
      Tab(1).Control(12)=   "textECD11"
      Tab(1).Control(13)=   "textECD01"
      Tab(1).Control(14)=   "textECD13"
      Tab(1).Control(15)=   "textECD10"
      Tab(1).Control(16)=   "textECD02"
      Tab(1).Control(17)=   "LabelECD02"
      Tab(1).Control(18)=   "Label1(3)"
      Tab(1).Control(19)=   "Label1(2)"
      Tab(1).Control(20)=   "Label8(3)"
      Tab(1).Control(21)=   "Label8(2)"
      Tab(1).Control(22)=   "Label8(1)"
      Tab(1).Control(23)=   "Label9"
      Tab(1).Control(24)=   "Label11"
      Tab(1).Control(25)=   "Label10"
      Tab(1).Control(26)=   "Label8(0)"
      Tab(1).Control(27)=   "Label7"
      Tab(1).Control(28)=   "Label6"
      Tab(1).Control(29)=   "Label1(0)"
      Tab(1).Control(30)=   "Label3(0)"
      Tab(1).Control(31)=   "Label2"
      Tab(1).ControlCount=   32
      Begin VB.TextBox txtEmailSameCnt 
         Height          =   300
         Left            =   -68400
         MaxLength       =   6
         TabIndex        =   38
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox textECA03 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox textECA01 
         Height          =   300
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   0
         Top             =   624
         Width           =   735
      End
      Begin MSForms.TextBox textECD14 
         Height          =   300
         Left            =   -73470
         TabIndex        =   16
         Top             =   3480
         Width           =   405
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "714;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD16 
         Height          =   765
         Left            =   -73470
         TabIndex        =   18
         Top             =   4140
         Width           =   5535
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "9763;1349"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD15 
         Height          =   300
         Left            =   -73470
         TabIndex        =   17
         Top             =   3810
         Width           =   1935
         VariousPropertyBits=   671105051
         MaxLength       =   12
         Size            =   "3413;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD12 
         Height          =   300
         Left            =   -70620
         TabIndex        =   8
         Top             =   1500
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD04 
         Height          =   300
         Left            =   -70620
         TabIndex        =   6
         Top             =   1170
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD03 
         Height          =   300
         Left            =   -73470
         TabIndex        =   5
         Top             =   1170
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD05 
         Height          =   300
         Left            =   -73470
         TabIndex        =   9
         Top             =   1830
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD06 
         Height          =   300
         Left            =   -70620
         TabIndex        =   10
         Top             =   1830
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD07 
         Height          =   300
         Left            =   -73470
         TabIndex        =   11
         Top             =   2160
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD08 
         Height          =   300
         Left            =   -70620
         TabIndex        =   12
         Top             =   2160
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD09 
         Height          =   300
         Left            =   -73470
         TabIndex        =   13
         Top             =   2490
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD11 
         Height          =   300
         Left            =   -73470
         TabIndex        =   7
         Top             =   1500
         Width           =   2775
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4895;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD01 
         Height          =   300
         Left            =   -73470
         TabIndex        =   4
         Top             =   840
         Width           =   1215
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD13 
         Height          =   300
         Left            =   -73470
         TabIndex        =   15
         Top             =   3150
         Width           =   5535
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "9763;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD10 
         Height          =   300
         Left            =   -73470
         TabIndex        =   14
         Top             =   2820
         Width           =   735
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECD02 
         Height          =   300
         Left            =   -73470
         TabIndex        =   3
         Top             =   510
         Width           =   615
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "1085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LabelECD02 
         Height          =   255
         Left            =   -71670
         TabIndex        =   39
         Top             =   515
         Width           =   3525
         VariousPropertyBits=   27
         Size            =   "6218;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textECA02 
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         Top             =   1104
         Width           =   5655
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "9975;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(N:不寄)"
         Height          =   180
         Index           =   3
         Left            =   -72930
         TabIndex        =   37
         Top             =   3510
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄電子報："
         Height          =   180
         Index           =   2
         Left            =   -74790
         TabIndex        =   36
         Top             =   3510
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   3
         Left            =   -74070
         TabIndex        =   35
         Top             =   4170
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "狀態："
         Height          =   180
         Index           =   2
         Left            =   -74070
         TabIndex        =   34
         Top             =   3840
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "E-MAIL："
         Height          =   180
         Index           =   1
         Left            =   -74310
         TabIndex        =   33
         Top             =   3180
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "xxx"
         Height          =   180
         Left            =   -72600
         TabIndex        =   32
         Top             =   2850
         Width           =   270
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "公司名稱："
         Height          =   180
         Left            =   -74430
         TabIndex        =   31
         Top             =   1530
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -72660
         TabIndex        =   30
         Top             =   3060
         Width           =   4365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "目前編號："
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   1584
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "國       籍："
         Height          =   180
         Index           =   0
         Left            =   -74385
         TabIndex        =   27
         Top             =   2850
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "地        址："
         Height          =   180
         Left            =   -74430
         TabIndex        =   26
         Top             =   1860
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "收  件  人："
         Height          =   180
         Left            =   -74430
         TabIndex        =   25
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "目前編號："
         Height          =   180
         Index           =   0
         Left            =   -74430
         TabIndex        =   24
         Top             =   870
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "屬性名稱："
         Height          =   180
         Index           =   0
         Left            =   -72630
         TabIndex        =   23
         Top             =   552
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "屬性代號："
         Height          =   180
         Left            =   -74430
         TabIndex        =   22
         Top             =   540
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "屬性代號："
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   624
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "屬性名稱："
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1104
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm081020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; textECA02、LabelECD02、textECD01~16
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim RcMain As New ADODB.Recordset, RsAdo As New ADODB.Recordset
' 變數宣告區
Dim m_EditMode As Integer
Dim m_SubMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM   '使用於 tf_ECA
Dim m_FieldList2() As FIELDITEM '使用於 tf_ECD
' 第一筆資料的本所案號
Dim m_FirstKEY(2) As String
Dim m_FirstKEY2(2) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(2) As String
Dim m_LastKEY2(2) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(2) As String
Dim m_CurrKEY2(2) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_ECA As Integer
Dim tf_ECD As Integer
Dim STabNum As String '記錄目前STab頁籤

Private Sub Form_Activate()
'   SSTab1.Tab = 0
'   textECA01.SetFocus
End Sub

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   
   '屬性檔
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from ExpandCusAttr where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_ECA = rsA.Fields.Count
   
   '資料檔
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from ExpandCusDetail where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_ECD = rsA.Fields.Count
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
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   ReDim m_FieldList(tf_ECA) As FIELDITEM
   ReDim m_FieldList2(tf_ECD) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   STabNum = "0"
   textECA01.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   txtEmailSameCnt.Visible = False
   
   InitialField
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm081020 = Nothing
   Set frm081020_1 = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Caption = "基本" Then
      STabNum = "0"
   ElseIf SSTab1.Caption = "內容" Then
      STabNum = "1"
   End If
   UpdateCtrlData
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
   
   '屬性檔
   If STabNum = "0" Then
         For nIndex = 0 To tf_ECA - 1
            If strName = m_FieldList(nIndex).fiName Then
               If strData = "#==#" Then
                  m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
               Else
                  m_FieldList(nIndex).fiNewData = strData
               End If
               Exit For
            End If
         Next nIndex
   '資料檔
   ElseIf STabNum = "1" Then
         For nIndex = 0 To tf_ECD - 1
            If strName = m_FieldList2(nIndex).fiName Then
               If strData = "#==#" Then
                  m_FieldList2(nIndex).fiNewData = m_FieldList2(nIndex).fiOldData
               Else
                  m_FieldList2(nIndex).fiNewData = strData
               End If
               Exit For
            End If
         Next nIndex
   End If
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String
   
   '屬性檔
   If STabNum = "0" Then
         For nIndex = 0 To tf_ECA - 1
            If m_FieldList(nIndex).fiName <> Empty Then
               If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
                  m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
                  m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
               Else
                  m_FieldList(nIndex).fiOldData = Empty
                  m_FieldList(nIndex).fiNewData = Empty
               End If
            End If
         Next nIndex
   '資料檔
   ElseIf STabNum = "1" Then
         For nIndex = 0 To tf_ECD - 1
            If m_FieldList2(nIndex).fiName <> Empty Then
               If IsNull(rsTmp.Fields(m_FieldList2(nIndex).fiName)) = False Then
                  m_FieldList2(nIndex).fiOldData = rsTmp.Fields(m_FieldList2(nIndex).fiName)
                  m_FieldList2(nIndex).fiNewData = rsTmp.Fields(m_FieldList2(nIndex).fiName)
               Else
                  m_FieldList2(nIndex).fiOldData = Empty
                  m_FieldList2(nIndex).fiNewData = Empty
               End If
            End If
         Next nIndex
   End If
EXITSUB:
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strECA01 As String
Dim strECA02 As String
Dim strECD01 As String
Dim strECD02 As String
   
   AddRecord = False
   
   '屬性檔
   If STabNum = "0" Then
         strECA01 = textECA01
         
         ' 檢查記錄是否已存在
         If IsRecordExist(strECA01, strECA02) = True Then
            strTit = "新增資料"
            strMsg = "該筆記錄已存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
            Exit Function
         End If
         
         bFirst = True
         bDifference = False
         strSql = "INSERT INTO ExpandCusAttr ("
         For nIndex = 0 To tf_ECA - 1
            strTmp = Empty
            If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
               strTmp = m_FieldList(nIndex).fiName
            End If
            If strTmp <> Empty Then
               bDifference = True
               If bFirst = True Then
                  strSql = strSql & strTmp
                  bFirst = False
               Else
                  strSql = strSql & "," & strTmp
               End If
            End If
         Next nIndex
         strSql = strSql & ") "
         strSql = strSql & "VALUES ("
      
         bFirst = True
         For nIndex = 0 To tf_ECA - 1
            strTmp = Empty
            If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
               If m_FieldList(nIndex).fiType = 0 Then
                  strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
               Else
                  strTmp = m_FieldList(nIndex).fiNewData
               End If
            End If
            If strTmp <> Empty Then
               If bFirst = True Then
                  strSql = strSql & strTmp
                  bFirst = False
               Else
                  strSql = strSql & "," & strTmp
               End If
            End If
         Next nIndex
         strSql = strSql & ")"
      
      On Error GoTo ErrHand
         cnnConnection.BeginTrans
      
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      
         If ((strECA01) < (m_FirstKEY(0))) Or ((strECA01) > (m_LastKEY(0))) Then
            RefreshRange
         End If
         cnnConnection.CommitTrans
      
         ShowCurrRecord strECA01, strECA02
         AddRecord = True
         Exit Function
         
   '資料檔
   ElseIf STabNum = "1" Then
         strECD01 = textECD01
         strECD02 = textECD02
         
         ' 檢查記錄是否已存在
         If IsRecordExist(strECD01, strECD02) = True Then
            strTit = "新增資料"
            strMsg = "該筆記錄已存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
            Exit Function
         End If
         
         bFirst = True
         bDifference = False
         strSql = "INSERT INTO ExpandCusDetail ("
         For nIndex = 0 To tf_ECD - 1
            strTmp = Empty
            If m_FieldList2(nIndex).fiOldData <> m_FieldList2(nIndex).fiNewData Then
               strTmp = m_FieldList2(nIndex).fiName
            End If
            If strTmp <> Empty Then
               bDifference = True
               If bFirst = True Then
                  strSql = strSql & strTmp
                  bFirst = False
               Else
                  strSql = strSql & "," & strTmp
               End If
            End If
         Next nIndex
         strSql = strSql & ") "
         strSql = strSql & "VALUES ("
      
         bFirst = True
         For nIndex = 0 To tf_ECD - 1
            strTmp = Empty
            If m_FieldList2(nIndex).fiOldData <> m_FieldList2(nIndex).fiNewData Then
               If m_FieldList2(nIndex).fiType = 0 Then
                  strTmp = "'" & ChgSQL(m_FieldList2(nIndex).fiNewData) & "'"
               Else
                  strTmp = m_FieldList2(nIndex).fiNewData
               End If
            End If
            If strTmp <> Empty Then
               If bFirst = True Then
                  strSql = strSql & strTmp
                  bFirst = False
               Else
                  strSql = strSql & "," & strTmp
               End If
            End If
         Next nIndex
         strSql = strSql & ")"
      
      On Error GoTo ErrHand
         cnnConnection.BeginTrans
         
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
         
         '更新屬性檔流水號
         strSql = "Update ExpandCusAttr Set ECA03=" & CNULL(textECD01) & " Where ECA01=" & CNULL(textECD02)
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
         
         If ((strECD01 & strECD02) < (m_FirstKEY2(0) & m_FirstKEY2(1))) Or ((strECD01 & strECD02) > (m_LastKEY2(0) & m_LastKEY2(1))) Then
            RefreshRange
         End If
         cnnConnection.CommitTrans
      
         ShowCurrRecord strECD01, strECD02
         AddRecord = True
         Exit Function
   End If
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 新增失敗！" & vbCrLf & Err.Description
   
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strECA01 As String
Dim strECA02 As String
Dim strECD01 As String
Dim strECD02 As String
   
   ModRecord = False
   
   '屬性檔
   If STabNum = "0" Then
         strECA01 = m_CurrKEY(0)
         
         strSql = "begin user_data.user_enabled:=1; UPDATE ExpandCusAttr SET "
         
         bFirst = True
         bDifference = False
         For nIndex = 0 To tf_ECA - 1
            strTmp = Empty
            If nIndex < 3 Or nIndex > 8 Then
                  If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
                     If m_FieldList(nIndex).fiType = 0 Then
                        If m_FieldList(nIndex).fiNewData = Empty Then
                           strTmp = m_FieldList(nIndex).fiName & " = NULL "
                        Else
                           strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
                        End If
                     Else
                        If m_FieldList(nIndex).fiNewData = Empty Then
                           strTmp = m_FieldList(nIndex).fiName & " = NULL "
                        Else
                           strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
                        End If
                     End If
                  End If
                  If strTmp <> Empty Then
                     bDifference = True
                     If bFirst = True Then
                        strSql = strSql & strTmp
                        bFirst = False
                     Else
                        strSql = strSql & "," & strTmp
                     End If
                  End If
              End If
         Next nIndex
      
         strSql = strSql & " " & _
                        "WHERE ECA01 = '" & strECA01 & "' ; end; "
      On Error GoTo ErrHand
         cnnConnection.BeginTrans
         If bDifference = True Then
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
         cnnConnection.CommitTrans
         
         ShowCurrRecord strECA01, strECA02
         
         ModRecord = True
         Exit Function
         
   '資料檔
   ElseIf STabNum = "1" Then
         strECD01 = m_CurrKEY2(0)
         strECD02 = m_CurrKEY2(1)
         
         strSql = "begin user_data.user_enabled:=1; UPDATE ExpandCusDetail SET "
         
         bFirst = True
         bDifference = False
         For nIndex = 0 To tf_ECD - 1
            strTmp = Empty
            If nIndex < 16 Or nIndex > 1 Then
                  If m_FieldList2(nIndex).fiOldData <> m_FieldList2(nIndex).fiNewData Then
                     If m_FieldList2(nIndex).fiType = 0 Then
                        If m_FieldList2(nIndex).fiNewData = Empty Then
                           strTmp = m_FieldList2(nIndex).fiName & " = NULL "
                        Else
                           strTmp = m_FieldList2(nIndex).fiName & " = '" & ChgSQL(m_FieldList2(nIndex).fiNewData) & "'"
                        End If
                     Else
                        If m_FieldList2(nIndex).fiNewData = Empty Then
                           strTmp = m_FieldList2(nIndex).fiName & " = NULL "
                        Else
                           strTmp = m_FieldList2(nIndex).fiName & " = " & m_FieldList2(nIndex).fiNewData
                        End If
                     End If
                  End If
                  If strTmp <> Empty Then
                     bDifference = True
                     If bFirst = True Then
                        strSql = strSql & strTmp
                        bFirst = False
                     Else
                        strSql = strSql & "," & strTmp
                     End If
                  End If
              End If
         Next nIndex
         
         strSql = strSql & " " & _
                        "WHERE ECD01 = '" & strECD01 & "' AND ECD02='" & strECD02 & "' ; end; "
      On Error GoTo ErrHand
         cnnConnection.BeginTrans
         If bDifference = True Then
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
         cnnConnection.CommitTrans
         
         ShowCurrRecord strECD01, strECD02
         
         ModRecord = True
         Exit Function
   End If
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
   
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strECA01 As String
Dim strECA02 As String
Dim strECD01 As String
Dim strECD02 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   '屬性檔
   If STabNum = "0" Then
         strECA01 = m_CurrKEY(0)
         
         strSql = "DELETE FROM ExpandCusAttr WHERE ECA01 = '" & strECA01 & "' "
         
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
         
         If (strECA01 = m_LastKEY(0)) Or (strECA01 = m_FirstKEY(0)) Then
            RefreshRange
         End If
         ShowCurrRecord strECA01, strECA02
         
   '資料檔
   ElseIf STabNum = "1" Then
         strECD01 = m_CurrKEY2(0)
         strECD02 = m_CurrKEY2(1)
         
         strSql = "DELETE FROM ExpandCusDetail WHERE ECD01 = '" & strECD01 & "' AND ECD02 = '" & strECD02 & "' "
         
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
         
         If (strECD01 = m_LastKEY2(0) And strECD02 = m_LastKEY2(1)) Or (strECD01 = m_FirstKEY2(0) And strECD02 = m_FirstKEY2(1)) Then
            RefreshRange
         End If
         ShowCurrRecord strECD01, strECD02
   End If
   
   DelRecord = True
   cnnConnection.CommitTrans
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
   
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strECA01 As String
Dim strECA02 As String
Dim strECD01 As String
Dim strECD02 As String
   
   QueryRecord = False
   
   '屬性檔
   If STabNum = "0" Then
         strECA01 = textECA01
         If IsRecordExist(strECA01, strECA02) = True Then
            m_CurrKEY(0) = strECA01
            QueryRecord = True
            UpdateCtrlData
         Else
            QueryRecord = False
         End If
   '資料檔
   ElseIf STabNum = "1" Then
         strECD01 = textECD01
         strECD02 = textECD02
         If IsRecordExist(strECD01, strECD02) = True Then
            m_CurrKEY2(0) = strECD01
            m_CurrKEY2(1) = strECD02
            QueryRecord = True
            UpdateCtrlData
         Else
            QueryRecord = False
         End If
   End If
   
   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse
   
   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If AddRecord = True Then
                RefreshRange
            Else
                Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '修改
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
         Else
            Exit Function
         End If
      Case 4: '查詢
         '屬性檔
         If STabNum = "0" Then
            If textECA01 = "" Then GoTo EXITSUB
         '資料檔
         ElseIf STabNum = "1" Then
            If textECD01 = "" Or textECD02 = "" Then GoTo EXITSUB
         End If
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   '屬性檔
   If STabNum = "0" Then
         Select Case m_EditMode
            Case 1: If Me.Visible = True Then textECA01.SetFocus
            Case 2: If Me.Visible = True Then textECA02.SetFocus
            Case 4: If Me.Visible = True Then textECA01.SetFocus
         End Select
   '資料檔
   ElseIf STabNum = "1" Then
         Select Case m_EditMode
            Case 1: If Me.Visible = True Then textECD02.SetFocus
            Case 2: If Me.Visible = True Then textECD03.SetFocus
            Case 4: If Me.Visible = True Then textECD01.SetFocus
         End Select
   End If
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   
   '屬性檔
   If STabNum = "0" Then
      strSql = "SELECT * FROM ExpandCusAttr " & _
               "WHERE ECA01 = '" & strKEY01 & "' "
   '資料檔
   ElseIf STabNum = "1" Then
      strSql = "SELECT * FROM ExpandCusDetail " & _
               "WHERE ECD01 = '" & strKEY01 & "' AND ECD02 = '" & strKEY02 & "' "
   End If
   
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

Private Function ISExistData(strECA01 As String, strName As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   strName = ""
   strSql = "SELECT * FROM EXPANDCUSATTR WHERE ECA01='" & strECA01 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.EOF = False Then
      If IsNull(rsTmp.Fields("ECA02")) = False Then: strName = rsTmp.Fields("ECA02")
      ISExistData = True
   Else
      ISExistData = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function ISExistAttr(strECD02 As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   strSql = "SELECT * FROM EXPANDCUSDETAIL WHERE ECD02='" & strECD02 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.EOF = False Then
      ISExistAttr = True
   Else
      ISExistAttr = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   '屬性檔
   If STabNum = "0" Then
         If IsRecordExist(strKEY01, strKEY02) = True Then
            m_CurrKEY(0) = strKEY01
         Else
            strSql = "SELECT * FROM ExpandCusAttr " & _
                     "WHERE ECA01 = '" & m_CurrKEY(0) & "' "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               If IsNull(rsTmp.Fields("ECA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ECA01")
               rsTmp.Close
               UpdateCtrlData
               GoTo EXITSUB
            End If
            rsTmp.Close
      
            strSql = "SELECT * FROM ExpandCusAttr Order by 1 ASC "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               If IsNull(rsTmp.Fields("ECA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ECA01")
            Else
               ShowLastRecord
               GoTo EXITSUB
            End If
            rsTmp.Close
         End If
         
   '資料檔
   ElseIf STabNum = "1" Then
         If IsRecordExist(strKEY01, strKEY02) = True Then
            m_CurrKEY2(0) = strKEY01
            m_CurrKEY2(1) = strKEY02
         Else
            strSql = "SELECT * FROM ExpandCusDetail " & _
                     "WHERE ECD01 = '" & m_CurrKEY2(0) & "' AND ECD02 = '" & m_CurrKEY2(1) & "' "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               If IsNull(rsTmp.Fields("ECD01")) = False Then: m_CurrKEY2(0) = rsTmp.Fields("ECD01")
               If IsNull(rsTmp.Fields("ECD02")) = False Then: m_CurrKEY2(1) = rsTmp.Fields("ECD02")
               rsTmp.Close
               UpdateCtrlData
               GoTo EXITSUB
            End If
            rsTmp.Close
            
            strSql = "SELECT * FROM ExpandCusDetail Order by 1 ASC "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               If IsNull(rsTmp.Fields("ECD01")) = False Then: m_CurrKEY2(0) = rsTmp.Fields("ECD01")
               If IsNull(rsTmp.Fields("ECD02")) = False Then: m_CurrKEY2(1) = rsTmp.Fields("ECD02")
            Else
               ShowLastRecord
               GoTo EXITSUB
            End If
            rsTmp.Close
         End If
   End If
   
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   m_CurrKEY2(0) = m_FirstKEY2(0)
   m_CurrKEY2(1) = m_FirstKEY2(1)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   '屬性檔
   If STabNum = "0" Then
         If m_CurrKEY(0) = m_FirstKEY(0) Then
            ShowMsg MsgText(9008)
            GoTo EXITSUB
         End If
         strSql = "SELECT * FROM ExpandCusAttr " & _
                  "WHERE ECA01 < '" & m_CurrKEY(0) & "' Order by 1 DESC "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("ECA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ECA01")
         End If
         
   '資料檔
   ElseIf STabNum = "1" Then
         If m_CurrKEY2(0) = m_FirstKEY2(0) And m_CurrKEY2(1) = m_FirstKEY2(1) Then
            ShowMsg MsgText(9008)
            GoTo EXITSUB
         End If
         strSql = "SELECT * FROM ExpandCusDetail " & _
                         "WHERE ECD02 = '" & m_CurrKEY2(1) & "' " & _
                              "AND ECD01 = (SELECT MAX(ECD01) FROM ExpandCusDetail " & _
                                                         "WHERE ECD02 = '" & m_CurrKEY2(1) & "' " & _
                                                              "AND ECD01 < '" & m_CurrKEY2(0) & "' )"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("ECD01")) = False Then: m_CurrKEY2(0) = rsTmp.Fields("ECD01")
            If IsNull(rsTmp.Fields("ECD02")) = False Then: m_CurrKEY2(1) = rsTmp.Fields("ECD02")
            rsTmp.Close
            UpdateCtrlData
            GoTo EXITSUB
         End If
         rsTmp.Close
         
         strSql = "SELECT * FROM ExpandCusDetail " & _
                         "WHERE ECD02 = (SELECT MAX(ECD02) FROM ExpandCusDetail " & _
                                                         "WHERE ECD02 < '" & m_CurrKEY2(1) & "') " & _
                              "AND ECD01 = (SELECT MAX(ECD01) FROM ExpandCusDetail " & _
                                                         "WHERE ECD02 = (SELECT MAX(ECD02) FROM ExpandCusDetail " & _
                                                                                         "WHERE ECD02 < '" & m_CurrKEY2(1) & "')) "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("ECD01")) = False Then: m_CurrKEY2(0) = rsTmp.Fields("ECD01")
            If IsNull(rsTmp.Fields("ECD02")) = False Then: m_CurrKEY2(1) = rsTmp.Fields("ECD02")
         End If
   End If
   
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   '屬性檔
   If STabNum = "0" Then
         If m_CurrKEY(0) = m_LastKEY(0) Then
            ShowMsg MsgText(9009)
            GoTo EXITSUB
         End If
         strSql = "SELECT * FROM ExpandCusAttr " & _
                  "WHERE ECA01 > '" & m_CurrKEY(0) & "' Order by 1 ASC "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("ECA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ECA01")
         End If
         
   '資料檔
   ElseIf STabNum = "1" Then
         If m_CurrKEY2(0) = m_LastKEY2(0) And m_CurrKEY2(1) = m_LastKEY2(1) Then
            ShowMsg MsgText(9009)
            GoTo EXITSUB
         End If
         strSql = "SELECT * FROM ExpandCusDetail " & _
                         "WHERE ECD02 = '" & m_CurrKEY2(1) & "' " & _
                              "AND ECD01 = (SELECT MIN(ECD01) FROM ExpandCusDetail " & _
                                                         "WHERE ECD02 = '" & m_CurrKEY2(1) & "' " & _
                                                              "AND ECD01 > '" & m_CurrKEY2(0) & "' )"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("ECD01")) = False Then: m_CurrKEY2(0) = rsTmp.Fields("ECD01")
            If IsNull(rsTmp.Fields("ECD02")) = False Then: m_CurrKEY2(1) = rsTmp.Fields("ECD02")
            rsTmp.Close
            UpdateCtrlData
            GoTo EXITSUB
         End If
         rsTmp.Close
         
         strSql = "SELECT * FROM ExpandCusDetail " & _
                         "WHERE ECD02 = (SELECT MIN(ECD02) FROM ExpandCusDetail " & _
                                                         "WHERE ECD02 > '" & m_CurrKEY2(1) & "') " & _
                              "AND ECD01 = (SELECT MIN(ECD01) FROM ExpandCusDetail " & _
                                                         "WHERE ECD02 = (SELECT MIN(ECD02) FROM ExpandCusDetail " & _
                                                                                         "WHERE ECD02 > '" & m_CurrKEY2(1) & "')) "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("ECD01")) = False Then: m_CurrKEY2(0) = rsTmp.Fields("ECD01")
            If IsNull(rsTmp.Fields("ECD02")) = False Then: m_CurrKEY2(1) = rsTmp.Fields("ECD02")
         End If
   End If
   
   rsTmp.Close
   UpdateCtrlData

EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   m_CurrKEY2(0) = m_LastKEY2(0)
   m_CurrKEY2(1) = m_LastKEY2(1)
   UpdateCtrlData
End Sub

' 執行指令
'SONIA
'Private Sub OnAction(ByVal KeyCode As Integer)
Public Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   'm_SubMode = 0
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         If STabNum = "0" Then
            Me.SSTab1.TabEnabled(0) = True
            Me.SSTab1.TabEnabled(1) = False
            SSTab1.Tab = 0
         ElseIf STabNum = "1" Then
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
         End If
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         If STabNum = "0" Then
            Me.SSTab1.TabEnabled(0) = True
            Me.SSTab1.TabEnabled(1) = False
            SSTab1.Tab = 0
         ElseIf STabNum = "1" Then
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
         End If
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         If STabNum = "0" Then
            If ISExistAttr(textECA01) = True Then
               MsgBox "開拓客戶資料檔有此屬性代號，不可刪除!", vbExclamation, "開拓客戶資料維護"
               Exit Sub
            End If
         End If
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         If STabNum = "0" Then
            Me.SSTab1.TabEnabled(0) = True
            Me.SSTab1.TabEnabled(1) = False
            SSTab1.Tab = 0
         ElseIf STabNum = "1" Then
            Me.SSTab1.TabEnabled(0) = False
            Me.SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
         End If
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
         PUB_FilterFormText Me 'Add by Morgan 2008/6/20 修正畫面所有含跳行符號的文字框
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         If OnWork = True Then
            If STabNum = "0" Then
               Me.SSTab1.TabEnabled(0) = True
               Me.SSTab1.TabEnabled(1) = True
               SSTab1.Tab = 0
            ElseIf STabNum = "1" Then
               Me.SSTab1.TabEnabled(0) = True
               Me.SSTab1.TabEnabled(1) = True
               SSTab1.Tab = 1
            End If
            UpdateToolbarState
         Else
            Exit Sub
         End If
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  If STabNum = "0" Then
                     Me.SSTab1.TabEnabled(0) = True
                     Me.SSTab1.TabEnabled(1) = True
                     SSTab1.Tab = 0
                  ElseIf STabNum = "1" Then
                     Me.SSTab1.TabEnabled(0) = True
                     Me.SSTab1.TabEnabled(1) = True
                     SSTab1.Tab = 1
                  End If
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               If STabNum = "0" Then
                  Me.SSTab1.TabEnabled(0) = True
                  Me.SSTab1.TabEnabled(1) = True
                  SSTab1.Tab = 0
               ElseIf STabNum = "1" Then
                  Me.SSTab1.TabEnabled(0) = True
                  Me.SSTab1.TabEnabled(1) = True
                  SSTab1.Tab = 1
               End If
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
         txtEmailSameCnt = "" ' Add By Sindy 98/03/05
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   '屬性檔
   'If STabNum = "0" Then
         strSql = "SELECT * FROM ExpandCusAttr Order BY 1 ASC "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("ECA01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("ECA01")
         End If
         rsTmp.Close
         
         strSql = "SELECT * FROM ExpandCusAttr Order BY 1 DESC "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("ECA01")) = False Then: m_LastKEY(0) = rsTmp.Fields("ECA01")
         End If
         rsTmp.Close
         
   '資料檔
   'ElseIf STabNum = "1" Then
         strSql = "SELECT * FROM ExpandCusDetail Order BY 2 ASC,1 ASC "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("ECD01")) = False Then: m_FirstKEY2(0) = rsTmp.Fields("ECD01")
            If IsNull(rsTmp.Fields("ECD02")) = False Then: m_FirstKEY2(1) = rsTmp.Fields("ECD02")
         End If
         rsTmp.Close
         
         strSql = "SELECT * FROM ExpandCusDetail Order BY 2 DESC,1 DESC "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("ECD01")) = False Then: m_LastKEY2(0) = rsTmp.Fields("ECD01")
            If IsNull(rsTmp.Fields("ECD02")) = False Then: m_LastKEY2(1) = rsTmp.Fields("ECD02")
         End If
         rsTmp.Close
   'End If
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer
   
   '屬性檔
   If STabNum = "0" Then
         strSql = "SELECT * FROM ExpandCusAttr " & _
                  "WHERE ECA01='" & m_CurrKEY(0) & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            ClearField
            If IsNull(rsTmp.Fields("ECA01")) = False Then: textECA01 = rsTmp.Fields("ECA01")
            If IsNull(rsTmp.Fields("ECA02")) = False Then: textECA02 = rsTmp.Fields("ECA02")
            If IsNull(rsTmp.Fields("ECA03")) = False Then: textECA03 = rsTmp.Fields("ECA03")
            ' 更新暫存區的資料
            UpdateFieldOldData rsTmp
         End If
         rsTmp.Close
         
   '資料檔
   ElseIf STabNum = "1" Then
         strSql = "SELECT * FROM ExpandCusDetail,ExpandCusAttr " & _
                  "WHERE ECA01=ECD02 " & _
                       "AND ECD01='" & m_CurrKEY2(0) & "' AND ECD02='" & m_CurrKEY2(1) & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            ClearField
            If IsNull(rsTmp.Fields("ECD01")) = False Then: textECD01 = rsTmp.Fields("ECD01")
            If IsNull(rsTmp.Fields("ECD02")) = False Then: textECD02 = rsTmp.Fields("ECD02")
            If IsNull(rsTmp.Fields("ECD03")) = False Then: textECD03 = rsTmp.Fields("ECD03")
            If IsNull(rsTmp.Fields("ECD04")) = False Then: textECD04 = rsTmp.Fields("ECD04")
            If IsNull(rsTmp.Fields("ECD05")) = False Then: textECD05 = rsTmp.Fields("ECD05")
            If IsNull(rsTmp.Fields("ECD06")) = False Then: textECD06 = rsTmp.Fields("ECD06")
            If IsNull(rsTmp.Fields("ECD07")) = False Then: textECD07 = rsTmp.Fields("ECD07")
            If IsNull(rsTmp.Fields("ECD08")) = False Then: textECD08 = rsTmp.Fields("ECD08")
            If IsNull(rsTmp.Fields("ECD09")) = False Then: textECD09 = rsTmp.Fields("ECD09")
            If IsNull(rsTmp.Fields("ECD10")) = False Then: textECD10 = rsTmp.Fields("ECD10")
            If IsNull(rsTmp.Fields("ECD11")) = False Then: textECD11 = rsTmp.Fields("ECD11")
            If IsNull(rsTmp.Fields("ECD12")) = False Then: textECD12 = rsTmp.Fields("ECD12")
            If IsNull(rsTmp.Fields("ECD13")) = False Then: textECD13 = rsTmp.Fields("ECD13")
            If IsNull(rsTmp.Fields("ECD14")) = False Then: textECD14 = rsTmp.Fields("ECD14")
            If IsNull(rsTmp.Fields("ECD15")) = False Then: textECD15 = rsTmp.Fields("ECD15")
            If IsNull(rsTmp.Fields("ECD16")) = False Then: textECD16 = rsTmp.Fields("ECD16")
            
            '屬性名稱
            LabelECD02.Caption = rsTmp.Fields("ECA02")
            '國籍名稱
            If IsNull(rsTmp.Fields("ECD10")) = False Then
               If ClsPDGetNation(rsTmp.Fields("ECD10"), strExc(0)) = True Then
                  Label9.Caption = strExc(0)
               End If
            End If
            
            ' 更新暫存區的資料
            UpdateFieldOldData rsTmp
         End If
         rsTmp.Close
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
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
      ' 新增
      Case 1, 2, 3, 4:
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

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   '屬性檔
   If STabNum = "0" Then
         textECA01.Locked = bEnable
         If bEnable Then textECA01.BackColor = &H8000000F Else textECA01.BackColor = &H80000005
   '資料檔
   ElseIf STabNum = "1" Then
         textECD01.Locked = bEnable
         textECD02.Locked = bEnable
         If bEnable Then textECD01.BackColor = &H8000000F Else textECD01.BackColor = &H80000005
         If bEnable Then textECD02.BackColor = &H8000000F Else textECD02.BackColor = &H80000005
   End If
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   '屬性檔
   'If STabNum = "0" Then
         textECA01.Locked = bEnable
         If bEnable Then textECA01.BackColor = &H8000000F Else textECA01.BackColor = &H80000005
         textECA02.Locked = bEnable
         textECA03.Locked = bEnable
   '資料檔
   'ElseIf STabNum = "1" Then
         'textECD01.Locked = bEnable
         textECD01.Locked = True
         textECD02.Locked = bEnable
         'If bEnable Then textECD01.BackColor = &H8000000F Else textECD01.BackColor = &H80000005
         textECD01.BackColor = &H8000000F
         If bEnable Then textECD02.BackColor = &H8000000F Else textECD02.BackColor = &H80000005
         textECD03.Locked = bEnable
         textECD04.Locked = bEnable
         textECD05.Locked = bEnable
         textECD06.Locked = bEnable
         textECD07.Locked = bEnable
         textECD08.Locked = bEnable
         textECD09.Locked = bEnable
         textECD10.Locked = bEnable
         textECD11.Locked = bEnable
         textECD12.Locked = bEnable
         textECD13.Locked = bEnable
         textECD14.Locked = bEnable
         textECD15.Locked = bEnable
         textECD16.Locked = bEnable
   'End If
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   '屬性檔
   'If STabNum = "0" Then
         textECA01 = Empty
         textECA02 = Empty
         textECA03 = Empty
         For nIndex = 0 To tf_ECA - 1
            m_FieldList(nIndex).fiOldData = Empty
            m_FieldList(nIndex).fiNewData = Empty
         Next nIndex
   '資料檔
   'ElseIf STabNum = "1" Then
         textECD01 = Empty
         textECD02 = Empty
         textECD03 = Empty
         textECD04 = Empty
         textECD05 = Empty
         textECD06 = Empty
         textECD07 = Empty
         textECD08 = Empty
         textECD09 = Empty
         textECD10 = Empty
         textECD11 = Empty
         textECD12 = Empty
         textECD13 = Empty
         textECD14 = Empty
         textECD15 = Empty
         textECD16 = Empty
         LabelECD02.Caption = Empty
         Label9.Caption = Empty
         For nIndex = 0 To tf_ECD - 1
            m_FieldList2(nIndex).fiOldData = Empty
            m_FieldList2(nIndex).fiNewData = Empty
         Next nIndex
   'End If
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   '屬性檔
   If STabNum = "0" Then
         '若新增資料
         If m_EditMode = 1 Then
            SetFieldNewData "ECA01", textECA01
         End If
         SetFieldNewData "ECA02", textECA02
         SetFieldNewData "ECA03", textECA03
   '資料檔
   ElseIf STabNum = "1" Then
         '若新增資料
         If m_EditMode = 1 Then
            SetFieldNewData "ECD01", textECD01
            SetFieldNewData "ECD02", textECD02
         End If
         SetFieldNewData "ECD03", textECD03
         SetFieldNewData "ECD04", textECD04
         SetFieldNewData "ECD05", textECD05
         SetFieldNewData "ECD06", textECD06
         SetFieldNewData "ECD07", textECD07
         SetFieldNewData "ECD08", textECD08
         SetFieldNewData "ECD09", textECD09
         SetFieldNewData "ECD10", textECD10
         SetFieldNewData "ECD11", textECD11
         SetFieldNewData "ECD12", textECD12
         SetFieldNewData "ECD13", textECD13
         SetFieldNewData "ECD14", textECD14
         SetFieldNewData "ECD15", textECD15
         SetFieldNewData "ECD16", textECD16
   End If
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   
   ' 初始化欄位陣列
   ' 開拓客戶屬性檔 ExpandCusAttr
   For nIndex = 1 To tf_ECA
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "ECA" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 3:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
   
    ' 開拓客戶資料檔 ExpandCusDetail
   For nIndex = 1 To tf_ECD
      strTmp = Format(nIndex, "00")
      m_FieldList2(nIndex - 1).fiName = "ECD" & strTmp
      m_FieldList2(nIndex - 1).fiOldData = Empty
      m_FieldList2(nIndex - 1).fiNewData = Empty
      m_FieldList2(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 1:
            m_FieldList2(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub


Private Function CheckDataValid() As Boolean
Dim nResponse
Dim strTmp  As String
Dim strChkText As String
   
   CheckDataValid = False
   
   '屬性檔
   If STabNum = "0" Then
         If textECA01.Text = "" Then
             MsgBox "屬性代號不可以空白！", vbExclamation
             textECA01.SetFocus
             Exit Function
         End If
         If textECA02.Text = "" Then
             MsgBox "屬性名稱不可以空白！", vbExclamation
             textECA02.SetFocus
             Exit Function
         End If
   '資料檔
   ElseIf STabNum = "1" Then
         If textECD01.Text = "" Then
             MsgBox "目前編號不可以空白！", vbExclamation
             textECD01.SetFocus
             Exit Function
         End If
         If textECD02.Text = "" Then
             MsgBox "屬性代號不可以空白！", vbExclamation
             textECD02.SetFocus
             Exit Function
         End If
         If textECD03.Text <> "" Or textECD04.Text <> "" Then
            If textECD05.Text = "" And textECD06.Text = "" And _
               textECD07.Text = "" And textECD08.Text = "" And _
               textECD09.Text = "" Then
               MsgBox "有收件人時，地址不可以空白！", vbExclamation
               textECD05.SetFocus
               Exit Function
            End If
         End If
         If textECD03.Text = "" And textECD04.Text = "" And textECD13.Text = "" Then
            MsgBox "收件人及E-Mail至少輸入一項！", vbExclamation
            textECD03.SetFocus
            Exit Function
         End If
         If textECD10.Text = "" Then
            MsgBox "國籍不可以空白！", vbExclamation
            textECD10.SetFocus
            Exit Function
         End If
         ' Add By Sindy 98/03/05 檢查E-Mail是否存在
         If txtEmailSameCnt.Text = "Y" Then
            txtEmailSameCnt.Text = ""
         Else
            txtEmailSameCnt.Text = 0
            Me.Enabled = False
   '         If fnSaveParentForm(Me) = False Then
   '             Me.Enabled = True
   '             Exit Function
   '         End If
            Screen.MousePointer = vbHourglass
            frm081020_1.Show
            frm081020_1.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Me.Hide
            If txtEmailSameCnt.Text <> "" And txtEmailSameCnt.Text <> 0 Then
               'nResponse = MsgBox(strChkText & "，資料已存在！是否要繼續？", vbYesNo + vbCritical + vbDefaultButton2, "詢問")
               'If nResponse = vbNo Then
                  'Me.textECD13.SetFocus
                  Exit Function
               'End If
            Else
               Me.Show
            End If
   '         If CheckEmailExist(strChkText) = True Then
   '           nResponse = MsgBox(strChkText & "，資料已存在！是否要繼續？", vbYesNo + vbCritical + vbDefaultButton2, "詢問")
   '           If nResponse = vbNo Then
   '              textECD13.SetFocus
   '              Exit Function
   '           End If
   '         End If
         End If
         ' 98/03/05 End
   End If
   
   CheckDataValid = True
EXITSUB:
End Function


Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   '屬性檔
   If STabNum = "0" Then
         If Me.textECA01.Enabled = True Then
            Cancel = False
            textECA01_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
   '資料檔
   ElseIf STabNum = "1" Then
         If Me.textECD02.Enabled = True Then
            Cancel = False
            textECD02_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD03.Enabled = True Then
            Cancel = False
            textECD03_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD04.Enabled = True Then
            Cancel = False
            textECD04_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD05.Enabled = True Then
            Cancel = False
            textECD05_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD06.Enabled = True Then
            Cancel = False
            textECD06_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD07.Enabled = True Then
            Cancel = False
            textECD07_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD08.Enabled = True Then
            Cancel = False
            textECD08_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD09.Enabled = True Then
            Cancel = False
            textECD09_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD10.Enabled = True Then
            Cancel = False
            textECD10_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD11.Enabled = True Then
            Cancel = False
            textECD11_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD12.Enabled = True Then
            Cancel = False
            textECD12_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD13.Enabled = True Then
            Cancel = False
            textECD13_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD14.Enabled = True Then
            Cancel = False
            textECD14_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD15.Enabled = True Then
            Cancel = False
            textECD15_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
         If Me.textECD16.Enabled = True Then
            Cancel = False
            textECD16_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
   End If
   
   'Added by Lydia 2021/09/22 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If

   TxtValidate = True
End Function


Private Sub textECA01_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textECA01
   End If
End Sub
Private Sub textECA01_KeyPress(KeyAscii As Integer)
   'KeyAscii = Pub_NumAscii(KeyAscii)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textECA01_Validate(Cancel As Boolean)
   If m_EditMode = 1 And textECA01 <> "" Then
       If IsRecordExist(textECA01, "") = True And textECA01.Enabled = True And textECA01.Locked = False Then
           MsgBox "已有此筆資料不可重複新增!", vbExclamation, "開拓客戶資料維護"
           Call textECA01_GotFocus
           Cancel = True
           Exit Sub
       End If
   End If
End Sub

Private Sub textECA02_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textECA02
       CloseIme
   End If
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub textECA02_KeyPress(KeyAscii As Integer)
Private Sub textECA02_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textECA02_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub textECD02_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textECD02
   End If
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub textECD02_KeyPress(KeyAscii As Integer)
Private Sub textECD02_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textECD02_Validate(Cancel As Boolean)
Dim strName As String
Dim Rc1 As New ADODB.Recordset
   
   If textECD02 = "" Then LabelECD02.Caption = ""
   If m_EditMode = 1 And textECD02 <> "" Then
      If ISExistData(textECD02, strName) = False And textECD02.Enabled = True And textECD02.Locked = False Then
         MsgBox "屬性代號不存在!", vbExclamation, "開拓客戶資料維護"
         Call textECD02_GotFocus
         Cancel = True
         Exit Sub
      Else
         LabelECD02.Caption = strName
      End If
      
      '自動給號
      strExc(0) = "Select ECA03 From ExpandCusAttr Where ECA01=" & CNULL(textECD02.Text)
      intI = 1
      Set Rc1 = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If IsNull(Rc1.Fields(0).Value) = False Then
            textECD01.Text = Rc1.Fields(0).Value + 1
         Else
            textECD01.Text = 1
         End If
      End If
   End If
   CloseIme
End Sub

Private Sub textECD03_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textECD03
       OpenIme
   End If
End Sub
Private Sub textECD03_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD03 <> "" Then
    If CheckLengthIsOK(textECD03, textECD03.MaxLength) = False Then
        Call textECD03_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textECD04_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textECD04
       OpenIme
   End If
End Sub
Private Sub textECD04_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD04 <> "" Then
    If CheckLengthIsOK(textECD04, textECD04.MaxLength) = False Then
        Call textECD04_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textECD05_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textECD05
    OpenIme
End If
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub textECD05_KeyPress(KeyAscii As Integer)
Private Sub textECD05_KeyPress(KeyAscii As MSForms.ReturnInteger)
'KeyAscii = ChangeZIP(KeyAscii)
End Sub
Private Sub textECD05_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD05 <> "" Then
    If CheckLengthIsOK(textECD05, textECD05.MaxLength) = False Then
        Call textECD05_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textECD06_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textECD06
    OpenIme
End If
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub textECD06_KeyPress(KeyAscii As Integer)
Private Sub textECD06_KeyPress(KeyAscii As MSForms.ReturnInteger)
'KeyAscii = ChangeZIP(KeyAscii)
End Sub
Private Sub textECD06_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD06 <> "" Then
    If CheckLengthIsOK(textECD06, textECD06.MaxLength) = False Then
        Call textECD06_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textECD07_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textECD07
    OpenIme
End If
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub textECD07_KeyPress(KeyAscii As Integer)
Private Sub textECD07_KeyPress(KeyAscii As MSForms.ReturnInteger)
'KeyAscii = ChangeZIP(KeyAscii)
End Sub
Private Sub textECD07_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD07 <> "" Then
    If CheckLengthIsOK(textECD07, textECD07.MaxLength) = False Then
        Call textECD07_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textECD08_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textECD08
    OpenIme
End If
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub textECD08_KeyPress(KeyAscii As Integer)
Private Sub textECD08_KeyPress(KeyAscii As MSForms.ReturnInteger)
'KeyAscii = ChangeZIP(KeyAscii)
End Sub
Private Sub textECD08_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD08 <> "" Then
    If CheckLengthIsOK(textECD08, textECD08.MaxLength) = False Then
        Call textECD08_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textECD09_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textECD09
    OpenIme
End If
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub textECD09_KeyPress(KeyAscii As Integer)
Private Sub textECD09_KeyPress(KeyAscii As MSForms.ReturnInteger)
'KeyAscii = ChangeZIP(KeyAscii)
End Sub
Private Sub textECD09_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD09 <> "" Then
    If CheckLengthIsOK(textECD09, textECD09.MaxLength) = False Then
        Call textECD09_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textECD10_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textECD10
   End If
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub textECD10_KeyPress(KeyAscii As Integer)
Private Sub textECD10_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textECD10_Validate(Cancel As Boolean)
   If textECD10 = "" Then Label9.Caption = ""
   If m_EditMode <> 0 And textECD10 <> "" Then
      '國籍名稱
      If ClsPDGetNation(textECD10, strExc(0)) = True Then
         Label9.Caption = strExc(0)
      Else
         'MsgBox "國籍代號不存在!", vbExclamation, "開拓客戶資料維護"
         Call textECD10_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   CloseIme
End Sub

Private Sub textECD11_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textECD11
       OpenIme
   End If
End Sub
Private Sub textECD11_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD11 <> "" Then
    If CheckLengthIsOK(textECD11, textECD11.MaxLength) = False Then
        Call textECD11_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textECD12_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textECD12
       OpenIme
   End If
End Sub
Private Sub textECD12_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD12 <> "" Then
    If CheckLengthIsOK(textECD12, textECD12.MaxLength) = False Then
        Call textECD12_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textECD13_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textECD13
End If
End Sub
Private Sub textECD13_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD13 <> "" Then
    If CheckLengthIsOK(textECD13, textECD13.MaxLength) = False Then
        Call textECD13_GotFocus
        Cancel = True
        Exit Sub
    End If
    If PUB_CheckMail(textECD13.Text) = False Then
       Call textECD13_GotFocus
       Cancel = True
       Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textECD14_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textECD14
   End If
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub textECD14_KeyPress(KeyAscii As Integer)
Private Sub textECD14_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textECD14_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textECD14 <> "" Then
      Select Case textECD14
      Case "N", ""
      Case Else
          MsgBox "是否寄電子報只可以輸入 N 或 空白！", vbInformation, "輸入錯誤！"
          Call textECD14_GotFocus
          Cancel = True
         Exit Sub
      End Select
   End If
   CloseIme
End Sub

Private Sub textECD15_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textECD15
       OpenIme
   End If
End Sub
Private Sub textECD15_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD15 <> "" Then
    If CheckLengthIsOK(textECD15, textECD15.MaxLength) = False Then
        Call textECD15_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

Private Sub textECD16_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textECD16
       OpenIme
   End If
End Sub
Private Sub textECD16_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textECD16 <> "" Then
    If CheckLengthIsOK(textECD16, textECD16.MaxLength) = False Then
        Call textECD16_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
CloseIme
End Sub

' Add By Sindy 98/03/05
' 檢查E-Mail是否存在於客戶檔, 國外代理人檔, 潛在客戶檔(含國內), 外法開拓客戶檔
' 若有, 顯示提示訊息, 並且可以選擇忽略
Private Function CheckEmailExist(strChkText As String) As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strCUText As String, strFAText As String, strPCUText As String, strPOCText As String, strECDText As String
   
   CheckEmailExist = False
   
   strCUText = ""
   '客戶檔
   'Modified by Lydia 2024/09/18 +財務副本信箱(CU200)
   strSql = "SELECT CU01||CU02 FROM Customer " & _
                  "Where (instr(NLS_Upper(CU20),'" & UCase(ChgSQL(textECD13)) & "')>0 " & _
                  "or instr(NLS_Upper(CU115),'" & UCase(ChgSQL(textECD13)) & "')>0 " & _
                  "or instr(NLS_Upper(CU116),'" & UCase(ChgSQL(textECD13)) & "')>0 " & _
                  "or instr(NLS_Upper(CU117),'" & UCase(ChgSQL(textECD13)) & "')>0 " & _
                  "or instr(NLS_Upper(CU118),'" & UCase(ChgSQL(textECD13)) & "')>0 " & _
                  "or instr(NLS_Upper(CU200),'" & UCase(ChgSQL(textECD13)) & "')>0) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      With rsTmp
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If strCUText = "" Then
               strCUText = "客戶檔：" & rsTmp.Fields(0)
            Else
               strCUText = strCUText & "," & rsTmp.Fields(0)
            End If
            rsTmp.MoveNext
         Loop
      End With
   End If
   rsTmp.Close
   
   strFAText = ""
   '國外代理人檔
   'Modified by Lydia 2018/07/20 +FA105 財務信箱(CF)
   'Modified by Lydia 2024/09/18 +財務副本信箱(FA134)
   strSql = "SELECT FA01||FA02 FROM Fagent " & _
                  "Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(textECD13)) & "')> 0 " & _
                  "or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(textECD13)) & "')> 0 " & _
                  "Or InStr(NLS_Upper(fa105),'" & UCase(ChgSQL(textECD13)) & "') > 0 " & _
                  "or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(textECD13)) & "')> 0 " & _
                  "or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(textECD13)) & "') > 0 " & _
                  "Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(textECD13)) & "') > 0 " & _
                  "Or InStr(NLS_Upper(FA134),'" & UCase(ChgSQL(textECD13)) & "') > 0) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      With rsTmp
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If strFAText = "" Then
               strFAText = "國外代理人檔：" & rsTmp.Fields(0)
            Else
               strFAText = strFAText & "," & rsTmp.Fields(0)
            End If
            rsTmp.MoveNext
         Loop
      End With
   End If
   rsTmp.Close
   
   strPCUText = ""
   '潛在客戶檔
   strSql = "SELECT PCU01||PCU02 FROM potcustomer " & _
                  "Where (instr(NLS_Upper(pcu18),'" & UCase(ChgSQL(textECD13)) & "')> 0) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      With rsTmp
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If strPCUText = "" Then
               strPCUText = "潛在客戶檔：" & rsTmp.Fields(0)
            Else
               strPCUText = strPCUText & "," & rsTmp.Fields(0)
            End If
            rsTmp.MoveNext
         Loop
      End With
   End If
   rsTmp.Close
   
   strPOCText = ""
   '國內潛在客戶檔
   strSql = "SELECT POC01||POC02 FROM potcustomer1 " & _
                  "Where (instr(NLS_Upper(poc09),'" & UCase(ChgSQL(textECD13)) & "')> 0) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      With rsTmp
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If strPOCText = "" Then
               strPOCText = "國內潛在客戶檔：" & rsTmp.Fields(0)
            Else
               strPOCText = strPOCText & "," & rsTmp.Fields(0)
            End If
            rsTmp.MoveNext
         Loop
      End With
   End If
   rsTmp.Close
   
   strECDText = ""
   '外法開拓客戶檔
   strSql = "SELECT ECD02||'-'||ECD01 FROM expandcusdetail " & _
                  "Where (instr(NLS_Upper(ecd13),'" & UCase(ChgSQL(textECD13)) & "')> 0) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      With rsTmp
         rsTmp.MoveFirst
         Do While Not rsTmp.EOF
            If strECDText = "" Then
               strECDText = "外法開拓客戶檔：" & rsTmp.Fields(0)
            Else
               strECDText = strECDText & "," & rsTmp.Fields(0)
            End If
            rsTmp.MoveNext
         Loop
      End With
   End If
   rsTmp.Close
   
   If strCUText <> "" Or strFAText <> "" Or strPCUText <> "" Or strPOCText <> "" Or strECDText <> "" Then
      CheckEmailExist = True
      If strCUText <> "" Then strChkText = strChkText & strCUText & vbCrLf
      If strFAText <> "" Then strChkText = strChkText & strFAText & vbCrLf
      If strPCUText <> "" Then strChkText = strChkText & strPCUText & vbCrLf
      If strPOCText <> "" Then strChkText = strChkText & strPOCText & vbCrLf
      If strECDText <> "" Then strChkText = strChkText & strECDText & vbCrLf
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Function
