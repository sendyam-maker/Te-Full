VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm170022 
   BorderStyle     =   1  '單線固定
   Caption         =   "年終獎金維護"
   ClientHeight    =   5604
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8376
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5604
   ScaleWidth      =   8376
   Begin VB.TextBox txtYB 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   24
      Left            =   6200
      MaxLength       =   1
      TabIndex        =   62
      Text            =   "1"
      Top             =   1250
      Width           =   300
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   26
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   5
      Text            =   "99999999"
      Top             =   1670
      Width           =   800
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   25
      Left            =   6830
      MaxLength       =   8
      TabIndex        =   17
      Text            =   "99999999"
      Top             =   4740
      Width           =   800
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   4
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "99999999"
      Top             =   1050
      Width           =   800
   End
   Begin VB.TextBox txtYB 
      Height          =   270
      Index           =   3
      Left            =   6000
      TabIndex        =   2
      Text            =   "96"
      Top             =   770
      Width           =   500
   End
   Begin VB.TextBox txtYB 
      Height          =   270
      Index           =   1
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "96"
      Top             =   770
      Width           =   500
   End
   Begin VB.TextBox txtYB 
      Height          =   270
      Index           =   2
      Left            =   3270
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "123456"
      Top             =   770
      Width           =   735
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   7
      Left            =   2010
      MaxLength       =   7
      TabIndex        =   7
      Text            =   "9999999"
      Top             =   2150
      Width           =   735
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   5
      Left            =   4110
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "99999999"
      Top             =   1670
      Width           =   800
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   9
      Left            =   2010
      MaxLength       =   7
      TabIndex        =   9
      Text            =   "9999999"
      Top             =   2570
      Width           =   735
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   6
      Left            =   4110
      MaxLength       =   8
      TabIndex        =   6
      Text            =   "99999999"
      Top             =   1910
      Width           =   800
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   17
      Left            =   6830
      MaxLength       =   8
      TabIndex        =   16
      Text            =   "99999999"
      Top             =   4515
      Width           =   800
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   10
      Left            =   2010
      MaxLength       =   7
      TabIndex        =   10
      Text            =   "9999999"
      Top             =   2810
      Width           =   735
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   15
      Left            =   4110
      MaxLength       =   8
      TabIndex        =   15
      Text            =   "99999999"
      Top             =   3770
      Width           =   800
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
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
            Picture         =   "frm170022.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   11
      Left            =   2010
      MaxLength       =   7
      TabIndex        =   11
      Text            =   "9999999"
      Top             =   3050
      Width           =   735
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   12
      Left            =   2010
      MaxLength       =   7
      TabIndex        =   12
      Text            =   "9999999"
      Top             =   3290
      Width           =   735
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   13
      Left            =   2010
      MaxLength       =   7
      TabIndex        =   13
      Text            =   "9999999"
      Top             =   3530
      Width           =   735
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   14
      Left            =   2010
      MaxLength       =   7
      TabIndex        =   14
      Text            =   "9999999"
      Top             =   3770
      Width           =   735
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   8
      Left            =   4110
      MaxLength       =   8
      TabIndex        =   8
      Text            =   "99999999"
      Top             =   2150
      Width           =   800
   End
   Begin VB.TextBox txtYB 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   16
      Left            =   6830
      MaxLength       =   8
      TabIndex        =   18
      Text            =   "99999999"
      Top             =   4985
      Width           =   800
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7500
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
            Picture         =   "frm170022.frx":20F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":2410
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":272C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":2908
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":2C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":2F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":325C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":3578
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":3894
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":3BB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170022.frx":3ECC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   8376
      _ExtentX        =   14774
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
   Begin VB.Label lblDsp 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "100"
      Height          =   180
      Index           =   6
      Left            =   4005
      TabIndex        =   63
      Top             =   1305
      Width           =   270
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   120
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5300
      Width           =   5610
      VariousPropertyBits=   671105055
      Size            =   "7223;529"
      Value           =   "textCUID"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "       未休假代金計算月薪＝12月的全月基本薪資+午餐津貼+職務津貼"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   26
      Left            =   144
      TabIndex        =   61
      Top             =   5076
      Width           =   5376
   End
   Begin VB.Label lblDsp 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "99,999,999"
      Height          =   180
      Index           =   11
      Left            =   6768
      TabIndex        =   60
      Top             =   2196
      Width           =   816
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "未休假代金計算月薪："
      Height          =   180
      Index           =   24
      Left            =   5004
      TabIndex        =   59
      Top             =   2196
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "紅　　利："
      Height          =   180
      Index           =   20
      Left            =   5900
      TabIndex        =   58
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代扣補充保費："
      Height          =   180
      Index           =   17
      Left            =   5490
      TabIndex        =   57
      Top             =   4800
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "       代扣補充保費＝（扣單總額－４＊投保金額）＊ 費率 "
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   13
      Left            =   120
      TabIndex        =   56
      Top             =   4800
      Width           =   4545
   End
   Begin VB.Label lblDsp 
      Caption         =   "台一國際專利"
      Height          =   180
      Index           =   10
      Left            =   6555
      TabIndex        =   55
      Top             =   1290
      Width           =   1700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公司別："
      Height          =   180
      Index           =   10
      Left            =   5370
      TabIndex        =   54
      Top             =   1290
      Width           =   720
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   180
      Index           =   33
      Left            =   4395
      TabIndex        =   53
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS : 代扣稅額＝(年終獎金＋特殊功績獎金＋紅利－缺勤扣款)＊稅率 "
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   52
      Top             =   4560
      Width           =   5340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "應發金額："
      Height          =   180
      Index           =   4
      Left            =   5900
      TabIndex        =   51
      Top             =   1950
      Width           =   900
   End
   Begin VB.Label lblDsp 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "99,999,999"
      Height          =   180
      Index           =   9
      Left            =   6780
      TabIndex        =   49
      Top             =   5295
      Width           =   810
   End
   Begin VB.Label lblDsp 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "99,999,999"
      Height          =   180
      Index           =   8
      Left            =   6780
      TabIndex        =   48
      Top             =   4230
      Width           =   810
   End
   Begin VB.Label lblDsp 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "99,999,999"
      Height          =   180
      Index           =   7
      Left            =   6780
      TabIndex        =   47
      Top             =   1950
      Width           =   810
   End
   Begin MSForms.Label lblName 
      Height          =   285
      Left            =   4080
      TabIndex        =   45
      Top             =   780
      Width           =   870
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1535;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "獎金年度： "
      Height          =   180
      Index           =   3
      Left            =   300
      TabIndex        =   44
      Top             =   810
      Width           =   945
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   8280
      Y1              =   1600
      Y2              =   1600
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   8280
      Y1              =   2500
      Y2              =   2500
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   8280
      Y1              =   4120
      Y2              =   4120
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   8280
      Y1              =   4485
      Y2              =   4485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "核發獎金基數："
      Height          =   180
      Index           =   32
      Left            =   2370
      TabIndex        =   43
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "考　　績："
      Height          =   180
      Index           =   31
      Left            =   300
      TabIndex        =   42
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "優等"
      Height          =   180
      Index           =   5
      Left            =   1250
      TabIndex        =   41
      Top             =   1290
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年度工作總天數："
      Height          =   180
      Index           =   30
      Left            =   5370
      TabIndex        =   40
      Top             =   1050
      Width           =   1440
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "365"
      Height          =   180
      Index           =   4
      Left            =   6900
      TabIndex        =   39
      Top             =   1050
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "基準月數："
      Height          =   180
      Index           =   12
      Left            =   300
      TabIndex        =   38
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "4.0"
      Height          =   180
      Index           =   3
      Left            =   1250
      TabIndex        =   37
      Top             =   1050
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "平均基準月薪："
      Height          =   180
      Index           =   1
      Left            =   2370
      TabIndex        =   36
      Top             =   1050
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公傷假時數："
      Height          =   180
      Index           =   15
      Left            =   915
      TabIndex        =   35
      Top             =   3810
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "實領金額："
      Height          =   180
      Index           =   11
      Left            =   5850
      TabIndex        =   34
      Top             =   5295
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "流產假時數："
      Height          =   180
      Index           =   9
      Left            =   915
      TabIndex        =   33
      Top             =   3570
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代扣稅額："
      Height          =   180
      Index           =   8
      Left            =   5850
      TabIndex        =   32
      Top             =   4560
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "扣  除        病假時數："
      Height          =   180
      Index           =   6
      Left            =   300
      TabIndex        =   31
      Top             =   2610
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "借支金額："
      Height          =   180
      Index           =   5
      Left            =   5850
      TabIndex        =   30
      Top             =   5040
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   2370
      TabIndex        =   29
      Top             =   810
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門："
      Height          =   180
      Index           =   2
      Left            =   5370
      TabIndex        =   28
      Top             =   810
      Width           =   540
   End
   Begin VB.Label lblDsp 
      AutoSize        =   -1  'True
      Caption         =   "台一部"
      Height          =   180
      Index           =   2
      Left            =   6525
      TabIndex        =   27
      Top             =   810
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年     終       獎      金："
      Height          =   180
      Index           =   14
      Left            =   300
      TabIndex        =   26
      Top             =   1710
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "事假時數："
      Height          =   180
      Index           =   16
      Left            =   1095
      TabIndex        =   25
      Top             =   2850
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "特  殊  功  績  獎  金："
      Height          =   180
      Index           =   18
      Left            =   300
      TabIndex        =   24
      Top             =   1950
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "缺勤扣款："
      Height          =   180
      Index           =   19
      Left            =   3085
      TabIndex        =   23
      Top             =   3810
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "未 休 特 別 假 代 金：                   時"
      Height          =   180
      Index           =   21
      Left            =   300
      TabIndex        =   22
      Top             =   2196
      Width           =   2820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "曠職時數："
      Height          =   180
      Index           =   22
      Left            =   1095
      TabIndex        =   21
      Top             =   3090
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "應領金額："
      Height          =   180
      Index           =   23
      Left            =   5850
      TabIndex        =   20
      Top             =   4230
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "產假時數："
      Height          =   180
      Index           =   25
      Left            =   1095
      TabIndex        =   19
      Top             =   3330
      Width           =   900
   End
End
Attribute VB_Name = "frm170022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/20 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2008/12/27 add by sonia
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_YB As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim m_bConfirmCheck As Boolean
Dim m_bActived As Boolean
Dim NHI() As String


Private Sub Form_Activate()
   If m_bActived = False Then
      SetInputEntry
      m_bActived = True
   End If
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
   UpdateToolbarState
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170022 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from YearBonus where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_YB = .Fields.Count
      ReDim m_FieldList(TF_YB) As FIELDITEM
      For Each oText In txtYB
         idx = oText.Index
         m_FieldList(idx).fiName = "YB" & Format(idx, "00")
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
   
   ReDim NHI(TF_NHI) As String  '2013/1/22 ADD BY SONIA
   
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
Dim stKey01 As String
Dim stKey02 As String
Dim adoRst As New ADODB.Recordset
   
   stKey01 = Val(txtYB(1)) + 1911
   stKey02 = txtYB(2)
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM YearBonus" & _
            " WHERE yb01 = '" & stKey01 & "' and yb02= '" & stKey02 & "'"
      Case -2
         strExc(0) = "SELECT * FROM YearBonus order by 1 ASC"
      Case -1
         strExc(0) = "SELECT * FROM YearBonus" & _
            " WHERE yb01||yb02 <'" & stKey01 & stKey02 & "' order by 1 DESC"
      Case 1
         strExc(0) = "SELECT * FROM YearBonus" & _
            " WHERE yb01||yb02 >'" & stKey01 & stKey02 & "' order by 1 ASC"
      Case 2
         strExc(0) = "SELECT * FROM YearBonus order by 1 DESC"
   End Select
   intI = 1
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
      txtYB(1).SetFocus
      txtYB_GotFocus 1
   End If
End Function

Private Sub txtYB_GotFocus(Index As Integer)
   TextInverse txtYB(Index)
   CloseIme
End Sub

Private Sub ClearField()
   lblName = Empty 'Modify By Sindy 2021/12/20
   For Each oText In txtYB
      oText.Text = Empty
   Next
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_YB
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   m_bConfirmCheck = False
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   With p_Rst
   If .RecordCount > 0 Then
      For Each oText In txtYB
         idx = oText.Index
         '獎金年度轉民國年
         If idx = 1 Then
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName) - 1911
         Else
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         End If
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         oText.Text = m_FieldList(idx).fiOldData
         '2010/12/30 add by sonia
         Select Case idx
            Case 5, 6, 8, 15, 26
               oText.Tag = oText.Text
         End Select
         '2010/12/30 emd
      Next
      
      If ClsPDGetStaffN(txtYB(2), strExc(1), , True) Then
         lblName = strExc(1) 'Modify By Sindy 2021/12/20
         lblDsp(2) = GetDepartmentName(txtYB(3))
      End If
      'Added by Morgan 2024/2/2 年終隔年才發要-1
      If txtYB(1) >= (Left(新部門啟用日, 4) - 1911 - 1) Then
         lblDsp(2) = GetPrjSalesBlack(txtYB(3), True)
      End If
      'end 2024/2/2
                     
      '取得年終獎金基準月數
      lblDsp(3) = GetYearBonusMonth(txtYB(1), txtYB(2))
      '取得年度工作總天數
      lblDsp(4) = GetYearWorkDay(txtYB(1), txtYB(2))
      '取得考績及核發獎金基數
      If GetYearMerit(txtYB(1), txtYB(2), strExc(1), strExc(2)) = True Then
         lblDsp(5) = strExc(1)
         lblDsp(6) = strExc(2)
      End If
      lblDsp(10) = CompNameQuery(txtYB(24))   '2008/12/31 add by sonia
      '計算應發金額,應領金額,實領金額
      SetRefData
      
      CUID(1) = "" & .Fields("yb18")
      CUID(2) = "" & .Fields("yb19")
      CUID(3) = "" & .Fields("yb20")
      CUID(4) = "" & .Fields("yb21")
      CUID(5) = "" & .Fields("yb22")
      CUID(6) = "" & .Fields("yb23")
      
      'Added by Morgan 2024/1/30 若獎金從沒有補充保費改成有補充保費時會用到
      'Removed by Morgan 2024/2/27 改存檔檢查時設定
      'NHI(2) = "" & .Fields("yb19")
      'NHI(10) = "" & .Fields("yb20")
      'end 2024/2/27
      'end 2024/1/30
   End If
   End With
   UpdateCUID CUID, textCUID
   txtYB(1).Tag = txtYB(1)
   txtYB(2).Tag = txtYB(2)
   
   '2013/1/28 add by sonia 取得補充保費明細資料
   strExc(0) = "select * from NHI2ND WHERE NHI01='" & txtYB(2) & "' AND SUBSTR(NHI02,1,4)=" & Val(txtYB(1) + 1912) & " AND NHI03='" & "50" & "' AND NHI04='" & "1" & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      NHI(2) = "" & Val(RsTemp.Fields("NHI02"))
      NHI(10) = "" & Val(RsTemp.Fields("NHI10"))
   End If
   '2013/1/28 end
   
   '2020/1/13 add by sonia 未休假代金計算月薪＝12月基本薪資+午餐津貼+職務津貼
   strExc(0) = "select NVL(M1.SM26,0)+NVL(M2.SM26,0) from YEARBONUS,SALARYMONTH M1,SALARYMONTH M2 " & _
               "WHERE YB02='" & txtYB(2) & "' AND YB01=" & Val(txtYB(1) + 1911) & " AND YB02=M1.SM01(+) AND YB01||'12'=M1.SM02(+) AND SUBSTR(M1.SM01,1,2)||'A'||SUBSTR(M1.SM01,4,2)=M2.SM01(+) AND M1.SM02=M2.SM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      lblDsp(11) = "" & Val(RsTemp.Fields(0))
      lblDsp(11) = Format(lblDsp(11), "#,###")
   End If
   '2020/1/13 end
   
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtYB
      oText.Locked = bLocked
   Next
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
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF3 ' 修改
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

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
         SetCtrlReadOnly True
         ClearField
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
      Case vbKeyF9 ' 確定
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         SetInputEntry
         
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
            txtYB(1) = txtYB(1).Tag
            txtYB(2) = txtYB(2).Tag
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
         If m_bUpdate And txtYB(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtYB(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtYB(1) <> "" Then
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
         txtYB(1).Locked = False
         If Me.Visible = True Then
            txtYB(1).SetFocus
         End If
      Case 2
         txtYB(1).Locked = True
         If Me.Visible = True Then
            txtYB(4).SetFocus
         End If
      Case 4
         txtYB(1).Locked = False
         txtYB(2).Locked = False
         If Me.Visible = True Then
            txtYB(1).SetFocus
         End If
      Case Else
         txtYB(1).Locked = True
         If Me.Visible = True Then
            txtYB(1).SetFocus
         End If
   End Select
   txtYB(3).Locked = True    '部門別鎖住
   txtYB(3).Enabled = False  '部門別鎖住
   txtYB(24).Locked = True   '公司別鎖住
   txtYB(24).Enabled = False '公司別鎖住
   'txtYB(17).Locked = True   '代扣稅額鎖住   '2009/1/10 add by sonia  2009/1/16放開,因為76028
   '2013/1/17 add by sonia
   txtYB(25).Locked = True   '代扣補充保費
   txtYB(25).Enabled = False '代扣補充保費
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
               txtYB(1).SetFocus
               txtYB_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   For Each oText In txtYB
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtYB_Validate idx, bCancel
         If bCancel = True Then
            txtYB(idx).SetFocus
            txtYB_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   '查詢
   If m_EditMode = 4 Then
      If txtYB(1) = "" Then
         ShowMsg "請輸入獎金年度 !"
         txtYB(1).SetFocus
         txtYB_GotFocus 1
         GoTo EscPoint
      End If
      If txtYB(2) = "" Then
         ShowMsg "請輸入員工代號 !"
         txtYB(2).SetFocus
         txtYB_GotFocus 2
         GoTo EscPoint
      End If
      
   '維護
   Else
      If txtYB(1) = "" And txtYB(1).Locked = False Then
         ShowMsg "請輸入獎金年度 !"
         txtYB(1).SetFocus
         txtYB_GotFocus 1
         GoTo EscPoint
      End If
      If txtYB(2) = "" And txtYB(2).Locked = False Then
         ShowMsg "請輸入員工代號 !"
         txtYB(2).SetFocus
         txtYB_GotFocus 2
         GoTo EscPoint
      End If
      If lblDsp(7) = "" Or lblDsp(7) = "0" Then
         ShowMsg "請輸入獎金資料 !"
         txtYB(5).SetFocus
         txtYB_GotFocus 5
         GoTo EscPoint
      End If
      
      'Added by Morgan 2024/2/27
      '新增或原來沒有補充保費時設定代扣日期時間為該年度的年終獎金補充保費最一筆的日期時間
      If m_EditMode = "1" Or NHI(2) = "" Then
         strExc(0) = "select nhi02,nhi10 from yearbonus,nhi2nd where yb01=" & (Val(txtYB(1)) + 1911) & " and nhi02(+)=yb19 and nhi01(+)=yb02 order by 1,2 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            NHI(2) = RsTemp.Fields("nhi02")
            NHI(10) = RsTemp.Fields("nhi10")
         Else
            MsgBox txtYB(1) & "年度的年終獎金補充保費代扣日期時間設定失敗，請確認該年度年終獎金是否已計算！", vbCritical
            GoTo EscPoint
         End If
      End If
      'end 2024/2/27
      
      'Added by Morgan 2024/1/30
      '檢查不可有晚於該筆資料的補充保費
      NHI(1) = txtYB(2)
      If PUB_ChkNHi2nd(NHI(1), NHI(2), NHI(10), IIf(m_EditMode = "1", True, False), False) = False Then
         GoTo EscPoint
      End If
      'end 2024/1/30
      
   End If
   TxtValidate = True
   
EscPoint:
   m_bConfirmCheck = False
    
End Function

Private Function AddRecord() As Boolean
Dim stCols As String, stValues As String, stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Added by Morgan 2024/2/27 從 UpdateFieldNewData 移過來
   NHI(1) = txtYB(2)
   NHI(3) = "50"
   NHI(4) = "1"
   NHI(7) = Val(txtYB(5)) + Val(txtYB(6)) + Val(txtYB(26)) - Val(txtYB(15))
   NHI(5) = 0: NHI(6) = 0: NHI(8) = 0
   NHI(11) = txtYB(24)
   PUB_NHI2nd NHI(1), NHI(2), NHI(3), NHI(4), NHI(7), NHI(5), NHI(6), NHI(8), NHI(10), NHI(11), NHI(13)
   txtYB(25) = NHI(6)
   
   m_FieldList(25).fiNewData = txtYB(25)
   '新增補充保費
   PUB_InsertNHI2nd NHI
   'end 2024/2/27
   
   
   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtYB
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            '日期
            'If idx = 1 Then
            '   stValues = stValues & "," & CNULL(Val((m_FieldList(idx).fiNewData) + 1911), True)
            'Else
               stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
            'End If
         End If
      End If
   Next
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO YearBonus (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
'   stSQL = "select max(yb02) from YearBonus where yb01='" & txtYB(1) & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
'   If intI = 1 Then
'      txtYB(2) = RsTemp.Fields(0)
'   End If
   
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
   
   'Added by Morgan 2024/2/27 從 UpdateFieldNewData 移過來
   NHI(1) = txtYB(2)
   NHI(3) = "50"
   NHI(4) = "1"
   NHI(7) = Val(txtYB(5)) + Val(txtYB(6)) + Val(txtYB(26)) - Val(txtYB(15))
   NHI(5) = 0: NHI(6) = 0: NHI(8) = 0
   NHI(11) = txtYB(24)
   PUB_NHI2nd NHI(1), NHI(2), NHI(3), NHI(4), NHI(7), NHI(5), NHI(6), NHI(8), NHI(10), NHI(11), NHI(13)
   txtYB(25) = NHI(6)
   
   m_FieldList(25).fiNewData = txtYB(25)
   '新增補充保費
   PUB_InsertNHI2nd NHI
   'end 2024/2/27
      
      
   stSQL = "begin user_data.user_enabled:=1; UPDATE YearBonus SET "
   stSet = ""
   For Each oText In txtYB
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
      stSQL = stSQL & stSet & " where yb01=" & Val(txtYB(1)) + 1911 & " and yb02='" & txtYB(2) & "'; end; "
      
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   
   '2013/1/21 ADD BY SONIA新增補充保費
   'PUB_InsertNHI2nd NHI
   '2013/1/21 END
   
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

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
   
   'Removed by Morgan 2024/2/27 移到存檔才能確保在同一個Transaction
   ''2013/1/21 ADD BY SONIA
   'NHI(1) = txtYB(2)
   'NHI(3) = "50"
   'NHI(4) = "1"
   
   ''Removed by Morgan 2024/2/27 改存檔檢查時設定為該年度年終獎金的日期及最大時間
   ''If m_EditMode = 1 Then
   ''   NHI(2) = strSrvDate(1)
   ''   NHI(10) = ServerTime
   ''End If
   
   ''modify by sonia 2018/1/11 +YB26
   'NHI(7) = Val(txtYB(5)) + Val(txtYB(6)) + Val(txtYB(26)) - Val(txtYB(15))
   'NHI(5) = 0: NHI(6) = 0: NHI(8) = 0
   'NHI(11) = txtYB(24) 'Added by Morgan 2013/2/26
   'PUB_NHI2nd NHI(1), NHI(2), NHI(3), NHI(4), NHI(7), NHI(5), NHI(6), NHI(8), NHI(10), NHI(11), NHI(13) 'Modified by Morgan 2013/3/12 +NHI13 2014/5/1 +NHI11
   'txtYB(25) = NHI(6)
   
   ''新增補充保費
   'PUB_InsertNHI2nd NHI
   ''2013/1/21 END
   'end 2024/2/27
         
   For Each oText In txtYB
      idx = oText.Index
      Select Case idx
         Case 1
            m_FieldList(idx).fiNewData = Val(oText.Text) + 1911
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub

Private Sub txtYB_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 2, 24
      Case 7, 9, 10, 11, 12, 13, 14
         KeyAscii = Pub_NumAscii(KeyAscii, True)
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub
'2011/4/20 ADD BY SONIA
Private Sub txtYB_LostFocus(Index As Integer)
   Select Case Index
      Case 17   '計算應發金額,應領金額,實領金額
         SetRefData (Index)
   End Select
End Sub
'2011/4/20 END

Private Sub txtYB_Validate(Index As Integer, Cancel As Boolean)
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Select Case Index
         Case 2
            If txtYB(Index) <> "" Then
               lblName = "" 'Modify By Sindy 2021/12/20
               'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
               'If ChkStaffID(Replace(txtYB(Index), "A", "0")) = True Then
               If ChkStaffID(Left(txtYB(Index), 1) & Replace(Mid(txtYB(Index), 2), "A", "0")) = True Then
                  Cancel = True
               End If
               If Cancel = False Then
                  If ClsPDGetStaffN(txtYB(Index), strExc(1), , True) = False Then
                     Cancel = True
                  Else
                     lblName = strExc(1) 'Modify By Sindy 2021/12/20
                     'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
                     'txtYB(3) = GetStaffDepartment(Replace(txtYB(Index), "A", "0"))
                     txtYB(3) = GetStaffDepartment(Left(txtYB(Index), 1) & Replace(Mid(txtYB(Index), 2), "A", "0"))
                     lblDsp(2) = GetDepartmentName(txtYB(3))
                     'Added by Morgan 2024/2/2 年終隔年才發要-1
                     If txtYB(1) = "" Or txtYB(1) >= (Left(新部門啟用日, 4) - 1911 - 1) Then
                        txtYB(3) = PUB_GetST93(txtYB(2))
                        If txtYB(3) <> "" Then
                           lblDsp(2) = GetPrjSalesBlack(txtYB(3), True)
                        End If
                     End If
                     'end 2024/2/2
                  
                     '取得年終獎金基準月數
                     lblDsp(3) = GetYearBonusMonth(txtYB(1), txtYB(Index))
                     '取得年度工作總天數
                     lblDsp(4) = GetYearWorkDay(txtYB(1), txtYB(Index))
                     '取得考績及核發獎金基數
                     If GetYearMerit(txtYB(1), txtYB(Index), strExc(1), strExc(2)) = True Then
                        lblDsp(5) = strExc(1)
                        lblDsp(6) = strExc(2)
                     End If
                  End If
                  '2008/12/31 add by sonia
                  If Cancel = False Then
                     If ClsPDGetStaffComp(txtYB(Index), strExc(1), True) = False Then
                        Cancel = True
                     Else
                        txtYB(24) = strExc(1)
                        lblDsp(10) = CompNameQuery(txtYB(24))
                     End If
                  End If
                  '2008/12/31 END
               End If
            End If
         '2010/12/30 modify by sonia
         'Case 5, 6, 8, 15, 16, 17  '計算應發金額,應領金額,實領金額
         '   SetRefData
         Case 5, 6, 15, 26   '計算應發金額,應領金額,實領金額   '2018/1/12 +yb26
            If txtYB(Index).Text <> txtYB(Index).Tag Then
               If MsgBox("是否重新計算代扣稅額？", vbExclamation + vbOKCancel) = vbOK Then
                  txtYB(17) = 0
               End If
            End If
            '2013/1/21 ADD BY SONIA有修改時清除代扣補充保費,存檔時重新計算再顯示
            If txtYB(Index).Tag <> txtYB(Index).Text Then txtYB(25) = 0
            '2013/1/21 END
            txtYB(Index).Tag = txtYB(Index).Text
            SetRefData
         Case 8, 16 '計算應發金額,應領金額,實領金額
            SetRefData
         '2010/12/30 end
      End Select
      
      If Cancel = True Then TextInverse txtYB(Index)
      
      '若是案確定的檢查時略過, 檢查代號檔
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
         End Select
      End If
   End If
End Sub

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   '刪除
   stSQL = "delete from YearBonus where yb01=" & Val(txtYB(1)) + 1911 & " and yb02='" & txtYB(2) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   '刪除當筆資料
   strSql = "DELETE NHI2ND WHERE NHI01='" & txtYB(2) & "' AND SUBSTR(NHI02,1,4)=" & Val(txtYB(1) + 1912) & " AND NHI03='" & "50" & "' AND NHI04='" & "1" & "' "
   cnnConnection.Execute strSql, intI
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtYB(1).Tag = ""
   txtYB(2).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Function SetRefData(Optional ByVal Index As Integer = 0)
Dim m_taxrate As String   '2010/12/30 add by sonia 非固定之薪資所得扣繳稅率
   
   lblDsp(7) = "": lblDsp(8) = "": lblDsp(9) = ""
   'modify by sonia 2018/1/11 +YB26
   lblDsp(7) = Val(txtYB(5)) + Val(txtYB(6)) + Val(txtYB(26)) + Val(txtYB(8))
   'modify by sonia 2018/1/30 婧瑄說應領不可扣除借支,實領再扣除
   lblDsp(8) = Val(lblDsp(7)) - Val(txtYB(15))
   '2013/1/21 add by sonia 重新計算代扣補充保費
   
   '2013/1/21 end
   
   '2009/1/10 add by sonia
   '2011/4/20 ADD BY SONIA 先問是否重新計算稅額,否則不扣稅者無法作業
   If Index = 17 Then
      If MsgBox("是否重新計算代扣稅額？", vbExclamation + vbOKCancel) = vbCancel Then
         GoTo SetTag
      Else
         txtYB(17) = 0
      End If
   End If
   '2011/4/20 END
   If Val(txtYB(17)) = 0 Then
      '2010/12/30 modify by sonia 非固定之薪資所得扣繳稅率改抓 翻譯所得oc01='01'的稅率
      'txtYB(17) = Round((Val(txtYB(5)) + Val(txtYB(6)) - Val(txtYB(15))) * 6 / 100, 0)
      m_taxrate = 0
      strExc(0) = "select oc04 from OtherSalaryCode where oc01='01'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_taxrate = "" & RsTemp.Fields(0)
      End If
      'modify by sonia 2018/1/11 +YB26
      txtYB(17) = Round((Val(txtYB(5)) + Val(txtYB(6)) + Val(txtYB(26)) - Val(txtYB(15))) * m_taxrate / 100, 0)
      '2010/12/30 end
      If txtYB(17) < 2000 Then txtYB(17) = ""
      'add by sonia 2023/1/17 2018年以後年終獎金+特殊功績獎金-缺勤扣款計算所得稅 < 84501 不扣稅,
      If Val(txtYB(1)) >= 107 And Val(txtYB(5)) + Val(txtYB(6)) + Val(txtYB(26)) - Val(txtYB(15)) * m_taxrate / 100 < 84501 Then
         txtYB(17) = ""
      End If
      'end 2023/1/17
   End If
   '2009/1/10 end
SetTag:  '2011/4/20 ADD BY SONIA
   '2013/1/17 modify by sonia 再減補充保費
   'modify by sonia 2018/1/30 婧瑄說應領不可扣除借支,實領再扣除
   lblDsp(9) = Val(lblDsp(8)) - Val(txtYB(17)) - Val(txtYB(25)) - Val(txtYB(16))
   lblDsp(7) = Format(lblDsp(7), "#,###")
   lblDsp(8) = Format(lblDsp(8), "#,###")
   lblDsp(9) = Format(lblDsp(9), "#,###")
End Function

' 取得年度工作總天數
Public Function GetYearWorkDay(ByVal strYear As String, ByVal StrStaff As String) As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   GetYearWorkDay = Empty
   strSql = "SELECT sum(sm27) FROM SalaryMonth WHERE sm01='" & StrStaff & "' and sm02>= '" & Val(strYear) + 1911 & "01' and sm02<='" & Val(strYear) + 1911 & "12' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then
         GetYearWorkDay = rsTmp.Fields(0)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
